#!/usr/bin/env python3
"""Excel を使わずに src/ の VBA を実行してテストする。

    python3 tests/run.py            # 全部 (コンパイル確認 + テスト)
    python3 tests/run.py --compile  # コンパイル確認だけ (速い)
    python3 tests/run.py --bisect mColorUtil
                                    # コンパイルに失敗するモジュールの原因行を絞り込む

LibreOffice の Basic は VBA と同じ言語ではないため、ここで緑になっても
Excel で緑になることを保証はしない (docs/TESTING.md の「再現度の限界」を参照)。
それでも、AI が書きがちな「そもそも動かない VBA」はここでほぼ落ちる。
"""

import argparse
import pathlib
import re
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from harness import lo, source  # noqa: E402

VBA_DIR = HERE / "vba"
# ハングしたらテスト失敗として扱うための上限 (LibreOffice が
# モーダルダイアログを出して固まるケースを CI で放置しないため)。
WORKER_TIMEOUT = 300


# ---------------------------------------------------------------- 実行本体
def compile_check(book, modules):
    """モジュールを1つずつ叩いてコンパイルが通っているか確かめる。

    LibreOffice Basic はコンパイルエラーを例外ではなく「関数が何も返さない」
    形で表面化させることがあり、しかも 1 つ壊れたモジュールがライブラリ全体を
    道連れにする。そのため全モジュールを個別に確認する。
    """
    broken = []
    for name in modules:
        try:
            got = book.call(name, source.PROBE_NAME)
        except Exception as exc:  # noqa: BLE001
            got = "<{}: {}>".format(type(exc).__name__, exc)
        if got != source.PROBE_OK:
            broken.append(name)
    return broken


def build_test_main(test_modules):
    """各テストモジュールの Run を呼ぶエントリポイントを生成する。

    Basic には名前による動的ディスパッチが無いため、モジュールごとに
    ラッパー Sub を生成する。1 つのモジュールが実行時エラーで落ちても
    他のテスト結果が失われないよう、呼び出しは個別にトラップする。
    """
    body = ["Option VBASupport 1", "Option Explicit", "",
            "Public Function RunAll() As String", "    TestLib.TBegin"]
    for name in test_modules:
        body.append("    Run_{}".format(name))
    body += ["    RunAll = TestLib.TReport()", "End Function", ""]
    for name in test_modules:
        body += [
            "Private Sub Run_{}()".format(name),
            "    On Error GoTo Trapped",
            '    TestLib.TModule "{}"'.format(name),
            "    {}.Run".format(name),
            "    Exit Sub",
            "Trapped:",
            '    TestLib.Failed "{}.Run が実行時エラーで中断した", "E" & Err.Number & ": " & Err.Description'.format(name),
            "End Sub",
            "",
        ]
    return "\n".join(body) + "\n"


def parse_report(report):
    results, summary = [], None
    for line in report.split("\n"):
        if not line.strip():
            continue
        parts = line.split("\t")
        if parts[0] == "SUMMARY":
            summary = (int(parts[1]), int(parts[2]))
        elif parts[0] in ("PASS", "FAIL"):
            results.append(dict(kind=parts[0], module=parts[1], desc=parts[2],
                                got=parts[3], want=parts[4]))
    return results, summary


def check_tb_borders(book):
    """tb が引いた罫線を UNO 側から検証する。

    VBA 互換層の Borders(...).LineStyle は読み戻しが機能しないため、
    LibreOffice のセル属性 (TopBorder / BottomBorder) を直接見る。
    LibreOffice が再現しない部分 (破線の線種・Union への一括適用など) には
    踏み込まず、確実に再現される「外枠が引かれたか」だけを確認する。
    """
    out = []
    try:
        sheet = book.doc.Sheets.getByName("TbCase")
    except Exception:  # noqa: BLE001
        return [dict(kind="FAIL", module="Test_mUtil(UNO)", desc="TbCase シートが作られていない",
                     got="(なし)", want="TbCase")]

    def width(cell, side):
        return getattr(sheet.getCellRangeByName(cell), side).LineWidth

    checks = [
        ("tb: 表の上端に外枠が引かれる", width("A1", "TopBorder") > 0, True),
        ("tb: 表の下端に外枠が引かれる", width("A4", "BottomBorder") > 0, True),
        ("tb: ヘッダー行の下端に罫線が引かれる", width("A1", "BottomBorder") > 0, True),
        ("tb: 範囲外のセルには罫線を引かない", width("A6", "TopBorder") > 0, False),
    ]
    for desc, got, want in checks:
        out.append(dict(kind="PASS" if got == want else "FAIL", module="Test_mUtil(UNO)",
                        desc=desc, got=str(got), want=str(want)))
    return out


def run(compile_only=False):
    src_modules = source.load_src_modules()
    test_modules = source.load_test_modules(VBA_DIR)
    test_runners = [n for n in test_modules if n.startswith("Test_")]

    office = lo.LibreOffice().start()
    book = office.new_workbook()
    modules = dict(src_modules)
    if not compile_only:
        modules.update(test_modules)
        modules["TestMain"] = build_test_main(test_runners)
    book.install(modules)

    broken = compile_check(book, sorted(src_modules))
    if not compile_only:
        broken += compile_check(book, sorted(test_modules))

    results, summary = [], (0, 0)
    if broken:
        for name in broken:
            results.append(dict(kind="FAIL", module=name, desc="モジュールのコンパイルに失敗",
                                got="(コンパイル不可)", want="コンパイル成功"))
    elif not compile_only:
        report = book.call("TestMain", "RunAll")
        if report is None:
            results.append(dict(kind="FAIL", module="TestMain", desc="テストの実行に失敗",
                                got="(戻り値なし)", want="テスト結果"))
        else:
            results, summary = parse_report(report)
            results += check_tb_borders(book)
    else:
        results.append(dict(kind="PASS", module="(compile)", desc="全モジュールがコンパイルできた",
                            got="ok", want="ok"))

    book.close()
    # テストの後に行う (別のドキュメントを開くと ActiveWorkbook が変わるため)
    results += check_real_win32(office)
    office.stop()
    return results


def check_real_win32(office):
    """偽物に差し替えた本物の Win32 モジュールを確かめる。

    - 偽物が本物の公開手続きをすべて持っているか (差し替えても呼び出し側が壊れないか)
    - 本物の構文 (PtrSafe を取り除いて別のドキュメントでコンパイルする)

    LibreOffice は未宣言の変数や手続き名の誤りを実行するまで報告しないので、
    本物について分かるのは構文 (ブロックの対応、Declare や #If の書式) までになる。
    """
    out = []
    mismatches = source.fake_api_mismatches()
    for name in source.fake_names():
        missing = mismatches.get(name)
        out.append(dict(kind="FAIL" if missing else "PASS", module="Win32 (本物と偽物)",
                        desc="{}: 偽物が本物の公開手続きをすべて持つ".format(name),
                        got=", ".join(missing) if missing else "ok", want="ok"))

    # 壊れたモジュールはライブラリ全体を道連れにするので、1 モジュールずつ別のドキュメントで確かめる
    real = source.load_real_win32_modules()
    for name in sorted(real):
        book = office.new_workbook()
        book.install({name: real[name]})
        ok = not compile_check(book, [name])
        book.close()
        out.append(dict(kind="PASS" if ok else "FAIL", module="Win32 (本物と偽物)",
                        desc="{}: 本物の構文が正しい (PtrSafe を除いてコンパイル)".format(name),
                        got="ok" if ok else "(コンパイル不可)", want="ok"))
    return out


def bisect(module_name):
    """モジュールを手続き単位で切り詰めて、どこでコンパイルが壊れるか探す。"""
    path = next(p for p in source.SRC_DIR.iterdir() if p.stem == module_name)
    code = source.strip_vbe_header(source.read_source(path)).replace("\r\n", "\n")
    if module_name in source.fake_names():
        code = source.strip_ptrsafe(code)  # 本物の Win32 モジュールは構文確認と同じ条件で調べる
    lines = code.split("\n")
    cuts = [i + 1 for i, l in enumerate(lines)
            if re.match(r"^(End Sub|End Function)\s*$", l) or re.match(r"^\s*(Public |Private )?Sub .*End Sub\s*$", l)]

    office = lo.LibreOffice().start()
    last_ok = 0
    try:
        for cut in cuts:
            book = office.new_workbook()
            head = "\n".join(lines[:cut])
            prologue = "Option ClassModule\nOption VBASupport 1\n" if path.suffix.lower() == ".cls" else "Option VBASupport 1\n"
            book.install({module_name: prologue + head + source._PROBE_SRC})
            ok = book.call(module_name, source.PROBE_NAME) == source.PROBE_OK
            book.close()
            if not ok:
                print("コンパイルが壊れる箇所: {}:{}-{}".format(path.name, last_ok + 1, cut))
                print("--- 該当範囲 ---")
                for i in range(last_ok, cut):
                    print("{:5}: {}".format(i + 1, lines[i]))
                return 1
            last_ok = cut
    finally:
        office.stop()
    print("{}: 全体がコンパイルできました".format(path.name))
    return 0


# ---------------------------------------------------------------- 出力
def report(results):
    failures = [r for r in results if r["kind"] == "FAIL"]
    by_module = {}
    for r in results:
        by_module.setdefault(r["module"], []).append(r)
    for module, rs in by_module.items():
        print("\n{}".format(module))
        for r in rs:
            if r["kind"] == "PASS":
                print("  ok   {}".format(r["desc"]))
            else:
                print("  FAIL {}".format(r["desc"]))
                print("         got : {}".format(r["got"]))
                print("         want: {}".format(r["want"]))
    print("\n{} passed, {} failed".format(len(results) - len(failures), len(failures)))
    return 1 if failures else 0


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--compile", action="store_true", help="コンパイル確認のみ")
    ap.add_argument("--bisect", metavar="MODULE", help="コンパイル失敗箇所を絞り込む")
    ap.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()

    if not args.worker:
        # LibreOffice が固まっても CI が止まらないよう、実処理は子プロセスで動かす。
        cmd = [sys.executable, str(pathlib.Path(__file__).resolve()), "--worker"] + sys.argv[1:]
        try:
            return subprocess.call(cmd, timeout=WORKER_TIMEOUT)
        except subprocess.TimeoutExpired:
            print("TIMEOUT: LibreOffice が {} 秒以内に応答しませんでした".format(WORKER_TIMEOUT))
            subprocess.call(["pkill", "-9", "-f", "[s]office"])
            return 1

    if args.bisect:
        return bisect(args.bisect)
    return report(run(compile_only=args.compile))


if __name__ == "__main__":
    sys.exit(main())
