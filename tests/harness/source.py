"""src/*.bas, *.cls を LibreOffice Basic に読み込める形へ変換する。

VBA の .bas / .cls は、そのままでは LibreOffice Basic のモジュールとして
読み込めない (VBE 固有の Attribute 行などが含まれるため)。ここでは

  * VBE 固有のヘッダ行 (Attribute VB_*, VERSION 1.0 CLASS, BEGIN/END/MultiUse) を除去
  * .cls には Option ClassModule を付与 (LibreOffice でクラスモジュールとして扱わせる)
  * 全モジュールに Option VBASupport 1 を付与 (Excel の VBA 方言・定数を有効化)
だけを行う。コードの本文には手を入れない。

Win32 API を使うモジュール (tests/fakes/ に同名の偽物があるもの) は、テストでは
偽物に差し替える。本物は構文の確認だけに使い、そのときに限り Declare 文の PtrSafe を
取り除く (LibreOffice Basic は PtrSafe を解釈できず、#If で分岐させてもモジュールごと
コンパイルに失敗する。Declare と LongPtr は通る)。
"""

import pathlib
import re

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC_DIR = REPO_ROOT / "src"
FAKES_DIR = REPO_ROOT / "tests" / "fakes"

# VBE が .bas / .cls の先頭に書き出すメタ情報。LibreOffice Basic は解釈できない。
_VBE_HEADER = re.compile(
    r"^\s*(Attribute\s+VB_|VERSION\s+[0-9.]+\s+CLASS\b|BEGIN\s*$|\s*MultiUse\s*=|END\s*$)"
)

# 実行時に各モジュールへ差し込む、コンパイル成功を確認するための関数。
# LibreOffice Basic はコンパイルエラーをダイアログではなく「関数が Void を返す」
# 形で表面化させることがあるため、モジュール単位で明示的に叩いて確認する。
PROBE_NAME = "ZZ_CompileProbe"
PROBE_OK = "COMPILED"
_PROBE_SRC = (
    "\n\nPublic Function {name}() As String\n"
    '    {name} = "{ok}"\n'
    "End Function\n"
).format(name=PROBE_NAME, ok=PROBE_OK)


# Declare 文の PtrSafe。Declare の直後に来る (Private/Public Declare PtrSafe Function ...)
_PTRSAFE = re.compile(r"^(\s*(?:(?:Private|Public)\s+)?Declare\s+)PtrSafe\s+", re.IGNORECASE | re.MULTILINE)


def strip_ptrsafe(code):
    return _PTRSAFE.sub(r"\1", code)


def strip_vbe_header(code):
    return "\n".join(l for l in code.splitlines() if not _VBE_HEADER.match(l))


def read_source(path):
    """build.py と同じ規則で読む (正は UTF-8、VBE エクスポート由来は cp932)。"""
    raw = pathlib.Path(path).read_bytes()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp932")


def to_basic_module(path, with_probe=True, ptrsafe=False):
    """ptrsafe=True なら Declare 文の PtrSafe を取り除く (本物の Win32 モジュールの構文確認用)。"""
    path = pathlib.Path(path)
    code = strip_vbe_header(read_source(path)).replace("\r\n", "\n")
    if ptrsafe:
        code = strip_ptrsafe(code)
    prologue = "Option VBASupport 1\n"
    if path.suffix.lower() == ".cls":
        # LibreOffice ではクラスモジュールも通常モジュールとして格納されるため、
        # Option ClassModule で明示する必要がある。
        prologue = "Option ClassModule\n" + prologue
    return prologue + code + (_PROBE_SRC if with_probe else "")


def _src_paths():
    return [p for p in sorted(SRC_DIR.iterdir()) if p.suffix.lower() in (".bas", ".cls")]


def fake_names():
    """偽物に差し替えるモジュール名 (tests/fakes/*.bas)。"""
    return sorted(p.stem for p in FAKES_DIR.glob("*.bas"))


def load_src_modules(with_probe=True):
    """テストで動かすモジュールを {モジュール名: Basicソース} で返す。
    tests/fakes/ に偽物があるモジュールは、src/ の本物の代わりに偽物を入れる。"""
    fakes = set(fake_names())
    mods = {}
    for path in _src_paths():
        if path.stem in fakes:
            path = FAKES_DIR / (path.stem + ".bas")
        mods[path.stem] = to_basic_module(path, with_probe=with_probe)
    return mods


def load_real_win32_modules(with_probe=True):
    """偽物に差し替えた本物のモジュールを、PtrSafe を取り除いて返す (構文確認用)。"""
    fakes = set(fake_names())
    return {p.stem: to_basic_module(p, with_probe=with_probe, ptrsafe=True)
            for p in _src_paths() if p.stem in fakes}


_PUBLIC_PROC = re.compile(r"^\s*Public\s+(?:Sub|Function)\s+(\w+)", re.IGNORECASE | re.MULTILINE)


def public_procs(path):
    return set(_PUBLIC_PROC.findall(read_source(path)))


def fake_api_mismatches():
    """本物にあって偽物にない公開手続き {モジュール名: [手続き名]}。"""
    out = {}
    for name in fake_names():
        real = SRC_DIR / (name + ".bas")
        if not real.exists():
            out[name] = ["(src/{}.bas がない)".format(name)]
            continue
        missing = sorted(public_procs(real) - public_procs(FAKES_DIR / (name + ".bas")))
        if missing:
            out[name] = missing
    return out


def load_test_modules(test_dir, with_probe=True):
    mods = {}
    for path in sorted(pathlib.Path(test_dir).glob("*.bas")):
        mods[path.stem] = to_basic_module(path, with_probe=with_probe)
    return mods
