#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pyopenvba",
# ]
# ///

import os
import re
import struct
import subprocess
import zipfile
from pyopenvba import ExcelFile, VBAModuleKind

BUILD_DIR     = os.path.join(os.path.dirname(__file__), "build")
NEW_PATH      = os.path.join(BUILD_DIR, "PERSONAL.xlsb")
SRC_DIR       = os.path.join(os.path.dirname(__file__), "src")

# src/*.bas, *.cls は UTF-8 で保存する（このリポジトリでの正）。
# VBAプロジェクト自体のコードページ（VBE上での文字表示に使われる）は SRC_CODEPAGE で
# 別途指定しており、ソースファイルのディスク上のエンコーディングとは独立している。
# ただし、コンパイル後のVBAプロジェクトへ埋め込む際は最終的に SRC_CODEPAGE
# （cp932 = Shift-JIS）へエンコードされるため、ソースの文字種は cp932 の範囲内に収めること。
SRC_CODEPAGE = 932

# ドキュメントモジュール (テンプレートに最初からあるもの)。src に同名の .cls があれば
# ヘッダーを除いた本文だけを差し替える。
DOCUMENT_MODULES = {"ThisWorkbook", "Sheet1"}

# UserForm のデザイン。コード (src/<名前>.frm) はフォームのコードビハインドで、
# コントロールの配置・プロパティはここで定義する (.frx は使わない)。
# 見た目の大半は Launcher.frm の UserForm_Initialize で上書きされる。
FORMS = {
    "Launcher": {
        # FormUtil.MakeFormTransparent が FindWindow でキャプションから探すため、一意な名前にする
        "caption": "PersonalLauncher",
        "width": 180,
        "height": 120,
        "font": ("MS UI Gothic", 9, 128),
        "controls": [
            {"kind": "TextBox", "name": "TextBox1", "left": 0, "top": 0, "width": 180, "height": 24,
             "props": {"VariousPropertyBits": 746604571}},
            {"kind": "ListBox", "name": "ListBox1", "left": 0, "top": 24, "width": 180, "height": 90,
             "props": {"ScrollBars": 3, "DisplayStyle": 2, "MatchEntry": 0}},
        ],
    },
}

# ランチャーの補完表を埋め込む生成モジュール (src には置かない)
COMP_MODULE = "mLauncherCompData"


def _read_source(path):
    """src/*.bas, *.cls を読み込む。

    正は UTF-8。ただし VBE の「Export File...」はシステムのANSIコードページ
    (日本語環境では cp932) で書き出すため、VBE上で直接編集してエクスポートした
    ファイルを混在させても壊れないよう、UTF-8 で読めない場合は cp932 にフォールバックする。
    """
    with open(path, "rb") as f:
        raw = f.read()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp932")


def _set_project_codepage(project, code_page):
    """VBAプロジェクトのPROJECTCODEPAGEレコードをcode_pageへ書き換える。

    create_new()のテンプレートは欧文Excel由来でcode_page=1252固定。
    project.code_pageはモジュール本文のエンコードにのみ使われ、dirストリームに
    書き込まれるPROJECTCODEPAGEレコード自体はdir_rawの生バイト列がそのまま
    再利用されるため、project.code_pageを変えるだけではファイルに反映されない。
    そのためdir_raw内のレコードを直接書き換える。
    """
    old_record = struct.pack("<HIH", 0x0003, 2, project.code_page)
    new_record = struct.pack("<HIH", 0x0003, 2, code_page)
    if old_record not in project.dir_raw:
        raise RuntimeError("PROJECTCODEPAGE record not found in dir_raw")
    project.dir_raw = project.dir_raw.replace(old_record, new_record, 1)
    project.code_page = code_page
    project.dir_structure_dirty = True


def _check_cp932(fname, source):
    try:
        source.encode("cp932")
    except UnicodeEncodeError as e:
        raise ValueError(
            f"{fname}: cp932(Shift-JIS) で表現できない文字が含まれています "
            f"(位置 {e.start}: {source[e.start:e.start + 1]!r})。"
            "VBAプロジェクトのコードページが cp932 のため、使用文字を絞ってください。"
        ) from e


def _strip_header(source):
    """VERSION/BEGIN...END/Attribute 行を除いた本文を返す (VBE の書き出し形式にも対応)。"""
    lines = source.splitlines(keepends=True)
    i = 0
    if i < len(lines) and lines[i].startswith("VERSION "):
        i += 1
        if i < len(lines) and lines[i].strip().upper() == "BEGIN":
            while i < len(lines) and lines[i].strip().upper() != "END":
                i += 1
            i += 1
    while i < len(lines) and lines[i].startswith("Attribute "):
        i += 1
    return "".join(lines[i:])


def _add_form(wb, name, spec, source):
    form = wb.add_form(name, caption=spec["caption"], width=spec["width"], height=spec["height"])
    # add_form はキャプションをヘッダー (\x03VBFrame) とフォームのレコードの両方に書く。
    # レコード側の Caption は MSForms がフォーム内に枠付きの見出しとして描いてしまうため消す
    # (VBE で作ったフォームと同じく、ウィンドウのタイトルはヘッダー側だけが持つ)。
    form.set_property("Caption", None)
    font_name, font_size, charset = spec["font"]
    for c in spec["controls"]:
        ctl = form.add_control(c["kind"], c["name"], left=c["left"], top=c["top"],
                               width=c["width"], height=c["height"])
        ctl.set_font(font_name, font_size, charset=charset)
        for key, value in c["props"].items():
            ctl.set_property(key, value)
    # 本文だけを渡し、add_form が作ったヘッダー (VB_Base など) を残す
    wb.set_module(name, _strip_header(source))


# --- ランチャーの補完表 -------------------------------------------------------
#
# 各モジュールのヘッダーコメントにある COMPLETIONS 節から作る。書式:
#
#     ' COMPLETIONS
#     '     bg              セル/図形の塗り色
#     '     bg copy         色をコピー
#
# 左側は空白区切りのトークン列 (最後が候補、それより前が文脈)、2つ以上の空白を
# 挟んで右側が説明。空のコメント行 (') か罫線 ('===) で節が終わる。

_COMP_HEADING = re.compile(r"^'\s*COMPLETIONS\s*$")
_COMP_ENTRY = re.compile(r"^'\s+(\S.*?)\s*$")
_SUB_DEF = re.compile(r"^\s*(?:Public\s+)?Sub\s+(\w+)", re.IGNORECASE | re.MULTILINE)


def _parse_completions(fname, source):
    entries = []
    in_section = False
    for line in source.splitlines():
        if _COMP_HEADING.match(line):
            in_section = True
            continue
        if not in_section:
            continue
        m = _COMP_ENTRY.match(line)
        if not m or m.group(1).startswith("="):
            in_section = False
            continue
        parts = re.split(r"\s{2,}", m.group(1), maxsplit=1)
        desc = parts[1] if len(parts) > 1 else ""
        entries.append((fname, parts[0].split(), desc))
    return entries


def _completion_module(entries, sub_names):
    """補完表を返す VBA モジュールのソースを作る。トップレベルの候補は実在する Sub か検査する。"""
    by_ctx = {"": []}
    seen = set()
    for fname, path, desc in entries:
        key = tuple(path)
        if key in seen:
            raise ValueError(f"{fname}: COMPLETIONS の項目が重複しています: {' '.join(path)}")
        seen.add(key)
        if path[0].lower() not in sub_names:
            raise ValueError(f"{fname}: COMPLETIONS の {path[0]!r} に対応する Sub がありません")
        by_ctx.setdefault(" ".join(path[:-1]), []).append((path[-1], desc))

    by_ctx[""].sort(key=lambda item: item[0])  # トップレベルはコマンド名順、サブコマンドは記述順
    lines = []
    for ctx, items in by_ctx.items():
        if ctx:
            lines.append(f"[{ctx}]")
        lines += [f"{tok} : {desc}" if desc else tok for tok, desc in items]

    out = [
        f'Attribute VB_Name = "{COMP_MODULE}"',
        "Option Explicit",
        "Option Private Module",
        "' build.py が src/*.bas の COMPLETIONS 節から生成する。直接編集しないこと。",
        "",
        "Public Function CompletionText() As String",
        "    Dim s As String",
    ]
    out += ['    s = s & "' + line.replace('"', '""') + '" & vbLf' for line in lines]
    out += ["    CompletionText = s", "End Function", ""]
    return "\r\n".join(out)


def build_new(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path):
        os.remove(path)
    with ExcelFile.create_new(path) as wb:
        project = wb.vba_project()
        _set_project_codepage(project, SRC_CODEPAGE)
        src_names = set()
        comp_entries = []
        sub_names = set()
        for fname in sorted(os.listdir(SRC_DIR)):
            name, ext = os.path.splitext(fname)
            ext = ext.lower()
            if ext not in (".bas", ".cls", ".frm"):
                continue
            source = _read_source(os.path.join(SRC_DIR, fname))
            _check_cp932(fname, source)
            if ext == ".bas":
                comp_entries += _parse_completions(fname, source)
                sub_names |= {n.lower() for n in _SUB_DEF.findall(source)}

            if ext == ".frm":
                if name not in FORMS:
                    raise ValueError(f"{fname}: build.py の FORMS にフォームの定義がありません")
                _add_form(wb, name, FORMS[name], source)
            elif name in DOCUMENT_MODULES:
                wb.set_module(name, _strip_header(source))
            else:
                kind = VBAModuleKind.standard if ext == ".bas" else VBAModuleKind.other
                try:
                    wb.set_module(name, source)
                except KeyError:
                    project.add_module(name, source, kind=kind)
            src_names.add(name)

        project.add_module(COMP_MODULE, _completion_module(comp_entries, sub_names),
                           kind=VBAModuleKind.standard)

        # テンプレート同梱の"Module1"がsrcに存在しない場合は空のまま残るため削除
        if "Module1" not in src_names:
            try:
                project.delete_module("Module1")
            except KeyError:
                pass

        wb.save()

    _hide_workbook_window(path)


# BrtBookView (MS-XLSB 2.4.304) のレコード種別と、fHidden フラグの位置
_BRT_BOOK_VIEW = 158
_BOOK_VIEW_FLAGS_OFFSET = 28  # xWn..itabCur (4byte x 7) の直後の 1byte
_BOOK_VIEW_FHIDDEN = 0x01


def _read_varint(buf, i, max_bytes):
    value = 0
    for n in range(max_bytes):
        b = buf[i]
        i += 1
        value |= (b & 0x7F) << (7 * n)
        if not b & 0x80:
            break
    return value, i


def _hide_workbook_window(path):
    """xl/workbook.bin の全 BrtBookView に fHidden を立て、ウィンドウを非表示にする。

    create_new() のテンプレートは通常のブック (ウィンドウ表示) なので、
    そのまま PERSONAL.XLSB にすると Excel 起動時に表に出てしまう。
    """
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        data = {info.filename: zin.read(info) for info in infos}

    buf = bytearray(data["xl/workbook.bin"])
    i = 0
    patched = 0
    while i < len(buf):
        rec_type, i = _read_varint(buf, i, 2)
        rec_size, i = _read_varint(buf, i, 4)
        if rec_type == _BRT_BOOK_VIEW and rec_size > _BOOK_VIEW_FLAGS_OFFSET:
            buf[i + _BOOK_VIEW_FLAGS_OFFSET] |= _BOOK_VIEW_FHIDDEN
            patched += 1
        i += rec_size
    if patched == 0:
        raise RuntimeError("BrtBookView record not found in xl/workbook.bin")
    data["xl/workbook.bin"] = bytes(buf)

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as zout:
        for info in infos:
            zout.writestr(info, data[info.filename])
    os.replace(tmp, path)


def _appdata_dir():
    """%APPDATA% のディレクトリを、実行中の Python から使えるパスで返す。

    Windows ではそのまま環境変数を使う。WSL では Windows の環境変数が
    引き継がれない (PATH と WSLENV 指定分のみ) ため、cmd.exe から取得して
    wslpath で /mnt/c/... 形式へ変換する。
    """
    if os.name == "nt":
        return os.environ["APPDATA"]
    if "APPDATA" in os.environ and os.path.isdir(os.environ["APPDATA"]):
        return os.environ["APPDATA"]  # WSLENV=APPDATA/p で渡された場合など
    raw = subprocess.run(
        ["cmd.exe", "/d", "/c", "echo %APPDATA%"],
        capture_output=True, check=True, cwd="/mnt/c",
    ).stdout
    try:
        win_path = raw.decode("cp932").strip()
    except UnicodeDecodeError:
        win_path = raw.decode("utf-8").strip()
    if not win_path or win_path == "%APPDATA%":
        raise RuntimeError("Windows 側の APPDATA を取得できませんでした")
    return subprocess.run(
        ["wslpath", "-u", win_path], capture_output=True, check=True, text=True,
    ).stdout.strip()


def personal_path():
    return os.path.join(_appdata_dir(), "Microsoft", "Excel", "XLSTART", "PERSONAL.XLSB")


def deploy(new_path, target_path):
    """新ファイルを本番パスへ移動する（既存ファイルは上書き）。"""
    os.replace(new_path, target_path)


def main():
    target = personal_path()
    build_new(NEW_PATH)
    deploy(NEW_PATH, target)


if __name__ == "__main__":
    main()
