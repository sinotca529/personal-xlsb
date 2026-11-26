"""src/*.bas, *.cls を LibreOffice Basic に読み込める形へ変換する。

VBA の .bas / .cls は、そのままでは LibreOffice Basic のモジュールとして
読み込めない (VBE 固有の Attribute 行などが含まれるため)。ここでは

  * VBE 固有のヘッダ行 (Attribute VB_*, VERSION 1.0 CLASS, BEGIN/END/MultiUse) を除去
  * .cls には Option ClassModule を付与 (LibreOffice でクラスモジュールとして扱わせる)
  * 全モジュールに Option VBASupport 1 を付与 (Excel の VBA 方言・定数を有効化)
  * Declare 文の PtrSafe を除去 (LibreOffice Basic は PtrSafe を解釈できず、
    #If で分岐させてもモジュールごとコンパイルに失敗する。Declare と LongPtr は通る)

だけを行う。それ以外のコードの本文には手を入れない。Win32 API は Linux では
呼べないので、Declare した関数を呼ぶ経路はテストしない。
"""

import pathlib
import re

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
SRC_DIR = REPO_ROOT / "src"

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


def to_basic_module(path, with_probe=True):
    path = pathlib.Path(path)
    code = strip_ptrsafe(strip_vbe_header(read_source(path)).replace("\r\n", "\n"))
    prologue = "Option VBASupport 1\n"
    if path.suffix.lower() == ".cls":
        # LibreOffice ではクラスモジュールも通常モジュールとして格納されるため、
        # Option ClassModule で明示する必要がある。
        prologue = "Option ClassModule\n" + prologue
    return prologue + code + (_PROBE_SRC if with_probe else "")


def load_src_modules(with_probe=True):
    """src/ 配下の全モジュールを {モジュール名: Basicソース} で返す。"""
    mods = {}
    for path in sorted(SRC_DIR.iterdir()):
        if path.suffix.lower() not in (".bas", ".cls"):
            continue
        mods[path.stem] = to_basic_module(path, with_probe=with_probe)
    return mods


def load_test_modules(test_dir, with_probe=True):
    mods = {}
    for path in sorted(pathlib.Path(test_dir).glob("*.bas")):
        mods[path.stem] = to_basic_module(path, with_probe=with_probe)
    return mods
