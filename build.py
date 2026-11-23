#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pyopenvba",
# ]
# ///

import os
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


def build_new(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if os.path.exists(path):
        os.remove(path)
    with ExcelFile.create_new(path) as wb:
        project = wb.vba_project()
        _set_project_codepage(project, SRC_CODEPAGE)
        src_names = set()
        for fname in os.listdir(SRC_DIR):
            name, ext = os.path.splitext(fname)
            if ext.lower() not in (".bas", ".cls"):
                continue
            source = _read_source(os.path.join(SRC_DIR, fname))
            try:
                source.encode("cp932")
            except UnicodeEncodeError as e:
                raise ValueError(
                    f"{fname}: cp932(Shift-JIS) で表現できない文字が含まれています "
                    f"(位置 {e.start}: {source[e.start:e.start + 1]!r})。"
                    "VBAプロジェクトのコードページが cp932 のため、使用文字を絞ってください。"
                ) from e
            kind = VBAModuleKind.standard if ext.lower() == ".bas" else VBAModuleKind.other
            try:
                wb.set_module(name, source)
            except KeyError:
                project.add_module(name, source, kind=kind)
            src_names.add(name)

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
