"""ヘッドレス LibreOffice を Excel の代わりの VBA 実行環境として起動する。

LibreOffice Calc のドキュメントに VBACompatibilityMode を立てると、
Basic から Excel の VBA オブジェクトモデル (Application / ActiveSheet /
Range / Interior / Borders など) と Excel 定数 (xlNone, xlEdgeTop など) が
使えるようになる。これを利用して、Excel 無し・GUI 無しで VBA を実行する。
"""

import atexit
import pathlib
import subprocess
import time

import uno  # LibreOffice 同梱の Python バインディング (python3-uno)
from com.sun.star.connection import NoConnectException

_HERE = pathlib.Path(__file__).resolve().parent
DEFAULT_PROFILE = _HERE / ".loprofile"
DEFAULT_PORT = 2002
LIB_NAME = "VbaTest"


class LibreOffice:
    def __init__(self, port=DEFAULT_PORT, profile=DEFAULT_PROFILE, soffice="soffice"):
        self.port = port
        self.profile = pathlib.Path(profile)
        self.soffice = soffice
        self.proc = None
        self.remote = None

    def start(self, timeout=90):
        self.proc = subprocess.Popen(
            [
                self.soffice,
                "--headless",
                "--invisible",
                "--nologo",
                "--norestore",
                "--nolockcheck",
                "-env:UserInstallation=file://{}".format(self.profile),
                "--accept=socket,host=127.0.0.1,port={};urp;StarOffice.ServiceManager".format(self.port),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        atexit.register(self.stop)
        ctx = uno.getComponentContext()
        resolver = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", ctx
        )
        url = "uno:socket,host=127.0.0.1,port={};urp;StarOffice.ComponentContext".format(self.port)
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                self.remote = resolver.resolve(url)
                return self
            except NoConnectException:
                time.sleep(0.3)
        raise RuntimeError("LibreOffice への接続がタイムアウトしました")

    def stop(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            try:
                self.proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                self.proc.kill()
        self.proc = None

    def new_workbook(self):
        """VBA 互換モードの Calc ドキュメントを 1 つ開く。

        Basic ライブラリはドキュメントごとに 1 回だけ入れ直すこと。
        同じドキュメントに再インストールすると古いコンパイル結果が残り、
        テスト結果が汚染される。
        """
        desktop = self.remote.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.remote
        )
        # NOTE: Hidden=True で開くと ViewShell が作られず、ActiveCell / Selection を
        # 使うマクロが実行時エラー91 (Object variable not set) になる。
        # --headless 起動なのでウィンドウは出ないため、Hidden は指定しない。
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        doc.setPropertyValue("VBACompatibilityMode", True)
        return Workbook(doc)


class Workbook:
    def __init__(self, doc):
        self.doc = doc

    def install(self, modules, lib_name=LIB_NAME):
        libs = self.doc.BasicLibraries
        if libs.hasByName(lib_name):
            libs.removeLibrary(lib_name)
        libs.createLibrary(lib_name)
        lib = libs.getByName(lib_name)
        for name, code in modules.items():
            lib.insertByName(name, code)

    def call(self, module, func, lib_name=LIB_NAME):
        url = "vnd.sun.star.script:{}.{}.{}?language=Basic&location=document".format(
            lib_name, module, func
        )
        return self.doc.getScriptProvider().getScript(url).invoke((), (), ())[0]

    def sheet(self, index=0):
        return self.doc.Sheets.getByIndex(index)

    def close(self):
        try:
            self.doc.close(False)
        except Exception:
            pass
