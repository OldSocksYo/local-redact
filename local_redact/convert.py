"""用本机 WPS 或 Microsoft Word 把旧版 .doc 转成临时 .docx。不改原稿。"""

from __future__ import annotations

from pathlib import Path

# 另存为 Office Open XML（.docx），不保留宏。
_WD_FORMAT_XML_DOCUMENT = 16
_WD_ALERTS_NONE = 0
_WD_DO_NOT_SAVE_CHANGES = 0
# 打开旧文档时不执行宏。
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
# WPS 有时会把 Word.Application 登记成会卡住的入口，所以先试 KWPS.Application。
_PROG_IDS = ("KWPS.Application", "Word.Application")


class DocConversionError(Exception):
    pass


class WordSession:
    """单独开一个 WPS 或 Word 进程，避免关掉用户正在用的窗口。"""

    def __init__(self) -> None:
        self._word = None
        self._com_ready = False

    def start(self) -> None:
        try:
            import pythoncom
            import win32com.client  # noqa: F401
        except ImportError as exc:
            raise DocConversionError("缺少 pywin32。请先执行：python -m pip install pywin32") from exc
        pythoncom.CoInitialize()
        self._com_ready = True
        try:
            word = _dispatch_office()
        except DocConversionError:
            self.close()
            raise
        except Exception as exc:
            self.close()
            raise DocConversionError("没有检测到可用的 Word 或 WPS，无法转换 .doc。") from exc
        word.Visible = False
        word.DisplayAlerts = _WD_ALERTS_NONE
        try:
            word.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        self._word = word

    def convert(self, src: Path, dest: Path) -> None:
        if self._word is None:
            raise DocConversionError("Word 尚未启动。")
        document = None
        try:
            document = self._word.Documents.Open(
                str(src.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
                NoEncodingDialog=True,
            )
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                dest.unlink()
            document.SaveAs2(str(dest.resolve()), FileFormat=_WD_FORMAT_XML_DOCUMENT)
            document.Saved = True
        except DocConversionError:
            raise
        except Exception as exc:
            raise DocConversionError("无法转成 docx，文件可能已加密或已损坏。") from exc
        finally:
            if document is not None:
                try:
                    document.Close(SaveChanges=_WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass

    def close(self) -> None:
        if self._word is not None:
            try:
                self._word.DisplayAlerts = _WD_ALERTS_NONE
                self._word.Quit()
            except Exception:
                pass
            self._word = None
        if self._com_ready:
            import pythoncom

            pythoncom.CoUninitialize()
            self._com_ready = False


def _dispatch_office():
    import win32com.client

    for prog_id in _PROG_IDS:
        server = _local_server(prog_id).lower()
        # 这个入口在本机指向会卡住的 wps.exe /Automation，前面已经试过 KWPS。
        if prog_id == "Word.Application" and "wps.exe" in server:
            continue
        try:
            return win32com.client.DispatchEx(prog_id)
        except Exception:
            continue
    raise DocConversionError("没有检测到可用的 Word 或 WPS，无法转换 .doc。")


def _local_server(prog_id: str) -> str:
    import winreg

    clsid = _registry_value(winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\{prog_id}\CLSID")
    if not clsid:
        clsid = _registry_value(winreg.HKEY_CLASSES_ROOT, rf"{prog_id}\CLSID")
    if not clsid:
        return ""
    candidates = (
        (winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\CLSID\{clsid}\LocalServer32"),
        (winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Classes\WOW6432Node\CLSID\{clsid}\LocalServer32"),
        (winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32"),
        (winreg.HKEY_CLASSES_ROOT, rf"WOW6432Node\CLSID\{clsid}\LocalServer32"),
    )
    for root, path in candidates:
        value = _registry_value(root, path)
        if value:
            return value
    return ""


def _registry_value(root, path: str) -> str:
    import winreg

    try:
        with winreg.OpenKey(root, path) as key:
            value, _ = winreg.QueryValueEx(key, "")
    except OSError:
        return ""
    return str(value)
