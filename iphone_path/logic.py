"""Browse iPhone photos exposed by Windows Portable Devices (MTP/WPD)."""

import io
import os
import tempfile
import time
from datetime import datetime
from urllib.parse import quote, unquote

from xefm.path_base import ReadOnlyPathImpl, UriStatResult

THIS_PC = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
_COPY_TIMEOUT = 180


def _connected_iphones(folder):
    found = []
    for item in folder.Items():
        if not item.IsFolder:
            continue
        label = str(item.Name)
        lowered = label.casefold()
        if "iphone" not in lowered:
            continue
        found.append(item)
    return found


def _items(shell, item):
    return _folder(item).Items()


def _folder(item):
    import pythoncom
    import win32com.client

    return win32com.client.Dispatch(item.GetFolder)


def _find(items, name):
    for item in items:
        if str(item.Name) == name:
            return item
    return None


def _resolve(shell, parts):
    folder = shell.Namespace(THIS_PC)
    if not parts:
        return None, folder
    item = _find(_connected_iphones(folder), parts[0])
    if item is None:
        return None, None
    for name in parts[1:]:
        if not item.IsFolder:
            return None, None
        folder = _folder(item)
        item = _find(folder.Items(), name)
        if item is None:
            return None, None
    return item, folder


def decode_parts(key):
    return [unquote(part) for part in key.split("/") if part]


def encode_component(name):
    return quote(str(name), safe="")


def _timestamp(item):
    value = item.ExtendedProperty("System.DateModified")
    if not value:
        value = item.ModifyDate
    if isinstance(value, datetime):
        return value.timestamp()
    return 0.0


class IPhonePathImpl(ReadOnlyPathImpl):
    """Read-only ``iphone://`` view of the Apple device visible in This PC."""

    SCHEME = "iphone"
    READ_ONLY_MESSAGE = "iPhone は読み取り専用です。コピー先を PC 側に指定してください"

    def _parts(self):
        return decode_parts(self.key)

    @property
    def name(self):
        return unquote(super().name)

    def _with_item(self, callback):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            if not self.key:
                return callback(shell, None, shell.Namespace(THIS_PC))
            item, parent = _resolve(shell, self._parts())
            return callback(shell, item, parent)
        finally:
            pythoncom.CoUninitialize()

    def exists(self):
        return self._with_item(
            lambda _shell, item, _parent: True if item is None and not self.key else item is not None)

    def is_dir(self):
        return self._with_item(
            lambda _shell, item, _parent: True if item is None and not self.key else bool(item and item.IsFolder))

    def iterdir(self):
        def list_items(shell, item, folder):
            folder = folder if item is None else (_folder(item) if item.IsFolder else None)
            if folder is None:
                raise NotADirectoryError(str(self))
            source = _connected_iphones(folder) if item is None else list(folder.Items())
            for child in source:
                yield self.child(encode_component(child.Name))

        # Keep the COM apartment alive while the lazy iterator is consumed.
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            if self.key:
                item, parent = _resolve(shell, self._parts())
            else:
                item, parent = None, shell.Namespace(THIS_PC)
            yield from list_items(shell, item, parent)
        finally:
            pythoncom.CoUninitialize()

    def stat(self):
        def get_stat(_shell, item, _parent):
            if item is None and not self.key:
                return UriStatResult(is_dir=True)
            if item is None:
                raise FileNotFoundError(str(self))
            is_dir = bool(item.IsFolder)
            try:
                size = int(item.ExtendedProperty("System.Size") or item.Size) if not is_dir else 0
            except (TypeError, ValueError):
                size = 0
            return UriStatResult(size=size, mtime=_timestamp(item), is_dir=is_dir)

        return self._with_item(get_stat)

    def open(self, mode="rb", buffering=-1, encoding=None, errors=None, newline=None):
        if mode not in ("r", "rb"):
            self._refuse()

        def read_item(shell, item, _parent):
            if item is None or item.IsFolder:
                raise FileNotFoundError(str(self))
            with tempfile.TemporaryDirectory(prefix="xefm-iphone-") as temp:
                destination = os.path.join(temp, str(item.Name))
                target_folder = shell.Namespace(temp)
                if target_folder is None:
                    raise OSError(f"一時フォルダを開けません: {temp}")
                target_folder.CopyHere(item, 20)  # no progress UI or overwrite prompt
                deadline = time.monotonic() + _COPY_TIMEOUT
                previous = -1
                stable = 0
                while time.monotonic() < deadline:
                    if os.path.isfile(destination):
                        size = os.path.getsize(destination)
                        if size > 0 and size == previous:
                            stable += 1
                            if stable >= 3:
                                with open(destination, "rb") as stream:
                                    data = stream.read()
                                stream = io.BytesIO(data)
                                if mode == "r":
                                    return io.TextIOWrapper(stream, encoding=encoding or "utf-8",
                                                            errors=errors, newline=newline)
                                return stream
                        else:
                            stable = 0
                        previous = size
                    time.sleep(0.5)
            raise TimeoutError(f"iPhoneからの読み込みが完了しませんでした: {item.Name}")

        return self._with_item(read_item)


IPHONE_SCHEME = {"class": IPhonePathImpl, "override": True}
