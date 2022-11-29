import subprocess
import sys
import time
from contextlib import suppress
import os
from textwrap import dedent
import numpy as np
import pythoncom
import win32com.client
import locate
from pathlib import Path
from win32com.universal import com_error
import hashlib

import win32con
import win32file
import win32api

import shutil
import threading

with locate.prepend_sys_path("../site-packages"):
    from aa_py_core.processes import kill_pid

alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
num2col = [i for i in alph] + [i + j for i in alph for j in alph]
col2num = {j: i for i, j in enumerate(num2col)}

numeric = set("1234567890")
alphanumeric = set("1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")


def try_n(f, *args, errors_list=None, tries=3, **kwargs):
    """
    Try to run a function n times before giving up

    >>> count = {'i':0}
    >>> # noinspection PyUnresolvedReferences
    ... def func():
    ...     count['i'] += 1
    ...     assert count['i'] >= 4
    ...     return 9

    >>> try_n(func, tries=3)
    Traceback (most recent call last):
    ...
    AssertionError

    >>> try_n(func, tries=4)
    9

    >>> try_n(func, tries=5)
    9

    """
    if errors_list is None:
        errors_list = Exception
    else:
        errors_list = tuple(errors_list)

    for i in range(tries - 1):
        try:
            return f(*args, **kwargs)
        except errors_list:
            time.sleep(0.05)

    return f(*args, **kwargs)


def try_thrice(f, *args, errors_list=None, **kwargs):
    """
    Try to run a function three times before giving up

    >>> count = {'i':0}
    >>> # noinspection PyUnresolvedReferences
    ... def func():
    ...     count['i'] += 1
    ...     assert count['i'] >= 3

    >>> try_thrice(func)

    """
    return try_n(f, *args, errors_list=errors_list, tries=3, **kwargs)


def get_named_range(wb, named_range):
    """
    Search all the Excel sheets for a named-range value
    """
    range = None
    for sheet in wb.Sheets:
        with suppress(com_error):
            range = sheet.Range(named_range)

    if range is None:
        raise ValueError(f"Range {named_range} cannot be found.")

    return range


def namify(s, replacer="_"):
    """
    Convert a piece of text to a valid table name

    """
    s = str(s).lower()
    if s[:1] in numeric:
        s = "_" + s
    s = "".join([i if i in alphanumeric else replacer for i in s])
    return s


def spread_iterator():
    for moniker in pythoncom.GetRunningObjectTable():
        try:
            # Workbook implements IOleWindow so only consider objects implementing that
            window = moniker.BindToObject(
                pythoncom.CreateBindCtx(0), None, pythoncom.IID_IOleWindow
            )
            disp = window.QueryInterface(pythoncom.IID_IDispatch)

            # Get a win32com Dispatch object from the PyIDispatch object as it's
            # easier to work with.
            book = win32com.client.Dispatch(disp)

        except pythoncom.com_error:
            # Skip any objects we're not interested in
            continue

        try:
            book.Sheets(1)  # Object is a book with sheets
        except:
            continue

        bookname = moniker.GetDisplayName(pythoncom.CreateBindCtx(0), None)

        yield bookname, book


def get_ignis_spreadsheet():
    for bookname, book in spread_iterator():
        print("Test workbook: ", bookname)

        inputs = [i for i in book.Sheets if i.Name.lower() == "inputs"]
        outputs = [i for i in book.Sheets if i.Name.lower() == "outputs"]

        if len(inputs) and len(outputs):
            print("Yes -->", bookname)
            return book, bookname, inputs[0], outputs[0]

    raise ValueError("No open spreadsheet found with inputs and outputs sheets")


def force_int(s):
    s = str(s)
    if "." in s:
        val = float(s)
    else:
        val = int(s, 0)

    return int(val)


def str2bits(s):
    result = []
    for c in s:
        bits = bin(ord(c))[2:]
        bits = "00000000"[len(bits) :] + bits
        result.extend([int(b) for b in bits])
    return result


def bits2str(bits):
    bitgroups = [bits[i : i + 8] for i in range(0, len(bits), 8)]
    int_list = []
    for bit in bitgroups:
        int_list.append(0)
        for i, val in enumerate(bit[::-1]):
            int_list[-1] += (2 ** (i)) * bool(val)
    str_out = "".join([chr(i) for i in int_list])
    return str_out


def bits2int(bits):
    out_int = 0
    for i, val in enumerate(bits[::-1]):
        out_int += (2 ** (i)) * bool(val)
    return out_int


class TimeStrober:
    def __init__(self, inpstr):
        self.set_timings(inpstr)

    def set_timings(self, inpstr):
        self.as_absolute_value = None

        try:
            inpstr * 1.0
        except:
            pass
        else:
            self.as_absolute_value = inpstr
            inpstr = "on" if inpstr else "off"

        inpstr = inpstr.lower().strip()

        if inpstr.startswith("t"):
            self.pperiod, self.pwidth = time.time(), float(inpstr[1:].strip())

        elif inpstr.startswith("s"):
            self.pperiod, self.pwidth = (
                float(inpstr[1:].split(",")[0].strip()),
                float(inpstr[1:].split(",")[1].strip()),
            )
            if self.pperiod == 0:
                self.pperiod += 0.05

        elif inpstr == "on":
            self.pperiod, self.pwidth = time.time(), time.time()

        elif inpstr == "off":
            self.pperiod, self.pwidth = time.time(), 0

        else:
            raise ValueError(
                "Error timestrobe inputs expected: "
                f'1/0/True/False/"on"/"off"/"t<seconds>"/"s<period>,<pulsewidth>", got {inpstr}.'
            )

    def get_value(self):
        if self.as_absolute_value is not None:
            return self.as_absolute_value

        now = time.time()
        if now - int(now / self.pperiod) * self.pperiod < self.pwidth:
            return True
        else:
            return False


def is_nan(val):
    """
    More extensive is_nan test than numpy's
    """
    try:
        return np.isnan(val)
    except TypeError:
        if val is None:
            return True
        else:
            return False


def exit_after_n_seconds(n=1):
    kill_pid  # noqa
    sitepack = locate.this_dir().joinpath("site-packages")
    assert sitepack.is_dir()

    pystr = dedent(
        f"""\
        import time;
        import sys;
        sys.path.insert(0, r'{sitepack}';
        time.sleep({n});
        from aa_py_core.processes import kill_pid;
        kill_pid({os.getpid()});
    """
    ).replace("\n", " ")

    subprocess.Popen([sys.executable, "-c", pystr], start_new_session=True)


def backup_workbook_to_highest_portable_drive(path, threaded=False):
    path = Path(path)
    drives = find_all_portable_drives_on_windows()
    if not drives:
        print("No portable drives found for backup.")
        return
    drive = drives[-1]

    uname = hashlib.sha1(path.as_posix().encode()).hexdigest()[:8]
    ext = path.suffix

    backupdir = Path(
        drive, "ignis_backups", Path(path).with_suffix("").name + f"_{uname}"
    )
    backupdir.mkdir(parents=True, exist_ok=True)

    # make sure all files are between 20 minutes and 3 hours
    for f in backupdir.glob("*"):
        if (
            time.time() - f.stat().st_mtime > 20 * 60
            and time.time() - f.stat().st_mtime < 3 * 60 * 60
        ):
            f.unlink()

    backuppath = backupdir.joinpath(
        f"{path.with_suffix('').name}{time.strftime('--%Y%m%d-%H%M%S')}{ext}"
    )
    if threaded:
        threading.Thread(target=shutil.copyfile, args=(path, backuppath)).start()
    else:
        shutil.copyfile(path, backuppath)


def find_all_portable_drives_on_windows():
    drives = win32api.GetLogicalDriveStrings().split("\000")
    drives = sorted([i.strip() for i in drives])
    drives = [
        drive
        for drive in drives
        if win32file.GetDriveType(drive) == win32con.DRIVE_REMOVABLE
    ]
    return drives
