from contextlib import suppress
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Union, List, Any, get_type_hints

import minimalmodbus
minimalmodbus.CLOSE_PORT_AFTER_EACH_CALL = False

from aa_py_core.xl.context import excel
import pandas as pd
from pathlib import Path
import xlwings as xw
from aa_py_core.xl.tables import LOTable
from pandas import Series
from serial import SerialException
from win32com.client import CDispatch

__dirpath__ = Path(globals().get("__file__", "./_")).absolute().parent

from misc import namify, force_int, str2bits, timeStrober, bits2int, try_n, is_nan
from scale_device import ScaleDevice


def wb_to_xw(wb):
    """
    Convert a CDispatch workbook object into an xlwings workbook object
    """
    return xw.books(wb.name)


def get_table_as_df(book: Union[xw.Book, str, CDispatch],
                    tablename: str) -> pd.DataFrame:
    """
    >>> with excel(__dirpath__.joinpath('test', 'tables.xlsx'), quiet=True, kill=True) as book:
    ...      list(get_table_as_df(book, "test_table").columns)
    ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']

    """
    if isinstance(book, CDispatch):
        book =  wb_to_xw(book)

    elif isinstance(book, str):
        book = xw.books[book]

    lo_table = LOTable.get_from_book(book, tablename)
    lo_data = lo_table.extract_data()
    df = pd.DataFrame(lo_data)
    df.columns = [namify(i) for i in df.columns]
    return df


@dataclass
class DeviceInfo:
    name: str
    com: str
    device: minimalmodbus.Instrument
    line: SimpleNamespace

    def read(self):
        if not is_nan(self.line.no_of_registers) and self.line.no_of_registers != 0:

            def do_read():
                # Reading a unsigned int normally
                if self.line.datatype == 'uint':
                    first = self.line.start_register
                    last = self.line.start_register + self.line.no_of_registers
                    values = []
                    for i in range(first, last, 8):
                        values.extend(self.device.read_registers(i, min(i+8, last)-i))
                    return values

                # Reading a unsigned int as a list of bits
                elif self.line.datatype == "char bits":
                    values = []
                    for i in range(force_int(self.line.start_register),
                                   force_int(self.line.start_register+self.line.no_of_registers)):
                        values.extend(
                            str2bits(self.device.read_string(i, 1))[::-1]
                        )
                    return values

                else:
                    ValueError('Datatype must be either "uint" or "char bits"')

            return try_n(do_read, tries=4)

    def write_bits(self,
                   strobe_settings: List[str],
                   register: int):
        strobes = [timeStrober(s) for s in strobe_settings]
        int_value = bits2int([s.is_on() for s in strobes][::-1])

        try_n(
            lambda: self.device.write_register(register, int_value),
            tries=4
        )

    def write(self,
              value,
              register):

        try_n(
            lambda: self.device.write_register(register, value),
            tries=4
        )

    def output_to_excel(self, sheet, line_number):
        if (vals := self.read()) is not None:
            # Add line number to dump_cols
            dump_cols = self.line.dump_cols
            dump_cols = ':'.join([f'{i}{line_number}' for i in dump_cols.split(':')])

            def do_output():
                range = sheet.Range(dump_cols)
                range.Value = vals[:len(range)]


            try_n(
                do_output,
                tries = 50
            )


class DeviceInfoScale(DeviceInfo):
    device: ScaleDevice
    def read(self):
        return [self.device.mass]


deviceinfo_subclasses = {}
for key, val in list(globals().items()):
    if key.startswith('DeviceInfo') and key != 'DeviceInfo':
        deviceinfo_subclasses[key[len('DeviceInfo'):].lower()] = val


def get_devices(device_info: pd.DataFrame) -> List[DeviceInfo]:
    """
    >>> import misc
    >>> import xlwings as xw
    >>> wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet() # doctest: +ELLIPSIS
    T...
    Y...
    >>> devices = get_devices_from_book(xw.books(wb.Name))
    """

    devices = []
    failures = []
    for line in [i for i in device_info.iloc if i.active and not is_nan(i.active) and str(i).strip() != '']:

        line = SimpleNamespace(**line)

        if line.device_name.lower() in deviceinfo_subclasses:
            device_class = deviceinfo_subclasses[line.device_name.lower()]
            device_instance = get_type_hints(device_class)['device']()
            devices.append(
                device_class(
                    name=line.device_name, com='Unknown', device=device_instance, line=line
                )
            )


        else:
            for col in ['address', 'start_register', 'baud', 'no_of_bits']:
                try:
                    line.__dict__[col] = force_int(line.__dict__[col])
                except ValueError:
                    raise ValueError(f"Invalid {col} for {line.device_name}: {line}")

            line.start_register = force_int(line.start_register)
            with suppress(ValueError):
                line.no_of_registers = force_int(line.no_of_registers)

            line.parity = str(line.parity).upper()
            line.communication_format = str(line.communication_format).lower()

            found_it = False
            for com in [f"COM{i}" for i in range(1 ,12)]:
                try:
                    dev = minimalmodbus.Instrument(com, line.address, mode=line.communication_format)
                    dev.serial.stopbits = line.stop
                    dev.serial.parity = line.parity
                    dev.serial.bytesizes = line.no_of_bits
                    dev.serial.baudrate = line.baud

                except SerialException:
                    continue

                for f in (dev.read_bits, dev.read_string):
                    try:
                        f(line.start_register, 1)
                        found_it = True

                        devices.append(
                            DeviceInfo(name=line.device_name, com=com, device=dev, line=line)
                        )
                        break

                    except (minimalmodbus.NoResponseError, minimalmodbus.InvalidResponseError, minimalmodbus.IllegalRequestError) :
                        pass

                if found_it:
                    break

            if not found_it:
                failures.append(str(line.device_name))

    if failures:

        raise ConnectionError(f"Could not read devices: {failures}")

    return devices


def get_devices_from_book(wb: Union[xw.Book, str, CDispatch]) -> List[DeviceInfo]:
    return get_devices(
        get_table_as_df(wb, 'device_info')
    )

