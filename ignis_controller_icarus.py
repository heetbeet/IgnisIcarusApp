import os
import subprocess
import sys
from types import SimpleNamespace
import xlwings as xw
from pathlib import Path
from aa_py_core.util import kill_pid
from win32com.universal import com_error

from datetime import datetime
import misc
import time
from device_information import  get_devices_from_book
from misc import is_nan
import numpy as np
import traceback
import crccheck

modb = crccheck.crc.CrcModbus()


def relay_crc(x):
    return modb.calcbytes(x)[::-1]


def get_harcoded_parameters(wb):
    return SimpleNamespace(
        reading_interval = misc.force_int(wb.Sheets['Parameters'].Range("reading_interval").Value),
    )


def update_write_values(wb, devices):
    '''
    Don't overtax this function
    '''

    # If it is a named range, search everywhere, if it
    def get_value_from_cell_reference(source_val):
        try:
            return wb.Sheets['Outputs'].Range(source_val).Value
        except com_error:
            return misc.get_named_range(wb, source_val).Value

    devices_mapping = []
    for d in devices:
        values_i = {}
        for i in range(1, 1+8):
            source = f'source_{i}'
            dest = f'write_{i}'

            source_val = lnk = d.line.__dict__[source]
            if not is_nan(lnk):
                source_val = get_value_from_cell_reference(lnk)

            dest_val = d.line.__dict__[dest]

            if str(source_val).lower().startswith("0x"):
                values_i[i] = source_val

            elif (not is_nan(source_val) and not is_nan(dest_val)):

                dest_val = misc.force_int(dest_val)
                values = get_value_from_cell_reference(source_val)

                if isinstance(values, tuple) or isinstance(values, list):
                    values_i[dest_val] = [misc.TimeStrober(i) for i in values[0]]
                else:
                    values_i[dest_val] = misc.TimeStrober(values)

        devices_mapping.append(values_i)

    return devices_mapping


def dump_dict_to_excel(dump_dict, sheet, line_number):
    misc.num2col

    minmax = SimpleNamespace(min=0, max=0)

    lst = []
    def grow_lst(lst, idx):
        while idx >= len(lst):
            lst.append(None)

    for cols, vals in dump_dict.items():
        if cols is None:
            continue

        def add_val_to_list(col, val):
            idx = misc.col2num[col]

            if minmax.min > idx:
                minmax.min = idx

            if minmax.max < idx:
                minmax.max = idx

            grow_lst(lst, idx)
            lst[idx] = val

        if not ":" in cols:
            add_val_to_list(cols, vals[0])

        else:
            ci, cj = cols.upper().split(":")
            i, j = misc.col2num[ci], misc.col2num[cj]

            for ii in range(i, j+1):
                add_val_to_list(misc.num2col[ii], vals[ii-i])

    lst_slice = lst[minmax.min:minmax.max+1]
    excel_range = f"{misc.num2col[minmax.min]}{line_number}:{misc.num2col[minmax.max]}{line_number}"

    def do_output():
        range = sheet.Range(excel_range)
        range.Value = lst_slice

    misc.try_n(do_output, tries=50)


if __name__ == "__main__":

    wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet()

    xwbook = xw.books(wb.Name)
    devices = get_devices_from_book(xwbook)

    p = get_harcoded_parameters(wb)

    prev = -9999999
    prev_save = -9999999

    print('Start')
    last_success = np.inf
    last_update = None

    line_number = misc.force_int(devices[0].line.start_row_no)
    for i in range(line_number, 60000):
        if not inputs_sheet.Range(f'A{i}').Value:
            break
        line_number += 1

    try:
        for i in range(int(1e16)):

            update = outputs_sheet.Range('B3').Value
            if update != last_update:
                values_to_write = update_write_values(wb, devices)

            trigger_only_after_reading_interval = []

            for device, val_dict in zip(devices, values_to_write):
                for register, value in val_dict.items():
                    if isinstance(value, list):
                        device.write_bits(value, register)
                    elif str(value).lower().startswith("0x"):
                        trigger_only_after_reading_interval.append(
                            (lambda device, value: # To form a closure over device and value
                                (lambda: (device.device.serial.write(msg:=(b:=bytes.fromhex(value[2:]))+relay_crc(b)),
                                          time.sleep(0.05))))(device, value))
                    else:
                        device.write(value.get_value(), register)

            if(time.time() - prev > p.reading_interval):
                _prev = time.time()

                for f in trigger_only_after_reading_interval:
                    f()


                dump_dict = {}
                for device in devices:
                    dump_dict[device.line.dump_cols] = device.read()
                dump_dict["A"] = [str(datetime.now())]

                dump_dict_to_excel(dump_dict, inputs_sheet, line_number)
                line_number += 1

                prev = _prev


            if(time.time() - prev_save > 60*3):
                prev_save = time.time()
                wb.Save()


            if (i+1)%10 ==0:
                print("y", end='\n' if i%600==0 else '', flush=True)

            last_success = time.time()
    except:
        traceback.print_exc()
        misc.exit_after_n_seconds(0.5)