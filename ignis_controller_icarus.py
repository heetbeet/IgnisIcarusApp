import os
import subprocess
import sys
import xlwings as xw
from types import SimpleNamespace
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


def get_harcoded_parameters(wb):
    return SimpleNamespace(
        reading_interval =  misc.force_int(wb.Sheets['Parameters'].Range("reading_interval").Value),
    )

def update_write_values(wb, devices):
    '''
    Don't overtax this function
    '''
    devices_mapping = []
    for d in devices:
        values_i = {}
        for i in range(1,1+8):
            source = f'source_{i}'
            dest = f'write_{i}'

            if (not is_nan(source_val := d.line.__dict__[source]) and
                not is_nan(dest_val := d.line.__dict__[dest])):

                dest_val = misc.force_int(dest_val)

                # If it is a named range, search everywhere, if it
                try:
                    values = wb.Sheets['Outputs'].Range(source_val).Value
                except com_error:
                    values = misc.get_named_range(wb, source_val).Value

                if isinstance(values, tuple):
                    values_i[dest_val] = list(values[0])
                else:
                    values_i[dest_val] = misc.force_int(values)

        devices_mapping.append(values_i)

    return devices_mapping

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
                values_to_write = update_write_values(wb,devices)

                for device, val_dict  in zip(devices, values_to_write):
                    for register, value in val_dict.items():
                        if not isinstance(value, list):
                            device.write(value, register)

            for device, val_dict in zip(devices, values_to_write):
                for register, value in val_dict.items():
                    if isinstance(value, list):
                        device.write_bits(value, register)

            if(time.time() - prev > p.reading_interval):
                _prev = time.time()

                for device in devices:
                    device.output_to_excel(inputs_sheet, line_number)

                inputs_sheet.Range(f"A{line_number}").Value = str(datetime.now())

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