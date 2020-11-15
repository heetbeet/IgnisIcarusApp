from types import SimpleNamespace
from win32com.universal import com_error

import misc
import sys
import time
from device_information import  get_devices_from_book
from misc import is_nan
import numpy as np
import traceback


def get_harcoded_parameters(wb):
    return SimpleNamespace(
        reading_interval =  misc.force_int(wb.Sheets['Parameters'].Range("reading_interval").value),
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
                    values = wb.Sheets['Outputs'].Range(source_val).value
                except com_error:
                    values = misc.get_named_range(wb, source_val).value

                if isinstance(values, tuple):
                    values_i[dest_val] = list(values[0])
                else:
                    values_i[dest_val] = misc.force_int(values)

        devices_mapping.append(values_i)

    return devices_mapping

if __name__ == "__main__":


    wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet()
    devices = get_devices_from_book(wb)

    p = get_harcoded_parameters(wb)


    iwrite = misc.inputs_writer_icarus()
    prev = -9999999
    prev_save = -9999999


    print('Start')
    last_success = np.inf
    last_update = None

    line_number = devices[0].start_row_no
    for i in range(line_number, 60000):
        if not inputs_sheet.Range(f'A{i}').value:
            line_number += i
            break

    for i in range(np.inf):

        update = outputs_sheet.Range('B3').Value
        if update != last_update:
            values_to_write = update_write_values(wb,devices)

        for device, val_dict  in zip(devices, values_to_write):
            for register, value in val_dict.items():
                if isinstance(value, list):
                    device.write_bits(value, register)

                else:
                    device.write(value, register)


        if(time.time() - prev > p.reading_interval):
            _prev = time.time()

            for device in devices:
                device.output_to_excel(inputs_sheet, line_number)

            line_number += 1

            prev = _prev



        if(time.time() - prev_save > 60*3):
            prev_save = time.time()
            wb.Save()


        if (i+1)%10 ==0:
            print("y", end='\n' if i%600==0 else '', flush=True)

        last_success = time.time()


