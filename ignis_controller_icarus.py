from types import SimpleNamespace

import misc
import sys
import time
import importlib
importlib.reload(misc)
import  pywintypes
import numpy as np
import devices
import scale_device
import traceback

# For some reason Pycharm doesn't pick up com_error
com_error = getattr(pywintypes, 'com_error')


def get_devices_to_use(wb):
    devices = wb.Sheets['Parameters'].Range("device_numbers_to_address").Value
    devices = [i.strip().lower() for i in devices.split(',')]

    scale = False
    if 'scale' in devices:
        scale = True

    devices = [int(i) for i in devices if i != 'scale']

    return scale, devices


def get_fan_parameters(wb):
    return SimpleNamespace(
        run_reg =  misc.force_int(wb.Sheets['Parameters'].Range("fan_run_register")),
        run_val =  misc.force_int(wb.Sheets['Parameters'].Range("fan_run_value")),
        freq_reg = misc.force_int(wb.Sheets['Parameters'].Range("fan_frequency_register")),
        freq_val = misc.force_int(wb.Sheets['Parameters'].Range("fan_frequency_value")),
    )


if __name__ == "__main__":

    wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet()

    time_interval_val = outputs_sheet.Range('B5').Value

    use_scale, device_ids = get_devices_to_use(wb)

    if use_scale:
        scale = scale_device.ScaleDevice()
    else:
        # Mock scale with no value
        scale = SimpleNamespace(mass=None)

    instruments = devices.get_instruments(device_ids)

    ins1, ins2, ins3, ins4, ins6 = [instruments.get(i, [None, None])[1] for i in [1, 2, 3, 4, 6]]


    iwrite = misc.inputs_writer_icarus()
    prev = -9999999
    prev_save = -9999999

    ins1_ok = []
    ins6_ok = []
    read_ok = []

    print('Start')
    last_success = np.inf
    try:
        last_update = None

        i=-1
        while(1):
            i+=1
            time.sleep(0.02)
            try:
                update = outputs_sheet.Range('B3').Value
                if update != last_update:
                    strobe_settings = outputs_sheet.Range('C3:J3').Value[0]
                    strobes1 = [misc.timeStrober(s) for s in strobe_settings]

                    strobe_settings = outputs_sheet.Range('M3:T3').Value[0]
                    strobes2 = [misc.timeStrober(s) for s in strobe_settings]
                    last_update = update

                fan_params = get_fan_parameters(wb)

                if ins1 is None:
                    ins1_ok = [True]
                else:
                    ins1_ok = ins1_ok[-10:]+[misc.write_to_inst(ins1, [s.is_on() for s in strobes1])]

                if ins6 is None:
                    ins6_ok = [True]
                else:
                    ins6_ok = ins6_ok[-20:]+[misc.write_to_inst(ins6, fan_params.run_val, reg=fan_params.run_reg),
                                             misc.write_to_inst(ins6, fan_params.freq_val, reg=fan_params.freq_reg)]


                if(time.time() - prev > time_interval_val):
                    _prev = time.time()
                    _ok = iwrite.do_readings(wb, inputs_sheet, ins1, ins2, ins3, ins4)
                    read_ok = read_ok[-3:]+[_ok]


                    #Hack in the scale readings
                    try:
                        iwrite.inputs_sheet.Range(f"AH{iwrite.curr_line}").Value = scale.mass
                    except AttributeError:
                        print(traceback.format_exc())
                        sys.exit()


                    if read_ok[-1]:
                        prev = _prev

                if(time.time() - prev_save > 60*3):
                    prev_save = time.time()
                    wb.Save()

                if ~np.any(ins1_ok): #not even once in 4 times
                    raise OSError("Can't write to instrument 1.")

                if ~np.any(ins6_ok): #not even once in 4 times
                    raise OSError("Can't write to instrument 6.")

                if ~np.any(read_ok):
                    raise OSError("One instrument can't be read do_readings(self, ...).")

                if (i+1)%10 ==0:
                    print("y", end='\n' if i%600==0 else '', flush=True)

                last_success = time.time()

            #Give alarm if the excel sheet is blocked
            except com_error:
                if i%10 == 0:
                    print('x', end='\n' if i%600==0 else '', flush=True)

                if time.time() - last_success > 10:
                    bits = [s.is_on() for s in strobes1]
                    bits[-2] = round(np.random.rand()) #Overwrite alarm bit

                    if ins1 is not None:
                        misc.write_to_inst(ins1, bits)


    except KeyboardInterrupt:
        print('\nStop')

