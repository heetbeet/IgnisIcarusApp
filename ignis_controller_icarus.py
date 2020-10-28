from types import SimpleNamespace
import misc
import time
import importlib
importlib.reload(misc)
from pywintypes import com_error
import numpy as np
import devices
import scale_device

wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet()


def get_fan_parameters():
    return SimpleNamespace(
        run_reg =  misc.force_int(wb.Sheets['Parameters'].Range("fan_run_register")),
        run_val =  misc.force_int(wb.Sheets['Parameters'].Range("fan_run_value")),
        freq_reg = misc.force_int(wb.Sheets['Parameters'].Range("fan_frequency_register")),
        freq_val = misc.force_int(wb.Sheets['Parameters'].Range("fan_frequency_value")),
    )


time_interval_val = outputs_sheet.Range('B5').Value

scale = scale_device.ScaleDevice()
instruments = devices.get_instruments([1,2,3,4,6])
ins1, ins2, ins3, ins4, ins6 = [instruments[i][1] for i in [1,2,3,4,6]]


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
            
            ins1_ok = ins1_ok[-10:]+[misc.write_to_inst(ins1, [s.is_on() for s in strobes1])]

            fan_params = get_fan_parameters()
            ins6_ok = ins6_ok[-20:]+[misc.write_to_inst(ins6, fan_params.run_val, reg=fan_params.run_reg),
                                     misc.write_to_inst(ins6, fan_params.freq_val, reg=fan_params.freq_reg)]


            if(time.time() - prev > time_interval_val):
                _prev = time.time()
                read_ok = read_ok[-3:]+[iwrite.do_readings(wb, inputs_sheet, ins1, ins2, ins3, ins4)]

                #Hack in the scale readings
                iwrite.inputs_sheet.Range(f"AH{iwrite.curr_line}").Value = scale.mass


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
            if i%10 ==0:
                print('x', end='\n' if i%600==0 else '', flush=True)
                    
            if time.time() - last_success > 10:
                bits = [s.is_on() for s in strobes1]
                bits[-2] = round(np.random.rand()) #Overwrite alarm bit
                
                misc.write_to_inst(ins1, bits)
            
        #except ValueError:
        #    print('\nValueError occured')

except KeyboardInterrupt:
    print('\nStop')

