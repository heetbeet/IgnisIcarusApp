import misc; 
import time
import importlib; importlib.reload(misc)
from pywintypes import com_error
import numpy as np
import traceback


wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet()

try:
    comms_sheet = [i for i in wb.Sheets if i.Name.lower() == 'comms'][0]
except IndexError:
    raise ValueError("Spreadsheet doesn't have sheet named 'comms'")


com_val = "COM%d"%(outputs_sheet.Range('B5').Value)

ins1, ins2, ins3, ins4, ins5, ins6, ins7 = [None]*7
for i in range(1, 7+1):
    if comms_sheet.Range(f"A{i+4}").Value:
        exec(f"ins{i} = misc.get_instrument(com_val, i)", globals())
    else:
        print(f"Not reading instrument {i}")


inst = misc.inputs_writer()
prev = -9999999
prev_save = -9999999

ins1_ok = [True]
ins6_ok = [True]
read_ok = [True]

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

            if ins1 is not None:
                ins1_ok = ins1_ok[-10:]+[misc.write_to_inst(ins1, [s.is_on() for s in strobes1])]
            if ins6 is not None:
                ins6_ok = ins6_ok[-10:]+[misc.write_to_inst(ins6, [s.is_on() for s in strobes2])]
            
            if(time.time() - prev > 5):
                _prev = time.time()
                read_ok = read_ok[-3:]+[inst.do_readings(wb, inputs_sheet, ins1, ins2, ins3, ins4, ins5, ins6)]
                
                if read_ok[-1]:
                    prev = _prev
                
            if(time.time() - prev_save > 60*3):
                prev_save = time.time()
                wb.Save()
            
            if ~np.any(ins1_ok): #not even once in 4 times
                raise OSError("Can't write to instrument 1.")
            if ~np.any(ins6_ok):
                raise OSError("Can't write to instrument 6.")
            if ~np.any(read_ok):
                raise OSError("One instrument can't be read by do_readings(self, ...).")
            
            if i%10 ==0: 
                print("y", end='\n' if i%600==0 else '', flush=True)
                
            last_success = time.time()
            
        #Give alarm if the excel sheet is blocked
        except com_error:
            print(traceback.format_exc())
            if i%10 ==0:
                print('x', end='\n' if i%600==0 else '', flush=True)
                    
            if time.time() - last_success > 10:
                bits = [s.is_on() for s in strobes1]
                bits[-2] = round(np.random.rand()) #Overwrite alarm bit

                if ins1 is not None:
                    misc.write_to_inst(ins1, bits)
                if ins7 is not None:
                    ins7.write_register(0x0310, int(outputs_sheet.Range("M12").Value))

        #except ValueError:
        #    print('\nValueError occured')

except KeyboardInterrupt:
    print('\nStop')

