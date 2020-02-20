
# coding: utf-8

# In[2]:


''' >NUL  2>NUL
@echo off
cd /d %~dp0
:loop
python %0 %*
goto loop
'''

import misc; 
import time
import importlib; importlib.reload(misc)
from pywintypes import com_error
import numpy as np
import os


# In[4]:


wb, inputs_sheet, outputs_sheet = misc.get_ignis_spreadsheet()


# In[ ]:


com_val = "COM%d"%(outputs_sheet.Range('B5').Value)
ins1, ins2, ins3, ins4, ins5, ins6, ins7 = misc.get_instruments(com_val, 7) #add comm number
inst = misc.inputs_writer()
prev = -9999999
prev_save = -9999999

ins1_ok = []
ins6_ok = []
read_ok = []

print('Start')
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
            ins6_ok = ins6_ok[-10:]+[misc.write_to_inst(ins6, [s.is_on() for s in strobes2])]
            
            if(time.time() - prev > 5):
                _prev = time.time()
                read_ok = read_ok[-3:]+[inst.do_readings(wb, inputs_sheet, [ins1, ins2, ins3, ins4, ins5, ins6, ins7])]
                
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
                raise OSError("One instrument can't be read do_readings(self, ...).")

        except com_error:
            if i%10 ==0:
                print('.', end='\n' if i%600==0 else '')
        #except ValueError:
        #    print('\nValueError occured')

except KeyboardInterrupt:
    print('\nStop')

