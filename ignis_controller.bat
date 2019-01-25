
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
import importlib; importlib.reload(misc)
from misc import *
from pywintypes import com_error
import numpy as np
import os

excel =  register_excel_application()


# In[3]:


if is_interactive():
    import os
    import subprocess
    subprocess.call(['jupyter', 'nbconvert', '--to', 'script', 'ignis_controller.ipynb'])
    try:
        os.remove('ignis_controller.bat')
    except: pass
    os.rename('ignis_controller.py', 'ignis_controller.bat')


# In[3]:


wb, inputs_sheet, outputs_sheet = register_excel_workbook(excel)


# In[4]:


com_val = "COM%d"%(outputs_sheet.Range('B5').Value)
ins1, ins2, ins3, ins4, ins5, ins6 = get_instruments(com_val) #add comm number
inst = inputs_writer()
prev = -9999999
prev_save = -9999999

print('Start')
try:
    last_update = None
    strobes = [timeStrober('off') for i in range(8)]
    
    i=-1
    while(1):
        i+=1
        time.sleep(0.02)
        try:
            
            update = outputs_sheet.Range('B3').Value
            if update != last_update:
                strobe_settings = outputs_sheet.Range('C3:J3').Value[0]
                strobes1 = [timeStrober(s) for s in strobe_settings]

                strobe_settings = outputs_sheet.Range('M3:T3').Value[0]
                strobes2 = [timeStrober(s) for s in strobe_settings]
                
                last_update = update
            
            bits = [s.is_on() for s in strobes1][::-1]
            ins1.write_register(320, bits2int(bits))

            bits = [s.is_on() for s in strobes2][::-1]
            ins6.write_register(320, bits2int(bits))
            
            if(time.time() - prev > 5):
                prev = time.time()
                inst.do_readings(inputs_sheet, ins1, ins2, ins3, ins4, ins5, ins6)
                
            if(time.time() - prev_save > 60*3):
                prev_save = time.time()
                wb.Save()

        except com_error:
            if i%10 ==0:
                print('.', end='\n' if i%600==0 else '')
        #except ValueError:
        #    print('\nValueError occured')

except KeyboardInterrupt:
    print('\nStop')
    
    

