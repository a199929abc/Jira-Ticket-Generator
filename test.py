from request import *
import pandas as pd
import numpy as np
df = pd.DataFrame(np.empty((0, 7))) 
df.columns=['Site/Location','Ticket Link','Instrument Category','Instrument',
                    'Serial Number','Created Ticket','rowNum']
df.insert(7,'DeviceID',26719)
print(df.head)
for index, row in df.iterrows():
    local_instrument,local_instrument_category,Serial_Number=request(row)   
    print('local_instrument_category is '+local_instrument_category)
    print('Serial N:'+ Serial_Number)
    print('instrment name : '+ local_instrument)