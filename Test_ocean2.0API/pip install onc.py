from onc.onc import ONC
import json
import requests
import re
Instrument_Category=''
Instrument=''
Serial_Number= ''
deviceId=23661

url = 'https://data.oceannetworks.ca/api/devices'
parameters = {'method':'get',
            'token':'71f23a7a-8b7f-4b13-bd24-0948bc76eab0', # replace YOUR_TOKEN_HERE with your personal token obtained from the 'Web Services API' tab at https://data.oceannetworks.ca/Profile when logged in.
            'deviceId':deviceId}
  
response = requests.get(url,params=parameters)
  
if (response.ok):
    devices = json.loads(str(response.content,'utf-8')) # convert the json response to an object
    for device in devices:
        Instrument=device.get('deviceName')
        InstrumentCategory=device.get('deviceCategoryCode')
       # print(InstrumentCategory)
        Instrument_Categorurl = 'https://data.oceannetworks.ca/api/deviceCategories'
else:
    if(response.status_code == 400):
        error = json.loads(str(response.content,'utf-8'))
        print(error) # json response contains a list of errors, with an errorMessage and parameter
    else:
        print ('Error {} - {}'.format(response.status_code,response.reason))

url = 'https://data.oceannetworks.ca/api/deviceCategories'
parameters = {'method':'get',
            'token':'71f23a7a-8b7f-4b13-bd24-0948bc76eab0', # replace YOUR_TOKEN_HERE with your personal token obtained from the 'Web Services API' tab at https://data.oceannetworks.ca/Profile when logged in.
            'deviceCategoryCode':InstrumentCategory}
  
response = requests.get(url,params=parameters)
  
if (response.ok):
    deviceCategories = json.loads(str(response.content,'utf-8')) # convert the json response to an object
    for deviceCategory in deviceCategories:
        InstrumentCategory=deviceCategory.get('deviceCategoryName')
else:
    if(response.status_code == 400):
        error = json.loads(str(response.content,'utf-8'))
        print(error) # json response contains a list of errors, with an errorMessage and parameter
    else:
        print ('Error {} - {}'.format(response.status_code,response.reason))



def processString(instrumentName):
    Instrument=instrumentName
    SNtemp = Instrument.split()
    if len(SNtemp)<=1:
        Serial_Number=None
    elif 'SN' in Instrument:
        sep='SN'
        head,sep,tail = Instrument.partition(sep)
        Instrument=head
        Instrument = Instrument.replace('(', '').replace(')', '')
        Serial_Number=tail.replace('(', '').replace(')', '')
    elif 'S/N' in Instrument:
        sep='S/N'
        head,sep,tail = Instrument.partition(sep)
        Instrument=head
        Instrument = Instrument.replace('(', '').replace(')', '')
        Serial_Number=tail.replace('(', '').replace(')', '')
    else: 
        st1=' '
        Serial_Number=SNtemp[-1]
        Serial_Number=Serial_Number.replace('(', '').replace(')', '')
        Instrument= st1.join(SNtemp[:-1])
    return Instrument,Serial_Number


InstrumentName,Serial_Number=processString(Instrument)
print('InstrumentCategory is '+InstrumentCategory)
print('Serial N:'+ Serial_Number)
print('instrment name : '+ InstrumentName)

