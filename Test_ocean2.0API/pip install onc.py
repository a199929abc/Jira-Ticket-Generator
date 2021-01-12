from onc.onc import ONC
import json
import requests

Instrument_Category=''
Instrument=''
Serial_Number= ''
deviceId=11057

url = 'https://data.oceannetworks.ca/api/devices'
parameters = {'method':'get',
            'token':'71f23a7a-8b7f-4b13-bd24-0948bc76eab0', # replace YOUR_TOKEN_HERE with your personal token obtained from the 'Web Services API' tab at https://data.oceannetworks.ca/Profile when logged in.
            'deviceId':deviceId}
  
response = requests.get(url,params=parameters)
  
if (response.ok):
    devices = json.loads(str(response.content,'utf-8')) # convert the json response to an object
    for device in devices:
        Instrument=device.get('deviceName')
        Instrument_Category=device.get('deviceCategoryCode')
        
        #Instrument_Category=device.get('devicecategoryCode')
else:
    if(response.status_code == 400):
        error = json.loads(str(response.content,'utf-8'))
        print(error) # json response contains a list of errors, with an errorMessage and parameter
    else:
        print ('Error {} - {}'.format(response.status_code,response.reason))

SNtemp = Instrument.split()
if len(SNtemp>1):
    Serial_Number=None
else:
    Serial_Number=SNtemp[-1]

print(Instrument_Category)
print(Instrument)
print(Serial_Number)