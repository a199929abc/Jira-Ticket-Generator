#from onc.onc import ONC
import json
import requests
import re
Serial_Number= ''
def onc_request(row):
    local_instrument_category=''
    local_instrument=''
    
    
  
    
    deviceId=int(row['DeviceID'])
    
 
    url = 'https://data.oceannetworks.ca/api/devices'
    parameters = {'method':'get',
                'token':'044668ea-6492-478e-998b-2f5dfeb1123b', # replace YOUR_TOKEN_HERE with your personal token obtained from the 'Web Services API' tab at https://data.oceannetworks.ca/Profile when logged in.
                'deviceId':deviceId}
    
    response = requests.get(url,params=parameters)

    
    if (response.ok):
        devices = json.loads(str(response.content,'utf-8')) # convert the json response to an object
        for device in devices:
            local_instrument=device.get('deviceName')
            local_instrument_category=device.get('deviceCategoryCode')


    else:
        if(response.status_code == 400):
            error = json.loads(str(response.content,'utf-8'))
            print(error) # json response contains a list of errors, with an errorMessage and parameter
        else:
            print ('Error {} - {}'.format(response.status_code,response.reason))

    url = 'https://data.oceannetworks.ca/api/deviceCategories'
    parameters = {'method':'get',
                'token':'044668ea-6492-478e-998b-2f5dfeb1123b', # replace YOUR_TOKEN_HERE with your personal token obtained from the 'Web Services API' tab at https://data.oceannetworks.ca/Profile when logged in.
                'deviceCategoryCode':local_instrument_category}

    response = requests.get(url,params=parameters)
    
    if (response.ok):
        deviceCategories = json.loads(str(response.content,'utf-8')) # convert the json response to an object
        for deviceCategory in deviceCategories:
            local_instrument_category=deviceCategory.get('deviceCategoryName')
           
    else:
        if(response.status_code == 400):
            error = json.loads(str(response.content,'utf-8'))
            print(error) # json response contains a list of errors, with an errorMessage and parameter
        else:
            print ('Error {} - {}'.format(response.status_code,response.reason))
    return local_instrument,local_instrument_category


def processString(local_instrumentName):
    local_instrument=local_instrumentName
    SNtemp = local_instrument.split()
    if len(SNtemp)<=1:
        Serial_Number=None
    elif 'SN' in local_instrument:
        sep='SN'
        head,sep,tail = local_instrument.partition(sep)
        local_instrument=head
        local_instrument = local_instrument.replace('(', '').replace(')', '')
        Serial_Number=tail.replace('(', '').replace(')', '')
    elif 'S/N' in local_instrument:
        sep='S/N'
        head,sep,tail = local_instrument.partition(sep)
        local_instrument=head
        local_instrument = local_instrument.replace('(', '').replace(')', '')
        Serial_Number=tail.replace('(', '').replace(')', '')
    elif SNtemp[-1].isalpha(): 
        Serial_Number=None
        local_instrument=local_instrumentName
    else:
        st1=' '
        Serial_Number=SNtemp[-1]
        Serial_Number=Serial_Number.replace('(', '').replace(')', '')
        local_instrument= st1.join(SNtemp[:-1])
    return local_instrument,Serial_Number


'''
local_instrument,local_instrument_category=onc_request()
local_instrument, Serial_Number = processString(local_instrument)  
print('local_instrument_category is '+local_instrument_category)
print('Serial N:'+ Serial_Number)
print('instrment name : '+ local_instrument)'''
