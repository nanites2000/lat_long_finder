
import requests, os
import json
import openpyxl


import glob
states = ['Alabama',
'Alaska',
'Arizona',
'Arkansas',
'California',
'Colorado',
'Connecticut',
'Delaware',
'Florida',
'Georgia',
'Hawaii',
'Idaho',
'Illinois',
'Indiana',
'Iowa',
'Kansas',
'Kentucky',
'Louisiana',
'Maine',
'Maryland',
'Massachusetts',
'Michigan',
'Minnesota',
'Mississippi',
'Missouri',
'Montana',
'Nebraska',
'Nevada',
'New Hampshire',
'New Jersey',
'New Mexico',
'New York',
'North Carolina',
'North Dakota',
'Ohio',
'Oklahoma',
'Oregon',
'Pennsylvania',
'Rhode Island',
'South Carolina',
'South Dakota',
'Tennessee',
'Texas',
'Utah',
'Vermont',
'Virginia',
'Washington',
'West Virginia',
'Wisconsin',
'Wyoming',
'District of Columbia',
'Puerto Rico',
'Guam',
'American Samoa',
'U.S. Virgin Islands',
'Northern Mariana Islands']





files = glob.glob('./*.xlsx')
print("You need to have this in a directory with the only .xlsx file is the one you want the addresses of")
for file in files:

    xl = openpyxl.load_workbook(file)
    sheet_name = xl.sheetnames[0]

    sheet = xl[sheet_name]

    print('Make sure rowid is in column A, latitude is in column B, and longitude in C')
    print('The location data will be in column D')

    try:
        i=1
        rowlist =[]
        while 1:

            latitude = sheet['B'+str(i)].value
            longitude = sheet['C'+str(i)].value

            if latitude == None or longitude == None:
                break
            #rowlist.append([lat , long])

            try:
                url = f'http://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/reverseGeocode?f=pjson&featureTypes=StreetInt&location={longitude},{latitude}'
                res = requests.get(url)
                x = json.loads(res.content)
                location = x['address']['Region']
                if location in states:
                    sheet['D' + str(i)] = location
                    print('OK')
                else:
                    sheet['D' + str(i)] = 'ZZZ  ' + location
                    print('Outside US')
            except:
                sheet['D' + str(i)] = 'Check This'
                print(f'{i} it failed')
            i = i + 1
    except:
        print('This run had an error of some kind')

    cwd = os.getcwd()
    path = cwd + '/results'
    if not os.path.exists(path):
        os.mkdir(path)

    xl.save('results/youdidit.xlsx')
    final = input("Program Complete. Hit ENTER to end.")
