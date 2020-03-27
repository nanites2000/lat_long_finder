
import requests, os
import json
import pandas as pd
import xlrd # to make pandas work


import glob
files = glob.glob('./*.xlsx')


for file in files:

    xl = pd.ExcelFile(file)
    sheet_name = xl.sheet_names[0]
    sheet = xl.parse(sheet_name)


    rowlist = []
    for index, row in sheet.iterrows():
        latitude, longitude = sheet.iloc[index, [3, 4]]

        try:
            url = f'http://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/reverseGeocode?f=pjson&featureTypes=StreetInt&location={longitude},{latitude}'
            res = requests.get(url)
            x = json.loads(res.content)
            rowlist.append(x['address']['Region'])
        except:
            rowlist.append('Check This')
            print(f'{index+1} it failed')

    sheet['State'] =rowlist
    cwd = os.getcwd()
    path = cwd + '/results'
    if not os.path.exists(path):
        os.mkdir(path)

    sheet.to_excel('./results/Result_Yay.xlsx', sheet_name='Sheet 1', index=False)
#final = input

    # xl = openpyxl.load_workbook(file)
    # sheet_name = xl.sheetnames[0]
    # print(sheet_name)
    # sheet = xl[sheet_name]
    # for row in sheet:
    #     print()


# latitude = 36.20171
# longitude = -81.1286
# url = f'http://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/reverseGeocode?f=pjson&featureTypes=StreetInt&location={longitude},{latitude}'
# print(url)
# res = requests.get(url)
# x = json.loads(res.content)
# print(x['address']['Region'])




# import requests, os,bs4,re,random,time
#
# url = 'https://www.latlong.net/Show-Latitude-Longitude.html'
# print("hello")
#
#
# try:
#     res = requests.get(urlbase+date)
#     res.raise_for_status()
#     soup = bs4.BeautifulSoup(res.text, 'html.parser')
#     #print(soup)
#     soupData = soup.findAll('picture')
#    #  print(soupData)
#     correct_picture = soupData[1]
#     #print(correct_picture)
#     #pic_only = correct_picture.find('img')
#     re_pattern = re.compile(r'assets.amuniversal.com/\S+')
#     #print('regular expression')
#
#     imageLink = re_pattern.findall(str(correct_picture))[0]
#     #print(imageLink)
#     # imageLink = []
#     # for tag in soup.findAll('assets.amuniversal.com/', alt=True):
#     #         # data = tag['src']
#     #         # if data[0:3] == "//a":
#     #         #     print(data)
#     #         #     imageLink = data
#     #         imageLink = tag
#     #         print(imageLink)
#
#     imageFromWeb = requests.get("https://" + imageLink)
#    # print(imageFromWeb)
#     #print(imageFromWeb.raw)
#
#     with open('CalvinHobbs/'+file_date+'.gif', "wb") as f:
#         f.write(imageFromWeb.content)
#     time.sleep(random.randint(6,11))
# except:
#     print(date)
#     print("FAILUREEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE!")
#
#
# # print("Done")
# import requests, os,bs4,re,random,time
#
# url = 'https://www.latlong.net/Show-Latitude-Longitude.html'
# res = requests.get(url)
# soup = bs4.BeautifulSoup(res.text, 'html.parser')
# print(soup)
# from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
# url = 'https://www.latlong.net/Show-Latitude-Longitude.html'
# browser = webdriver.Firefox()
# print(type(browser))
# browser.get(url)
# latitude_element = browser.find_element_by_id('latitude')
# latitude_element.send_keys('30')
# longitude_element = browser.find_element_by_id('longitude')
# longitude_element.send_keys('40')
# longitude_element.send_keys(Keys.ENTER)
# address = browser.find_element_by_id('address')
# final_address = address.get_attribute('address')
# print(final_address)




# soup = bs4.BeautifulSoup(res.text, 'html.parser')
# print(soup)
# print(soup.Region)
# x = json.loads(soup)
# print(x)
# print(x['Region'])