import requests as req
import json
import xlrd
import csv
import time
import re

workbook = xlrd.open_workbook("data.xlsx")
sheet = workbook.sheet_by_index(0)

not_found = []
url = "https://geocode-maps.yandex.ru/1.x/"

with open("result.csv", "w") as f:
	fieldnames = ['Полное название', 'Категории', 'Город', 'Адрес', 'Vkontakte', 'тел.', 'Район']
	writer = csv.writer(f, delimiter=";")
	writer.writerow(fieldnames)
	f.close()

count = 1

border = 1

for row in range(1, sheet.nrows):
	cols = sheet.row_values(row)
	border += 1
	if border < 11780: # default 0. позиция на которой парсер споткнулся
		continue
	
	#print(cols)

	line = []
	for j in range(6):
		line.append(cols[j])

	city = cols[2]
	address = cols[3]
	district = ''
	#print("{}, {}".format(city, address))

	params = {
		'format'  : 'json',
		'geocode' :	"{}, {}".format(city, address)
	}

	try:
		r = req.get(url, params = params)
	except Exception as e:
		print(str(e))
		time.sleep(10)
		row = row - 1
		continue
	resp = json.loads(r.text)
	try:
		obj = resp['response']['GeoObjectCollection']['featureMember'][0]
	except:
		continue
	lng, lat = obj['GeoObject']['Point']['pos'].split(" ") # lng, lat

	params = {
		'format'	: 'json',
		'geocode' 	: '{},{}'.format(lng, lat),
		'kind'		: 'district'
	}
	try:
		r = req.get(url, params = params)
	except Exception as e:
		print(str(e))
		time.sleep(10)
		row = row - 1
		continue

	resp = json.loads(r.text)
	
	try:
		featureMembers = resp['response']['GeoObjectCollection']['featureMember']
	except Exception as e:
		print(resp)
		time.sleep(10)
		continue
	if len(featureMembers) == 0:
		not_found.append((city, address))

	for item in featureMembers:
		components = item['GeoObject']['metaDataProperty']['GeocoderMetaData']['Address']['Components']
		for comp in components:
			if comp['kind'] == 'district':
				buff = re.findall(r'(\sрайон$)|(\sокруг$)', comp['name'])
				if len(buff) > 0:
					district = comp['name']
					
	line.append(district)

	with open("result.csv", "a") as f:
		writer = csv.writer(f, delimiter=";")
		writer.writerow(line)
		f.close()

	count = count + 1
	#print("{} / 12601".format(count), end='\r', flush=True)
	print("{} / 12601".format(count))

'''
with open("not_found.txt", "w") as nff:
	nff.write(not_found)
	nff.close()'''