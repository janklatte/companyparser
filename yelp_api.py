#!/usr/local/bin/python

# A simple tool to query a certain number of businesses based on location and keyword from the Yelp Databank into
# an excel spreadsheet (.xls) and .csv (UTF-8) file for importing into e.g. a CRM software

import xlwt
import csv
import json

from yelp.client import Client
from yelp.oauth1_authenticator import Oauth1Authenticator

#get yelp access data
with open('yelp_cred.json') as yelp_data_file:
	yelp_data = json.load(yelp_data_file)

YOUR_CONSUMER_KEY = yelp_data['Consumer_Key']
YOUR_CONSUMER_SECRET = yelp_data['Consumer_Secret']
YOUR_TOKEN = yelp_data['Token']
YOUR_TOKEN_SECRET = yelp_data['Token_Secret']

#input data
town = raw_input("Town: ")
keyword = raw_input("Business: ")
num_result = raw_input('Number of Results: ')


#preparing excel spreadsheet
wb = xlwt.Workbook()
ws = wb.add_sheet('Company List')

#adding header to spreadsheet
ws.write(0,0,'Name')
ws.write(0,1,'Phone')
ws.write(0,2,'URL')
ws.write(0,3,'Street')
ws.write(0,4,'City')

#preparing csv file
csvfile = open('companylist.csv','w')
#header for csv
header = '"Company Name";"Logo";"Company Type";"Industry";"Employees";"Annual Revenue";"Currency";"Comment";"Responsible person";"Address";"Street";"Apartment / Suite";"City";"Region";"State / Province";"Zip";"Country";"Billing Address";"Street (billing)";"Apartment / Suite (billing)";"City (billing)";"Region (billing)";"State / Province (billing)";"Zip (billing)";"Country (billing)";"Work Phone";"Mobile";"Fax";"Home Phone";"Pager Number";"Other Phone Number";"Corporate Website";"Personal Page";"Facebook Page";"LiveJournal";"Twitter";"Other Website";"Work E-mail";"Home E-mail";"Other E-mail";"Skype ID";"ICQ Number";"MSN/Live!";"Jabber";"Other Contact";"Payment Details";"Available to everyone";\n'
csvfile.write(header)

#calculating number of loops and rest to make multiple calls
#Yelp only allows a max of 20 responses
try:
	rest = int(num_result)%20
	loops = (int(num_result)-rest)/20
except:
	print("Wrong Values only Integers (1,2,3,4,5,6,100,1000..)")
	exit()

#@linenumber: number of lines in the excel spreadsheet @offset: offset for yelp api
linenumber = 0
offset = 0

for i in range(0,loops+1):

	#API Authentication
	auth = Oauth1Authenticator(
	    consumer_key=YOUR_CONSUMER_KEY,
	    consumer_secret=YOUR_CONSUMER_SECRET,
	    token=YOUR_TOKEN,
	    token_secret=YOUR_TOKEN_SECRET
	)

	client = Client(auth)

	#test if last loop in that case only rest for limit
	if i is loops:
		limit = rest
	else:
		limit = 20

	# parameters for API call
	params = {
		'term': keyword,
		'cc': 'DE',
		'lang': 'de',
		'sort': 0,
		'offset': offset,
		'limit': limit
	}

	offset += 20

	#get response
	response = client.search(town,**params)
	index = len(response.businesses)

	for x in range(0,index):
		#object with street ([0] and city [1])
		#sometimes city is only a district of a city
		adressobj =  response.businesses[x].location.display_address

		#set google search link for url
		url = 'https://www.google.de/search?q='+response.businesses[x].name+'&ie=utf-8&oe=utf-8&gws_rd=cr&ei=V9_pVsmbHMXa6ASrvr64AQ'

		print(response.businesses[x].name)


		#test if phone number excist if not set to 0
		if response.businesses[x].phone is None:
			response.businesses[x].phone = "0"

		line = '"' + response.businesses[x].name + '"' + ';"";"Partner";"' + keyword + '";"less than 50";"";"USD";"";"Jan Kleine-Klatte";"";"' + adressobj[0] + '";"";"' + adressobj[1] + '";"' + town + '";"";"";"Deutschland";"";"";"";"";"";"";"";"";"' + response.businesses[x].phone + '";"";"";"";"";"";"";"";"";"";"";"' + url + '";"";"";"";"";"";"";"";"";"";"yes";\n'
		
		#write csv and xls file

		csvfile.write(line.encode('utf-8'))

		ws.write(linenumber+1,0,response.businesses[x].name)
		ws.write(linenumber+1,1,response.businesses[x].phone)
		ws.write(linenumber+1,2,url)
		ws.write(linenumber+1,3,adressobj[0])
		ws.write(linenumber+1,4,adressobj[1])
		linenumber+=1

#close and save files
wb.save('companylist.xls')
csvfile.close()

#Program end






