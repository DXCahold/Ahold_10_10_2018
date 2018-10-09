#!/usr/bin/env python
import os,sys,json,xlrd,requests,logging
from flask import Flask
from flask import request
from flask import make_response

def excel2json(workbook):
	book = xlrd.open_workbook(workbook)
	sheets = book.sheet_names()
	source={}
	for sheet in sheets:
		source[sheet] = []
		page = book.sheet_by_name(sheet)
		for row in range(1,page.nrows):
			data = {}
			for col in range(0,page.ncols):
				data[str(page.cell(0,col).value)] = str(page.cell(row,col).value)
			source[sheet].append(data)
	return source

workbook = "Ahold.xlsx"
book = excel2json(workbook)
phonenumber,signedin = "",False

app = Flask(__name__)
app.logger.addHandler(logging.StreamHandler(sys.stdout))
app.logger.setLevel(logging.ERROR)

@app.route('/', methods=['POST', 'GET'])
def webhook():
	if request.method == 'POST':
		req = json.loads((request.data).decode("utf-8"))
		request_data = {"known":{}, "unknown":"", "fulfillmentText":"", "result" : ""}
		for key in req['queryResult']['parameters'].keys():
			request_data["known"].update({key : req['queryResult']['parameters'][key]})
		request_data["unknown"] = str(req['queryResult']['intent']['displayName'])
		request_data["fulfillmentText"] = str(req['queryResult']['fulfillmentText'])
		print (request_data)
		
		if request_data["unknown"] == "welcome":
			if signedin:
				request_data["result"] = "how may i assist you?"
			else:
				request_data["result"] = request_data["fulfillmentText"]
			
		if request_data["unknown"] == "phonenumber-yes":
			phonenumber,signedin,request_data["result"] = request_data['known']['phone-number'],True,request_data["fulfillmentText"].replace("*result",str([request_data['known']['phone-number'][i:i+1] for i in range(0,len(request_data['known']['phone-number']),1)]).replace(" ","").replace("'","").replace("[","").replace("]","").replace(","," "))
		
		if request_data["unknown"] == "phonenumber-no":
			phonenumber,signedin,request_data["result"] = "",True,request_data["fulfillmentText"]
			
		if request_data["unknown"] == "Thankyou":
			phonenumber,signedin,request_data["result"] = "",False, request_data["fulfillmentText"]
					
		
		print (request_data["result"])
		return json.dumps({"fulfillmentText":request_data["result"]})
	else:
		return "<h1>Home</h1>"

if __name__ == '__main__':
	port = int(os.getenv('PORT', 5000))
	print("Starting app on port %d" % port)
	app.run(debug=True, port=port, host='0.0.0.0')
