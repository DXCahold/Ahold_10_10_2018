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
		result = ""
		print (request_data)
		if request_data["unknown"] == "product":
			availables,outofstocks = [],[]
			for sheet in book.keys():
				for row in book[sheet]:
					headers = row.keys()
					if "content" in headers:
						for key in request_data['known']:
							if len(request_data['known'][key])>0:
								if request_data['known'][key] in row[key]:
									if int(float(row["quantity"]))>0:
										availables.append(row[request_data["unknown"]])
									else:
										outofstocks.append(row[request_data["unknown"]])
			if len(availables)>0:
				request_data["result"] = str(request_data["fulfillmentText"]).replace('*result','available').replace('*availables ',str(availables).replace("[","").replace("]","").replace("'","").replace('"',''))
			else:
				request_data["result"] = str(request_data["fulfillmentText"]).replace('*result','not available').replace("you can find *availables  in stock"," sorry for inconvenience!")
			if len(outofstocks)>0:
				request_data["result"] = str(request_data["result"]).replace("*outofstocks",str(outofstocks).replace("[","").replace("]","").replace("'","").replace('"',''))
			else:
				request_data["result"] = str(request_data["result"]).replace(" *outofstocks currently unavailable","")
		else:
			for sheet in book.keys():
				for row in book[sheet]:
					headers = row.keys()
					if request_data["unknown"] in headers:
						for key in request_data['known']:
							if request_data['known'][key] == row[key]:
								result = row[request_data["unknown"]]
			request_data["result"] = request_data["fulfillmentText"].replace("*result",result)
		print (request_data["result"])
		return json.dumps({"fulfillmentText":request_data["result"]})
	else:
		return "<h1>Home</h1>"

if __name__ == '__main__':
	port = int(os.getenv('PORT', 5000))
	print("Starting app on port %d" % port)
	app.run(debug=True, port=port, host='0.0.0.0')
