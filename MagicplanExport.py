# importing required modules 
import os
import csv
import requests
import time
import datetime
import configparser
import tkinter as tk
# initialize the Moisture Point list
kv=[['PlanId','Plan','Room','Room Number','Symbol','Symbol ID','Form','Field','Field ID','Text']] #- have removed the headers now, but left this

# initialize the text lookup list
plan_list = [['id', 'name', 'street', 'city', 'country', 'createdate', 'update', 'email', 'Inspector']]

# initialize the failed exports
failed_export=[['id','name','reason']]

#set up session with headers for Magicplan
#directory = os.path.dirname(os.path.realpath(__file__))
directory = os.getcwd()
os.chdir(directory)
configFile = directory+"\\configurations.ini"
try:
    config = configparser.ConfigParser()
    config.read(r'configurations.ini')

    skey = "da66c28a3f24782410494e3c7497165a300d"
    #skey = config.get('global', 'skey')

    scustomer = "60b6804ac2a4d"
    #scustomer = config.get('global', 'scustomer')
    syear = config.get('global', 'syear')
    smonth = config.get('global', 'smonth')
    sday = config.get('global', 'sday')
    eyear = config.get('global', 'eyear')
    emonth = config.get('global', 'emonth')
    eday = config.get('global', 'eday')
    outputdir = config.get('global', 'outputdir')
except:
    skey = "da66c28a3f24782410494e3c7497165a300d"
    scustomer = "60b6804ac2a4d"
    syear = '2024'
    smonth = '1'
    sday = '1'
    eyear = '2024'
    emonth = '12'
    eday = '31'
    outputdir = directory

window = tk.Tk()
window.title("MagicPlan Export")
outputfolder = tk.Label(window,text="Output Folder")
exportstart = tk.Label(window,text="Starting Date:")
exportstart1 = tk.Label(window,text="/")
exportstart2 = tk.Label(window,text="/")
exportend = tk.Label(window,text="Ending Date:")
exportend1 = tk.Label(window,text="/")
exportend2 = tk.Label(window,text="/")
greeting2 = tk.Label(window,text="Ending - "+emonth+"/"+eday+"/"+eyear)
numberofplans = tk.Label(window,text="Number of plans")
output2 = tk.Label(window,text="Number of plans to export")
output3 = tk.Label(window,text="Time to search")
output4 = tk.Label(window,text="Current plan - ")
output5 = tk.Label(window,text="Current plan Export time")
output6 = tk.Label(window,text="Total Exported")
output7 = tk.Label(window,text="Total Export time")
output7a = tk.Label(window,text="Average Export time")
output8 = tk.Label(window,text="Total Export time remaining")
output8a = tk.Label(window,text="Plans remaining")
status = tk.Label(window,text="Status: Running")

tsmonth = tk.Entry(window,width=2)
tsmonth.insert(0,smonth)
temonth = tk.Entry(window,width=2)
temonth.insert(0,emonth)

tsday = tk.Entry(window,width=2)
tsday.insert(0,sday)
teday = tk.Entry(window,width=2)
teday.insert(0,eday)

tsyear = tk.Entry(window,width=4)
tsyear.insert(0,syear)
teyear = tk.Entry(window,width=4)
teyear.insert(0,eyear)

aoutputfolder = tk.Entry(window,width=30)
aoutputfolder.insert(0,outputdir)

anumberofplans = tk.Label(window,text=": 0")
aoutput2 = tk.Label(window,text=": 0")
aoutput3 = tk.Label(window,text=": 0")
aoutput4 = tk.Label(window,text=": 0")
aoutput5 = tk.Label(window,text=": 0")
aoutput6 = tk.Label(window,text=": 0")
aoutput7 = tk.Label(window,text=": 0")
aoutput7a = tk.Label(window,text=": 0")
aoutput8 = tk.Label(window,text=": 0")
aoutput8a = tk.Label(window,text=": 0")

p_len = 0
plan_count = 0
plan_time = 0

start_time_all = 0
start_time = 0

plan_id = ""
plan_name = ""

s = requests.Session()
s.headers.update({"accept": "application/json"})
s.headers.update({"key": skey})
s.headers.update({"customer": scustomer})


def extractXML(r,plan_id,plan_name):
	#create dictionary of the entire magicplan project
	dd = r.json()
	data = dd["data"]

	#it is made up of many imbeded dictionaries 
	eRoom = "NONE"
	eRoomID = "NONE"
	eSymbol = "NONE"
	eSymbolID = "NONE"

	for symbol in data:
		symbolname = symbol["symbol_name"]
		symbolID = symbol["symbol_instance_id"]
		symboltype = symbol["symbol_type"]
		if symboltype == "room":
			eRoom = symbolname
			eRoomID = symbolID
			eSymbol = "NONE"
			eSymbolID = "NONE"
		if symboltype == "furniture":
			eSymbol = symbolname
			eSymbolID = symbolID

		symdata = symbol["forms"]
		for sd in symdata:
			form = sd["title"]
			sections = sd["sections"]
			for sec in sections:
				secname = sec["name"]
				secfields = sec["fields"]
				for scf in secfields:
					fieldname = scf["label"]
					fieldID = scf["id"]
					fieldtype = scf["type_as_string"]
					if ((fieldtype == "date") or (fieldtype == "time")):
						fieldvalue = scf["value"]["formatted"]
					else:
						fieldvalue = scf["value"]["value"]
					kvs=[]
					kvs.append(plan_id)
					kvs.append(plan_name)
					kvs.append(eRoom)
					kvs.append(eRoomID)
					kvs.append(eSymbol)
					kvs.append(eSymbolID)
					kvs.append(form)
					kvs.append(fieldname)
					kvs.append(fieldID)
					kvs.append(fieldvalue)
					if fieldvalue != None:
						kv.append(kvs)


def getplancount():
    global plan_list
    status["text"] = "Status: Getting plan count"
    window.config(cursor="watch")
    window.update()
    plan_list.clear()
    plan_list = [['id', 'name', 'street', 'city', 'country', 'createdate', 'update', 'email', 'Inspector']]
    more_pages = True
    page_number = 0
    smonth = tsmonth.get()
    emonth = temonth.get()
    sday = tsday.get()
    eday = teday.get()
    syear = tsyear.get()
    eyear = teyear.get()
    outputdir = aoutputfolder.get()
    start_date1 = datetime.datetime(int(syear), int(smonth), int(sday))
    end_date1 = datetime.datetime(int(eyear), int(emonth), int(eday))

    start_time = time.perf_counter()

    while (more_pages):
        page_number += 1
        page_number_string = str(page_number)
        get_string = "https://cloud.magicplan.app/api/v2/workgroups/plans?page="+page_number_string+"&sort=Plans.name&direction=asc"
        r = s.get(get_string)

	#create dictionary of the entire magicplan project
        dd = r.json()

        data = dd["data"]
        plans = data["plans"]
        paging = data["paging"]
        plan_count = paging["count"]
        p_len = len(plans)
        more_pages = paging["next_page"]

        # for each plan
        for p in range(p_len):
            plan_dict = plans[p]
            plan_id = plan_dict["id"]
            plan_name = plan_dict["name"]
            plan_address = plan_dict["address"]
            plan_created_by = plan_dict["created_by"]
            plan_created_by_email = plan_created_by["email"]
            plan_street = plan_address["street"]
            plan_city = plan_address["city"]
            plan_country = plan_address["country"]
            plan_create = plan_dict["creation_date"][0:10]
            try:
                plan_update = plan_dict["update_date"][0:10]
            except:
                plan_update = ""

            year = int(plan_create[0:4])
            month = int(plan_create[5:7])
            day = int(plan_create[8:10])
            try:
                plan_inspector = plan_created_by["firstname"] + " " + plan_created_by["lastname"]
            except:
                plan_inspector = "Unknown"

            plan_update = plan_dict["update_date"]
            datetimeobj = datetime.datetime(year,month,day,0,0,0)
            if ((datetimeobj >= start_date1) and (datetimeobj < end_date1)):
                pl = []
                pl.append(plan_id)
                pl.append(plan_name)
                pl.append(plan_street)
                pl.append(plan_city)
                pl.append(plan_country)
                pl.append(plan_create)
                pl.append(plan_update)
                pl.append(plan_created_by_email)
                pl.append(plan_inspector)
                plan_list.append(pl)


    end_time = time.perf_counter()
    pl_len = len(plan_list)
    plan_time = round(end_time - start_time,2)
    anumberofplans["text"] = ": "+ str(plan_count)
    aoutput2["text"] = ": "+str(pl_len-1)
    btn_export["state"] = "active"
    status["text"] = "Status: Finished getting plan count"

    try:
        #create csv file for plan list
        with open(outputdir +'\\planlist.csv', 'w', newline='') as f3:
            # using csv.writer method from CSV package
            write = csv.writer(f3)
            for t in plan_list:
                try:
                    write.writerow(t)
                except:
                    errstr = 'MP Error - ' + t[0] + '\n'
                    print(errstr)
                    tracestr = traceback.format_exc() + '\n'
                    print(tracestr)
    except Exception as error:
        status["text"] = "Status: Writing Plan List failed because of "+error.strerror
    window.config(cursor="")
    window.update()
    return [plan_count,plan_time]

def startExport():
    status["text"] = "Status: Exporting plans"
    window.config(cursor="watch")
    window.update()
    hold = True
    pl_len = len(plan_list)


    start_time_all = time.perf_counter()
    start_time = start_time_all


    for i in range(pl_len):
        if i == 0: #skip header
            continue
        plan_dict = plan_list[i]
        plan_id = plan_dict[0]
        plan_name = plan_dict[1]
        plan_u_id = plan_name + " " + plan_id

        get_string = "https://cloud.magicplan.app/api/v2/plans/forms/"+plan_id

        error = "no"
        er = ""
        try:
            er = "get"
            r = s.get(get_string)
            if r.status_code == 200:
               er = "process"
               extractXML(r,plan_id,plan_name)
            else:
                error = "yes"
                fd = []
                fd.append(plan_id)
                fd.append(plan_name)
                fd.append("error downloading from Magiplan")
                failed_export.append(fd)
        except:
            error = "yes"
            fd = []
            fd.append(plan_id)
            fd.append(plan_name)
            fd.append("error processing file")



        end_time = time.perf_counter()
        x_time = round(end_time - start_time,1)
        tx_time = round((end_time - start_time_all),1)
        avg_time = round(tx_time / i,2)
        files_left = pl_len - i -1
        remain_time_seconds = files_left * avg_time
        remain_time_minutes = round(remain_time_seconds/60,2)
        aoutput4["text"] = ": "+plan_name
        aoutput5["text"] = ": "+str(x_time)+" seconds"
        aoutput6["text"] = ": "+str(i)
        if tx_time < 120:
            aoutput7["text"] = ": "+str(tx_time)+" seconds"
        else:
            aoutput7["text"] = ": "+str(round(tx_time/60,2))+" minutes"
        aoutput7a["text"] = ": "+str(avg_time)+" seconds"
        aoutput8["text"] = ": "+str(remain_time_minutes)+" minutes"
        aoutput8a["text"] = ": "+str(files_left)
        start_time = end_time
        window.update()
    status["text"] = "Status: Finished Exporting plans"

    # save the Values
    outputdir = aoutputfolder.get()
    try:
        with open(outputdir +'\\mp_data.csv', 'w', newline='', encoding='utf-8-sig') as f1:
            # using csv.writer method from CSV package
            write = csv.writer(f1)
            for m in kv:
                try:
                    write.writerow(m)
                except:
                    errstr = 'MP Error - ' + m[0] + '\n'
                    print(errstr)
                    tracestr = traceback.format_exc() + '\n'
                    print(tracestr)
    except Exception as error:
        status["text"] = "Status: Writing Custom Field List failed because of " + error.strerror

    try:
        with open(outputdir +'\\plan_fail.csv', 'w', newline='', encoding='utf-8-sig') as f2:
            # using csv.writer method from CSV package
            write = csv.writer(f2)
            for m in failed_export:
                try:
                    write.writerow(m)
                except:
                    errstr = 'MP Error - ' + m[0] + '\n'
                    print(errstr)
                    tracestr = traceback.format_exc() + '\n'
                    print(tracestr)
    except Exception as error:
        status["text"] = "Status: Writing Failed plans List failed because of " + error.strerror
    window.config(cursor="")
    window.update()


#-------------------------------------------------START-------------------------------

# create a tkinter window
# Create Object
# Open window having dimension 100x100
window.geometry('600x500')


#greeting.place(x =200, y = 10)
exportstart.place(x=10,y=20)
tsmonth.place(x =90, y = 20)
exportstart1.place(x=105, y = 20)
tsday.place(x = 115, y = 20)
exportstart2.place(x=130, y = 20)
tsyear.place(x = 140, y = 20)

exportend.place(x=10,y=40)
temonth.place(x =90, y = 40)
exportend1.place(x=105, y = 40)
teday.place(x = 115, y = 40)
exportend2.place(x=130, y = 40)
teyear.place(x = 140, y = 40)

outputfolder.place(x=10,y=60)
aoutputfolder.place(x=170,y=60)

btn_getnumberofplans = tk.Button(master=window, text="Get Plan Count", command=getplancount)
btn_getnumberofplans.place(x = 25,y = 90)
numberofplans.place(x = 10, y = 120)
output2.place(x = 10, y = 140)
anumberofplans.place(x = 170, y = 120)
aoutput2.place(x = 170, y = 140)
btn_export = tk.Button(master=window, text="Export Plans", command=startExport)
btn_export["state"] = "disabled"
btn_export.place(x =25, y = 170)
output4.place(x = 10, y = 200)
output5.place(x = 10, y = 220)
output6.place(x = 10, y = 240)
output7.place(x = 10, y = 260)
output7a.place(x = 280, y = 260)
output8.place(x = 10, y = 280)
output8a.place(x = 280, y = 280)
aoutput4.place(x = 170, y = 200)
aoutput5.place(x = 170, y = 220)
aoutput6.place(x = 170, y = 240)
aoutput7.place(x = 170, y = 260)
aoutput7a.place(x = 420, y = 260)
aoutput8.place(x = 170, y = 280)
aoutput8a.place(x = 420, y = 280)
status.place(x=10, y = 470)

window.mainloop()








