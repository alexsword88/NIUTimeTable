#Table Capture Require Library
import re
import os
import time
import json
from pathlib import Path
from datetime import datetime,timedelta
from selenium import webdriver
from bs4 import BeautifulSoup
from lxml import etree
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException,NoSuchFrameException, WebDriverException
#Google Require Library
import httplib2
from apiclient import discovery ,errors
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

#driver = webdriver.PhantomJS()
#datafilepath = Path(os.path.abspath("國立宜蘭大學教務行政資訊系統_files/portal.html")).as_uri()
#driver.get(datafilepath)
SCOPES = "https://www.googleapis.com/auth/calendar"
CLIENT_SECRET_FILE = "client_id.json"
APPLICATION_NAME = "NIU Time Table"

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,'NIUTT.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        credentials = tools.run_flow(flow, store)
        #print('Storing credentials to ' + credential_path)
    return credentials

def weekdayconvert(weekday,startdate):
    startdayadjust = {"MO":0,"TU":1,"WE":2,"TH":3,"FR":4,"SA":5,"SU":6}
    startdateweekday = startdate.weekday()
    adjustdays = startdayadjust[weekday] - startdateweekday
    return adjustdays if adjustdays >= 0 else adjustdays+7

def tag2timedate(NIU_timetag):
    timestart = {"1":"08:10:00","2":"09:10:00","3":"10:10:00","4":"11:10:00","5":"13:10:00","6":"14:10:00","7":"15:10:00","8":"16:10:00","9":"17:10:00","A":"18:20:00","B":"19:15:00","C":"20:10:00","D":"21:05:00"}
    timeend = {"1":"09:00:00","2":"10:00:00","3":"11:00:00","4":"12:00:00","5":"14:00:00","6":"15:00:00","7":"16:00:00","8":"17:00:00","9":"18:00:00","A":"19:10:00","B":"20:05:00","C":"21:00:00","D":"21:55:00"}
    weekday = {"1":"MO","2":"TU","3":"WE","4":"TH","5":"FR","6":"SA","7":"SU"}
    returnformat = "T{}+08:00"
    alltimelist = NIU_timetag.split(",")
    timeweekdata = {}
    for time in alltimelist:
        time_week_split = time.split("0")
        if time_week_split[0] not in timeweekdata:
            timeweekdata[time_week_split[0]]=[]
        timeweekdata[time_week_split[0]].append(time_week_split[1])
    jsonresult = []
    jsontemp = {}
    if len(timeweekdata.keys()) == 1:
        key = list(timeweekdata.keys())[0]
        jsontemp["weekday"] = weekday[key]
        starttime = returnformat.format(timestart[timeweekdata[key][0]])
        endtime = returnformat.format(timeend[timeweekdata[key][-1]])
        jsontemp["timerange"] = [starttime,endtime]
        jsonresult.append(jsontemp)
    else:
        for key in timeweekdata.keys():
            jsontemp = {}
            jsontemp["weekday"] = weekday[key]
            starttime = returnformat.format(timestart[timeweekdata[key][0]])
            endtime = returnformat.format(timeend[timeweekdata[key][-1]])
            jsontemp["timerange"] = [starttime,endtime]
            jsonresult.append(jsontemp)
    return jsonresult


def table_capture():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-infobars")
    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.get("https://acade.niu.edu.tw/NIU/Default.aspx")
    driver.implicitly_wait(2)
    table = ""
    place_n_time = []
    while True:
        try:
            try:
                driver.switch_to.frame(driver.find_element_by_css_selector("#topFrame #downFrame #rightFrame #pageFrame frame[name=mainFrame]"))
            except (NoSuchElementException,WebDriverException):
                #print("Frame Not Found")
                driver.switch_to_default_content()
                continue
            try:
                table = driver.find_element_by_css_selector("#Span4")
            except (NoSuchElementException,WebDriverException):
                #print("Title Not Found")
                driver.switch_to_default_content()
                continue
            if table.text.find("學生個人選課清單課表列印") != -1:
                try:
                    table = driver.find_element_by_css_selector("#DataGrid")
                except NoSuchElementException:
                    print("Table Not Found")
                    driver.switch_to_default_content()
                    continue
                classlinks = driver.find_elements_by_css_selector("#DataGrid a[id*=DataGrid_ct]")
                table = driver.page_source
                time.sleep(0.5)
                for index in range(0,len(classlinks)):
                    driver.find_elements_by_css_selector("#DataGrid a[id*=DataGrid_ct]")[index].click()
                    while len(driver.window_handles)<2:
                        time.sleep(0.5)
                    driver.switch_to.window(driver.window_handles[1])
                    driver.switch_to.frame(driver.find_element_by_css_selector("#topFrame frameset[name=downFrame] #rightFrame #pageFrame frame[name=mainFrame]"))
                    time.sleep(0.5)
                    place_n_time.append(driver.page_source)
                    driver.close()
                    time.sleep(0.5)
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(0.5)
                    driver.switch_to.frame(driver.find_element_by_css_selector("#topFrame #downFrame #rightFrame #pageFrame frame[name=mainFrame]"))
                break
        except UnexpectedAlertPresentException:
            #print("Alert Box Appear")
            pass
    driver.quit()

    print("Time Table Detected!")
    time_table = BeautifulSoup(table, 'lxml')
    table = time_table.find_all("table",{"id":"DataGrid"})[0].find_all("tr")[1:]

    jsondata = {"calender_added":0,"timetable_data":{}}
    for row in table:
        tddata = row.find_all("td")
        tablerow = []
        for rowdata in tddata:
            tablerow.append(rowdata.text.strip())
        reqtablerow = {"class_name":tablerow[3],
                       "class_teacher":tablerow[6]}
        jsondata["timetable_data"][tablerow[2]]=reqtablerow

    for place_n_time_source in place_n_time:
        table_pt = BeautifulSoup(place_n_time_source, 'lxml')
        place_time_table = table_pt.select(".welcome table")[0].find_all("td")
        requiredata = ["課程代碼","上課時間","上課地點"]
        record = False
        popdata = ""
        collecteddata = {}
        for tddata in place_time_table:
            try:
                if tddata["class"][0] == "welcome":
                    if record:
                        collecteddata[popdata] = tddata.text.strip()
                        record = False
                        if len(requiredata) == 0:
                            break
                if tddata["class"][0] == "table_title":
                    if tddata.text.strip() in requiredata:
                        popdata = requiredata.pop(requiredata.index(tddata.text.strip()))
                        record = True
            except KeyError:
                print("Class Attr Not Found")
        jsondata["timetable_data"][collecteddata["課程代碼"]]["class_time"]=collecteddata["上課時間"]
        jsondata["timetable_data"][collecteddata["課程代碼"]]["class_place"]=collecteddata["上課地點"]

    with open("timetable.json","w") as file:
        file.write(json.dumps(jsondata,indent=4))

def calendar_event_insert(service,timetable_json):
    try:
        startdate = None
        enddate = None
        startdate = datetime.strptime(timetable_json["startdate"],"%Y%m%d")
        print("Start Date OK")
        enddate = datetime.strptime(timetable_json["enddate"],"%Y%m%d")
        print("End Date OK")
    except (KeyError,ValueError):
        try:
            while True:
                try:
                    if not startdate:
                        startdate = datetime.strptime(input("Start Date(YYYYMMDD):"),"%Y%m%d")
                        print("Start Date OK")
                    break
                except ValueError:
                    print("Start Date Wrong input Format(YYYYMMDD)")
                    print("Y=Year, M=Month, D=Day")
            timetable_json["startdate"] = startdate.strftime("%Y%m%d")
            file.seek(0)
            file.truncate()
            file.write(json.dumps(timetable_json,indent=4))
            while True:
                try:
                    if not enddate:
                        enddate = datetime.strptime(input("End Date(YYYYMMDD):"),"%Y%m%d")
                        print("End Date OK")
                    break
                except ValueError:
                    print("End Date Wrong input Format(YYYY-MM-DD)")
                    print("Y=Year, M=Month, D=Day")
            timetable_json["enddate"] = enddate.strftime("%Y%m%d")
            file.seek(0)
            file.truncate()
            file.write(json.dumps(timetable_json,indent=4))
        except IOError:
            print("\nWrite file Failed")
        except KeyboardInterrupt:
            print("\nUser Exit")
            quit()
    timetable_data = timetable_json["timetable_data"]
    for classdataindex in timetable_data.keys():
        classdata = timetable_data[classdataindex]
        timedata = tag2timedate(classdata["class_time"])
        for calendarevent in timedata:
            newstartdate = startdate + timedelta(days=weekdayconvert(calendarevent["weekday"],startdate))
            bodyparameter ={"start": {
                                "dateTime": "{}{}".format(newstartdate.strftime("%Y-%m-%d"),calendarevent["timerange"][0]),
                                "timeZone": "Asia/Kuala_Lumpur"},
                            "end": {
                                "dateTime": "{}{}".format(newstartdate.strftime("%Y-%m-%d"),calendarevent["timerange"][1]),
                                "timeZone": "Asia/Kuala_Lumpur"},
                            "recurrence": ["RRULE:FREQ=WEEKLY;UNTIL={}T160000Z;BYDAY={}".format(enddate.strftime("%Y%m%d"),calendarevent["weekday"])],
                            "reminders": {"useDefault": False},
                            "summary": classdata["class_name"],
                            "description": classdata["class_teacher"],
                            "location": classdata["class_place"]}
            event = service.events().insert(calendarId='primary', body=bodyparameter).execute()
            #print(event)
            print("Class:\"{}\" Added".format(classdata["class_name"]))
            with open("calendarID","a") as cidfile:
                cidfile.write(event.get("id")+"\n")
        if timetable_json["calender_added"] == 0:
            timetable_json["calender_added"] = 1
            file.seek(0)
            file.truncate()
            file.write(json.dumps(timetable_json,indent=4))
    print("All Class Add Completed")

def calendar_event_delete(service,timetable_json):
    if os.path.exists("calendarID"):
        with open("calendarID","r") as cidfile:
            cids = cidfile.read().splitlines()
            for cid in cids:
                if cid.strip() == "":
                    continue
                try:
                    service.events().delete(calendarId='primary', eventId=cid).execute()

                except errors.HttpError as error:
                    if json.loads(error.content.decode("utf-8"))["error"]["message"] == "Resource has been deleted":
                        continue
                    else:
                        print(json.loads(error.content.decode("utf-8"))["error"]["message"])
                        print("ERROR")
                        quit()
        os.remove("calendarID")
        timetable_json["calender_added"] = 0
        file.seek(0)
        file.truncate()
        file.write(json.dumps(timetable_json,indent=4))
    else:
        print("calendarID File Not exists!")

if __name__ == "__main__":
    if not os.path.exists("timetable.json"):
        print("Timetable json data no exists")
        table_capture()
    with open("timetable.json","r+") as file:
        timetable_json = json.loads(file.read())
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('calendar', 'v3', http=http)
        if timetable_json["calender_added"] == 0:
            calendar_event_delete(service,timetable_json)
            calendar_event_insert(service,timetable_json)
        else:
            flag_del = input("Calendar Already exists,Want to Delete the previous data?[Y/N]:")
            flag_del = flag_del.upper()
            if flag_del == "Y":
                print("Start Delete Calendar")
                calendar_event_delete(service,timetable_json)
                print("Delete Completed")
            else:
                print("No Delete Any Data")

"""#table write to csv file function
colname = "序號,學期,課號,課名,開課單位,年級班別,授課老師,老師單位,學分,選別,人數,人數限制上/下限,實習,時數,合開"
with open("class.csv","w",encoding="big5") as file:
    file.write(colname+"\n")
    for row in table:
        tddata = row.find_all("td")
        print(len(tddata))
        count = 0
        for rowdata in tddata:
            file.write(re.sub(",","/",rowdata.text.strip()))
            if count != len(tddata)-1:
                file.write(",")
            else:
                file.write("\n")
            count += 1
"""
