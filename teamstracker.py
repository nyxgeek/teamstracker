#!/usr/bin/env python3
# Teams Enum
# 2022 @nyxgeek - TrustedSec

import csv
import requests
from requests.exceptions import ConnectionError, ReadTimeout, Timeout
import datetime
from datetime import date
import os
import time
import sys
import re
import threading
from threading import Semaphore
import sqlite3

from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

requests.packages.urllib3.disable_warnings()

import argparse
import json
import array

sqldb_location = 'data/teamstracker.db'


def sql_insert_presence(uuid, email_address, displayname, availability, ooo_enabled, statusmsg, currenttime, currentdate, qh_period, hh_period): 
    try:
        conn = sqlite3.connect(sqldb_location)
        sql_query = f"INSERT OR IGNORE INTO presence_log (uuid, emailaddress, displayname, availability, ooo_enabled, statusmsg, currenttime, currentdate, qh_period, hh_period) VALUES ('{uuid}','{email_address}', '{displayname}', '{availability}', '{ooo_enabled}', '{statusmsg}', '{currenttime}', '{currentdate}', '{qh_period}','{hh_period}');"
        conn.execute(sql_query)
        conn.commit()
        conn.close()
    except Exception as e: print(e)
    except:
        print("Some SQLite error in sql_insert_user! Maybe write some better logging next time.")

def sql_insert_ooo(objectidline, emailaddress, displayname, ooo_text, currentdate, currenttime): 
    try:
        conn = sqlite3.connect(sqldb_location)
        print("inserting")
        #ooo_text = "safe_text"
        sql_query = f"REPLACE INTO ooo_log (uuid, emailaddress, displayname, ooo_text, currentdate, currenttime) VALUES ('{objectidline}', '{emailaddress}', '{displayname}', '{ooo_text}', '{currentdate}', '{currenttime}');"
        conn.execute(sql_query)
        conn.commit()
        conn.close()
    except Exception as e: print(e)
    except:
        print("Some SQLite error in sql_insert_user! Maybe write some better logging next time.")



writeLock = Semaphore(value = 1)

#with open('teams_headers.txt', 'r') as file:
#    teams_header = file.read().rstrip()

# initiate the parser
parser = argparse.ArgumentParser()
parser.add_argument("-c", "--csv", help="csv export of email addresses from TeamFiltration")
parser.add_argument("-u", "--uuid", help="Azure UUID to target")
parser.add_argument("-U", "--uuidfile", help="file containing Azure UUIDs of users to target")
parser.add_argument("-v", "--verbose", help="enable verbose output", action='store_true')
parser.add_argument("-T", "--threads", help="total number of threads (defaut: 50)")

if len(sys.argv)==1:
    parser.print_help(sys.stderr)
    sys.exit(1)


username = "FakeUser"
domain = "Placeholder"
verbose = False
isUser = False
isUserFile = False
isCSVFile = False
failedList = []

# read arguments from the command line
args = parser.parse_args()


if args.verbose:
    verbose = True

if args.csv:
    isCSVFile = True

if args.uuid:
    #print("Checking username: %s" % args.username)
    username = args.uuid
    isUser = True

if args.uuidfile:
    #print("Reading users from file: %s" % args.userfile)
    global userfile
    userfile = args.uuidfile
    #checkUserFile()
    global isUserfile
    isUserFile = True

if args.threads:
    thread_count = args.threads
else:
    thread_count = 50


def requests_retry_session(
    retries=3,
    backoff_factor=0.3,
    status_forcelist=(500, 502, 504),
    session=None,
):
    session = session or requests.Session()
    retry = Retry(
        total=retries,
        read=retries,
        connect=retries,
        backoff_factor=backoff_factor,
        status_forcelist=status_forcelist,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session



### BANNER GOES HERE

print("\n\n*******************************************************************************************")
print("**************************   Microsoft Teams Presence Logger   ****************************")
print("*************************** 2023.04.28 - @nyxgeek - TrustedSec*****************************")
print("**************************************** v 0.3 ********************************************")
print("*******************************************************************************************\n\n")


#@retry
def checkURL(objectidline, name='', email=''):
    #global r
    global domain
    objectidline = (objectidline.rstrip())

    header = {"Host":"graph.office.net", "Sdkversion":"GraphExplorer/4.0", "Content-Type":"application/json", "Authorization":"Bearer {token:https://graph.microsoft.com/}", 
"Origin":"https://developer.microsoft.com", "Referer":"https://developer.microsoft.com/", "Accept-Encoding":"gzip, deflate", "Accept-Language":"en-US,en;q=0.9" }
    url = 'https://graph.office.net/en-us/graph/api/proxy?url=https%3A%2F%2Fgraph.microsoft.com%2Fbeta%2Fusers%2F' + objectidline + '%2Fpresence'
    #print(url)

    requests.packages.urllib3.disable_warnings()


    try:
        r = requests_retry_session().get(url, headers=header, timeout=5.0)
        #print(r)
        #print(r.content)

        if r.status_code == 200:

            displayname = name
            emailaddress = email


            currenttime = str(int(time.time()))
            currentdate = str(date.today())
            now = datetime.datetime.now()
            #print(str(int(now)))
            totalminutes = now.hour*60 + now.minute
            #print(now.hour*60)
            #print(now.minute)
            #print(totalminutes)

            # this is the quarter-hour period of the day, and half-hour period of the day
            qh_period = int(totalminutes/15)
            hh_period = int(totalminutes/30)
            #print("Here is {0}".format(qh_period))


            #print(objectidline)
            #print(len(r.content))
            if ( (len(r.content) > 2 )):
                result_content=r.content.decode("utf-8")
                #print(result_content)
                decoded_results=json.loads(result_content)
                #print(decoded_results)

                if ( str(decoded_results['availability']) != 'PresenceUnknown' ):
                    availability = str(decoded_results['availability'])
                    activity = str(decoded_results['activity'])
                    statusmsg = str(decoded_results['statusMessage'])

                    ooo = str(decoded_results['outOfOfficeSettings'])



                    #chopping down to 3997 chars max - arbitrary choice. could increase but don't go past db column max
                    ooo = (ooo[:3997] + '..') if len(ooo) > 3997 else ooo
                    if len(ooo) > 30:
                        print("We have an OOO!")
                        #print(ooo)

                        ooo_text = (str(decoded_results['outOfOfficeSettings']['message']))
                        ooo_text = ooo_text.strip().replace("'","").replace("\n","").replace("|","")
                        #print(f"New ooo: {ooo_text}")

                        ooo_isOOO = (str(decoded_results['outOfOfficeSettings']['isOutOfOffice']))

                        print("Are we ooo? {0}".format(ooo_isOOO))
                        if ooo_isOOO:
                            ooo_enabled = 1
                        else:
                            ooo_enabled = 0

                    else:
                        ooo_enabled = 0

                    #print(f"DisplayName: {displayname}, Email: {emailaddress}, UUID: {objectidline}, Presence: {availability}, OOO: {ooo}")

                    #print("Currentdate is: {0}".format(currentdate))
                    sqldata = (objectidline, emailaddress, displayname, availability, ooo_enabled, statusmsg, currenttime, currentdate, qh_period, hh_period)
                    print(sqldata)
                    sql_insert_presence(objectidline, emailaddress, displayname.replace("'","''"), availability, ooo_enabled, statusmsg, currenttime, currentdate, qh_period, hh_period)

                    if ooo_enabled == 1:
                        print("**** going to insert ooo data now ***")
                        sql_insert_ooo(objectidline, emailaddress, displayname.replace("'","''"), ooo_text, currentdate, currenttime)

                else:
                    print("Adding bad ID to array: " + objectidline)
                    failedList.append(objectidline)
        else:
            RESPONSE = "[?] [" + str(r.status_code) + "] UNKNOWN RESPONSE"




#################

    except requests.ConnectionError as e:
        print("Error: %s" % e)
        time.sleep(3)
    except requests.Timeout as e:
        print("Error: %s" % e)
        #print("Read Timeout reached, sleeping for 3 seconds")
        time.sleep(3)
    except requests.RequestException as e:
        print("Error: %s" % e)
        time.sleep(3)
    except requests as e:
        print("Well, I'm not sure what just happened. Onward we go...")
        print(e)
        time.sleep(3)


def checkUserFile():

    currenttime=datetime.datetime.now()

    f = open(userfile)
    listthread=[]
    for userline in f:
        #if threading.active_count() < thread_count:
        while int(threading.active_count()) >= int(thread_count):
            #print "We have enough threads, sleeping."
            time.sleep(1)

        #print "Spawing thread for: " + userline + " thread(" + str(threading.active_count()) +")"
        x = threading.Thread(target=checkURL, args=(userline,))

        listthread.append(x)
        x.start()

    f.close()

    for i in listthread:
        i.join()
    return


def checkUser(username):
    checkURL(username)


def checkCSVFile():

    currenttime=datetime.datetime.now()

    filename = args.csv
    listthread=[]

    with open(filename, 'r') as csvfile:
        datareader = csv.reader(csvfile, delimiter=';')
        next(datareader)
        for row in datareader:
            try:
                email = row[1]
                displayname = row[3]
                uuid = row[2]
                #print(f"Here is: {email}, for {displayname} with uuid: {uuid}")
            except:
                print("doh")

            #if threading.active_count() < thread_count:
            while int(threading.active_count()) >= int(thread_count):
                #print "We have enough threads, sleeping."
                time.sleep(1)

            #print "Spawing thread for: " + uuid + " thread(" + str(threading.active_count()) +")"
            x = threading.Thread(target=checkURL, args=(uuid,displayname,email))

            listthread.append(x)
            x.start()


        for i in listthread:
            i.join()
        return


def testConnect():
    requests.packages.urllib3.disable_warnings()
    if isUser:
        checkURL(username)
    if isUserFile:
        checkUserFile()
    if isCSVFile:
        checkCSVFile()


testConnect()


