# -*- coding: utf-8 -*-
"""
Created on Thu Sep 26 14:24:04 2019

@author: haywoojr
"""

import json
from pathlib import Path

#for importing the config file
path = Path(__file__).parent / "config.json"
#print(path)
with path.open() as config_file:
    config_data = json.load(config_file)

userInfo=config_data['userInfo']
#for decrypting password
from cryptography.fernet import Fernet
key = userInfo['key']
password=Fernet(key).decrypt(userInfo['password'].encode('utf-8'))
userInfo['password']=password.decode('utf-8')

import pyodbc
#import datetime
import pandas as pd
from pandas.tseries.offsets import BDay
import datetime
import dateutil.parser as parser
import win32com.client
import shutil,os,time
from shareplum import Site
from shareplum import Office365

#open the sharepoint list to edit
authcookie = Office365('yoursharepointsite', 
                       username=userInfo['username'], 
                       password=userInfo['password']).GetCookies()
site = Site('yoursharepointlistsite',
            authcookie=authcookie)

listName='TxTimeOutDates'

sp_list = site.List(listName)
#data holds all record (I think)
#data = sp_list.GetListItems('PythonView')
#fields have to have a default value and something in all rows or the column won't be returned
data = sp_list.GetListItems(fields=['ID','Title','PatientId','NumTimeoutsNeeded',
                                    'NumTimeoutsPerformed',
                                    'MissingTimeouts','ActuallyMissingTimeouts',
                                    'DatesTimeoutsMissed','Comments','Approved'])

#directory to store files locally
tempDir=Path('C:/Temp3/finals/')

# Create connection
con = pyodbc.connect(driver="{SQL Server}",server='servername',database='variansystem',uid='username',pwd='password',autocommit=True)
cur = con.cursor()
#get patients that have had a final approved in this date range
n=datetime.datetime(2021,1,1)
fnlstrtdate=n.strftime("%m/%d/%y")
n=datetime.datetime(2021,3,31)
fnlenddate=n.strftime("%m/%d/%y")
#check for treatments 3 months before the fnl date? maybe
n=datetime.datetime(2021,1,1)
strtdate=n.strftime("%m/%d/%y")
n=datetime.datetime(2021,3,31)
enddate=n.strftime("%m/%d/%y")
#all completed final physics checks
getPatientList = """
select distinct Patient.PatientId,Patient.LastName,Patient.FirstName,
NonScheduledActivity.WorkFlowActiveFlag,Activity.ActivityCode
from Patient,NonScheduledActivity,ActivityInstance,Activity
where (NonScheduledActivity.DueDateTime between '{strtdate}' and dateadd(dd,1,'{enddate}'))
and (Patient.PatientSer = NonScheduledActivity.PatientSer)
and (ActivityInstance.ActivityInstanceSer=NonScheduledActivity.ActivityInstanceSer)
and (ActivityInstance.ActivityInstanceRevCount = NonScheduledActivity.ActivityInstanceRevCount)
and (ActivityInstance.ActivitySer=Activity.ActivitySer)
and (Activity.ActivityCode like 'Final Physics%')
--and (NonScheduledActivity.WorkFlowActiveFlag like '%0%')
and (NonScheduledActivity.NonScheduledActivityCode like '%Completed%')
order by Patient.PatientId
""".format(strtdate=fnlstrtdate,enddate=fnlenddate)

#
PtTab=cur.execute(getPatientList)
#Put ptTab into python object not pyodbc opbject
ptsListIDs=[]
ptsListName=[]
ptFinalDocs=[]
ptTimeouts=[]
ptNoTimeOutDoc=[]
ptMissingTimeOuts=[]
AppList=[]

for row in PtTab:
    #print(row[0],row[4])
    ptsListIDs.append(row[0])
    ptsListName.append((row[1],row[2]))
print("len of patients= ",len(ptsListIDs))
#print(ptsListIDs)
#get word ready
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
#loop through patient to get their treatment timeout document and days they had appointments
k=0
for ptID in ptsListIDs:
    #this variable chooses between updating and creating a new item
    update=False
    datesString=['-']
     # reset counters
    numTimeoutsNeeded=0
    numTimeoutsPerformed=0
    #a default patient
    #think about making the Title a date?
    thisPatient={'Title':str(strtdate)+"-"+str(enddate),
                'PatientId': str(ptID),
                'NumTimeoutsNeeded': 0,
                'NumTimeoutsPerformed': 0,
                'MissingTimeouts': "Yes",
                'ActuallyMissingTimeouts': "Yes",
                'DatesTimeoutsMissed': "-",
                'Comments': "-",
                'Approved':"No"}
    
    for item in data:
        if item['PatientId']==str(ptID) and item['Title']==str(strtdate)+"-"+str(enddate):
            update=True
            thisPatient=item
            print('found matching patient in list '+str(ptID))
    datesString.append(thisPatient['DatesTimeoutsMissed'])
    #change the title on all 
    thisPatient['Title']=str(strtdate)+"-"+str(enddate)
     # reset counters to what is there ? maybe
    # numTimeoutsNeeded=thisPatient['NumTimeoutsNeeded']
    # numTimeoutsPerformed=thisPatient['NumTimeoutsPerformed']

    #get patient appointments/ currently ignores Simulation and new starts
    getPtApp="""
select distinct Patient.PatientId,Patient.LastName,Patient.FirstName,
ScheduledActivity.ScheduledStartTime,ScheduledActivity.ActualStartDate,
ScheduledActivity.ScheduledActivityCode,Hospital.HospitalName,Department.DepartmentName,
Activity.ActivityCode,Activity.ObjectStatus,Patient.PatientSer,Machine.MachineId,vv_ActivityLng.Expression1
from Activity, vv_ActivityLng,Department,Hospital,Patient,ScheduledActivity,Machine,ResourceActivity,ActivityInstance
where Patient.PatientId ='{patid}' and
(Patient.PatientSer=ScheduledActivity.PatientSer) and
(ActivityInstance.DepartmentSer=Department.DepartmentSer) and
(ScheduledActivity.ActivityInstanceSer=ActivityInstance.ActivityInstanceSer) and
(Department.HospitalSer=Hospital.HospitalSer) and
(ActivityInstance.ActivitySer=Activity.ActivitySer) and
(Activity.ActivityCode=vv_ActivityLng.LookupValue) and
(Machine.ResourceSer=ResourceActivity.ResourceSer) and
(ScheduledActivity.ScheduledActivitySer=ResourceActivity.ScheduledActivitySer) and
((ScheduledActivity.ObjectStatus='Active') and (ScheduledActivity.ScheduledStartTime between '{strtdate}' and dateadd(dd,1,'{enddate}')))
and (ScheduledActivity.ScheduledActivityCode like '%Complete%')
and (vv_ActivityLng.SubSelector = 1)
--removed comment for counting new starts 
--and (vv_ActivityLng.Expression1 not like 'New%')
and (vv_ActivityLng.Expression1 not like '%QA%') and (vv_ActivityLng.Expression1 not like '%Physician%')
and (vv_ActivityLng.Expression1 not like '%IMRT%')  
and (vv_ActivityLng.Expression1 not like '%Simulation%')
and (vv_ActivityLng.Expression1 not like '%PerFr%')    
    """.format(patid=ptID,strtdate=strtdate,enddate=enddate)
    ptAppTab=cur.execute(getPtApp)
    ptAppList=[]
    
    for ptApp in ptAppTab:
        ptAppList.append(ptApp[3])
        
    getPtDocuments="""--Patient Documents
select distinct CONVERT(Varchar(20), note_tstamp, 100) as v11note_tstamp,note_typ_desc,template_name,visit_note.appr_flag,
CONVERT(VARCHAR(20), visit_note.appr_tstamp, 100) as appr_tstamp,FileLocation.DriveName,FileLocation.FolderName1,FileLocation.FileName,visit_note.valid_entry_ind
--ServerName,DriveName,
--upper(FolderName1) as FolderName1,doc_file_loc
from varianenm.dbo.visit_note,variansystem.dbo.FileLocation,varianenm.dbo.pt,variansystem.dbo.Patient,
varianenm.dbo.note_typ
where  (variansystem.dbo.Patient.PatientId ='{patid}')
and (variansystem.dbo.Patient.PatientSer=varianenm.dbo.pt.patient_ser)
and (varianenm.dbo.visit_note.pt_id=varianenm.dbo.pt.pt_id)
and (varianenm.dbo.visit_note.doc_file_loc=variansystem.dbo.FileLocation.FileName)
and (varianenm.dbo.visit_note.note_typ=varianenm.dbo.note_typ.note_typ)
and (visit_note.valid_entry_ind like 'Y')
and ((note_typ_desc like '%Timeout%'))
order by FileLocation.FileName asc
""".format(patid=ptID)
    ptDocTab=cur.execute(getPtDocuments)
    #looks for all treatment timeout documents add them to the list and adds none if no doc exists
    #reset the ptFinalDocs so I am only looking at this patients docs
    ptFinalDocs=[]
    TimeoutDocFound=True
    try:
        i=0
        for rows in ptDocTab:
            i+=1
            #print(i,ptID,rows[5],rows[6],rows[7])
            temp=(ptID,Path(rows[5]+"/"+rows[6]+"/"+rows[7]),rows[7],rows[3])
            ptFinalDocs.append(temp)
    except:
        temp=(ptID,None,None,None)
        ptFinalDocs.append(temp)
        #print(ptID+' Missing timeout document')
    #check if got here and no final doc
    if(len(ptFinalDocs)==0):
        temp=(ptID,None,None,None)
        ptFinalDocs.append(temp)
        #print(ptID+' Missing timeout document')

    #while still in the ptid loop, find all the dates that the timeout document has in it
    print("len of final doc= ",len(ptFinalDocs))
    
    #loop through final documents to get number of timeouts performed
    
    ptTimeOutDates=[]

    #first copy the file to local temp so not messing with the one on the server
    
    for i in range(len(ptFinalDocs)):
        #print(ids,str(fl))
        ids=ptFinalDocs[i][0]
        fl=ptFinalDocs[i][1]
        fname=ptFinalDocs[i][2]
        if fname is not None:
            shutil.copyfile(str(fl),str(tempDir)+"\\"+ids+"-"+fname)
    
            wb=word.Documents.Open(str(tempDir)+"\\"+ids+"-"+fname)
        
            doc=word.ActiveDocument
            numTables=doc.Tables.Count
            numTimeOuts=0
            #TODO Need to make this table in tables only reads tables could search for paragraphs
            for k in range(numTables):
                table=doc.Tables(k+1)
                cont=table.Cell(Row=1,Column=2).Range.Text
                tableString=" ".join(cont.split())
                #print(cont)
                if tableString!='\x07':
                    #know that it is not empty try to parse the text as a date
                    dateText=tableString.split(" ")
                    numTimeOuts+=1
                    try:
                        ptTimeOutDates.append(parser.parse(dateText[0]))
                    except:
                        #could not parse will need to be checked by hand
                        continue
            word.Documents.Close()
            #temp=(ids,numTimeOuts)
            #ptTimeouts.append(temp)
            os.remove(str(tempDir)+"\\"+ids+"-"+fname)
        else:
            ptNoTimeOutDoc.append((ids,'No Timeout Document Found!'))
            TimeoutDocFound=False
            print(ptID+' Missing timeout document')
    
    #no compare all the dates and see if any are missing
   
    
    missingYesNo="No"
    approvedYesNo="Yes"
    
    delIndex=0
    for date in ptAppList:
        print('len ptTimeOutDates ',len(ptTimeOutDates), 'len ptAppList', len(ptAppList))
        thisDate=date.strftime('%m/%d/%Y')
        #always add one since must do a time out if pt was treated that date
        numTimeoutsNeeded+=1
        
        #assume it is missing a timeout
        missingTimeoutDate=True
        
        #loop through all timeouts and check if this date exists in the timeout document
        #will be empty if no timeout document was found so all dates will be missingtimeoutdate=true
        for t,timeOutDate in enumerate(ptTimeOutDates):
            thisTimeOutDate=timeOutDate.strftime('%m/%d/%Y')
            if thisDate == thisTimeOutDate:
                missingTimeoutDate=False
                delIndex=t
                numTimeoutsPerformed+=1
                #now break out so we don't double count ie newstarts and treats should have 2 this would give more
                break
        #did not find matching timeoutdate to thisdate do add it to missing list
        if missingTimeoutDate:
            datesString.append(thisDate)
            missingYesNo="Yes"
            approvedYesNo="No"
        #delet this element from ptTimeOutDates so it doesn't get used again
        #ie if the newstart and first treatment are on the same day there should
        #be two unique dates in the ptTimeOutDates array
        try:
            ptTimeOutDates.pop(t)
        except:
            #if error here that means we ran out of timeout dates before appointment dates
            continue
       
        
        #append this date to check total number of timeouts
        #ptTimeouts.append((ptID,thisDate,missingTimeoutDate))
        #k+=1
    
    thisPatient['NumTimeoutsNeeded']=numTimeoutsNeeded
    thisPatient['NumTimeoutsPerformed']=numTimeoutsPerformed
    thisPatient['MissingTimeouts']=missingYesNo
    thisPatient['ActuallyMissingTimeouts']=missingYesNo
    thisPatient['DatesTimeoutsMissed']=", ".join(datesString)
    thisPatient['Approved']=approvedYesNo
    # add the items to the patient
    if update:
        sp_list.UpdateListItems(data=[thisPatient], kind='Update')
    else:
        sp_list.UpdateListItems(data=[thisPatient], kind='New')
    # AppList.append({'Title':str(k)+"-"+str(strtdate)+"-"+str(enddate),
    #                     'PatientId': str(ptID),
    #                     'NumTimeoutsNeeded': 11,
    #                     'NumTimeoutsPerformed': 10,
    #                     'MissingTimeOutDate': missingYesNo,
    #                     'ActuallyMissing': missingYesNo,
    #                     'DatesTimeoutsMissed': datesString})

#I have gone through all the patients in the list in the current dates
# add these to the sharepoint list
#add in batched of 100? maybe
# batchSize=100
# firstRange=0
# if AppList:
#     #commented lines are for batching if needed
#     # ntimes=len(AppList)/batchSize
#     # nloops=int(ntimes)
#     # if(ntimes-nloops>0):
#     #     firstRange=nloops*batchSize
#     #     lastRange=len(AppList)-1
#     authcookie = Office365('sharepointsite', username=userInfo['username'], password=userInfo['password']).GetCookies()
#     site = Site('sharepointlistsite', authcookie=authcookie)
#     listName='TxTimeOutDates'
#     sp_list = site.List(listName)
#     data = sp_list.GetListItems('All Items')
#     sp_list.UpdateListItems(data=AppList, kind='New')
#     #at this point I have the timeout dates and the list to add them to in Sharepoint
#     # for i in range(nloops):
#     #     firstIndex=i*batchSize
#     #     lastIndex=firstIndex+batchSize-1
#     #     sp_list.UpdateListItems(data=AppList[firstIndex:lastIndex], kind='New')
#     # if firstRange:
#     #     sp_list.UpdateListItems(data=AppList[firstRange:lastRange], kind='New')


word.Quit()

    
