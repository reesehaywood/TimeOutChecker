# -*- coding: utf-8 -*-
"""
Created on Thu Sep 26 14:24:04 2019

@author: haywoojr
"""

import pyodbc
#import datetime
import pandas as pd
from pandas.tseries.offsets import BDay
import datetime
from pathlib import Path
import win32com.client
import shutil,os,time

#directory to store files locally
tempDir=Path('C:/Temp3/finals/')

# Create connection
con = pyodbc.connect(driver="{SQL Server}",server='mshsaptvard1',database='variansystem',uid='reports',pwd='reports',autocommit=True)
cur = con.cursor()
#get patients that have had a final
#last date run 10-31-19
n=datetime.datetime(2021,3,1)
strtdate=n.strftime("%m/%d/%y")
n=datetime.datetime(2021,4,1)
enddate=n.strftime("%m/%d/%y")
getPatientList = """
select distinct Patient.PatientId,Patient.LastName,Patient.FirstName,NonScheduledActivity.WorkFlowActiveFlag,Activity.ActivityCode
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
""".format(strtdate=strtdate,enddate=enddate)

#
PtTab=cur.execute(getPatientList)
#Put ptTab into python object not pyodbc opbject
ptsListIDs=[]
ptsListName=[]
ptFinalDocs=[]
for row in PtTab:
    #print(row[0],row[4])
    ptsListIDs.append(row[0])
    ptsListName.append((row[1],row[2]))
print("len of patients= ",len(ptsListIDs))
#print(ptsListIDs)
#loop through patient to get their treatment timeout document
for ptID in ptsListIDs:
    #get patient appointments/ currently ignores Simulation
    getPtApp="""
select distinct Patient.PatientId,Patient.LastName,Patient.FirstName,ScheduledActivity.ScheduledStartTime,ScheduledActivity.ActualStartDate,ScheduledActivity.ScheduledActivityCode,Hospital.HospitalName,Department.DepartmentName,Activity.ActivityCode,Activity.ObjectStatus,Patient.PatientSer,Machine.MachineId,vv_ActivityLng.Expression1
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
((ScheduledActivity.ObjectStatus='Active') and (ScheduledActivity.ActualStartDate between '{strtdate}' and dateadd(dd,1,'{enddate}')) 
and (ScheduledActivity.ScheduledActivityCode like '%Complete%')
--removed comment for counting new starts 
--and (vv_ActivityLng.Expression1 not like 'New%')
and (vv_ActivityLng.Expression1 not like '%QA%') and (vv_ActivityLng.Expression1 not like '%Physician%')
and (vv_ActivityLng.Expression1 not like '%IMRT%')  
and (vv_ActivityLng.Expression1 not like '%Simulation%')    
    """.format(patid=ptID,strtdate=strtdate,enddate=enddate)
    ptAppTab=cur.execute(getPtApp)
    ptAppList=[]
    for ptApp in ptAppTab:
        ptAppList.append(ptApp)
        
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
    try:
        i=0
        for rows in ptDocTab:
            i+=1
            #print(i,ptID,rows[5],rows[6],rows[7])
            temp=(ptID,Path(rows[5]+"/"+rows[6]+"/"+rows[7]),rows[7],rows[3])
            ptFinalDocs.append(temp)
    except:
        temp=(ptID,Path(None))
        ptFinalDocs.append(temp)
print("len of final doc= ",len(ptFinalDocs))
#loop through final documents to get number of timeouts performed
for i in range(len(ptFinalDocs)):
    #print(ids,str(fl))
    ids=ptFinalDocs[i][0]
    fl=ptFinalDocs[i][1]
    fname=ptFinalDocs[i][2]
    shutil.copyfile(str(fl),str(tempDir)+"\\"+ids+"-"+fname)

ptTimeouts=[]
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
ptTimeOutDates=[]
for i in range(len(ptFinalDocs)):
    ids=ptFinalDocs[i][0]
    fl=ptFinalDocs[i][1]
    fname=ptFinalDocs[i][2]
    wb=word.Documents.Open(str(tempDir)+"\\"+ids+"-"+fname)

    doc=word.ActiveDocument
    numTables=doc.Tables.Count
    numTimeOuts=0
    #TODO Need to make this table in tabled
    for k in range(numTables):
        table=doc.Tables(k+1)
        cont=table.Cell(Row=1,Column=2).Range.Text
        ptTimeOutDates.append(cont)
        #print(cont)
        if " ".join(cont.split())!='\x07':
            numTimeOuts+=1
            
    word.Documents.Close()
    temp=(ids,numTimeOuts)
    ptTimeouts.append(temp)
    #os.remove(str(tempDir)+"\\"+fname)
for i in range(len(ptFinalDocs)):
    ids=ptFinalDocs[i][0]
    fl=ptFinalDocs[i][1]
    fname=ptFinalDocs[i][2]    
    #print(ids,str(tempDir)+"\\"+ids+"-"+fname)
    os.remove(str(tempDir)+"\\"+ids+"-"+fname)

word.Quit()
print(ptTimeouts)
ptTreats=[]
for ptID in ptsListIDs:
    getPatientScheduledAppoint="""--Count of Scheduled Appointments
select distinct Patient.PatientId,Patient.LastName,Patient.FirstName,convert(varchar,ScheduledActivity.ScheduledStartTime,110)
as NumSchedTrtsleft
from Patient,ScheduledActivity,ActivityInstance,Activity
where (Patient.PatientId like '{patid}')
and (Patient.PatientSer = ScheduledActivity.PatientSer)
--and (ScheduledActivity.ScheduledStartTime between @startdate and @enddate)
and (ActivityInstance.ActivityInstanceSer=ScheduledActivity.ActivityInstanceSer)
and (ActivityInstance.ActivityInstanceRevCount = ScheduledActivity.ActivityInstanceRevCount)
and (ActivityInstance.ActivitySer=Activity.ActivitySer)
and ((Activity.ActivityCode like '%Treatment%') or (Activity.ActivityCode like '%Tx%') or (Activity.ActivityCode like 'RMS-2000%')
or (Activity.ActivityCode like '%Stereo%') or (Activity.ActivityCode like '%Final%'))
--and (Activity.ActivityCode not like '%Final%')
and (ScheduledActivity.ScheduledActivityCode like '%Completed%')
--group by Patient.PatientId,Patient.LastName,Patient.FirstName""".format(patid=ptID)
    ptTreatsTab=cur.execute(getPatientScheduledAppoint)
    cnt=len(list(cur.fetchall()))
    #for row in ptTreatsTab:
    #    print(row)
    temp=(ptID,cnt)
    ptTreats.append(temp)
mtch=[]
nomtch=[]
i=0
for ids in ptsListIDs:
    for j in range(len(ptTreats)):
        if ptTreats[j][0]==ids:
            tempTx=ptTreats[j][1]
    for j in range(len(ptTimeouts)):
        if ptTimeouts[j][0]==ids:
            tempTimes=ptTimeouts[j][1]
    
    if tempTx==tempTimes:
        mtch.append(1.0)
    else:
        nomtch.append(ids)
        mtch.append(0.0)
    print(ids,ptsListName[i][0],ptsListName[i][1],tempTx,tempTimes)
    i+=1       
    
