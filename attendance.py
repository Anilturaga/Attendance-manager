#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime
import numpy as np
data = pd.read_excel("/home/anil/Desktop/final.xlsx")
data = data.drop(columns=['CompanyName', 'DeptName','LocationName','MachineId','MachineNo','PunchType','Latitude','Longitude'])
def student(year):    
    xls = pd.ExcelFile('/home/anil/Desktop/Student List 2018-19( final).xlsx')
    df1 = pd.read_excel(xls, 'CSE')
    df2 = pd.read_excel(xls, 'EEE')
    df3 = pd.read_excel(xls, 'ME')
    df4 = pd.read_excel(xls, 'CIVIL')
    f = pd.concat([df1,df2,df3,df4],sort=True)
    flist = list(f['HALL TICKET NO.'])
    fname = list(f['NAME OF THE STUDENT '])
    year = year*1000
    for i in range(len(flist)):
        flist[i] = int(flist[i][-3:])
        flist[i] = flist[i] + year
    return flist,fname

"""
8:20-9:20 => 500-560
9:25-10:20 => 565-620
10:35-11:30 => 635-690
11:35-12:30 => 695-750
12:35-1:30 => 755-810
"""
def Diff(li1, li2): 
    li_dif = [i for i in li1 + li2 if i not in li1 or i not in li2] 
    return li_dif 
def PH202():
    df1 = pd.DataFrame(columns=['CardNo','Name','PunchDate','PunchTime'])
    final = pd.DataFrame(columns=['CardNo','Name','count'])
    count = 0
    li = []
    for i in data['CardNo']:
        if int(i/1000) == 18:
            li.append(data.iloc[count,:])
        count = count + 1
    df1 = df1.append(li)
    li1 = []
    df1.PunchDate.unique()
    print(len(df1.CardNo.unique()))
    for i in df1.CardNo.unique():
        li1.append(df1.loc[df1['CardNo'] == i])
    for k in range(0,len(li1)):
        li2 = [] 
        for j in li1[k]['PunchTime']:
            j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 500 and j.hour*60+j.minute < 560):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Monday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 500 and j.hour*60+j.minute < 560):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Wednesday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 635 and j.hour*60+j.minute < 690):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Friday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
        x = 0
        if len(li2)>0:
            df3 = pd.concat(li2)
            x = len(df3.PunchDate.unique())
        idc = pd.DataFrame({"CardNo":li1[k].CardNo.unique()[0],"Name":li1[k].Name.unique()[0],"count":x}, index=[0])
        final = pd.concat([final,idc])
    finalist = list(final.CardNo[0])
    addlist = Diff(flist,finalist)
    for ans in addlist:
        tempo = {'CardNo': ans, 'Name': fname[flist.index(ans)], 'count': 0}
        final = final.append(tempo, ignore_index=True)
    final.sort_values(["CardNo", "Name"], axis=0,ascending=True, inplace=True)
    final.to_excel("/home/anil/Desktop/attendance/PH202.xlsx")
def ES210():
    df1 = pd.DataFrame(columns=['CardNo','Name','PunchDate','PunchTime'])
    final = pd.DataFrame(columns=['CardNo','Name','count'])
    count = 0
    li = []
    for i in data['CardNo']:
        if int(i/1000) == 18:
            li.append(data.iloc[count,:])
        count = count + 1
    df1 = df1.append(li)
    li1 = []
    df1.PunchDate.unique()
    print(len(df1.CardNo.unique()))
    for i in df1.CardNo.unique():
        li1.append(df1.loc[df1['CardNo'] == i])
    for k in range(0,len(li1)):
        li2 = [] 
        for j in li1[k]['PunchTime']:
            j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 500 and j.hour*60+j.minute < 560):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Tuesday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 565 and j.hour*60+j.minute < 620):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Thursday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
        x = 0
        if len(li2)>0:
            df3 = pd.concat(li2)
            x = len(df3.PunchDate.unique())
        idc = pd.DataFrame({"CardNo":li1[k].CardNo.unique()[0],"Name":li1[k].Name.unique()[0],"count":x}, index=[0])
        final = pd.concat([final,idc])
    finalist = list(final.CardNo[0])
    addlist = Diff(flist,finalist)
    for ans in addlist:
        tempo = {'CardNo': ans, 'Name': fname[flist.index(ans)], 'count': 0}
        final = final.append(tempo, ignore_index=True)
    final.sort_values(["CardNo", "Name"], axis=0,ascending=True, inplace=True)
    final.to_excel("/home/anil/Desktop/attendance/ES210.xlsx")
def ES208():
    df1 = pd.DataFrame(columns=['CardNo','Name','PunchDate','PunchTime'])
    final = pd.DataFrame(columns=['CardNo','Name','count'])
    count = 0
    li = []
    for i in data['CardNo']:
        if int(i/1000) == 18:
            li.append(data.iloc[count,:])
        count = count + 1
    df1 = df1.append(li)
    li1 = []
    df1.PunchDate.unique()
    print(len(df1.CardNo.unique()))
    for i in df1.CardNo.unique():
        li1.append(df1.loc[df1['CardNo'] == i])
    for k in range(0,len(li1)):
        li2 = [] 
        for j in li1[k]['PunchTime']:
            j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 685 and j.hour*60+j.minute < 750):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Thursday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 565 and j.hour*60+j.minute < 620):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Tuesday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
        x = 0
        if len(li2)>0:
            df3 = pd.concat(li2)
            x = len(df3.PunchDate.unique())
        idc = pd.DataFrame({"CardNo":li1[k].CardNo.unique()[0],"Name":li1[k].Name.unique()[0],"count":x}, index=[0])
        final = pd.concat([final,idc])
    finalist = list(final.CardNo[0])
    addlist = Diff(flist,finalist)
    for ans in addlist:
        tempo = {'CardNo': ans, 'Name': fname[flist.index(ans)], 'count': 0}
        final = final.append(tempo, ignore_index=True)
    final.sort_values(["CardNo", "Name"], axis=0,ascending=True, inplace=True)
    final.to_excel("/home/anil/Desktop/attendance/ES208.xlsx")
def MA203():
    df1 = pd.DataFrame(columns=['CardNo','Name','PunchDate','PunchTime'])
    final = pd.DataFrame(columns=['CardNo','Name','count'])
    count = 0
    li = []
    for i in data['CardNo']:
        if int(i/1000) == 18:
            li.append(data.iloc[count,:])
        count = count + 1
    df1 = df1.append(li)
    li1 = []
    df1.PunchDate.unique()
    print(len(df1.CardNo.unique()))
    for i in df1.CardNo.unique():
        li1.append(df1.loc[df1['CardNo'] == i])
    for k in range(0,len(li1)):
        li2 = [] 
        for j in li1[k]['PunchTime']:
            j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 500 and j.hour*60+j.minute < 560):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Thursday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 565 and j.hour*60+j.minute < 620):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Wednesday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
        x = 0
        if len(li2)>0:
            df3 = pd.concat(li2)
            x = len(df3.PunchDate.unique())
        idc = pd.DataFrame({"CardNo":li1[k].CardNo.unique()[0],"Name":li1[k].Name.unique()[0],"count":x}, index=[0])
        final = pd.concat([final,idc])
    finalist = list(final.CardNo[0])
    addlist = Diff(flist,finalist)
    for ans in addlist:
        tempo = {'CardNo': ans, 'Name': fname[flist.index(ans)], 'count': 0}
        final = final.append(tempo, ignore_index=True)
    final.sort_values(["CardNo", "Name"], axis=0,ascending=True, inplace=True)
    final.to_excel("/home/anil/Desktop/attendance/MA203.xlsx")
def ES209():
    df1 = pd.DataFrame(columns=['CardNo','Name','PunchDate','PunchTime'])
    final = pd.DataFrame(columns=['CardNo','Name','count'])
    count = 0
    li = []
    for i in data['CardNo']:
        if int(i/1000) == 18:
            li.append(data.iloc[count,:])
        count = count + 1
    df1 = df1.append(li)
    li1 = []
    df1.PunchDate.unique()
    print(len(df1.CardNo.unique()))
    for i in df1.CardNo.unique():
        li1.append(df1.loc[df1['CardNo'] == i])
    for k in range(0,len(li1)):
        li2 = [] 
        for j in li1[k]['PunchTime']:
            j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 565 and j.hour*60+j.minute < 620):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Friday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 635 and j.hour*60+j.minute < 690):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Wednesday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
            if(j.hour*60+j.minute > 755 and j.hour*60+j.minute < 810):
                j = j.strftime("%H:%M:%S")
                mytemp = li1[k].loc[li1[k]['PunchTime'] == j]
                j = datetime.strptime(j, '%H:%M:%S').time()
                if pd.to_datetime(mytemp.PunchDate).dt.day_name().all() == 'Thursday':
                    j = j.strftime("%H:%M:%S")
                    li2.append(li1[k].loc[li1[k]['PunchTime'] == j])
                    j = datetime.strptime(j, '%H:%M:%S').time()
        x = 0
        if len(li2)>0:
            df3 = pd.concat(li2)
            x = len(df3.PunchDate.unique())
        idc = pd.DataFrame({"CardNo":li1[k].CardNo.unique()[0],"Name":li1[k].Name.unique()[0],"count":x}, index=[0])
        final = pd.concat([final,idc])
    finalist = list(final.CardNo[0])
    addlist = Diff(flist,finalist)
    for ans in addlist:
        tempo = {'CardNo': ans, 'Name': fname[flist.index(ans)], 'count': 0}
        final = final.append(tempo, ignore_index=True)
    final.sort_values(["CardNo", "Name"], axis=0,ascending=True, inplace=True)
    final.to_excel("/home/anil/Desktop/attendance/ES209.xlsx")


flist,fname = student(18)
PH202()
ES210()
ES208()
MA203()
ES209()