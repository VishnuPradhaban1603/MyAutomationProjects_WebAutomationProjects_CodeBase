import os
import win32com.client
import json
import pyad
import pyad.adquery

from win32com.client import *

 

def EmailtoUser(userEmail,ADGroup):
   

    valid=False

    ADID=""

    q = pyad.adquery.ADQuery()

    q.execute_query(

    attributes = ["sAMAccountName","distinguishedName","userPrincipalName"],

    where_clause = "userPrincipalName='{}'".format(userEmail),

    base_dn = "DC=VFCORP, DC=VFC, DC=com"
   

    )

    for row in q.get_results():

        ADID=(row['sAMAccountName'])

        print(ADID)

    if ADID !="":

        AD_data = {ADGroup}

        finalList = []

        domain = os.getenv('userdomain')

        userPath = GetObject(Pathname = 'WinNT://%s/%s,user' % (domain,ADID))

        for x in userPath.Groups():

            if str(x.Name) in AD_data:

                finalList.append(x.Name)

        if finalList !=[]:
           

            valid=True

    else:

        valid=False
        

        

    return valid
