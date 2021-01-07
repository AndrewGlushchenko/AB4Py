#! /usr/bin/env python


# (C)DGV Group.2019
# AB4Py module v0.1
# Working with MacOS/OS X Address Book
# Version MacOS 10.12 and above
# Version python is 2.7 or 3.6 or above

# Needed the Python to Objective-C bridge library
# Actual version PyObjC 5.1.2 was released on 2018-12-13
# https://pyobjc.readthedocs.io/en/latest/index.html#



from AddressBook import *


# Persons record properties key:

# kABFirstNameProperty
# kABMiddleNameProperty
# kABLastNameProperty
#
# kABEmailProperty
# kABPhoneProperty
# kABAddressProperty
#
# kABJobTitleProperty
# kABOrganizationProperty
# kABDepartmentProperty
#
# kABSocialProfileProperty
# kABInstantMessageProperty
#
# kABBirthdayProperty
#
# kABNoteProperty



ad_label_iphone = "iPhone"
ad_label_work = "_$!<Work>!$_"
ad_label_home = "_$!<Home>!$_"
ad_label_workfax = "_$!<WorkFAX>!$_"
ad_label_mobile = "_$!<Mobile>!$_"

adr_tmpl = {'City': '', 'Country': '', 'CountryCode': 'RU', 'Street': '', 'ZIP': ''}

# Check python version 
# Return True if you current version of python meets the requirements of this  module

def CheckVersion() :
    current_v = sys.version_info
    if current_v.major == 2 :
        if current_v.minor > 6 :
            reload(sys)
            sys.setdefaultencoding('utf8')    		
            return True
        else:
            return False
    if current_v.major == 3 :
        if current_v.minor > 4 :
            return True
        else:
            return False
    return False

# Searching Person
# return array of searching person

def SearchPersonByName(Addr_book, Name):
    res_persons = []
    persons = ABCopyArrayOfAllPeople(Addr_book)
    for per1 in persons :
        if Name in GetFullNamePerson(per1) :
            res_persons.append(per1)
    return res_persons


# Return string with full name of the person
# Parameter: reference to person record (ABPerson)
# Return: string with persons full name

def GetFullNamePerson(RefPerson):
    first = RefPerson.valueForProperty_(kABFirstNameProperty)
    if not first:
        first = ''
    middle = RefPerson.valueForProperty_(kABMiddleNameProperty)
    if not middle:
        middle = ''
    last = RefPerson.valueForProperty_(kABLastNameProperty)
    if not last:
        last = ''
    return last + ' ' + first + ' ' + middle


# Return array with fjob description of the person
# Parameter: reference to person record (ABPerson)
# Return: list with name of duty, departments name and organisations name

def GetJobDataPerson(RefPerson):
    idx = [kABJobTitleProperty, kABDepartmentProperty, kABOrganizationProperty]
    dd = []
    for i in range(3):
        nn = RefPerson.valueForProperty_(idx[i])
        if not nn:
            nn = ''
        dd.append(nn)
    return dd


# Return array of phones of person
# Parameter: reference to person record (ABPerson)
# Return: list of phones and labels of it

# the handler ABMultiValueSetPrimaryIdentifier(phones) will be added in the future

def GetPhonesPerson(RefPerson):
    phones = ABRecordCopyValue(RefPerson, kABPhoneProperty)
    ph_array =[]
    if phones != None :
        ph = []
        for i in range(phones.count()):
            ph=[]
            ss = ABMultiValueCopyLabelAtIndex(phones, i)
            ph.append(ss[4:-4])
            ph.append(ABMultiValueCopyValueAtIndex(phones, i))
            ph_array.append(ph)
    return ph_array


# Return array of E-mails of person
# Parameter: reference to person record (ABPerson)
# Return: list of e-mails and labels of it

# the handler ABMultiValueSetPrimaryIdentifier(phones) will be added in the future

def GetEmailsPerson(RefPerson):
    emails = ABRecordCopyValue(RefPerson, kABEmailProperty)
    ph_array =[]
    if emails != None :
        ph = []
        for i in range(emails.count()):
            ph=[]
            ss = ABMultiValueCopyLabelAtIndex(emails, i)
            ph.append(ss[4:-4])
            ph.append(ABMultiValueCopyValueAtIndex(emails, i))
            ph_array.append(ph)
    return ph_array


# Return array of addresses of person
# Parameter: reference to person record (ABPerson)
# Return: list of addresses and labels of it

# the handler ABMultiValueSetPrimaryIdentifier(phones) will be added in the future

def GetAddressesPerson(RefPerson):
    adds = ABRecordCopyValue(RefPerson, kABAddressProperty)
    ph_array = []
    for i in range(adds.count()):
        ph = []
        ss = ABMultiValueCopyLabelAtIndex(adds, i)
        ph.append(ss[4:-4])
        adr = (ABMultiValueCopyValueAtIndex(adds, i))
        ph_array.append(str(adr)[1:-2])
    return ph_array





#  Create new person record (ABPerson) and return reference on it
#  Parameter: last name, first name and middle name

def SetNewPersonRecord(LastName, FirstName='', MiddleName=''):
    new_person = ABPersonCreate()

    res = ABRecordSetValue(new_person, kABLastNameProperty, LastName)
    if res == False:
        return None
    res = ABRecordSetValue(new_person, kABFirstNameProperty, FirstName)
    if res == False:
        return None
    res = ABRecordSetValue(new_person, kABMiddleNameProperty, MiddleName)
    if res == False:
        return None
    return new_person


# Set Organisation name, JobTitle and Department to the person record
# Parameter: reference to person record (ABPerson), names of organisation, position and department
# Return: execution result (True if successfully :)

def SetJobName(RefPerson, Organization, JobTitle ='', Department=''):    
    res = ABRecordSetValue(RefPerson, kABOrganizationProperty, Organization)
    if res == False:
        return None
    if len(JobTitle) > 0 :
        res = ABRecordSetValue(RefPerson, kABJobTitleProperty, JobTitle)
        if res == False:
            return None
    if len(Department) > 0 :        
        res = ABRecordSetValue(RefPerson, kABDepartmentProperty, Department)
        if res == False:
            return None
    return True


# Forming new_person record for writing in the Address Book
# Parameter: list of persons data and name of organisation
# Return: reference to to person record (ABPerson) if successfully else None

def FormingPersonalRecord(record_ar, Organization = ''):
    new_person = SetNewPersonRecord(record_ar[2].strip(), record_ar[3].strip(), record_ar[4].strip())
    if new_person == None :
        return None
    res = SetJobName(new_person, Organization, record_ar[9].strip(), record_ar[10].strip())    
    if not res:
        return None
        
    emails = ABMultiValueCreateMutable()
    phones = ABMultiValueCreateMutable()
    
    ln = record_ar[7].strip()
    ph_ar = ln.split(" ")
    for ph in ph_ar:
        if not ABMultiValueInsert(phones, rus_phone_numb_norm(ph), ad_label_mobile, 0, None)[0]:
            return None
    
    ln = record_ar[8].strip()
    em_ar = ln.split(" ")        
    for em in em_ar:        
        if not ABMultiValueInsert(emails, record_ar[8].strip(), ad_label_work, 0, None)[0]:
            return None
    
    if not ABRecordSetValue(new_person,kABEmailProperty,emails):
        return None    
    if not ABRecordSetValue(new_person,kABPhoneProperty,phones):    
        return None
    
    dt_birth = datetime.strptime( record_ar[11], '%d.%m.%Y') 
    if not ABRecordSetValue(new_person,kABBirthdayProperty, dt_birth):    
        return None
    str_note = "Extension- {0}, Room - {1}".format(record_ar[5], record_ar[6])
    if not ABRecordSetValue(new_person, kABNoteProperty, str_note):
        return None
        
    return new_person

# Print to console data of person
# Parameter: reference to person record (ABPerson)
# Return: None

def print_person_date(RefPerson):
    print(GetFullNamePerson(RefPerson))
    stg_ar = [] 
    stg_ar = GetJobDataPerson(RefPerson)
    print("{:<40}\n{:<40}\n{:<40}\n".format(stg_ar[2], stg_ar[0], stg_ar[1]))
     
    phones = []
    phones  = GetPhonesPerson(RefPerson)
    for ln in range(len(phones)) :
        print(" {:<9} phone - {:<18}".format(phones[ln][0], phones[ln][1]))
    emails = []
    emails  = GetEmailsPerson(RefPerson)
    for ln in range(len(emails)) :
        print(" {:<9} e-mail - {:<18}".format(emails[ln][0], emails[ln][1])) 
    print("")


# Phone russian number normalisation
# from 8-999-555-55-55 as parameter to +7(999)555-55-55 (return)

def rus_phone_numb_norm(ph_numb):
    extn = []
    r0 = re.sub(r'\D','',ph_numb)
    if len(r0) > 11 :
        return '+7({0}){1}-{2}-{3};{4}'.format(r0[1:4],r0[4:7],r0[7:9],r0[9:11],r0[11:])
    else:
        return '+7({0}){1}-{2}-{3}'.format(r0[1:4],r0[4:7],r0[7:9],r0[9:11])
        
        
#Header for .csv file

header_file = ["#", "FullName", "LastName", "MiddleName", "LastName", "Extn",
               "Room", "ph_numb", "e-mail", "JobTitle", "Departmet", "Birthdate"]

import sys
import importlib
import re
import csv
from datetime import datetime


def main():
    if not (CheckVersion()) :
        print("Sorry, your python version is invalid.")
        return 0
    lenst = len(sys.argv)
    if lenst < 2 :
        print("Usage : AB4Py.py -[key]")
        print("        AB4Py.py -h for help")
        return 0
    else :
        if "-h" in sys.argv[1] :
            print("Usage : AB4Py.py -[key]")
            print("        AB4Py.py -m information about login person")
            print("        AB4Py.py -f 'Name' searching person by name")
            print("        AB4Py.py -i 'csv.file' 'Organization' - import records to Address book from .csv file")
            print("                    and setting organsations name for them")
            return 0
        elif "-m" in sys.argv[1] :
            book = ABGetSharedAddressBook()
            if book == None :
                print("Access to Adderess Book denid. Please check the rights of this application.")
                return 0
            me = ABGetMe(book)
            print_person_date(me)
            return 0
        elif "-f" in sys.argv[1] :
            if len(sys.argv) < 3:
                print("Name for the searching is not signed.")
                return 3
            else:
                s_name = sys.argv[2]
                book = ABGetSharedAddressBook()
                if book == None :                
                    print("Access to Adderess Book denid. Please check the rights of this application.")
                    return 0
                persons = SearchPersonByName(Abook, s_ame,)
                for per in persons :
                    print_person_date(per)
                return 0
        elif "-i" in sys.argv[1] :
            if len(sys.argv) < 3:
                print("File .csv is not signed.")
                return 3
            else:
                csv_path = sys.argv[2]
                if len(sys.argv) > 3:
                    Organiz = sys.argv[3]
                else:
                    Organiz = ""
                line_ar = []
                with open(csv_path, "r") as f_obj:
                    reader = csv.reader(f_obj)
                    for row in reader:
                       line_ar.append(row)
                del(line_ar[0])
                del(line_ar[0]) 
                i = 0
                book = ABGetSharedAddressBook()
                if book == None :                
                    print("Access to Adderess Book denid. Please check the rights of this application.")
                    return 0
                for row in line_ar :
                    record_ar = row[0].split(";")
                    new_person = FormingPersonalRecord(record_ar, Organiz)
                    if new_person == None:
                        continue
                    print("Add: "+ row[0])
                    ABAddRecord(book, new_person)
                    i += 1
                else:
                    print("Added " + str(i)+" records.")
                    if i > 0:
                        ABSave(book)
                return 0
        else :
            print("Wrong key!")
            print("AB4Py.py -h for help")
            return 1
            
if __name__ == "__main__":
    main()