#! /usr/bin/env python
# (C)Andrey Glushchenko. AG IT Consulting LLC. 2019..2025
# AB4Py module v1.01
# Working with macOS/OS X Address Book
# Version macOS 10.12 and above
# Version python is  3.7 or above

# Needed the Python to Objective-C bridge library
# Actual version PyObjC 5.1.2 was released on 2018-12-13
# https://pyobjc.readthedocs.io/en/latest/index.html#
from __future__ import annotations
from AddressBook import *
import sys
import re
import csv
from datetime import datetime
from typing import TYPE_CHECKING, Optional, Any

ABPerson: Any

if TYPE_CHECKING:
    ABPerson = Any

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


def CheckVersion() -> bool:
    """
    Check python version.
    :return: bool, return True if you current version of python meets the requirements of this  module
    """
    current_v = sys.version_info
    if current_v.major == 3:
        if current_v.minor > 7:
            return True
        else:
            return False
    return False


def SearchPersonByName(addr_book: object, name: str) -> list:
    """
    Searching Person by string
    :param addr_book: object - reference to AddressBook
    :param name: str - string for the searching
    :return: list of the persons
    """
    res_persons = []
    persons = ABCopyArrayOfAllPeople(addr_book)
    for per1 in persons:
        if name in GetFullNamePerson(per1):
            res_persons.append(per1)
    return res_persons


def GetFullNamePerson(ref_person: ABPerson) -> str:
    """
    Returns string with full name of the person
    :param ref_person: ABPerson - reference to person record
    :return: str - string with persons full name
    """
    first = ref_person.valueForProperty_(kABFirstNameProperty)
    if not first:
        first = ''
    middle = ref_person.valueForProperty_(kABMiddleNameProperty)
    if not middle:
        middle = ''
    last = ref_person.valueForProperty_(kABLastNameProperty)
    if not last:
        last = ''
    return last + ' ' + first + ' ' + middle


def GetJobDataPerson(ref_person: ABPerson) -> list:
    """
    Returns an array of strings with the company name, department name, and job title
    :param ref_person: ABPerson - reference to person record
    :return: list of strings
    """
    idx = [kABJobTitleProperty, kABDepartmentProperty, kABOrganizationProperty]
    dd = []
    for i in range(3):
        nn = ref_person.valueForProperty_(idx[i])
        if not nn:
            nn = ''
        dd.append(nn)
    return dd


# the handler ABMultiValueSetPrimaryIdentifier(phones) will be added in the future

def GetPhonesPerson(ref_person: ABPerson) -> list:
    """
    Returns array of phones of person
    :param ref_person: ABPerson - reference to person record
    :return: list of phones and their labels
    """
    phones = ABRecordCopyValue(ref_person, kABPhoneProperty)
    ph_array = []
    if phones is not None:
        for i in range(phones.count()):
            ph = []
            ss = ABMultiValueCopyLabelAtIndex(phones, i)
            ph.append(ss[4:-4])
            ph.append(ABMultiValueCopyValueAtIndex(phones, i))
            ph_array.append(ph)
    return ph_array


# the handler ABMultiValueSetPrimaryIdentifier(phones) will be added in the future

def GetEmailsPerson(ref_person: ABPerson) -> list:
    """
    Return array of E-mails of person
    :param ref_person: ABPerson - reference to person record
    :return: list of e-mails and their labels
    """
    emails = ABRecordCopyValue(ref_person, kABEmailProperty)
    ph_array = []
    if emails is not None:
        for i in range(emails.count()):
            ph = []
            ss = ABMultiValueCopyLabelAtIndex(emails, i)
            ph.append(ss[4:-4])
            ph.append(ABMultiValueCopyValueAtIndex(emails, i))
            ph_array.append(ph)
    return ph_array


# the handler ABMultiValueSetPrimaryIdentifier(phones) will be added in the future

def GetAddressesPerson(ref_person: ABPerson) -> list:
    """
    Returns array of addresses of person
    :param ref_person: ABPerson - reference to person record
    :return: list of addresses and their labels
    """
    adds = ABRecordCopyValue(ref_person, kABAddressProperty)
    ph_array = []
    for i in range(adds.count()):
        ph = []
        ss = ABMultiValueCopyLabelAtIndex(adds, i)
        ph.append(ss[4:-4])
        adr = (ABMultiValueCopyValueAtIndex(adds, i))
        ph_array.append(str(adr)[1:-2])
    return ph_array


def SetNewPersonRecord(last_name: str, first_name='', middle_name='') -> Optional[ABPerson]:
    """
    Creates new person record
    :param last_name: str - last name
    :param first_name: str - first name, by default ''
    :param middle_name: str - middle name, by default ''
    :return: ABPerson - reference to person record
    """
    new_person = ABPersonCreate()

    res = ABRecordSetValue(new_person, kABLastNameProperty, last_name)
    if not res:
        return None
    res = ABRecordSetValue(new_person, kABFirstNameProperty, first_name)
    if not res:
        return None
    res = ABRecordSetValue(new_person, kABMiddleNameProperty, middle_name)
    if not res:
        return None
    return new_person


def SetJobName(ref_person: ABPerson, organization: str, job_title='', department='') -> bool:
    """
    Sets Organisation name, JobTitle and Department to the person record
    :param ref_person: ABPerson - reference to person record
    :param organization: str - company name
    :param job_title: str - position name
    :param department: str - department name
    :return: execution result (True if successfully)
    """
    res = ABRecordSetValue(ref_person, kABOrganizationProperty, organization)
    if not res:
        return False
    if len(job_title) > 0:
        res = ABRecordSetValue(ref_person, kABJobTitleProperty, job_title)
        if not res:
            return False
    if len(department) > 0:
        res = ABRecordSetValue(ref_person, kABDepartmentProperty, department)
        if not res:
            return False
    return True


def FormingPersonalRecord(record_ar: list, organization='') -> ABPerson:
    """
    Forming new_person record for writing in the Address Book
    :param record_ar: list of persons data and name of organisation
    :param organization: str - company name
    :return: ABPerson - reference to person record or None if the result is unsuccessful
    """
    new_person = SetNewPersonRecord(record_ar[2].strip(), record_ar[3].strip(), record_ar[4].strip())
    if new_person is None:
        return None
    res = SetJobName(new_person, organization, record_ar[9].strip(), record_ar[10].strip())
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
    for _ in em_ar:
        if not ABMultiValueInsert(emails, record_ar[8].strip(), ad_label_work, 0, None)[0]:
            return None

    if not ABRecordSetValue(new_person, kABEmailProperty, emails):
        return None
    if not ABRecordSetValue(new_person, kABPhoneProperty, phones):
        return None

    dt_birth = datetime.strptime(record_ar[11], '%d.%m.%Y')
    if not ABRecordSetValue(new_person, kABBirthdayProperty, dt_birth):
        return None
    str_note = "Extension- {0}, Room - {1}".format(record_ar[5], record_ar[6])
    if not ABRecordSetValue(new_person, kABNoteProperty, str_note):
        return None

    return new_person


def print_person_date(ref_person: ABPerson) -> None:
    """
    Prints to console data of person
    :param ref_person: ABPerson - reference to person
    :return: None
    """
    print('-> {:<40}'.format(GetFullNamePerson(ref_person)))

    stg_ar = GetJobDataPerson(ref_person)
    print('Organization: {:<40}\nJob title: {:<40}\nDepartment: {:<40}'.format(stg_ar[2], stg_ar[0], stg_ar[1]))

    phones = GetPhonesPerson(ref_person)
    for ln in range(len(phones)):
        print(' {:<9} phone - {:<18}'.format(phones[ln][0], phones[ln][1]))

    emails = GetEmailsPerson(ref_person)
    for ln in range(len(emails)):
        print(' {:<9} e-mail - {:<18}'.format(emails[ln][0], emails[ln][1]))
    print("")


def rus_phone_numb_norm(ph_numb: str):
    """
    Phone number normalisation
    :param ph_numb: str source phone number string
    :return: str formatted phone number string
    """

    r0 = re.sub(r'\D', '', ph_numb)
    if len(r0) > 11:
        return '+7({0}){1}-{2}-{3};{4}'.format(r0[1:4], r0[4:7], r0[7:9], r0[9:11], r0[11:])
    else:
        return '+7({0}){1}-{2}-{3}'.format(r0[1:4], r0[4:7], r0[7:9], r0[9:11])


# Header for .csv file

header_file = ["#", "FullName", "LastName", "MiddleName", "LastName", "Extn",
               "Room", "ph_numb", "e-mail", "JobTitle", "Department", "Birthdate"]


def main():
    if not (CheckVersion()):
        print("Sorry, your python version is invalid.")
        return 0
    lenst = len(sys.argv)
    if lenst < 2:
        print("Usage : AB4Py.py -[key]")
        print("        AB4Py.py -h for help")
        return 0
    else:
        if "-h" in sys.argv[1]:
            print("Usage : AB4Py.py -[key]")
            print("        AB4Py.py -m information about login person")
            print("        AB4Py.py -f 'Name' searching person by name")
            print("        AB4Py.py -a List of total contacts of your address book")
            print("        AB4Py.py -i 'csv.file' 'Organization' - import records to address book from .csv file")
            print("                    and setting organizations name for them")
            return 0
        elif "-m" in sys.argv[1]:
            book = ABGetSharedAddressBook()
            if book is None:
                print("Access to the Address Book denied. Please check the rights of this application.")
                return 0
            me = ABGetMe(book)
            print_person_date(me)
            return 0
        elif "-f" in sys.argv[1]:
            if len(sys.argv) < 3:
                print("Name for the searching is not signed.")
                return 3
            else:
                s_name = sys.argv[2]
                book = ABGetSharedAddressBook()
                if book is None:
                    print("Access to the Address Book denied. Please check the rights of this application.")
                    return 0
                persons = SearchPersonByName(book, s_name, )
                for per in persons:
                    print_person_date(per)
                print(f'Total records: {len(persons)}.')
                return 0
        elif "-i" in sys.argv[1]:
            if len(sys.argv) < 3:
                print("File .csv is not signed.")
                return 3
            else:
                csv_path = sys.argv[2]
                if len(sys.argv) > 3:
                    organiz = sys.argv[3]
                else:
                    organiz = ""
                line_ar = []
                with open(csv_path, "r") as f_obj:
                    reader = csv.reader(f_obj)
                    for row in reader:
                        line_ar.append(row)
                del (line_ar[0])
                del (line_ar[0])
                i = 0
                book = ABGetSharedAddressBook()
                if book is None:
                    print("Access to the Address Book is denied. Please check the rights of this application.")
                    return 0
                for row in line_ar:
                    record_ar = row[0].split(";")
                    new_person = FormingPersonalRecord(record_ar, organiz)
                    if new_person is None:
                        continue
                    print("Add: " + row[0])
                    ABAddRecord(book, new_person)
                    i += 1
                else:
                    print("Added " + str(i) + " records.")
                    if i > 0:
                        ABSave(book)
                return 0
        else:
            print("Wrong key!")
            print("AB4Py.py -h for help")
            return 1


if __name__ == "__main__":
    main()
