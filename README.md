

# AB4Py python module for working with MacOS/OS X Address Book

### Needed the Python to Objective-C bridge library PyObjC. 
### Actual version 5.1.2 was released on 2018-12-13 
### https://pyobjc.readthedocs.io/en/latest/index.html


Installing or upgrading PyObjC using pip is easy:

    $ pip install -U pyobjc

=======

This module was verified with python 2.7 and 3.6.  In the March 3, 2019 issue was added an adaptation for python 2.7, a version check of python, and a check of the accessing the MacOs address book result.

Special thanks to Vladimir Litovko and  Eugene Prusakov for testing, tips and support.

### Using function from AddressBook module:

Getting reference on Address Book

    book = ABGetSharedAddressBook()
    book = AddressBook.ABAddressBook.sharedAddressBook()

Getting reference on Address Book

    book = ABGetSharedAddressBook()
    book = AddressBook.ABAddressBook.sharedAddressBook()

Getting reference on Login Person

    me = ABGetMe(book)

Creating of new person record

    new_person = ABPersonCreate()

Get properties of persons record in accordance with key

    RefPerson.valueForProperty_(kABMiddleNameProperty)

Create MultiValue record

    emails = ABMultiValueCreateMutable()
    phones = ABMultiValueCreateMutable()
    addresses = ABMultiValueCreateMutable()

Template for address record

    adr_tmpl['City'] = "Москва"
    adr_tmpl['Country'] = 'Россия'
    adr_tmpl['Street'] = "Тупик нечистой силы, 13"
    adr_tmpl['ZIP'] = "123123"

Insert element of MultiValue record

    ABMultiValueInsert(emails,"glav@psih.ru", ad_label_work, 0, None)
    ABMultiValueInsert(phones,"+7(985)999-99-99", ad_label_work, 0, None)
    ABMultiValueInsert(new_address,adr_tmpl, ad_label_work, 0, None)

Add MultiValue record to the Personal record

    ABRecordSetValue(new_person,kABAddressProperty,new_address)
    ABRecordSetValue(new_person,kABEmailProperty,emails)
    ABRecordSetValue(new_person,kABPhoneProperty,phones)

Add new_person record to address book 

    ABAddRecord(book, new_person)

Save address book  

    ABSave(book)


### Function in this module:

    SearchPersonByName(Addr_book, Name)
    GetFullNamePerson(RefPerson)
    GetJobDataPerson(RefPerson)
    GetPhonesPerson(RefPerson)
    GetEmailsPerson(RefPerson)
    GetAddressesPerson(RefPerson)

    SetNewPersonRecord(LastName, FirstName='', MiddleName='')
    SetJobName(RefPerson, Organization, JobTitle ='', Department='')
    FormingPersonalRecord(record_ar, Organization = '')
 
    print_person_date(RefPerson)
    rus_phone_numb_norm(ph_numb)
==========

### You can use it as separete shell application

    Usage : AB4Py.py -[key]
        -m information about login person
        -f 'Name' searching person by name
        -i 'csv.file' 'Organization' - import records to Address book from .csv file
        and setting organsations name for them
        
        Format of csv.file describs in the list header_file
