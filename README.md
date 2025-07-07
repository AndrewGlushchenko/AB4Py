

# AB4Py python module for working with MacOS/OS X Address Book

### Needed the Python to Objective-C bridge library PyObjC. 
### Actual version 11.1 was released on 2025-06-14 
### https://pyobjc.readthedocs.io/en/latest/index.html


Installing or upgrading PyObjC using pip is easy:

    $ pip install -U pyobjc

=======

Current version 1.01 dated July 6, 2025. This module has been tested with Python 3.7 and above. Support for Python 2.7 has been discontinued as obsolete.

I am still grateful to Vladimir Litovko and Eugene Prusakov for their help in testing the first version of this module.

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

    adr_tmpl['Street'] = "435 N Michigan Ave"
    adr_tmpl['City'] = "Chicago"
    adr_tmpl['ZIP'] = "60001"
    adr_tmpl['Country'] = 'USA'

Insert element of MultiValue record

    ABMultiValueInsert(emails,"j_dou@google.com", ad_label_work, 0, None)
    ABMultiValueInsert(phones,"+1(555)555-12-34", ad_label_work, 0, None)
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

The purpose and parameters of the functions are described in the module code.

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
            AB4Py.py -m information about login person
            AB4Py.py -f 'Name' search for person by name
            AB4Py.py -a List of all contacts in your address book
            AB4Py.py -i 'csv.file' 'Organization' - import records to the address book from a .csv file
                        with the company name specified in the “Organization” field
        
Format of csv.file describs in the value "header_file" in the module.
