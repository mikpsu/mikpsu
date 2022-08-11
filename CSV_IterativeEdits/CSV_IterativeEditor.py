# bring in libraries
import re, csv

# file path of CSV opened as file object
file_Obj = open("/Users/mike/Downloads/contacts.csv")

# reader object
csv_reader = csv.reader(file_Obj)

# create list from reader object
contact_info = list(csv_reader)

person_number = 0
for person in contact_info:
    # skip first line because it is the header
    if person_number == 0:
        person_number += 1
        continue
    Email_Address = person[8]

    # Replace the email address line with proper email
    # regex search
    email_wild = re.compile("\-.+")
    email_search = re.search(email_wild, Email_Address)
    email_trimmed = email_search.group()

    #trim off the hyphen
    email_prefix = email_trimmed.replace("-","")

    #append the prefix with the domain
    email_full = email_prefix + "@adigc.com"

    #replace the email address in the list
    person[8] = email_full
    contact_info[person_number]=person

# export list to csv

with open('Contacts_Updated.csv', 'w') as f:
    # using csv.writer method from CSV package
    write = csv.writer(f)
    write.writerows(contact_info)





# Need find the correct email prefix, remove garbage before it, and append '@adigc.com'
# the email should be found in the same index for every row
    # email is in 9th column of each row
    # loop through each row
        # use regex to find the email prefix in 9th column
            # store the prefix
            # append the prefix with '@adigc.com'
        # clear the entry, fill with the email
        # move to next row --->

