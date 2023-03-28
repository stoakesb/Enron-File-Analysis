import csv
import email

#path to the eml file
filename = '/Users/brittonstoakes/Documents/Projects/Enron-Dashboards/edrm-enron-v2_kaminski-v_xml_1of2/native_000/3.311777.CD3AZ41RRTREN0KKJ3TKSIQKYR4POEBEA.eml'

#open the eml file
f = open(filename, 'r')

#parse the eml file
msg = email.message_from_file(f)

#grab email body

if msg.is_multipart():
    for part in msg.walk():
        ctype = part.get_content_type()
        cdispo = str(part.get('Content-Disposition'))

        # skip any text/plain (txt) attachments
        if ctype == 'text/plain' and 'attachment' not in cdispo:
            body = part.get_payload(decode=True)  # decode
            break
        
# not multipart - i.e. plain text, no attachments, keeping fingers crossed
else:
    body = msg.get_payload(decode=True)

#open a csv file to write the parsed data
csvfile = open('/Users/brittonstoakes/Documents/Projects/Enron-Dashboards/edrm-enron-v2_kaminski-v_xml_1of2/csv-output/output7.csv', 'w')
writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)

#write the headers
writer.writerow(['Subject', 'From', 'To', 'Body'])

#write the parsed data to the csv file
writer.writerow([msg['Subject'], msg['From'], msg['To'], body])

#close the csv file
csvfile.close()