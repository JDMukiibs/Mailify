import gmailapi as gapi
import csv
from string import Template


# Service variables
client_secret = "The name of the .json file you got from going through the initial steps."
api_name = "gmail"
api_version = "v1"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
# There are different scopes depending on the use case and to read more about them check https://developers.google.com/gmail/api/auth/scopes


# Build service
service = gapi.create_service(client_secret, api_name, api_version, SCOPES)

# Email variables. Modify this! These variables in the usual case will remain the same since you're sending a similar
# email with same content and attachment to multiple people.
EMAIL_FROM = 'Your Email'
EMAIL_SUBJECT = 'Subject of the Email'
EMAIL_ATTACHMENT = "The file you want to send or the path to it"

# Get the template
with open("message.txt", "r") as content:
    t = Template(content.read())


# Read contact information from contacts.csv file or any file where you have the contacts stored. Then, of course, you wouldn't need to use csv library
# And create the email content for a specific contact and then send the email message
with open('contacts.csv', newline='') as contacts:
    reader = csv.DictReader(contacts)
    for row in reader:
        EMAIL_TO = row['email']
        EMAIL_CONTENT = t.substitute(PERSON_NAME=row['name'])

        # Call the API
        emailMessage = gapi.create_message(EMAIL_FROM, EMAIL_TO, EMAIL_SUBJECT, EMAIL_CONTENT, EMAIL_ATTACHMENT)
        sentMessage = gapi.send_message(service, 'me', emailMessage)
        print("To: {}\nSent: {}\n".format(row['email'], sentMessage))  # Can be removed. But it's use is to track that all contacts were looped through and which labels the message got
