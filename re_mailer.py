from datetime import datetime
import pyodbc
from exchangelib import Account, Credentials, Configuration, DELEGATE

# This is instantiated outside of the methods to prevent attempted double-connections (which, IIRC, never succeed)
connection = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER=ServerGoesHere;DATABASE=DBNameGoesHere;TrustedConnection=yes;')


class SentItemsError(Exception):
    def __init__(self, message):
        super().__init__(message)


cred = Credentials('USerGoesHere', 'PasswordGoesHere')
config = Configuration(server='ServerGoesHere', credentials=cred)
a = Account(primary_smtp_address='User@ServerGoesHere', autodiscover=False, config=config, access_type=DELEGATE)


def forwarder(email_subject, unique_id):
    a.root.refresh()
    print(email_subject)
    message = a.sent.get(subject=email_subject)
    message.forward(
        subject='Fwd: 10-Minute Reminder',
        body='Hey [name of recipient can go here],\n\nThis is a 10-minute reminder for [whatever the reminders were meant for].',
        to_recipients=['ExampleRecipient@host.com', 'AnExampleOfASecondRecipientOrForOneJustRemoveThis&Comma@host.com']
    )


# Note the code in the last two conditional statements of the following method. They should never happen to begin with,
# but if they do, the program will recover and finish its scheduled run. The odds of the last error occurring are
# slim-to-none, as the original emails' subjects are all unique so long as our incoming data
# (to the database at the very start) is unique. The odds of the first error occurring are a bit higher. There could,
# eventually, be data put into the DB that never got emailed for one reason or another.
def get_full_subject(unique_id):
    qs = a.sent.all()
    partial_subject = f'&{unique_id} '
    items = qs.filter(subject__contains=partial_subject)
    if len(items) == 1:
        return items[0].subject
    elif len(items) == 0:
        # TODO: Raise exception to allow this method's caller to try/except: log and continue from this failure.
        raise SentItemsError(
            f'No Emails found with partial subject "&{unique_id} " for a UniqueID of "{unique_id}"')
    else:
        # TODO: Raise exception to allow this method's caller to try/except: log and continue from this failure.
        raise SentItemsError(f'More than one email has been sent by this user with a subject containing "&{unique_id} "'
                             f' for a UniqueID of "{unique_id}"')


def insert_reminded(unique_id):
    # Update the origin table so that the reminder won't constantly go out.
    cursor = connection.execute("UPDATE \"Origin_Table_Name_Here\" SET RemindedDateTime = ? WHERE \"Unique_ID\" = ?",
                                datetime.now().__format__('%m-%d-%Y %I:%M %p'), unique_id)
    connection.commit()


# This is the method to call to run this whole script.
def re_mailer():
    # Query a purpose-built view for all UniqueID's where:
    #  CreatedDateTime is less than or equal to the result of adding -10 (negative 10) minutes to the current DateTime.
    cursor = connection.execute(
        "SELECT \"UniqueID_Name_Goes_Here\" FROM \"View_Name_Here\" WHERE CreatedDateTime <= DATEADD(minute, -10, GETDATE());")
    remind_list = list(cursor.fetchall())
    for remind in remind_list:
        # TODO: Ensure that the value of each to_remind is actually the right value and not still packed inside a list.
        try:
            to_remind = remind.UniqueID #  TODO: Change the ".UniqueID" to the name of the column previously selected.
            forwarder(get_full_subject(str(to_remind)), str(to_remind))
            # TODO: Note that the following line will not run if the previous line raised an error,
            #  this may or may not be best practice.
            insert_reminded(to_remind)
        except SentItemsError as errorMessage:
            print(errorMessage)
            # TODO: Email that errorMessage (the print statement above has been tested to ensure that it gives just
            #  the message)

            # Continue the loop
            continue
