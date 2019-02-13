import pymssql
import pandas as pd
import pymssql
import datetime
from O365 import Message, Account, FileSystemTokenBackend, oauth_authentication_flow

#creating month/year string for email
today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
lastMonthFinal = lastMonth.strftime("%m-%Y")

#connect to MSSQL
cnxn = pymssql.connect(
    host='HOSTIDHERE',
    user='USERNAMEHERE',
    password='PASSWORDHERE',
    database='DATABASEIDHERE'
)
#Run MSSQL query
cursor = cnxn.cursor()
script = """
select * from whatever with (nolock)
"""
cursor.execute(script)

#write SQL query to excel and name using DATE string
df = pd.read_sql_query(script, cnxn)
writer = pd.ExcelWriter('Query for %s.xlsx' % lastMonthFinal)
df.to_excel(writer, sheet_name='My Query')
writer.save()


#API and Scope for O365
credentials = ('CLIENT_ID','CLIENT_SECRET')
scopes = ['message_all']

#Authenticate O365 Account using API
account = Account(credentials)

if not account.is_authenticated:  # will check if there is a token and has not expired
    # ask for a login
    account.authenticate(scopes=scopes)

#Send message via Office 365
m = account.new_message()
m.to.add('EMAIL ADDRESS HERE')
m.attachments.add('Query for %s.xlsx' % lastMonthFinal)
m.subject = 'Query for ' + lastMonthFinal
m.body = 'Hello USER,<br> <br> Please see the attached QUERY. <br> <br> Regards, <br> <br> QUERYMASTERHERE'
m.send()
