from tableau_api_lib import TableauServerConnection
from tableau_api_lib.utils import querying, flatten_dict_column
from urllib import parse
import pandas as pd

from Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os
from email.mime.base import MIMEBase
from email import encoders

df = pd.read_excel (r'HCA_TEAM_MEMBERS_DATA.xlsx')
df

config = {
    'tableau_prod': {
        'server': 'https://tableauserver.hcaanalytics.com/',
        'api_version': '3.8',
        'personal_access_token_name': 'Ian_test',
        'personal_access_token_secret': 'YxVWp0lGQRa1DC0pQ0/SBQ==:0Ffbq2O0nAecNgVYMgPaAacSc9SGxsaJ',
        'site_name': '',
        'site_url': '' 
    }
}

conn = TableauServerConnection(config_json =  config, env = 'tableau_prod')
response = conn.sign_in()
response.json()

views_df = querying.get_views_dataframe(conn)
views_df = flatten_dict_column(views_df,keys = ['name','id'], col_name = 'workbook')

relevant_views = views_df[views_df['workbook_name'] == 'HCL_commission_auto']
relevant_views

pdf_view_id = relevant_views['id'][888]
pdf_view_id

tabeleau_hcl_member_field_name = parse.quote('HCA Team Member')
tabeleau_hcl_member_field_name

def user_hcl(hcl_member_name):
    tabeleau_hcl_member_field_value = parse.quote(hcl_member_name)
    return tabeleau_hcl_member_field_value
    
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)


members_inact = list()
for x in range(len(df)):
    if df.loc[x].Active_comission_flag == 1:
        #Message Body
        emailMsg = 'Hi '+ str(df.loc[x].HCA_Team_member) + ',\n\nThis is an automated message that includes your current commission. Please do not respond this message  \n\nBest Regards, \nThe Key'
        mimeMessage = MIMEMultipart()
        mimeMessage ['to'] = df.loc[x].Email
        mimeMessage['subject'] = 'Test Message with commission attached'
        mimeMessage.attach(MIMEText(emailMsg,'plain'))
    
        #PDF Creation for each user 
        pdf_params = {
        'pdf_filter': f'vf_{tabeleau_hcl_member_field_name}={user_hcl(df.loc[x].HCA_Team_member)}',
        }
    
        view_pdf_commission = conn.query_view_pdf(view_id = pdf_view_id,parameter_dict = pdf_params)
    
        #PDF Attached to message
        myFile = MIMEBase('application', 'pdf')
        myFile.set_payload(view_pdf_commission.content)
        myFile.add_header('Content-Disposition', 'attachment', filename='pdf1')
        encoders.encode_base64(myFile)
        mimeMessage.attach(myFile)
    
        raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

        message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
        print(message)
    else:
        print(df.loc[x].HCA_Team_member + ' was not included for this iteration')
        members_inact.append(df.loc[x].HCA_Team_member)
        
        
        
        
emailMsg = 'Hi Preston, ' + '\n\nThis is an automated message. Here you have the Members that were excluded from this iteration of the HCL Commissions. Please do not respond this message: \n'+str(members_inact)+'  \n\nBest Regards, \nThe Key'
mimeMessage = MIMEMultipart()
#mimeMessage ['to'] = 'pharlow@homecareassistance.com'
mimeMessage ['to'] = 'iperrilliat@homecareassistance.com'
mimeMessage['subject'] = 'Members Excluded from Commission iteration'
mimeMessage.attach(MIMEText(emailMsg,'plain'))

raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
print(message)