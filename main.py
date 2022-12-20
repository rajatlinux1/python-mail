import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas
import glob
import environ
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent

env = environ.Env(DEBUG=(bool, False))
environ.Env.read_env('config.env')

EXCEL_FILE = glob.glob(f"{str(BASE_DIR)}/*.xlsx")[0]
xls_data = pandas.read_excel(EXCEL_FILE)
receivers_data = xls_data.to_dict('record')


sender = env("EMAIL_HOST_USER")
password = env("EMAIL_HOST_PASSWORD")


# Setup the MIME
message = MIMEMultipart()
message['From'] = sender
files_size = 0

all_files = glob.glob(f"{str(BASE_DIR)}/attachments/*")

has_zip = False
files_names = []
print("\n")
for file in all_files:
    if os.path.isfile(file):
        file_name = file.split('/')[-1]
        if file_name.endswith(".zip") or file_name.endswith(".tar") or file_name.endswith(".rar"):
            print(f">>\"{file_name}\" not attached")
            has_zip = True

        file_stats = os.stat(file)
        files_size += file_stats.st_size

if has_zip:
    print("You can't attach any compressed file for security reasons, You have to uncompressed them")

if files_size > 25000000:
    print(
        f"\nAttached files size too large, you attached {(files_size/1000)/1000:.1f} MB and it accepts only 25 MB")

else:
    print(f"Attached files size {(files_size/1000)/1000:.1f} MB")
    for data in receivers_data:
        receiver = data['Email']
        body = f'''Hello {data["First Name"]},
        This is the body of the email
        Sincerely,
        Testing
        '''
        message['To'] = receiver
        message['Subject'] = 'This is a testing mail from RJT'

        message.attach(MIMEText(body, 'plain'))

        for file in all_files:
            if os.path.isfile(file):
                file_name = file.split('/')[-1]
                file_path = file
                # open the file in bynary
                if file_name.endswith(".zip") or file_name.endswith(".tar") or file_name.endswith(".rar"):
                    continue
                else:
                    files_names.append(file_name)
                    binary_pdf = open(file_path, 'rb')
                    payload = MIMEBase(
                        'application', 'octate-stream', Name=file_name)
                    # payload = MIMEBase('application', 'pdf', Name=file_name)
                    payload.set_payload((binary_pdf).read())

                    # enconding the binary into base64
                    encoders.encode_base64(payload)

                    # add header with pdf name
                    payload.add_header('Content-Decomposition',
                                    'attachment', filename=file_name)
                    message.attach(payload)

        # use gmail with port
        session = smtplib.SMTP(env("SMTP_NAME"), env("SMTP_PORT"))

        # enable security
        session.starttls()

        # login with mail_id and password
        session.login(sender, password)

        text = message.as_string()
        session.sendmail(sender, receiver, text)
        print(f"Mail sent to {data['First Name']}")

    session.quit()
    for file, position in enumerate(files_names, start=1):
        print("\n")
        print(f"\033[1m {position} - {file_name}\033[0m", )
    print('All mail sent successfully')
