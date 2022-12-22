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
sender = env("EMAIL_HOST_USER")
password = env("EMAIL_HOST_PASSWORD")


def home():
    env_file = glob.glob(f"{str(BASE_DIR)}/*.env")
    excel_file = glob.glob(f"{str(BASE_DIR)}/*.xlsx")
    all_files = glob.glob(f"{str(BASE_DIR)}/*")
    directory = []
    EXCEL_FILE_NAME = ''
    ATTACHMENT_DIR_NAME = ''
    message = MIMEMultipart()
    message['From'] = sender

    if not env_file:
        raise FileNotFoundError("Please configure .env file in base directory with these variables\nEMAIL_HOST_USER=\nEMAIL_HOST_PASSWORD=\nSMTP_NAME=\nSMTP_PORT=")
    
    if not env("EMAIL_HOST_USER"):
        raise ValueError("EMAIL_HOST_USER not available in env file")

    if not env("EMAIL_HOST_PASSWORD"):
        raise ValueError("EMAIL_HOST_PASSWORD not available in env file")

    if not env("SMTP_NAME"):
        raise ValueError("SMTP_NAME not available in env file")

    if not env("SMTP_PORT"):
        raise ValueError("SMTP_PORT not available in env file")

    if not excel_file:
        raise FileNotFoundError("Please provide emails excel file in base directory")

    for file in all_files:
        if not os.path.isfile(file):
            directory.append(file)

    for position, name in enumerate(excel_file, start=1):
        print(position, name.split("/")[-1])

    desired_number = input("choose your targeted excel file : ")

    if desired_number.isdigit():
        EXCEL_FILE_NAME = excel_file[int(desired_number)-1]

    for position, name in enumerate(directory, start=1):
        print(position, name.split("/")[-1])

    desired_number = input("choose your targeted attachment directory : ")

    if desired_number.isdigit():
        ATTACHMENT_DIR_NAME = directory[int(desired_number)-1]

    action(EXCEL_FILE_NAME, ATTACHMENT_DIR_NAME)


def action(excel, attachment):
    EXCEL_FILE = glob.glob(f"{excel}")[0]
    xls_data = pandas.read_excel(EXCEL_FILE)
    receivers_data = xls_data.to_dict('records')

    sender = env("EMAIL_HOST_USER")
    password = env("EMAIL_HOST_PASSWORD")
    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender
    files_size = 0

    all_files = glob.glob(f"{attachment}/*")


    has_zip = False
    files_names = []

    for file in all_files:
        if os.path.isfile(file):
            file_name = file.split('/')[-1]
            if file_name.endswith(".zip") or file_name.endswith(".tar") or file_name.endswith(".rar"):
                print(f">>\"{file_name}\" not attached")
                has_zip = True

            file_stats = os.stat(file)
            files_size += file_stats.st_size

    if has_zip:
        print(f"\033[91m You can't attach any compressed file for security reasons, You have to uncompressed them \033[0m", )

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
                        if file_name not in files_names:
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
            print(f"Mailed to {data['First Name']}")

        session.quit()
        for position, file in enumerate(files_names, start=1):
            print(f"\033[1m {position} - {file}\033[0m", )
        print("\033[92m All mailed successfully \033[0m", )


if __name__ == "__main__":
    home()
