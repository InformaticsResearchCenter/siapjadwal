from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pymysql
import config

import email
import smtplib
import ssl
import os 

def dbConnectSiap():
    db = pymysql.connect(config.db_host_siap, config.db_username_siap,
                         config.db_password_siap, config.db_name_siap)
    return db

def getEmailDosen(dosen):
    email = ''
    db = dbConnectSiap()
    sql = "select Email from simak_mst_dosen where Nama = '"+dosen+"'"
    
    with db:
        cur = db.cursor()
        cur.execute(sql)

        rows = cur.fetchone()
        return rows[0]

def sendEmail(file):
    subject = "An email with attachment from Python"
    body = "This is an email with attachment sent from Python"
    
    sender_email = config.email_iteung
    receiver_email = file['tujuan']
    password = config.pass_iteung

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email

    message.attach(MIMEText(body, "plain"))

    os.chdir(r'absensi/'+file['prodi'])
    os.rename(file['nama_lama'], file['nama_baru'])
    filename = file['nama_baru']

    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    message.attach(part)
    text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
        
    os.rename(file['nama_baru'], file['nama_lama'])
    os.chdir(r'../../')
    

def sendFileUjian(list_prodi_ujian, filters):
    for prodi_selected in list_prodi_ujian:
        directory = 'absensi/'+prodi_selected
        for filename in os.listdir(directory):
            if filename.endswith(".pdf") and filename.startswith(filters['tahun']+'-'+setUjian(filters['jenis'])+'-'+filters['program']):
                nama_baru = filename[:-4].split("-")
                email_dosen = nama_baru[5]
                if email_dosen == 'NULL':
                    continue
                else:
                    file = {'nama_lama': filename,
                            'prodi': prodi_selected,
                            'nama_baru': nama_baru[0]+'-'+nama_baru[1]+'-'+nama_baru[2]+'-'+nama_baru[3]+'-'+nama_baru[4]+'.pdf',
                            'tujuan': email_dosen}
                    sendEmail(file)
                    print(email_dosen)
                    continue
            else:
                continue
    

def setUjian(ujian):
    if 1:
        return 'UTS'
    elif 2:
        return 'UAS'
