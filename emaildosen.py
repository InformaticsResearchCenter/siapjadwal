from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pymysql
import config
from dosensiap import *
import email
import smtplib
import ssl
import os 
import string

def sendEmail(file):
    subject = "Absensi {} Mata Kuliah {} Kelas ".format(file['ujian'], file['matkul'], file['kelas'])
    body = "Ini baru percobaan pengiriman file absensi oleh iteung ya... :)"
    
    sender_email = config.email_iteung
    receiver_email = file['tujuan']
    password = config.pass_iteung

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email

    message.attach(MIMEText(body, "plain"))

    os.chdir(r'absensi/{}'.format(file['prodi']))
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
                            'tujuan': email_dosen,
                            'ujian': nama_baru[1],
                            'matkul': nama_baru[3],
                            'kelas': nama_baru[4]}
                    sendEmail(file)
                    print('File '+filename+' berhasil dikirim ke '+email_dosen)
    

def sendFileUjianDosen(dosens, filters):
    for dosen in dosens:
        matkul = getMatkulDosen(dosen, filters['tahun'])
        for prodi_selected in matkul['prodi'].unique():
            directory = 'absensi/'+prodi_selected
            for ind in matkul.index:
                for filename in os.listdir(directory):
                    if filename.endswith(".pdf") and filename.startswith(filters['tahun']+'-'+setUjian(filters['jenis'])+'-'+filters['program']):
                        nama_baru = filename[:-4].split("-")
                        email_dosen = nama_baru[5]
                        matkul_select = matkul['nama_matkul'][ind].replace(" ", "_")
                        matkul_select = matkul_select.replace("-", "_")
                        
                        if nama_baru[3] == matkul_select and nama_baru[4] == convertKelas(int(matkul['kelas'][ind].strip("0"))):
                            
                            if email_dosen == 'NULL':
                                continue
                            else:
                                file = {'nama_lama': filename,
                                        'prodi': prodi_selected,
                                        'nama_baru': nama_baru[0]+'-'+nama_baru[1]+'-'+nama_baru[2]+'-'+nama_baru[3]+'-'+nama_baru[4]+'.pdf',
                                        'tujuan': email_dosen,
                                        'ujian': nama_baru[1],
                                        'matkul': nama_baru[3],
                                        'kelas': nama_baru[4]}
                                sendEmail(file)
                                print('File '+filename+' berhasil dikirim ke '+email_dosen)


def setUjian(ujian):
    ujian = int(ujian)
    if ujian == 1:
        return 'UTS'
    elif ujian == 2:
        return 'UAS'
    else:
        return 'XXX'

def convertKelas(kelas):
    list_kelas = list(string.ascii_lowercase)
    list_nomor = list(range(1, 27))
    dict_kelas = dict(zip(list_nomor, list_kelas))
    for k, v in dict_kelas.items():
        if k == kelas:
            return v.upper()
            break
    return "Kelas tidak terdaftar"
