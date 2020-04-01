import pymysql
import config
import smtplib

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

def sent
