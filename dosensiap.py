import pymysql
import config
import pandas as pd

def dbConnectSiap():
    db = pymysql.connect(config.db_host_siap, config.db_username_siap,
                         config.db_password_siap, config.db_name_siap)
    return db

def getEmailDosen(matkul, kelas, tahun):
    db = dbConnectSiap()
    
    sql = """
    select d.Email,j.JadwalID,j.TahunID,j.NamaKelas,CASE
        WHEN j.ProdiID ='.13.' THEN 'D3 Teknik Informatika'
        WHEN j.ProdiID ='.14.' THEN 'D4 Teknik Informatika'
        WHEN j.ProdiID ='.23.' THEN 'D3 Manajemen Informatika'
        WHEN j.ProdiID ='.33.' THEN 'D3 Akuntansi'
        WHEN j.ProdiID ='.34.' THEN 'D4 Akuntansi Keuangan'
        WHEN j.ProdiID ='.43.' THEN 'D3 Manajemen Pemasaran'
        WHEN j.ProdiID ='.44.' THEN 'D4 Manajemen Perusahaan'
        WHEN j.ProdiID ='.53.' THEN 'D3 Logistik Bisnis'
        WHEN j.ProdiID ='.54.' THEN 'D4 Logistik Bisnis'
        END AS namaprodi,j.MKKode,j.DosenID,d.Nama,j.Nama,j.HariID,
        j.JamMulai,j.JamSelesai,j.DosenID 
        from simak_trn_jadwal as j join simak_mst_dosen as d
        where  j.dosenid=d.login and j.MKKode = '"""+matkul+"""' and j.NamaKelas = '"""+kelas+"""' and TahunID = '"""+tahun+"""';
    """
    
    with db:
        cur = db.cursor()
        cur.execute(sql)
        rows = cur.fetchone()
        if "/" in rows[0]:
            return rows[0][0:rows[0].find("/")-1].replace(" ", "")
        elif rows[0] != None and rows[0] != "":
            return rows[0].replace(" ", "")
        else:
            return 'NULL'

def getMatkulDosen(dosen, tahun):
    db = dbConnectSiap()
    sql = """
    select j.JadwalID,j.TahunID,j.NamaKelas,CASE
        WHEN j.ProdiID ='.13.' THEN 'D3 Teknik Informatika'
        WHEN j.ProdiID ='.14.' THEN 'D4 Teknik Informatika'
        WHEN j.ProdiID ='.23.' THEN 'D3 Manajemen Informatika'
        WHEN j.ProdiID ='.33.' THEN 'D3 Akuntansi'
        WHEN j.ProdiID ='.34.' THEN 'D4 Akuntansi Keuangan'
        WHEN j.ProdiID ='.43.' THEN 'D3 Manajemen Pemasaran'
        WHEN j.ProdiID ='.44.' THEN 'D4 Manajemen Perusahaan'
        WHEN j.ProdiID ='.53.' THEN 'D3 Logistik Bisnis'
        WHEN j.ProdiID ='.54.' THEN 'D4 Logistik Bisnis'
        END AS namaprodi,j.MKKode,j.Nama,j.HariID,
        j.JamMulai,j.JamSelesai,j.DosenID 
        from simak_trn_jadwal as j join simak_mst_dosen as d
        where  j.dosenid=d.login and j.DosenID = '"""+dosen+"""' and TahunID = '"""+tahun+"""';
    """
    
    with db:
        matkul = []
        
        cur = db.cursor()
        cur.execute(sql)
        rows = cur.fetchall()
        
        for row in rows:
            matkul.append([row[2], row[3], row[4], row[5]])
            
    return pd.DataFrame(matkul, columns=['kelas', 'prodi', 'matkul', 'nama_matkul'])

