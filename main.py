##### Cara Kerja #####
# 1. Siapin data dummy untuk SIAP
# 2. Ubah data xls ke data frame
# 3. Ambil semua data ujian di tiap prodi sesuai xls
# 4. Cek valid data xls ada di siap
# 5. Masukkan data ke siap jika data valid

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
import string
from selenium.webdriver import ChromeOptions, Chrome
from selenium.common.exceptions import NoSuchElementException

# Data dummy
username = 'xxxx'
password = 'xxxx'
tahun = '20191'
jenis = '1'
program = 'REG'

# Jalanin Chrome
opts = ChromeOptions()
opts.add_experimental_option("detach", True)
browser = Chrome(options=opts)

# Mengisi jadwal sesuai dengan excel


def fillJadwal(ujian):
    tanggal_ujian = Select(
        browser.find_element_by_xpath("//select[@name='TGL_d']"))
    tanggal_ujian.select_by_visible_text(ujian['tanggal'])
    bulan_ujian = Select(
        browser.find_element_by_xpath("//select[@name='TGL_m']"))
    bulan_ujian.select_by_value(ujian['bulan'])
    tahun_ujian = Select(
        browser.find_element_by_xpath("//select[@name='TGL_y']"))
    tahun_ujian.select_by_visible_text(ujian['tahun'])
    jam_mulai_ujian = browser.find_element_by_xpath("//input[@name='JM']")
    jam_mulai_ujian.clear()
    jam_mulai_ujian.send_keys(ujian['waktu mulai'])
    jam_selesai_ujian = browser.find_element_by_xpath(
        "//input[@name='JS']")
    jam_selesai_ujian.clear()
    jam_selesai_ujian.send_keys(ujian['waktu selesai'])
    ruangan_ujian = browser.find_element_by_xpath(
        "//input[@name='RuangID1']")
    ruangan_ujian.clear()
    ruangan_ujian.send_keys(ujian['ruang'])
    simpan_jadwal_ujian = browser.find_element_by_xpath(
        "//input[@name='Simpan']")
    simpan_jadwal_ujian.send_keys(Keys.ENTER)
    time.sleep(1)

# Pilih ujian, tampilkan sesuai data, edit ujian


def chooseUjian(ujian):
    tahun_ujian = Select(
        browser.find_element_by_xpath("//select[@name='tahun']"))
    tahun_ujian.select_by_value(tahun)
    jenis_ujian = Select(
        browser.find_element_by_xpath("//select[@name='ujian']"))
    jenis_ujian.select_by_value(jenis)
    program_ujian = Select(
        browser.find_element_by_xpath("//select[@name='prid']"))
    program_ujian.select_by_value(program)
    prodi_ujian = Select(
        browser.find_element_by_xpath("//select[@name='prodi']"))

    for prodis in prodi_ujian.options:
        prodi = prodis.text[5:].lower()
        if ujian['prodi'] in prodi:
            prodi_ujian.select_by_value(prodis.text[:2])
            break
    tampil_ujian = browser.find_element_by_xpath("//input[@name='Tampilkan']")
    tampil_ujian.send_keys(Keys.ENTER)

    time.sleep(2)
    tabel_select = browser.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(2)

    edit_select = tabel_select.find_element_by_xpath(
        "//tr[" + ujian['index'] + "]/td[1]/a")
    time.sleep(1)
    edit_select.send_keys(Keys.ENTER)

    fillJadwal(ujian)

# Membuat data frame semua ujian per prodi


def generateDataFrameUjian(prodi_selected):
    tabel_select = browser.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(1)
    index = 1
    dict_data = []
    while True:
        try:
            index += 1
            matkul_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[8]").text.lower()
            time.sleep(1)
            kelas_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[9]").text
            time.sleep(1)
            kelas_select = int(kelas_select.strip("0"))
            data = {'prodi': prodi_selected, 'matkul': matkul_select,
                    'kelas': kelas_select, 'index': index}
            dict_data.append(data)
            print(str(index))
        except NoSuchElementException:
            break

    df_data = pd.DataFrame(dict_data)
    return df_data

# Membuat data frame semua ujian


def getAllUjianFromSiap(column):
    tahun_ujian = Select(
        browser.find_element_by_xpath("//select[@name='tahun']"))
    tahun_ujian.select_by_value(tahun)
    jenis_ujian = Select(
        browser.find_element_by_xpath("//select[@name='ujian']"))
    jenis_ujian.select_by_value(jenis)
    program_ujian = Select(
        browser.find_element_by_xpath("//select[@name='prid']"))
    program_ujian.select_by_value(program)

    list_ujian = pd.DataFrame()
    for prodi_selected in column:
        prodi_ujian = Select(
            browser.find_element_by_xpath("//select[@name='prodi']"))
        for prodis in prodi_ujian.options:
            prodi = prodis.text[5:].lower()
            prodi_selected = prodi_selected.lower()
            if prodi_selected == prodi:
                prodi_ujian.select_by_value(prodis.text[:2])
                break
        tampil_ujian = browser.find_element_by_xpath(
            "//input[@name='Tampilkan']")
        tampil_ujian.send_keys(Keys.ENTER)

        data = generateDataFrameUjian(prodi_selected)
        list_ujian = pd.concat([data, list_ujian])

    return list_ujian

# Buka SIAP, pilih menu jadwal


def launchSiapToJadwal():
    browser.get("http://siap.poltekpos.ac.id/")
    user_siap = browser.find_element_by_xpath("//input[@name='user_name']")
    user_siap.send_keys(username)
    pass_siap = browser.find_element_by_xpath("//input[@name='user_pass']")
    pass_siap.send_keys(password)
    login_siap = browser.find_element_by_xpath("//input[@name='login']")
    login_siap.send_keys(Keys.ENTER)
    jadwal_menu = browser.find_element_by_link_text("Jadwal Ujian 1")
    jadwal_menu.click()

# Konversi kelas sesuai ketentuan SIAP


def convertKelas(kelas):
    kelas = kelas[-1:].lower()
    list_kelas = list(string.ascii_lowercase)
    list_nomor = list(range(1, 27))
    dict_kelas = dict(zip(list_kelas, list_nomor))
    for k, v in dict_kelas.items():
        if k == kelas:
            return v
            break

# Set ruang


def checkTypeRuang(ruang):
    if type(ruang) is str:
        return ruang
    else:
        return str(int(ruang))


###### Mulai disini ######
# Ambil data xls
data = pd.read_excel(r'revisi JADWAL UTS GANJIL 2019 2020 - 1.xls',
                     skiprows=3, usecols=['MATAKULIAH', 'KELAS', 'PRODI', 'TANGGAL', 'WAKTU', 'RUANG'])
# Hapus NA
data = data.dropna()
# Ambil list prodi
column = data['PRODI'].unique()
# Jalankan siap
launchSiapToJadwal()
# Ambil semua ujian sesuai prodi
all_ujian = getAllUjianFromSiap(column)
# Masukin jadwal ujian

total = len(data)
gagal = 0
berhasil = 0

for index, row in data.iterrows():
    ujian = all_ujian.loc[(all_ujian['matkul'] == row['MATAKULIAH'].lower()) & (
        all_ujian['kelas'] == convertKelas(row['KELAS']))]
    if(ujian.empty):
        print("Jadwal "+row['MATAKULIAH']+" kelas " +
              row['KELAS']+" tidak ada di SIAP")
        gagal += 1
    else:
        index_ujian = ujian.iloc[0]['index']
        tanggal = row['TANGGAL'].strftime('%d/%m/%Y').split("/")
        waktu = row['WAKTU'].split("-")
        ujian = {"matkul": row['MATAKULIAH'].lower(),
                 "prodi": row['PRODI'].lower(),
                 "kelas": convertKelas(row['KELAS']),
                 "tanggal": tanggal[0],
                 "bulan": tanggal[1],
                 "tahun": tanggal[2],
                 "waktu mulai": waktu[0],
                 "waktu selesai": waktu[1],
                 "ruang": checkTypeRuang(row['RUANG']),
                 "index": str(index_ujian)}
        chooseUjian(ujian)
        print("Jadwal "+row['MATAKULIAH']+" kelas " +
              row['KELAS']+" berhasil diinput")
        berhasil += 1

print("Total {}, gagal {}, berhasil {}".format(total, gagal, berhasil))
