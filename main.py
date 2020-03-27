from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import pandas as pd
import string
from selenium.webdriver import ChromeOptions, Chrome

username = 'siap_akademik'
password = 'poltek'
tahun = '20191'
jenis = '1'
program = 'REG'

opts = ChromeOptions()
opts.add_experimental_option("detach", True)
browser = Chrome(options=opts)

#define browser
def launchBrowser():
    #open siap
    browser.get("http://siap.poltekpos.ac.id/")

    #login to siap
    user_siap = browser.find_element_by_xpath("//input[@name='user_name']")
    user_siap.send_keys(username)
    pass_siap = browser.find_element_by_xpath("//input[@name='user_pass']")
    pass_siap.send_keys(password)
    login_siap = browser.find_element_by_xpath("//input[@name='login']")
    login_siap.send_keys(Keys.ENTER)

    #choose jadwal menu
    jadwal_menu = browser.find_element_by_link_text("Jadwal Ujian 1")
    jadwal_menu.click()

def convertKelas(kelas):
    kelas = kelas[-1:].lower()
    list_kelas = list(string.ascii_lowercase)
    list_nomor = list(range(1, 27))
    dict_kelas = dict(zip(list_kelas, list_nomor))
    for k, v in dict_kelas.items():
        if k == kelas:
            return v
            break

def setJadwal(tabel_select, matkul_selected, kelas_selected):
    index = 1
    while True:
        index += 1
        matkul_select = tabel_select.find_element_by_xpath(
            "//tr[" + str(index) + "]/td[8]").text.lower()
        time.sleep(1)
        kelas_select = tabel_select.find_element_by_xpath(
            "//tr[" + str(index) + "]/td[9]").text
        time.sleep(1)

        kelas_select = int(kelas_select.strip("0"))

        if matkul_selected == matkul_select and kelas_selected == kelas_select:
            print(matkul_select, kelas_select)
            edit_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[1]/a")
            time.sleep(1)
            edit_select.send_keys(Keys.ENTER)
            return True
            break

def setUjian(prodi_selected, matkul_selected, kelas_selected):
    prodi_selected = prodi_selected.lower()
    matkul_selected = matkul_selected.lower()
    kelas_selected = convertKelas(kelas_selected)

    #choose tahun
    tahun_ujian = Select(browser.find_element_by_xpath("//select[@name='tahun']"))
    tahun_ujian.select_by_value(tahun)
    jenis_ujian = Select(browser.find_element_by_xpath("//select[@name='ujian']"))
    jenis_ujian.select_by_value(jenis)
    program_ujian = Select(browser.find_element_by_xpath("//select[@name='prid']"))
    program_ujian.select_by_value(program)
    prodi_ujian = Select(browser.find_element_by_xpath("//select[@name='prodi']"))
    for prodis in prodi_ujian.options:
        prodi = prodis.text[5:].lower()
        if prodi_selected in prodi:
            prodi_ujian.select_by_value(prodis.text[:2])
            break

    tampil_ujian = browser.find_element_by_xpath("//input[@name='Tampilkan']")
    tampil_ujian.send_keys(Keys.ENTER)

    time.sleep(2)
    tabel_select = browser.find_element_by_xpath("//table[@cellpadding='4' and @cellspacing='1']/tbody")

    if setJadwal(tabel_select, matkul_selected, kelas_selected):
        tanggal_ujian = Select(browser.find_element_by_xpath("//select[@name='TGL_d']"))
        tanggal_ujian.select_by_visible_text('25')
        bulan_ujian = Select(browser.find_element_by_xpath("//select[@name='TGL_m']"))
        bulan_ujian.select_by_value('11')
        tahun_ujian = Select(browser.find_element_by_xpath("//select[@name='TGL_y']"))
        tahun_ujian.select_by_visible_text('2019')
        jam_mulai_ujian = browser.find_element_by_xpath("//input[@name='JM']")
        jam_mulai_ujian.send_keys('08:00:00')
        jam_selesai_ujian = browser.find_element_by_xpath("//input[@name='JS']")
        jam_selesai_ujian.send_keys('10:00:00')
        ruangan_ujian = browser.find_element_by_xpath("//input[@name='RuangID1']")
        ruangan_ujian.send_keys('')
        return True


###### Mulai disini ######

data = pd.read_excel(r'revisi JADWAL UTS GANJIL 2019 2020 - 1.xls',skiprows=3, usecols=['MATAKULIAH', 'KELAS', 'PRODI'])
data = data.dropna()

launchBrowser()

# for index, row in data.iterrows():
#     matkul = row['MATAKULIAH'].lower()
#     kelas = convertKelas(row['KELAS'])
#     prodi = row['PRODI'].lower()
#     print(matkul, kelas, prodi)

setUjian('D4 Teknik Informatika', 'Matematika Diskrit', '2A')


