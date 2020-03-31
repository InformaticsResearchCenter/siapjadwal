import pandas as pd
import config
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
from selenium.common.exceptions import NoSuchElementException
import string
import sys


def inputJadwalUjian(driver, filters, filename):
    data_ujian = prepareFileJadwal(filename)
    launchJadwalMenu(driver)
    all_ujian = getAllUjianFromSiap(driver, data_ujian['prodi'], filters)
    print(inputJadwal(driver, all_ujian, data_ujian['data'], filters))
    driver.quit()

def prepareFileJadwal(filejadwal):
    data = pd.read_excel(filejadwal,
                         skiprows=3, usecols=['MATAKULIAH', 'KELAS', 'PRODI', 'TANGGAL', 'WAKTU', 'RUANG'])
    data = data.dropna()
    prodi = data['PRODI'].unique()
    return {'data': data, 'prodi': prodi}


def launchJadwalMenu(driver):
    driver.get("http://siap.poltekpos.ac.id/")
    user_siap = driver.find_element_by_xpath("//input[@name='user_name']")
    user_siap.send_keys(config.username_siap)
    pass_siap = driver.find_element_by_xpath("//input[@name='user_pass']")
    pass_siap.send_keys(config.password_siap)
    login_siap = driver.find_element_by_xpath("//input[@name='login']")
    login_siap.send_keys(Keys.ENTER)
    jadwal_menu = driver.find_element_by_link_text("Jadwal Ujian 1")
    jadwal_menu.click()


def getAllUjianFromSiap(driver, prodis, filters):
    tahun_ujian = Select(
        driver.find_element_by_xpath("//select[@name='tahun']"))
    tahun_ujian.select_by_value(filters['tahun'])
    jenis_ujian = Select(
        driver.find_element_by_xpath("//select[@name='ujian']"))
    jenis_ujian.select_by_value(filters['jenis'])
    program_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prid']"))
    program_ujian.select_by_value(filters['program'])

    list_ujian = pd.DataFrame()
    for prodi_selected in prodis:
        if getProdiFromDropdown(driver, prodi_selected):
            tampil_ujian = driver.find_element_by_xpath(
                "//input[@name='Tampilkan']")
            tampil_ujian.send_keys(Keys.ENTER)

            result_generate = generateDataFrameUjian(driver, prodi_selected)
            list_ujian = pd.concat([result_generate, list_ujian])

    return list_ujian


def getProdiFromDropdown(driver, prodi_selected):
    prodi_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prodi']"))
    for prodis in prodi_ujian.options:
        prodi = prodis.text[5:].lower()
        prodi_selected = prodi_selected.lower()
        if prodi_selected == prodi:
            prodi_ujian.select_by_value(prodis.text[:2])
            return True
            break
    print('Prodi {} tidak ada'.format(prodis.text[5:]))
    return False


def generateDataFrameUjian(driver, prodi_selected):
    time.sleep(3)
    tabel_select = driver.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(3)
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
        except NoSuchElementException:
            break

    df_data = pd.DataFrame(dict_data)
    return df_data


def inputJadwal(driver, all_ujian, data_ujian, filters):
    total = len(data_ujian)
    gagal = 0
    berhasil = 0

    for index, row in data_ujian.iterrows():
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
            chooseUjian(driver, ujian, filters)
            print("Jadwal "+row['MATAKULIAH']+" kelas " +
                  row['KELAS']+" berhasil diinput")
            berhasil += 1
    return "Total {}, gagal {}, berhasil {}".format(total, gagal, berhasil)


def chooseUjian(driver, ujian_data, filters):
    tahun_ujian = Select(
        driver.find_element_by_xpath("//select[@name='tahun']"))
    tahun_ujian.select_by_value(filters['tahun'])
    jenis_ujian = Select(
        driver.find_element_by_xpath("//select[@name='ujian']"))
    jenis_ujian.select_by_value(filters['jenis'])
    program_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prid']"))
    program_ujian.select_by_value(filters['program'])
    prodi_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prodi']"))

    for prodis in prodi_ujian.options:
        prodi = prodis.text[5:].lower()
        if ujian_data['prodi'] in prodi:
            prodi_ujian.select_by_value(prodis.text[:2])
            break
    tampil_ujian = driver.find_element_by_xpath("//input[@name='Tampilkan']")
    tampil_ujian.send_keys(Keys.ENTER)

    time.sleep(3)
    tabel_select = driver.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(3)

    edit_select = tabel_select.find_element_by_xpath(
        "//tr[" + ujian_data['index'] + "]/td[1]/a")
    time.sleep(1)
    edit_select.send_keys(Keys.ENTER)

    fillJadwalUjian(driver, ujian_data)


def fillJadwalUjian(driver, ujian_data):
    tanggal_ujian = Select(
        driver.find_element_by_xpath("//select[@name='TGL_d']"))
    tanggal_ujian.select_by_visible_text(ujian_data['tanggal'])
    bulan_ujian = Select(
        driver.find_element_by_xpath("//select[@name='TGL_m']"))
    bulan_ujian.select_by_value(ujian_data['bulan'])
    tahun_ujian = Select(
        driver.find_element_by_xpath("//select[@name='TGL_y']"))
    tahun_ujian.select_by_visible_text(ujian_data['tahun'])
    jam_mulai_ujian = driver.find_element_by_xpath("//input[@name='JM']")
    jam_mulai_ujian.clear()
    jam_mulai_ujian.send_keys(ujian_data['waktu mulai'])
    jam_selesai_ujian = driver.find_element_by_xpath(
        "//input[@name='JS']")
    jam_selesai_ujian.clear()
    jam_selesai_ujian.send_keys(ujian_data['waktu selesai'])
    ruangan_ujian = driver.find_element_by_xpath(
        "//input[@name='RuangID1']")
    ruangan_ujian.clear()
    ruangan_ujian.send_keys(ujian_data['ruang'])
    simpan_jadwal_ujian = driver.find_element_by_xpath(
        "//input[@name='Simpan']")
    simpan_jadwal_ujian.send_keys(Keys.ENTER)
    time.sleep(1)


def convertKelas(kelas):
    kelas = kelas[-1:].lower()
    list_kelas = list(string.ascii_lowercase)
    list_nomor = list(range(1, 27))
    dict_kelas = dict(zip(list_kelas, list_nomor))
    for k, v in dict_kelas.items():
        if k == kelas:
            return v
            break
    return "Kelas tidak terdaftar"


def checkTypeRuang(ruang):
    if type(ruang) is str:
        return ruang
    else:
        return str(int(ruang))
