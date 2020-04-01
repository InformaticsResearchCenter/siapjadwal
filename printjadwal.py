import config
from selenium.webdriver.common.keys import Keys
import time
import unicodedata
import re
import urllib.request
import os
import sys
from fpdf import FPDF
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
import string
from emaildosen import getEmailDosen


def printJadwalUjian(driver, list_prodi_ujian, filters):
    launchJadwalMenuPrint(driver)
    printAbsensiUjian(driver, list_prodi_ujian, filters)
    driver.quit()


def launchJadwalMenuPrint(driver):
    driver.get("http://siap.poltekpos.ac.id/")
    user_siap = driver.find_element_by_xpath("//input[@name='user_name']")
    user_siap.send_keys(config.username_siap)
    pass_siap = driver.find_element_by_xpath("//input[@name='user_pass']")
    pass_siap.send_keys(config.password_siap)
    login_siap = driver.find_element_by_xpath("//input[@name='login']")
    login_siap.send_keys(Keys.ENTER)
    jadwal_menu = driver.find_element_by_link_text("Jadwal Ujian 1")
    jadwal_menu.click()


def printAbsensiUjian(driver, list_prodi_ujian, filters):
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

    for prodi_selected in list_prodi_ujian:
        if getProdiFromDropdown(driver, prodi_selected, filters):
            continue
        else:
            break


def getProdiFromDropdown(driver, prodi_selected, filters):
    prodi_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prodi']"))
    for prodis in prodi_ujian.options:
        prodi = prodis.text[5:].lower()
        prodi_select = prodi_selected.lower()
        if prodi_select == prodi:
            prodi_ujian.select_by_value(prodis.text[:2])
            tampil_ujian = driver.find_element_by_xpath(
                "//input[@name='Tampilkan']")
            tampil_ujian.send_keys(Keys.ENTER)
            getAbsensiPDFUjian(driver, filters, prodi_selected)
            return True
            break
    print('Prodi {} tidak ada'.format(prodis.text[5:]))
    return False


def getAbsensiPDFUjian(driver, filters, prodi):
    time.sleep(3)
    tabel_select = driver.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(3)
    index = 1
    while True:
        try:
            index += 1
            matkul_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[8]").text
            time.sleep(1)
            matkul_select = matkul_select.replace(" ", "_")
            kelas_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[9]").text
            time.sleep(1)
            kelas_select = int(kelas_select.strip("0"))
            kelas_select = convertKelas(kelas_select)
            nama_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[13]").text
            time.sleep(2)
            print(str(index))
            email_select = getEmailDosen(nama_select)
            filename = ''
            if email_select != None:
                filename = "{}-{}-{}-{}-{}-{}".format(filters['tahun'], setUjian(filters['jenis']), filters['program'], matkul_select, kelas_select, email_select)
            else:
                filename = "{}-{}-{}-{}-{}-NULL".format(filters['tahun'], setUjian(filters['jenis']), filters['program'], matkul_select, kelas_select)

            checkDir(prodi)
            if os.path.exists('absensi/'+prodi+'/'+filename+'.pdf'):
                continue
            else:
                try:
                    edit_select = tabel_select.find_element_by_xpath(
                        "//tr[" + str(index) + "]/td[14]/a")
                    print(str(index))
                    time.sleep(1)
                    edit_select.send_keys(Keys.ENTER)
                    time.sleep(2)
                    driver.switch_to.window(driver.window_handles[1])
                    time.sleep(2)
                    url_select = driver.find_element_by_link_text(
                        "Cetak Laporan").get_attribute('href')
                    urllib.request.urlretrieve(
                        url_select, 'absensi/'+prodi+'/'+filename+'.txt')
                    time.sleep(2)
                    makeAbsensiPDFUjian(filename, prodi)
                    time.sleep(2)
                    if os.path.exists('absensi/'+prodi+'/'+filename+'.txt'):
                        os.remove('absensi/'+prodi+'/'+filename+'.txt')
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    time.sleep(2)
                    print(filename)
                except NoSuchElementException:
                    continue
        except NoSuchElementException:
            break


def makeAbsensiPDFUjian(filename, prodi):
    all_chars = (chr(i) for i in range(sys.maxunicode))
    control_chars = ''.join(
        c for c in all_chars if unicodedata.category(c) == 'Cc')
    control_char_re = re.compile('[%s]' % re.escape(control_chars))

    pdf = FPDF('P', 'mm', 'A4')
    pdf.add_font('Consolas', '', 'consola.ttf', uni=True)
    pdf.set_font("Consolas", size=7.4)
    file = open('absensi/'+prodi+'/'+filename+'.txt')
    all_lines = file.readlines()

    pdf.add_page()
    if len(all_lines) < 61:
        for i in range(1, len(all_lines)):
            pdf.cell(0, 4, txt=control_char_re.sub(
                '', all_lines[i]), ln=1, border=0)
    else:
        for i in range(1, 61):
            pdf.cell(0, 4, txt=control_char_re.sub(
                '', all_lines[i]), ln=1, border=0)

    if len(all_lines) > 61:
        pdf.add_page()
        for i in range(63, len(all_lines)):
            pdf.cell(0, 4, txt=control_char_re.sub(
                '', all_lines[i]), ln=1, border=0)
    pdf.output('absensi/'+prodi+'/'+filename+'.pdf')


def convertKelas(kelas):
    list_kelas = list(string.ascii_lowercase)
    list_nomor = list(range(1, 27))
    dict_kelas = dict(zip(list_nomor, list_kelas))
    for k, v in dict_kelas.items():
        if k == kelas:
            return v.upper()
            break
    return "Kelas tidak terdaftar"


def checkDir(prodi):
    if os.path.exists('absensi/'+prodi):
        return True
    else:
        os.makedirs('absensi/'+prodi)
        return True


def setUjian(ujian):
    if 1:
        return 'UTS'
    elif 2:
        return 'UAS'
