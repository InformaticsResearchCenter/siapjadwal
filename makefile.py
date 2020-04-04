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
from dosensiap import *


def makeFile(driver, list_prodi_ujian, filters):
    launchJadwalUjianMenu(driver)
    getFile(driver, list_prodi_ujian, filters)
    driver.quit()


def getFile(driver, list_prodi_ujian, filters):
    chooseUjian(driver, filters)
    prodi_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prodi']"))
    for prodi_selected in list_prodi_ujian:
        if getProdiFromDropdown(driver, prodi_selected, filters):
            printAbsensiUjian(driver, filters, prodi_selected)


def printAbsensiUjian(driver, filters, prodi):
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
            matkul_select = matkul_select.replace("-", "_")
            matkul_select = matkul_select.replace("(", "")
            matkul_select = matkul_select.replace(")", "")
            kelas_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[9]").text
            time.sleep(1)
            kode_matkul_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[7]").text
            time.sleep(2)
            email_select = getEmailDosen(
                kode_matkul_select, kelas_select, filters['tahun'])
            kelas_select = int(kelas_select.strip("0"))
            kelas_select = convertKelas(kelas_select)
            filename = ''
            if email_select != None:
                filename = "{}-{}-{}-{}-{}-{}".format(filters['tahun'], setUjian(
                    filters['jenis']), filters['program'], matkul_select, kelas_select, email_select)
            else:
                filename = "{}-{}-{}-{}-{}-NULL".format(filters['tahun'], setUjian(
                    filters['jenis']), filters['program'], matkul_select, kelas_select)
            checkDir(prodi)
            if os.path.exists('absensi/'+prodi+'/'+filename+'.pdf'):
                os.remove('absensi/'+prodi+'/'+filename+'.pdf')
            try:
                edit_select = tabel_select.find_element_by_xpath(
                    "//tr[" + str(index) + "]/td[16]/a")
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
                makePDFOfAbsensiUjian(filename, prodi)
                time.sleep(2)
                if os.path.exists('absensi/'+prodi+'/'+filename+'.txt'):
                    os.remove('absensi/'+prodi+'/'+filename+'.txt')
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(2)
                print('File '+filename+'.pdf berhasil dibuat')
            except NoSuchElementException:
                continue
        except NoSuchElementException:
            break

#########################

def makeFileForDosen(driver, dosens, filters):
    launchJadwalUjianMenu(driver)
    prodis = []
    for dosen in dosens:
        matkul = getMatkulDosen(dosen, filters['tahun'])
        prodis.extend(matkul['prodi'].tolist())
        
    prodis = list(set(prodis))

    all_ujian = getAllUjian(driver, prodis, filters)
    print("\nSelesai mengambil semua ujian")
    for dosen in dosens:
        matkul = getMatkulDosen(dosen, filters['tahun'])
        getFileForDosen(driver, all_ujian, matkul, filters)
    driver.quit()

def getFileForDosen(driver, all_ujian, matkul, filters):
    total = len(matkul)
    gagal = 0
    berhasil = 0
    
    for index, row in matkul.iterrows():
        ujian = all_ujian.loc[(all_ujian['matkul'] == row['nama_matkul']) & (all_ujian['prodi'] == row['prodi'])
                              & (all_ujian['kelas'] == int(row['kelas'].strip("0")))]
        if(ujian.empty):
            print("Jadwal "+row['nama_matkul']+" kelas " +
                  row['kelas']+" tidak ada di SIAP")
            gagal += 1
        else:
            index_ujian = ujian.iloc[0]['index']
            prodi_ujian = ujian.iloc[0]['prodi']
            matkul = {"prodi": prodi_ujian,"index": str(index_ujian)}
            printAbsensiUjianForDosen(driver, matkul, filters)

def printAbsensiUjianForDosen(driver, matkul, filters):
    chooseUjian(driver, filters)
    prodi_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prodi']"))
    for prodis in prodi_ujian.options:
        prodi = prodis.text[5:]
        if matkul['prodi'] == prodi:
            prodi_ujian.select_by_value(prodis.text[:2])
            break
        
    tampil_ujian = driver.find_element_by_xpath("//input[@name='Tampilkan']")
    tampil_ujian.send_keys(Keys.ENTER)
    
    time.sleep(2)
    tabel_select = driver.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(2)
    
    matkul_select = tabel_select.find_element_by_xpath(
        "//tr[" + matkul['index'] + "]/td[8]").text
    time.sleep(1)
    kelas_select = tabel_select.find_element_by_xpath(
        "//tr[" + matkul['index'] + "]/td[9]").text
    time.sleep(1)
    kode_matkul_select = tabel_select.find_element_by_xpath(
        "//tr[" + matkul['index'] + "]/td[7]").text
    time.sleep(2)
    
    email_select = getEmailDosen(
    kode_matkul_select, kelas_select, filters['tahun'])
    kelas_select = int(kelas_select.strip("0"))
    kelas_select = convertKelas(kelas_select)
    matkul_select = matkul_select.replace(" ", "_")
    matkul_select = matkul_select.replace("-", "_")

    filename = ''
    if email_select != None:
        filename = "{}-{}-{}-{}-{}-{}".format(filters['tahun'], setUjian(
            filters['jenis']), filters['program'], matkul_select, kelas_select, email_select)
    else:
        filename = "{}-{}-{}-{}-{}-NULL".format(filters['tahun'], setUjian(
            filters['jenis']), filters['program'], matkul_select, kelas_select)
    
    checkDir(matkul['prodi'])
    
    if os.path.exists('absensi/'+prodi+'/'+filename+'.pdf'):
        os.remove('absensi/'+prodi+'/'+filename+'.pdf')
    try:
        edit_select = tabel_select.find_element_by_xpath(
            "//tr[" + matkul['index'] + "]/td[16]/a")
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
        
        makePDFOfAbsensiUjian(filename, prodi)
        time.sleep(2)
        if os.path.exists('absensi/'+prodi+'/'+filename+'.txt'):
            os.remove('absensi/'+prodi+'/'+filename+'.txt')
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(2)
        print('File '+filename+'.pdf berhasil dibuat')
    except NoSuchElementException:
        print('Jadwal {} kelas {} belum diatur'.format(matkul_select, kelas_select))
            

def launchJadwalUjianMenu(driver):
    driver.get("http://siap.poltekpos.ac.id/")
    user_siap = driver.find_element_by_xpath("//input[@name='user_name']")
    user_siap.send_keys(config.username_siap)
    pass_siap = driver.find_element_by_xpath("//input[@name='user_pass']")
    pass_siap.send_keys(config.password_siap)
    login_siap = driver.find_element_by_xpath("//input[@name='login']")
    login_siap.send_keys(Keys.ENTER)
    jadwal_menu = driver.find_element_by_link_text("Jadwal Ujian 1")
    jadwal_menu.click()

def chooseUjian(driver, filters):
    tahun_ujian = Select(
        driver.find_element_by_xpath("//select[@name='tahun']"))
    tahun_ujian.select_by_value(filters['tahun'])
    jenis_ujian = Select(
        driver.find_element_by_xpath("//select[@name='ujian']"))
    jenis_ujian.select_by_value(filters['jenis'])
    program_ujian = Select(
        driver.find_element_by_xpath("//select[@name='prid']"))
    program_ujian.select_by_value(filters['program'])

def getAllUjian(driver, prodis, filters):
    chooseUjian(driver, filters)
    list_ujian = pd.DataFrame()
    for prodi_selected in prodis:
        if getProdiFromDropdown(driver, prodi_selected):
            tampil_ujian = driver.find_element_by_xpath(
                "//input[@name='Tampilkan']")
            tampil_ujian.send_keys(Keys.ENTER)

            result_generate = genDataFrameUjian(driver, prodi_selected)
            list_ujian = pd.concat([result_generate, list_ujian])

    return list_ujian


def genDataFrameUjian(driver, prodi_selected):
    time.sleep(2)
    tabel_select = driver.find_element_by_xpath(
        "//table[@cellpadding='4' and @cellspacing='1']/tbody")
    time.sleep(2)
    index = 1
    dict_data = []
    while True:
        try:
            index += 1
            matkul_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[8]").text
            time.sleep(1)
            kelas_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[9]").text
            time.sleep(1)
            kode_matkul_select = tabel_select.find_element_by_xpath(
                "//tr[" + str(index) + "]/td[7]").text
            time.sleep(2)
            kelas_select = int(kelas_select.strip("0"))
            data = {'prodi': prodi_selected, 'matkul': matkul_select,
                    'kelas': kelas_select, 'kode_matkul': kode_matkul_select, 'index': index}
            dict_data.append(data)
            print('.', end='', flush=True)
        except NoSuchElementException:
            break

    df_data = pd.DataFrame(dict_data)
    return df_data

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


def makePDFOfAbsensiUjian(filename, prodi):
    pdf = setPdfFormat()
    
    #Read file txt
    path = 'absensi/{}/{}.txt'.format(prodi, filename)
    file = open(path)
    all_lines = file.readlines()

    #Get length lines file txt
    length_lines = len(all_lines)

    #Make first page
    if length_lines < 61:
        makeCellPdf(pdf, all_lines, 1, length_lines)
    else:
        makeCellPdf(pdf, all_lines, 1, 61)
    
    #Make second page if exist
    if length_lines > 61:
        makeCellPdf(pdf, all_lines, 63, length_lines)
    
    pdf.output('absensi/{}/{}.pdf'.format(prodi, filename))
    
def makeCellPdf(pdf, all, begin, end):
    char_re = removeSpecialChar()
    pdf.add_page()
    for i in range(begin, end):
            pdf.cell(0, 4, txt=char_re.sub('', all[i]), ln=1, border=0)
            
def setPdfFormat():
    pdf = FPDF('P', 'mm', 'A4')
    pdf.add_font('Consolas', '', 'consola.ttf', uni=True)
    pdf.set_font("Consolas", size=7.4)
    return pdf

def removeSpecialChar():
    all_chars = (chr(i) for i in range(sys.maxunicode))
    control_chars = ''.join(
        c for c in all_chars if unicodedata.category(c) == 'Cc')
    control_char_re = re.compile('[%s]' % re.escape(control_chars))
    return control_char_re

def convertKelas(kelas):
    list_kelas = list(string.ascii_lowercase)
    list_nomor = list(range(1, 27))
    dict_kelas = dict(zip(list_nomor, list_kelas))
    for k, v in dict_kelas.items():
        if k == kelas:
            return v.upper()
            break
    return 'Kelas tidak terdaftar'

def checkDir(prodi):
    path = 'absensi/{}'.format(prodi)
    if not os.path.exists(path):
        os.makedirs(path)
        print('Direktori {} telah dibuat'.format(prodi))

def setUjian(ujian):
    ujian = int(ujian)
    if ujian == 1:
        return 'UTS'
    elif ujian == 2:
        return 'UAS'
    else:
        return 'XXX'
