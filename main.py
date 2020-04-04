from inputjadwal import inputJadwalUjian
from selenium.webdriver import ChromeOptions, Chrome
from emaildosen import *
from makefile import *

opts = ChromeOptions()
opts.add_argument("--headless")
opts.add_experimental_option("detach", True)
driver = Chrome(options=opts)

# data yang dibutuhkan
filters = {'tahun': '20192',
           'jenis': '1',
           'program': 'REG'}
# list_prodi_ujian = ['D4 Teknik Informatika', 'D4 Manajemen Perusahaan',
#                     'D3 Logistik Bisnis', 'D4 Logistik Bisnis']

# inputJadwalUjian(driver, filters, "jadwal_uts_coba.xls")

# makeFile(driver, list_prodi_ujian, filters)

# sendFileUjian(list_prodi_ujian, filters)

dosens = ['NN155L', 'NN056L', 'NN222L']

makeFileForDosen(driver, dosens, filters)

# sendFileUjianDosen(dosens, filters)
