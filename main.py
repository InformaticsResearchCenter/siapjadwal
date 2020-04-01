from inputjadwal import inputJadwalUjian
from printjadwal import *
from selenium.webdriver import ChromeOptions, Chrome

#config selenium
opts = ChromeOptions()
# opts.add_argument("--headless")
opts.add_experimental_option("detach", True)
driver = Chrome(options=opts)

#data yang dibutuhkan
filters = {'tahun': '20192',
           'jenis': '1',
           'program': 'REG'}
list_prodi_ujian = ['D4 Teknik Informatika', 'D3 Teknik Informatika']


# inputJadwalUjian(driver, filters, "jadwal_uts_coba.xls")

printJadwalUjian(driver, list_prodi_ujian, filters)


