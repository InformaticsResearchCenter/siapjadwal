from inputjadwal import inputJadwalUjian
from printjadwal import *
from selenium.webdriver import ChromeOptions, Chrome

#config selenium
opts = ChromeOptions()
# opts.add_argument("--headless")
opts.add_experimental_option("detach", True)
driver = Chrome(options=opts)

#data yang dibutuhkan
filters = {'tahun': '20191',
           'jenis': '1',
           'program': 'REG'}
list_prodi_ujian = ['D4 Teknik Informatika', 'D3 Teknik Informatika']

#inputjadwal
# inputJadwalUjian(driver, filters, "revisi JADWAL UTS GANJIL 2019 2020 - 1.xls")
#printjadwal
printJadwalUjian(driver, list_prodi_ujian, filters)
