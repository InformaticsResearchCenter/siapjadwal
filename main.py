from inputjadwal import inputJadwalUjian
from selenium.webdriver import ChromeOptions, Chrome
from emaildosen import *
from makefile import *

###########################
###### Buat Selenium ######
###########################

opts = ChromeOptions()
opts.add_argument("--headless")
opts.add_experimental_option("detach", True)
driver = Chrome(options=opts)

############################
###### Buat Per Prodi ######
############################

filters = {'tahun': '20192',
           'jenis': '1',
           'program': 'REG'}
prodis = ['D4 Teknik Informatika', 
                    'D4 Manajemen Perusahaan',
                    'D3 Logistik Bisnis', 
                    'D4 Logistik Bisnis']

# inputJadwalUjian(driver, filters, "jadwal_uts_coba.xls")
# Buat generate PDF
# makeFile(driver, prodis, filters)
# Buat ngirim pdf ke email
# sendFileUjian(prodis, filters)

############################
###### Buat Per Dosen ######
############################

filters = {'tahun': '20192',
           'jenis': '1',
           'program': 'REG'}
dosens = ['NN155L',
          'NN222L']

# Buat generate pdf
makeFileForDosen(driver, dosens, filters)
# Buat ngirim pdf ke email
sendFileUjianDosen(dosens, filters)
