from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time

username = 'siap_akademik'
password = 'poltek'
tahun = '20192'
jenis = '1'
program = 'REG'
prodi = '14'

#define browser
browser = webdriver.Chrome('E:/IRC/Jadwal/chromedriver')

#open siap
browser.get("http://siap.poltekpos.ac.id/")

#login to siap
time.sleep(1)
user_siap = browser.find_element_by_xpath("//input[@name='user_name']")
user_siap.send_keys(username)
pass_siap = browser.find_element_by_xpath("//input[@name='user_pass']")
pass_siap.send_keys(password)
login_siap = browser.find_element_by_xpath("//input[@name='login']")
login_siap.send_keys(Keys.ENTER)

#choose jadwal menu
time.sleep(1)
jadwal_menu = browser.find_element_by_link_text("Jadwal Ujian 1")
jadwal_menu.click()

#choose tahun
time.sleep(1)
tahun_ujian = Select(browser.find_element_by_xpath("//select[@name='tahun']"))
tahun_ujian.select_by_value(tahun)
jenis_ujian = Select(browser.find_element_by_xpath("//select[@name='ujian']"))
jenis_ujian.select_by_value(jenis)
program_ujian = Select(browser.find_element_by_xpath("//select[@name='prid']"))
program_ujian.select_by_value(program)
prodi_ujian = Select(browser.find_element_by_xpath("//select[@name='prodi']"))
prodi_ujian.select_by_value(prodi)
tampil_ujian = browser.find_element_by_xpath("//input[@name='Tampilkan']")
tampil_ujian.send_keys(Keys.ENTER)



