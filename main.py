from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

#define browser
browser = webdriver.Chrome('E:/IRC/Jadwal/chromedriver')

#open siap
browser.get("http://siap.poltekpos.ac.id/")

#login to siap
user_siap = browser.find_element_by_xpath("//input[@name='user_name']")
user_siap.send_keys('siap_akademik')
pass_siap = browser.find_element_by_xpath("//input[@name='user_pass']")
pass_siap.send_keys('poltek')
login_siap = browser.find_element_by_xpath("//input[@name='login']")
login_siap.send_keys(Keys.ENTER)

time.sleep(2)
#choose jadwal menu
jadwal_menu = browser.find_element_by_link_text("Jadwal Ujian 1")
jadwal_menu.click()
