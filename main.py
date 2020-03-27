from selenium import webdriver
from selenium.webdriver.common.keys import Keys

browser = webdriver.Chrome('E:/IRC/Jadwal/chromedriver')

browser.get("http://siap.poltekpos.ac.id/")
user_siap = browser.find_element_by_xpath("//input[@name='user_name']")
user_siap.send_keys('siap_akademik')
pass_siap = browser.find_element_by_xpath("//input[@name='user_pass']")
pass_siap.send_keys('poltek')
login_siap = browser.find_element_by_xpath("//input[@name='login']")
login_siap.send_keys(Keys.ENTER)
