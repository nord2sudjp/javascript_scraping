# -*- coding: utf-8 -*-
import sys
import os
import datetime

import configparser

from shutil import copyfile
from shutil import move
from time import sleep 

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select

inifile = configparser.RawConfigParser()
inifile.read('./config.ini', 'UTF-8')
TargetUrl = inifile.get('settings', 'targeturl')
src = inifile.get('settings', 'src')
dst = inifile.get('settings', 'dst')

Browser = None 
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : inifile.get('settings', 'downloadfolder')}
chromeOptions.add_experimental_option("prefs",prefs)

def main():

  global Browser
  Browser = None
  Browser = webdriver.Chrome(executable_path = ".\chromedriver", chrome_options=chromeOptions)
  
  print('Navigating....' + TargetUrl, file=sys.stderr)

  Browser.get(TargetUrl)
  link_to_report = Browser.find_element_by_css_selector('#ctl32_ctl05_ctl04_ctl00_Menu > div:nth-child(6) > a') # This is OK
  print('INFO:main():Following link...', link_to_report.get_attribute("alt"), file=sys.stderr)
  sleep(5)
  
  Browser.execute_script("javascript:$find('ctl32').exportReport('EXCELOPENXML')");
  sleep(10)
  
  Browser.close()
  
  copyfile(src, dst)
  
  datestring = '_' + datetime.datetime.today().strftime("%m%d") + '.xls'
  dst_mv = dst.replace('.xls', datestring)
  move(src, dst_mv)

if __name__ == '__main__':
    main()
