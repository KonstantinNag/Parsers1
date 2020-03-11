from selenium import webdriver
import requests
import datetime
import io
import re
import os
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import json
from sqlalchemy import desc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import configparser
import sqlalchemy
import psycopg2

def rename_file(inn):
    files = os.listdir('D:\\Data\\pdf\\')
    for item in files:
        if re.search(inn, item):
            file_name = item
            break
    try:
        os.rename('D:\\Data\\pdf\\' + file_name, 'D:\\Data\\send\\' + str(inn) + '.pdf')
    except:
        print('уже существует')
    return 'D:\\Data\\send\\' + str(inn) + '.pdf'
def download_pdf():
    fp = webdriver.FirefoxProfile()
    mime_types = "application/pdf,application/vnd.adobe.xfdf,application/vnd.fdf,application/vnd.adobe.xdp+xml"
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)
    fp.set_preference("browser.download.dir", "D:\\Data\\pdf\\")
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
    fp.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)
    fp.set_preference("pdfjs.disabled", True)

    url='https://pb.nalog.ru/'
    driver = webdriver.Firefox(firefox_profile=fp, executable_path='D:\Anaconda3\geckodriver-v0.26.0-win64\geckodriver.exe')
    driver.get(url)
    editor = driver.find_element_by_id('query')
    search = 'Домашний интерьер Москва'
    editor.send_keys(search)
    find_block = 'search.html#quick-result?query='
    tail = '&mode=quick&page=1&pageSize=10'
    transition = url + find_block + search + tail
    driver.get(transition)
    editor = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME , 'result-group')))

    link = editor.get_attribute('data-href')
    if link != None:
        link = url + link
        driver.get(link)
        list_data = driver.find_elements_by_class_name('field')
    vipiska = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div[3]/div/div[1]/div/div/div[2]/div[2]/a[1]'))).click()
    full_title = driver.find_element_by_xpath('//*[@id="pnlCompany"]/div/div[1]/div/div[2]/div[2]/a').text
    name = driver.find_element_by_xpath('//*[@id="pnlCompany"]/div/div[1]/div/div[3]/div[2]/a').text
    inn = driver.find_element_by_xpath('//*[@id="pnlCompany"]/div/div[1]/div/div[5]/div[1]/div/div[2]/a').text
    kind_activity = driver.find_element_by_xpath('//*[@id="pnlCompany"]/div/div[1]/div/div[7]/div[2]/a').text
    address = driver.find_element_by_xpath('//*[@id="pnlCompany"]/div/div[1]/div/div[8]/div[1]/div/div[2]/a').text
    rename_file(inn)
    driver.close()

    url = 'postgresql://{}:{}@{}:5432/{}'.format('test1', '12341234', '195.133.146.22', 'test')
    # postgres://{user}:{password}@{hostname}:{port}/{database-name}
    con = sqlalchemy.create_engine(url, echo=True)

    from sqlalchemy import Column, Integer, String, create_engine
    from sqlalchemy.ext.declarative import declarative_base

    Base = declarative_base()
    from sqlalchemy.sql import select
    class User(Base):
        __tablename__ = 'company'
        id = Column(Integer, primary_key=True)
        inn = Column(String)
        full_title = Column(String)
        name = Column(String)
        kind_activity = Column(String)
        address = Column(String)

        def __init__(self, inn, full_title, name, kind_activity, address):
            self.inn = inn
            self.full_title = full_title
            self.name = name
            self.kind_activity = kind_activity
            self.address = address
        def __repr__(self):
            return "<company('%s', %s, '%s', %s, %s)>" % (self.inn, self.full_title, self.name, self.kind_activity, self.address)

    Base.metadata.create_all(con)
    from sqlalchemy.orm import sessionmaker
    Session = sessionmaker(bind=con)
    session = Session()

    interer_User = User(inn, full_title, name, kind_activity, address)
    session.add(interer_User)
    session.commit()

download_pdf()