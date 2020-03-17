from selenium import webdriver
from selenium.webdriver.common.proxy import Proxy, ProxyType
import requests
import datetime
import io
#from DB import Add, End
import re
import os
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
from bs4 import BeautifulSoup
import json
from sqlalchemy import desc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import configparser
from docxtpl import DocxTemplate
encoding ='UTF-8'
#prox = Proxy()
#prox.proxy_type = ProxyType.MANUAL
#prox.http_proxy = "1.0.0.40:80"
#prox.ssl_proxy = "163.172.147.94:8811"

capabilities = webdriver.DesiredCapabilities.FIREFOX
#prox.add_to_capabilities(capabilities)


def extract_text_from_pdf():
    file = os.listdir('D:\\Data\\send\\')
    files_name = ''
    for item in file:
        if re.search(str(), item):
            files_name ='D:\\Data\\send\\'+ item
            break
    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager , fake_file_handle)
    page_interpreter = PDFPageInterpreter(resource_manager , converter)
    count=0
    with open(files_name, 'rb') as fh:
        for page in PDFPage.get_pages(fh,
                                      caching=True,
                                      check_extractable=True):
            page_interpreter.process_page(page)
            count+=1
            if count==3:
                break

        text = fake_file_handle.getvalue()
    converter.close()
    fake_file_handle.close()

#def patterns1():
    p1 = re.compile(r'([а-я])([А-Я])')
    p2 = re.compile(r'([А-Я]{7})([а-я])')
    p3 = re.compile(r'([A-я])([0-9])')
    p4 = re.compile(r'([0-9])([A-я])')
    p5 = re.compile(r'(\w)(№)')
    p6 = re.compile(r'([.!?)])([A-я])')
    p7 = re.compile(r'(\d{2}.\d{2}.\d{4})([A-я])')
    p8 = re.compile(r'([A-я])(\d{2}.\d{2}.\d{4})')
    p9 = re.compile(r'([,])([A-я])')
    p10 = re.compile(r'([А-я]")([A-я])')
#def no_CaMeLs():
    text2 = re.sub(p1, r"\1&\2", text)
    text3 = re.sub(p2, r"\1 \2", text2)
    text4 = re.sub(p3, r"\1&\2", text3)
    text5 = re.sub(p4, r"\1&\2", text4)
    text6 = re.sub(p5, r"\1 \2", text5)
    text7 = re.sub(p6, r"\1&\2", text6)
    text8 = re.sub(p7, r"\1 \2", text7)
    text9 = re.sub(p8, r"\1 \2", text8)
    text10 = re.sub(p9, r"\1 \2", text9)
    text11 = re.sub(p10, r'\1&\2', text10)
#def patterns2():
    p_vipiska_date = re.compile(r'(.+)(ИСКА из Единого государственного реестра юридических лиц&)(\d{2}.\d{2}.\d{4})(.+)')
    p_number_date = re.compile(r'(.+)(№.+)(&дата формирования выписки)(.+)')
    p_full_title = re.compile(r'(.+)(Настоящая выписка содержит сведения о юридическом лице&)([А-я\s"]+)(.+)')
    p_inn = re.compile(r'(.+)(Сведения об учете в налоговом органе&\d\d&ИНН&)([0-9]+)(\d\d&КПП&)(.+)')
    p_ogrn = re.compile(r'(.+)(полное наименование юридического лица&ОГРН&)([0-9]+)(.+)')
    p_name = re.compile(r'(.+)(Сокращенное наименование&)([А-я\s"]+)(.+)')
    p_date_egrul = re.compile(r'(.+)(3&ГРН и дата внесения в ЕГРЮЛ записи, содержащей указанные сведения&)([0-9\.]+)(\d{2}.\d{2}.\d{4})(&Адрес )(.+)')
    p_index = re.compile(r'(.+)(4&Почтовый индекс&)([0-9]{6})(5&Субъект Российской Федерации&)(.+)')
    p_subject_RF = re.compile(r'(.+)(5&Субъект Российской Федерации&)([А-я\s]+)(&6&Улица )(.+)')
    p_street = re.compile(r'(.+)(проспект, переулок и т\.&д\..&)(.+)(&7&Дом )(.+)')
    p_house = re.compile(r'(.+)(&ДОМ )([0-9]+)(&Корпус )(.+)')
    p_corpus = re.compile(r'(.+)(&СТРОЕНИЕ )([0-9]+)(&Офис )(.+)')
    p_flat = re.compile(r'(.+)(&)(.+)(10&ГРН и дата внесения в ЕГРЮЛ)(.+)')
    p_date_egrul2 = re.compile(r'(.+)(10&ГРН и дата внесения в ЕГРЮЛ записи, содержащей указанные сведения&)([0-9\.]+)(\d{2}.\d{2}.\d{4})(&Сведения о регистрации&11&)(.+)')
    p_registration = re.compile(r'(.+)(11&Способ образования&)([А-я\s]+)(&12&ОГРН)(.+)')
    p_date_registration = re.compile(r'(.+)(13&Дата регистрации&)(\d{2}.\d{2}.\d{4})(14&ГРН и дата внесения в ЕГРЮЛ записи, содержащей указанные сведения&)(.+)')
    p_date_egrul3 = re.compile(r'(.+)(14&ГРН и дата внесения в ЕГРЮЛ записи, содержащей указанные сведения&)([0-9\.]+)(\d{2}.\d{2}.\d{4})(&Сведения о регистрирующем органе по месту нахождения юридического лица&15)(.+)')
    p_kpp = re.compile(r'(.+)(&КПП&)([0-9]+)(&Дата постановки на учет&)(.+)')
    p_date_inn = re.compile(r'(.+)(&Дата постановки на учет&)(\d{2}.\d{2}.\d{4})(21&Наименование налогового органа&)(.+)')
    p_tax_office = re.compile(r'(.+)(&Наименование налогового органа&)(.+)(22&ГРН и дата внесения в ЕГРЮЛ записи)(.+)')
    p_capital = re.compile(r'(.+)(31&Вид&)(.+)(&32&Размер )(.+)')
    p_capital2 = re.compile(r'(.+)(&32&Размер .в рублях.)(.+)(33&ГРН и дата внесения в ЕГРЮЛ)(.+)')
    p_surname = re.compile(r'(.+)(35&Фамилия&)(.+)(&36&Имя&)(.+)')
    p_name_gener = re.compile(r'(.+)(&36&Имя&)(.+)(&37&Отчество&)(.+)')
    p_patronymic = re.compile(r'(.+)(&37&Отчество&)(.+)(&38&ИНН&)(.+)')
    p_inn_gener = re.compile(r'(.+)(&38&ИНН&)(.+)(39&ГРН и дата внесения в ЕГРЮЛ записи,)(.+)')
    p_position = re.compile(r'(.+)(40&Должность&)(.+)(&41&ГРН и дата внесения)(.+)')
    p_founder = re.compile(r'(.+)(43&Полное наименование&)(.+)(&44&ГРН и дата внесения)(.+)')
    p_fou_country = re.compile(r'(.+)(45&Страна происхождения&)(.+)(&46&Дата регистрации&)(.+)')
    p_fou_address = re.compile(r'(.+)(&49&Адрес .место нахождения. в странепроисхождения&)(.+)(&50&ГРН и дата внесения )(.+)')
    p_fou_capital = re.compile(r'(.+)(51&Номинальная стоимость доли .в рублях.)(.+)(52&Размер доли)(.+)')
    p_percent = re.compile(r'(.+)(52&Размер доли .в процентах.)(.+)(53&ГРН и дата внесения в ЕГРЮЛ записи)(.+)')
    p_activity = re.compile(r'(.+)(&54&Код и наименование вида деятельности&)(.+)(&55&ГРН и дата внесения в ЕГРЮЛ записи)(.+)')

#def stuff_docx():
    vipiska_date = re.sub(p_vipiska_date, r'\3', text11)
    vipiska_number = re.sub(p_number_date, r'\2', text11)
    full_title = re.sub(p_full_title, r'\3', text11)
    inn = re.sub(p_inn, r'\3', text11)
    ogrn = re.sub(p_ogrn, r'\3', text11)
    name = re.sub(p_name, r'\3', text11)
    date_egrul = re.sub(p_date_egrul, r'\3 \4', text11)
    index = re.sub(p_index, r'\3', text11)
    subject_RF = re.sub(p_subject_RF, r'\3', text11)
    street = re.sub(p_street, r'\3', text11)
    house = re.sub(p_house, r'\3', text11)
    corpus = re.sub(p_corpus, r'\3', text11)
    flat = re.sub(p_flat, r'\3', text11)
    date_egrul2 = re.sub(p_date_egrul2, r'\3 \4', text11)
    registration = re.sub(p_registration, r'\3', text11)
    date_registration = re.sub(p_date_registration, r'\3', text11)
    date_egrul3 = re.sub(p_date_egrul3, r'\3 \4', text11)
    kpp = re.sub(p_kpp, r'\3', text11)
    date_inn = re.sub(p_date_inn, r'\3', text11)
    tax_office = re.sub(p_tax_office, r'\3', text11)
    capital = re.sub(p_capital, r'\3', text11)
    capital2 = re.sub(p_capital2, r'\3', text11)
    surname = re.sub(p_surname, r'\3', text11)
    name_gener = re.sub(p_name_gener, r'\3', text11)
    patronymic = re.sub(p_patronymic, r'\3', text11)
    inn_gener = re.sub(p_inn_gener, r'\3', text11)
    position = re.sub(p_position, r'\3', text11)
    founder = re.sub(p_founder, r'\3', text11)
    fou_country = re.sub(p_fou_country, r'\3', text11)
    fou_address = re.sub(p_fou_address, r'\3', text11)
    fou_capital = re.sub(p_fou_capital, r'\3', text11)
    percent = re.sub(p_percent, r'\3', text11)
    activity = re.sub(p_activity, r'\3', text11)
    print(text11)

#def make_docx():
    doc = DocxTemplate("D:\\Data\\document\\выписка.docx")
    context = {'framing_date': vipiska_date,
               'vipiska_number': vipiska_number,
               'full_title': full_title,
               'inn': inn,
               'ogrn': ogrn,
               'name': name,
               'date_egrul': date_egrul,
               'index': index,
               'subject_RF': subject_RF,
               'street': street,
               'house': house,
               'corpus': corpus,
               'flat': flat,
               'date_egrul2': date_egrul2,
               'registration': registration,
               'date_registration': date_registration,
               'date_egrul3': date_egrul3,
               'kpp': kpp,
               'date_inn': date_inn,
               'tax_office': tax_office,
               'capital': capital,
               'capital2': capital2,
               'surname': surname,
               'name_gener': name_gener,
               'patronymic': patronymic,
               'inn_gener': inn_gener,
               'position': position,
               'founder': founder,
               'fou_country': fou_country,
               'fou_address': fou_address,
               'fou_capital': fou_capital,
               'percent': percent,
               'activity': activity}

    doc.render(context)
    doc.save('D:\\Data\\document\\выписка_' + str(inn) + '.docx')


extract_text_from_pdf()

