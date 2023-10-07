from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from selenium.common.exceptions import NoSuchElementException


class Milena():
    def __init__(self):
        self.driver = webdriver.Chrome()
        self.site_link = 'https://www.millenamoveiseeletro.com.br/eletrodomesticos'

    def main(self):
        self.abre()
        self.raspa()
        sleep(131231231)

    def abre(self):
        self.driver.get(self.site_link)
        self.driver.maximize_window()

    def raspa(self):
        site_map = {
            'XP':{
                'mais': '#product-search-results > div:nth-child(2) > div.col-sm-12.col-md-9.p > div.row.product-grid.product-tile-plp-space > div.col-12.p-0.grid-footer > div > button'
            }          
        }
        while True:
            try:
                mais = self.driver.find_element(By.CSS_SELECTOR, site_map['XP']['mais'])
                mais.click()
                sleep(3)

            except:
                print('nao tem mais resultados')
        



milena = Milena()
milena.main()
