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
        sleep(131231231)

    def abre(self):
        self.driver.get(self.site_link)
        self.driver.maximize_window()
    



milena = Milena()
milena.main()
