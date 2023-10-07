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
        sleep(2)
        self.mais()
        sleep(2)
        self.raspa()
        sleep(131231231)

    def abre(self):
        self.driver.get(self.site_link)
        self.driver.maximize_window()
        sleep(2)


    def mais(self):
        passa = {
            'XP':{
                'mais': '#product-search-results > div:nth-child(2) > div.col-sm-12.col-md-9.p > div.row.product-grid.product-tile-plp-space > div.col-12.p-0.grid-footer > div > button'
            }          
        }
        while True:
            try:
                mais = self.driver.find_element(By.CSS_SELECTOR, passa['XP']['mais'])
                mais.click()
                sleep(1)

            except NoSuchElementException:
                try:
                    break
                except:
                    print('algo deu errado')

            except Exception as e:
                print(f'error {e}')

    def raspa(self):
        contador = 1
        while True:
            try:
                produtos = {
                    'XP':{
                        'nome': f'/html/body/div[2]/div/div[2]/div/div/div[1]/div[2]/div[2]/div[2]/div[{contador}]/div/div/div[2]/div[1]/a',        
                        'preco': f'/html/body/div[2]/div/div[2]/div/div/div[1]/div[2]/div[2]/div[2]/div[{contador}]/div/div/div[2]/div[2]/span'
                    }
                }

                nome = self.driver.find_element(By.XPATH, produtos['XP']['nome']).text
                preco = self.driver.find_element(By.XPATH, produtos['XP']['preco']).text

                print(nome)
                print(preco)
                sleep(1)
                contador += 1
                print(contador)
            except NoSuchElementException:
                print('sexo')

            except Exception as e:
                print(f'error {e}')

    

milena = Milena()
milena.main()
