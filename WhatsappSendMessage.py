from selenium import webdriver
from selenium.webdriver.common.by import By
import urllib
import os
import pandas as pd
import time


class WhatsappSendMessage:
    """My class to send messages to Whatsapp"""

    def __init__(self):
        """This is the constructor of the class WhatsappSendMessage"""
        self.driver = webdriver.Chrome()
        self.driver.get("https://web.whatsapp.com")
        self.driver.maximize_window()
        self.driver.implicitly_wait(10)

    def import_table(self):
        """This method import the table with the contacts"""
        self.tabela = pd.read_excel("files\contatos.xlsx")
        print(self.tabela)

        for linha in self.tabela.index:
            """Locate the contact in the table"""
            self.nome = self.tabela.loc[linha, "nome"]
            self.mensagem = self.tabela.loc[linha, "mensagem"]
            self.arquivo = self.tabela.loc[linha, "arquivo"]
            self.telefone = self.tabela.loc[linha, "telefone"]

    def send_mensage(self):
        """This method send the message to the contact"""
        texto = self.mensagem.replace("fulano", self.nome)
        texto = urllib.parse.quote(texto)
        print(texto)

        link = f"https://web.whatsapp.com/send?phone={self.telefone}&text={texto}"
        self.driver.get(link)

        """Wait for the page to load, if the page is not loaded and wait for 1 second, the program will not continue"""
        while len(self.driver.find_elements(By.ID, 'pane-side')) < 1:
            time.sleep(1)
        time.sleep(2)

        """Verify if the number is invalid, if it is invalid, locate the button to send the message"""
        if len(self.driver.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
            print("Número inválido")

            self.driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()

            if self.arquivo != "N":
                caminho_arquivo = os.path.abspath(f'files\{self.arquivo}')

                clip = self.driver.find_element(By.XPATH , '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span')
                clip.click()

                attach = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/span')
                attach.click()

                input = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input')
                input.send_keys(caminho_arquivo)
                time.sleep(2)

                send_file = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
                send_file.click()

            time.sleep(2)

if __name__ == "__main__":
    main = WhatsappSendMessage()
    main.import_table()
    main.send_mensage()

