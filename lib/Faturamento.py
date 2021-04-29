from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys

import sys

from lib import AuxFunc
from lib import GUI
import pandas as pd
import time
import os
import wx
import psutil as ps

class fatura(object):
    def __init__(self):
        self.basePath = os.getcwd()
        #Opening chromedriver
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        self.browser = webdriver.Chrome(self.basePath +"/chromedriver.exe", chrome_options=options)

    #Login SSW
    def loginSSW(self,dom,cpfUser,nameUser,pswUser):
        self.browser.get('https://sistema.ssw.inf.br/bin/ssw0422') 
        element = AuxFunc.checkElementByID('5', self.browser)
        if(element):
            dominio = self.browser.find_element_by_id("1")
            cpf = self.browser.find_element_by_id("2")
            username = self.browser.find_element_by_id("3")
            password = self.browser.find_element_by_id("4")
            dominio.clear()
            cpf.clear()
            username.clear()
            password.clear()
            dominio.send_keys(dom)
            cpf.send_keys(cpfUser)
            username.send_keys(nameUser)
            password.send_keys(pswUser)
            element.click()
             
            #check login
            try:
                element = WebDriverWait(self.browser,3).until(EC.visibility_of_element_located((By.XPATH, "//*[text()='Menu Principal']")))
            except TimeoutException:
                element =  AuxFunc.checkElementByID('0', self.browser)
                element.click()
                return False

            return True

    #opening opc 147 in new tab       
    def openOPC(self):
        time.sleep(0.5)
        self.browser.execute_script("window.open('https://sistema.ssw.inf.br/bin/ssw1478/', '_blank')")
        self.browser.switch_to_window(self.browser.window_handles[1])
        time.sleep(0.5)
         
    
    def createFilesXLXS(self,pathFatura):
        #Open invoice file
        self.df = pd.read_excel(pathFatura,header=0,index_col=False,keep_default_na=True)

        #creating 'capas' names file
        try:
            self.dfcapas = pd.read_excel(self.basePath +"/relacao_capas.xlsx" ,header=0,index_col=False,keep_default_na=True)
        except FileNotFoundError:
            COLUMNS = ['UDI', 'PMI', 'PCB', 'PTC', 'MON', 'IBI', 'PTM']
            self.dfcapas = pd.DataFrame(columns=COLUMNS)
        
        #Opening/creating temporary files of subcontract names and capas
        self.fileTemp = open(self.basePath +"/tmp/TempFile.txt", 'a+')
        
    
    def removingSubcontracts(self, numCapa):

        inputKeyCapa = self.browser.find_element_by_id("1")
        bntNext = self.browser.find_element_by_id("2")
        inputKeyCapa.send_keys(numCapa)
        bntNext.click()
        logResult = ""
        element = AuxFunc.checkElementByID('4', self.browser)
        if(element):
            lenghtCapa = int(self.browser.find_element_by_xpath("//*[@id='frm']/div[19]").text)
            element.click()
            element= AuxFunc.checkElementByID(str(lenghtCapa+1), self.browser)
            if(element):
                noSearch = [""]
                content = AuxFunc.remnoveSubcontractDF(self.browser,self.df,lenghtCapa,noSearch)
                element.click()
                element = AuxFunc.checkElementByID('5', self.browser)
                if(element):
                    if not (content ==""):
                        if(content =="Subcontrato iguais encontrado na fatura"):
                            logResult = content
                        else:
                            self.fileTemp.writelines(content)
                            uni = self.browser.find_element_by_xpath("//*[@id='frm']/div[9]").text
                            capaNum = self.browser.find_element_by_xpath("//*[@id='frm']/div[10]").text
                            try:
                                self.dfcapas.at[self.dfcapas[uni].count(), uni] = float(capaNum.replace("-",""))
                                self.dfcapas.to_excel("Relacao_Capas.xlsx", index=False)
                            except:
                                logResult = "Não foi possivel incluir a capa na relação"
                            
                            if not(noSearch[0] ==""):
                                logResult = "Subcontratos não encontrados: "+ noSearch[0]
                    else:
                        logResult = "capa não esta na fatura"
                        
                    element.click()
                else:
                    logResult = "Ocorreu um Erro inesperado"
            else:
                logResult = "Ocorreu um Erro inesperado"                        
        else:
            logResult = "Numero de capa invalido"

        element = AuxFunc.checkElementByID('0', self.browser)
        if(element):
            element.click()

        return logResult
        
    def checkExtensionFile(self, pathFile):
        len_pathFile = len(pathFile)
        extensionFile = ""
        while pathFile[len_pathFile-1] != '.':
            extensionFile += pathFile[len_pathFile-1]
            len_pathFile -= 1
        if(extensionFile == "xslx"):
            return True
        else:
            return False

    def saveprocess(self, pathFatura):
        self.fileTemp.close()
        AuxFunc.setIndexDataFrame(self.basePath + "/tmp/TempFile.txt",self.df)
        self.df.to_excel(pathFatura, index=False) 
        #self.dfcapas.to_excel("Relacao_Capas.xlsx", index=False) 
        self.browser.quit()
        os.remove(self.basePath+"/tmp/TempFile.txt")

    def checkProcessExcel(self):
        PROC_NAME = "EXCEL.EXE"
        for proc in ps.process_iter():
            if proc.name() == PROC_NAME:
                return True         
        return False