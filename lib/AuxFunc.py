from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time


def checkElementByID(id, browser):
    time.sleep(0.5)
    try:
        element = WebDriverWait(browser,2).until(EC.visibility_of_element_located((By.ID, id)))
    except TimeoutException:
        return False
    return element


def setIndexDataFrame(pathfile, df):
    file = open(pathfile, 'r')
    contents = file.readline()
    contents = contents.split('-')
    lenContents = len(contents) -1
    for i in range(lenContents):
        try:
            df.drop(int(contents[i]) , inplace=True)
        except:
            print("index no exist")
    file.close()
        
def remnoveSubcontractDF(browser, df, lenghtCapa, noSearch):
    content = ""
    for x in range(lenghtCapa):
        line = str(2 + x)
        subcontract = browser.find_element_by_xpath("//*[@id='tblsr']/tbody/tr["+line+"]/td[1]/div/a/u").text
        indexNames = df[ df['NUMERO CTRC'] == subcontract ].index
        if not (indexNames.empty):
            try:
                content += str(int(indexNames.values)) + '-'
            except:
                content ="Subcontratos iguais encontrado na fatura"
        else:
            noSearch[0] += '\n'+subcontract + ' '
    return content

