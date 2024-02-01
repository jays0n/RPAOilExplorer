#Imports:
from selenium import webdriver
from selenium.common.exceptions import JavascriptException
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.firefox.options import Options as MozilaOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from msedge.selenium_tools import EdgeOptions as EdgeOptions2
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup as bs
from html import *
from time import sleep
import pandas as pd
import matplotlib.pyplot as plt
from datetime import *
import os


#Constantes:
url_site=r"https://oilxplorer.lexsis.com.br/"
user_login="loginengenharia"
password_login="login123"
File_Exit='oilexplorer_geral.xlsx'


# Auxiliary functions:
def Timer(seconds):
  sleep(seconds)

def Go(driver_object):
  try:
    driver_object.get(url_site)
  except Exception as err:
    print("Erro no browser ao tentar carregar a url:"+url_site + "\nVerifique sua conexao.")

def Login(driver_object):
  Timer(3)
  analyse=Wait_Until(driver_object,By.ID,'usuario')
  scripts=["document.getElementsByClassName('swal2-actions')[0].getElementsByTagName('button')[0].click()","document.getElementById('usuario').value='"+user_login+"'",
          "document.getElementById('senha').value='"+password_login+"'","document.getElementById('btn-enviar').click()"]
  for script in scripts:
    try:
      driver_object.execute_script(script)
      Timer(1)
    except Exception as err:
      pass

def Quit(driver_object):
  try:
    driver_object.Close()
  except:
    pass

def Wait_Until(driver_object,by,tag,expected_condition=None,seconds=1,max_attempts=10):
  attempts=max_attempts
  expected_condition_try=EC.presence_of_element_located if expected_condition is None else expected_condition
  while attempts>0:
      try:
          WebDriverWait(driver_object,seconds).until(expected_condition_try((by,tag)))
          return True
      except Exception as err:
          pass
      Timer(1)
      attempts-=1
  print("Erro no procedimento 'Wait' apos {0} tentativas de encontrar '{1}' sem sucesso!".format(max_attempts,tag))
  return False

def Close_Alerts(driver_object):
  k=0
  msg_err=None
  while k<10:
    try:
      driver_object.switch_to.alert().dismiss()
      driver_object.switch_to.default_content()
      break
    except Exception as err:
      msg_err=str(err)
      pass
    k+=1
  if msg_err!=None:
    print("Erro ao fechar o 'Alert':\n",msg_err)

#Chrome Initialize:
def Get_Driver():

    try:    
        #Options:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')

        #Driver:
        service = Service(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(options=chrome_options,service=service)
        #driver = webdriver.Edge(EdgeChromiumDriverManager().install())
    except Exception as err:
        print("O seguinte erro ocorreu:\n",err)
        edge_options=EdgeOptions2()
        edge_options.use_chromium = True
        #edge_options.add_argument('')
        edge_options.add_argument("disable-gpu")
        driver = webdriver.Edge(executable_path=EdgeChromiumDriverManager().install())
        
  
    driver.implicitly_wait(10)
    return driver

#Overview:
driver=Get_Driver()
Go(driver)
Login(driver)

soup=bs(driver.page_source,"html.parser")
class_name='card-label'
divs=soup.find('div',{'class':class_name})
divs=divs.find_all('div')[1]
total_anual=int(divs.text.split(' ')[0])

class_name='mt-10 mb-5'
divs=soup.find('div',{'class':class_name})
div_row=divs.find_all('div')

columns=[]
values=[]
percents=[]
columns.append('Total')
values.append(total_anual)
percents.append(1)

for div in div_row:
  class_name='col'
  div_col=div.find_all('div',{'class':class_name})
  for col in div_col:
      columns.append(col.text.strip().split('\n')[1])
      values.append(round(float(col.text.strip().split('\n')[0]),1))
      percents.append(round(100*float(col.text.strip().split('\n')[0])/total_anual,3))

driver.quit()

df_overview=pd.DataFrame({'Status':columns,'Quantidade':values,'Percents':percents})
print(df_overview,'\n\n')




"""
# History:

Go(driver)
Login(driver)
Wait_Until(driver, By.ID,'kt_aside')
script="selecionar_departamento('laudo-index')"
driver.execute_script(script)
Wait_Until(driver, By.ID,'btn-esconde-mostra-filtros')

attempts=0
while attempts<6:
  attempts+=1
  driver.execute_script("limpar_filtros()")
  Timer(3)
  driver.execute_script("filtrar()")
  times=0
  while times<120:
    trs=driver.find_element(By.ID,'tbl-laudos').find_elements(By.TAG_NAME,'tr')
    if len(trs)>3:
        total_rows=len(trs)
        attempts=100
        break
    else:
        Timer(1)
    times+=1

soup=bs(driver.page_source,"html.parser")
tbl=soup.find('div',{'id':'tbl-laudos'})
df_report=pd.read_html(str(tbl))[0]
Timer(2)

#Sub extracts:
start_sample=False
start_teste=False
first_time=True
first_occurrence=True

last_record_sample=None
last_record_teste=None

for row in range(total_rows):
    webElement=driver.find_element(By.ID,'tbl-laudos').find_elements(By.TAG_NAME,'table')[0]
    webElement=webElement.find_elements(By.TAG_NAME,'tr')[row]
    if webElement is None or webElement==[]:
        print("Na tabela lauldos: O elemento da linha {0} não existe ou não pode ser acessado!".format(row))
    else:
        if len(webElement.find_elements(By.TAG_NAME,'td'))>0:
            vessel=str(webElement.find_elements(By.TAG_NAME,'td')[2].text).strip()
            if vessel in (""," ","\n"," \n"):
                script="return document.getElementsByClassName('datatable-table')[0].getElementsByTagName('tr')[{0}].".format(row)
                script+="getElementsByTagName('td')[2].innerText"
                vessel=driver.execute_script(script)

            equipment_number=str(webElement.find_elements(By.TAG_NAME,'td')[7].text).strip()
            if equipment_number in (""," ","\n"," \n"):
                script="return document.getElementsByClassName('datatable-table')[0].getElementsByTagName('tr')[{0}].".format(row)
                script+="getElementsByTagName('td')[7].innerText"
                equipment_number=driver.execute_script(script)

            equipment_name=str(webElement.find_elements(By.TAG_NAME,'td')[8].text).strip()
            if equipment_name in (""," ","\n"," \n"):
                script="return document.getElementsByClassName('datatable-table')[0].getElementsByTagName('tr')[{0}].".format(row)
                script+="getElementsByTagName('td')[8].innerText"
                equipment_name=driver.execute_script(script)


            date_end_sample=str(webElement.find_elements(By.TAG_NAME,'td')[5].text).strip()
            if date_end_sample in (""," ","\n"," \n"):
                script="return document.getElementsByClassName('datatable-table')[0].getElementsByTagName('tr')[{0}].".format(row)
                script+="getElementsByTagName('td')[5].innerText"
                date_end_sample=driver.execute_script(script)

            if last_record_sample == None:
                start_sample=True
            else:
                if datetime.strptime(str(last_record_sample['data'])[0:10],'%d/%m%Y') < datetime.strptime(str(date_end_sample)[0:10],'%d/%m%Y') \
                    and last_record_sample['equipamento']==equipment_name and  last_record_sample['navio']==vessel:
                    start_sample=True

            if start_sample:
                try:
                    webElement.click()
                except:
                    script="document.getElementsByClassName('datatable-table')[0].getElementsByTagName('tr')[{0}].getElementsByTagName('td')[6].click()"
                    script=script.format(row)
                    driver.execute_script(script)

                times=0
                while times<120:
                    trs=driver.find_element(By.ID,'tbl-amostras').find_elements(By.TAG_NAME,'tr')
                    if len(trs)>1:
                        total_samples=len(trs)
                        break
                    else:
                        Timer(1)
                    times+=1

                if times>=120: #Estouro?
                    pass
                else:
                    soup=bs(driver.page_source,"html.parser")
                    tbl=soup.find('div',{'id':'tbl-amostras'})

                    df_sample_temp=None
                    df_sample_temp=pd.read_html(str(tbl))[0]

                    total_temp=len(df_sample_temp['SEQ'])

                    df_sample_temp['Embarcacao']=[vessel]*total_temp
                    df_sample_temp['Numero Equipamento']=[equipment_number]*total_temp
                    df_sample_temp['Equipamento']=[equipment_name]*total_temp

                    if first_time:
                        df_sample=df_sample_temp
                        first_time=False
                    else:
                        df_sample=pd.concat([df_sample,df_sample_temp],ignore_index=True)

                    for sample in range(total_samples):
                        sample_element=driver.find_element(By.ID,'tbl-amostras').find_elements(By.TAG_NAME,'tr')[sample]
                        if sample_element is None:
                            print("Na tabela amostras do laudo {0}: O elemento da linha {1} não existe ou não pode ser acessado!".format(row,sample))
                        else:
                            if len(sample_element.find_elements(By.TAG_NAME,'td'))>0:

                                seq=sample_element.find_elements(By.TAG_NAME,'td')[1].text.strip()
                                if seq in (""," ","\n"," \n"):
                                    script="return document.getElementsByClassName('datatable-table')[1].getElementsByTagName('tr')[{0}].".format(sample)
                                    script+="getElementsByTagName('td')[1].innerText"
                                    seq=driver.execute_script(script)

                                RA=sample_element.find_elements(By.TAG_NAME,'td')[2].text.strip()
                                if RA in (""," ","\n"," \n"):
                                    script="return document.getElementsByClassName('datatable-table')[1].getElementsByTagName('tr')[{0}].".format(sample)
                                    script+="getElementsByTagName('td')[2].innerText"
                                    RA=driver.execute_script(script)

                                date_teste=sample_element.find_elements(By.TAG_NAME,'td')[8].text.strip()
                                if date_teste in (""," ","\n"," \n"):
                                    script="return document.getElementsByClassName('datatable-table')[1].getElementsByTagName('tr')[{0}].".format(sample)
                                    script+="getElementsByTagName('td')[8].innerText"
                                    date_teste=driver.execute_script(script)

                                if last_record_teste == None:
                                    start_teste=True
                                else:
                                    if str(last_record_teste['RA'])==str(RA) \
                                        and datetime.strptime(str(last_record_teste['data'])[0:10],'%d/%m/%Y') < datetime.strptime(str(date_teste)[0:10],'%d/%m/%Y'):
                                        start_teste=True

                                if start_teste:
                                    Timer(1)
                                    sample_element=driver.find_element(By.ID,'tbl-amostras').find_elements(By.TAG_NAME,'tr')[sample]
                                    try:
                                        sample_element.click()
                                    except:
                                        script="document.getElementsByClassName('datatable-table')[1].getElementsByTagName('tr')[{0}].getElementsByTagName('td')[6].click()"
                                        script=script.format(sample)
                                        driver.execute_script(script)
                                    Timer(2)
                                    times=0
                                    while times<120:
                                        trs=driver.find_element(By.ID,'tbl-teste').find_elements(By.TAG_NAME,'tr')
                                        if len(trs)>1:
                                            break
                                        else:
                                            Timer(1)
                                        times+=1

                                    soup=bs(driver.page_source,"html.parser")
                                    tbl=soup.find('table',{'id':'tbl-teste'})

                                    if times>=120:
                                        pass
                                    else:
                                        df_testes_temp=pd.read_html(str(tbl))[0]
                                        total_temp=len(df_testes_temp['Tipo'])
                                        df_testes_temp['Embarcacao']=[vessel]*total_temp
                                        df_testes_temp['Numero Equipamento']=[equipment_number]*total_temp
                                        df_testes_temp['Equipamento']=[equipment_name]*total_temp
                                        df_testes_temp['RA']=[RA]*total_temp
                                        df_testes_temp['Seq']=[seq]*total_temp

                                        if first_occurrence:
                                            first_occurrence=False
                                            df_testes=df_testes_temp
                                        else:
                                            df_testes=pd.concat([df_testes,df_testes_temp],ignore_index=True)

                                        Timer(1)
                                        driver.execute_script("document.getElementById('div-teste-fechar').click()")


try:
  df_report.to_csv()
  df_sample.to_csv()
  df_testes.to_csv()
except Exception as err:
  print(err)


"""