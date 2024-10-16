from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.select import Select
import openpyxl

numero_oab = '259155'
estado_uf = 'UF'
planilha_dados_consulta = openpyxl.load_workbook('dados da consulta.xlsx')
pagina_processos = planilha_dados_consulta['Plan1']

#Entrar no Site

site = 'https://pje-consulta-publica.tjmg.jus.br/'

driver = webdriver.Chrome()

driver.get(site)
sleep(2)

#Digitar Algo no Site

campo_numero_oab = driver.find_element(By.XPATH,"//input[@id='fPP:Decoration:numeroOAB']") # F12 para encontrar onde sera buscado o local onde ira digitar
sleep(2)
campo_numero_oab.click()
sleep(1)
campo_numero_oab.send_keys(numero_oab)
sleep(1)

#Escolher o Estado

selecao_UF = driver.find_element(By.XPATH,"//select[@id='fPP:Decoration:estadoComboOAB']")
sleep(1)
opcoes_uf = Select(selecao_UF)
sleep(1)
opcoes_uf.select_by_visible_text(estado_uf)  # Faz a selecao de qual estado quer

# Clicar na parte de Pesquisar

botao_pesquisar = driver.find_element(By.XPATH,"//input[@id='fPP:searchProcessos']")
sleep(1)
botao_pesquisar.click()

# Entrar nos processos
sleep(5)
links_abrir_processo = driver.find_elements(By.XPATH,"//a[@title='Ver Detalhes']")



for link in links_abrir_processo:
    janela_principal = driver.current_window_handle
    link.click()
    sleep(2)
    janelas_abertas = driver.window_handles
    for janela in janelas_abertas:
        if janela not in janela_principal:
            driver.switch_to.window(janela)
            sleep(2)

            # Salvando nome dos participantes, numero do processo

            numero_processo = driver.find_elements(By.XPATH,"//div[@class='propertyView ']//div[@class='col-sm-12 ']")[0]
            
            participantes = driver.find_elements(By.XPATH,"//tbody[contains(@id,'processoPartesPoloAtivoResumidoList:tb')]//span[@class='text-bold']")
            
            lista_participantes = []

            for participantes in participantes:
                lista_participantes.append(participantes.text)

                # Guardar os nomes dos participantes
                if len(lista_participantes) == 1:
                    # n da oab, n do processo, nome dos participantes
                    pagina_processos.append([numero_oab,numero_processo.text,lista_participantes[0]])

                else:
                    # ','.join(lista_participantes) para sair os nomes dos participantes entre ,
                    pagina_processos.append([numero_oab,numero_processo.text,','.join(lista_participantes)])

                planilha_dados_consulta.save('dados da consulta.xlsx')

                driver.close()

        driver.switch_to.window(janela_principal)    
            