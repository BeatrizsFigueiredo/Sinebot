from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import re
import pandas as pd
from datetime import datetime
from obs_comum import processo_secundario

#Carrega os dados do excel
alimentação = pd.read_excel('cartao comum.xlsx', sheet_name='Resultado da consulta', engine='openpyxl')

#cria coluna nova vazia
alimentação['Cartão Criado'] = ''
#abre navegador
navegador = webdriver.Firefox()
#declara variavel para aguardar um elemento
def wait_for_element(xpath, timeout=17):
    return WebDriverWait(navegador, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
def xpath_inserir():
    wait_for_element('//*[@id="btnInsert"]').click()
    
#bloco de retorno de erros
try:
   #pagina inicial do VTADMIN
   navegador.get("https://vtadmin.manaus.prodatamobility.com.br/wfm_Home.aspx")

   #informações usuarios
   wait_for_element('//*[@id="txtLogin"]').send_keys("Comercial.sntr")
   wait_for_element('//*[@id="txtSenha"]').send_keys("Botnelly")
   wait_for_element('//*[@id="loginbutton"]').click()
   #clica para pesquisar o UCA
   wait_for_element('//*[@id="parent_uca"]/a').click()
   wait_for_element('//*[@id="xmlmenu_right"]/li[1]/ul/li[1]/a').click()

   #necessita vim para essa pagina para conseguir inserir
   navegador.get('https://vtadmin.manaus.prodatamobility.com.br//pages/uca/wfm_Users_Ins.aspx?NEW=1')

   #looping cadastros
   for index, row in alimentação.iterrows():
        #bloco de erros
         try:
           #declara a variavel para nome
           usuarios = row['Nome']
           #variavel para o campo / aguarda o elemento aparecer
           inserir_nome = wait_for_element('//*[@id="txtNome"]')
           #limpa o campo antes de colocar o nome
           inserir_nome.clear()
           #define a variavel como texto
           inserir_nome.send_keys(str(usuarios))
          
           #declara a variavel para data
           data_nascimento = row ['Data_Nascimento']

           #variavel para o campo / aguarda o elemento aparecer
           inserir_data = wait_for_element('//*[@id="txtDataNascimento"]')
           #limpa o elemento antes de inserir
           inserir_data.clear()
           #define a variavel como texto para inserir
           inserir_data.send_keys(str(data_nascimento))
           
           #necessita entrar no site documentos/cpf para inserir
           navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Documents_Ins.aspx')

           #declara a variavel da coluna cpf
           #usuarios_cpf = row ['CPF']
           #declara variavel como texto 
           usuarios_cpf = re.sub(r'\D', '', str(row['CPF'])).zfill(11)
           inserir_cpf = wait_for_element('//*[@id="txtDoc"]')
           inserir_cpf.clear()
           inserir_cpf.send_keys(f"'{usuarios_cpf}")  # Envia o CPF
          
           #elemento para incluir
           wait_for_element('//*[@id="btnIncluir"]').click()

           #necessita voltar duas vezes para conseguir prosseguir o cadastro
           navegador.back()
           navegador.back()

           #site para o telefone
           navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Telephones_Ins.aspx')

           #declara a variavel de DDD
           usuarios_ddd = row['DDD']

           #declara variavel para o campo / aguarda o elemento aparecer
           inserir_ddd = wait_for_element('//*[@id="txtArea"]')
           #limpa o elemento antes de inserir
           inserir_ddd.clear()
           #define a variavel como texto para inserir
           inserir_ddd.send_keys(str(usuarios_ddd))

           #declara a variavel de telefone
           numero_telefone = row ['Telefone_contato']
           #declara variavel para o campo / aguarda o campo aparecer
           inserir_numero = wait_for_element('//*[@id="txtPhone"]')
           #limpa o elemento antes de inserir
           inserir_numero.clear()
           #define a variavel como texto para inserir
           inserir_numero.send_keys(str(numero_telefone))
           #botao de inserir
           xpath_inserir()

           #volta para o navegador anterior
           navegador.back()
           #site para incluir o email
           navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Emails_Ins.aspx')

           #declara a variavel para o email
           usuarios_email = row ['Email']
           #declara variavel para o campo / aguarda o campo aparecer
           inserir_email = wait_for_element('//*[@id="txtEmail"]')
           #limpa o elemento antes de inserir
           inserir_email.clear()
           #define a variavel como texto para inserir
           inserir_email.send_keys(str(usuarios_email))
           #botao de inserir
           xpath_inserir()
           #necessita voltar 3 vezes para conseguir imprimir\nao remover 
           navegador.back()
           navegador.back()
           navegador.back()
           
           # Cria um objeto selecionavel / Seleciona o dropdown para o tipo de usuário e aguarda o elemento aparecer
           dropdown_comum = Select(wait_for_element('//*[@id="cboTypeUser"]'))
           #Seleciona a opção no dropdown de tipo de usuário
           dropdown_comum.select_by_visible_text("COMUM")
           #aguarda o elemento aparecer e clica para inserir
           xpath_inserir()
        
           #Cria um objeto selecionavel / Seleciona o dropdown para o tipo de usuário e aguarda o elemento aparecer
           dropdown_compra = Select(wait_for_element('//*[@id="cboTypeUser"]'))
           #Seleciona a opção no dropdown de tipo de usuário
           dropdown_compra.select_by_visible_text("COMPRA RECARGA")

           #botao inserir apenas aplicações 
           xpath_inserir()

           #botao para ADICIONAR cadastro
           wait_for_element('//*[@id="btnAdd"]').click()
           time.sleep(2)
           #bloco de erro caso ja esta algum cadastro
           try:
                wait_for_element('//*[@id="lblAlertMessage"]', timeout=2)  
                alimentação.at[index, 'Cartão Criado'] = "JÁ POSSUI CADASTRO"
           #salva na coluna nova as informações
           except TimeoutException:
                alimentação.at[index, 'Cartão Criado'] = "FEITO"

           #volta para a pagina de inserir para recomeçar o looping
           navegador.get('https://vtadmin.manaus.prodatamobility.com.br//pages/uca/wfm_Users_Ins.aspx?NEW=1')

         #final da captura o erro do codigo
         except Exception as e:
          print(f"Erro ao processar o CPF {usuarios_cpf}: {e}")
          alimentação.at[index, 'Criados'] = "ERRO"

         with pd.ExcelWriter('cartao comum.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
          alimentação.to_excel(writer, sheet_name='Resultado da consulta', index=False)  

finally:
    processo_secundario(navegador)
    navegador.quit()