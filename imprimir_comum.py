from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import re
import time

# Carregar os dados do Excel
criar_cadastro = pd.read_excel('cartao comum.xlsx', sheet_name='Resultado da consulta', engine='openpyxl')
#cria uma coluna nova vazia 
criar_cadastro['Criados'] = ''

# Inicia o navegador
navegador = webdriver.Firefox()
#declara variavel para aguardar um elemento
def wait_for_element(xpath):
    return WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
#bloco de retorno de erros
try:
    # Acessar a página de login
    navegador.get("https://vtadmin.manaus.prodatamobility.com.br/wfm_Home.aspx")

    # Login usuario
    wait_for_element('//*[@id="txtLogin"]').send_keys("Comercial.sntr")
    wait_for_element('//*[@id="txtSenha"]').send_keys("Botnelly")
    wait_for_element('//*[@id="loginbutton"]').click()

    # Vai para a página do UCA
    wait_for_element('//*[@id="parent_uca"]/a').click()
    wait_for_element('//*[@id="xmlmenu_right"]/li[1]/ul/li[2]/a').click()

    #necessita vim para essa pagina para conseguir pesquisar
    navegador.get("https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx")

    #Seleciona o status "TODOS" no dropdown
    dropdown_status = Select(wait_for_element('//*[@id="cboStatus"]'))
    dropdown_status.select_by_visible_text("TODOS")

    #looping cadastros
    for index, row in criar_cadastro.iterrows():
        try:
            #declara variavel como texto 
            usuarios_cpf = re.sub(r'\D', '', str(row['CPF'])).zfill(11)
            inserir_cpf = wait_for_element('//*[@id="txtDoc"]')
            #limpa o elemento antes de inserir
            inserir_cpf.clear()
            #define a variavel como texto para inserir
            inserir_cpf.send_keys(f"'{usuarios_cpf}")  # Envia o CPF

            #elemento para incluir
            wait_for_element('//*[@id="btnEldery"]').click()
            #clica no codigo do usuario 
            wait_for_element('//*[@id="dgUser"]/tbody/tr[2]/td[1]/a').click()

            #PROCESSO DE ESCOLHER A APLICAÇÃO ONDE SERÁ GERADO O CARTÃO
            #Aguarda até que a tabela esteja presente na página
            wait = WebDriverWait(navegador, 10)
            wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")))

            #Encontra todas as linhas na tabela
            linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")

            #Percorre as linhas para encontrar a linha com o nome "COMUM"
            for linha in linhas:
                descricao = linha.find_element(By.XPATH, "./td[1]").text  # Declara que a primeira coluna é "Descrição"
                if descricao.startswith("COMUM"):  #Verifica se a descrição começa com "COMUM"
                 #Encontra o botão em "Dados Adicionais" e clica nele
                  botao_dados_adicionais = linha.find_element(By.XPATH, ".//input[contains(@id, 'imgEdit')]")
                  botao_dados_adicionais.click()
                  break

            #Clica na imagem para imprimir o cartão 
            wait_for_element('//*[@id="dgProviders__ctl2_imgNewCard"]').click()
            #Necessita ir nesse site para imprimir
            navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Cards_Ins.aspx')

            #Seleciona o elemento dropdown e escolhe a opçao "COMUM"
            dropdown_desenho = Select(wait_for_element('//*[@id="cboDesign"]'))
            dropdown_desenho.select_by_visible_text("COMUM")
            #Seleciona o elemento dropdown e escolhe a opção 'MIFARE 1K'
            dropdown_tipo = Select(wait_for_element('//*[@id="cboType"]'))
            dropdown_tipo.select_by_visible_text("MIFARE 1K")
            # Seleciona o elemento dropdown e escolhe a opção '1K - COMUM'
            dropdown_template = Select(wait_for_element('//*[@id="cboTemplate"]'))
            dropdown_template.select_by_visible_text("1K - COMUM")
            #Clica no botão para inserir as aplicações
            wait_for_element('//*[@id="btnInsertApplications"]').click()
            # Clica no botão para inserir o cartão para impressora remota
            wait_for_element('//*[@id="btnInserCardRemotePrinter"]').click()

            #Seleciona o elemento dropdown escolhe a opção 'VT'
            dropdown_impressora = Select(wait_for_element('//*[@id="cboRemotePrinter"]') )
            dropdown_impressora.select_by_visible_text("VT")

            #Clica no botão de confirmação para concluir o processo
            wait_for_element('//*[@id="btnConfirm"]').click()
          
            #Volta para a página de pesquisa de usuários
            navegador.get("https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx")

            #Marca como 'OK' na planilha que o cadastro foi criado com sucesso
            criar_cadastro.at[index, 'Criados'] = "OK"
        except Exception as e:
         print(f"Erro ao processar o CPF {usuarios_cpf}: {e}")
         criar_cadastro.at[index, 'Criados'] = "ERRO"

        with pd.ExcelWriter('cartao comum.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
         criar_cadastro.to_excel(writer, sheet_name='Resultado da consulta', index=False)  

finally:
    navegador.quit()