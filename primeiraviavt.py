from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import time
import re
from selenium.common.exceptions import StaleElementReferenceException
from primeiraiviaobs import processo_vt2

#IMPRIME 1ª VIA VT 
#Carregar os dados do Excel
cartao_criado = pd.read_excel('cartao vt.xlsx')
#Cria uma coluna nova vazia 
cartao_criado['Cartão Feito'] = ''

# Inicia o navegador
navegador = webdriver.Firefox()
#declara variavel para aguardar um elemento
def wait_for_element(xpath):
    return WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))

def inserir_cpf():
     #declara variavel como texto 
        usuarios_cpf = re.sub(r'\D', '', str(row['CPF'])).zfill(11)
          
            #variavel para o campo / aguarda o elemento aparecer
        inserir_cpf = wait_for_element('//*[@id="txtDoc"]')
            #limpa o elemento antes de inserir
        inserir_cpf.clear()
            #define a variavel como texto para inserir
        inserir_cpf.send_keys(f"'{usuarios_cpf}")  # Envia o CPF

            # botao pesquisar
        wait_for_element('//*[@id="btnEldery"]').click()

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
    for index, row in cartao_criado.iterrows():
        try:
            inserir_cpf()

            #clica no codigo do usuario 
            wait_for_element("/html/body/div/form/table/tbody/tr/td/fieldset/table/tbody/tr[21]/td/table/tbody/tr[2]/td[1]/a").click()

            #Aguarda até que a tabela esteja presente na página
            wait_for_element("//table[@id='gvCards']")  # Aguarda a tabela de cartões ser carregada

            #Encontra todas as linhas na tabela
            linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")

            #looping cadastros
            for linha in linhas:
                try:
                    cartao = linha.find_element(By.XPATH, "./td[1]").text  # A primeira coluna é "Cartão"
                    status = linha.find_element(By.XPATH, "./td[2]").text  # A segunda coluna é "Status"

                    # Verifica se o cartão começa com "58.04" e o status é "Aguardando"
                    if cartao.startswith("58.04") and status == "Aguardando":
                        # Clica no botão  de aguardando
                        link_cartao = linha.find_element(By.XPATH, "./td[1]/a")
                        link_cartao.click()

                    # Se o cartão começa com "58.04" e o status é "ATIVO", pula para o próximo CPF
                    elif cartao.startswith("58.04") and status == "ATIVO":
                        cartao_criado.at[index, 'Cartão Feito'] = "Ativo"
                        #emite uma mensagem no terminal se ja estiver ativo o cartao
                        print(f"Cartão {cartao} está ATIVO. Pulando para o próximo CPF.")
                        #salva na planilha que ja estava ativo o cartao
                        break  # Sai do loop para procurar o próximo CPF
                except StaleElementReferenceException:
                    #Se o erro for StaleElementReferenceException, tente encontrar novamente o elemento\nao sei qual erro é mas as vezes some o elemento, nao mexer pf
                    print("Elemento obsoleto, tentando localizar novamente...")
                    #Encontra todas as linhas na tabela novamente
                    linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")
                    continue

            # Se o cartão foi encontrado e o status é "Aguardando", o código continua o processamento
            if cartao.startswith("58.04") and status == "Aguardando":
                #Clica no botão para inserir o cartão para impressora remota
                wait_for_element('//*[@id="lnkRemotePrinter"]').click()
                #Seleciona o elemento dropdown escolhe a opção 'VT'
                dropdown_impressora = Select(wait_for_element('//*[@id="cboRemotePrinter"]'))
                dropdown_impressora.select_by_visible_text("VT")
                #confirma a impressao
                wait_for_element('//*[@id="btnConfirm"]').click()
                cartao_criado[index, 'Cartão Feito'] = "Feito"

            #Volta para a página de pesquisa de usuários sem reiniciar a página
            navegador.get("https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx")

            #Aguardar a página carregar e recomeçar o processo para o próximo CPF
            wait_for_element('//*[@id="txtDoc"]')  # Verifica se o campo para inserir o CPF está carregado

            #Marca como 'feito' na planilha que o cadastro foi criado com sucesso
        except Exception as e:
         print(f"Erro ao processar o CPF {inserir_cpf()}: {e}")
         cartao_criado.at[index, 'Cartão Feito'] = "ERRO"
         
        with pd.ExcelWriter('cartao vt.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
         cartao_criado.to_excel(writer, sheet_name='Resultado da consulta', index=False)  

finally:
    processo_vt2(navegador)
    navegador.quit()
    