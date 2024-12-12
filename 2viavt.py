import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import openpyxl
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
import time
import re
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from segundaviaobs import processo_secundario

# Carregar os dados do Excel
criar_cadastro = pd.read_excel('cartao vt_2.xlsx', sheet_name='GARANTIDO', engine='openpyxl')

criar_cadastro['Criados'] = ''
criar_cadastro['Transferencia'] = ''  # Nova coluna para registrar o status da transferência

# Inicia o navegador
navegador = webdriver.Firefox()

def wait_for_element(xpath):
    return WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))

def inserir_cpf():
     usuarios_cpf = re.sub(r'\D', '', str(row['CPF'])).zfill(11)
     inserir_cpf = wait_for_element('//*[@id="txtDoc"]')
     inserir_cpf.clear()
     inserir_cpf.send_keys(f"'{usuarios_cpf}")  # Envia o CPF

def voltar_para_pagina_inicial():
    navegador.get("https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx")
    wait_for_element('//*[@id="txtDoc"]')  # Aguarda o campo para inserir o CPF

try:
    # Acessar a página de login
    navegador.get("https://vtadmin.manaus.prodatamobility.com.br/wfm_Home.aspx")

    # Login usuário
    wait_for_element('//*[@id="txtLogin"]').send_keys("Comercial.sntr")
    wait_for_element('//*[@id="txtSenha"]').send_keys("Botnelly")
    wait_for_element('//*[@id="loginbutton"]').click()

    # Vai para a página do UCA
    wait_for_element('//*[@id="parent_uca"]/a').click()
    wait_for_element('//*[@id="xmlmenu_right"]/li[1]/ul/li[2]/a').click()

    # Outro site dos usuários
    voltar_para_pagina_inicial()

    # Selecionar o status
    dropdown_status = Select(wait_for_element('//*[@id="cboStatus"]'))
    dropdown_status.select_by_visible_text("TODOS")

    # Processa cada linha do dataframe
    for index, row in criar_cadastro.iterrows():
        try:
            inserir_cpf()

            # Aqui vem o processo de buscar o cartão
            wait_for_element("/html/body/div/form/table/tbody/tr/td/fieldset/table/tbody/tr[21]/td/table/tbody/tr[2]/td[1]/a").click()

            # Agora vamos adicionar o processo de verificar a tabela de cartões
            wait_for_element("//table[@id='gvCards']")  # Aguarda a tabela de cartões ser carregada
            linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")
            
            # Loop para verificar cada cartão
            for linha in linhas:
                try:
                    cartao = linha.find_element(By.XPATH, "./td[1]").text  # A primeira coluna é "Cartão"
                    status = linha.find_element(By.XPATH, "./td[2]").text  # A segunda coluna é "Status"

                    # Verifica se o cartão começa com "58.04" e o status é "Aguardando"
                    if cartao.startswith("58.04") and status == "Aguardando":
                        # Clique no link ou botão necessário nesta linha
                        link_cartao = linha.find_element(By.XPATH, "./td[1]/a")
                        link_cartao.click()

                    # Se o cartão começa com "58.04" e o status é "ATIVO", pula para o próximo CPF
                    elif cartao.startswith("58.04") and status == "ATIVO":
                        criar_cadastro.at[index, 'Criados'] = "Ativo"
                        break  # Sai do loop para procurar o próximo CPF

                except StaleElementReferenceException:
                    # Se o erro for StaleElementReferenceException, tente encontrar novamente o elemento
                    print("Elemento obsoleto, tentando localizar novamente...")
                    linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")
                    continue

            # Se o cartão foi encontrado e o status é "Aguardando", o código continua o processamento
            if cartao.startswith("58.04") and status == "Aguardando":
                wait_for_element('//*[@id="lnkRemotePrinter"]').click()

                dropdown_impressora = Select(wait_for_element('//*[@id="cboRemotePrinter"]'))
                dropdown_impressora.select_by_visible_text("VT")
                wait_for_element('//*[@id="btnConfirm"]').click()

            # Volta para a página de pesquisa de usuários sem reiniciar a página
            voltar_para_pagina_inicial()  # Função para voltar à página inicial
            # Atualiza a coluna 'Criados' após concluir o processo
            criar_cadastro.at[index, 'Criados'] = "Feito"
        except Exception as e:
            print(f"Erro ao processar o ID {inserir_cpf()}: {e}")
            with pd.ExcelWriter('cartao vt_2.xlsx', engine='openpyxl', mode='a') as writer:
             criar_cadastro.at[index, 'Criados'] = "ERRO"

    # Segunda parte: 
    for index, row in criar_cadastro.iterrows():
        tentativa_transferencia = 0  # Contador para tentativas de transferência

        while tentativa_transferencia < 3:  # Limita a 3 tentativas para evitar loops infinitos
            try:
                inserir_cpf()

                # Espera pelo elemento e clica no link
                wait_for_element("/html/body/div/form/table/tbody/tr/td/fieldset/table/tbody/tr[21]/td/table/tbody/tr[2]/td[1]/a").click()

                # Aguarda a tabela de cartões ser carregada
                wait_for_element("//table[@id='gvCards']")  
                linhas = navegador.find_elements(By.XPATH, "//tr[contains(@class, 'GridLinha') or contains(@class, 'GridLinhaAlternada')]")

                # descorre pelas linhas da tabela de cartões
                for linha in linhas:
                    try:
                        # Verifica se o link "Transferir" existe
                        transf_credito = linha.find_elements(By.XPATH, "./td[4]//a[contains(text(), 'Transferir')]")
                        # Se o link "Transferir" existir, clica nele
                        if transf_credito:
                            transf_credito[0].click()  # Clica no primeiro link encontrado
                            time.sleep(3)  # Aguarda o processo de transferência
                            # Verificar se há um popup e aceitá-lo automaticamente
                            try:
                                alert = WebDriverWait(navegador, 5).until(EC.alert_is_present())
                                alert.accept()  # função OK
                                time.sleep(2)  
                                wait_for_element('//*[@id="btnTransfer"]').click() #botao para transferir o credito
                                time.sleep(2)  
                                voltar_para_pagina_inicial()
                                criar_cadastro.at[index, 'Transferencia'] = "Transferido"  # Atualiza a coluna de transferência
                                break  # Sai do loop, pois a transferência foi realizada
                            except TimeoutException:
                                print("Nenhum alerta apareceu.")
                                criar_cadastro.at[index, 'Transferencia'] = "Não Transferido"
                                break

                    except NoSuchElementException:
                        continue  # Se a linha não tiver a estrutura esperada, continua para a próxima linha

                # Se não encontrar o link "Transferir", volta para a página de pesquisa
                if not transf_credito:
                    print(f"Link 'Transferir' não encontrado para o CPF {inserir_cpf()}. Retornando à página inicial.")
                    voltar_para_pagina_inicial()
                    tentativa_transferencia += 1  # numero de tentativas
                    break  # Sai do loop de tentativas

            except Exception as e:
                print(f"Erro ao processar o CPF {inserir_cpf()}: {e}")
                criar_cadastro.at[index, 'Transferencia'] = "Erro"
                tentativa_transferencia += 1  # Incrementa a tentativa, para sair do loop após 3 tentativas

    # Salvar os dados no Excel
    with pd.ExcelWriter('cartao vt_2.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
     criar_cadastro.to_excel(writer, sheet_name='GARANTIDO', index=False)  # Atualiza a planilha "PONTE"
     processo_secundario(navegador)

finally:
    navegador.quit()