from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import re
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException

# Carregar o arquivo Excel
def processo_secundario(navegador):
    try:
        Obs_c = pd.read_excel('cartao comum.xlsx', sheet_name='Resultado da consulta', engine='openpyxl')
        Obs_c['Observação'] = ''
        
        # Função para aguardar um elemento
        def wait_for_element(xpath, timeout=17):
            return WebDriverWait(navegador, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))  
        
        navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx')
        
        # Seleciona o status "TODOS" no dropdown
        dropdown_status = Select(wait_for_element('//*[@id="cboStatus"]'))
        dropdown_status.select_by_visible_text('TODOS')

        # Iteração através dos dados do arquivo Excel
        for index, row in Obs_c.iterrows():
            try:
                # Processar o CPF e inserir na pesquisa
                usuarios_cpf = re.sub(r'\D', '', str(row['CPF'])).zfill(11)
                inserir_cpf = wait_for_element('//*[@id="txtDoc"]')
                inserir_cpf.clear()
                inserir_cpf.send_keys(f"'{usuarios_cpf}")  # Envia o CPF

                # Iniciar busca e abrir o perfil do usuário
                wait_for_element('//*[@id="btnEldery"]').click()
                wait_for_element('/html/body/div/form/table/tbody/tr/td/fieldset/table/tbody/tr[21]/td/table/tbody/tr[2]/td[1]/a').click()

                # Verificar e inserir a observação
                obs = wait_for_element('//*[@id="txtObservacao"]')
                texto_atual = obs.get_attribute("value")
                
                if texto_atual:
                    obs.send_keys("\n \n")  # Adiciona uma nova linha
                obs.send_keys("1ª via comum - 06/12/24 - BOTB2 - VTONLINE")

                # Verifica se a opção "COMUM" já está presente na coluna 'Descrição'
                descricao_comum = navegador.find_elements(By.XPATH, "//tr/td[text()='COMUM']")
                
                if descricao_comum:
                    # Se "COMUM" já estiver presente, imprime uma mensagem no terminal
                    print("A opção 'COMUM' já está presente na coluna 'Descrição'.")
                else:
                    # Se "COMUM" não estiver presente, imprime que a opção será inserida
                    print("A opção 'COMUM' não está presente na coluna 'Descrição'. Incluindo agora...")
                    dropdown_compra = Select(wait_for_element('//*[@id="cboTypeUser"]'))
                    dropdown_compra.select_by_visible_text("COMUM")
                    wait_for_element('//*[@id="btnInsert"]').click()

                # Verifica se a opção "COMPRA RECARGA" está presente na coluna 'Descrição'
                descricao_compra_recarga = navegador.find_elements(By.XPATH, "//tr/td[text()='COMPRA RECARGA']")
                
                if descricao_compra_recarga:
                    print("A opção 'COMPRA RECARGA' já está presente na coluna 'Descrição'.")
                else:
                    print("A opção 'COMPRA RECARGA' não está presente na coluna 'Descrição'. Incluindo agora...")
                    dropdown_compra = Select(wait_for_element('//*[@id="cboTypeUser"]'))
                    dropdown_compra.select_by_visible_text("COMPRA RECARGA")
                    wait_for_element('//*[@id="btnInsert"]').click()
                    
                # Atualiza a página após a inserção
                wait_for_element('//*[@id="btnUpdate"]').click()

                # Marca como feito no Excel
                Obs_c.at[index, 'Observação'] = "Feito"
                    
                # Retorna à página inicial para processar o próximo CPF
                navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx')
                    
                # Reconfigura o dropdown de status após retornar
                dropdown_status = Select(wait_for_element('//*[@id="cboStatus"]'))
                dropdown_status.select_by_visible_text('TODOS')

            except (NoSuchElementException, StaleElementReferenceException) as e:
                print(f"Erro ao processar o CPF {usuarios_cpf}: {e}")
                Obs_c.at[index, 'Observação'] = f"Erro: {e}"

        # Salva as alterações no arquivo Excel
        with pd.ExcelWriter('cartao comum.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            Obs_c.to_excel(writer, sheet_name='Resultado da consulta', index=False)  # Atualiza a planilha "Resultado da consulta"
    
    finally:
        # Garantir que o navegador será fechado, mesmo que ocorra um erro
        navegador.quit()
