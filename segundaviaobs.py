from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException

# Carregar o arquivo Excel
def processo_secundario(navegador):
    criar_cadastro = pd.read_excel('cartao vt_2.xlsx', sheet_name='GARANTIDO', engine='openpyxl')
    criar_cadastro['Observação'] = '' 

    # Função para aguardar um elemento
    def wait_for_element(xpath, timeout=17):
        return WebDriverWait(navegador, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
    
    # Seleciona o status "TODOS" no dropdown
    dropdown_status = Select(wait_for_element('//*[@id="cboStatus"]'))
    dropdown_status.select_by_visible_text('TODOS')

    # Iteração através dos dados do arquivo Excel
    for index, row in criar_cadastro.iterrows():
        try:
            # Processar o CPF e inserir na pesquisa
            usuarios_cpf = str(row['CPF']).zfill(11)  # Formatar CPF com 11 dígitos
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
            obs.send_keys("2ª via VT - 28/11/24 - BOTB2 - VTONLINE")

            # Verifica se a opção "COMPRA RECARGA" está no dropdown e a seleciona, se necessário
            descricao_comum = navegador.find_elements(By.XPATH, "//tr/td[text()='COMPRA RECARGA']")

            if descricao_comum:
                print("A opção 'COMPRA RECARGA' já está presente na coluna 'Descrição'.")
            else:
                print("A opção 'COMPRA RECARGA' não está presente na coluna 'Descrição'. Incluindo agora...")

                # Selecionar a opção "COMPRA RECARGA" no dropdown
            dropdown_compra = Select(wait_for_element('//*[@id="cboTypeUser"]'))
            dropdown_compra.select_by_visible_text("COMPRA RECARGA")

                # Clica para inserir a nova opção
            wait_for_element('//*[@id="btnInsert"]').click()

                # Atualiza a página após a inserção
            wait_for_element('//*[@id="btnUpdate"]').click()
            criar_cadastro.at[index, 'Observação'] = "Feito"
                
                # Retorna à página inicial para processar o próximo CPF
            navegador.get('https://vtadmin.manaus.prodatamobility.com.br/pages/uca/wfm_Users_Lst.aspx')
                
                # Reconfigura o dropdown de status após retornar
            dropdown_status = Select(wait_for_element('//*[@id="cboStatus"]'))
            dropdown_status.select_by_visible_text('TODOS')

        except (NoSuchElementException, StaleElementReferenceException) as e:
            print(f"Erro ao processar o CPF {usuarios_cpf}: {e}")
            criar_cadastro.at[index, 'Observação'] = f"Erro: {e}"

    # Salva as alterações no arquivo Excel
    with pd.ExcelWriter('cartao vt_2.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        criar_cadastro.to_excel(writer, sheet_name='GARANTIDO', index=False)  # Atualiza a planilha "RIO"
