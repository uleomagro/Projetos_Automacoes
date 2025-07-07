import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import time
import os

print('='*80 + '\n' + '||' + '\033[1mINICIANDO...\033[0m'.center(84) + '||' + '\n' + '='*80)

def main():
    # Carrega a planilha Excel
    excel_path = os.path.join(os.path.dirname(__file__), "lista.xlsx")
    df = pd.read_excel(excel_path)

    # Inicia o navegador com o ChromeDriver atualizado automaticamente
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://app.acessorias.com/index.php") 
    time.sleep(1)

    # Aguarda o usuário fazer login manualmente no sistema
    print('='*80)
    print('||' + "🔒  Por favor, faça o login \033[1mMANUALMENTE\033[0m no"
    " sistema Acessórias. ".center(83) + '||')
    input('||' + " 🔓  Apos o login \033[1mCLIQUE NESTE TERMINAL\033[0m e"
    " pressione \033[1mENTER\033[0m para continuar... ".center(84) + '||' + '\n' + '='*80 + '\n')
    time.sleep(1)

    print('='*80 + '\n' + '||' + 'ACESSANDO EMPRESAS'.center(76) + '||' + '\n' + '='*80)

    # Acessa o menu "Empresas" antes de buscar
    driver.find_element(By.LINK_TEXT, "Empresas").click()
    time.sleep(1)

    for _, row in df.iterrows():
        # Define as variaves com os dados do Excel
        sucesso = False
        tentativas = 0
        codigo = str(row["Codigo"])
        departamento = row["Departamento"]
        novo_responsavel = row["Novo Responsavel"]

        while not sucesso and tentativas < 2:
            try:
                tentativas += 1

                # Busca pela empresa
                busca_input = driver.find_element(By.XPATH, "//input[@placeholder='Procurar na relação']")
                busca_input.clear()
                busca_input.send_keys(codigo)
                busca_input.send_keys(Keys.RETURN)
                time.sleep(2)

                # Clica no ícone de edição dos responsáveis pelos departamentos
                driver.find_element(By.XPATH, "(//*[@title='Responsáveis pelos departamentos'])[1]").click()
                time.sleep(2)

                from selenium.webdriver.support.ui import WebDriverWait
                from selenium.webdriver.support import expected_conditions as EC

                # Ids dos Departamentos criados no sistema              
                departamento_ids = {
                    "ADMINISTRAÇÃO DE PESSOAL": "3",
                    "ARQUIVO - EXPEDIENTE": "8",
                    "ROTEIRISTA": "16",
                    "CONTABILIDADE": "1",
                    "COORDENADORES": "12",
                    "FINANCEIRO": "4",
                    "FISCAL": "2",
                    "INFORMÁTICA": "13",
                    "LEGALIZAÇÃO": "5",
                    "ALVARÁS E LICENÇAS": "22",
                    "CERTIDÕES": "21",
                    "CERTIFICAÇÃO DIGITAL": "20",
                    "COMERCIAL": "11",
                    "TAXA MUNICIPAL": "18",
                    "RECEPÇÃO": "17",
                    "MALA DIRETA": "19"
                }

                id_departamento = departamento_ids.get(departamento.strip().upper())
                if not id_departamento:
                    raise ValueError(f"⚠️ ID do departamento '{departamento}' não encontrado.")

                xpath_dinamico = f"//div[@class='col-sm-12 col-xs-12 no-padding' and contains(@id, 'editDptEmpResp_{id_departamento}_')]//button[@title='Editar os responsáveis']"

                # Clica no botão de edição do respectivo departamento
                elemento = driver.find_element(By.XPATH, f"//*[contains(text(), '{departamento}')]")
                if elemento:
                    botao_editar_elemento = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, xpath_dinamico))
                    )

                    driver.execute_script("document.body.style.zoom='80%'")
                    time.sleep(1)

                    botao_editar_elemento.click()

                # Entra na lista suspensa
                id_dropdown = f"inpResp_{id_departamento}_{codigo}"
                dropdown_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, id_dropdown))
                )

                select = Select(dropdown_element)
                opcoes = [o.text.strip() for o in select.options]

                # Define o novo responsavel
                if novo_responsavel.strip() in opcoes:
                    select.select_by_visible_text(novo_responsavel.strip())
                    print(f"✅ Empresa {codigo} - Responsável '{novo_responsavel}' selecionado com sucesso.")
                else:
                    raise ValueError(f"⚠️ Nome '{novo_responsavel}' não está disponível na lista.")

                # Salvar a alteração
                id_botao_salvar = f"RespSave_{id_departamento}_{codigo}"
                salvar_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        f"//div[@id='{id_botao_salvar}']//button[@title='Salvar alterações']"
                    ))
                )
                salvar_button.click()
                print(f"💾 Alterações salvas com sucesso para '{departamento}'")

                print('~'*20)
                sucesso = True  # Se chegou até aqui, tudo deu certo!

            # Em casos de erros, recomeça o processo
            except Exception as e:
                print(f"🔁 Tentativa {tentativas} falhou para código {codigo} / departamento '{departamento}'.")
                print(f"Erro: {str(e).splitlines()[0]}")
                print("Aguardando antes de tentar novamente...\n")
                time.sleep(3)

        if not sucesso:
            print(f"❌ Falha definitiva ao processar código {codigo} após 2 tentativas.\n{'='*80}")

    driver.quit()

if __name__ == "__main__":
    main()

input('='*80 + '\n' + '||' + 'FINALIZANDO... Pressione ENTER para sair.'.center(76) + '||' + '\n' + '='*80)
