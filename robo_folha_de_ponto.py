import os
from components.configuracao_selenium_drive import configura_selenium_driver
from components.importacao_diretorios_windows import listagem_arquivos, listagem_arquivos_downloads
from components.procura_elementos_web import procura_elemento, procura_todos_elementos
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from time import sleep
from datetime import datetime
import tkinter as tk
from components.importacao_caixa_dialogo import DialogBox
from components.enviar_emails import enviar_email_com_anexos
from sys import argv
from flask import Flask, request
from flask_restful import Resource, Api, reqparse
import textwrap

PAGE_TIMEOUT = 5
ACTION_TIMEOUT = 1

load_dotenv()

cnpj_email = os.getenv('SELENIUM_CNPJ_EMAIL')
cnpj_password = os.getenv('SELENIUM_CNPJ_PASSWORD')
anexos = []

def rename_files(file, new_name: str = None):
    """
    Renomeia um arquivo mantendo a mesma extensão.

    arquivo: O caminho do arquivo que será renomeado.
    novo_nome: O novo nome para o arquivo (sem a extensão).
    Se não for fornecido, o arquivo será renomeado mantendo o nome original.
    """
    try:
        # Divida o nome do arquivo e sua extensão
        nome_arquivo, extensao = os.path.splitext(file)
        if extensao == '.pdf':
          # Se um novo nome for fornecido, use-o. Caso contrário, mantenha o nome original
          new_file_name = new_name if new_name else nome_arquivo

          today = pd.Timestamp.today().strftime('%d-%m-%Y')

          # Renomeie o arquivo com o novo nome e a mesma extensão
          new_path = f"{os.path.dirname(file)}\\{new_file_name} {today}{extensao}"
          os.rename(file, new_path)

          return new_path
    except FileNotFoundError as not_found_error:
        print(f"Arquivo não encontrado: {not_found_error}")
    except Exception as exc:
        print(f"Ocorreu um erro ao renomear o arquivo: {exc}")

def find_all_datasheet(directory: str = "C://"):
    """Find all Excel datasheet in the given directory.

    Args:
        directory: The path to the directory to search for Excel datasheet.

    Returns:
        A list of file paths to all found datasheet, or an empty list if no datasheet are found.
    """
    # List all files in the directory
    all_files = listagem_arquivos(directory)

    # Filter and return Excel datasheet excluding temporary files
    return [file for file in all_files if file.endswith('.xlsx') and not file.startswith('~$')]

def get_from_datasheet_raw(datasheet: str = "C://", data: str = None):
    """Extracts raw data from Excel datasheet in the specified directory.

    Args:
        path: The directory path where Excel datasheet are located.
        data: The specific column name to extract data from.

    Returns:
        A list containing the data from the specified column across all datasheet.
    """
    extracted_data = []
    df = pd.read_excel(datasheet)
    if data in df.columns:
      for value in df[data].tolist():
        if isinstance(value, str) and value != 'nan':
          extracted_data.append(value)

    return extracted_data

def get_from_datasheet(datasheet: str = """C://""", data: str = None):
  base_data = get_from_datasheet_raw(datasheet, data)
  ret = []

  if data == 'email para envio':
    for i in range(len(base_data)):
      if isinstance(base_data[i], str):
        ret.append(base_data[i].strip().split(','))
    return ret
  
  if data == 'Clientes':
    for i in range(len(base_data)):
      if isinstance(base_data[i], str):
        if isinstance(base_data[i], str):
          if base_data[i].strip().lower() == 's':
            ret.append(True)
          elif base_data[i].strip().lower() == 'n':
            ret.append(False)
    return ret
  
  if data == 'Colaboradores':
    for i in range(len(base_data)):
      if isinstance(base_data[i], str):
        if isinstance(base_data[i], str):
          if base_data[i].strip().lower() == 's':
            ret.append(True)
          elif base_data[i].strip().lower() == 'n':
            ret.append(False)
    return ret

def login(driver, email: str, password: str):
  try:
    driver.get("https://app.tangerino.com.br/Tangerino/pages/LoginPage")
    email_input = procura_elemento(driver, 'xpath', """//*[@id="id4"]""", PAGE_TIMEOUT)
    password_input = procura_elemento(driver, 'xpath', """//*[@id="id8"]""", PAGE_TIMEOUT)
    login_button = procura_elemento(driver, 'xpath', """//*[@id="id9"]""", PAGE_TIMEOUT)

    email_input.send_keys(email)
    password_input.send_keys(password)
    login_button.click()

    sleep(PAGE_TIMEOUT)
  except Exception as e:
    if isinstance(e, NoSuchElementException):
      print('Elemento não encontrado')
    if isinstance(e, TimeoutException):
      print('Tempo de espera excedido')

def ir_para_folha_ponto(driver):
  relatorio_button = procura_elemento(driver, 'xpath', """//*[@id="idc"]/nav[2]/ul/li[5]/div""", PAGE_TIMEOUT)
  if relatorio_button:
    actions = ActionChains(driver)
    actions.move_to_element(relatorio_button).perform()
    sleep(ACTION_TIMEOUT)
    folha_button = procura_elemento(driver, 'xpath', """//*[@id="id39"]""", PAGE_TIMEOUT)
    folha_button.click()
    sleep(PAGE_TIMEOUT)

def preenche_folha_ponto(driver, start_date: str, end_date: str, cliente_nome: str = 'Todos', saldo_horas=True, descanso_semanal=True):
  try:
    try:
      nome_cliente_input = procura_elemento(driver, 'xpath', """//*[@id="mat-input-2"]""", PAGE_TIMEOUT)
      if(nome_cliente_input):
        nome_cliente_input.click()
        nome = cliente_nome.strip().split()
        for index, palavra in enumerate(nome):
          if index < len(nome) - 1:
            nome_cliente_input.send_keys(palavra + ' ')
            sleep(0.1)
          else:
            nome_cliente_input.send_keys(palavra)
            sleep(0.1)
        #nome_cliente_input.send_keys(cliente_nome)
        clientes_encontrados = procura_todos_elementos(driver, 'class_name', 'select-option-custom', PAGE_TIMEOUT)
        if clientes_encontrados:
          for cliente in clientes_encontrados:
            if cliente.text.lower().strip() == cliente_nome.lower().strip():
              cliente.click()
              try:
                saldo_horas = procura_elemento(driver, 'xpath', """//*[@id="checkbox-showHours"]/label""", PAGE_TIMEOUT)
                if(saldo_horas):
                  saldo_horas.click()
              except Exception as e:
                print(f"Erro ao selecionar saldo de horas: {e}")
                if isinstance(e, NoSuchElementException):
                  print('Elemento não encontrado')
                if isinstance(e, TimeoutException):
                  print('Tempo de espera excedido')

              try:
                descanso_semanal = procura_elemento(driver, 'xpath', """//*[@id="checkbox-showDsr"]/label""", PAGE_TIMEOUT)
                if(descanso_semanal):
                  descanso_semanal.click()
              except Exception as e:
                print(f"Erro ao selecionar descanso semanal: {e}")
                if isinstance(e, NoSuchElementException):
                  print('Elemento não encontrado')
                if isinstance(e, TimeoutException):
                  print('Tempo de espera excedido')

              try:
                start_date_input = procura_elemento(driver, 'id', """datepicker-startDate""", PAGE_TIMEOUT)
                if start_date_input:
                  start_date_input.click()
                  start_date_input.send_keys(Keys.CONTROL + 'A')
                  start_date_input.send_keys(Keys.DELETE)
                  start_date_input.send_keys(start_date)
                  start_date_input.send_keys(Keys.ESCAPE)

                end_date_input = procura_elemento(driver, 'id', """datepicker-endDate""", PAGE_TIMEOUT)
                if end_date_input:
                  end_date_input.click()
                  end_date_input.send_keys(Keys.CONTROL + 'A')
                  end_date_input.send_keys(Keys.DELETE)
                  end_date_input.send_keys(end_date)
                  end_date_input.send_keys(Keys.ESCAPE)
              except Exception as e:
                print(f"Erro ao preencher datas: {e}")
                if isinstance(e, NoSuchElementException):
                  print('Elemento não encontrado')
                if isinstance(e, TimeoutException):
                  print('Tempo de espera excedido')
              
              folha_de_ponto = download_folha_ponto(driver)
              sleep(2)               
              return folha_de_ponto
          else:
            print("Nenhum cliente encontrado")
            return
    except Exception as e:
      print(f"Erro ao preencher folha de ponto: {e}")
      if isinstance(e, NoSuchElementException):
        print('Elemento não encontrado')
      if isinstance(e, TimeoutException):
        print(f"Tempo de espera excedido {e.msg}\n{e.stacktrace}")

  except Exception as e:
    print(f"Erro ao preencher folha de ponto: {e}")
    if isinstance(e, NoSuchElementException):
      print('Elemento não encontrado')
    if isinstance(e, TimeoutException):
      print(f"Tempo de espera excedido {e.msg}\n{e.stacktrace}")

def download_folha_ponto(driver):
  try:
    gerar_button = procura_elemento(driver, 'xpath', """//*[@id="btn-generate-simple"]""", PAGE_TIMEOUT)
    if gerar_button:
      gerar_button.click()
      sleep(12)
      download_button = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[1]/a""", PAGE_TIMEOUT)
      if download_button:
        download_button_class = download_button.get_attribute('class')
        while 'disabled' in download_button_class:
          sleep(PAGE_TIMEOUT)
          driver.refresh()
          sleep(1)
          download_button = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[1]/a""", PAGE_TIMEOUT)
          download_button_class = download_button.get_attribute('class')
          message = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[7]/span""", PAGE_TIMEOUT).text
          if message.strip() == 'Erro no processamento!':
            print('Erro ao baixar folha de ponto')
            break

        if not 'disabled' in download_button_class:
          download_button.click()
          folha_ponto_name = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[2]""", PAGE_TIMEOUT).text
          return folha_ponto_name
        
  except Exception as e:
    if isinstance(e, NoSuchElementException):
      print('Elemento não encontrado')
    if isinstance(e, TimeoutException):
      print('Tempo de espera excedido')

# Função para formatar o corpo do e-mail
def format_email_body(body: str) -> str:
    # Dividir os parágrafos em cada ponto seguido por um espaço
    paragrafos = body.strip().split('\n')
    
    # Adicionar tabulação no início de cada parágrafo
    paragrafos_com_tabulacao = [f"\t{paragrafo.strip()}" for paragrafo in paragrafos if paragrafo.strip()]

    # Adicionar quebra de linha dupla entre parágrafos
    email_body_formatado = "\n\n".join(paragrafos_com_tabulacao)

    return email_body_formatado

def gerar_folha(start_date: str, end_date: str, particao: str):
  datasheet = f"""{particao}:\\Meu Drive\\15. Arquivos_Automacao\\tangRh\\informacoes-robo-tangrh-correto.xlsx"""
  if len(datasheet) > 0:
    chrome_options, service = configura_selenium_driver()
    driver = webdriver.Chrome(options=chrome_options, service=service)

    clientes = get_from_datasheet_raw(datasheet, 'centro de custo')
    emails = get_from_datasheet(datasheet, 'email para envio')
    na_plataforma = get_from_datasheet(datasheet, 'Clientes')
    tem_colaborador = get_from_datasheet(datasheet, 'Colaboradores')

    login(driver, cnpj_email, cnpj_password)
    ir_para_folha_ponto(driver)
    
    for i in range(len(clientes)):
      if not na_plataforma[i]:
        continue

      if not tem_colaborador[i]:
        continue

      anexos.clear()

      prev_url = driver.current_url
      embed = procura_elemento(driver, 'tag_name', """embed""", PAGE_TIMEOUT)
      if embed:
        embed_src = embed.get_attribute('src')
        driver.get(embed_src)

      folha_de_ponto = preenche_folha_ponto(driver, start_date, end_date, clientes[i])
      driver.get(prev_url)

      if folha_de_ponto:
        arquivos_download = listagem_arquivos_downloads()
        arquivo_mais_recente = max(arquivos_download, key=os.path.getmtime)
        sleep(PAGE_TIMEOUT)
        arquivo_mais_recente = rename_files(arquivo_mais_recente, f"Folha de Ponto - {clientes[i]}")
        anexos.append(arquivo_mais_recente)

        email_body = format_email_body(f"""
                                  Gostaríamos de informar que a folha de ponto referente a {datetime.strptime(start_date, '%d%m%Y').strftime("%d/%m/%Y")} - {datetime.strptime(end_date, "%d%m%Y").strftime("%d/%m/%Y")} foi gerada com sucesso e está disponível para análise e eventual correção, caso necessário.
                                  Por favor, acesse {os.path.basename(arquivo_mais_recente)} para visualizar e verificar as informações registradas. 
                                  Caso identifique qualquer inconsistência ou discrepância em seu registro, por gentiliza entre em contato imediatamente.
                                  Salientamos a importância da verificação cuidadosa dos registros de ponto, a fim de garantir a precisão e integridade das informações relacionadas à jornada de trabalho da sua empresa.
                                  Agradecemos antecipadamente pela sua atenção e colaboração neste processo.
                                  """)
        enviar_email_com_anexos("bruno.apolinario010@gmail.com", f"Folha de Ponto - {clientes[i]}", email_body, anexos)
        os.remove(arquivo_mais_recente)
    
    driver.quit()
    return True

app = Flask(__name__)
api = Api(app)

class execute(Resource):
  def post(self):
    print("Requisição recebida")
    # Verifique se todos os campos obrigatórios foram fornecidos no JSON
    parser = reqparse.RequestParser()
    parser.add_argument('Robo Folha de Ponto - Data Inicial', required=True)
    parser.add_argument('Robo Folha de Ponto - Data Final', required=True)
    parser.add_argument('Robo Folha de Ponto - Particao', required=True)
    json_data = parser.parse_args()

    # Se todos os campos estiverem presentes, prossiga com a execução do programa
    data1 = json_data['Robo Folha de Ponto - Data Inicial']
    data2 = json_data['Robo Folha de Ponto - Data Final']
    particao = json_data['Robo Folha de Ponto - Particao']

    print(f"Data 1: {data1}")
    print(f"Data 2: {data2}")
    print(f"Partição: {particao}")
    data1 = ''.join(reversed(data1.split('-')))
    data2 = ''.join(reversed(data2.split('-')))
    sucesso = gerar_folha(data1, data2, particao)

    if sucesso:
      return {'message': 'Folha de ponto gerada com sucesso'}, 200
    else:
      return {'message': 'Erro ao gerar folha de ponto'}, 500

class shutdown(Resource):
  def post(self):
    os._exit(0)

api.add_resource(execute, '/')
api.add_resource(shutdown, '/shutdown')

if __name__ == "__main__":
  app.run(debug=True, port=5000)