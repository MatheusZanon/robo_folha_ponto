import os
from components.configuracao_selenium_drive import configura_selenium_driver
from components.importacao_diretorios_windows import listagem_arquivos, listagem_arquivos_downloads
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from time import sleep

PAGE_TIMEOUT = 5
ACTION_TIMEOUT = 1

load_dotenv()

cnpj_email = os.getenv('SELENIUM_CNPJ_EMAIL')
cnpj_password = os.getenv('SELENIUM_CNPJ_PASSWORD')

def procura_elemento(driver, tipo_seletor:str, elemento, tempo_espera=PAGE_TIMEOUT):
  """
    Function to search for an element using the specified selector type, element, and wait time.
    driver: The WebDriver instance to use for element search.
    tipo_seletor: The type of selector to use (e.g., 'ID', 'CLASS_NAME', 'XPATH', 'TAG_NAME').
    elemento: The element to search for.
    tempo_espera: The maximum time to wait for the element to be located.
    :return: The located element if found, otherwise None.
  """
  try:
    seletor = getattr(By, tipo_seletor.upper())
    WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_element_located((seletor, elemento)))
    sleep(0.1)
    elemento = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_element_located((seletor, elemento)))
    if elemento.is_displayed() and elemento.is_enabled():
      return elemento
  except TimeoutException:
    return None

def procura_todos_elementos(driver, tipo_seletor:str, elemento, tempo_espera=PAGE_TIMEOUT):
  """
    A function that searches for all elements based on the given selector type and element, within a specified waiting time.
    
    Args:
      driver: The WebDriver instance to use for locating the elements.
      tipo_seletor: A string representing the type of selector to use (e.g., 'ID', 'CLASS_NAME', 'XPATH', 'TAG_NAME').
      elemento: The element to search for.
      tempo_espera: The maximum time to wait for the elements to be present before throwing a TimeoutException.
      
    Returns:
      A list of WebElement objects representing the found elements, or None if the elements are not found within the specified waiting time.
  """
  try:
    seletor = getattr(By, tipo_seletor.upper())
    WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_all_elements_located((seletor, elemento)))
    sleep(0.1)
    elementos = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_all_elements_located((seletor, elemento)))
    if elementos:
      return elementos
  except TimeoutException:
    return None

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

          print(f"Arquivo renomeado para: {new_path}")
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
    print(all_files)

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
    email_input = procura_elemento(driver, 'xpath', """//*[@id="id4"]""")
    password_input = procura_elemento(driver, 'xpath', """//*[@id="id8"]""")
    login_button = procura_elemento(driver, 'xpath', """//*[@id="id9"]""")

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
  relatorio_button = procura_elemento(driver, 'xpath', """//*[@id="id31"]""")
  if relatorio_button:
    actions = ActionChains(driver)
    actions.move_to_element(relatorio_button).perform()
    sleep(ACTION_TIMEOUT)
    folha_button = procura_elemento(driver, 'xpath', """//*[@id="id34"]""")
    folha_button.click()
    sleep(PAGE_TIMEOUT)

def preenche_folha_ponto(driver, cliente_nome: str = 'Todos', saldo_horas=True, descanso_semanal=True):
  try:
    try:
      nome_cliente_input = procura_elemento(driver, 'xpath', """//*[@id="mat-input-2"]""")
      if(nome_cliente_input):
        nome_cliente_input.click()
        nome_cliente_input.send_keys(cliente_nome)
        input()
        sleep(1)
        clientes_encontrados = procura_todos_elementos(driver, 'class_name', 'select-option-custom')
        if clientes_encontrados:
          for cliente in clientes_encontrados:
            if cliente.text.lower() == cliente_nome.lower():
              print(f"Cliente selecionado: {cliente.text}")
              sleep(1)
              cliente.click()
              break
    except Exception as e:
      if isinstance(e, NoSuchElementException):
        print('Elemento não encontrado')
      if isinstance(e, TimeoutException):
        print(f"Tempo de espera excedido {e.msg}\n{e.stacktrace}")

    try:
      saldo_horas = procura_elemento(driver, 'xpath', """//*[@id="checkbox-showHours"]/label""")
      if(saldo_horas):
        saldo_horas.click()
    except Exception as e:
      if isinstance(e, NoSuchElementException):
        print('Elemento não encontrado')
      if isinstance(e, TimeoutException):
        print('Tempo de espera excedido')

    try:
      descanso_semanal = procura_elemento(driver, 'xpath', """//*[@id="checkbox-showDsr"]/label""")
      if(descanso_semanal):
        descanso_semanal.click()
    except Exception as e:
      if isinstance(e, NoSuchElementException):
        print('Elemento não encontrado')
      if isinstance(e, TimeoutException):
        print('Tempo de espera excedido')
  except Exception as e:
    print(f"Erro ao preencher folha de ponto: {e}")
    if isinstance(e, NoSuchElementException):
      print('Elemento não encontrado')
    if isinstance(e, TimeoutException):
      print(f"Tempo de espera excedido {e.msg}\n{e.stacktrace}")

"""def esvazia_folha_ponto(driver):
  try:
    if driver.current_url != "https://app.tangerino.com.br/Tangerino/pages/folha-ponto?funcionalidade=24":
      driver.get("https://app.tangerino.com.br/Tangerino/pages/folha-ponto?funcionalidade=24")
      sleep(PAGE_TIMEOUT)
  except Exception as e:
    if isinstance(e, NoSuchElementException):
      print('Elemento não encontrado')
    if isinstance(e, TimeoutException):
      print('Tempo de espera excedido')"""

def download_folha_ponto(driver, cliente_nome: str = 'Todos'):
  try:
    gerar_button = procura_elemento(driver, 'xpath', """//*[@id="btn-generate-simple"]""")
    if gerar_button:
      gerar_button.click()
      sleep(12)
      print(f"Downloading...")
      download_button = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[1]/a""")
      if download_button:
        download_button_class = download_button.get_attribute('class')
        while 'disabled' in download_button_class:
          print(f"Download button is disabled, refreshing in {PAGE_TIMEOUT} seconds")
          sleep(PAGE_TIMEOUT)
          driver.refresh()
          print(f"Downloading...")
          sleep(1)
          download_button = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[1]/a""")
          download_button_class = download_button.get_attribute('class')

        if not 'disabled' in download_button_class:
          download_button.click()
          folha_ponto_name = procura_elemento(driver, 'xpath', """/html/body/app-root/app-report-time-sheet/div/section/div[4]/table/tbody/tr[1]/td[2]""").text
          return folha_ponto_name
        
  except Exception as e:
    if isinstance(e, NoSuchElementException):
      print('Elemento não encontrado')
    if isinstance(e, TimeoutException):
      print('Tempo de espera excedido')


def main():
  datasheet = """H:\\Meu Drive\\15. Arquivos_Automacao\\tangRh\\informacoes-robo-tangrh-correto.xlsx"""
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

      print(f"{i:>6}| {clientes[i]:<60} {na_plataforma[i]} {tem_colaborador[i]} {emails[i]}\n")

      prev_url = driver.current_url
      embed = procura_elemento(driver, 'tag_name', """embed""")
      if embed:
        embed_src = embed.get_attribute('src')
        driver.get(embed_src)

      preenche_folha_ponto(driver, clientes[i])
      downloads_antigos = listagem_arquivos_downloads()
      folha_de_ponto = download_folha_ponto(driver, clientes[i])

      driver.get(prev_url)

      if folha_de_ponto:
        sleep(PAGE_TIMEOUT)
        downloads_novos = listagem_arquivos_downloads()
        arquivos_renomear = list(set(downloads_novos) - set(downloads_antigos))
        if len(arquivos_renomear) > 0:
          rename_files(arquivos_renomear[0], f"Folha de Ponto - {clientes[i]}")
    
    input()
    driver.quit()

main()