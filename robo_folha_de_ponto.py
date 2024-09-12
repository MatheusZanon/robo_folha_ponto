import os
from components.importacao_diretorios_windows import listagem_arquivos_downloads
from components.configuracao_db import configura_db, ler_sql
import mysql.connector
from dotenv import load_dotenv
from time import sleep
from datetime import datetime, timezone, timedelta
from components.enviar_emails import enviar_email_com_anexos
import os, requests, json, base64, platform
from shutil import copy
from typing import Literal
import traceback
import boto3

PAGE_TIMEOUT = 5
ACTION_TIMEOUT = 1

cnpj_email = os.getenv('SELENIUM_CNPJ_EMAIL')
cnpj_password = os.getenv('SELENIUM_CNPJ_PASSWORD')
anexos = []

# ================= CARREGANDO VARIÁVEIS DE AMBIENTE=======================
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf = configura_db()

# ============================ AUXILIARES =================================

VALID_TIMEZONES = {
    'UTC-12': -12, 'UTC-11': -11, 'UTC-10': -10, 'UTC-9': -9,
    'UTC-8': -8, 'UTC-7': -7, 'UTC-6': -6, 'UTC-5': -5,
    'UTC-4': -4, 'UTC-3': -3, 'UTC-2': -2, 'UTC-1': -1,
    'UTC+0': 0, 'UTC+1': 1, 'UTC+2': 2, 'UTC+3': 3,
    'UTC+4': 4, 'UTC+5': 5, 'UTC+6': 6, 'UTC+7': 7,
    'UTC+8': 8, 'UTC+9': 9, 'UTC+10': 10, 'UTC+11': 11,
    'UTC+12': 12, 'UTC+13': 13, 'UTC+14': 14
}

TimeZoneStr = Literal[
    'UTC-12', 'UTC-11', 'UTC-10', 'UTC-9', 'UTC-8', 'UTC-7', 'UTC-6', 'UTC-5',
    'UTC-4', 'UTC-3', 'UTC-2', 'UTC-1', 'UTC+0', 'UTC+1', 'UTC+2', 'UTC+3',
    'UTC+4', 'UTC+5', 'UTC+6', 'UTC+7', 'UTC+8', 'UTC+9', 'UTC+10', 'UTC+11',
    'UTC+12', 'UTC+13', 'UTC+14'
]

def get_download_path():
    """Returns the default downloads path for Linux, MacOS, or Windows."""
    if platform.system() == 'Windows':
        from winreg import OpenKey, QueryValueEx, HKEY_CURRENT_USER
        
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        
        with OpenKey(HKEY_CURRENT_USER, sub_key) as key:
            downloads_path = QueryValueEx(key, downloads_guid)[0]
        
        return str(downloads_path)
    
    else:
        from pathlib import Path
        
        return str(Path.home() / 'Downloads')

def copia_folha_baixada(nome_cliente, mes, ano, pasta_cliente):
    try:
        arquivos_downloads = listagem_arquivos_downloads()
        arquivo_mais_recente = max(arquivos_downloads, key=os.path.getmtime)
        if (arquivo_mais_recente.__contains__(".pdf") 
            and not arquivo_mais_recente.__contains__(f"Boleto_Recebimento_{nome_cliente.replace("S/S", "S S")}_{ano}.{mes}")):
            caminho_pdf = os.Path(arquivo_mais_recente)
            novo_nome_boleto = caminho_pdf.with_name(f"Boleto_Recebimento_{nome_cliente.replace("S/S", "S S")}_{ano}.{mes}.pdf")
            caminho_pdf_mod = caminho_pdf.rename(novo_nome_boleto)
            sleep(0.5)
            copy(caminho_pdf_mod, pasta_cliente / caminho_pdf_mod.name)
            if os.path.exists(caminho_pdf_mod):
                os.remove(caminho_pdf_mod)
            else:
                print("Arquivo nao encontrado no caminho para remocão!")
        else:
            print("Arquivo de boleto não encontrado!")
    except Exception as error:
        print(f"Erro ao copiar o arquivo: {error}")
        input("Pressione ENTER para sair...")

def convert_datetimes(data):
		"""
		Converts datetime values in the input data dictionary to a specific string format.
		
		Args:
			data (dict): A dictionary with key-value pairs to be converted.
		
		Returns:
			dict: The data dictionary with datetime values converted to the specified string format.
		"""
		for key, value in data.items():
			if isinstance(value, datetime):
				data[key] = value.strftime('%Y-%m-%dT%H:%M:%S')
		return data

def row_to_dict(row, column_names):
	"""
	Creates a dictionary from the input 'row' and 'column_names' by zipping them together.
	
	Args:
		row (iterable): The row to be converted into a dictionary.
		column_names (iterable): The names of the columns to be used as keys in the dictionary.
	
	Returns:
		dict: A dictionary where 'column_names' are keys and 'row' values are values.
	"""
	return dict(zip(column_names, row))

def process_clientes(clientes, column_names):
	clientes_dicts = []
	for cliente in clientes:
		cliente_dict = row_to_dict(cliente, column_names)
		cliente_dict = convert_datetimes(cliente_dict)
		clientes_dicts.append(cliente_dict)
	return clientes_dicts

def formatCNPJ(self, cnpj):
	"""
	Method to format cnpj numbers.
	Tests:
	
	>>> print Cnpj().format('53612734000198')
	>>> 53.612.734/0001-98
	"""
	return "%s.%s.%s/%s-%s" % ( cnpj[0:2], cnpj[2:5], cnpj[5:8], cnpj[8:12], cnpj[12:14] )

# ============================ FUNÇOES ====================================

def unix_time_millis(dt_str, tz_str: TimeZoneStr = 'UTC+0'):
	# Converte a string para um objeto datetime
    dt = datetime.strptime(dt_str, "%Y-%m-%d")
    
    # Define o fuso horário especificado pelo offset em horas
    offset_hours = VALID_TIMEZONES[tz_str]
    offset = timezone(timedelta(hours=offset_hours))
    dt = dt.replace(tzinfo=offset)
    
    # Converte o datetime para UTC
    dt_utc = dt.astimezone(timezone.utc)
    
    # Define a época (1970-01-01 00:00:00 UTC)
    epoch = datetime(1970, 1, 1, tzinfo=timezone.utc)
    
    # Calcula a diferença em milissegundos
    return int((dt_utc - epoch).total_seconds() * 1000)

def consultar_empresa_por_razao_social(razao_social):
	"""
	Procura uma empresa por razão social.
	"""
	TANGERINO_API_BASE_URL = os.getenv('TANGERINO_API_BASE_URL')
	TANGERINO_API_TOKEN = os.getenv('TANGERINO_API_TOKEN')

	headers = {
		'Authorization': f'Basic {TANGERINO_API_TOKEN}',
		'Content-Type': 'application/json'
	}

	try:
		response = requests.get(f"{TANGERINO_API_BASE_URL}/employer/companies/filter?socialReason={razao_social}", headers=headers)

		if response.status_code == 200:
			empresas_data = response.json()[0]
			return empresas_data
		else:
			return None
	except Exception as e:
		print(e)
		input()

def baixar_folha_de_ponto(empresa: dict, data_inicial, data_final):
	"""
	Baixa uma folha de ponto para uma empresa.
	"""
	try:
		TANGERINO_API_BASE_URL = os.getenv('TANGERINO_API_BASE_URL')
		TANGERINO_API_TOKEN = os.getenv('TANGERINO_API_TOKEN')

		headers = {
			'Authorization': f'Basic {TANGERINO_API_TOKEN}',
			'Content-Type': 'application/json'
		}

		startDate = unix_time_millis(data_inicial, 'UTC-3')
		endDate = unix_time_millis(data_final, 'UTC-3')

		data_inicial = datetime.strptime(data_inicial, "%Y-%m-%d")
		data_final = datetime.strptime(data_final, "%Y-%m-%d")
		response = requests.get(f"{TANGERINO_API_BASE_URL}/report/time-sheet?companyId={empresa['id']}&employeeStatus=ADMITIDOS&startDate={startDate}&endDate={endDate}&format=PDF", headers=headers)

		if response.status_code == 200:
			response_data = response.json()
			file_content = base64.b64decode(response_data['base64FileContent'])
			with open(f"{get_download_path()}\\Folha de Ponto - {empresa['socialReason'].replace("S/S", "S S")} {data_inicial.strftime('%d-%m-%Y')} - {data_final.strftime('%d-%m-%Y')}.{response_data['fileExtension']}", 'wb') as f:
				f.write(file_content)
				return True
		else:
			print(f"Erro ao baixar folha de ponto: {response.status_code} - {response.text}")
			return False
	except Exception as e:
		print(e)
		input()
		return False

# Função para formatar o corpo do e-mail
def format_email_body(body: str) -> str:
		# Dividir os parágrafos em cada ponto seguido por um espaço
		paragrafos = body.strip().split('\n')
		
		# Adicionar tabulação no início de cada parágrafo
		paragrafos_com_tabulacao = [f"\t{paragrafo.strip()}" for paragrafo in paragrafos if paragrafo.strip()]

		# Adicionar quebra de linha dupla entre parágrafos
		email_body_formatado = "\n\n".join(paragrafos_com_tabulacao)

		return email_body_formatado

def gerar_folha(start_date: str, end_date: str):
	query_procura_clientes_folha = ler_sql('sql/procura_clientes_folha_ponto.sql')
	try:
		with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
			print(f"Conectado ao MySQL? {conn.is_connected()}")
			cursor.execute(query_procura_clientes_folha)
			clientes = cursor.fetchall()
			column_names = [desc[0] for desc in cursor.description]  # Obtém os nomes das colunas
			conn.commit()
			conn.close()
		if clientes:
			quantidade_sucessos = 0
			clientes_dict = process_clientes(clientes, column_names) # Converte cada linha para um dicionário e converte datetimes
			for cliente in clientes_dict:
				empresa_cliente = consultar_empresa_por_razao_social(cliente['nome_razao_social'])

				if not empresa_cliente or empresa_cliente == None:
					print(f"Cliente [{cliente['nome_razao_social']}] não encontrado na plataforma Tangerino RH")

					if (clientes_dict.index(cliente) + 1) < len(clientes_dict):
						input("Pressione enter para ir para o proximo cliente")
						continue

				if empresa_cliente:
					sucesso = baixar_folha_de_ponto(empresa_cliente, start_date, end_date)

					if sucesso:
						print(f"Sucesso ao baixar folha de ponto para o cliente [{cliente['nome_razao_social']}]")

						arquivos_download = listagem_arquivos_downloads()
						arquivo_mais_recente = max(arquivos_download, key=os.path.getmtime)

						anexos.append(arquivo_mais_recente)

						start_date = datetime.strptime(start_date, '%Y-%m-%d').strftime("%d/%m/%Y")
						end_date = datetime.strptime(end_date, '%Y-%m-%d').strftime("%d/%m/%Y")
						nome_razao_social = empresa_cliente['socialReason']

						email_body = format_email_body(f"""
												Gostaríamos de informar que a folha de ponto referente a {start_date} - {end_date} foi gerada com sucesso e está disponível para análise e eventual correção, caso necessário.
												Por favor, acesse {os.path.basename(arquivo_mais_recente)} para visualizar e verificar as informações registradas. 
												Caso identifique qualquer inconsistência ou discrepância em seu registro, por gentiliza entre em contato imediatamente.
												Salientamos a importância da verificação cuidadosa dos registros de ponto, a fim de garantir a precisão e integridade das informações relacionadas à jornada de trabalho da sua empresa.
												Agradecemos antecipadamente pela sua atenção e colaboração neste processo.
												""")
						enviar_email_com_anexos("faleconoscorj@human.rh.com.br", f"Folha de Ponto - {nome_razao_social}", email_body, anexos)
						os.remove(arquivo_mais_recente)
						quantidade_sucessos += 1
				
					if (clientes_dict.index(cliente) + 1) < len(clientes_dict):
						print(f"\n{clientes_dict.index(cliente) + 1} / {len(clientes_dict)}")

				if (clientes_dict.index(cliente) + 1) == len(clientes_dict):
					print(f"\n{quantidade_sucessos} Sucesso(s) de {len(clientes_dict)}")
					input("Pressione enter para sair")
					return True
		else:
			return None
	except Exception as e:
		print(''.join(traceback.format_exception(etype=type(e), value=e, tb=e.__traceback__)))
		input()
  
def lambda_handler(event, context):
	try:
		start_date = event['Robo Folha de Ponto - Data Inicial']
		end_date = event['Robo Folha de Ponto - Data Final']
		
		if not start_date or not end_date:
			return {'statusCode': 400, 'body': json.dumps({"message": 'Data Inicial ou Data Final ausente'})}
		try:
			print(f"Data Inicial: {start_date}")
			print(f"Data Final: {end_date}")

			# Cliente do Lambda
			"""lambda_client = boto3.client('lambda')
			
			response = lambda_client.invoke(
		        FunctionName='human_activate_rds_proxy',
		        InvocationType='RequestResponse',  # Pode ser 'Event' para execução assíncrona
		    )
			print(response)"""
			
			sucesso = gerar_folha(start_date, end_date)
			
			if sucesso:
				return {'statusCode': 200, 'body': json.dumps({"message": 'Folha de ponto gerada com sucesso'})}
			else:
				return {'statusCode': 500, 'body': json.dumps({"message": 'Erro ao gerar folha de ponto'})}
		except Exception as e:
			return {'statusCode': 500, 'body': json.dumps({"message": f'Erro ao processar a requisição: {e}'})}
		"""finally:
			response = lambda_client.invoke(
			    FunctionName='human_deactivate_rds_proxy',
			    InvocationType='RequestResponse'
			)
			print(response)"""
	except Exception as e:
		print(''.join(traceback.format_exception(etype=type(e), value=e, tb=e.__traceback__)))
		return {'statusCode': 500, 'body': json.dumps({"message": f'Erro ao iniciar o processamento da requisição: {e}'})}