"""IMPORTAÇÃO DE BIBLIOTECA PARA MEXER COM DIRETÓRIOS E ARQUIVOS NO WINDOWS"""
import os
from pathlib import Path

def listagem_pastas(diretorio):
    lista_pastas = []
    pastas = Path(diretorio)
    for pasta in pastas.iterdir():
        lista_pastas.append(f"{pasta}")
    return lista_pastas 