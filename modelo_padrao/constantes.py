import os
import sys
import re

PASTA_PROJETO = os.path.dirname(os.path.realpath(__file__))
PASTA_TEMP = os.path.join(PASTA_PROJETO, 'temp')
PASTA_INPUT = os.path.join(PASTA_PROJETO, 'input')
PASTA_OUTPUT = os.path.join(PASTA_PROJETO, 'output')
PASTA_LOGS = os.path.join(PASTA_PROJETO, 'logs')
CAMINHO_PASTA_LOGS = os.path.join(PASTA_LOGS, 'log.txt')
CAMINHO_ARQUIVO_OUTPUT = os.path.join(PASTA_OUTPUT, 'resultado.csv')
ARQUIVO = os.path.join(PASTA_INPUT, 'Petição.pdf')

VALOR_CAUSA = r'Valor da causa: R\$ ([\d.,]+)'
PARTE_CREDORA = r'Parte Credora:\s*(.+)'
DOCUMENTO_CREDOR = r'Parte Credora:\s*(.+)\s*CPF/CNPJ:\s*((?:\d{3}\.\d{3}\.\d{3}-\d{2})|(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}))'
ADVOGADO = r'Advogado\(s\):\s*(.+)'
DOCUMENTO_ADV = r'Advogado\(s\):\s*(.+)\s*CPF/CNPJ:\s*((?:\d{3}\.\d{3}\.\d{3}-\d{2})|(?:\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}))'
OAB = r'OAB:\s*\d\d.\d\d\d'
NATUREZA = r'Natureza.+Alimentar.+\(.+X.+\).+Patrimonial'
VL_PRINCIPAL = r'Valor Principal:\s*([^:]+)Juros:'
VL_JUROS= r'Juros:\s*([^:]+)Índices/taxa Selic:'
DATA_BASE = r'Data base utilizada para os cálculos:\s*(\d{2}\.\d{2}\.\d{4})'
TOTAL = r'Total\s+\s*([^:]+)DADOS COMPLEMENTARES'
HONORARIO = r'Honorários Contratuais:\s%\sValor.+\s.+\s.+\s.+\s.+'

valor_causa = []
credor = []
adv = []
documento = []        
doc_adv = []
adv_oab = []
natureza_do_precatorio = []
valor_principal = []
valor_juros = []
data_base = []
total = []
honorario = []
porcentagem_honorario = []
valor_honorario = []