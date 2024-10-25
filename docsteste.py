from docx import Document
from typing import Iterator
from docx.table import _Row
from docx.shared import Inches
import pandas as pd
import csv
import re
from typing import List, Tuple

# Função 1: Remover pontuações dos conteúdos de cada célula
def remove_punctuation(table: List[List[str]]) -> List[List[str]]:
    # Usar regex para remover pontuações
    table_no_punctuation = [[re.sub(r'[^\w\s]', '', cell) for cell in row] for row in table]
    return table_no_punctuation

# Função 2: Dividir células mescladas e replicar conteúdo
def expand_merged_cells(table: List[List[str]]) -> List[List[str]]:
    expanded_table = []
    for row in table:
        expanded_row = []
        for cell in row:
            expanded_row.append(cell)
        expanded_table.append(expanded_row)
    return expanded_table

# Função 3: Converter para DataFrame e salvar como CSV
def save_to_csv(table: List[List[str]], file_name: str):
    df = pd.DataFrame(table[1:], columns=table[0])  # Usando a primeira linha como cabeçalho
    df.to_csv(file_name, index=False)
    print(f"Dados salvos em '{file_name}' com sucesso!")

document = Document("teste1.1.docx")#Abrir o documento
table = document.tables[1]#buscar a tabela
extracted_table = [[cell.text for cell in row.cells] for row in table.rows]#transformar em uma tupla
extracted_table = [[cell.text.replace("\n", " ") for cell in row.cells] for row in table.rows]#Remover as quebras de linah


nome_arquivo = "tabela_extraída.csv"#definir o nome do arquivo

table_no_punctuation = remove_punctuation(extracted_table)#remover pontuações
expanded_table = expand_merged_cells(table_no_punctuation)#expandir linhas mescladas
save_to_csv(expanded_table, nome_arquivo )#salvar em um arquivo csv

