import os
import xmltodict
import pandas as pd
import json
import openpyxl

# Função para ler informações de um arquivo XML e extrair dados específicos
def ler_informacoes(nome_arquivo, valores):
    print(f'pegou as informações {nome_arquivo}')
    try:
        # Abre o arquivo XML para leitura
        # Inserir o local das CTes aqui
        with open(f'ctes/{nome_arquivo}', 'rb') as arquivo_xml:
            # Converte o conteúdo XML para um dicionário Python
            dicionario_arquivo = xmltodict.parse(arquivo_xml)
            # Imprime o dicionário para depuração
            print(json.dumps(dicionario_arquivo, indent=4))
            # Acessa a seção específica do XML ( Personalizável )
            # Exemplos: chave de acesso e chave complementar da CTe
            if "cteProc" in dicionario_arquivo:
                infos_nf = dicionario_arquivo["cteProc"]
            chave_acesso = infos_nf["CTe"]["infCte"]["infCteComp"]
            chave_cte = infos_nf["CTe"]["infCte"]["@Id"]
            # Adiciona os valores extraídos à lista de valores
            valores.append([chave_acesso, chave_cte])
            return chave_acesso
    except Exception as e:
        # Tratamento de erro caso ocorra algum problema ao processar o arquivo
        print(f'Erro ao processar {nome_arquivo}: {e}')
        return None

# Lista todos os arquivos no diretório 'ctes'
lista_arquivos = os.listdir('ctes')

# Define as colunas para o DataFrame
colunas = ['Chave CTe', 'Chave CTe Complementar']
valores = []
# Itera sobre cada arquivo na lista de arquivos
for arquivo in lista_arquivos:
    ler_informacoes(arquivo, valores) # Chama a função para ler e extrair informações do arquivo


# Cria uma tabela com os dados extraídos
tabela = pd.DataFrame(columns=colunas, data=valores)
# Salva a tabela em uma planilha Excel
tabela.to_excel('CTes.xlsx', index=False)

