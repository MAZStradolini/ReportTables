import os
import pandas as pd
import re
import Dados.DictionaryRT

# Solicita ao usuário que insira o caminho para a pasta na rede
path = input("Por favor, insira o caminho para a pasta na rede: ")

# Verifica se o caminho existe
if not os.path.exists(path):
    print(f"O caminho '{path}' não existe. Verifique o caminho e tente novamente.")
else:
    # Inicializa lista para armazenar resultados
    resultados = []

    # Loop para ler todos os arquivos .xlsx na pasta
    for filename in os.listdir(path):
        if filename.endswith(".xlsx"):
            # Leitura do arquivo
            df = pd.read_excel(os.path.join(path, filename), sheet_name=1, skiprows=0)

            for index, row in df.iterrows():
                if row.iloc[0] in Dados.DictionaryRT.Services: 
                    secretaria_encontrada = 'não identificado'
                    for secretaria_possivel in Dados.DictionaryRT.Setores:
                        if secretaria_possivel in str(df.iloc[:]):
                            secretaria_encontrada = secretaria_possivel

            for index, row in df.iterrows():
                if row.iloc[0] in Dados.DictionaryRT.Services:  
                    mes_competencia_encontrada = 'não identificado'
                    for mes_competencia_possivel in Dados.DictionaryRT.mes_competencia:
                        if mes_competencia_possivel in str(df.iloc[0:0]):
                            mes_competencia_encontrada = mes_competencia_possivel

            for index, row in df.iterrows():
                if row.iloc[0] in Dados.DictionaryRT.Services:  
                    ano_competencia_encontrada = 'não identificado'
                    for ano_competencia_possivel in Dados.DictionaryRT.ano_competencia:
                        if ano_competencia_possivel in str(df.iloc[0:0]):
                            ano_competencia_encontrada = ano_competencia_possivel

            for index, row in df.iterrows():
                if row.iloc[0] in Dados.DictionaryRT.Services: 
                    sei = 'não identificado'
                    for texto in df.iloc[:]:
                        # Expressão regular para buscar a string no formato 00.00.000000000-0
                        match = re.search(r"\d{2}\.\d{2}\.\d{9}-\d", str(texto))
                        if match:
                            sei = match.group()
                            break     
                    
              # Loop para buscar objeto na coluna 0
            for index, row in df.iterrows():
                if row.iloc[0] in Dados.DictionaryRT.Services:  
                    resultados.append({'Nome do Arquivo Original': filename,
                                       'Secretaria': secretaria_encontrada,
                                       'SEI': sei,
                                       'Mês': mes_competencia_encontrada,
                                       'Ano': ano_competencia_encontrada,
                                       'Serviços': row.iloc[0],  
                                       'Vlr Total': row.iloc[3],  
                                       'Vlr Faturado': row.iloc[6]})  

    # Cria dataframe com resultados
    df_resultados = pd.DataFrame(resultados)

    # Exibe resultados
    print(df_resultados)
    df_resultados.to_excel('consolidado.xlsx', index=False)

