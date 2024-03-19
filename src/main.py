import pandas as pd 
import os
import glob

# Criando caminho para ler os arquivos
folder_path = 'src\\data\\raw'

# lista todos os arquivos de excel
excel_files = glob.glob(os.path.join(folder_path , '*.xlsx'))

if not excel_files:
    print('Nenhum arquivo compativel encontrado')
else:

# Dataframe = tabela na memória para guardar os conteudos dos arquivos
    dfs = []

for excel_file in excel_files:


    try:
        # Leio o arquivo de excel 
        df_temp = pd.read_excel(excel_file)
        

        #pegar o nome dos arquivos
        file_name = os.path.basename(excel_file)

        # Criando mais uma coluna que não existe no excel e dando nome a ela
        if 'brasil' in file_name.lower():
            df_temp['Location'] = 'br'

        elif 'france' in file_name.lower():
            df_temp['Location'] = 'fr'

        elif 'italian' in file_name.lower():
            df_temp['Location'] = 'it'

        #criando uma nova coluna chamada campaing
        df_temp['Campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

        #criando mais uma coluna
        df_temp['name_arquiv'] = file_name

        df_temp['plan'] = df_temp['Contracted Plan'].str.extract(r'Plano(.*)')

        # Guarda dados tratados dentro de uma dataframe comum
        dfs.append(df_temp)
        
    except Exception as e:
        print(f'Erro ao ler o arquivo {excel_files} : {e}') 


if dfs: 
    
   
    # Concatenando todas as tabelas salvas no dfs em uma unica tabela
    result = pd.concat(dfs, ignore_index=True)

    # Caminho de saida
    output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')

    # Configurar motor de escrita
    write = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # Levar os dados do resultado a serem escritos no motor de excel configurado
    result.to_excel(write, index=False)

    # salva o arquivo de excel
    write._save()

else:
    print('Nenhum dado para ser salvo')

