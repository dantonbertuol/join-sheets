import pandas as pd
import datetime
import warnings

warnings.simplefilter("ignore")

# Data do primeiro arquivo dentro da pasta - a leitura inicia a partir deste
data_inicial = '08-11-2021'
# Data final sempre o dia de hoje pra ler todos os arquivos da pasta
data_final = datetime.date.strftime(datetime.datetime.now(), '%d-%m-%Y')

# Converte as datas para datetime
data_inicial_dt = datetime.datetime.strptime(data_inicial, '%d-%m-%Y')
data_final_dt = datetime.datetime.strptime(data_final, '%d-%m-%Y')

# Cria um dataframe vazio para armazenar os dados consolidados
final_sheet = pd.DataFrame()

# Loop para percorrer todos os itens
while data_inicial_dt < data_final_dt:
    data_inicial = datetime.date.strftime(data_inicial_dt, '%d-%m-%Y')
    caminho = f"C:/Users/297754/OneDrive/Unoesc/Desenvolvimento/BI/Power BI/Unoesc/Utilização BI/{data_inicial}.xlsx"

    # Caso não exista o arquivo da data especificada pula para a próxima
    try:
        exc = pd.read_excel(caminho)
    except FileNotFoundError:
        data_inicial_dt += datetime.timedelta(days=1)
        continue

    # Remove as duas últimas linhas pois são informações não utilizadas
    exc = exc.drop([len(exc)-1, len(exc)-2])

    # Adiciona a coluna data, que é a data do nome do arquivo
    exc['Data'] = data_inicial

    # Concatena o consolidado com o arquivo que está lendo
    final_sheet = pd.concat([final_sheet, exc], ignore_index=True)

    # Incrementa a data
    data_inicial_dt += datetime.timedelta(days=1)

# nome_arquivo = 'Consolidado ' + datetime.date.strftime(datetime.datetime.now(), '%d-%m-%Y' + '.xlsx')
nome_arquivo = 'Consolidado.xlsx'
caminho_arquivo = 'C:/Users/297754/OneDrive/Unoesc/Desenvolvimento/BI/Power BI/Unoesc/Utilização BI/'

# Salva o consolidado em um arquivo de excel
final_sheet.to_excel(caminho_arquivo + nome_arquivo,
                     index=False, sheet_name='Consolidado')

print('Processo Concluído.')
