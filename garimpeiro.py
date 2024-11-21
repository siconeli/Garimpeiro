import pandas as pd
import os

# Definir os grupos de códigos [ EXERCÍCIO 2022 a 2024 ]
grupos_de_codigos = [
    ['1112.50.0.1.00.00.00', '1112.50.0.2.00.00.00'],  # IPTU
    ['1112.53.0.1.00.00.00', '1112.53.0.2.00.00.00'], # ITBI
    ['1113.03.1.1.00.00.00', '1113.03.4.3.00.00.00', '1113.03.4.1.00.00.00'], # IRRF
    ['1114.51.1.1.00.00.00', '1114.51.1.2.00.00.00'], # ISSQN
    ['1121.01.0.1.00.00.00', '1121.01.0.2.00.00.00'], # ALVARÁ
    ['1122.01.0.1.00.00.00', '1122.01.0.2.00.00.00'], # TAXA DE LIXO
    ['1240.00.0.0.00.00.00'], # COSIP
    ['1711.52.0.1.00.00.00'], # ITR
    ['1112.50.0.3.00.00.00', '1112.50.0.4.00.00.00', '1112.53.0.3.00.00.00', '1112.53.0.4.00.00.00', 
     '1114.51.1.3.00.00.00', '1114.51.1.4.00.00.00', '1121.01.0.3.00.00.00', '1121.01.0.4.00.00.00', 
     '1122.01.0.3.00.00.00', '1122.01.0.4.00.00.00']  # DÍVIDA ATIVA
]

# Definir os grupos de códigos [ EXERCÍCIO 2020 ]
# grupos_de_codigos = [
#     ['1118.01.1.1.00.', '1118.01.1.2.00'],  # IPTU
#     ['1118.01.4.1.00'], # ITBI
#     ['1113.03.1.1.00'], # IRRF
#     ['1118.02.3.1.00', '1118.02.3.2.00'], # ISSQN
#     ['1121.01.1.1.00', '1121.01.1.2.00', '1128.01.1.1.00', '1128.01.1.2.00'], # ALVARÁ
#     ['1122.01.1.1.00', '1122.01.1.2.00'], # TAXA DE LIXO
#     ['1240.00.1.1.00'], # COSIP
#     ['1718.01.5.1.00'], # ITR
#     ['1113.03.1.3.00', '1113.03.1.4.00', '1118.01.1.3.00', '1118.01.1.4.00', '1118.02.3.3.00', '1118.02.3.4.00', '1121.01.1.3.00', '1121.01.1.4.00', '1122.01.1.3.00', '1122.01.1.4.00']  # DÍVIDA ATIVA
# ]


# Definir os grupos de códigos [ EXERCÍCIO 2021 ]
# grupos_de_codigos = [
#     ['1118.01.1.1.00.00.00', '1118.01.1.2.00.00.00'],  # IPTU
#     ['1118.01.4.1.00.00.00'], # ITBI
#     ['1113.03.1.1.00.00.00'], # IRRF
#     ['1118.02.3.1.00.00.00', '1118.02.3.2.00.00.00'], # ISSQN
#     ['1121.01.1.1.00.00.00', '1121.01.1.2.00.00.00', '1128.01.1.1.00.00.00', '1128.01.1.2.00.00.00'], # ALVARÁ
#     ['1122.01.1.1.00.00.00', '1122.01.1.2.00.00.00'], # TAXA DE LIXO
#     ['1240.00.1.1.00.00.00'], # COSIP
#     ['1718.01.5.1.00.00.00'], # ITR
#     ['1113.03.1.3.00.00.00', '1113.03.1.4.00.00.00', '1118.01.1.3.00.00.00', '1118.01.1.4.00.00.00', '1118.02.3.3.00.00.00', '1118.02.3.4.00.00.00', '1121.01.1.3.00.00.00', '1121.01.1.4.00.00.00', '1122.01.1.3.00.00.00', '1122.01.1.4.00.00.00']  # DÍVIDA ATIVA
# ]



# Mapeamento de meses para números para garantir a ordenação correta
meses = {
    "janeiro": 1,
    "fevereiro": 2,
    "março": 3,
    "abril": 4,
    "maio": 5,
    "junho": 6,
    "julho": 7,
    "agosto": 8,
    "setembro": 9,
    "outubro": 10,
    "novembro": 11,
    "dezembro": 12
}

# Inicializa as listas para todos os meses, mas agora só adiciona os tributos fixos uma vez
dados = {
    'tributo': ['IPTU', 'ITBI', 'IRRF', 'ISSQN', 'ALVARÁ', 'TX LIXO', 'COSIP', 'ITR', 'D.A'], 
    'janeiro': [None] * 9,
    'fevereiro': [None] * 9,
    'março': [None] * 9,
    'abril': [None] * 9,
    'maio': [None] * 9,
    'junho': [None] * 9,
    'julho': [None] * 9,
    'agosto': [None] * 9,
    'setembro': [None] * 9,
    'outubro': [None] * 9,
    'novembro': [None] * 9,
    'dezembro': [None] * 9,
}

# Diretório onde os arquivos Excel estão localizados
diretorio_arquivos = r"C:\Users\luisf\OneDrive\Área de Trabalho\DADOS -TRIBUTOS\DADOS MENSAL\Sidrolândia\2023"

ano = 2023

# Obter todos os arquivos do diretório
arquivos = [f for f in os.listdir(diretorio_arquivos) if f.endswith('.xls') or f.endswith('.xlsx')]

# Ordenar os arquivos de acordo com o mês, usando o mapeamento de meses
arquivos_ordenados = sorted(arquivos, key=lambda x: meses.get(x.split('.')[0].lower(), 0))

# Iterar sobre cada arquivo Excel no diretório
for nome_arquivo in arquivos_ordenados:
    # Carregar a planilha para o DataFrame
    caminho_arquivo = os.path.join(diretorio_arquivos, nome_arquivo)
    df = pd.read_excel(caminho_arquivo, engine='openpyxl' if nome_arquivo.endswith('.xlsx') else 'xlrd')

    # Substituir NaN por 0 na coluna 'Arrec. Período'
    df['Arrec. Período'] = df['Arrec. Período'].fillna(0)

    # Adicionar uma coluna com o nome do mês (extraído do nome do arquivo)
    mes = nome_arquivo.split('.')[0].lower()  # Assumindo que o nome do arquivo é o mês (ex: janeiro.xlsx)

    # Verificar se o mês está no mapeamento de meses
    if mes in meses:
        for i, grupo in enumerate(grupos_de_codigos):
            # Filtrar o DataFrame para o grupo de códigos
            df_filtrado = df[df['Código'].isin(grupo)]
            print(df_filtrado)

            # Inicializar a lista para armazenar os maiores valores
            maiores_valores = []

            # Encontrar o maior valor de 'Arrec. Período' para cada código
            for codigo, grupo_codigo in df_filtrado.groupby('Código'):
                maior_valor_por_codigo = grupo_codigo.loc[grupo_codigo['Arrec. Período'].idxmax()]
                maiores_valores.append(maior_valor_por_codigo)

            # Criar o DataFrame com os maiores valores
            df_maiores_valores = pd.DataFrame(maiores_valores)

            # Somar os maiores valores de 'Arrec. Período' do grupo
            total_do_grupo = df_maiores_valores['Arrec. Período'].sum()

            # Adicionar o valor ao dicionário, preenchendo o valor no mês correto
            dados[mes][i] = total_do_grupo  # Preenche a lista do mês com o total calculado

# Criar o DataFrame
dados_df = pd.DataFrame(dados)

# Salvar o DataFrame final em um arquivo Excel
dados_df.to_excel(f"dados_garimpados_{ano}.xlsx", index=False, engine="openpyxl")
