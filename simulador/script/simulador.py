import os
import pandas as pd # pandas  2.2.2
import datetime
import numpy as np

count = 0

#########################################################################################################################################################
# Painel de controlo
##################
lista_percentagem_minima_circulo_nacional = [(0, "0"), (0.01, "1"), (0.025, "2_5"), (0.05, "5")]
##################
lista_ano = [2015, 2019, 2022, 2024, 2025]
##################
# (numero_cirulos_uninominais, numero_mandatos_circulo_nacional, X)
lista_numero_circulos_mandatos = [(100, 126, "A"), (113, 113, "B"), (126, 100, "C")]
#########################################################################################################################################################
def calcular_votos_usados_e_transferidos(df_votos_agregados_circulo_eleitoral):
    # Cria um DataFrame para armazenar os votos usados no sistema uninominal
    votos_usados_uninominal_df = pd.DataFrame(index=df_votos_agregados_circulo_eleitoral.index, columns=df_votos_agregados_circulo_eleitoral.columns)
    # Cria um DataFrame para armazenar os votos transferidos 
    votos_transferidos_para_compensacao_df = pd.DataFrame(index=df_votos_agregados_circulo_eleitoral.index, columns=df_votos_agregados_circulo_eleitoral.columns)

    # Itera sobre cada círculo eleitoral
    for ciruclo_eleitoral in df_votos_agregados_circulo_eleitoral.index:
        # Obtém os votos do círculo eleitoral atual
        votos_circulo_eleitoral = df_votos_agregados_circulo_eleitoral.loc[ciruclo_eleitoral]

        # Ordena os partidos pelos votos para determinar o vencedor e o segundo classificado
        partidos_ordenados = votos_circulo_eleitoral.sort_values(ascending=False)
        partido_primeiro_classificado = partidos_ordenados.index[0]  # Primeiro classificado
        votos_primeiro_classificado = partidos_ordenados.iloc[0]  # Votos do primeiro classificado
        votos_segundo_classificado = partidos_ordenados.iloc[1] if len(partidos_ordenados) > 1 else 0  # Votos do segundo classificado

        # Votos usados pelo partido vencedor no sistema uninominal: um a mais que o segundo classificado
        votos_usados_primeiro_classificado = votos_segundo_classificado + 1 if len(partidos_ordenados) > 1 else votos_primeiro_classificado

        # Define os votos usados para o vencedor e os votos transferidos para todos os partidos
        votos_usados_uninominal_df.loc[ciruclo_eleitoral] = 0  # Inicializa os votos usados como zero
        votos_usados_uninominal_df.at[ciruclo_eleitoral, partido_primeiro_classificado] = votos_usados_primeiro_classificado  # Define os votos usados pelo vencedor

        # Calcula os votos transferidos para compensação
        votos_transferidos_para_compensacao_df.loc[ciruclo_eleitoral] = votos_circulo_eleitoral - votos_usados_uninominal_df.loc[ciruclo_eleitoral]

    # Retorna os DataFrames de votos usados no sistema uninominal e de votos transferidos
    return votos_usados_uninominal_df, votos_transferidos_para_compensacao_df

#########################################################################################################################################################

def alocar_deputados_metodo_hondt(df_votos, numero_de_lugares_por_circulo):

    votos_totais = df_votos.sum()

    # Dicionário "lugares", com partidos/número de deputados. o número de deputados é igual a 0, antes do início do processo de alocação. 
    lugares = {partido: 0 for partido in df_votos.columns}

    # Dicionário "quocientes", com partidos/quocientes. 
    quocientes = {partido: votos_totais[partido] / (lugares[partido] + 1) for partido in df_votos.columns}

    votos_utilizados_metodo_hondt = {partido: [] for partido in df_votos.columns}
    for _ in range(numero_de_lugares_por_circulo): 
        # Seleciona o partido com o quociente mais alto
        partido_maximo_quociente = max(quocientes, key=quocientes.get)
        # Atribui um deputado a esse partido, com o quociente máximo
        lugares[partido_maximo_quociente] += 1

        # Votos utilizados na alocação
        votos_para_alocacao = votos_totais[partido_maximo_quociente] / lugares[partido_maximo_quociente]

        # Atualiza o dicionário
        votos_utilizados_metodo_hondt[partido_maximo_quociente].append(round(votos_para_alocacao, 2))

        # Atualiza o quociente do partido com o quociente máximo atual
        quocientes[partido_maximo_quociente] = votos_totais[partido_maximo_quociente] / (lugares[partido_maximo_quociente] + 1)

    # Para garantir que todas as listas têm o mesmo tamanho
    tamanho_maximo = max(len(votos) for votos in votos_utilizados_metodo_hondt.values())
    for votos in votos_utilizados_metodo_hondt.values():
        votos.extend([0] * (tamanho_maximo - len(votos)))

    # Converter os dicionários em DataFrames
    df_lugares = pd.DataFrame([lugares], index=["Total lugares"])
    df_votos_utilizados_metodo_hondt = pd.DataFrame(votos_utilizados_metodo_hondt)
    # Ajustar o índice, para começar em 1, em vez de 0
    df_votos_utilizados_metodo_hondt.index += 1

    return df_lugares, df_votos_utilizados_metodo_hondt

#########################################################################################################################################################

def ordenar_linhas_colunas_e_adicionar_totais_df(df):

    # Ordenar alfabeticamente as linhas
    df.sort_index(axis=0, inplace=True)

    # Ordenar alfabeticamente as colunas 
    df.sort_index(axis=1, inplace=True) 


    # Adicionar os totais
    df['Total'] = df.sum(axis=1)
    df.loc['Total'] = df.sum(axis=0)
    return df

#########################################################################################################################################################

def ordenar_linhas_e_adicionar_totais_df(df):

    # Ordenar alfabeticamente as linhas
    df.sort_index(axis=0, inplace=True)

    # Adicionar os totais
    df['Total'] = df.sum(axis=1)
    df.loc['Total'] = df.sum(axis=0)
    return df

#########################################################################################################################################################

def ordenar_colunas_linhas_df(df):

    # Ordenar alfabeticamente as linhas
    df.sort_index(axis=0, inplace=True)

    # Ordenar alfabeticamente as colunas 
    df.sort_index(axis=1, inplace=True) 
    return df

#########################################################################################################################################################


for percentagem_minima_circulo_nacional, limite_per in lista_percentagem_minima_circulo_nacional:
    for ano in lista_ano:
        for numero_cirulos_uninominais, numero_mandatos_circulo_nacional, cenario in lista_numero_circulos_mandatos:

#########################################################################################################################################################

            # Dicionários com o número de deputados, por círculo atual (varia de eleição para eleição, em função do número de eleitores)
            if ano == 2015:
                dicionario_sistema_atual={
                'Aveiro': 16,
                'Beja': 3,
                'Braga': 19,
                'Bragança': 3,
                'Castelo Branco': 4,
                'Coimbra': 9,
                'Évora': 3,
                'Faro': 9,
                'Guarda': 4,
                'Leiria': 10,
                'Lisboa': 47,
                'Portalegre': 2,
                'Porto': 39,
                'Santarém': 9,
                'Setúbal': 18,
                'Viana do Castelo': 6,
                'Vila Real': 5,
                'Viseu': 9,
                'Açores': 5,
                'Madeira': 6,
                'Europa': 2,
                'Fora da Europa': 2,
                }

            elif ano in [2019, 2022]:
                dicionario_sistema_atual={
                'Aveiro': 16,
                'Beja': 3,
                'Braga': 19,
                'Bragança': 3,
                'Castelo Branco': 4,
                'Coimbra': 9,
                'Évora': 3,
                'Faro': 9,
                'Guarda': 3,
                'Leiria': 10,
                'Lisboa': 48,
                'Portalegre': 2,
                'Porto': 40,
                'Santarém': 9,
                'Setúbal': 18,
                'Viana do Castelo': 6,
                'Vila Real': 5,
                'Viseu': 8,
                'Açores': 5,
                'Madeira': 6,
                'Europa': 2,
                'Fora da Europa': 2,
                }

            elif ano == 2024:
                dicionario_sistema_atual={
                'Aveiro': 16,
                'Beja': 3,
                'Braga': 19,
                'Bragança': 3,
                'Castelo Branco': 4,
                'Coimbra': 9,
                'Évora': 3,
                'Faro': 9,
                'Guarda': 3,
                'Leiria': 10,
                'Lisboa': 48,
                'Portalegre': 2,
                'Porto': 40,
                'Santarém': 9,
                'Setúbal': 19,
                'Viana do Castelo': 5,
                'Vila Real': 5,
                'Viseu': 8,
                'Açores': 5,
                'Madeira': 6,
                'Europa': 2,
                'Fora da Europa': 2,
                }

            elif ano == 2025:
                dicionario_sistema_atual={
                'Aveiro': 16,
                'Beja': 3,
                'Braga': 19,
                'Bragança': 3,
                'Castelo Branco': 4,
                'Coimbra': 9,
                'Évora': 3,
                'Faro': 9,
                'Guarda': 3,
                'Leiria': 10,
                'Lisboa': 48,
                'Portalegre': 2,
                'Porto': 40,
                'Santarém': 9,
                'Setúbal': 19,
                'Viana do Castelo': 5,
                'Vila Real': 5,
                'Viseu': 8,
                'Açores': 5,
                'Madeira': 6,
                'Europa': 2,
                'Fora da Europa': 2,
                }

#########################################################################################################################################################

            caminho_ficheiro_input = f"../input/{ano}/dados_{ano}_{numero_cirulos_uninominais}_uninominais.xlsx"

####################################################################################################################################################

            df_original = pd.read_excel(caminho_ficheiro_input, sheet_name="Sheet1")

#########################################################################################################################################################

            df_territorio_nacional = df_original[df_original['tipo_territorio'] == 'Territorio nacional']

            partido_colunas_territorio_nacional = [col for col in df_territorio_nacional.columns if col.startswith('partido_')]

            df_territorio_nacional_votos_agregados = df_territorio_nacional.groupby('circulo_uninominal')[partido_colunas_territorio_nacional].sum()

            df_territorio_nacional_votos_agregados.columns = [col.replace('partido_', '') if col.startswith('partido_') else col for col in df_territorio_nacional_votos_agregados.columns]
#########################################################################################################################################################

            # Para determinar quais os partidos que têm menos que x% de votos no total do território nacional)

            # Passo 1: Calcular o total de votos para cada partido
            total_por_partido = df_territorio_nacional_votos_agregados.sum()

            # Passo 2: Calcular o total  de votos
            total_votos = total_por_partido.sum()

            # Passo 3: Calcular a percentagem_minima_circulo_nacional no total de votos, em termos de votos
            limiar =  percentagem_minima_circulo_nacional * total_votos

            # Passo 4: Determinar quais os partidos que têm votos totais >= limiar
            lista_partidos_acima_do_limiar = total_por_partido[total_por_partido >= limiar].index.tolist()

#########################################################################################################################################################

            df_territorio_estrangeiro = df_original[df_original['tipo_territorio'] == 'Estrangeiro']

            partido_colunas_territorio_estrangeiro = [col for col in df_territorio_estrangeiro.columns if col.startswith('partido_')]

            df_territorio_estrangeiro_votos_agregados = df_territorio_estrangeiro.groupby('circulo_atual')[partido_colunas_territorio_estrangeiro].sum()

            df_territorio_estrangeiro_votos_agregados.columns = [col.replace('partido_', '') if col.startswith('partido_') else col for col in df_territorio_estrangeiro_votos_agregados.columns]

#########################################################################################################################################################

            df_territorio_estrangeiro_europa = df_original[df_original['circulo_atual'] == 'Europa']

            partido_colunas_territorio_estrangeiro_europa = [col for col in df_territorio_estrangeiro_europa.columns if col.startswith('partido_')]

            df_territorio_estrangeiro_europa_votos_agregados = df_territorio_estrangeiro_europa.groupby('circulo_atual')[partido_colunas_territorio_estrangeiro_europa].sum()

            df_territorio_estrangeiro_europa_votos_agregados.columns = [col.replace('partido_', '') if col.startswith('partido_') else col for col in df_territorio_estrangeiro_europa_votos_agregados.columns]

#########################################################################################################################################################

            df_territorio_estrangeiro_fora_da_europa = df_original[df_original['circulo_atual'] == 'Fora da Europa']

            partido_colunas_territorio_estrangeiro_fora_da_europa = [col for col in df_territorio_estrangeiro_fora_da_europa.columns if col.startswith('partido_')]

            df_territorio_estrangeiro_fora_da_europa_votos_agregados = df_territorio_estrangeiro_fora_da_europa.groupby('circulo_atual')[partido_colunas_territorio_estrangeiro_fora_da_europa].sum()

            df_territorio_estrangeiro_fora_da_europa_votos_agregados.columns = [col.replace('partido_', '') if col.startswith('partido_') else col for col in df_territorio_estrangeiro_fora_da_europa_votos_agregados.columns]

#########################################################################################################################################################

            df_territorio_total_com_europa_e_fora_europa_sistema_atual = df_original.copy()

            partido_colunas_territorio_total_com_europa_e_fora_europa_sistema_atual = [col for col in df_territorio_total_com_europa_e_fora_europa_sistema_atual.columns if col.startswith('partido_')]

            df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados = df_territorio_total_com_europa_e_fora_europa_sistema_atual.groupby('circulo_atual')[partido_colunas_territorio_total_com_europa_e_fora_europa_sistema_atual].sum()

            df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.columns = [col.replace('partido_', '') if col.startswith('partido_') else col for col in df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.columns]

#########################################################################################################################################################

            df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados = df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.copy()
            # Excluir as linhas onde o índice é 'Europa' ou 'Fora da Europa'
            df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados = df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados[
                ~df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados.index.isin(['Europa', 'Fora da Europa'])
            ]

#########################################################################################################################################################

            df_votos_usados_uninominal_agregados, df_votos_transferidos_para_nacional_agregados = calcular_votos_usados_e_transferidos(df_territorio_nacional_votos_agregados)

            # Para limiar
            df_votos_transferidos_para_nacional_agregados_com_limiar = df_votos_transferidos_para_nacional_agregados[lista_partidos_acima_do_limiar]


            # Aplicação do método de Hondt
            # Para o círculo nacional de compensação
            df_deputados_alocados_nacional, df_quocientes_metodo_hondt_nacional = alocar_deputados_metodo_hondt(df_votos_transferidos_para_nacional_agregados_com_limiar, numero_mandatos_circulo_nacional)
            # Para o círculo Europa
            df_deputados_alocados_europa, df_quocientes_metodo_hondt_europa = alocar_deputados_metodo_hondt(df_territorio_estrangeiro_europa_votos_agregados, 2)
            # Para o círculo Fora da Europa
            df_deputados_alocados_fora_da_europa, df_quocientes_metodo_hondt_fora_da_europa = alocar_deputados_metodo_hondt(df_territorio_estrangeiro_fora_da_europa_votos_agregados, 2)

#########################################################################################################################################################

            # Número de deputados eleitos nos círculos uninominais, por círculo uninominal
            df_deputados_alocados_diretamente_circulos_uninominais = pd.DataFrame(0, index=df_votos_usados_uninominal_agregados.index, columns=df_votos_usados_uninominal_agregados.columns)

            # Deputados eleitos diretamente pelos círculos uninominais
            for circulo in df_votos_usados_uninominal_agregados.index:
                for party in df_votos_usados_uninominal_agregados.columns:
                    if df_votos_usados_uninominal_agregados.at[circulo, party] > 0:
                        df_deputados_alocados_diretamente_circulos_uninominais.at[circulo, party] = 1

#########################################################################################################################################################

            # Número de deputados, no sistema atual, por círculo atual

            # Com Europa e Fora da Europa 
            df_deputados_com_europa_e_fora_europa_sistema_atual = pd.DataFrame(index=df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.index, columns=df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.columns)

            for district, seats in dicionario_sistema_atual.items():
                district_votes = df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.loc[district]
                result_seats, _ = alocar_deputados_metodo_hondt(district_votes.to_frame().T, seats)
                
                result_seats_series = result_seats.iloc[0]  
                result_seats_series.name = district 

                df_deputados_com_europa_e_fora_europa_sistema_atual.loc[district] = result_seats_series

            # Sem Europa e Fora da Europa 
            df_deputados_sem_europa_e_fora_europa_sistema_atual = df_deputados_com_europa_e_fora_europa_sistema_atual.copy()
            df_deputados_sem_europa_e_fora_europa_sistema_atual = df_deputados_sem_europa_e_fora_europa_sistema_atual[
                ~df_deputados_sem_europa_e_fora_europa_sistema_atual.index.isin(['Europa', 'Fora da Europa'])
            ]

#########################################################################################################################################################

            # Lista de candidatos no círculo nacional, por círculo e tipo (efetivo/suplente)

            df_lista_candidatos_circulo_nacional = pd.DataFrame(columns=['partido', 'circulo', 'candidato', 'votos_nao_usados_no_uninominais'])

            # Para cada círulo uninominal, determinar o candidatos a adicionar à lista de candidatos, por partido, ao círculo nacional
            for circulo in df_votos_transferidos_para_nacional_agregados.index:
                for partido in df_votos_transferidos_para_nacional_agregados.columns:
                    tipo_candidato = "Suplente" if df_deputados_alocados_diretamente_circulos_uninominais.at[circulo, partido] == 1 else "Efetivo"
                    votos_nao_usados_no_uninominais = df_votos_transferidos_para_nacional_agregados.at[circulo, partido]      

                    df_lista_candidatos_circulo_nacional = df_lista_candidatos_circulo_nacional._append({
                        'Partido': partido,
                        'Círculo': circulo,
                        'Tipo candidato': tipo_candidato,
                        'Votos transferidos': votos_nao_usados_no_uninominais
                    }, ignore_index=True)

            df_lista_candidatos_circulo_nacional = df_lista_candidatos_circulo_nacional.sort_values(by=['Partido', 'Votos transferidos'], ascending=[True, False])

            df_lista_candidatos_circulo_nacional['Ordenação no partido, por votos transferidos'] = df_lista_candidatos_circulo_nacional.groupby('Partido').cumcount() + 1

            df_lista_candidatos_circulo_nacional = df_lista_candidatos_circulo_nacional.sort_values(by=['Votos transferidos'])
            df_lista_candidatos_circulo_nacional.reset_index(inplace=True)
            df_lista_candidatos_circulo_nacional['Ordenação, por votos transferidos'] = df_lista_candidatos_circulo_nacional.index + 1
            df_lista_candidatos_circulo_nacional['Ordenação, por votos transferidos'] = df_lista_candidatos_circulo_nacional['Ordenação, por votos transferidos'].iloc[::-1].values

            df_lista_candidatos_circulo_nacional.drop(columns='index', inplace=True)

            columns_order = ['Ordenação, por votos transferidos', 'Partido', 'Ordenação no partido, por votos transferidos', 'Círculo', 'Tipo candidato', 'Votos transferidos']
            df_lista_candidatos_circulo_nacional = df_lista_candidatos_circulo_nacional[columns_order]
            df_lista_candidatos_circulo_nacional = df_lista_candidatos_circulo_nacional.sort_values(by=['Ordenação, por votos transferidos', 'Partido', 'Ordenação no partido, por votos transferidos', 'Tipo candidato', 'Votos transferidos'])

            df_lista_candidatos_circulo_nacional['Votos transferidos'] = pd.to_numeric(df_lista_candidatos_circulo_nacional['Votos transferidos'], errors='coerce')

            df_lista_candidatos_circulo_nacional['Votos transferidos'].fillna(0, inplace=True)

            df_lista_candidatos_circulo_nacional['Eleito'] = 0

            for partido in df_deputados_alocados_nacional.columns:
                # Determina o número de mandatos atribuidos, por partido
                num_seats = df_deputados_alocados_nacional[partido].iloc[0]
                # Filtra o df dos candidatos, por partido
                partido_candidatos = df_lista_candidatos_circulo_nacional[df_lista_candidatos_circulo_nacional['Partido'] == partido]
                # Seleciona os X maiores candidatos, com base nos votos transferidos
                elected_candidatos = partido_candidatos.nlargest(num_seats, 'Votos transferidos').index
                df_lista_candidatos_circulo_nacional.loc[elected_candidatos, 'Eleito'] = 1

#########################################################################################################################################################


            # Para limiar
            df_lista_candidatos_circulo_nacional = df_lista_candidatos_circulo_nacional[df_lista_candidatos_circulo_nacional['Partido'].isin(lista_partidos_acima_do_limiar)]


            pivot_df = df_lista_candidatos_circulo_nacional.pivot_table(index='Círculo', columns='Partido', values='Eleito', aggfunc='sum', fill_value=0)

            df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional = pivot_df.rename_axis(None, axis=1).rename_axis(None, axis=0)

#########################################################################################################################################################

            df_deputados_alocados_diretamente_e_indiretamente = df_deputados_alocados_diretamente_circulos_uninominais.add(df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional, fill_value=0)

#########################################################################################################################################################

            # Para calcular o número de deputados eleitos, pelo sistema proposto, nos círculos atuais

            df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_copia=df_deputados_alocados_diretamente_circulos_uninominais.copy()
            df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_copia['ciruclo_atual'] = df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_copia.index.to_series().str.replace(" nº ", "", regex=False).str.replace(r'\d', '', regex=True)
            df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais = df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_copia.groupby('ciruclo_atual').sum()

            df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacionalnos_nos_circulos_atuais_copia=df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional.copy()
            df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacionalnos_nos_circulos_atuais_copia['ciruclo_atual'] = df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacionalnos_nos_circulos_atuais_copia.index.to_series().str.replace(" nº ", "", regex=False).str.replace(r'\d', '', regex=True)
            df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais = df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacionalnos_nos_circulos_atuais_copia.groupby('ciruclo_atual').sum()

            df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_copia=df_deputados_alocados_diretamente_e_indiretamente.copy()
            df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_copia['ciruclo_atual'] = df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_copia.index.to_series().str.replace(" nº ", "", regex=False).str.replace(r'\d', '', regex=True)
            df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais = df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_copia.groupby('ciruclo_atual').sum()

#########################################################################################################################################################

            #  Votos totais, SP, por círculo, território nacional e território estrangeiro

            df_votos_usados_uninominal_agregados_por_partido = df_votos_usados_uninominal_agregados.sum()
            df_votos_transferidos_para_nacional_agregados_por_partido = df_votos_transferidos_para_nacional_agregados.sum()
            df_territorio_estrangeiro_europa_votos_agregados_por_partido = df_territorio_estrangeiro_europa_votos_agregados.sum()
            df_territorio_estrangeiro_fora_da_europa_votos_agregados_por_partdo = df_territorio_estrangeiro_fora_da_europa_votos_agregados.sum()

            df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa = pd.DataFrame({
                'Uninominais': df_votos_usados_uninominal_agregados_por_partido,
                'Nacional': df_votos_transferidos_para_nacional_agregados_por_partido,
            })

            df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa = pd.DataFrame({
                'Uninominais': df_votos_usados_uninominal_agregados_por_partido,
                'Nacional': df_votos_transferidos_para_nacional_agregados_por_partido,
                "Europa":df_territorio_estrangeiro_europa_votos_agregados_por_partido,
                "Fora da Europa": df_territorio_estrangeiro_fora_da_europa_votos_agregados_por_partdo
            })


            # Com limiar
            #  Votos totais, SP, por círculo, território nacional e território estrangeiro

            df_votos_usados_uninominal_agregados_por_partido = df_votos_usados_uninominal_agregados.sum()
            df_votos_transferidos_para_nacional_agregados_com_limiar =df_votos_transferidos_para_nacional_agregados[lista_partidos_acima_do_limiar]
            df_votos_transferidos_para_nacional_agregados_por_partido_com_limiar = df_votos_transferidos_para_nacional_agregados_com_limiar.sum()
            df_territorio_estrangeiro_europa_votos_agregados_por_partido = df_territorio_estrangeiro_europa_votos_agregados.sum()
            df_territorio_estrangeiro_fora_da_europa_votos_agregados_por_partdo = df_territorio_estrangeiro_fora_da_europa_votos_agregados.sum()

            df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar = pd.DataFrame({
                'Uninominais': df_votos_usados_uninominal_agregados_por_partido,
                'Nacional': df_votos_transferidos_para_nacional_agregados_por_partido_com_limiar,
            })

            df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar = df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar.fillna(0)

            df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar = pd.DataFrame({
                'Uninominais': df_votos_usados_uninominal_agregados_por_partido,
                'Nacional': df_votos_transferidos_para_nacional_agregados_por_partido_com_limiar,
                "Europa":df_territorio_estrangeiro_europa_votos_agregados_por_partido,
                "Fora da Europa": df_territorio_estrangeiro_fora_da_europa_votos_agregados_por_partdo
            })

            df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar = df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar.fillna(0)

#########################################################################################################################################################

            # Deputados totais, SP, por círculo, territorio nacional e território estrangeiro

            df_deputados_alocados_diretamente_circulos_uninominais_por_partido = df_deputados_alocados_diretamente_circulos_uninominais.sum()
            df_deputados_alocados_nacional_por_partido = df_deputados_alocados_nacional.sum()
            df_deputados_alocados_europa_por_partido = df_deputados_alocados_europa.sum()
            df_deputados_alocados_fora_da_europa_por_partido = df_deputados_alocados_fora_da_europa.sum()

            df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa = pd.DataFrame({
                'Uninominais': df_deputados_alocados_diretamente_circulos_uninominais_por_partido,
                'Nacional': df_deputados_alocados_nacional_por_partido,
            })


            df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa = pd.DataFrame({
                'Uninominais': df_deputados_alocados_diretamente_circulos_uninominais_por_partido,
                'Nacional': df_deputados_alocados_nacional_por_partido,
                "Europa":df_deputados_alocados_europa_por_partido,
                "Fora da Europa": df_deputados_alocados_fora_da_europa_por_partido
            })

#########################################################################################################################################################

            # Número de votos desperdiçados, por círculo 

            # Sistema proposto
            # Território nacional e território estrangeiro
            df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa = df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa.copy()


            for district in df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa.index:
                for party in df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa.columns:

                    if df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa.loc[district, party] > 0:

                        df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa.loc[district, party] = 0

            # Território nacional 
            df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa = df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa.copy()


            for district in df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa.index:
                for party in df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa.columns:

                    if df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa.loc[district, party] > 0:

                        df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa.loc[district, party] = 0

            # Sistema atual
            # Território nacional e território estrangeiro
            df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual = df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.copy()


            for district in df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual.index:
                for party in df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual.columns:

                    if df_deputados_com_europa_e_fora_europa_sistema_atual.loc[district, party] > 0:

                        df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual.loc[district, party] = 0

            # Território nacional
            df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual = df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados.copy()


            for district in df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual.index:
                for party in df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual.columns:

                    if df_deputados_sem_europa_e_fora_europa_sistema_atual.loc[district, party] > 0:

                        df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual.loc[district, party] = 0

#########################################################################################################################################################

            # Ordenar colunas/linhas alfabeticamente (e adicionar totais)

            lista_dataframes_ordenar_linhas_e_colunas_e_adicionar_totais = [
                df_territorio_nacional_votos_agregados,
                df_votos_usados_uninominal_agregados,
                df_votos_transferidos_para_nacional_agregados,
                df_territorio_estrangeiro_votos_agregados,
                df_deputados_alocados_diretamente_circulos_uninominais,
                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional,
                df_deputados_alocados_diretamente_e_indiretamente,
                df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais,
                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais,
                df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais,
                df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados,
                df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados,
                df_deputados_com_europa_e_fora_europa_sistema_atual,
                df_deputados_sem_europa_e_fora_europa_sistema_atual,
                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual,
                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual
            ]


            for i in range(len(lista_dataframes_ordenar_linhas_e_colunas_e_adicionar_totais)):
                lista_dataframes_ordenar_linhas_e_colunas_e_adicionar_totais[i] = ordenar_linhas_colunas_e_adicionar_totais_df(lista_dataframes_ordenar_linhas_e_colunas_e_adicionar_totais[i])


            lista_dataframes_ordenar_linhas_e_adicionar_totais = [
                df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa,
                df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa,
                df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar,
                df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar,
                df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa,
                df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa,
                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa,
                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa,
            ]

            for i in range(len(lista_dataframes_ordenar_linhas_e_adicionar_totais)):
                lista_dataframes_ordenar_linhas_e_adicionar_totais[i] = ordenar_linhas_e_adicionar_totais_df(lista_dataframes_ordenar_linhas_e_adicionar_totais[i])


            lista_dataframes_ordenar_linhas_e_colunas = [
                df_quocientes_metodo_hondt_nacional,
                df_quocientes_metodo_hondt_europa,
                df_quocientes_metodo_hondt_fora_da_europa,
            ]

            for i in range(len(lista_dataframes_ordenar_linhas_e_colunas)):
                lista_dataframes_ordenar_linhas_e_colunas[i] = ordenar_colunas_linhas_df(lista_dataframes_ordenar_linhas_e_colunas[i])
                
#########################################################################################################################################################

            # Votos médios

            # Sistema proposto
            # Território nacional e território estrangeiro
            df_numero_medio_votos_por_deputado_com_europa_e_fora_europa = df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar / df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa.replace(0, np.nan)
            df_numero_medio_votos_por_deputado_com_europa_e_fora_europa = df_numero_medio_votos_por_deputado_com_europa_e_fora_europa + 0.0000001 # + 0.0000001 devido a possivel bug no arredondamento
            df_numero_medio_votos_por_deputado_com_europa_e_fora_europa = df_numero_medio_votos_por_deputado_com_europa_e_fora_europa.apply(pd.to_numeric, errors='coerce').round()
            # Território nacional
            df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa = df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar / df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa.replace(0, np.nan)
            df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa = df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa + 0.0000001 # + 0.0000001 devido a possivel bug no arredondamento
            df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa = df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa.apply(pd.to_numeric, errors='coerce').round()

            # Sistema atual
            # Território nacional e território estrangeiro
            df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual = df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados / df_deputados_com_europa_e_fora_europa_sistema_atual.replace(0, np.nan)
            df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual = df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual + 0.0000001 # + 0.0000001 devido a possivel bug no arredondamento
            df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual = df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual.apply(pd.to_numeric, errors='coerce').round()
            # Território nacional
            df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual = df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados / df_deputados_sem_europa_e_fora_europa_sistema_atual.replace(0, np.nan)
            df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual = df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual + 0.0000001 # + 0.0000001 devido a possivel bug no arredondamento
            df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual = df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual.apply(pd.to_numeric, errors='coerce').round()

#########################################################################################################################################################

            momento_agora = datetime.datetime.now()
            ficheiro_output = f"eleicao_{ano}_limite_{limite_per}_cenario_{cenario}_com_{numero_cirulos_uninominais}_uninominais_e_{numero_mandatos_circulo_nacional}_no_c_nacional_simulado_em_{momento_agora.strftime('%Y')}_{momento_agora.strftime('%m')}_{momento_agora.strftime('%d')}_{momento_agora.strftime('%H')}_{momento_agora.strftime('%M')}_{momento_agora.strftime('%S')}.xlsx"

#########################################################################################################################################################

            # Para criar um ficheiro excel com os df's'
          
            with pd.ExcelWriter(f"../output/{ano}/{limite_per}/{ficheiro_output}", engine='openpyxl') as writer:
                df_original.to_excel(writer, sheet_name='Dados originais', index=False)
                
#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_territorio_nacional_votos_agregados.index.name = "Votos, por círculo uninominal e partido (TN)"
                df_territorio_nacional_votos_agregados.to_excel(writer, sheet_name='Votos por c. uni. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_territorio_nacional_votos_agregados.columns) + 2

                df_votos_usados_uninominal_agregados.index.name = "Votos usados nos círculos uninominais, por círculo uninominal e partido (TN)"
                df_votos_usados_uninominal_agregados.to_excel(writer, sheet_name='Votos por c. uni. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_votos_usados_uninominal_agregados.columns) + 2

                df_votos_transferidos_para_nacional_agregados.index.name = "Votos transferidos, por círculo uninominal e partido (TN)"
                df_votos_transferidos_para_nacional_agregados.to_excel(writer, sheet_name='Votos por c. uni. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_territorio_estrangeiro_votos_agregados.index.name = "Votos, por círculo plurinominal e partido (TE)"
                df_territorio_estrangeiro_votos_agregados.to_excel(writer, sheet_name='Votos por c. pluri. (TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_quocientes_metodo_hondt_nacional.index.name = "Quocientes, pelo método de Hondt, no círculo plurinominal Nacional (TN)"
                df_quocientes_metodo_hondt_nacional.to_excel(writer, sheet_name='Quocientes (TN ou TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_territorio_estrangeiro_votos_agregados.columns) + 2

                df_quocientes_metodo_hondt_europa.index.name = "Quocientes, pelo método de Hondt, no círculo plurinominal Europa (TE)"
                df_quocientes_metodo_hondt_europa.to_excel(writer, sheet_name='Quocientes (TN ou TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_territorio_estrangeiro_votos_agregados.columns) + 2

                df_quocientes_metodo_hondt_fora_da_europa.index.name = "Quocientes, pelo método de Hondt, no círculo plurinominal Fora da Europa (TE)"
                
                df_quocientes_metodo_hondt_fora_da_europa.to_excel(writer, sheet_name='Quocientes (TN ou TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)
#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_lista_candidatos_circulo_nacional.index.name = "Resultado, por candidato, no círculo Nacional (TN)"
                df_lista_candidatos_circulo_nacional.to_excel(writer, sheet_name='R. por candidato, c. Nac. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=False, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_deputados_alocados_diretamente_circulos_uninominais.index.name = "Resultado nos círculos uninominais, por círculo uninominal e partido (TN)"
                df_deputados_alocados_diretamente_circulos_uninominais.to_excel(writer, sheet_name='R. por c. uni. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_alocados_diretamente_circulos_uninominais.columns) + 2

                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional.index.name = "Resultado no círculo Nacional, por círculo uninominal (de origem) e partido (TN)"
                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional.to_excel(writer, sheet_name='R. por c. uni. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional.columns) + 2

                df_deputados_alocados_diretamente_e_indiretamente.index.name = "Resultado total, por círculo uninominal e partido (TN)"
                df_deputados_alocados_diretamente_e_indiretamente.to_excel(writer, sheet_name='R. por c. uni. (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_transpose = df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais.transpose()
                df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_transpose.index.name = "Resultado nos círculos uninominais, por partido e círculo atual (TN)"
                df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_transpose.to_excel(writer, sheet_name='R. por c. atual (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_alocados_diretamente_circulos_uninominais_nos_circulos_atuais_transpose.columns) + 2

                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais_transpose = df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais.transpose()
                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais_transpose.index.name = "Resultado no círculo Nacional, por partido e círculo atual (de origem) (TN)"
                df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais_transpose.to_excel(writer, sheet_name='R. por c. atual (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_alocados_indiretamente_circulos_uninominais_pelo_circulo_nacional_nos_circulos_atuais_transpose.columns) + 2

                df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_transpose = df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais.transpose()
                df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_transpose.index.name = "Resultado total, por partido e círculo atual (TN)"
                df_deputados_alocados_diretamente_e_indiretamente_nos_circulos_atuais_transpose.to_excel(writer, sheet_name='R. por c. atual (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa.index.name = "Votos, por partido e círculo (círculos uninominais agregados) no SP (TN)"
                df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa.columns) + 2

                df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar.index.name = "Votos, por partido e círculo (círculos uninominais agregados), filtrado pela percent. limite no Nacional, no SP (TN)"
                df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar.to_excel(writer, sheet_name='Resumo SP (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_votos_totais_por_tipo_circulo_sem_europa_e_fora_europa_com_limiar.columns) + 2

                df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa.index.name = "Deputados, por partido e círculo (círculos uninominais agregados) no SP (TN)"
                df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_totais_por_tipo_circulo_sem_europa_e_fora_europa.columns) + 2

                df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa.index.name = "Votos médios, por partido e círculo (círculos uninominais agregados) no SP (TN)"
                df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa.columns) + 2

                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa.index.name = "Votos desperdiçados, por partido e círculo (círculos uninominais agregados) no SP (TN)"
                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa.index.name =  "Votos, por partido e círculo (círculos uninominais agregados) no SP (TN e TE)" 
                df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa.columns) + 2


                df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar.index.name ="Votos, por partido e círculo (círculos uninominais agregados), filtrado pela percent. limite no Nacional, (TN e TE)"
                df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar.to_excel(writer, sheet_name='Resumo SP (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_votos_totais_por_tipo_circulo_com_europa_e_fora_europa_com_limiar.columns) + 2

                df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa.index.name = "Deputados, por partido e círculo (círculos uninominais agregados) no SP (TN e TE)"
                df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_totais_por_tipo_circulo_com_europa_e_fora_europa.columns) + 2

                df_numero_medio_votos_por_deputado_com_europa_e_fora_europa.index.name = "Votos médios, por partido e círculo (círculos uninominais agregados) no SP (TN e TE)"
                df_numero_medio_votos_por_deputado_com_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_numero_medio_votos_por_deputado_com_europa_e_fora_europa.columns) + 2

                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa.index.name = "Votos desperdiçados, por partido e círculo (círculos uninominais agregados) no SP (TN e TE)"
                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa.to_excel(writer, sheet_name='Resumo SP (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados_transpose = df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados.transpose()
                df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados_transpose.index.name = "Votos, por partido e círculo atual no SA (TN)"
                df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados_transpose.to_excel(writer, sheet_name='Resumo SA (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_territorio_total_sem_europa_e_fora_europa_sistema_atual_agregados_transpose.columns) + 2

                df_deputados_sem_europa_e_fora_europa_sistema_atual_transpose = df_deputados_sem_europa_e_fora_europa_sistema_atual.transpose()
                df_deputados_sem_europa_e_fora_europa_sistema_atual_transpose.index.name = "Deputados, por partido e círculo atual no SA (TN)"
                df_deputados_sem_europa_e_fora_europa_sistema_atual_transpose.to_excel(writer, sheet_name='Resumo SA (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_sem_europa_e_fora_europa_sistema_atual_transpose.columns) + 2

                df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual_transpose = df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual.transpose()
                df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual_transpose.index.name = "Votos médios, por partido e círculo atual no SA (TN)"
                df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual_transpose.to_excel(writer, sheet_name='Resumo SA (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_numero_medio_votos_por_deputado_sem_europa_e_fora_europa_sistema_atual_transpose.columns) + 2

                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual_transpose = df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual.transpose()
                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual_transpose.index.name = "Votos desperdiçados, por partido e círculo atual no SA (TN)"
                df_votos_desperdicados_por_tipo_circulo_sem_europa_e_fora_europa_sistema_atual_transpose.to_excel(writer, sheet_name='Resumo SA (TN)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

#########################################################################################################################################################

                linha_inicio = 0
                coluna_inicio = 0

                df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados_transpose = df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados.transpose()
                df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados_transpose.index.name = "Votos, por partido e círculo atual no SA (TN e TE)"
                df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados_transpose.to_excel(writer, sheet_name='Resumo SA (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_territorio_total_com_europa_e_fora_europa_sistema_atual_agregados_transpose.columns) + 2

                df_deputados_com_europa_e_fora_europa_sistema_atual_transpose = df_deputados_com_europa_e_fora_europa_sistema_atual.transpose()
                df_deputados_com_europa_e_fora_europa_sistema_atual_transpose.index.name = "Deputados, por partido e círculo atual no SA (TN e TE)"
                df_deputados_com_europa_e_fora_europa_sistema_atual_transpose.to_excel(writer, sheet_name='Resumo SA (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_deputados_com_europa_e_fora_europa_sistema_atual_transpose.columns) + 2

                df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual_transpose = df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual.transpose()
                df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual_transpose.index.name = "Votos médios, por partido e círculo atual no SA (TN e TE)"
                df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual_transpose.to_excel(writer, sheet_name='Resumo SA (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)

                coluna_inicio += len(df_numero_medio_votos_por_deputado_com_europa_e_fora_europa_sistema_atual_transpose.columns) + 2

                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual_transpose = df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual.transpose()
                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual_transpose.index.name = "Votos desperdiçados, por partido e círculo atual no SA (TN e TE)"
                df_votos_desperdicados_por_tipo_circulo_com_europa_e_fora_europa_sistema_atual_transpose.to_excel(writer, sheet_name='Resumo SA (TN e TE)', startrow=linha_inicio, startcol=coluna_inicio, index=True, index_label=False)


            print(f"Ficheiro {ficheiro_output} guardado na pasta output/{limite_per}/{ano}")

            count = count + 1
            print(f"Ficheiros criados: {count}")
        print("Done")
          

