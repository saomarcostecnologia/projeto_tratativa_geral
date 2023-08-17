# use_cases.py
import pandas as pd
from datetime import date
from datetime import datetime
import datetime
from entities import FileTreatment
import tkinter.filedialog as filedialog
from tkinter import messagebox
import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import NamedStyle
from openpyxl.styles import Font
from openpyxl.styles import Border
from openpyxl.styles import Side
from openpyxl.styles import Alignment


class FileTreatmentUseCase:
    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            print("Arquivo selecionado: ", file_path)
            return FileTreatment(file_path)  # Por padrão, sempre será "" neste exemplo
        return None

    #Função tratamento de razao salario
    def tratar_salario(self, file_path):
        try:
            # Leitura do dataframe a partir do arquivo Excel fornecido pelo caminho 'file_path'
            arquivo = pd.read_excel(file_path, header=None)

            # Armazenando a primeira linha como cabeçalho
            novo_cabecalho = arquivo.iloc[9]

            # Definindo a primeira linha como cabeçalho
            arquivo = arquivo[1:]
            arquivo.columns = novo_cabecalho

            # Excluir as 9 primeiras linhas (0 a 8) para remover informações indesejadas
            arquivo = arquivo.iloc[9:]

            # Resetando os índices após a exclusão das linhas
            arquivo = arquivo.reset_index(drop=True)

            # Primeira etapa de tratativa: filtrar e manipular os dados para a categoria 'REC'
            tratativa01 = arquivo[arquivo['Categoria'] == 'REC']
            tratativa01['tipo_documento'] = tratativa01['Historico'].str.split(';').str[-1]
            tratativa01['nf'] = tratativa01['Historico'].str.split('Doc').str[-1].str.split('-').str[0]
            df_REC = tratativa01.reset_index(drop=True)

            # Segunda etapa de tratativa: filtrar e manipular os dados para a categoria 'NFFs de Compra'
            tratativa02 = arquivo[arquivo['Categoria'] == 'NFFs de Compra']
            tratativa02['tipo_documento'] = None
            tratativa02['nf'] = tratativa02['Historico'].str.split('NFF').str[-1].str.split('Forn').str[0]
            df_nff_compra = tratativa02.reset_index(drop=True)

            # Terceira etapa de tratativa: filtrar e manipular os dados para a categoria 'Pagamentos'
            tratativa03 = arquivo[arquivo['Categoria'] == 'Pagamentos']
            tratativa03['tipo_documento'] = None
            tratativa03['nf'] = tratativa03['Historico'].str.split(' ').str[-1]
            df_pagamentos = tratativa03.reset_index(drop=True)

            # Concatenar os dataframes tratados das diferentes categorias em um único dataframe
            df_rec_nff_pagamento = pd.concat([df_REC, df_nff_compra, df_pagamentos], axis=0)
            df_rec_nff_pagamento['nf'] = df_rec_nff_pagamento['nf'].str.replace(' ', '')

            # Filtrar e manipular os dados para o dataframe com nota fiscal e tipo de documento
            tratamento_nf_categoria = df_rec_nff_pagamento[['nf', 'tipo_documento']]
            tratamento_nf_categoria = tratamento_nf_categoria.dropna(subset=['tipo_documento'])

            # Mesclar o dataframe df_rec_nff_pagamento com o dataframe tratamento_nf_categoria usando a coluna 'nf'
            df_rec_nff_pagamento_categoria = pd.merge(df_rec_nff_pagamento, tratamento_nf_categoria, on='nf', how='outer')

            # Filtrar e manipular os dados para as categorias 'PEOPLESOFT' e 'XRT'
            df_peoplesoft = arquivo[arquivo['Categoria'] == 'PEOPLESOFT']

            df_peoplesoft['tipo_documento'] = df_peoplesoft['Historico'].str.split('refere').str[0]
            df_peoplesoft['tipo_documento'] = df_peoplesoft['tipo_documento'].str.replace("'", '')

            df_peoplesoft['tipo_documento'] = df_peoplesoft['tipo_documento'].str.replace("VL LIQUIDO RESC", "RESCISAO")
            df_peoplesoft['tipo_documento'] = df_peoplesoft['tipo_documento'].str.replace("VL LIQ FERIAS", "FERIAS")
            df_peoplesoft['tipo_documento'] = df_peoplesoft['tipo_documento'].str.replace("VL LIQUIDO", "FOLHADEPAGAMENTO")
            df_peoplesoft['tipo_documento'] = df_peoplesoft['tipo_documento'].str.replace("VL REST DESC INDEV", "VL REST DESC INDEV")

            df_xrt = arquivo[arquivo['Categoria'] == 'XRT']

            df_rec_nff_pagamento_categoria = df_rec_nff_pagamento_categoria.rename(columns={'tipo_documento_y': 'tipo_documento'})

            # Concatenar os dataframes tratados das categorias 'PEOPLESOFT' e 'XRT' com o dataframe existente
            df_salarios_pgto = pd.concat([df_rec_nff_pagamento_categoria, df_xrt, df_peoplesoft], axis=0)
            df_salarios_pgto = df_salarios_pgto.reset_index(drop=True)

            # Remover colunas desnecessárias e formatar valores
            df_salarios_pgto = df_salarios_pgto.drop(['tipo_documento_x'], axis=1)
            df_salarios_pgto['tipo_documento'] = df_salarios_pgto['tipo_documento'].str.replace(' ', '')
            df_salarios_pgto['Credito'] = df_salarios_pgto['Credito'] * -1

            # Criar um nome dinâmico para o arquivo tratado com base na hora atual
            file_path_sem_extensao = file_path[:-5]  # Remove a extensão do arquivo
            hora_atual = datetime.datetime.now().strftime("%H%M%S")  # Obtém a hora atual no formato HHMMSS
            nome_arquivo_tratado = f"{file_path_sem_extensao}.{hora_atual}.xlsx"

            # Salvar o dataframe tratado no novo arquivo Excel com o nome dinâmico
            df_salarios_pgto.to_excel(nome_arquivo_tratado, sheet_name='Tratado', index=False)

            # Exibir mensagem de sucesso
            messagebox.showinfo("Sucesso", "O arquivo foi tratado com sucesso!")
            return True

        except Exception as a:
            # Em caso de erro, exibir mensagem de erro e o motivo específico do erro
            messagebox.showwarning("Erro", "Ocorreu um erro durante o tratamento do arquivo.")
            print("Ocorreu um erro durante o tratamento do arquivo:", str(a))
            return False
        
    #Função de tratamento de Balancete
    def tratar_balancete(self, file_path):
        try:
            #Leitura de dataframe e manipulação de planilha
            balancete = pd.read_excel(file_path, header=None)
            balancetemoeda = balancete
            
            #subir linha 12 para cabeçalho
            # Armazenando a primeira linha como cabeçalho
            novo_cabecalho = balancete.iloc[12]
            #subir o cabeçalho para coluna.
            balancete.columns = novo_cabecalho

            # Resetando os índices
            balancete = balancete.reset_index(drop=True)
            #Remoção das primeiras linhas.

            balancete = balancete.drop(balancete.index[:13])
            balancete = balancete.dropna(subset=["CONTA"])

            #numero de conta de niveis

            #balancete['concat_2'] = balancete['CONTA'].str.split(('-')).str[0] 
            balancete['Nivel'] = balancete['CONTA'].str.split(('                      ')).str[0].str.split().str[0]
            balancete['id'] = balancete['CONTA'].str.split(('-')).str[0].str.split().str[0]

            #Tratamento de DataFrame para conecatenar linhas vazias.

            tratamento01 = pd.DataFrame({'id': balancete['id'],'valor': balancete['Nivel']})
            tratamento01 = tratamento01[tratamento01['valor'].isna()]
            tratamento01['Setimo_nivel'] = balancete['id']

            #Mesclagem da base balancete e tratamento01

            balancete_tratado = pd.merge(balancete, tratamento01, left_on= 'id', right_on= 'id', how='left')

            #Preenchimento de colunas vazias de Nivel

            balancete_tratado['Nivel'].fillna(method='ffill', inplace=True)

            #Concatenando colunas.

            balancete_tratado['Nivel_Correto'] = balancete_tratado['Nivel'].astype(str) + balancete_tratado['Setimo_nivel'].astype(str)

            #removendo duplicidade em coluna conta.

            balancete_tratado = balancete_tratado.drop_duplicates(subset='Nivel_Correto', keep='last')

            balancete_tratado['Nivel_Correto'] = balancete_tratado['Nivel_Correto'].str.split(('nan')).str[0]

            #Organizando Balancete

            #removendo colunas.

            balancete_tratado = balancete_tratado.drop(['Setimo_nivel','valor','id','Nivel'], axis=1)

            #Mudando nome de colunas

            balancete_tratado.rename(columns={'Nivel_Correto': 'Nivel'}, inplace=True)

            balancete_tratado.head(10)

            nome_coluna = ['CONTA', 'SALDO ANTERIOR','C/D1','DEBITO','CREDITO','SALDO ATUAL','C/D2','Nivel']

            balancete_tratado.columns = nome_coluna

            # organizando data frame

            balancete_tratado =  balancete_tratado[['Nivel','CONTA', 'SALDO ANTERIOR','C/D1','DEBITO','CREDITO','SALDO ATUAL','C/D2']]

            # Resetar o índice

            balancete_tratado = balancete_tratado.reset_index(drop=True)

            balancete_etl2 = balancete_tratado.tail(5)

            # Adicione uma nova coluna 'Nova_Coluna' com valores nulos (None)
            balancete_etl2.loc[:, 'Nova_Coluna'] = None

            #Organizando o dataframe

            balancete_etl2 =  balancete_etl2[['Nivel','CONTA','SALDO ANTERIOR','Nova_Coluna','C/D1','DEBITO','CREDITO','SALDO ATUAL','C/D2',]]

            # Realiza as renomeações de colunas
            balancete_etl2.rename(columns={'C/D1': 'DEBITO', 'DEBITO': 'CREDITO', 'CREDITO': 'SALDO ATUAL', 'SALDO ATUAL': 'C/D1'}, inplace=True)

            nova_ordem = balancete_etl2[['Nivel', 'CONTA', 'SALDO ANTERIOR', 'C/D1', 'DEBITO', 'CREDITO', 'SALDO ATUAL', 'C/D2']]
            balancete_etl2 = nova_ordem


            #remoção das ultimas linhas do balancete

            balancete_tratado = balancete_tratado.drop(balancete_tratado.tail(5).index)

            # Assuming you have a DataFrame called balancete_tratado already defined
            dados = {
                'Nivel': [None],
                'CONTA': [None],
                'SALDO ANTERIOR': [None],
                'C/D1': [None],
                'DEBITO': [None],
                'CREDITO': [None],
                'SALDO ATUAL': [None],
                'C/D2': [None],
            }

            nova_linha = pd.DataFrame(dados)

            # Concatenating the new row to the original DataFrame
            balancete_tratado = pd.concat([balancete_tratado, nova_linha], ignore_index=True)
            # Print the DataFrame
            balancete_tratado.tail(5)

            # Concatena o dataframe 'balancete_etl_01' no final do dataframe 'balancete_tratado'
            balancete_tratado = pd.concat([balancete_tratado, balancete_etl2], ignore_index=True)

            # Criar um nome dinâmico para o arquivo tratado
            file_path_sem_extensao = file_path[:-5] # Remover os últimos 5 caracteres da variável file_path (assumindo que sejam ".xlsx")
            hora_atual = datetime.datetime.now().strftime("%H%M%S")  # Obtém a hora atual no formato HHMMSS
            nome_arquivo_tratado = f"{file_path_sem_extensao}.{hora_atual}.xlsx"
            # Salvar o arquivo tratado com o nome dinâmico
            balancete_tratado.to_excel(nome_arquivo_tratado, sheet_name='Tratado', index=False)
            
            # Exibir mensagem de sucesso
            messagebox.showinfo("Sucesso", "O arquivo foi tratado com sucesso!")
            return True

        except Exception as a:
            # Em caso de erro, exibir mensagem de erro e o motivo específico do erro
            messagebox.showwarning("Erro", "Ocorreu um erro durante o tratamento do arquivo.")
            print("Ocorreu um erro durante o tratamento do arquivo:", str(a))
            return False
        
    #Função de tratamento de Razão     
    def tratar_razao(self, file_path):
        try:
            # Leitura de dataframe e manipulação de planilha
            razao = pd.read_excel(file_path, header=None)
            
            diretorio = razao
            razao_cabecalho = diretorio
            razao = diretorio
            razao = razao.drop(razao.index[:11])

            #subir linha 15 para cabeçalho

            # Armazenando a primeira linha como cabeçalho
            novo_cabecalho = razao_cabecalho.iloc[15]

            # Definindo a primeira linha como cabeçalho
            razao = razao[1:]
            razao.columns = novo_cabecalho

            # Resetando os índices
            razao = razao.reset_index(drop=True)

            #df para a segunda parte do tratamento
            razao_pt2 = razao

            #Coluna para tratamento

            etl_contas = razao[['DIA','UA']]

            etl_contas = pd.DataFrame(etl_contas, columns=['DIA','UA'])

            etl_contas_2 = etl_contas

            etl_contas_2 = etl_contas_2.drop(etl_contas_2[(etl_contas_2 == 0).any(axis=1) |
                                                        etl_contas_2.isnull().any(axis=1) |
                                                        etl_contas_2['UA'].str.startswith('UA') |
                                                        etl_contas_2['DIA'].str.startswith('CONTAS ANTERI') |
                                                        etl_contas_2['UA'].str.startswith('PROP') 
                                                        ]
                                                        .index)
            etl_contas_2 = etl_contas_2.reset_index(drop=True)

            # Criar uma lista com a sequência "banco" e "estrutura"
            sequencia = ["1","2","3"]
            # Atribuir a sequência à coluna "ColunaNova" do DataFrame
            etl_contas_2["ColunaNova"] = sequencia * (len(etl_contas_2) // len(sequencia)) + sequencia[:len(etl_contas_2) % len(sequencia)]

            # Realizar o pivô
            etl_contas_3 = etl_contas_2.pivot(columns='ColunaNova', values='UA')

            #Preenchimento de colunas vazias.

            etl_contas_3['1'].fillna(method='ffill', inplace=True)
            etl_contas_3['2'].fillna(method='ffill', inplace=True)

            #Exclusão de vazios na coluna 3
            etl_contas_3 = etl_contas_3.dropna(subset=["3"])

            #Reset index
            etl_contas_3 = etl_contas_3.reset_index(drop=True)


            #Seleção de colunas
            etl_contas_3 = etl_contas_3[['2','3']]

            etl_contas_3 = pd.DataFrame(etl_contas_3, columns=['2','3'])

            #Mudando nome de Coluna

            etl_contas_3.rename(columns={'2': 'Estrutura de Contas'}, inplace=True)
            etl_contas_3.rename(columns={'3': 'Conta'}, inplace=True)

            estrutura = etl_contas_3

            # Segunda parte do Tratamento

            #Base INICIAL COM CABEÇALHO TRATADO
            razao = razao_pt2

            #Mesclagem de Estrutura de conta + razão

            razao = pd.merge(razao, estrutura, left_on= 'UA', right_on= 'Estrutura de Contas', how='left')

            #Preenchimento de colunas vazias de Estrtura de Contas e Contas

            razao['Estrutura de Contas'].fillna(method='ffill', inplace=True)
            razao['Conta'].fillna(method='ffill', inplace=True)

            #removendo o periodo da coluna 0
            razao = razao.drop(razao.index[:6])

            # Resetando os índices
            razao = razao.reset_index(drop=True)

            razao['Nova Coluna'] = razao['DIA']
            razao['Nova Coluna'] = razao['DIA'].str[:5]

            ### Nome da coluna que contém as informações
            nome_coluna = 'Nova Coluna'

            # Nome específico para identificar o início do range
            nome_inicio_range = 'MOVTO'

            # Nome específico para identificar o fim do range (que contém "PERÍODO")
            nome_fim_range = 'PERÍO'

            # Identificando o índice da linha de início do range
            indice_inicio_range = razao[razao[nome_coluna] == nome_inicio_range].index[0]
            indice_fim_range = razao[razao[nome_coluna] == nome_fim_range].index[0]

            x = len(razao['DIA'])

            # Nome específico para identificar o início do range
            nome_inicio_range = 'MOVTO'

            # Nome específico para identificar o fim do range (que contém "PERÍODO")
            nome_fim_range = 'PERÍO'

            # Inicializando os índices de início e fim
            indice_inicio_range = 0
            indice_fim_range = 0

            # Loop para remover todos os ranges de linhas
            while True:
                # Verificando se há um novo range a ser removido
                if nome_inicio_range in razao[nome_coluna].values and nome_fim_range in razao[nome_coluna].values:
                    # Identificando o índice da linha de início do range
                    indice_inicio_range = razao[razao[nome_coluna] == nome_inicio_range].index[0]
                    # Identificando o índice da linha de fim do range
                    indice_fim_range = razao[razao[nome_coluna] == nome_fim_range].index[0]
                    
                    # Removendo o range de linhas
                    razao = razao.drop(range(indice_inicio_range, indice_fim_range + 1))
                    # Resetando os índices
                    razao = razao.reset_index(drop=True)
                else:
                    # Se não há mais ranges a serem removidos, sair do loop
                    break

            #TERCEIRA PARTE DO TRATAMENTO

            # Definir o tamanho do conjunto de linhas
            tamanho_conjunto_linhas = 4

            # Inicializar o índice de início do conjunto de linhas
            indice_inicio_conjunto = 0

            #Colunas que serao armazenadas
            coluna1 = []
            coluna2 = []
            coluna3 = []
            coluna4 = []


            while indice_inicio_conjunto < len(razao):
                # Selecionar o conjunto de linhas atual
                linhas_selecionadas = razao['LOTE'].iloc[indice_inicio_conjunto:indice_inicio_conjunto+tamanho_conjunto_linhas]
                
                # Verificar se há linhas suficientes para formar o conjunto completo
                if len(linhas_selecionadas) == tamanho_conjunto_linhas:
                    # Extrair as informações das linhas selecionadas para novas colunas
                    df_extracted = linhas_selecionadas.str.extract(r'(\w+.*)')
                    
                    # Transpor o DataFrame resultante
                    df_transposed = df_extracted.transpose()
                    
                    # Resetar o índice
                    df_transposed = df_transposed.reset_index(drop=True)
                    
                    # Adicionar as informações nas listas correspondentes
                    coluna1.extend(df_transposed.iloc[:, 0])
                    coluna2.extend(df_transposed.iloc[:, 1])
                    coluna3.extend(df_transposed.iloc[:, 2])
                    coluna4.extend(df_transposed.iloc[:, 3])
            
                # Atualizar o índice de início do próximo conjunto de linhas
                indice_inicio_conjunto += tamanho_conjunto_linhas

            # Remover linhas vazias da coluna "dia" da planilha razao
            razao = razao.dropna(subset=["DIA"])

            # Resetar o índice
            razao = razao.reset_index(drop=True)

            # Criar um novo DataFrame com as colunas extraídas
            df_extraido = pd.DataFrame({
                'Coluna1': coluna1,
                'Coluna2': coluna2,
                'Coluna3': coluna3,
                'Coluna4': coluna4
            })

            # Concatenar o DataFrame extraído com a planilha original
            razao_concatenada = pd.concat([razao, df_extraido], axis=1)

            # Excluir as 3 últimas linhas da planilha razao
            razao_concatenada = razao_concatenada.iloc[:-3]

            # Resetar o índice
            razao_concatenada = razao_concatenada.reset_index(drop=True)

            razao = razao_concatenada

            razao.rename(columns={'Coluna1': 'Origem'}, inplace=True)
            razao.rename(columns={'Coluna2': 'Categoria'}, inplace=True)
            razao.rename(columns={'Coluna3': 'Lote'}, inplace=True)
            razao.rename(columns={'Coluna4': 'Moeda'}, inplace=True)
            razao = razao.drop(['Nova Coluna','LOTE'], axis=1)
            razao = razao.dropna(axis=1, how='all')

            #Quarta parte do Tratamento

            #Mudando tipo do dado.
            razao['CREDITO'] = razao['CREDITO'].astype(float)
            razao['DEBITO'] = razao['DEBITO'].astype(float)

            #substituindo valores null para 0

            razao['CREDITO']  = razao['CREDITO'] .fillna(0)
            razao['DEBITO']  = razao['DEBITO'] .fillna(0)

            # CONCATENANDO VALORES.

            razao['valores'] = - razao['CREDITO']   + razao['DEBITO']  

            #Removendo colunas

            razao = razao.drop(['CREDITO','DEBITO'], axis=1)

            #Quinta parte do tratamento

            #Remoção dos nomes antes do " : "
            razao['Origem'] = razao['Origem'].str.replace('.*:', '', regex=True)
            razao['Categoria'] = razao['Categoria'].str.replace('.*:', '', regex=True)
            razao['Moeda'] = razao['Moeda'].str.replace('.*:', '', regex=True)


            # Remove tudo antes do caractere ':' na coluna 'Origem'
            razao['Origem'] = razao['Origem'].str.replace('.*:', '', regex=True)

            # Remove tudo antes do caractere ':' na coluna 'Categoria'
            razao['Categoria'] = razao['Categoria'].str.replace('.*:', '', regex=True)

            # Remove tudo antes do caractere ':' na coluna 'Moeda'
            razao['Moeda'] = razao['Moeda'].str.replace('.*:', '', regex=True)

            # Substitui o caractere '.' por '/' na coluna 'DIA'
            razao['DIA'] = razao['DIA'].str.replace('.', '/')

            # Obtém a data de hoje
            data_hoje = date.today()

            # Obtém o ano da data de hoje
            ano = data_hoje.year

            # Converte o ano para string
            ano = str(ano)

            # Concatena o dia do DataFrame com o ano para formar 'dia/mes/ano'
            razao['dia_mes_ano']  = razao['DIA'] + '/' + ano

            # Remove espaços em branco da coluna 'dia_mes_ano'
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace(' ', '')

            # Substitui ex. 'JUN' por ex.'06' na coluna 'dia_mes_ano'
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('JAN', '01')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('FEV', '02')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('MAR', '03')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('ABR', '04')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('MAI', '05')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('JUN', '06')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('JUL', '07')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('AGO', '08')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('SET', '09')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('OUT', '10')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('NOV', '11')
            razao['dia_mes_ano'] = razao['dia_mes_ano'].str.replace('DEZ', '12')

            # Converte a coluna 'dia_mes_ano' para o tipo de dado datetime usando o formato "%d/%m/%Y"
            razao['Data'] = pd.to_datetime(razao['dia_mes_ano'], format="%d/%m/%Y")

            # Remove as colunas 'dia_mes_ano' e 'DIA' do DataFrame
            razao = razao.drop(['dia_mes_ano','DIA'], axis=1)

            # Obtém a coluna 'Data' tratada
            data_tratada = razao['Data']

            # Converte a coluna 'Data' para o tipo de dado datetime e extrai apenas a data, sem a informação do horário
            data_tratada = pd.to_datetime(data_tratada, format="%d/%m/%Y").dt.date

            # Substitui a coluna 'Data' pelo resultado do tratamento
            razao['Data'] = data_tratada

            # organizando data frame

            nova_ordem = ['Estrutura de Contas','Conta','Data','HISTORICO','Lote','valores','UA','C/C','CONTRAP','Origem','Categoria','Moeda']

            razao = razao[nova_ordem]

            # Criar um nome dinâmico para o arquivo tratado
            file_path_sem_extensao = file_path[:-5]
            hora_atual = datetime.datetime.now().strftime("%H%M%S")  # Obtém a hora atual no formato HHMMSS
            nome_arquivo_tratado = f"{file_path_sem_extensao}.{hora_atual}.xlsx" 
            # Salvar o arquivo tratado com o nome dinâmico
            razao.to_excel(nome_arquivo_tratado, sheet_name='Tratado', index=False)
            
            # Exibir mensagem de sucesso
            messagebox.showinfo("Sucesso", "O arquivo foi tratado com sucesso!")
            return True

        except Exception as a:
            # Em caso de erro, exibir mensagem de erro e o motivo específico do erro
            messagebox.showwarning("Erro", "Ocorreu um erro durante o tratamento do arquivo.")
            print("Ocorreu um erro durante o tratamento do arquivo:", str(a))
            return False