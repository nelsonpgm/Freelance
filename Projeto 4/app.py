#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os 
import pandas as pd
import PySimpleGUI as sg
from datetime import datetime
import numpy as np

import os 
import pandas as pd
import PySimpleGUI as sg
from datetime import datetime
import numpy as np


def case_450(row): 
# USADO PARA OS PEDIDOS - separa os pedidios da coluna texto filtrando pelo inicio 450 e contando 10 caracteres totais    
    if '450' in str(row['Texto']):
        
        texto = str(row['Texto'])
        posicao_450 = texto.find('450')
        val = texto[posicao_450:posicao_450+10]
        
    else:
        
        val = np.nan
        
    return val


print=sg.Print
sg.theme('DarkBlue12')

menu_def = [      
                ['Ajuda', 'Instruções'], ]    
layout = [
        [sg.Menu(menu_def, tearoff=False)],
        [sg.Text('Insira os arquivos abaixo no formato xlsx :',size=(35, 1))],
        [sg.Text('Planilha Excel .xlsx', size=(15, 1), auto_size_text=False, justification='right'),
        sg.Input(key= 'diretorio'),sg.FileBrowse('Selecionar')],
        [sg.Text('_'  * 80)], 
        [sg.Button('Analisar partidas')]
]


# Display the window and get values

window = sg.Window('Analisar partidas', layout)
event, values = window.read()
diretorio = values['diretorio']

while True:
    event, value = window.read()
    if event == 'Analisar partidas':
        inicio = ('  Horário de início {}:{} - {}/{}/{} '.format(datetime.now().hour, datetime.now().minute,datetime.now().day,datetime.now().month,datetime.now().year ))
        print(inicio)
        print('[+] Importando planilhas do excel, por favor aguarde...')
        
        relat_aberto_e_comp = pd.read_excel(diretorio)
        Mapeando_proposta = pd.read_excel('local/folder/archive')
        Mapeando_proposta =  Mapeando_proposta[['COD_SAP','CONSOLIDADO','CATEGORIA MACRO (PRINCIPAIS FORNECEDORES)','PREVISÃO DE PAGAMENTO']]
        Mapeando_proposta.rename(columns={'CONSOLIDADO': 'RAZAO_SOCIAL'},inplace = True)
        Tipo_de_docs = pd.read_excel('G:\\Ctpg\\GESTÃO\\CONSULTA FORNECEDOR\\Tipo de documentos SAP e Área.xlsx')
        print('  [-] Iniciando o cruzamento de dados...')
        relat_aberto_e_comp.rename(columns={'Conta':'COD_SAP'}, inplace = True)
        relat_aberto_e_comp['Montante em moeda interna'] = relat_aberto_e_comp['Montante em moeda interna'].round(2)
        relat_aberto_e_comp.dropna(subset=["COD_SAP"],inplace = True)
        relat_aberto_e_comp.COD_SAP = relat_aberto_e_comp.COD_SAP.astype(int)

        relat_aberto_e_comp = pd.merge(relat_aberto_e_comp, Mapeando_proposta, on=['COD_SAP'], how='left')


        Partidas_em_aberto = relat_aberto_e_comp[relat_aberto_e_comp['Doc.compensação'].isnull()].reset_index(drop= True)
        df = relat_aberto_e_comp[relat_aberto_e_comp['Doc.compensação'].notnull()].reset_index(drop= True)

        primeiro_caracter = []
        for i in df['Nº documento']:
            i = str(i)
            i = int(i[0])
            primeiro_caracter.append(i)
        df['primeiro_caracter_numero_doc'] = primeiro_caracter
        primeiro_caracter = []
        for i in df['Doc.compensação']:
            i = str(i)
            i = int(i[0])
            primeiro_caracter.append(i)
        df['primeiro_caracter_numero_doc_compensacao'] = primeiro_caracter
        # Telemont - tirar toda referecia que for = ADIANTAMENTO
        Telemont = (df[df['RAZAO_SOCIAL'].str.contains('TELEMONT', na = False)]).reset_index(drop=True)
        Telemont['RAZAO_SOCIAL'] = 'TELEMONT'
        mask = (Telemont['Referência'] != 'ADIANTAMENTO')
        Telemont = Telemont.loc[mask].reset_index(drop = True)
        Telemont['Nº documento'].replace(np.nan, '-')
        Telemont['Doc.compensação'].replace(np.nan, '-')
        print('  [-] Ajustando tabelas...')

        # FILTRAR SO O MEIO DE PAGAMENTO V - DATAFRAME - BAIXA DO ADIANTAMENTO
        mask = (Telemont['FrmPgto'] == 'V')
        Telemont_Baixa_adiantamento = Telemont.loc[mask].reset_index(drop = True)
        # EXLUIR EM UM NOVO DF, TODO NUMERO DE DOCUMENTO DA TELEMONT QUE COMEÇA COM 2*  E DEIXAR SÓ O DOC DE COMPENSAÇÃO QUE COMEÇA COM 2
        mask = (Telemont['primeiro_caracter_numero_doc'] != 2)
        Telemont_2 = Telemont.loc[mask].reset_index(drop = True)
        mask = (Telemont['primeiro_caracter_numero_doc_compensacao'] == 2)
        Telemont_2 = Telemont.loc[mask].reset_index(drop = True)
        y = Telemont_Baixa_adiantamento['Montante em moeda interna'].sum()
        x = Telemont_2['Montante em moeda interna'].sum()
        Telemont = pd.concat([Telemont_Baixa_adiantamento, Telemont_2]).reset_index(drop = True)

        Serede = (df[df['RAZAO_SOCIAL'].str.contains('SEREDE', na = False)]).reset_index(drop=True)
        Serede['RAZAO_SOCIAL'] = 'SEREDE'
        mask = (Serede['Referência'] != 'ADIANTAMENTO')
        Serede = Serede.loc[mask].reset_index(drop = True)

        mask = (Serede['Referência'] != 'TRANSFERENCIA')
        Serede = Serede.loc[mask].reset_index(drop = True)

        mask = (Serede['FrmPgto'] == 'V')
        Serede_baixa_de_adiantamento = Serede.loc[mask].reset_index(drop = True)

        Serede_baixa_de_adiantamento2 = Serede[Serede['Atribuição'].str.contains('BX ADTO', na = False)]
        Serede_baixa_de_adiantamento2 = Serede_baixa_de_adiantamento2.reset_index(drop = True)
        baixa_adto = pd.concat([Serede_baixa_de_adiantamento, Serede_baixa_de_adiantamento2]).reset_index(drop = True)

        mask = (Serede['primeiro_caracter_numero_doc'] != 2)
        Serede_2 = Serede.loc[mask].reset_index(drop = True)
        mask = (Serede['primeiro_caracter_numero_doc_compensacao'] == 2)
        Serede_2 = Serede.loc[mask].reset_index(drop = True)

        Serede_2 = Serede_2[Serede_2["Atribuição"].str.contains('BX ADTO') == False]
        Serede_2.reset_index(drop = True)



        Serede = pd.concat([baixa_adto, Serede_2]).reset_index(drop = True)
        Relatorio = pd.concat([Serede, Telemont]).reset_index(drop = True)

        df2 = df[df["RAZAO_SOCIAL"].str.contains('TELEMONT') == False]
        df2 = df2.reset_index(drop = True)
        df2 = df2[df2["RAZAO_SOCIAL"].str.contains('SEREDE') == False]
        df2 = df2.reset_index(drop = True)

        print('  [-] Ajustando parâmetros e exportando arquivo excel...')

        Relatorio = pd.concat([Relatorio, df2]).reset_index(drop = True)
        Relatorio = Relatorio.drop_duplicates()
        Relatorio = Relatorio.reset_index(drop = True)
        Relatorio = Relatorio[Relatorio['Bloqueio pgto.'].isnull()]
        Relatorio.drop(columns=['primeiro_caracter_numero_doc','primeiro_caracter_numero_doc_compensacao'],inplace = True)
        Partidas_compensadas = pd.merge(Relatorio, Tipo_de_docs, on=['Tipo de documento'], how='left')
        Partidas_em_aberto = pd.merge(Partidas_em_aberto, Tipo_de_docs, on=['Tipo de documento'], how='left')
        Partidas_compensadas['Tipo de partida'] = 'Compensada'
        Partidas_em_aberto['Tipo de partida'] = 'Em aberto'
        Relatorio_full = pd.concat([Partidas_compensadas,Partidas_em_aberto]).reset_index(drop = True)    
        caminho_padrao = (diretorio.split("/"))
        del caminho_padrao[-1]
        separator = '/'
        caminho_padrao = separator.join(caminho_padrao) + separator
        text_naoloc = caminho_padrao + 'Fornecedores não localizados.xlsx'
        texto_relat = caminho_padrao + 'Partidas aberta e compensadas - Tratado.xlsx'
        Relatorio_full['Empresa_co'] = Relatorio_full['Empresa'].replace(['BTSA'],'ClientCo').replace(['SMPE'],'ClientCo').replace(['MRED'],'InfraCo')
        Relatorio_full['Pedido'] = Relatorio_full.apply(case_450, axis=1)
        Relatorio_full = Relatorio_full[['FrmPgto', 'Dias 1', 'Empresa', 'Nº documento', 'COD_SAP',
               'Dt.base prazo pgto.', 'Data de pagamento', 'Data do documento',
               'Referência', 'Bloqueio pgto.', 'Montante em moeda interna',
               'Data de lançamento', 'Texto','Pedido','Atribuição', 'Data de compensação',
               'Doc.compensação', 'Tipo de documento', 'RAZAO_SOCIAL',
               'CATEGORIA MACRO (PRINCIPAIS FORNECEDORES)', 'PREVISÃO DE PAGAMENTO',
               'Pre editado ', 'Denominação', 'Área','Empresa_co']]
        Relatorio_full.to_excel(texto_relat,index=False)
        naocadastrados = Relatorio_full[Relatorio_full['RAZAO_SOCIAL'].isnull()]
        naocadastrados = naocadastrados[['COD_SAP']].reset_index(drop = True)
        naocadastrados.drop_duplicates(subset=['COD_SAP'], keep='first',inplace = True)
        naocadastrados.to_excel(text_naoloc,index=False)
        print('Análise Finalizada!')

        
    elif event == 'Instruções': 
        print('Para utilizar o programa de gerar prévia, siga as seguintes instruções :\n • Insira o caminho de rede aonde está salvo a base SAP. \n • Insira os arquivo aonde indica "Planilha Shaepoint".  \n • O arquivo será gerado em html')

        
    else:
        break
        window.close()
