#!/usr/bin/env python
# coding: utf-8

# In[11]:


import pandas as pd
import PySimpleGUI as sg
import os
import datetime
from datetime import datetime as dtm
from dateutil.relativedelta import relativedelta

lista = ['Matrícula','Nome','Chave','Gestor','Deletar','Descricao','Horas','Deletar_2']
for x in range(2,37):
    lista.append('Descricao'+str(x))
    lista.append('Horas'+str(x))
    lista.append('Deletar'+str(x))
    


today = datetime.datetime.today()
days = []
month_count = -2
while month_count < 3:
    day = (today - relativedelta(months=month_count)).replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    day = day.strftime("%d/%m/%Y")
    days.append(day)
    month_count += 1


sg.theme('Material2')

menu_def = [      
                ['Ajuda', 'Instruções'], ]    
layout = [
        [sg.Menu(menu_def, tearoff=False)],
        [sg.Text('Insira os arquivos abaixo no formato xlsx :',size=(35, 1))],
        [sg.Text('Mês Referência :', size=(15, 1), auto_size_text=False, justification='right'),
         sg.Combo(list(days),default_value=days[2],key='diames')],
        [sg.Text('Dias :', size=(15, 1), auto_size_text=False, justification='right'),
        sg.Input(key= 'dia')],
        [sg.Text('Arquivo Parâmetro :', size=(15, 1), auto_size_text=False, justification='right'),
        sg.Input(key= 'diretorio1'),sg.FileBrowse('Selecionar')],
        [sg.Text('Arquivo Horas extra :', size=(15, 1), auto_size_text=False, justification='right'),
        sg.Input(key= 'diretorio2'),sg.FileBrowse('Selecionar')],
        [sg.Text('_'  * 80)], 
        [sg.Button('Analisar arquivos'),sg.CloseButton('Fechar')]
]

window = sg.Window('Analisar arquivos', layout)




while True:
    event, values = window.read()
    dia= values['dia']
    diames = values['diames']
    diretorio1 = values['diretorio1']
    diretorio2 = values['diretorio2']
    if event == 'Analisar arquivos':
        
        if dia == '':
            print('A quantidade de dias não foi informada!')
            dia = 1
        dia = int(dia)
        caminho_padrao1 = (diretorio1.split("/"))
        del caminho_padrao1[-1]
        separator = '/'
        caminho_padrao1 = separator.join(caminho_padrao1) + separator
        caminho_padrao2 = (diretorio2.split("/"))
        del caminho_padrao2[-1]
        caminho_padrao2 = separator.join(caminho_padrao2) + separator
        try:

                df_horas = pd.read_excel(diretorio2, header=None, names = (lista))
                df_horas = df_horas.drop(columns=['Deletar','Deletar_2'])
                base = pd.read_excel(diretorio1, sheet_name = 'Parametros')
                filial = pd.read_excel(diretorio1, sheet_name='Jornada')
                orcado = pd.read_excel(diretorio1, sheet_name='Orçado')
                Parametros = pd.read_excel(diretorio1, sheet_name='Rubricas')
                base = base.merge(filial, how = 'left', left_on = 'Descrição Filial', right_on = 'Filial')
                base['Salário_hora'] = (base['Salário/Soldada Base'] / base['Jornada']).round(2)
                lista_colunas = list(df_horas.columns)
                del(lista_colunas[0:6])
                lista_colunas
                df = df_horas[['Matrícula', 'Nome', 'Chave', 'Gestor', 'Descricao', 'Horas']].copy()
                while len(lista_colunas) != 0:
                    dados2 = df_horas[['Matrícula', 'Nome', 'Chave', 'Gestor', lista_colunas[0], lista_colunas[1]]].copy()
                    dados2 = dados2.rename({lista_colunas[0]: 'Descricao', lista_colunas[1]: 'Horas'}, axis=1)
                    del lista_colunas[0:3]
                    df = pd.concat([df,dados2])
                df["Horas"] = df["Horas"].replace("         ",0)
                df = df[df['Horas'] != 0].reset_index(drop=True)
                df = df[df['Horas'].notnull()].reset_index(drop=True)
                df['Descricao'] = df['Descricao'].str.strip()
                Parametros['RUBRICA'] = Parametros['RUBRICA'].str.strip()
                df = df.groupby(["Matrícula","Nome","Chave","Gestor","Descricao"])["Horas"].sum().reset_index()

                df = df.merge(base[['Matrícula','Salário_hora']],on = 'Matrícula', how = 'left')
                df_nulos = df[~df['Salário_hora'].notnull()].groupby(["Matrícula","Nome","Chave","Gestor"])["Horas"].sum().reset_index()
                df = df[df['Salário_hora'].notnull()].reset_index(drop=True)

                df['Valor Hora Extra'] = ''
                for x in df.index:
                        if df.iloc[x, 4] == 'QTD BANCO DE HORAS':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD BANCO DE HORAS' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] * Parametros.loc[y,'DADOS1'] * df.iloc[x, 5]
                    
                        elif df.iloc[x, 4] == 'QTD HORA EXTRA 50%':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA 50%' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] * Parametros.loc[y,'DADOS1'] * df.iloc[x, 5]
                            
                        elif df.iloc[x, 4] == 'QTD HORA EXTRA 100%':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA 100%' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] * Parametros.loc[y,'DADOS1'] * df.iloc[x, 5]
                            
                        elif df.iloc[x, 4] == 'QTD ADIC. NOTURNO':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD ADIC. NOTURNO' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] * Parametros.loc[y,'DADOS1'] * df.iloc[x, 5]
                            
                        elif df.iloc[x, 4] == 'AD  NOT 20% DE H.E 50%':
                            y = Parametros.index[Parametros['RUBRICA'] == 'AD  NOT 20% DE H.E 50%' ].tolist()[0]
                            df.iloc[x, 7] = (df.iloc[x, 6] * Parametros.loc[y,'DADOS1']) * Parametros.loc[y,'DADOS2'] * df.iloc[x, 5]
                        
                        elif df.iloc[x, 4] == 'AD  NOT 20% DE H.E 100%':
                            y = Parametros.index[Parametros['RUBRICA'] == 'AD  NOT 20% DE H.E 100%' ].tolist()[0]
                            df.iloc[x, 7] = (df.iloc[x, 6] * Parametros.loc[y,'DADOS1']) * Parametros.loc[y,'DADOS2']* df.iloc[x, 5]
                        
                        elif df.iloc[x, 4] == 'QTD HE 50% NOTUR':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HE 50% NOTUR' ].tolist()[0]
                            df.iloc[x, 7] = (df.iloc[x, 6] * Parametros.loc[y,'DADOS1']) * df.iloc[x, 5] + Parametros.loc[y,'DADOS2']
                        
                        elif df.iloc[x, 4] == 'QTD HE 100% NOT':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HE 100% NOT' ].tolist()[0]
                            df.iloc[x, 7] = (df.iloc[x, 6] * Parametros.loc[y,'DADOS1']) * df.iloc[x, 5] + Parametros.loc[y,'DADOS2']
                        
                        elif df.iloc[x, 4] == 'QTD HE 75% NOTUR':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HE 75% NOTUR' ].tolist()[0]
                            df.iloc[x, 7] = (df.iloc[x, 6] * Parametros.loc[y,'DADOS1']) * df.iloc[x, 5] + Parametros.loc[y,'DADOS2']
                        
                        elif df.iloc[x, 4] == 'DSR S/ H.E NOT 100%':
                            y = Parametros.index[Parametros['RUBRICA'] == 'DSR S/ H.E NOT 100%' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] * Parametros.loc[y,'DADOS1'] * df.iloc[x, 5]
                            
                        elif df.iloc[x, 4] == 'QTD HORA EXTRA 75%':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA 75%' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] * Parametros.loc[y,'DADOS1'] * df.iloc[x, 5]
                            
                        elif df.iloc[x, 4] == 'QTD HORA EXTRA SOBRE AVISO':
                            y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA SOBRE AVISO' ].tolist()[0]
                            df.iloc[x, 7] = df.iloc[x, 6] / Parametros.loc[y,'DADOS1']  * df.iloc[x, 5]

                y = Parametros[Parametros['RUBRICA'].str.contains("DSR")==True].reset_index().loc[0,'DADOS1']
                
                dsr_banco =  df[df['Descricao'] == 'QTD BANCO DE HORAS'].copy()
                dsr_banco['Valor Hora Extra'] = (dsr_banco['Valor Hora Extra']/y )* dia
                dsr_banco['Descricao'] = 'DSR BANCO DE HORAS'
                
                dsr_ad_not20 = df[df['Descricao'] == 'QTD ADIC. NOTURNO'].copy()
                dsr_ad_not20['Valor Hora Extra'] = (dsr_ad_not20['Valor Hora Extra']/ y) * dia
                dsr_ad_not20['Descricao'] = 'DSR S/ AD NOT 20%'
                
                dsr_ad_not50 = df[df['Descricao'] == 'AD  NOT 20% DE H.E 50%'].copy()
                dsr_ad_not50['Valor Hora Extra'] = (dsr_ad_not50['Valor Hora Extra']/ y) * dia
                dsr_ad_not50['Descricao'] = 'DSR AD NOT 50%'
                
                dsr_ad_not100 = df[df['Descricao'] == 'AD  NOT 20% DE H.E 100%'].copy()
                dsr_ad_not100['Valor Hora Extra'] = (dsr_ad_not100['Valor Hora Extra']/ y)* dia
                dsr_ad_not100['Descricao'] = 'DSR AD NOT 100%'
                
                dsr_she_50 = df[df['Descricao'] == 'QTD HORA EXTRA 50%'].copy()
                dsr_she_50['Valor Hora Extra'] = (dsr_she_50['Valor Hora Extra']/y )* dia
                dsr_she_50['Descricao'] = 'DSR S/ H.E 50%'
                
                dsr_she_100 = df[df['Descricao'] == 'QTD HORA EXTRA 100%'].copy()
                dsr_she_100['Valor Hora Extra'] = (dsr_she_100['Valor Hora Extra']/y )* dia
                dsr_she_100['Descricao'] = 'DSR S/ H.E 100%'
                
                dsr_she_not50 = df[df['Descricao'] == 'QTD HE 50% NOTUR'].copy()
                dsr_she_not50['Valor Hora Extra'] = (dsr_she_not50['Valor Hora Extra']/y) * dia
                dsr_she_not50['Descricao'] = 'DSR S/ H.E NOT 50%'
                
                dsr_she_not100 = df[df['Descricao'] == 'QTD HE 100% NOT'].copy()
                dsr_she_not100['Valor Hora Extra'] = (dsr_she_not100['Valor Hora Extra']/y)* dia
                dsr_she_not100['Descricao'] = 'DSR S/ H.E NOT 100%'

                dsr = pd.concat([dsr_banco,dsr_ad_not20,dsr_ad_not50,dsr_ad_not100,dsr_she_50,dsr_she_100,dsr_she_not50,dsr_she_not100])
    
                dsr = dsr[['Matrícula', 'Nome', 'Chave', 'Gestor', 'Descricao', 'Horas','Valor Hora Extra']]
                df = pd.concat([df,dsr])
                
                df['Valor Hora Extra'] = df['Valor Hora Extra'].astype(float).round(2)
                df = df[['Matrícula', 'Nome', 'Chave', 'Gestor', 'Descricao', 'Horas','Valor Hora Extra']]
                df = df.groupby(["Matrícula","Nome","Chave","Gestor","Descricao"])["Horas",'Valor Hora Extra'].sum().reset_index()
                df = df.merge(base[['Matrícula','Salário_hora','Descrição Centro de Custo','Descrição Filial','Gestor Imediato']],on = 'Matrícula', how = 'left')
                df = df.sort_values(['Nome'], ascending=[True])
                df = df.drop(columns=['Gestor'])
                df.rename(columns={'Descricao': "Rubrica",'Gestor Imediato':'Gestor'},inplace = True)
                hora = dtm.now()
                df = df[['Matrícula', 'Nome', 'Chave', 'Rubrica', 'Horas','Salário_hora', 'Descrição Centro de Custo', 'Descrição Filial','Gestor','Valor Hora Extra']].reset_index(drop = True)
                df['Dia_Mês'] = diames
                df = df.merge(base[['Matrícula','Email do Gestor']],on = 'Matrícula', how = 'left')
                df = df[['Dia_Mês','Matrícula','Nome','Chave','Rubrica','Horas','Salário_hora','Descrição Centro de Custo','Descrição Filial','Gestor','Email do Gestor','Valor Hora Extra']]
                
                realizado = df.groupby(['Gestor','Email do Gestor','Descrição Filial'])["Valor Hora Extra"].sum().reset_index()
                realizado['Orçado'] = 0
                realizado['Dia_Mês'] = diames
                realizado = realizado.rename({'Valor Hora Extra': 'Realizado'}, axis=1)
                
                arquivo = caminho_padrao2 + ('Relatorios gerados {}.{}.{}_hr_{}_{}').format(hora.day,hora.month,hora.year,hora.hour,hora.minute)
                isExist = os.path.exists(arquivo + separator + "Gestores")
                
                
                if not isExist:
                    os.makedirs(arquivo)
                df.to_excel(arquivo + separator + 'Relatório de horas consolidado.xlsx', index = False)
                realizado.to_excel(arquivo + separator + 'Orçado vs realizado.xlsx',index = False)
                if len(df_nulos) > 0:
                    df_nulos.to_excel(arquivo + separator + 'Funcionários não cadastrados.xlsx', index = False)
                    sg.popup('Atenção!!! Foi indentificado que alguns funcionários não estão cadastrados da planila "Base", um arquivo foi gerado com os nomes não cadastrados!')
                
                
                isExist = os.path.exists(arquivo + separator + "Gestores")

                if not isExist:
                    os.makedirs(arquivo + separator + "Gestores")
                    
                lista_gestores = list(df['Gestor'].drop_duplicates())                
                for gestor in lista_gestores:
                    df_gestor = df[df['Gestor'] == gestor]
                    df_gestor.to_excel(arquivo + separator + "Gestores" + separator + gestor.strip() + '.xlsx',index=False)
                
                
                window.close()
                break
                
        except FileNotFoundError:
            sg.popup('Selecione os dois arquivos no formato xlsx!')
            break
            window.close()
        except ValueError:
            sg.popup('O arquivo está fora do padrão esperado.')
            break
            window.close()
            
    elif event == 'Instruções': 
       
        sg.popup('Para utilizar o programa , siga as seguintes instruções :\n • Selecione a pasta aonde estão salvos apenas os arquivos Xlsx. \n • Clique em analisar arquivos.  \n • No final da análise será gerado um arquivo excel, contendo os erros presentes em cada relatório.')
    elif event == 'Fechar':
        break
        window.close()
        
    else:
        break
        window.close()

