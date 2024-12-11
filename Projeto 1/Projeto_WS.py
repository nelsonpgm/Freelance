import pandas as pd
import os
import datetime
from datetime import datetime as dtm
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Analisar arquivos")

        # Variáveis para armazenar arquivos selecionados e parâmetros
        self.diretorio1 = None
        self.diretorio2 = None
        self.diames = None
        self.dia = None

        self.create_interface()

    def create_interface(self):
        # Menu
        menubar = tk.Menu(self)
        ajuda_menu = tk.Menu(menubar, tearoff=False)
        ajuda_menu.add_command(label="Instruções", command=self.instrucoes)
        menubar.add_cascade(label="Ajuda", menu=ajuda_menu)
        self.config(menu=menubar)

        # Cálculo de datas para combo
        today = datetime.datetime.today()
        days = []
        month_count = -2
        while month_count < 3:
            day = (today - relativedelta(months=month_count)).replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            day = day.strftime("%d/%m/%Y")
            days.append(day)
            month_count += 1

        frame = tk.Frame(self, padx=10, pady=10)
        frame.pack(fill='both', expand=True)

        tk.Label(frame, text="Insira os arquivos abaixo no formato XLSX :").grid(row=0, column=0, columnspan=3, pady=5, sticky='w')

        tk.Label(frame, text="Mês Referência :", width=15, anchor='e').grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.combo_diames = ttk.Combobox(frame, values=days)
        self.combo_diames.set(days[2])
        self.combo_diames.grid(row=1, column=1, padx=5, pady=5, sticky='w')

        tk.Label(frame, text="Dias :", width=15, anchor='e').grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.entry_dias = tk.Entry(frame)
        self.entry_dias.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        tk.Label(frame, text="Arquivo Parâmetro :", width=15, anchor='e').grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.entry_parametro = tk.Entry(frame, width=50)
        self.entry_parametro.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        tk.Button(frame, text="Selecionar", command=self.selecionar_arquivo_parametro).grid(row=3, column=2, padx=5, pady=5)

        tk.Label(frame, text="Arquivo Horas extra :", width=15, anchor='e').grid(row=4, column=0, padx=5, pady=5, sticky='e')
        self.entry_horas = tk.Entry(frame, width=50)
        self.entry_horas.grid(row=4, column=1, padx=5, pady=5, sticky='w')
        tk.Button(frame, text="Selecionar", command=self.selecionar_arquivo_horas).grid(row=4, column=2, padx=5, pady=5)

        tk.Label(frame, text="_"*80).grid(row=5, column=0, columnspan=3, pady=10)

        frame_buttons = tk.Frame(frame)
        frame_buttons.grid(row=6, column=0, columnspan=3, pady=5)

        tk.Button(frame_buttons, text="Analisar arquivos", command=self.analisar_arquivos).pack(side='left', padx=5)
        tk.Button(frame_buttons, text="Fechar", command=self.fechar).pack(side='left', padx=5)

    def instrucoes(self):
        messagebox.showinfo(
            "Instruções",
            "Para utilizar o programa, siga as seguintes instruções:\n"
            "• Selecione os dois arquivos no formato XLSX.\n"
            "• Selecione o mês de referência e informe a quantidade de dias.\n"
            "• Clique em 'Analisar arquivos'.\n"
            "• Ao final, será gerado um arquivo Excel com os resultados e possíveis erros."
        )

    def selecionar_arquivo_parametro(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if filepath:
            self.entry_parametro.delete(0, tk.END)
            self.entry_parametro.insert(0, filepath)

    def selecionar_arquivo_horas(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if filepath:
            self.entry_horas.delete(0, tk.END)
            self.entry_horas.insert(0, filepath)

    def analisar_arquivos(self):
        self.dia = self.entry_dias.get()
        self.diames = self.combo_diames.get()
        self.diretorio1 = self.entry_parametro.get()
        self.diretorio2 = self.entry_horas.get()

        if not self.diretorio1 or not self.diretorio2:
            messagebox.showwarning('Atenção', 'Selecione os dois arquivos no formato XLSX!')
            return

        if self.dia == '':
            print('A quantidade de dias não foi informada! Será considerado 1.')
            self.dia = 1

        try:
            self.dia = int(self.dia)
        except ValueError:
            messagebox.showwarning("Atenção", "O campo 'Dias' deve ser um número inteiro. Assumindo valor 1.")
            self.dia = 1

        try:
            # ----------------- Início da lógica original -----------------
            lista = ['Matrícula','Nome','Chave','Gestor','Deletar','Descricao','Horas','Deletar_2']
            for x in range(2,37):
                lista.append('Descricao'+str(x))
                lista.append('Horas'+str(x))
                lista.append('Deletar'+str(x))

            caminho_padrao1 = (self.diretorio1.split("/"))
            del caminho_padrao1[-1]
            separator = '/'
            caminho_padrao1 = separator.join(caminho_padrao1) + separator

            caminho_padrao2 = (self.diretorio2.split("/"))
            del caminho_padrao2[-1]
            caminho_padrao2 = separator.join(caminho_padrao2) + separator

            df_horas = pd.read_excel(self.diretorio2, header=None, names=lista)
            df_horas = df_horas.drop(columns=['Deletar','Deletar_2'])
            base = pd.read_excel(self.diretorio1, sheet_name='Parametros')
            filial = pd.read_excel(self.diretorio1, sheet_name='Jornada')
            orcado = pd.read_excel(self.diretorio1, sheet_name='Orçado')
            Parametros = pd.read_excel(self.diretorio1, sheet_name='Rubricas')

            base = base.merge(filial, how='left', left_on='Descrição Filial', right_on='Filial')
            # Mantendo a mesma lógica de cálculo do salário/hora do código original atualizado
            base['Salário_hora'] = (base['Remuneração Total'] / base['Jornada']).round(2)

            lista_colunas = list(df_horas.columns)
            del(lista_colunas[0:6])

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

            df = df.merge(base[['Matrícula','Salário_hora']],on='Matrícula', how='left')
            df_nulos = df[~df['Salário_hora'].notnull()].groupby(["Matrícula","Nome","Chave","Gestor"])["Horas"].sum().reset_index()
            df = df[df['Salário_hora'].notnull()].reset_index(drop=True)

            df['Valor Hora Extra'] = ''
            for x in df.index:
                desc = df.iloc[x,4]
                sal_hora = df.iloc[x,6]
                horas = df.iloc[x,5]

                if desc == 'QTD BANCO DE HORAS':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD BANCO DE HORAS' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * Parametros.loc[y,'DADOS1'] * horas

                elif desc == 'QTD HORA EXTRA 50%':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA 50%' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * Parametros.loc[y,'DADOS1'] * horas

                elif desc == 'QTD HORA EXTRA 100%':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA 100%' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * Parametros.loc[y,'DADOS1'] * horas

                elif desc == 'QTD ADIC. NOTURNO':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD ADIC. NOTURNO' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * Parametros.loc[y,'DADOS1'] * horas

                elif desc == 'AD  NOT 20% DE H.E 50%':
                    y = Parametros.index[Parametros['RUBRICA'] == 'AD  NOT 20% DE H.E 50%' ].tolist()[0]
                    df.iloc[x,7] = (sal_hora * Parametros.loc[y,'DADOS1']) * Parametros.loc[y,'DADOS2'] * horas

                elif desc == 'AD  NOT 20% DE H.E 100%':
                    y = Parametros.index[Parametros['RUBRICA'] == 'AD  NOT 20% DE H.E 100%' ].tolist()[0]
                    df.iloc[x,7] = (sal_hora * Parametros.loc[y,'DADOS1']) * Parametros.loc[y,'DADOS2'] * horas

                elif desc == 'QTD HE 50% NOTUR':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HE 50% NOTUR' ].tolist()[0]
                    df.iloc[x,7] = (sal_hora * Parametros.loc[y,'DADOS1']) * horas + Parametros.loc[y,'DADOS2']

                elif desc == 'QTD HE 100% NOT':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HE 100% NOT' ].tolist()[0]
                    df.iloc[x,7] = (sal_hora * Parametros.loc[y,'DADOS1']) * horas + Parametros.loc[y,'DADOS2']

                elif desc == 'QTD HE 75% NOTUR':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HE 75% NOTUR' ].tolist()[0]
                    df.iloc[x,7] = (sal_hora * Parametros.loc[y,'DADOS1']) * horas + Parametros.loc[y,'DADOS2']

                elif desc == 'DSR S/ H.E NOT 100%':
                    y = Parametros.index[Parametros['RUBRICA'] == 'DSR S/ H.E NOT 100%' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * Parametros.loc[y,'DADOS1'] * horas

                elif desc == 'QTD HORA EXTRA 75%':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA 75%' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * Parametros.loc[y,'DADOS1'] * horas

                elif desc == 'QTD HORA EXTRA SOBRE AVISO':
                    y = Parametros.index[Parametros['RUBRICA'] == 'QTD HORA EXTRA SOBRE AVISO' ].tolist()[0]
                    df.iloc[x,7] = sal_hora * (1/Parametros.loc[y,'DADOS1']) * horas

            y = Parametros[Parametros['RUBRICA'].str.contains("DSR")==True].reset_index().loc[0,'DADOS1']

            dsr_banco =  df[df['Descricao'] == 'QTD BANCO DE HORAS'].copy()
            dsr_banco['Valor Hora Extra'] = (dsr_banco['Valor Hora Extra']/y )* self.dia
            dsr_banco['Descricao'] = 'DSR BANCO DE HORAS'

            dsr_ad_not20 = df[df['Descricao'] == 'QTD ADIC. NOTURNO'].copy()
            dsr_ad_not20['Valor Hora Extra'] = (dsr_ad_not20['Valor Hora Extra']/ y) * self.dia
            dsr_ad_not20['Descricao'] = 'DSR S/ AD NOT 20%'

            dsr_ad_not50 = df[df['Descricao'] == 'AD  NOT 20% DE H.E 50%'].copy()
            dsr_ad_not50['Valor Hora Extra'] = (dsr_ad_not50['Valor Hora Extra']/ y) * self.dia
            dsr_ad_not50['Descricao'] = 'DSR AD NOT 50%'

            dsr_ad_not100 = df[df['Descricao'] == 'AD  NOT 20% DE H.E 100%'].copy()
            dsr_ad_not100['Valor Hora Extra'] = (dsr_ad_not100['Valor Hora Extra']/ y)* self.dia
            dsr_ad_not100['Descricao'] = 'DSR AD NOT 100%'

            dsr_she_50 = df[df['Descricao'] == 'QTD HORA EXTRA 50%'].copy()
            dsr_she_50['Valor Hora Extra'] = (dsr_she_50['Valor Hora Extra']/y )* self.dia
            dsr_she_50['Descricao'] = 'DSR S/ H.E 50%'

            dsr_she_100 = df[df['Descricao'] == 'QTD HORA EXTRA 100%'].copy()
            dsr_she_100['Valor Hora Extra'] = (dsr_she_100['Valor Hora Extra']/y )* self.dia
            dsr_she_100['Descricao'] = 'DSR S/ H.E 100%'

            dsr_she_not50 = df[df['Descricao'] == 'QTD HE 50% NOTUR'].copy()
            dsr_she_not50['Valor Hora Extra'] = (dsr_she_not50['Valor Hora Extra']/y) * self.dia
            dsr_she_not50['Descricao'] = 'DSR S/ H.E NOT 50%'

            dsr_she_not100 = df[df['Descricao'] == 'QTD HE 100% NOT'].copy()
            dsr_she_not100['Valor Hora Extra'] = (dsr_she_not100['Valor Hora Extra']/y)* self.dia
            dsr_she_not100['Descricao'] = 'DSR S/ H.E NOT 100%'

            dsr = pd.concat([dsr_banco,dsr_ad_not20,dsr_ad_not50,dsr_ad_not100,dsr_she_50,dsr_she_100,dsr_she_not50,dsr_she_not100])
            dsr = dsr[['Matrícula', 'Nome', 'Chave', 'Gestor', 'Descricao', 'Horas','Valor Hora Extra']]
            df = pd.concat([df,dsr])

            df['Valor Hora Extra'] = df['Valor Hora Extra'].astype(float).round(2)
            df = df[['Matrícula', 'Nome', 'Chave', 'Gestor', 'Descricao', 'Horas','Valor Hora Extra']]

            df = df.groupby(["Matrícula","Nome","Chave","Gestor","Descricao"], as_index=False).agg({"Horas":"sum","Valor Hora Extra":"sum"})
            df = df.merge(base[['Matrícula','Salário_hora','Descrição Centro de Custo','Descrição Filial','Gestor Imediato']],on='Matrícula', how='left')
            df = df.sort_values(['Nome'], ascending=[True])
            df = df.drop(columns=['Gestor'])
            df.rename(columns={'Descricao': "Rubrica",'Gestor Imediato':'Gestor'},inplace=True)
            hora = dtm.now()
            df = df[['Matrícula', 'Nome', 'Chave', 'Rubrica', 'Horas','Salário_hora', 'Descrição Centro de Custo', 'Descrição Filial','Gestor','Valor Hora Extra']].reset_index(drop=True)
            df['Dia_Mês'] = self.diames
            df = df.merge(base[['Matrícula','Email do Gestor']],on='Matrícula', how='left')
            df = df[['Dia_Mês','Matrícula','Nome','Chave','Rubrica','Horas','Salário_hora','Descrição Centro de Custo','Descrição Filial','Gestor','Email do Gestor','Valor Hora Extra']]

            realizado = df.groupby(['Gestor','Email do Gestor','Descrição Filial'], as_index=False)["Valor Hora Extra"].sum()
            realizado['Orçado'] = 0
            realizado['Dia_Mês'] = self.diames
            realizado = realizado.rename({'Valor Hora Extra': 'Realizado'}, axis=1)

            arquivo = caminho_padrao2 + ('Relatorios gerados {}.{}.{}_hr_{}_{}').format(hora.day,hora.month,hora.year,hora.hour,hora.minute)
            isExist = os.path.exists(arquivo)

            if not isExist:
                os.makedirs(arquivo)
            df.to_excel(arquivo + separator + 'Relatório de horas consolidado.xlsx', index=False)
            realizado.to_excel(arquivo + separator + 'Orçado vs realizado.xlsx', index=False)

            if len(df_nulos) > 0:
                df_nulos.to_excel(arquivo + separator + 'Funcionários não cadastrados.xlsx', index=False)
                messagebox.showinfo('Atenção!!!', 'Foi identificado que alguns funcionários não estão cadastrados na planilha "Base". Um arquivo foi gerado com os nomes não cadastrados!')

            isExist = os.path.exists(arquivo + separator + "Gestores")

            if not isExist:
                os.makedirs(arquivo + separator + "Gestores")

            lista_gestores = list(df['Gestor'].drop_duplicates())
            for gestor in lista_gestores:
                df_gestor = df[df['Gestor'] == gestor]
                df_gestor.to_excel(arquivo + separator + "Gestores" + separator + gestor.strip() + '.xlsx', index=False)

            messagebox.showinfo('Sucesso', 'Análise concluída com sucesso! Arquivos gerados na pasta:\n' + arquivo)
            self.fechar()
            # ----------------- Fim da lógica original -----------------

        except FileNotFoundError:
            messagebox.showerror('Erro', 'Selecione os dois arquivos no formato XLSX!')
        except ValueError:
            messagebox.showerror('Erro', 'O arquivo está fora do padrão esperado.')

    def fechar(self):
        """Fecha o programa."""
        self.destroy()
        exit()

if __name__ == "__main__":
    app = App()
    app.mainloop()
