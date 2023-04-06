import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from tkcalendar import DateEntry
import pandas as pd
import yfinance as yf
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import smtplib
import re
import awesometkinter as atk

class App:
    def __init__(self):
        self.janela = tk.Tk()
        self.janela.title("Historico de Ações")
        self.janela.geometry("600x250")
        self.janela.resizable(False, False)
        self.janela.config(bg="#333333")
        self.estilo = ttk.Style(self.janela)
        self.estilo.configure('TEntry', width=20)
        self.frame = atk.Frame3d(self.janela)
        self.frame.grid(row=0, column=0, padx=10, pady=5, sticky='w', columnspan=6)
        self.frame.configure(width=90)

        #Funão verifica que o e_mail digitado é valido .
        def validar_email(email):
            padrao = r'^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$'
            if re.match(padrao, email) :
                return True
            else:
                return False

        #Função pega o key da ação selecionada
        def pegar_acao_selecionada():
            lista_selecionada = self.lista_acoes.get()
            acao_selecionada = self.dic_carteira[lista_selecionada]
            return acao_selecionada

        #Faz a consulta da ação selecionada e retorna o resultado
        def pegar_dados_ibov():
            data_inicio = self.data_inicial.get_date()
            data_fim = self.data_final.get_date()
            selecionada = pegar_acao_selecionada()
            if selecionada == '^BVSP':
                cotacao = yf.download('AAPL', start=data_inicio.strftime('%Y-%m-%d'),
                                      end=data_fim.strftime('%Y-%m-%d'))
            else:
                cotacao = yf.download(f'{selecionada}.SA', start=data_inicio.strftime('%Y-%m-%d'),
                                      end=data_fim.strftime('%Y-%m-%d'))

            return cotacao

        #Faz um download da Ação selecionada com periodo també selecionado com as colunas renomeadas
        def gerar_excel_acoes():
            dados_acoes = pegar_dados_ibov()
            selecionada = pegar_acao_selecionada()
            dados_acoes.to_excel(f'{selecionada}.xlsx')
            dados = pd.read_excel(f'{selecionada}.xlsx')
            dados.rename(
                columns={"Date": "Dt.Abertura",
                         "High": "Alta",
                         "Low": "Baixa",
                         "Open": "Aberto",
                         "Close": "Fechamento",
                         "Volume": "Volume",
                         "Adj Close": "Fechado"}
            )
            messagebox.showinfo('Mensagem', "Gerada com Sucesso")

        #Gera um grafico da ação seleciona da periodo
        def gerar_grafico_acoes():
            acao= self.lista_acoes.get()
            dados_acoes = pegar_dados_ibov()
            dados_acoes["Adj Close"].plot(figsize=(10,5), color ='g', label='Fechamento',lw="2")
            #plt.ion()
            plt.title(f'Grafico do historico da ({acao}) ')
            plt.xlabel("Data")
            plt.ylabel("Preço (R$)")
            plt.legend(loc= 1)
            plt.savefig(f'{acao}.png')
            plt.show()

        #Função envia eMail com Anexo de um planilha gerada pela função  gerar_excel_acoes():
        def enviar_email_acoes():
            email_valido = validar_email(self.email_text.get())
            acao_selecionada =pegar_acao_selecionada()
            email_destino = self.email_text.get()
            if email_valido:
                username = 'colocar e_mail aqui'
                password = 'colocar senha aqui'

                # Cria a mensagem
                msg = MIMEMultipart()
                msg['From'] = username
                msg['To'] = email_destino
                msg['Subject'] = f'Segue {acao_selecionada}'

                # Adiciona o texto da mensagem
                texto = 'Olá, segue em anexo um arquivo importante.'
                part_texto = MIMEText(texto, 'plain')
                msg.attach(part_texto)

                # Adiciona o arquivo em anexo

                arquivo = f'{acao_selecionada}.xlsx'
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(open(arquivo, 'rb').read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(arquivo))
                msg.attach(part)

                # Conecta-se ao servidor SMTP do Gmail e envia a mensagem
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(username, password)
                    smtp.sendmail(username, {email_destino}, msg.as_string())

                messagebox.showinfo('Sucesso',f'Email Enviado para : {email_destino}, verifique a Caixa de Entrada')
            else:
                messagebox.showwarning('Erro', f'Email digitado : {email_destino} Invalido')



        # Cria o título da janela
        titulo = tk.Label(self.frame, text="Selecione a Ação desejada", bg="#333333", fg='white',font=("Arial", 12))
        titulo.grid(row=1, column=0, padx=10, pady=5, sticky='w', columnspan=2)
         #Cria um frame dentro da janela

        # Cria a combobox para selecionar a carteira
        self.dic_carteira = {'ITAÚ - (ITUB4)': 'ITUB4', 'PETROBRÁS -(PETR4)': 'PETR4', 'BRADESCO - (BBDC4)': 'BBDC4',
                             'MAGALU - (MGLU3) ': 'MGLU3', 'IBOVESPA - (^BVSP)': '^BVSP', 'AMBEV - (ABEV3)': 'ABEV3',
                             'VALE-(VALE3)': 'VALE3', 'GERDAU - (GGBR4)': 'GGBR4'}
        lista_carteira = list(self.dic_carteira.keys())
        self.lista_acoes = ttk.Combobox(self.frame, values=lista_carteira,font=("Arial", 12))
        self.lista_acoes.grid(row=2, column=0, padx=10, pady=5, sticky='w', columnspan=6)
        self.lista_acoes.current(0)
        self.lista_acoes.configure(width=60)

        # cria a label e a entrada de data inicial dentro do fra,font=("Arial", 12))me
        l_datainicial = tk.Label(self.frame, text="Selecione a Data Inicial:", bg="#333333", fg='white',font=("Arial", 12))
        l_datainicial.grid(row=3, column=0, padx=10, pady=5, sticky='w')
        self.data_inicial = DateEntry(self.frame, date_pattern='dd/mm/yyyy',font=("Arial", 12))
        self.data_inicial.grid(row=4, column=0, padx=10, pady=5, sticky='w')
        self.data_inicial.configure(width=20)

        # cria a label e a entrada de data final dentro do frame
        l_datafinal = tk.Label(self.frame, text="Selecione a Data Final:", bg="#333333", fg='white',font=("Arial", 12))
        l_datafinal.grid(row=3, column=4, padx=10, pady=5, sticky='w')
        self.data_final = DateEntry(self.frame, date_pattern='dd/mm/yyyy',font=("Arial", 12))
        self.data_final.grid(row=4, column=4, padx=10, pady=5, sticky='w')
        self.data_final.configure(width=20)


        # Cria o botão para gerar a planilha

        self.botao_planilha = atk.Button3d(self.frame , text="Salvar Planilha", command=gerar_excel_acoes)
        self.botao_planilha.grid(row=5, column=0, padx=10, pady=10,sticky='w',  columnspan=2)
        self.botao_planilha.configure(width=20)

        # Cria o botão para gerar o gráfico
        botao_grafico = atk.Button3d(self.frame, text="Visualizar Gráfico", command=gerar_grafico_acoes)
        botao_grafico.grid(row=5, column=4, padx=10, pady=5, sticky='w')
        botao_grafico.configure(width=20)

       #Define o foco no campo da lista de tikers
        self.lista_acoes.focus()
        # Inicia a janela principal
        self.janela.mainloop()
if __name__ == '__main__':
    app = App()

