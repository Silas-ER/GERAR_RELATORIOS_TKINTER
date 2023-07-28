# importação das bibliotecas
import pandas as pd
import pyodbc 
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from tkcalendar import DateEntry
from datetime import datetime

# conexão com o banco de dados
connect = pyodbc.connect('Driver={SQL Server};' 'Server=SQL;' 'Database=DATA;' 'UID=USER;' 'PWD=PASSWORD;')
cursor = connect.cursor()

#classe de botões 
button_style = {"fg": "black", "font": ("Arial", 10)}

# criação da classe de visual da aplicação criada
class Aplicacao(tk.Tk):
    #init de comum de programas
    def __init__(self):
        super().__init__()  # init da classe pai Tk 
        self.title("Custo de Barcos")  # título do programa
        self.chave_entry = None

        self.widgets()  # inicia os componentes da interface gráfica 
        self.geometry("900x600")

    #criação de um cabeçalho de imagem para manter o código organizado e vai ser utilizado em outras funções
    def head(self):
        logo = Image.open("img/logo.png")
        logo = logo.resize((600, 150), Image.ANTIALIAS)
        self.logo_tk = ImageTk.PhotoImage(logo)

        label_logo = tk.Label(self, image=self.logo_tk)
        label_logo.pack()
        
    #widgets o que aparece em tela 
    def widgets(self):
        self.limpar_tela()

        self.head()
        # Criação de um Frame para conter os botões e centralizá-los
        frame_botoes = tk.Frame(self)
        frame_botoes.pack()

        botao_servicos = tk.Button(frame_botoes, text="Consultar valores de material por barco", command=self.materiais)
        botao_servicos.config(**button_style)
        botao_servicos.pack(padx=10,pady=10)

        botao_materiais = tk.Button(frame_botoes, text="Consultar valores de serviços por barco", command=self.servicos)
        botao_materiais.config(**button_style)
        botao_materiais.pack(padx=10, pady=10)

        botao_sair = tk.Button(frame_botoes, text="Sair do Programa", command=self.fechar_programa)
        botao_sair.config(**button_style)
        botao_sair.pack(padx=10, pady=10)

    #fazer consulta dos serviços 
    def gerar_servicos(self):
        data1s = self.data_inicials.get() #função para obter o conteudo da caixa de entrada
        data2s = self.data_finals.get()
        data1s_formatada = datetime.strptime(data1s, "%d/%m/%Y").strftime("%Y-%m-%d")
        data2s_formatada = datetime.strptime(data2s, "%d/%m/%Y").strftime("%Y-%m-%d")
        consulta = """
            SELECT 
                CONVERT(VARCHAR(10), 
                MCTB.DATA_LANCTO, (103)) AS DATA, 
                CONVERT(INT,TBE.NUM_DOCTO) AS NUM_DOCTO, 
                CONVERT(INT,TBE.COD_CLI_FOR) AS N_CADASTRO, 
                TCG.NOME_CADASTRO + ' - ' + MCTB.DESC_CONTA AS SERVIÇO, 
                MCTB.VALOR_DEBITO AS VALOR, 
                MCTB.NOME_CCUSTO AS CENTRO_DE_CUSTO 
            FROM 
                VWMOVTOCTB MCTB 
                LEFT JOIN TBENTRADAS TBE ON (TBE.CHAVE_FATO = MCTB.CHAVE_FATO) 
                LEFT JOIN TBCADASTROGERAL TCG ON (TBE.COD_CLI_FOR = TCG.COD_CADASTRO) 
            WHERE 
                TBE.STATUS NOT LIKE 'C' 
                AND TBE.COD_TIPO_MV IN ('T139', 'F105') 
                AND MCTB.STATUS_PARTIDA = 'D' AND DATA_LANCTO >= '{}' 
                AND DATA_LANCTO < '{}' 
                AND MCTB.DESC_CONTA NOT LIKE '%A RECUPERAR%'
            """.format(data1s_formatada, data2s_formatada)

        try:
            result = pd.read_sql(consulta, connect)
            result.to_excel(r"\\servidor\Compras\RELATORIOS\RELATORIO_SERVICOS_{}.xlsx".format(data2s_formatada), index=False)
            messagebox.showinfo("Sucesso","Consulta realizada com sucesso!")
        except Exception as e: #caso aconteça algum erro 
            messagebox.showerror("Erro", f"Ocorreu um erro ao executar a consulta:\n{str(e)}")

    #fazer consulta dos materiais
    def gerar_materiais(self):
        data1m = self.data_inicialm.get()
        data2m = self.data_finalm.get()

        data1m_formatada = datetime.strptime(data1m, "%d/%m/%Y").strftime("%Y-%m-%d")
        data2m_formatada = datetime.strptime(data2m, "%d/%m/%Y").strftime("%Y-%m-%d")

        print(data1m_formatada)

        consulta_mat = """
            SELECT 
                CONVERT(VARCHAR(10),
                TBS.DATA_MOVTO,(103)) AS DATA, 
                CONVERT(INT,TBS.NUM_DOCTO) AS DOCUMENTO, 
                CONVERT(INT,TBSI.COD_PRODUTO) AS COD_PRODUTO, 
                PR.DESC_PRODUTO_EST AS DESCRIÇÃO, TBSI.QTDE_UND, 
                COALESCE(TBSI.VALOR_UNITARIO_CUSTO_RCPE, TBSI.VALOR_UNITARIO) AS VALOR_UNITARIO, 
                COALESCE(TBSI.VALOR_CUSTO_RCPE, TBSI.VALOR_TOTAL) AS VALOR_TOTAL, 
                CC.NOME_CCUSTO 
                FROM 
                TBSAIDASITEM TBSI 
                INNER JOIN TBSAIDAS TBS ON (TBS.CHAVE_FATO = TBSI.CHAVE_FATO) 
                LEFT JOIN TBCENTROCUSTO CC ON (CC.COD_CCUSTO = TBSI.COD_CCUSTO) 
                LEFT JOIN TBPRODUTO PR ON (TBSI.COD_PRODUTO = PR.COD_PRODUTO) 
                WHERE 
                TBSI.STATUS_ITEM NOT LIKE 'C' AND 
                DATA_MOVTO >= '{}' AND DATA_MOVTO < '{}' AND 
                TBS.COD_TIPO_MV IN ('1151', '1152', '1153', '1154')
            """.format(data1m_formatada, data2m_formatada)
        
        try:
            result = pd.read_sql(consulta_mat, connect)
            result.to_excel(r"\\servidor\Compras\RELATORIOS\RELATORIO_MATERIAIS_{}.xlsx".format(data2m_formatada), index=False)
            messagebox.showinfo("Sucesso","Consulta realizada com sucesso!") 
        except Exception as e: #caso aconteça algum erro 
            messagebox.showerror("Erro", f"Ocorreu um erro ao executar a consulta:\n{str(e)}")

    #tela de serviços 
    def servicos(self):
        self.limpar_tela()
        self.head()

        label = tk.Label(self, text="Relatório de Serviços", font=("Helvetica", 14, "bold"))
        label.pack(pady=10)

        label = tk.Label(self, text="Insira a data inicial: ")
        label.pack(pady=10)

        #entrada de dados de data inicial
        self.data_inicials = DateEntry(self, date_pattern="dd/mm/yyyy", datefont=('Helvetica', 10), width=12, selectbackground='gray80', locale='pt_BR')
        self.data_inicials.pack(pady=10) #ajusta a caixa de input automaticamente

        label = tk.Label(self, text="Insira a data final: ")
        label.pack(pady=10)

        #entrada de dados de data final
        self.data_finals = DateEntry(self, date_pattern="dd/mm/yyyy", datefont=('Helvetica', 10), width=12, selectbackground='gray80', locale='pt_BR')
        self.data_finals.pack(pady=10) #ajusta a caixa de input automaticamente
        
        #botão para gerar relatório
        botao_materiais = tk.Button(text="Gerar relatório", command=self.gerar_servicos)
        botao_materiais.config(**button_style)
        botao_materiais.pack(pady=10)

        back_button = tk.Button(self, text="<< Back to Menu", command=self.widgets)
        back_button.pack(pady=10)
        
    def materiais(self):
        self.limpar_tela()
        self.head()

        label = tk.Label(self, text="Relatório de Materiais", font=("Helvetica", 14, "bold"))
        label.pack(pady=10)

        label = tk.Label(self, text="Insira a data inicial: ")
        label.pack(pady=10)

        #entrada de dados de data inicial
        self.data_inicialm = DateEntry(self, date_pattern="dd/mm/yyyy", datefont=('Helvetica', 10), width=12, selectbackground='gray80', locale='pt_BR')
        self.data_inicialm.pack(pady=10) #ajusta a caixa de input automaticamente

        label = tk.Label(self, text="Insira a data final: ")
        label.pack(pady=10)

        #entrada de dados de data final
        self.data_finalm = DateEntry(self, date_pattern="dd/mm/yyyy", datefont=('Helvetica', 10), width=12, selectbackground='gray80', locale='pt_BR')
        self.data_finalm.pack(pady=10) #ajusta a caixa de input automaticamente

        #botão para gerar relatório
        botao_materiais = tk.Button(text="Gerar relatório", command=self.gerar_materiais)
        botao_materiais.config(**button_style)
        botao_materiais.pack(pady=10)

        back_button = tk.Button(self, text="<< Back to Menu", command=self.widgets)
        back_button.pack(pady=10)

    def limpar_tela(self):
        # Limpa todos os widgets da tela
        for widget in self.winfo_children():
            widget.destroy()

        # Chama update para garantir que a tela seja limpa imediatamente
        self.update()

    def fechar_programa(self):
        self.quit()

if __name__ == "__main__":
    app = Aplicacao()
    app.mainloop()
