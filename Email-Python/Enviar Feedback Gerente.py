import tkinter as tk
from tkinter import filedialog
import threading
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

caminho_arquivo = ""

def selecionar_arquivo_e_enviar():
    global caminho_arquivo
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_arquivo:
        email_thread = threading.Thread(target=enviar_emails, args=(caminho_arquivo,))
        email_thread.start()
    else:
        print("Nenhum arquivo selecionado.")

def criar_linhas_tabela_coordenador(df):
    linhas = []
    for _, row in df.iterrows():
        linha = (
            f'<tr><td style="border: 2px solid black; padding: 6px;"><font color="red"><strong>{row["Coordenador(a)"]}</strong></font></td>'
            f'<td style="border: 2px solid black; padding: 6px;"><font size="2px"><strong>{row["Quantidade de monitorias"]}</strong></font></td></tr>'
        )
        linhas.append(linha)
    return '\n'.join(linhas)

def criar_linhas_tabela_supervisor(df):
    linhas = []
    for _, row in df.iterrows():
        linha = (
            f'<tr><td style="border: 2px solid black; padding: 6px;"><font color="red"><strong>{row["Superior"]}</strong></font></td>'
            f'<td style="border: 2px solid black; padding: 6px;"><font size="2px"><strong>{row["Quantidade de monitorias"]}</strong></font></td></tr>'
        )
        linhas.append(linha)
    return '\n'.join(linhas)

def enviar_emails(caminho_arquivo):
    if caminho_arquivo:  
        EnviarEmailGerente(caminho_arquivo)
        janela.destroy()

def EnviarEmailGerente(caminho_arquivo):
    if caminho_arquivo:
        df = pd.read_excel(caminho_arquivo)
        gerente_email = 'jose.raimundo@aec.com.br, vinicius.bezerra@aec.com.br, victor.evani@aec.com.br, lorranny.santos@aec.com.br'
        msg = MIMEMultipart()
        msg['Subject'] = "Relatório de Monitorias pendentes de aplicação de feedback"

        coluna_coordenador = 'Coordenador' 
        nomes_coordenadores = ['VINICIUS ALVES BEZERRA DA SILVA', 'LORRANNY ALVES DOS SANTOS', 'VICTOR EVANI SOUZA SOCHA']
        df_filtrado_coordenadores = df[df[coluna_coordenador].isin(nomes_coordenadores)]
        contagem_monitorias_pendente_feedback_coordenador = df_filtrado_coordenadores[coluna_coordenador].value_counts()

        df_contagem_coordenadores = pd.DataFrame({
            'Coordenador(a)': df_filtrado_coordenadores[coluna_coordenador].unique(),
            'Quantidade de monitorias': contagem_monitorias_pendente_feedback_coordenador.values
        })

        coluna_supervisor = 'Superior' 
        nomes_supervisores = ['AISLAYNE JORDANA MORAES DE OLIVEIRA', 'ANDREINA STEPHANE ALVES FARIAS', 'ANTHONNY MICHAEL BALBINO DA SILVA SANTOS', 'CARLOS EDUARDO DOS SANTOS OLIVEIRA', 'DANILO GUSTAVO DOS SANTOS MORAIS', 'DIEGO SANTOS DE SOUZA', 'DOUGLAS JOSUE NASCIMENTO DA SILVA', 'ICARO ALEXANDRE TORRES', 'JACIARA MILANE DA SILVA LIMA', 'JESICA DA SILVA OLIVEIRA', 'JOSE BRUNO SANTOS SILVA', 'JOSE CASSIO TAVARES DA SILVA', 'JOSE IALDO MAGALHAES BARBOSA', 'KAUA VICTOR NASCIMENTO PEREIRA', 'LAIS MORGANA CASSIANO DOS SANTOS', 'LARISSA FERREIRA DA SILVA', 'LARYSSA MONIQUE SANTOS LOPES', 'LAURA SABRINA TARGINO DOS SANTOS', 'LAURITA DOS SANTOS RIBEIRO', 'LUIZ GUSTAVO PAZ DO NASCIMENTO', 'MANOEL MESSIAS PEREIRA DE FREITAS', 'MARCIA MARIA BARBOSA', 'MARIA CRISTINA FLORENTINO DOS SANTOS', 'MARIA RICKELLY VITAL BARBOSA', 'MARIANA PEREIRA VERAS', 'MATEUS FELIPE DOS SANTOS', 'MILENA CARLA COSTA DOS SANTOS', 'PAULINA MAYARA FERREIRA DA SILVA', 'RAYLAN PLADION DOS SANTOS', 'SANSHAINE ALVES DE FRANCA', 'SARAH DOS SANTOS SALES', 'SIDREIA DE SOUZA SANTOS LIMA', 'THAIS DOS SANTOS COSTA', 'VANESSA CAVALCANTE DE SOUZA', 'VITORIA ELIZABETHE ARAUJO TORRES BARROS', 'ZAQUEU MONTEIRO DA SILVA']
        df_filtrado_supervisores = df[df[coluna_supervisor].isin(nomes_supervisores)]
        contagem_monitorias_pendente_feedback_supervisor = df_filtrado_supervisores[coluna_supervisor].value_counts()

        df_contagem_supervisores = df_filtrado_supervisores.groupby(coluna_supervisor).size().reset_index(name='Quantidade de monitorias')

        corpo_tabela_coordenador = criar_linhas_tabela_coordenador(df_contagem_coordenadores)
        corpo_tabela_supervisor = criar_linhas_tabela_supervisor(df_contagem_supervisores)

        messageemail = f"""\
        <html>
        <body>
        <p><h4>Olá prezados, tudo bem? Espero que sim!</h4></p>
        <p><h4>Segue relatório de quantidade de monitorias pendentes de aplicação de feedback por Coordenador e Supervisor. Qualquer dúvida, estou a disposição.</h4></p>
        <br>
        <br>
        <p><h3><font color="red">Monitorias pendentes de aplicação de feedback por coordenador(a):</font></h3></p>
        <table style="border-collapse: collapse; width: 60%; text-align: center;">
        <tr>
        <th style="border: 3px solid black; padding: 6px; background-color: #000080;"><font color="white" size="3px">Coordenador(a)</font></th>
        <th style="border: 3px solid black; padding: 6px; background-color: #000080;"><font color="white" size="3px">Quantidade de monitorias</font></th>
        </tr>
        {corpo_tabela_coordenador}
        </table>
        <br>
        <br>
        <p><h3><font color="red">Monitorias pendentes de aplicação de feedback por supervisor(a):</font></h3></p>
        <table style="border-collapse: collapse; width: 60%; text-align: center;">
        <tr>
        <th style="border: 3px solid black; padding: 6px; background-color: #000080;"><font color="white" size="3px">Supervisor(a)</font></th>
        <th style="border: 3px solid black; padding: 6px; background-color: #000080;"><font color="white" size="3px">Quantidade de monitorias</font></th>
        </tr>
        {corpo_tabela_supervisor}
        </table>
        <p>Atenciosamente</p>
        <p>Gerencia Rai, Coordenação Alpha e Loyalty</p>
        <p>Unidade Arapiraca</p>
        <img src="https://i.imgur.com/74arG9m.jpg" />
        </body>
        </html>
        """
        msg['From'] = 'alphaeloyalty@gmail.com'
        msg['To'] = gerente_email
        msg.attach(MIMEText(messageemail, 'html'))
        context = smtplib.SMTP('smtp.gmail.com', 587)
        context.starttls()
        context.login('alphaeloyalty@gmail.com', 'vxxggabmsmzfapff')
        try:
            context.sendmail(msg['From'], [msg['To']], msg.as_string())
            print(f"Email enviado com sucesso para o gerente iFood.")
        except Exception as e:
            print(f"Erro ao enviar email para gerente iFood: {str(e)}")

janela = tk.Tk()
janela.geometry("300x70")
janela.title("Enviar Email")

botao_selecionar_enviar = tk.Button(janela, text="Selecionar export do 2CLIX", command=selecionar_arquivo_e_enviar, bg="black", fg="white", font="-weight bold -size 15")
botao_selecionar_enviar.pack(padx=15, pady=20)

janela.mainloop()
