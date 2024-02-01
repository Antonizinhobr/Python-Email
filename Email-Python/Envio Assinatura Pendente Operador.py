import tkinter as tk
from tkinter import filedialog
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def selecionar_arquivo_e_enviar():
    global caminho_arquivo
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_arquivo:
        enviar_emails(caminho_arquivo)
    else:
        print("Nenhum arquivo selecionado.")

def enviar_emails(caminho_arquivo):
    if caminho_arquivo:
        global contagem_emails_supervisor
        contagem_emails_supervisor = 0
        EnviarEmailOperador(caminho_arquivo)

def EnviarEmailOperador(caminho_arquivo):
    if caminho_arquivo:
        global contagem_emails_supervisor

        df = pd.read_excel(caminho_arquivo)
        
        analista = df['Avaliado'].unique()
        
        for analista_email in analista:
            analista_df = df[df['Avaliado'] == analista_email]
            enviar_email_individual(analista_df)

def enviar_email_individual(df):
    global contagem_emails_supervisor
    msg = MIMEMultipart()
    msg['Subject'] = "Pendências de Assinatura! iFood CX - Alpha E Loyalty"
    messageemail = f"""\
    <html>
    <head>
        <style>
            .custom-row {{
                background-color: #ADFF2F;
            }}
        
            .normal-row {{
                background-color: #FA8072;
            }}
        </style>
    </head>
    <body>
    <p><h5>Olá <font color="red"><strong>{df.iloc[0]['Avaliado']},</font></strong> tudo bem? Espero que sim!</h5></p>
    <p><h5>Você tem feedbacks pendentes de assinatura na ferramenta 2CLIX. Peço que assine as monitorias pendentes após receber o feedback de seu gestor(a). Qualquer dúvida, estou a disposição.</h5></p>
    <br>
    <table style="margin: left; text-align: center; border-collapse: collapse;">
        <tr>
            <th style='border: 2px solid black; padding: 3px;'>Código da avaliação</th>
            <th style='border: 2px solid black; padding: 3px;'>Data da avaliação</th>
            <th style='border: 2px solid black; padding: 3px;'>Avaliado</th>
            <th style='border: 2px solid black; padding: 3px;'>Nota</th>
            <th style='border: 2px solid black; padding: 3px;'>Superior</th>
            <th style='border: 2px solid black; padding: 3px;'>Coordenador</th>
        </tr>
        {df.apply(lambda row: f"<tr class='{'custom-row' if row['Nota'] == 100 else 'normal-row'}'><td style='border: 2px solid black; padding: 3px;'>{row['Código da avaliação']}</td><td style='border: 2px solid black; padding: 3px;'>{row['Data da avaliação']}</td><td style='border: 2px solid black; padding: 3px;'>{row['Avaliado']}</td><td style='border: 2px solid black; padding: 3px;'>{row['Nota']}</td><td style='border: 2px solid black; padding: 3px;'>{row['Superior']}</td><td style='border: 2px solid black; padding: 3px;'>{row['Coordenador']}</td></tr>", axis=1).str.cat()}
    </table>
    <br>
    <p>Atenciosamente</p>
    <p>Gerencia Rai, Coordenação Alpha e Loyalty</p>
    <p>Unidade Arapiraca</p>
    <img src="https://i.imgur.com/74arG9m.jpg" />
    </body>
    </html>
    """
    
    msg['From'] = 'alphaeloyalty@gmail.com'
    msg['To'] = df.iloc[0]['Email Analista']
    msg.attach(MIMEText(messageemail, 'html'))

    context = smtplib.SMTP('smtp.gmail.com', 587)
    context.starttls()
    context.login('alphaeloyalty@gmail.com', 'vxxggabmsmzfapff')
    try:
        context.sendmail(msg['From'], msg['To'], msg.as_string())
        contagem_emails_supervisor += 1
        print(f"Email enviado para {msg['To']} com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar email para {msg['To']}: {str(e)}")

janela = tk.Tk()
janela.geometry("300x70")
janela.title("Enviar Email")

botao_selecionar_enviar = tk.Button(janela, text="Selecionar export do 2CLIX", command=selecionar_arquivo_e_enviar, bg="black", fg="white", font="-weight bold -size 15")
botao_selecionar_enviar.pack(padx=15, pady=20)

janela.mainloop()
