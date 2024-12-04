import pyodbc
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def conectar_sql_server():
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=----;'
        'DATABASE=robo_email;'
        'UID=sa;'
        'PWD=------;'
    )
    return conn

def criar_arquivo_excel(mensagem):
    emails_data = [{'Mensagem': mensagem}]  # Somente a mensagem
    df = pd.DataFrame(emails_data)
    arquivo_excel = 'mensagem.xlsx'  # Nome do arquivo
    df.to_excel(arquivo_excel, index=False)
    return arquivo_excel

def enviar_email(nome, email_destino, mensagem, arquivo_excel, email_origem, senha, smtp_server):
    smtp_port = 587

    msg = MIMEMultipart()
    msg['Subject'] = f"Mensagem para {nome}"
    msg['From'] = email_origem
    msg['To'] = email_destino

    msg.attach(MIMEText(mensagem, 'plain'))

    with open(arquivo_excel, 'rb') as anexo:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(anexo.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={arquivo_excel}')
        msg.attach(part)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_origem, senha)
            server.sendmail(email_origem, email_destino, msg.as_string())
        return 'Enviado com sucesso'
    except Exception as e:
        
        return f'Erro ao enviar: {str(e)}'


def processar_envio():
    conn = conectar_sql_server()
    cursor = conn.cursor()

    cursor.execute('SELECT id, nome, email, mensagem FROM emails')
    emails = cursor.fetchall()

    
    contas_email = [
        ('email1', 'senha', 'servidor'),
        ('email2', 'senha', 'servidor')
    ]

    for index, email in enumerate(emails):
        id_email = email.id
        nome = email.nome
        email_destino = email.email
        mensagem = email.mensagem

        arquivo_excel = criar_arquivo_excel(mensagem)

        email_origem, senha, smtp_server = contas_email[index // 10 % len(contas_email)]

        status_envio = enviar_email(nome, email_destino, mensagem, arquivo_excel, email_origem, senha, smtp_server)

        cursor.execute(''' 
            INSERT INTO log_envio (email_id, status, data_envio)
            VALUES (?, ?, GETDATE())
        ''', (id_email, status_envio))

        conn.commit()

    conn.close()

if __name__ == "__main__":
    processar_envio()
