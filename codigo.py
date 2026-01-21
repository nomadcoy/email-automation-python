import smtplib
import pandas as pd
from email.message import EmailMessage
import mimetypes
import os

# Configurações SMTP do Gmail
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587

EMAIL_ADDRESS = 'seu_email@gmail.com'          # Seu email Gmail
EMAIL_PASSWORD = 'sua_senha_de_app'             # Sua senha de app do Gmail

# Caminho para o arquivo PDF que vai ser anexado
ANEXO_CAMINHO = r'C:\Users\Admin\Desktop\...'

# Ler a planilha Excel com a lista de emails
tabela = pd.read_excel("sua_base_de_dados.xlsx")

def enviar_email(destinatario):
    msg = EmailMessage()
    msg['Subject'] = "Consultoria Técnica no SUAS"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = destinatario

    corpo_email = """Prezados(as),

sua mensagem no corpo do e-mail.

Atenciosamente,
"""
    msg.set_content(corpo_email)

    # Adicionar anexo se existir
    if os.path.isfile(ANEXO_CAMINHO):
        mime_type, _ = mimetypes.guess_type(ANEXO_CAMINHO)
        mime_type = mime_type or 'application/octet-stream'
        main_type, sub_type = mime_type.split('/', 1)
        with open(ANEXO_CAMINHO, 'rb') as f:
            msg.add_attachment(f.read(),
                               maintype=main_type,
                               subtype=sub_type,
                               filename=os.path.basename(ANEXO_CAMINHO))
    else:
        print(f"Aviso: arquivo de anexo não encontrado em {ANEXO_CAMINHO}")

    # Enviar o email via SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
            print(f"Email enviado para {destinatario}")
    except Exception as e:
        print(f"Erro ao enviar para {destinatario}: {e}")

import time  # certifique-se de importar no topo do arquivo

tamanho_lote = 100      # emails por lote
intervalo_lote = 3600   # pausa de 1 hora entre lotes (3600 segundos)

total_emails = len(tabela)

for i in range(0, total_emails, tamanho_lote):
    lote = tabela.iloc[i:i+tamanho_lote]  # pega o pedaço da tabela

    print(f"Enviando lote de emails de {i+1} a {i+len(lote)}...")

    for idx, row in lote.iterrows():
        email_destino = str(row['endereco']).strip().lower()
        enviar_email(email_destino)
        time.sleep(1)  # pausa de 1 segundo entre emails

    if i + tamanho_lote < total_emails:
        print(f"Pausa de {intervalo_lote//60} minutos antes do próximo lote...")
        time.sleep(intervalo_lote)  # pausa antes do próximo lote
