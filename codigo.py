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

Meu nome é Chico Alves e atuo como consultor técnico especializado no Sistema Único de Assistência Social (SUAS), com experiência em Vigilância Socioassistencial, análise de dados, trabalho social com famílias, gestão territorial e desenvolvimento de sistemas de informação aplicados à política de assistência social.

Encaminho este contato para apresentar uma proposta de consultoria técnica personalizada, voltada ao aprimoramento da Vigilância Socioassistencial e à qualificação do trabalho social, integrando diagnóstico territorial, construção de indicadores, fortalecimento da gestão e desenvolvimento de um sistema informatizado de atendimento e acompanhamento da Rede SUAS municipal.

A consultoria é desenvolvida em modalidade híbrida, com acompanhamento técnico contínuo, produção de relatórios periódicos e adaptação às especificidades do território e da rede local.

Anexo a este e-mail segue a proposta técnica detalhada, contendo objetivos, metodologia, etapas e cronograma.

Fico à disposição para esclarecimentos ou para agendarmos uma breve conversa.

Atenciosamente,

Chico Alves
Sociólogo | Analista de Dados
Consultor em Políticas Públicas e SUAS
Belo Horizonte – MG
f.neto.alves@hotmail.com
(31) 97145-0972
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
