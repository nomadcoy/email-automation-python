# Email Automation with Python

## Overview
This project automates the process of sending emails using Python based on a structured dataset.
It was developed as a portfolio project to demonstrate automation, data handling, and practical scripting for real-world tasks.

The focus is not only on automation itself, but on designing a reproducible and data-driven workflow.

---

## Technologies
- Python
- pandas
- pyautogui
- openpyxl
- PowerShell
- Virtual Environment (venv)

---

## How the Project Works
1. A dataset (Excel or CSV) contains the email recipients and relevant fields used for personalization.
2. The script reads and processes the dataset using pandas.
3. For each record, the script automates the email composition and sending process based on a predefined template.
4. Emails are sent individually, allowing controlled and personalized communication at scale.

---

## Technical Considerations
- This project uses GUI automation with `pyautogui`.
- While the script is running:
  - The computer should not be used for other tasks.
  - Screen resolution and interface layout must remain unchanged.
- This approach was chosen intentionally to explore automation under constrained environments.

---

## Limitations
- GUI-based automation is sensitive to interface changes.
- Accented characters may require specific handling depending on the environment.
- The script is not intended for large-scale or background execution.

---

## Possible Improvements
- Replace GUI automation with SMTP or email APIs.
- Add structured logging and exception handling.
- Implement validation for email addresses.
- Introduce rate control and execution reports.

---

## Project Structure

emails-automation/
├── data/
│ └── contacts.xlsx
├── codigo.py
├── README.md
└── env/

## Code

import pyautogui
import pyperclip
import time
import pandas as pd
import sys

# =========================
# Configurações
# =========================

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

ARQUIVO_EXCEL = "Mailing prefeitura, órgão gestor e conselho - 2025 2.xlsx"
ASSUNTO = "Consultoria Técnica em SUAS"

CORPO_EMAIL = """Prezados(as),

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

# =========================
# Funções auxiliares
# =========================

def abrir_chrome():
    pyautogui.press("win")
    time.sleep(3)
    pyautogui.write("chrome")
    pyautogui.press("enter")
    time.sleep(8)

def acessar_outlook():
    pyautogui.click(x=223, y=67)
    pyautogui.write("https://outlook.live.com/mail/0/")
    pyautogui.press("enter")
    time.sleep(12)

def novo_email():
    pyautogui.click(x=129, y=215)
    time.sleep(10)

def preencher_email(destinatario):
    pyautogui.click(x=771, y=325)
    pyautogui.write(destinatario)
    time.sleep(2)

    pyautogui.click(x=681, y=377)
    pyperclip.copy(ASSUNTO)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    pyautogui.click(x=697, y=431)
    pyperclip.copy(CORPO_EMAIL)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

def anexar_arquivo():
    anexar = pyautogui.locateCenterOnScreen("anexar.png")
    if anexar is None:
        pyautogui.alert("Botão 'Anexar' não encontrado.")
        sys.exit()

    pyautogui.click(anexar)
    time.sleep(3)

    navegar = pyautogui.locateCenterOnScreen("navegar_pc.png")
    if navegar is None:
        pyautogui.alert("Opção 'Navegar neste computador' não encontrada.")
        sys.exit()

    pyautogui.click(navegar)
    time.sleep(3)

    pyautogui.doubleClick(x=595, y=268)  # arquivo
    time.sleep(5)

def enviar_email():
    pyautogui.doubleClick(x=712, y=273)
    time.sleep(5)

def voltar_inicio():
    pyautogui.click(x=230, y=183)
    time.sleep(2)

# =========================
# Execução principal
# =========================

def main():
    time.sleep(2)
    abrir_chrome()
    acessar_outlook()

    tabela = pd.read_excel(ARQUIVO_EXCEL)

    for _, linha in tabela.iterrows():
        endereco = str(linha["endereco"]).lower()

        try:
            novo_email()
            preencher_email(endereco)
            anexar_arquivo()
            enviar_email()
            voltar_inicio()

        except Exception as erro:
            print(f"Erro ao enviar para {endereco}: {erro}")

if __name__ == "__main__":
    main()
