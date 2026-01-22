# Bot de Envio de Emails em Massa

Este projeto é um script Python para enviar emails em massa para uma lista de contatos, com anexo, usando a conta Gmail via SMTP com senha de app.  

É ideal para quem quer automatizar o envio de propostas, newsletters ou comunicações profissionais personalizadas.

---

## Funcionalidades

- Envia emails personalizados para cada destinatário listado em uma planilha Excel (.xlsx).  
- Anexa um arquivo PDF (ex: proposta técnica).  
- Envia em lotes para evitar bloqueios do Gmail (limite recomendado: 200 emails/dia).  
- Intervalos configuráveis entre envios e entre lotes.  
- Registra erros de envio no console para acompanhamento.

---

## Requisitos

- Python 3.7+  
- Bibliotecas Python:
  - pandas  
  - openpyxl  
- Conta Gmail com [Verificação em duas etapas](https://myaccount.google.com/security) ativada  
- Senha de app gerada para uso no script

