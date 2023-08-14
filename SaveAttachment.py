#!/usr/bin/env python
# coding: utf-8

# In[9]:


import os
import win32com.client
import datetime
from pathlib import Path


# In[10]:


# Configurações
YOUR_EMAIL = "seuemail_@outlook.com"
SEARCH_FOLDER = "Caixa de Entrada"
DOWNLOAD_FOLDER = Path("C:/Users/usuario/Downloads/Solicitacao-Documental")

# Conectar-se ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.Folders.Item(YOUR_EMAIL).Folders[SEARCH_FOLDER]
messages = inbox.Items

# Dicionário para rastrear anexos baixados por assunto e data
downloaded_attachments = {}

# Filtrar e processar e-mails
for message in messages:
    if "Solicitação de Documento" in message.Subject:
        # Obter informações do e-mail
        email_subject = message.Subject
        email_sender = message.SenderName
        email_date = message.ReceivedTime.date()

        # Verificar se o anexo já foi baixado para o mesmo assunto e data
        if (email_subject, email_date) in downloaded_attachments.keys():
            print(f"Anexos para '{email_subject}' do dia {email_date} já foram baixados.")
            continue

        # Salvar anexos em PDF na pasta
        for attachment in message.Attachments:
            if attachment.FileName.lower().endswith(".pdf"):
                new_filename = f"{email_subject}_{email_sender}_{attachment.FileName}"
                attachment_path = os.path.join(DOWNLOAD_FOLDER, new_filename)
                attachment.SaveAsFile(attachment_path)
                print(f"Anexo '{attachment.FileName}' salvo como '{new_filename}' em {DOWNLOAD_FOLDER}")

        # Registrar o download do anexo para evitar duplicações
        downloaded_attachments[(email_subject, email_date)] = True

print("Processo concluído.")


# In[ ]:





# In[ ]:




