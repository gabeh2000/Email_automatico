import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
import os
import pandas as pd
from io import BytesIO

'''
Bom dia, espero que esteja bem.
 
Escrevo para informar sobre a possibilidade de incremento de faturamento neste mês por meio da renovação de licenças. Abaixo compartilho imagem com as licenças e os respectivos equipamentos que as utilizam.
'''

class Emailer:
    def __init__(self, login, senha):
        self.login = login
        self.senha = senha

    def enviar_email_outlook(self,destinatario, assunto, corpo_html, planilha = None,copia=None, caminho_anexo=None):
        # Configurações do e-mail empresarial
        #Esses dois tem que ser modular
        sender_email = self.login
        password = self.senha

        smtp_server = "smtp.office365.com"
        port = 587

        # Criar a mensagem
        message = MIMEMultipart()
        message["From"] = sender_email
        if isinstance(destinatario,list):
            message["To"] = ", ".join(destinatario)
        else:
            message["To"] = destinatario
        if copia:
            if isinstance(copia,list):
                message["Cc"] = ", ".join(copia)
            else:
                message["Cc"] = copia
        message["Subject"] = assunto

        # Corpo do email (HTML)
        message.attach(MIMEText(corpo_html, "html"))

        # Anexo (opcional)
        if caminho_anexo:
            filename = os.path.basename(caminho_anexo)
            with open(caminho_anexo, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {filename}")
            message.attach(part)

        if planilha is not None:
            buffer = BytesIO()
            df_xlsx  = planilha.to_excel(buffer,index=False)
            buffer.seek(0)
            df_anexo = MIMEApplication(buffer.read(),_subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            df_anexo.add_header('content-disposition','attachment',filename="licencas_expirando.xlsx")
            message.attach(df_anexo)

        # Conectar e Enviar
        try:
            # Cria contexto SSL
            context = ssl.create_default_context()
            server = smtplib.SMTP(smtp_server, port)
            server.starttls(context=context) # Inicia TLS
            server.login(sender_email, password)
            server.send_message(message)
            return f"Email enviado com sucesso para {destinatario}!"
        except Exception as e:
            return f"Erro ao enviar: {e}"
        finally:
            server.quit()