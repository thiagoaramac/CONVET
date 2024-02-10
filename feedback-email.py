import os.path
import base64
import os
import google.auth
from email.message import EmailMessage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# E-mail destinatário --------------------------------------------------------------------------------------------------
destinatario = "thiagoaramac@gmail.com"

# Título do E-mail------------------------------------------------------------------------------------------------------
titulo = "Seu Feedback do Simulado CONVET nº 29"

# Texto do Feedback ----------------------------------------------------------------------------------------------------
texto_feedback = """
<p>Oi, André! Tudo bem? &#128516;</p><p><strong>Seu feedback na prova como um todo:</strong><br/>Sua Nota Final (objetiva + discursiva): <strong>29.4/30.00</strong><br/>Seu ranking da prova discursiva: <strong>1ª</strong> nota mais alta<br/>Seu ranking da prova Objetiva Básica: <strong>1ª</strong> nota mais alta<br/>Seu ranking da prova Objetiva Específica: <strong>1ª</strong> nota mais alta<br/>Seu ranking geral: <strong>1ª</strong> nota mais alta</p><p><strong>Seu feedback em cada disciplina:</strong><br/>Língua Portuguesa: <strong>2ª</strong> nota mais alta<br/>Noções de direito administrativo e constitucional: <strong>1ª</strong> nota mais alta<br/>Noções de raciocínio lógico e matemático: <strong>1ª</strong> nota mais alta<br/>Noções de Informática: <strong>1ª</strong> nota mais alta<br/>Disciplinas do Eixo Transversal: <strong>1ª</strong> nota mais alta<br/><p><strong>Esse simulado teve <u>108</u> alunos</strong><br/>Média geral(conhecimentos básicos): <strong>1</strong><br/>Média geral(conhecimentos específicos): <strong>1</strong><br/>Média geral(discursiva): <strong>1</strong><br/>Continue firme e bons estudos!!</p><br/><p><font color="gray"><b><i><h1 style="font-size:8pt; ">Lembrando que esse é um email automático do CONVET!</h1></i></b><i><h1 style="font-size:8pt; ">Se você tentar respondê-lo, ninguém vai ver &#128542;</h1></i></font></p><img src="https://www.concursosconvet.com.br/assets/img/logotipo/logotipo.png"  width="100" height="100">
"""
# ----------------------------------------------------------------------------------------------------------------------


def gmail_send_message(destinatario, titulo, texto_feedback):
    SCOPES = ["https://www.googleapis.com/auth/gmail.readonly",
              "https://www.googleapis.com/auth/gmail.send",
              "https://www.googleapis.com/auth/gmail.labels",
              "https://www.googleapis.com/auth/gmail.compose",
              "https://mail.google.com/",
              ]

    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            file = 'client_secret_904512780263-9hjfifvl0gpoa1edgviut8lldjqt72kg.apps.googleusercontent.com.json'
            flow = InstalledAppFlow.from_client_secrets_file(file, SCOPES)
            creds = flow.run_local_server(port = 0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("gmail", "v1", credentials = creds)
        message = EmailMessage()

        message.set_content(texto_feedback, subtype="html")

        message["To"] = destinatario
        message["From"] = "no-reply@convet.com.br"
        message["Subject"] = titulo

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {"raw": encoded_message}
        # pylint: disable=E1101
        send_message = (
            service.users()
            .messages()
            .send(userId = "me", body = create_message)
            .execute()
        )
        print("E-mail enviado com sucesso para: " + destinatario)
        print(f'Id da mensagem: {send_message["id"]}')
    except HttpError as error:
        print(f"An error occurred: {error}")
        send_message = None
    return send_message


# ----------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    gmail_send_message(destinatario, titulo, texto_feedback)
