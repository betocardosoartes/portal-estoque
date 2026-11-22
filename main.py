import imaplib
import email
import io
import os
import socket
import pandas as pd
from fastapi import FastAPI

app = FastAPI()

# Define um limite máximo de 10 segundos para não travar a página
socket.setdefaulttimeout(10)

EMAIL_USER = os.getenv("EMAIL_USER", "seu_email@gmail.com")
EMAIL_PASS = os.getenv("EMAIL_PASS", "sua_senha_de_app")
IMAP_SERVER = os.getenv("IMAP_SERVER", "://gmail.com")
ASSUNTO_BUSCA = os.getenv("ASSUNTO_BUSCA", "Relatorio")

# Rota 1: Teste rápido para saber se o site está online (abre na hora)
@app.get("/")
def inicio():
    return {"status": "Online!", "mensagem": "O servidor Render esta funcionando. Acesse /testar-email para ler a planilha."}

# Rota 2: Teste de conexão com o e-mail
@app.get("/testar-email")
def testar_email():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        status, mensagens = mail.search(None, f'(SUBJECT "{ASSUNTO_BUSCA}")')
        ids_emails = mensagens[0].split()

        if not ids_emails:
            mail.logout()
            return {"status": "Aviso", "mensagem": f"Conectou no e-mail, mas nao achou assunto '{ASSUNTO_BUSCA}'."}

        ultimo_id = ids_emails[-1]
        _, dados_email = mail.fetch(ultimo_id, "(RFC822)")
        msg = email.message_from_bytes(dados_email[0][1])

        for parte in msg.walk():
            if parte.get_content_maintype() == "multipart" or parte.get("Content-Disposition") is None:
                continue

            nome_arquivo = parte.get_filename()
            if nome_arquivo and (nome_arquivo.endswith(".xlsx") or nome_arquivo.endswith(".csv")):
                conteudo = parte.get_payload(decode=True)
                df = pd.read_excel(io.BytesIO(conteudo)) if nome_arquivo.endswith(".xlsx") else pd.read_csv(io.BytesIO(conteudo))
                mail.logout()
                return {
                    "status": "SUCESSO!",
                    "arquivo": nome_arquivo,
                    "colunas": list(df.columns),
                    "primeiras_linhas": df.head(3).fillna("").to_dict(orient="records")
                }

        mail.logout()
        return {"status": "Aviso", "mensagem": "E-mail encontrado, mas sem planilha .xlsx ou .csv anexada."}
    except Exception as e:
        return {"status": "Erro na conexao", "detalhe": str(e)}
