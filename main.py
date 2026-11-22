import imaplib
import email
import io
import pandas as pd
from fastapi import FastAPI

app = FastAPI()

# --- PREENCHA SEUS DADOS AQUI (APENAS PARA TESTE) ---
EMAIL_USER = "estoque@amazonnorte.com"
EMAIL_PASS = "B3to@464"
IMAP_SERVER = "mail.amazonnorte.com"
ASSUNTO_BUSCA = "WEBSGEP - ESTOQUE DIÁRIO AMAZON"
# ----------------------------------------------------

@app.get("/")
def testar():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        status, mensagens = mail.search(None, f'(SUBJECT "{ASSUNTO_BUSCA}")')
        ids_emails = mensagens[0].split()

        if not ids_emails:
            mail.logout()
            return {"status": "Erro", "mensagem": "Nenhum e-mail encontrado com esse assunto."}

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
        return {"status": "Erro", "mensagem": "E-mail achado, mas sem planilha anexada."}
    except Exception as e:
        return {"status": "Falha", "erro": str(e)}
