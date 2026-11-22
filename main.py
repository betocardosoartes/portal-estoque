import imaplib
import email
import io
import os
import socket
import pandas as pd
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse

app = FastAPI()
socket.setdefaulttimeout(15)

# --- CONFIGURAÇÕES DO E-MAIL (Vêm do Render) ---
EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASS = os.getenv("EMAIL_PASS", "")
IMAP_SERVER = os.getenv("IMAP_SERVER", "://gmail.com")
ASSUNTO_BUSCA = os.getenv("ASSUNTO_BUSCA", "Relatorio")

# --- AJUSTE O NOME DA COLUNA DA SIGLA NA SUA PLANILHA ---
COLUNA_SIGLA = "SIGLA"

# --- CADASTRO DE SENHAS POR SIGLA (Coloque em maiúsculo) ---
CLIENTES_SENHAS = {
    "AMZ": "1234",
    "BEL": "senha2026",
    "PAR": "abc123",
}

def obter_planilha_do_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, mensagens = mail.search(None, f'(SUBJECT "{ASSUNTO_BUSCA}")')
    ids_emails = mensagens[0].split()

    if not ids_emails:
        mail.logout()
        raise Exception(f"Nenhum e-mail encontrado com o assunto '{ASSUNTO_BUSCA}'.")

    ultimo_id = ids_emails[-1]
    _, dados_email = mail.fetch(ultimo_id, "(RFC822)")
    msg = email.message_from_bytes(dados_email[0][1])

    for parte in msg.walk():
        if parte.get_content_maintype() == "multipart" or parte.get("Content-Disposition") is None:
            continue

        nome_arquivo = parte.get_filename()
        if nome_arquivo and (nome_arquivo.endswith(".xlsx") or nome_arquivo.endswith(".csv")):
            conteudo = parte.get_payload(decode=True)
            mail.logout()
            if nome_arquivo.endswith(".xlsx"):
                return pd.read_excel(io.BytesIO(conteudo))
            else:
                return pd.read_csv(io.BytesIO(conteudo))

    mail.logout()
    raise Exception("Nenhuma planilha (.xlsx ou .csv) encontrada no anexo do e-mail.")

# --- PÁGINA INICIAL COM FORMULÁRIO VISUAL ---
@app.get("/", response_class=HTMLResponse)
def pagina_login(erro: str = ""):
    alerta_html = f'<div class="erro">{erro}</div>' if erro else ""
    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Portal de Estoque</title>
        <style>
            body {{ font-family: Arial, sans-serif; background: #f0f2f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; padding: 15px; }}
            .card {{ background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); width: 100%; max-width: 380px; }}
            h2 {{ margin-top: 0; color: #1a202c; text-align: center; }}
            p {{ color: #4a5568; font-size: 14px; text-align: center; margin-bottom: 20px; }}
            label {{ display: block; font-weight: bold; margin-bottom: 5px; color: #2d3748; font-size: 14px; }}
            input {{ width: 100%; padding: 12px; margin-bottom: 16px; border: 1px solid #cbd5e0; border-radius: 8px; box-sizing: border-box; font-size: 16px; }}
            button {{ width: 100%; padding: 14px; background: #2563eb; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; }}
            button:hover {{ background: #1d4ed8; }}
            .erro {{ background: #fee2e2; color: #b91c1c; padding: 10px; border-radius: 8px; margin-bottom: 15px; font-size: 14px; text-align: center; }}
        </style>
    </head>
    <body>
        <div class="card">
            <h2>Portal de Estoque</h2>
            <p>Informe sua sigla e senha para baixar sua planilha atualizada.</p>
            {alerta_html}
            <form action="/baixar" method="post">
                <label for="sigla">Sigla do Cliente</label>
                <input type="text" id="sigla" name="sigla" placeholder="Ex: AMZ" required style="text-transform: uppercase;">
                
                <label for="senha">Senha de Acesso</label>
                <input type="password" id="senha" name="senha" placeholder="Digite sua senha" required>
                
                <button type="submit">Baixar Planilha (.xlsx)</button>
            </form>
        </div>
    </body>
    </html>
    """

# --- ROTA QUE VALIDA, FILTRA E GERA O DOWNLOAD ---
@app.post("/baixar")
def baixar_planilha(sigla: str = Form(...), senha: str = Form(...)):
    sigla_limpa = sigla.strip().upper()

    # 1. Valida Sigla e Senha
    if sigla_limpa not in CLIENTES_SENHAS or CLIENTES_SENHAS[sigla_limpa] != senha:
        return HTMLResponse(content=pagina_login(erro="Sigla ou senha inválidos."), status_code=401)

    try:
        # 2. Busca a planilha no e-mail
        df = obter_planilha_do_email()

        # 3. Verifica se a coluna existe
        if COLUNA_SIGLA not in df.columns:
            return HTMLResponse(
                content=pagina_login(erro=f"Erro interno: Coluna '{COLUNA_SIGLA}' não encontrada na planilha."),
                status_code=500
            )

        # 4. Filtra os dados apenas daquela sigla
        df_filtrado = df[df[COLUNA_SIGLA].astype(str).str.strip().str.upper() == sigla_limpa]

        if df_filtrado.empty:
            return HTMLResponse(
                content=pagina_login(erro=f"Nenhum registro encontrado para a sigla {sigla_limpa}."),
                status_code=404
            )

        # 5. Gera o arquivo Excel na memória para download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name=sigla_limpa)
        buffer.seek(0)

        nome_download = f"estoque_{sigla_limpa}.xlsx"
        headers = {"Content-Disposition": f'attachment; filename="{nome_download}"'}

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )

    except Exception as e:
        return HTMLResponse(content=pagina_login(erro=f"Falha ao processar: {str(e)}"), status_code=500)
