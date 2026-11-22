import imaplib
import email
import io
import os
import socket
import openpyxl
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, StreamingResponse

app = FastAPI()
socket.setdefaulttimeout(15)

# --- CONFIGURAÇÕES DO E-MAIL (Vêm do Render) ---
EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASS = os.getenv("EMAIL_PASS", "")
IMAP_SERVER = os.getenv("IMAP_SERVER", "://gmail.com")
ASSUNTO_BUSCA = os.getenv("ASSUNTO_BUSCA", "Relatorio")

# --- LISTA DE CLIENTES (Nome exato da Aba -> Senha) ---
CLIENTES_ABAS = {
    "AZUMA": "AZ482",
    "CARLOS AUGUSTO": "CA719",
    "ED PESCADOS": "ED305",
    "GF.BRASIL": "GF941",
    "LAMEI": "LA628",
    "MAURICIO LAMAS": "ML154",
    "PUROMAR >> JA PESCADOS": "JA837",
    "T M J SUL (DYNAMIC)": "TS590",
    "TUDO DO MAR >> GUILHERME": "GU263",
    "VIEIRA PESCADOS": "VP746",
    "WS FOODS": "WS318",
    "ZEZÃO": "ZE895",
    "ALLFOODS": "AF512",
}

def baixar_anexo_excel():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, mensagens = mail.search(None, f'(SUBJECT "{ASSUNTO_BUSCA}")')
    ids_emails = mensagens.split()

    if not ids_emails:
        mail.logout()
        raise Exception(f"Nenhum e-mail encontrado com o assunto '{ASSUNTO_BUSCA}'.")

    ultimo_id = ids_emails[-1]
    _, dados_email = mail.fetch(ultimo_id, "(RFC822)")
    msg = email.message_from_bytes(dados_email)

    for parte in msg.walk():
        if parte.get_content_maintype() == "multipart" or parte.get("Content-Disposition") is None:
            continue

        nome_arquivo = parte.get_filename()
        if nome_arquivo and nome_arquivo.endswith(".xlsx"):
            conteudo = parte.get_payload(decode=True)
            mail.logout()
            return io.BytesIO(conteudo)

    mail.logout()
    raise Exception("Nenhuma planilha (.xlsx) encontrada no anexo do e-mail.")

@app.get("/", response_class=HTMLResponse)
def pagina_clientes(erro: str = "", selecionado: str = ""):
    alerta_html = f'<div class="erro">{erro}</div>' if erro else ""
    
    # Gera os botões para cada cliente
    botoes_html = ""
    for aba in CLIENTES_ABAS.keys():
        botoes_html += f"""
        <button type="button" class="btn-cliente" onclick="abrirModal('{aba}')">
            <span>{aba}</span>
            <span class="seta">→</span>
        </button>
        """

    # Se houve erro de senha, já reabre o modal no cliente selecionado
    script_reabrir = f"abrirModal('{selecionado}');" if selecionado else ""

    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Portal de Estoque - Clientes</title>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f1f5f9; margin: 0; padding: 20px; color: #1e293b; }}
            .container {{ max-width: 480px; margin: 0 auto; }}
            .header {{ text-align: center; margin-bottom: 24px; }}
            .header h1 {{ font-size: 22px; margin: 0 0 6px 0; color: #0f172a; }}
            .header p {{ font-size: 14px; color: #64748b; margin: 0; }}
            .lista {{ display: flex; flex-direction: column; gap: 10px; }}
            .btn-cliente {{ background: white; border: 1px solid #e2e8f0; padding: 16px 20px; border-radius: 10px; font-size: 15px; font-weight: 600; color: #1e293b; display: flex; justify-content: space-between; align-items: center; cursor: pointer; box-shadow: 0 1px 3px rgba(0,0,0,0.05); transition: all 0.15s; text-align: left; }}
            .btn-cliente:hover {{ border-color: #2563eb; background: #eff6ff; color: #1d4ed8; }}
            .seta {{ color: #94a3b8; font-weight: bold; }}
            .erro {{ background: #fee2e2; color: #b91c1c; padding: 12px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; text-align: center; font-weight: 500; }}
            
            /* Modal de Senha */
            .modal-bg {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(15, 23, 42, 0.6); justify-content: center; align-items: center; padding: 15px; box-sizing: border-box; }}
            .modal {{ background: white; padding: 25px; border-radius: 14px; width: 100%; max-width: 360px; box-shadow: 0 10px 25px rgba(0,0,0,0.2); }}
            .modal h3 {{ margin: 0 0 8px 0; font-size: 18px; color: #0f172a; }}
            .modal p {{ margin: 0 0 18px 0; font-size: 13px; color: #64748b; }}
            input[type="password"] {{ width: 100%; padding: 12px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 16px; margin-bottom: 16px; box-sizing: border-box; text-transform: uppercase; }}
            .acoes {{ display: flex; gap: 10px; }}
            .btn-cancelar {{ flex: 1; padding: 12px; background: #e2e8f0; border: none; border-radius: 8px; font-weight: 600; color: #475569; cursor: pointer; }}
            .btn-baixar {{ flex: 2; padding: 12px; background: #2563eb; border: none; border-radius: 8px; font-weight: 600; color: white; cursor: pointer; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Estoque Diário - Clientes</h1>
                <p>Selecione sua empresa abaixo para baixar sua planilha</p>
            </div>
            {alerta_html}
            <div class="lista">
                {botoes_html}
            </div>
        </div>

        <div class="modal-bg" id="modalSenha">
            <div class="modal">
                <h3 id="tituloCliente">Cliente</h3>
                <p>Digite sua senha de acesso para baixar a planilha atualizada.</p>
                <form action="/baixar-aba" method="post">
                    <input type="hidden" name="aba" id="inputAba">
                    <input type="password" name="senha" id="inputSenha" placeholder="Senha (ex: AB123)" required>
                    <div class="acoes">
                        <button type="button" class="btn-cancelar" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-baixar">Baixar Excel</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModal(nomeAba) {{
                document.getElementById('tituloCliente').innerText = nomeAba;
                document.getElementById('inputAba').value = nomeAba;
                document.getElementById('modalSenha').style.display = 'flex';
                document.getElementById('inputSenha').focus();
            }}
            function fecharModal() {{
                document.getElementById('modalSenha').style.display = 'none';
            }}
            {script_reabrir}
        </script>
    </body>
    </html>
    """

@app.post("/baixar-aba")
def baixar_aba(aba: str = Form(...), senha: str = Form(...)):
    senha_digitada = senha.strip().upper()

    # 1. Valida se o cliente existe e se a senha confere
    if aba not in CLIENTES_ABAS or CLIENTES_ABAS[aba] != senha_digitada:
        return HTMLResponse(
            content=pagina_clientes(erro=f"Senha incorreta para {aba}.", selecionado=aba),
            status_code=401
        )

    try:
        # 2. Baixa o arquivo Excel completo do e-mail
        arquivo_bytes = baixar_anexo_excel()
        wb = openpyxl.load_workbook(arquivo_bytes)

        # 3. Verifica se a aba existe na planilha
        if aba not in wb.sheetnames:
            return HTMLResponse(
                content=pagina_clientes(erro=f"A aba '{aba}' não foi encontrada na planilha de hoje."),
                status_code=404
            )

        # 4. Remove todas as outras abas, mantendo apenas a do cliente selecionado
        for nome_aba in wb.sheetnames:
            if nome_aba != aba:
                del wb[nome_aba]

        # 5. Salva o novo arquivo Excel com apenas 1 aba na memória
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        # Limpa caracteres especiais (como >>) para o nome do arquivo baixado
        nome_limpo = aba.replace(">>", "-").replace(" ", "_")
        headers = {"Content-Disposition": f'attachment; filename="Estoque_{nome_limpo}.xlsx"'}

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )

    except Exception as e:
        return HTMLResponse(
            content=pagina_clientes(erro=f"Erro ao gerar planilha: {str(e)}"),
            status_code=500
        )
