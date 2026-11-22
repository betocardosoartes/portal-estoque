import imaplib
import email
from email.header import decode_header
import io
import os
import socket
import time
import openpyxl
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, StreamingResponse

app = FastAPI()
socket.setdefaulttimeout(15)

EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASS = os.getenv("EMAIL_PASS", "")
IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.titan.email")
ASSUNTO_BUSCA = os.getenv("ASSUNTO_BUSCA", "SGEP - Estoque")

# --- SENHA -> NOME EXATO DA ABA DO CLIENTE ---
SENHAS_CLIENTES = {
    "AZ482": "AZUMA",
    "CA719": "CARLOS AUGUSTO",
    "ED305": "ED PESCADOS",
    "GF941": "GF.BRASIL",
    "LA628": "LAMEI",
    "ML154": "MAURICIO LAMAS",
    "JA837": "PUROMAR >> JA PESCADOS",
    "TS590": "T M J SUL (DYNAMIC)",
    "GU263": "TUDO DO MAR >> GUILHERME",
    "VP746": "VIEIRA PESCADOS",
    "WS318": "WS FOODS",
    "ZE895": "ZEZÃO",
    "AF512": "ALLFOODS",
}

# Cache simples em memória (10 minutos) para deixar o portal instantâneo
CACHE_PLANILHA = {"timestamp": 0, "bytes": None, "nome_arquivo": ""}
TEMPO_CACHE_SEGUNDOS = 600

def decodificar_texto(texto_bruto):
    if not texto_bruto:
        return ""
    partes = decode_header(texto_bruto)
    resultado = ""
    for texto, encoding in partes:
        if isinstance(texto, bytes):
            resultado += texto.decode(encoding or "utf-8", errors="ignore")
        else:
            resultado += texto
    return resultado

def obter_planilha_mais_recente():
    agora = time.time()
    if CACHE_PLANILHA["bytes"] and (agora - CACHE_PLANILHA["timestamp"] < TEMPO_CACHE_SEGUNDOS):
        return io.BytesIO(CACHE_PLANILHA["bytes"]), CACHE_PLANILHA["nome_arquivo"]

    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, mensagens = mail.search(None, f'(SUBJECT "{ASSUNTO_BUSCA}")')
    ids_emails = mensagens[0].split()

    if not ids_emails:
        mail.logout()
        raise Exception(f"Nenhum relatório encontrado no e-mail ({ASSUNTO_BUSCA}).")

    for email_id in reversed(ids_emails[-10:]):
        _, dados_email = mail.fetch(email_id, "(RFC822)")
        msg = email.message_from_bytes(dados_email[0][1])

        for parte in msg.walk():
            if parte.get_content_maintype() == "multipart" or parte.get("Content-Disposition") is None:
                continue

            nome_arquivo = decodificar_texto(parte.get_filename())
            if nome_arquivo and nome_arquivo.lower().endswith(".xlsx"):
                conteudo = parte.get_payload(decode=True)
                mail.logout()
                CACHE_PLANILHA["timestamp"] = agora
                CACHE_PLANILHA["bytes"] = conteudo
                CACHE_PLANILHA["nome_arquivo"] = nome_arquivo
                return io.BytesIO(conteudo), nome_arquivo

    mail.logout()
    raise Exception("Relatório encontrado, mas nenhuma planilha .xlsx estava anexada.")

def localizar_aba_cliente(wb, nome_cliente):
    alvo = nome_cliente.strip().upper()
    for nome_real in wb.sheetnames:
        if nome_real.strip().upper() == alvo or alvo in nome_real.strip().upper():
            return nome_real
    return None

def formatar_valor_celula(val):
    if val is None:
        return ""
    if isinstance(val, float):
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if isinstance(val, int):
        return f"{val:,}".replace(",", ".")
    return str(val).strip()

# --- TELA 1: LOGIN EXCLUSIVO POR SENHA ---
@app.get("/", response_class=HTMLResponse)
def tela_login(erro: str = ""):
    alerta_html = f'<div class="erro">{erro}</div>' if erro else ""
    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Portal de Estoque - Amazon Norte</title>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }}
            .card {{ background: white; padding: 32px 28px; border-radius: 16px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.3); width: 100%; max-width: 380px; text-align: center; }}
            .badge {{ display: inline-block; background: #dbeafe; color: #1e40af; font-size: 12px; font-weight: 700; padding: 4px 12px; border-radius: 999px; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 0.5px; }}
            h1 {{ font-size: 22px; margin: 0 0 8px 0; color: #0f172a; }}
            p {{ font-size: 14px; color: #64748b; margin: 0 0 24px 0; line-height: 1.5; }}
            input[type="password"] {{ width: 100%; padding: 14px; border: 2px solid #cbd5e1; border-radius: 10px; font-size: 18px; text-align: center; letter-spacing: 3px; font-weight: bold; margin-bottom: 16px; box-sizing: border-box; text-transform: uppercase; outline: none; transition: border-color 0.2s; }}
            input[type="password"]:focus {{ border-color: #2563eb; }}
            button {{ width: 100%; padding: 14px; background: #2563eb; color: white; border: none; border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s; }}
            button:hover {{ background: #1d4ed8; }}
            .erro {{ background: #fee2e2; color: #b91c1c; padding: 12px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; font-weight: 500; }}
            .rodape {{ margin-top: 20px; font-size: 12px; color: #94a3b8; }}
        </style>
    </head>
    <body>
        <div class="card">
            <span class="badge">Amazon Norte Pescados</span>
            <h1>Consulta de Estoque</h1>
            <p>Digite sua chave de acesso individual para visualizar e baixar seu saldo atualizado.</p>
            {alerta_html}
            <form action="/painel" method="post">
                <input type="password" name="senha" placeholder="•••••" maxlength="15" required autofocus>
                <button type="submit">Acessar Meu Estoque</button>
            </form>
            <div class="rodape">Acesso restrito e individual por cliente</div>
        </div>
    </body>
    </html>
    """

# --- TELA 2: VISUALIZAÇÃO DO ESTOQUE DO CLIENTE NA TELA ---
@app.post("/painel", response_class=HTMLResponse)
def painel_cliente(senha: str = Form(...)):
    senha_limpa = senha.strip().upper()

    if senha_limpa not in SENHAS_CLIENTES:
        return HTMLResponse(content=tela_login(erro="Chave de acesso inválida. Verifique e tente novamente."), status_code=401)

    nome_cliente = SENHAS_CLIENTES[senha_limpa]

    try:
        arquivo_bytes, nome_original = obter_planilha_mais_recente()
        wb = openpyxl.load_workbook(arquivo_bytes, data_only=True)
        aba_encontrada = localizar_aba_cliente(wb, nome_cliente)

        if not aba_encontrada:
            return HTMLResponse(content=tela_login(erro=f"Não há saldo registrado para {nome_cliente} no relatório de hoje."), status_code=404)

        ws = wb[aba_encontrada]

        data_relatorio = "Atualizado"
        if "_" in nome_original:
            parte_data = nome_original.split("_")[-1].replace(".xlsx", "")
            partes = parte_data.split("-")
            if len(partes) == 3:
                data_relatorio = f"{partes[2]}/{partes[1]}/{partes[0]}"

        linhas_html = ""
        for row in ws.iter_rows(values_only=True):
            valores = [v for v in row if v is not None and str(v).strip() != ""]
            if not valores:
                continue

            texto_linha = " ".join(str(v) for v in valores).upper()
            if "VOLTAR AO RESUMO" in texto_linha:
                continue

            celulas = list(row)
            while celulas and (celulas[-1] is None or str(celulas[-1]).strip() == ""):
                celulas.pop()

            eh_total = "TOTAL" in texto_linha or "SUBTOTAL" in texto_linha
            eh_cabecalho = "PRODUTO" in texto_linha and ("VOL" in texto_linha or "SALDO" in texto_linha)

            classe_tr = ""
            if eh_total:
                classe_tr = 'class="linha-total"'
            elif eh_cabecalho:
                classe_tr = 'class="linha-cabecalho"'

            tds = ""
            for idx, cel in enumerate(celulas):
                val_formatado = formatar_valor_celula(cel)
                alinhamento = 'style="text-align: right;"' if isinstance(cel, (int, float)) else ""
                tds += f"<td {alinhamento}>{val_formatado}</td>"

            linhas_html += f"<tr {classe_tr}>{tds}</tr>"

        return f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Estoque - {nome_cliente}</title>
            <style>
                body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f1f5f9; margin: 0; padding: 15px; color: #1e293b; }}
                .container {{ max-width: 960px; margin: 0 auto; }}
                .topo {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 15px; margin-bottom: 16px; }}
                .info h1 {{ margin: 0 0 4px 0; font-size: 20px; color: #0f172a; }}
                .info span {{ font-size: 13px; color: #64748b; font-weight: 500; }}
                .acoes {{ display: flex; gap: 10px; align-items: center; }}
                .btn-excel {{ background: #16a34a; color: white; border: none; padding: 12px 18px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 6px; }}
                .btn-excel:hover {{ background: #15803d; }}
                .btn-sair {{ background: #e2e8f0; color: #475569; text-decoration: none; padding: 12px 14px; border-radius: 8px; font-size: 14px; font-weight: 600; }}
                
                .tabela-box {{ background: white; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); overflow-x: auto; }}
                table {{ width: 100%; border-collapse: collapse; font-size: 13px; white-space: nowrap; }}
                td {{ padding: 10px 12px; border-bottom: 1px solid #e2e8f0; }}
                .linha-cabecalho td {{ background: #1e293b; color: white; font-weight: 600; }}
                .linha-total td {{ background: #f8fafc; font-weight: 700; color: #0f172a; border-top: 2px solid #cbd5e1; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="topo">
                    <div class="info">
                        <h1>{nome_cliente}</h1>
                        <span>📅 Posição de Estoque: {data_relatorio}</span>
                    </div>
                    <div class="acoes">
                        <form action="/baixar-excel" method="post" style="margin:0;">
                            <input type="hidden" name="senha" value="{senha_limpa}">
                            <button type="submit" class="btn-excel">⬇ Baixar Planilha (.xlsx)</button>
                        </form>
                        <a href="/" class="btn-sair">Sair</a>
                    </div>
                </div>

                <div class="tabela-box">
                    <table>
                        {linhas_html}
                    </table>
                </div>
            </div>
        </body>
        </html>
        """

    except Exception as e:
        return HTMLResponse(content=tela_login(erro=f"Erro ao carregar estoque: {str(e)}"), status_code=500)

# --- ROTA DE DOWNLOAD DO ARQUIVO EXCEL DO CLIENTE ---
@app.post("/baixar-excel")
def baixar_excel(senha: str = Form(...)):
    senha_limpa = senha.strip().upper()
    if senha_limpa not in SENHAS_CLIENTES:
        return HTMLResponse(content=tela_login(erro="Sessão inválida."), status_code=401)

    nome_cliente = SENHAS_CLIENTES[senha_limpa]

    try:
        arquivo_bytes, nome_original = obter_planilha_mais_recente()
        wb = openpyxl.load_workbook(arquivo_bytes)
        aba_encontrada = localizar_aba_cliente(wb, nome_cliente)

        if not aba_encontrada:
            return HTMLResponse(content=tela_login(erro="Aba não encontrada."), status_code=404)

        for nome_aba in list(wb.sheetnames):
            if nome_aba != aba_encontrada:
                del wb[nome_aba]

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        nome_limpo = nome_cliente.replace(">>", "-").replace(" ", "_")
        sufixo_data = nome_original.replace("ESTOQUE", "").replace(".xlsx", "")
        headers = {"Content-Disposition": f'attachment; filename="{nome_limpo}{sufixo_data}.xlsx"'}

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
    except Exception as e:
        return HTMLResponse(content=tela_login(erro=str(e)), status_code=500)
