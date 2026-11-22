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

# --- LISTA DE CLIENTES (Nome do Estoque -> Senha) ---
CLIENTES_ESTOQUE = {
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

def formatar_numero(val):
    if val is None or val == "":
        return "0"
    if isinstance(val, float):
        if val.is_integer():
            return f"{int(val):,}".replace(",", ".")
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if isinstance(val, int):
        return f"{val:,}".replace(",", ".")
    return str(val).strip()

# --- TELA 1: LOGIN COM COMBO DE ESTOQUES + SENHA ---
@app.get("/", response_class=HTMLResponse)
def tela_login(erro: str = "", selecionado: str = ""):
    alerta_html = f'<div class="erro">{erro}</div>' if erro else ""
    
    opcoes_html = '<option value="" disabled selected>Selecione o seu estoque...</option>'
    for cliente in CLIENTES_ESTOQUE.keys():
        sel = "selected" if cliente == selecionado else ""
        if sel:
            opcoes_html = opcoes_html.replace('selected', '')
        opcoes_html += f'<option value="{cliente}" {sel}>{cliente}</option>'

    return f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Portal de Estoque - Amazon Norte</title>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0f172a; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 16px; box-sizing: border-box; }}
            .card {{ background: white; padding: 28px 24px; border-radius: 16px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.3); width: 100%; max-width: 400px; }}
            .topo-card {{ text-align: center; margin-bottom: 22px; }}
            .badge {{ display: inline-block; background: #dbeafe; color: #1e40af; font-size: 11px; font-weight: 700; padding: 4px 12px; border-radius: 999px; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 0.5px; }}
            h1 {{ font-size: 22px; margin: 0 0 6px 0; color: #0f172a; }}
            p {{ font-size: 14px; color: #64748b; margin: 0; }}
            label {{ display: block; font-size: 13px; font-weight: 600; color: #334155; margin-bottom: 6px; }}
            select, input[type="password"] {{ width: 100%; padding: 13px 14px; border: 2px solid #cbd5e1; border-radius: 10px; font-size: 16px; margin-bottom: 16px; box-sizing: border-box; background: white; color: #0f172a; outline: none; transition: border-color 0.2s; }}
            select:focus, input[type="password"]:focus {{ border-color: #2563eb; }}
            input[type="password"] {{ text-transform: uppercase; letter-spacing: 2px; font-weight: 600; }}
            button {{ width: 100%; padding: 14px; background: #2563eb; color: white; border: none; border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s; margin-top: 4px; }}
            button:hover {{ background: #1d4ed8; }}
            .erro {{ background: #fee2e2; color: #b91c1c; padding: 12px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; font-weight: 500; text-align: center; }}
        </style>
    </head>
    <body>
        <div class="card">
            <div class="topo-card">
                <span class="badge">Amazon Norte Pescados</span>
                <h1>Portal de Estoque</h1>
                <p>Escolha seu estoque e digite sua senha de acesso</p>
            </div>
            {alerta_html}
            <form action="/painel" method="post">
                <label for="cliente">Estoque / Cliente</label>
                <select id="cliente" name="cliente" required>
                    {opcoes_html}
                </select>

                <label for="senha">Senha de Acesso</label>
                <input type="password" id="senha" name="senha" placeholder="Digite sua senha" maxlength="15" required>

                <button type="submit">Visualizar Estoque</button>
            </form>
        </div>
    </body>
    </html>
    """

# --- TELA 2: VISUALIZAÇÃO OTIMIZADA PARA CELULAR (CARDS) + DOWNLOAD ---
@app.post("/painel", response_class=HTMLResponse)
def painel_cliente(cliente: str = Form(...), senha: str = Form(...)):
    senha_limpa = senha.strip().upper()

    if cliente not in CLIENTES_ESTOQUE or CLIENTES_ESTOQUE[cliente] != senha_limpa:
        return HTMLResponse(
            content=tela_login(erro="Senha incorreta para o estoque selecionado.", selecionado=cliente),
            status_code=401
        )

    try:
        arquivo_bytes, nome_original = obter_planilha_mais_recente()
        wb = openpyxl.load_workbook(arquivo_bytes, data_only=True)
        aba_encontrada = localizar_aba_cliente(wb, cliente)

        if not aba_encontrada:
            return HTMLResponse(
                content=tela_login(erro=f"Não há saldo registrado para {cliente} no relatório de hoje.", selecionado=cliente),
                status_code=404
            )

        ws = wb[aba_encontrada]

        data_relatorio = "Atualizado"
        if "_" in nome_original:
            parte_data = nome_original.split("_")[-1].replace(".xlsx", "")
            partes = parte_data.split("-")
            if len(partes) == 3:
                data_relatorio = f"{partes}/{partes}/{partes}"

        # Mapeia dinamicamente os índices das colunas a partir da linha de cabeçalho (Cod.Ext, Produto, Vol, Saldo Kgs)
        col_map = {"cod": 0, "nf": 1, "produto": 2, "sec": 3, "pri": 4, "vol": 5, "kgs": 6}
        itens_html = ""
        total_geral_vol = ""
        total_geral_kgs = ""

        for row in ws.iter_rows(values_only=True):
            valores_nao_vazios = [str(v).strip() for v in row if v is not None and str(v).strip() != ""]
            if not valores_nao_vazios:
                continue

            texto_linha = " ".join(valores_nao_vazios).upper()

            # Ignora linhas de título geral ou links de navegação
            if "VOLTAR AO RESUMO" in texto_linha or "DENTRO DA VALIDADE" in texto_linha or "DATA-BASE" in texto_linha:
                continue

            # Identifica cabeçalho da tabela para mapear as colunas exatas
            if "PRODUTO" in texto_linha and ("VOL" in texto_linha or "SALDO" in texto_linha):
                for idx, val in enumerate(row):
                    v_up = str(val or "").strip().upper()
                    if "COD" in v_up and "NF" not in v_up:
                        col_map["cod"] = idx
                    elif "NF" in v_up:
                        col_map["nf"] = idx
                    elif "PRODUTO" in v_up:
                        col_map["produto"] = idx
                    elif "SECUND" in v_up:
                        col_map["sec"] = idx
                    elif "PRIM" in v_up:
                        col_map["pri"] = idx
                    elif "VOL" in v_up:
                        col_map["vol"] = idx
                    elif "KG" in v_up or "SALDO" in v_up:
                        col_map["kgs"] = idx
                continue

            # Ignora a linha do título do cliente no topo da aba
            if valores_nao_vazios[0].upper() == cliente.strip().upper():
                continue

            # Captura Total Geral da Categoria
            if "TOTAL DA CATEGORIA" in texto_linha:
                vol_val = row[col_map["vol"]] if len(row) > col_map["vol"] else ""
                kgs_val = row[col_map["kgs"]] if len(row) > col_map["kgs"] else ""
                total_geral_vol = formatar_numero(vol_val)
                total_geral_kgs = formatar_numero(kgs_val)
                continue

            # Captura Subtotais ou Totais de Bloco (ex: NACIONAL / EXPORTAÇÃO)
            if "SUBTOTAL" in texto_linha or texto_linha.startswith("TOTAL "):
                titulo_sub = next((v for v in valores_nao_vazios if "TOTAL" in v.upper()), "Subtotal")
                vol_val = row[col_map["vol"]] if len(row) > col_map["vol"] else ""
                kgs_val = row[col_map["kgs"]] if len(row) > col_map["kgs"] else ""
                itens_html += f"""
                <div class="card-subtotal">
                    <span class="sub-titulo">{titulo_sub}</span>
                    <div class="sub-valores">
                        <span>Vol: <b>{formatar_numero(vol_val)}</b></span>
                        <span>Saldo: <b>{formatar_numero(kgs_val)} kg</b></span>
                    </div>
                </div>
                """
                continue

            # Captura divisores de bloco (ex: NACIONAL, EXPORTAÇÃO)
            if len(valores_nao_vazios) == 1 and valores_nao_vazios[0].upper() in ["NACIONAL", "EXPORTAÇÃO", "COM REGISTRO", "SEM REGISTRO"]:
                itens_html += f'<div class="divisor-bloco">{valores_nao_vazios[0].upper()}</div>'
                continue

            # Linha normal de produto
            if len(row) > col_map["produto"] and row[col_map["produto"]]:
                cod = str(row[col_map["cod"]] or "").strip()
                produto = str(row[col_map["produto"]] or "").strip()
                sec = str(row[col_map["sec"]] or "").strip()
                pri = str(row[col_map["pri"]] or "").strip()
                vol = formatar_numero(row[col_map["vol"]] if len(row) > col_map["vol"] else 0)
                kgs = formatar_numero(row[col_map["kgs"]] if len(row) > col_map["kgs"] else 0)

                embalagem = " · ".join([p for p in [sec, pri] if p])

                itens_html += f"""
                <div class="card-item">
                    <div class="item-topo">
                        <span class="item-cod">{cod}</span>
                        <span class="item-emb">{embalagem}</span>
                    </div>
                    <div class="item-nome">{produto}</div>
                    <div class="item-rodape">
                        <div class="metrica">
                            <span class="m-label">Volumes</span>
                            <span class="m-valor">{vol}</span>
                        </div>
                        <div class="metrica destaque">
                            <span class="m-label">Saldo Kgs</span>
                            <span class="m-valor">{kgs} kg</span>
                        </div>
                    </div>
                </div>
                """

        resumo_topo_html = ""
        if total_geral_vol or total_geral_kgs:
            resumo_topo_html = f"""
            <div class="resumo-geral">
                <div class="resumo-box">
                    <span class="r-label">Total Volumes</span>
                    <span class="r-num">{total_geral_vol}</span>
                </div>
                <div class="resumo-box">
                    <span class="r-label">Saldo Total</span>
                    <span class="r-num">{total_geral_kgs} kg</span>
                </div>
            </div>
            """

        if not itens_html:
            itens_html = '<div class="vazio">Sem estoque registrado neste relatório.</div>'

        return f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Estoque - {cliente}</title>
            <style>
                body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f1f5f9; margin: 0; padding: 14px; color: #1e293b; }}
                .container {{ max-width: 600px; margin: 0 auto; }}
                
                .header {{ background: #0f172a; color: white; padding: 20px; border-radius: 14px; margin-bottom: 14px; }}
                .header-top {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; margin-bottom: 14px; }}
                .header h1 {{ margin: 0 0 4px 0; font-size: 19px; line-height: 1.2; }}
                .header .data {{ font-size: 13px; color: #94a3b8; }}
                .btn-sair {{ background: rgba(255,255,255,0.12); color: white; text-decoration: none; padding: 8px 14px; border-radius: 8px; font-size: 13px; font-weight: 600; white-space: nowrap; }}
                
                .resumo-geral {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 14px; }}
                .resumo-box {{ background: rgba(255,255,255,0.08); padding: 12px; border-radius: 10px; }}
                .r-label {{ display: block; font-size: 11px; color: #94a3b8; text-transform: uppercase; margin-bottom: 2px; }}
                .r-num {{ font-size: 18px; font-weight: 700; color: #38bdf8; }}

                .btn-excel {{ width: 100%; background: #16a34a; color: white; border: none; padding: 13px; border-radius: 10px; font-size: 15px; font-weight: 600; cursor: pointer; display: flex; justify-content: center; align-items: center; gap: 8px; }}
                .btn-excel:hover {{ background: #15803d; }}

                .divisor-bloco {{ font-size: 12px; font-weight: 700; color: #475569; letter-spacing: 0.8px; margin: 18px 0 8px 4px; text-transform: uppercase; }}
                
                .card-item {{ background: white; border-radius: 12px; padding: 14px 16px; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; }}
                .item-topo {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; font-size: 12px; }}
                .item-cod {{ background: #eff6ff; color: #1d4ed8; font-weight: 700; padding: 2px 8px; border-radius: 6px; }}
                .item-emb {{ color: #64748b; font-weight: 500; }}
                .item-nome {{ font-size: 14px; font-weight: 600; color: #0f172a; line-height: 1.4; margin-bottom: 12px; }}
                .item-rodape {{ display: flex; justify-content: space-between; border-top: 1px solid #f1f5f9; padding-top: 10px; }}
                .metrica {{ display: flex; flex-direction: column; }}
                .metrica.destaque {{ text-align: right; }}
                .m-label {{ font-size: 11px; color: #64748b; }}
                .m-valor {{ font-size: 15px; font-weight: 700; color: #1e293b; }}
                .metrica.destaque .m-valor {{ color: #0f172a; }}

                .card-subtotal {{ background: #e2e8f0; border-radius: 10px; padding: 12px 16px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }}
                .sub-titulo {{ font-size: 13px; font-weight: 700; color: #334155; }}
                .sub-valores {{ display: flex; gap: 14px; font-size: 13px; color: #0f172a; }}
                .vazio {{ background: white; padding: 30px; text-align: center; border-radius: 12px; color: #64748b; font-size: 14px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <div class="header-top">
                        <div>
                            <h1>{cliente}</h1>
                            <div class="data">📅 Posição: {data_relatorio}</div>
                        </div>
                        <a href="/" class="btn-sair">Sair</a>
                    </div>
                    {resumo_topo_html}
                    <form action="/baixar-excel" method="post" style="margin:0;">
                        <input type="hidden" name="cliente" value="{cliente}">
                        <input type="hidden" name="senha" value="{senha_limpa}">
                        <button type="submit" class="btn-excel">⬇ Baixar Planilha (.xlsx)</button>
                    </form>
                </div>

                {itens_html}
            </div>
        </body>
        </html>
        """

    except Exception as e:
        return HTMLResponse(content=tela_login(erro=f"Erro ao carregar estoque: {str(e)}", selecionado=cliente), status_code=500)

# --- ROTA DE DOWNLOAD DO ARQUIVO EXCEL ---
@app.post("/baixar-excel")
def baixar_excel(cliente: str = Form(...), senha: str = Form(...)):
    senha_limpa = senha.strip().upper()
    if cliente not in CLIENTES_ESTOQUE or CLIENTES_ESTOQUE[cliente] != senha_limpa:
        return HTMLResponse(content=tela_login(erro="Sessão inválida."), status_code=401)

    try:
        arquivo_bytes, nome_original = obter_planilha_mais_recente()
        wb = openpyxl.load_workbook(arquivo_bytes)
        aba_encontrada = localizar_aba_cliente(wb, cliente)

        if not aba_encontrada:
            return HTMLResponse(content=tela_login(erro="Aba não encontrada."), status_code=404)

        for nome_aba in list(wb.sheetnames):
            if nome_aba != aba_encontrada:
                del wb[nome_aba]

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        nome_limpo = cliente.replace(">>", "-").replace(" ", "_")
        sufixo_data = nome_original.replace("ESTOQUE", "").replace(".xlsx", "")
        headers = {"Content-Disposition": f'attachment; filename="{nome_limpo}{sufixo_data}.xlsx"'}

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
    except Exception as e:
        return HTMLResponse(content=tela_login(erro=str(e)), status_code=500)
