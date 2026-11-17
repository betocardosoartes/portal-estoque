const express = require('express');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

const URL_LOGIN_SGEP = process.env.SGEP_URL || 'https://sistemassartec.com.br';
const URL_ESTOQUE_SGEP = process.env.SGEP_ESTOQUE_URL || 'https://sistemassartec.com.br/ConsEstoqueProdutoNovo/FiltroConsultaEstoque';
const USUARIO_SGEP = process.env.SGEP_USER || 'Beto';
const SENHA_SGEP = process.env.SGEP_PASS || '@464';

const CLIENTES_CADASTRADOS = {
  'CA': {
    nomeCompleto: 'C A DOS PASSOS',
    filtroSgep: 'C A DOS PASSOS',
    senhaAcesso: '123456'
  }
  /* EXEMPLO PARA ADICIONAR MAIS CLIENTES:
  ,
  'XY': {
    nomeCompleto: 'NOME COMPLETO DO CLIENTE',
    filtroSgep: 'NOME NO FILTRO DO SGEP',
    senhaAcesso: 'SENHA_DELE'
  }
  */
};

app.get('/', (req, res) => {
  const opcoesHtml = Object.entries(CLIENTES_CADASTRADOS)
    .map(([sigla]) => `<option value="${sigla}">${sigla}</option>`)
    .join('');

  res.send(`
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Consulta de Estoque</title>
      <style>
        body { font-family: Arial, sans-serif; background: #f0f2f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: white; padding: 28px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); width: 90%; max-width: 360px; }
        h2 { color: #1F3864; margin-top: 0; text-align: center; }
        label { display: block; margin-top: 15px; font-weight: bold; color: #333; font-size: 14px; }
        select, input { width: 100%; padding: 12px; margin-top: 6px; border: 1px solid #ccc; border-radius: 8px; font-size: 16px; box-sizing: border-box; }
        button { width: 100%; padding: 14px; background: #1F3864; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; margin-top: 22px; cursor: pointer; }
        button:disabled { background: #7a8ba8; }
        #status { margin-top: 15px; text-align: center; font-size: 14px; color: #1F3864; font-weight: bold; }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Estoque Atualizado</h2>
        <label for="sigla">Selecione sua Sigla:</label>
        <select id="sigla">${opcoesHtml}</select>

        <label for="senha">Senha de Acesso:</label>
        <input type="password" id="senha" placeholder="Digite sua senha">

        <button id="btnBaixar" onclick="solicitarPlanilha()">Baixar Planilha</button>
        <div id="status"></div>
      </div>

      <script>
        async function solicitarPlanilha() {
          const sigla = document.getElementById('sigla').value;
          const senha = document.getElementById('senha').value;
          const btn = document.getElementById('btnBaixar');
          const status = document.getElementById('status');

          if (!senha) {
            status.style.color = 'red';
            status.innerText = 'Digite a senha cadastrada.';
            return;
          }

          btn.disabled = true;
          status.style.color = '#1F3864';
          status.innerText = 'Consultando estoque e formatando planilha... Aguarde cerca de 15 segundos.';

          try {
            const response = await fetch('/baixar', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ sigla, senha })
            });

            if (!response.ok) {
              const erro = await response.json();
              throw new Error(erro.mensagem || 'Erro ao gerar arquivo');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Estoque_' + sigla + '.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);

            status.style.color = 'green';
            status.innerText = 'Download concluído com sucesso!';
          } catch (err) {
            status.style.color = 'red';
            status.innerText = err.message;
          } finally {
            btn.disabled = false;
          }
        }
      </script>
    </body>
    </html>
  `);
});

app.post('/baixar', async (req, res) => {
  const { sigla, senha } = req.body;
  const cliente = CLIENTES_CADASTRADOS[sigla];

  if (!cliente || cliente.senhaAcesso !== senha) {
    return res.status(401).json({ mensagem: 'Sigla ou senha incorreta.' });
  }

  let browser;
  try {
    browser = await chromium.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    await page.goto(URL_LOGIN_SGEP, { waitUntil: 'networkidle' });
    await page.fill('input[type="text"], input[name*="Login"], input[name*="Usuario"]', USUARIO_SGEP);
    await page.fill('input[type="password"]', SENHA_SGEP);
    await page.click('button[type="submit"], input[type="submit"]');
    await page.waitForLoadState('networkidle');

    await page.goto(URL_ESTOQUE_SGEP, { waitUntil: 'networkidle' });

    const seletorCliente = page.locator('.bootstrap-select').filter({ hasText: /Nada selecionado|Cliente/i }).first();
    if (await seletorCliente.count() > 0) {
      await seletorCliente.locator('button').click();
    } else {
      await page.locator('.bootstrap-select button').first().click();
    }

    await page.keyboard.type(cliente.filtroSgep, { delay: 50 });
    await page.keyboard.press('Enter');
    await page.waitForTimeout(1000);

    const btnPesquisar = page.locator('button:has-text("Pesquisar"), button:has-text("Consultar"), button:has-text("Filtrar")').first();
    if (await btnPesquisar.count() > 0) {
      await btnPesquisar.click();
      await page.waitForLoadState('networkidle');
    }

    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 30000 }),
      page.locator('button:has-text("Excel"), a:has-text("Excel"), .buttons-excel, [title*="Excel"]').first().click()
    ]);

    const caminhoBruto = path.join(__dirname, `temp_${sigla}_${Date.now()}.xlsx`);
    await download.saveAs(caminhoBruto);

    const agora = new Date();
    const bufferFormatado = await formatarPlanilha(caminhoBruto, cliente.nomeCompleto, sigla, agora);

    fs.unlinkSync(caminhoBruto);
    await browser.close();

    const dia = String(agora.getDate()).padStart(2, '0');
    const mes = String(agora.getMonth() + 1).padStart(2, '0');
    const ano = agora.getFullYear();
    const hora = String(agora.getHours()).padStart(2, '0');
    const minuto = String(agora.getMinutes()).padStart(2, '0');
    const nomeFinal = `Estoque_${cliente.nomeCompleto.replace(/\s+/g, '_')}_${dia}-${mes}-${ano}_${hora}h${minuto}.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename="${nomeFinal}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(bufferFormatado);

  } catch (erro) {
    if (browser) await browser.close();
    console.error('Erro na automação WebSGEP:', erro);
    res.status(500).json({ mensagem: 'Não foi possível gerar a planilha agora. Verifique o filtro ou tente novamente.' });
  }
});

function extrairEspecie(nomeProduto) {
  const texto = String(nomeProduto || '').trim().toUpperCase();
  if (!texto) return 'OUTROS';
  const partes = texto.split(/[\.\s]+/).filter(Boolean);
  const primeira = partes || '';
  const segunda = partes || '';

  if (primeira === 'P' || primeira === 'PESC' || primeira === 'PESCADA') {
    const mapa = { 'AMARELA': 'PESCADA AMARELA', 'BANANA': 'PESCADA BANANA', 'BRANCA': 'PESCADA BRANCA', 'CURUCA': 'PESCADA CURUCA', 'GO': 'PESCADA GO' };
    return segunda ? (mapa[segunda] || `PESCADA ${segunda}`) : 'PESCADA';
  }
  if (primeira === 'CAM' || primeira === 'CAMARAO') {
    return (segunda && segunda !== 'N') ? `CAMARAO ${segunda}` : 'CAMARAO ROSA';
  }
  return primeira || 'OUTROS';
}

async function formatarPlanilha(caminhoArquivo, nomeCliente, sigla, agora) {
  const wbOrig = new ExcelJS.Workbook();
  await wbOrig.xlsx.load(fs.readFileSync(caminhoArquivo));
  const sheetOrig = wbOrig.worksheets;

  let dados = [];
  sheetOrig.eachRow((row) => {
    const codExt = String(row.getCell(1).value || '').trim();
    const produto = String(row.getCell(3).value || '').trim();
    const prodUpper = produto.toUpperCase();

    if (!produto || prodUpper.startsWith('SGEP') || codExt.toUpperCase().startsWith('SGEP')) return;
    if (prodUpper === 'PRODUTO' || codExt.toUpperCase() === 'COD.EXT' || codExt === produto) return;
    if (prodUpper.startsWith('TOTAL') || prodUpper.startsWith('SUBTOTAL')) return;

    const parseNum = (v) => typeof v === 'number' ? v : (Number(String(v || 0).replace(/\./g, '').replace(',', '.')) || 0);

    dados.push({
      codExt: row.getCell(1).value,
      codNf: row.getCell(2).value,
      produto,
      secundaria: row.getCell(4).value,
      primaria: row.getCell(5).value,
      vol: parseNum(row.getCell(6).value),
      saldoKgs: parseNum(row.getCell(7).value),
      especie: extrairEspecie(produto)
    });
  });

  const grupos = {};
  dados.forEach(item => {
    if (!grupos[item.especie]) grupos[item.especie] = [];
    grupos[item.especie].push(item);
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Estoque Formatado');

  ws.getCell('A1').value = nomeCliente;
ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F3864' } };
const dataAtual = agora.toLocaleDateString('pt-BR');
ws.getCell('A2').value = Dentro da validade + cliente ${sigla}. • Data-base ${dataAtual};
ws.getCell('A2').font = { italic: true, size: 9, color: { argb: 'FF595959' } };
const headerRow = ws.getRow(5);
headerRow.values = ['Cod.Ext', 'Cód. NF', 'Produto', 'Secundária', 'Primária', 'Vol', 'Saldo Kgs'];
headerRow.eachCell(c => {
c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
});
let rowIdx = 6, totVol = 0, totKgs = 0;
for (const [especie, itens] of Object.entries(grupos)) {
let subVol = 0, subKgs = 0;
itens.forEach(i => {
const r = ws.getRow(rowIdx++);
r.values = [i.codExt, i.codNf, i.produto, i.secundaria, i.primaria, i.vol, i.saldoKgs];
r.getCell(6).numFmt = '#,##0';
r.getCell(7).numFmt = '#,##0.00';
subVol += i.vol;
subKgs += i.saldoKgs;
});
const subRow = ws.getRow(rowIdx++);
subRow.getCell(3).value = Subtotal ${especie};
subRow.getCell(6).value = subVol;
subRow.getCell(7).value = subKgs;
subRow.eachCell({ includeEmpty: true }, (c, col) => {
if (col <= 7) {
c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
c.font = { bold: true, italic: true };
}
});
subRow.getCell(6).numFmt = '#,##0';
subRow.getCell(7).numFmt = '#,##0.00';
totVol += subVol;
totKgs += subKgs;
}
ws.getCell('E1').value = 'TOTAL';
ws.getCell('F1').value = totVol;
ws.getCell('G1').value = totKgs;
ws.getCell('F1').numFmt = '#,##0';
ws.getCell('G1').numFmt = '#,##0.00';
const totRow = ws.getRow(rowIdx);
totRow.getCell(3).value = 'TOTAL DA CATEGORIA';
totRow.getCell(6).value = totVol;
totRow.getCell(7).value = totKgs;
totRow.font = { bold: true };
totRow.getCell(6).numFmt = '#,##0';
totRow.getCell(7).numFmt = '#,##0.00';
ws.columns = [{ width: 12 }, { width: 12 }, { width: 58 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 16 }];
return await wb.xlsx.writeBuffer();
}
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(Portal rodando na porta ${PORT}));
