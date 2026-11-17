const express = require('express');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const DOMINIO_SGEP = 'https://' + 'amazonnorte' + '.sistemassartec.com.br';
const URL_LOGIN_SGEP = DOMINIO_SGEP + '/WebSGEP';
const URL_ESTOQUE_SGEP = DOMINIO_SGEP + '/WebSGEP/ConsEstoqueProdutoNovo/FiltroConsultaEstoque';
const USUARIO_SGEP = process.env.SGEP_USER || 'Beto';
const SENHA_SGEP = process.env.SGEP_PASS || '@464';


const CLIENTES_CADASTRADOS = {
  CA: {
    nomeCompleto: 'C A DOS PASSOS',
    filtroSgep: 'C A DOS PASSOS',
    senhaAcesso: '123456'
  }
};

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.post('/baixar', async (req, res) => {
  const sigla = req.body.sigla;
  const senha = req.body.senha;
  const cliente = CLIENTES_CADASTRADOS[sigla];

  if (!cliente || cliente.senhaAcesso !== senha) {
    return res.status(401).json({ mensagem: 'Sigla ou senha incorreta.' });
  }

  let browser;
  let etapa = 'Iniciando navegador';
  try {
    browser = await chromium.launch({
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--no-zygote'
      ]
    });
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();
    page.setDefaultTimeout(25000);

    etapa = 'Acessando tela de login';
    await page.goto(URL_LOGIN_SGEP, { waitUntil: 'domcontentloaded' });
    await page.locator('input[type="text"], input[name*="Login"], input[name*="Usuario"]').first().fill(USUARIO_SGEP);
    await page.locator('input[type="password"]').first().fill(SENHA_SGEP);
    await page.locator('button[type="submit"], input[type="submit"], button:has-text("Entrar"), button:has-text("Acessar")').first().click();
    await page.waitForLoadState('networkidle');

    etapa = 'Abrindo pagina de estoque';
    await page.goto(URL_ESTOQUE_SGEP, { waitUntil: 'networkidle' });

    etapa = 'Selecionando cliente no filtro';
    const seletorCliente = page.locator('.bootstrap-select').filter({ hasText: /Nada selecionado|Cliente/i }).first();
    if (await seletorCliente.count() > 0) {
      await seletorCliente.locator('button').first().click();
    } else {
      await page.locator('.bootstrap-select button').first().click();
    }

    const caixaBusca = page.locator('.dropdown-menu.open .bs-searchbox input, .dropdown-menu.show .bs-searchbox input, .bs-searchbox input').first();
    if (await caixaBusca.isVisible()) {
      await caixaBusca.fill(cliente.filtroSgep);
    } else {
      await page.keyboard.type(cliente.filtroSgep, { delay: 50 });
    }
    await page.waitForTimeout(800);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(800);

    etapa = 'Clicando em Pesquisar';
    const btnPesquisar = page.locator('button:has-text("Pesquisar"), button:has-text("Consultar"), button:has-text("Filtrar"), input[value*="Pesquisar"]').first();
    if (await btnPesquisar.count() > 0 && await btnPesquisar.isVisible()) {
      await btnPesquisar.click();
      await page.waitForLoadState('networkidle');
    }

    etapa = 'Baixando arquivo Excel';
    const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
    await page.locator('button:has-text("Excel"), a:has-text("Excel"), .buttons-excel, [title*="Excel"], button:has-text("Exportar"), a:has-text("Exportar")').first().click();
    const download = await downloadPromise;

    etapa = 'Formatando planilha';
    const caminhoBruto = path.join(__dirname, 'temp_' + sigla + '_' + Date.now() + '.xlsx');
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
    const nomeFinal = 'Estoque_' + cliente.nomeCompleto.replace(/\s+/g, '_') + '_' + dia + '-' + mes + '-' + ano + '_' + hora + 'h' + minuto + '.xlsx';

    res.setHeader('Content-Disposition', 'attachment; filename="' + nomeFinal + '"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(bufferFormatado);

  } catch (erro) {
    if (browser) await browser.close();
    console.error('Erro na etapa [' + etapa + ']:', erro);
    res.status(500).json({ mensagem: 'Falha na etapa "' + etapa + '": ' + erro.message });
  }
});

function extrairEspecie(nomeProduto) {
  const texto = String(nomeProduto || '').trim().toUpperCase();
  if (!texto) return 'OUTROS';
  const partes = texto.split(/[\.\s]+/).filter(Boolean);
  const primeira = partes.at(0) || '';
  const segunda = partes.at(1) || '';

  if (primeira === 'P' || primeira === 'PESC' || primeira === 'PESCADA') {
    const mapa = { AMARELA: 'PESCADA AMARELA', BANANA: 'PESCADA BANANA', BRANCA: 'PESCADA BRANCA', CURUCA: 'PESCADA CURUCA', GO: 'PESCADA GO' };
    return segunda ? (mapa[segunda] || 'PESCADA ' + segunda) : 'PESCADA';
  }
  if (primeira === 'CAM' || primeira === 'CAMARAO') {
    return (segunda && segunda !== 'N') ? 'CAMARAO ' + segunda : 'CAMARAO ROSA';
  }
  return primeira || 'OUTROS';
}

async function formatarPlanilha(caminhoArquivo, nomeCliente, sigla, agora) {
  const wbOrig = new ExcelJS.Workbook();
  await wbOrig.xlsx.load(fs.readFileSync(caminhoArquivo));
  const sheetOrig = wbOrig.worksheets.at(0);

  const dados = [];
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
      produto: produto,
      secundaria: row.getCell(4).value,
      primaria: row.getCell(5).value,
      vol: parseNum(row.getCell(6).value),
      saldoKgs: parseNum(row.getCell(7).value),
      especie: extrairEspecie(produto)
    });
  });

  const grupos = {};
  dados.forEach((item) => {
    if (!grupos[item.especie]) grupos[item.especie] = [];
    grupos[item.especie].push(item);
  });

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Estoque Formatado');

  ws.getCell('A1').value = nomeCliente;
  ws.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FF1F3864' } };

  const dataAtual = agora.toLocaleDateString('pt-BR');
  ws.getCell('A2').value = 'Dentro da validade + cliente ' + sigla + '. • Data-base ' + dataAtual;
  ws.getCell('A2').font = { italic: true, size: 9, color: { argb: 'FF595959' } };

  const headerRow = ws.getRow(5);
  headerRow.values = ['Cod.Ext', 'Cód. NF', 'Produto', 'Secundária', 'Primária', 'Vol', 'Saldo Kgs'];
  headerRow.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
  });

  let rowIdx = 6;
  let totVol = 0;
  let totKgs = 0;

  Object.keys(grupos).forEach((especie) => {
    const itens = grupos[especie];
    let subVol = 0;
    let subKgs = 0;

    itens.forEach((i) => {
      const r = ws.getRow(rowIdx++);
      r.values = [i.codExt, i.codNf, i.produto, i.secundaria, i.primaria, i.vol, i.saldoKgs];
      r.getCell(6).numFmt = '#,##0';
      r.getCell(7).numFmt = '#,##0.00';
      subVol += i.vol;
      subKgs += i.saldoKgs;
    });

    const subRow = ws.getRow(rowIdx++);
    subRow.getCell(3).value = 'Subtotal ' + especie;
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
  });

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
app.listen(PORT, '0.0.0.0', () => console.log('Portal rodando na porta ' + PORT));
