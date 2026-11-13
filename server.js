'use strict';
/**
 * Servidor que usa o arquivo Excel como banco de dados.
 *  - API REST  : GET/POST /api/vendas, PUT/DELETE /api/vendas/:id
 *  - Tempo real: GET /api/stream (Server-Sent Events)
 *  - Vigia o arquivo: se alguém salvar a planilha no Excel, o painel atualiza sozinho.
 *
 * Uso:  npm install && node server.js
 * Opções (variáveis de ambiente): EXCEL_FILE=./vendas.xlsx  PORT=3000  USE_POLLING=1
 */
const express = require('express');
const ExcelJS = require('exceljs');
const chokidar = require('chokidar');
const fs = require('fs');
const path = require('path');

const FILE = path.resolve(process.env.EXCEL_FILE || path.join(__dirname, 'vendas.xlsx'));
const PORT = process.env.PORT || 3000;
const FRETE = { 'Centro-Oeste': 15, Nordeste: 20, Norte: 25, Sul: 15, Sudeste: 10 };
const HEADERS = ['ID', 'Região', 'estado', 'Produtos', 'data', 'Quantidade', 'Valor Unit.', 'Frete', 'TOTAL'];
const NCOLS = HEADERS.length;
const MOEDA = '"R$" #,##0.00';

/* ---------- fila: só uma operação por vez no arquivo ---------- */
let fila = Promise.resolve();
const serial = fn => { const p = fila.then(fn); fila = p.catch(() => {}); return p; };
let ultimaGravacao = 0;

/* ---------- utilidades do Excel ---------- */
async function abrir() {
  const wb = new ExcelJS.Workbook();
  if (!fs.existsSync(FILE)) {
    wb.addWorksheet('Vendas').addRow(HEADERS);
    await wb.xlsx.writeFile(FILE);
  }
  await wb.xlsx.readFile(FILE);
  return { wb, ws: wb.worksheets[0] }; // usa a primeira aba
}

async function salvar(wb) {
  ultimaGravacao = Date.now();
  await wb.xlsx.writeFile(FILE);
  ultimaGravacao = Date.now();
}

function valor(cell) {
  const v = cell.value;
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if ('result' in v && !(v.result && v.result.error)) return v.result; // fórmula: usa o resultado
    return null; // erro (#N/D etc.)
  }
  return v;
}

function dataISO(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(v || ''));
  return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
}
const num = v => (typeof v === 'number' ? v : null);

function lerVendas(ws) {
  const out = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const id = num(valor(ws.getCell(r, 1)));
    if (!id) continue;
    const s = {
      id, linha: r,
      reg: String(valor(ws.getCell(r, 2)) || '').replace(/^Região\s+/i, ''),
      uf: String(valor(ws.getCell(r, 3)) || ''),
      prod: String(valor(ws.getCell(r, 4)) || ''),
      data: dataISO(valor(ws.getCell(r, 5))),
      qtd: num(valor(ws.getCell(r, 6))),
      vu: num(valor(ws.getCell(r, 7))),
      fr: num(valor(ws.getCell(r, 8)))
    };
    s.tot = s.vu != null && s.fr != null && s.qtd != null ? s.qtd * (s.vu + s.fr) : null;
    out.push(s);
  }
  return out;
}

function gravarLinha(ws, r, s) {
  const set = (c, x, fmt) => { const cell = ws.getCell(r, c); cell.value = x; if (fmt) cell.numFmt = fmt; };
  set(1, s.id);
  set(2, 'Região ' + s.reg);
  set(3, s.uf);
  set(4, s.prod);
  set(5, new Date(s.data + 'T00:00:00Z'), 'dd/mm/yyyy');
  set(6, s.qtd);
  set(7, s.vu, MOEDA);
  set(8, s.fr, MOEDA);
  set(9, { formula: `F${r}*(G${r}+H${r})`, result: s.qtd * (s.vu + s.fr) }, MOEDA);
}

// copia o conteúdo de uma célula para a linha de cima, ajustando referências da própria linha
function moverValor(v, de, para) {
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    if (v.formula) {
      const re = new RegExp('(?<![A-Za-z$])(\\$?[A-Z]{1,2})' + de + '(?!\\d)', 'g');
      return { formula: v.formula.replace(re, '$1' + para), result: v.result };
    }
    if (v.sharedFormula) return v.result;
  }
  return v;
}

const erro = (status, msg) => Object.assign(new Error(msg), { status });

function validar(b) {
  const reg = String(b.reg || '').trim();
  if (!FRETE[reg]) throw erro(400, 'Região inválida. Use: ' + Object.keys(FRETE).join(', '));
  const uf = String(b.uf || '').trim().toUpperCase();
  if (!/^[A-Z]{2}$/.test(uf)) throw erro(400, 'UF inválida.');
  const prod = String(b.prod || '').trim();
  if (!prod) throw erro(400, 'Informe o produto.');
  const data = String(b.data || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || isNaN(Date.parse(data))) throw erro(400, 'Data inválida (use AAAA-MM-DD).');
  const qtd = Number(b.qtd);
  if (!Number.isInteger(qtd) || qtd < 1) throw erro(400, 'Quantidade inválida.');
  if (b.vu == null || b.vu === '' || !(Number(b.vu) >= 0)) throw erro(400, 'Valor unitário inválido.');
  const vu = Number(b.vu);
  const fr = b.fr == null || b.fr === '' ? FRETE[reg] : Number(b.fr);
  if (!(fr >= 0)) throw erro(400, 'Frete inválido.');
  return { reg, uf, prod, data, qtd, vu, fr };
}

/* ---------- operações (o "banco de dados") ---------- */
const listar = () => serial(async () => lerVendas((await abrir()).ws));

const criar = body => serial(async () => {
  const dados = validar(body);
  const { wb, ws } = await abrir();
  const vendas = lerVendas(ws);
  const r = vendas.reduce((m, s) => Math.max(m, s.linha), 1) + 1;
  for (let c = 1; c <= NCOLS; c++) {
    if (valor(ws.getCell(r, c)) != null) {
      throw erro(409, `Há conteúdo na linha ${r}, logo abaixo da tabela (ex.: a linha TOTAL). Mova-o para fora das colunas A:I.`);
    }
  }
  const s = { ...dados, id: vendas.reduce((m, x) => Math.max(m, x.id), 1000) + 1 };
  gravarLinha(ws, r, s);
  await salvar(wb);
  return { ...s, linha: r, tot: s.qtd * (s.vu + s.fr) };
});

const atualizar = (id, body) => serial(async () => {
  const { wb, ws } = await abrir();
  const atual = lerVendas(ws).find(s => s.id === id);
  if (!atual) throw erro(404, 'Venda não encontrada.');
  const novo = validar({ ...atual, ...body });
  const colunas = { reg: 2, uf: 3, prod: 4, data: 5, qtd: 6, vu: 7, fr: 8 };
  for (const [k, c] of Object.entries(colunas)) {
    if (novo[k] === atual[k]) continue; // só mexe no que mudou
    const cell = ws.getCell(atual.linha, c);
    cell.value = k === 'reg' ? 'Região ' + novo.reg : k === 'data' ? new Date(novo.data + 'T00:00:00Z') : novo[k];
  }
  const ci = ws.getCell(atual.linha, 9);
  if (!(ci.value && ci.value.formula)) ci.value = { formula: `F${atual.linha}*(G${atual.linha}+H${atual.linha})` };
  ci.numFmt = MOEDA;
  await salvar(wb);
  return { ...novo, id, linha: atual.linha, tot: novo.qtd * (novo.vu + novo.fr) };
});

const remover = id => serial(async () => {
  const { wb, ws } = await abrir();
  const vendas = lerVendas(ws);
  const i = vendas.findIndex(s => s.id === id);
  if (i < 0) throw erro(404, 'Venda não encontrada.');
  // sobe as linhas seguintes (somente colunas A:I, para não mexer nas tabelas ao lado)
  for (let k = i + 1; k < vendas.length; k++) {
    const de = vendas[k].linha, para = vendas[k - 1].linha;
    for (let c = 1; c <= NCOLS; c++) {
      const src = ws.getCell(de, c), dst = ws.getCell(para, c);
      dst.value = moverValor(src.value, de, para);
      dst.numFmt = src.numFmt;
    }
  }
  const ultima = vendas[vendas.length - 1].linha;
  for (let c = 1; c <= NCOLS; c++) ws.getCell(ultima, c).value = null;
  await salvar(wb);
  return { removido: id };
});

/* ---------- tempo real (Server-Sent Events) ---------- */
const clientes = new Set();
const avisar = () => clientes.forEach(res => res.write(`data: ${Date.now()}\n\n`));
setInterval(() => clientes.forEach(res => res.write(': ping\n\n')), 25000);

// vigia o arquivo: mudanças feitas direto no Excel também atualizam o painel
let timer;
chokidar.watch(FILE, {
  ignoreInitial: true,
  usePolling: !!process.env.USE_POLLING, // use em pasta de rede / OneDrive, se necessário
  awaitWriteFinish: { stabilityThreshold: 400, pollInterval: 100 }
}).on('all', ev => {
  if (ev !== 'change' && ev !== 'add') return;
  if (Date.now() - ultimaGravacao < 2000) return; // foi o próprio servidor que gravou
  clearTimeout(timer);
  timer = setTimeout(avisar, 200);
});

/* ---------- API ---------- */
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const rota = fn => (req, res) => fn(req, res).catch(err => {
  const ocupado = ['EBUSY', 'EPERM', 'EACCES'].includes(err.code);
  const status = ocupado ? 423 : err.status || 500;
  if (status === 500) console.error(err);
  res.status(status).json({
    erro: ocupado ? 'A planilha está aberta no Excel. Feche-a e tente de novo.' : err.message
  });
});

app.get('/api/vendas', rota(async (req, res) => res.json(await listar())));
app.post('/api/vendas', rota(async (req, res) => { const s = await criar(req.body); avisar(); res.status(201).json(s); }));
app.put('/api/vendas/:id', rota(async (req, res) => { const s = await atualizar(+req.params.id, req.body); avisar(); res.json(s); }));
app.delete('/api/vendas/:id', rota(async (req, res) => { const r = await remover(+req.params.id); avisar(); res.json(r); }));
app.get('/api/stream', (req, res) => {
  res.set({ 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' });
  res.flushHeaders();
  res.write('retry: 2000\n\n');
  clientes.add(res);
  req.on('close', () => clientes.delete(res));
});

app.listen(PORT, () => console.log(`Painel em http://localhost:${PORT}  |  Excel: ${FILE}`));