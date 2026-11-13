/**
 * Conecta o painel de vendas ao servidor (e, por ele, à planilha Excel).
 * Coloque no FINAL do index.html, depois do <script> principal:
 *     <script src="sync.js"></script>
 * Ele substitui os dados fixos e o armazenamento local pelo Excel.
 */
(() => {
  try { localStorage.removeItem('vendas-novas'); } catch (e) {}

  const status = document.createElement('p');
  status.className = 'sub';
  status.id = 'syncStatus';
  document.querySelector('.sub').after(status);
  const dizer = t => { status.textContent = t; };

  const paraLinha = s => ({
    id: s.id, reg: s.reg, uf: s.uf, prod: s.prod,
    dia: s.data.slice(8, 10) + '/' + s.data.slice(5, 7), ano: +s.data.slice(0, 4),
    qtd: s.qtd, vu: s.vu, fr: s.fr, tot: s.tot, novo: true // novo:true mostra o botão Remover em todas
  });

  async function req(metodo, url, corpo) {
    const r = await fetch(url, {
      method: metodo,
      headers: { 'Content-Type': 'application/json' },
      body: corpo ? JSON.stringify(corpo) : undefined
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.erro || r.statusText);
    return j;
  }

  async function carregar() {
    const dados = await req('GET', '/api/vendas');
    const validos = dados.filter(s => s.vu != null && s.fr != null && s.data);
    RAW.length = 0;
    RAW.push(...validos.map(paraLinha));
    PRICE && validos.forEach(s => (PRICE[s.prod] = s.vu));
    refreshSelects();
    render();
    const ign = dados.length - validos.length;
    dizer(`Conectado ao Excel: ${validos.length} vendas` + (ign ? ` (${ign} linhas ignoradas por valor ou data inválidos)` : '') + '.');
  }

  // adicionar venda -> grava no Excel
  $('nForm').onsubmit = async e => {
    e.preventDefault();
    try {
      const s = await req('POST', '/api/vendas', {
        reg: $('nReg').value, uf: $('nUf').value, prod: $('nProd').value, data: $('nData').value,
        qtd: $('nQtd').value, vu: $('nVu').value, fr: $('nFr').value
      });
      $('nMsg').textContent = `Venda ${s.id} gravada no Excel.`;
      $('nProd').value = ''; $('nVu').value = ''; $('nQtd').value = 1; nTotal();
      await carregar();
    } catch (err) { $('nMsg').textContent = 'Erro: ' + err.message; }
  };

  // remover venda -> apaga do Excel
  window.delNovo = async id => {
    try { await req('DELETE', '/api/vendas/' + id); await carregar(); $('nMsg').textContent = `Venda ${id} removida do Excel.`; }
    catch (err) { $('nMsg').textContent = 'Erro: ' + err.message; }
  };

  // tempo real: qualquer mudança (neste painel, em outro navegador ou direto no Excel) recarrega
  const es = new EventSource('/api/stream');
  es.onmessage = () => carregar().catch(err => dizer('Erro ao ler o Excel: ' + err.message));
  es.onerror = () => dizer('Reconectando ao servidor…');

  carregar().catch(err => dizer('Sem conexão com o servidor: ' + err.message));
})();