function pagStatusBadge(statusId) {
  const meta = getStatusMeta(statusId);
  const cls = meta.group === 'feito' || meta.group === 'fechado' ? 'ok' : meta.id === 'FALTANDO INFORMACAO' ? 'warn' : 'low';
  return `<span class="status-badge ${cls}">${esc(meta.label)}</span>`;
}

function pagChip(text, type) {
  return `<span class="chip ${type || 'blue'}">${esc(text)}</span>`;
}

function renderPagDetail(venda, { editable, onSave } = {}) {
  const u = getUnidadeMeta(venda.unidade);
  const overlay = document.createElement('div');
  overlay.className = 'pag-detail-overlay';
  overlay.innerHTML = `
    <div class="pag-detail-panel card">
      <div class="pag-detail-head">
        <div><h2>${esc(venda.cliente)}</h2><p class="row-sub">${venda.formType === 'revendedor' ? 'Revendedor' : 'Cliente'} · ${formatDateTime(venda.timestamp)}</p></div>
        <button class="pag-close" type="button">&times;</button>
      </div>
      <div class="pag-detail-body">
        <div style="margin-bottom:14px;display:flex;gap:8px;flex-wrap:wrap">${pagStatusBadge(venda.status)} ${pagChip(u.nome,'gold')}</div>
        ${pagField('T','CHASSI',venda.chassi)}${pagField('T','CPF',venda.cpf)}
        ${pagField('$','Entrada',formatCurrency(venda.entrada))}${pagField('$','Valor',formatCurrency(venda.valor))}
        ${pagField('T','Cliente',venda.cliente)}${pagField('T','Vendedor',pagChip(venda.vendedor,'blue'))}
        ${pagField('T','Modelo',pagChip(venda.modelo,'gold'))}${pagField('T','Canal',pagChip(venda.canal,'teal'))}
        ${pagField('T','Forma Pagamento',pagChip(venda.formaPagamento,'green'))}
        ${pagField('T','Rastreador',pagChip(venda.fechouRastreador, venda.fechouRastreador === 'Sim' ? 'green' : 'red'))}
        ${pagField('T','Porque Não Fechou',venda.motivoNaoFechou)}
        ${pagField('📄','Nota Fiscal', venda.notaFiscalUrl ? `<a href="${esc(venda.notaFiscalUrl)}" target="_blank" rel="noopener">Abrir NF</a>` : '<span style="color:var(--amber)">Sem anexo</span>')}
        ${venda.observacao ? pagField('💬','Observação',venda.observacao) : ''}
      </div>
      ${editable ? `
        <div class="pag-detail-foot">
          <label class="pag-field"><span>Alterar Status</span>
            <select id="pagStatusSel">${PAGAMENTOS_CONFIG.STATUS.map(s => `<option value="${s.id}" ${s.id===venda.status?'selected':''}>${s.label}</option>`).join('')}</select>
          </label>
          <label class="pag-field"><span>Observação</span><textarea id="pagObsIn" placeholder="Motivo...">${esc(venda.observacao||'')}</textarea></label>
          <button class="btn-primary" id="pagSaveBtn" type="button">Confirmar alteração</button>
        </div>` : ''}
    </div>`;
  document.body.appendChild(overlay);
  const close = () => overlay.remove();
  overlay.querySelector('.pag-close').addEventListener('click', close);
  overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
  if (editable) {
    overlay.querySelector('#pagSaveBtn').addEventListener('click', async () => {
      const btn = overlay.querySelector('#pagSaveBtn');
      btn.disabled = true; btn.textContent = 'Salvando...';
      try {
        await onSave(venda.id, overlay.querySelector('#pagStatusSel').value, overlay.querySelector('#pagObsIn').value);
        close();
      } catch (err) { alert(err.message); btn.disabled = false; btn.textContent = 'Confirmar alteração'; }
    });
  }
}

function pagField(icon, label, value) {
  return `<div class="pag-field-row"><span class="pag-ficon">${icon}</span><span class="pag-flbl">${label}</span><span class="pag-fval">${value||'—'}</span></div>`;
}

function renderPagTable(vendas, onRowClick) {
  if (!vendas.length) return '<div class="empty-state">Nenhum registro encontrado.</div>';
  const groups = {};
  PAGAMENTOS_CONFIG.STATUS.forEach(s => { groups[s.id] = []; });
  vendas.forEach(v => (groups[v.status] || groups['AGUARDANDO CONFIRMACAO']).push(v));

  let html = '';
  PAGAMENTOS_CONFIG.STATUS.forEach(st => {
    const items = groups[st.id];
    if (!items.length) return;
    html += `<div class="pag-status-group"><div class="pag-group-head">${pagStatusBadge(st.id)}<span class="chip gold">${items.length}</span></div>
      <div class="tbl-wrap"><table class="pag-table"><thead><tr>
        <th>Cliente</th><th>CPF</th><th class="num">Valor</th><th>Pagamento</th><th>Vendedor</th><th>Data</th><th>Rastreador</th>
      </tr></thead><tbody>${items.map(v => `
        <tr data-id="${v.id}"><td><b>${esc(v.cliente)}</b><div class="row-sub">${esc(v.modelo)}</div></td>
        <td>${esc(v.cpf)}</td><td class="num"><b>${formatCurrency(v.valor)}</b></td>
        <td>${pagChip(v.formaPagamento,'green')}</td><td>${pagChip(v.vendedor,'blue')}</td>
        <td>${formatDate(v.timestamp)}</td><td>${pagChip(v.fechouRastreador, v.fechouRastreador==='Sim'?'green':'red')}</td></tr>`).join('')}
      </tbody></table></div></div>`;
  });
  const wrap = document.createElement('div');
  wrap.innerHTML = html;
  wrap.querySelectorAll('tr[data-id]').forEach(r => r.addEventListener('click', () => onRowClick(r.dataset.id)));
  return wrap;
}
