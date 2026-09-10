function renderPagSidebar(activePage) {
  const p = PAGAMENTOS_CONFIG.PAGES;
  const session = PagAuth.getSession();
  
  let navItems = '';
  if (activePage === 'financeiro') {
    navItems = `
      <div class="sb-section-lbl">CONFIRMAÇÃO DE PAGAMENTOS</div>
      <a href="${p.financeiro}" class="nav-item nav-item-link active"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg><span>Financeiro</span></a>
    `;
  } else if (activePage === 'gerentes') {
    navItems = `
      <div class="sb-section-lbl">CONFIRMAÇÃO DE PAGAMENTOS</div>
      <a href="${p.gerentes}" class="nav-item nav-item-link active"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 21h18M5 21V9l7-5 7 5v12"/></svg><span>Gerentes</span></a>
    `;
  }

  return `
    <aside class="sidebar" id="sidebar">
      <div class="sb-brand">
        <a href="${p.index}" class="nav-item-link" style="display:flex;align-items:center;gap:12px;text-decoration:none">
          <div class="mark"><svg viewBox="0 0 100 100" fill="none"><circle cx="50" cy="50" r="46" stroke="#FF6B00" stroke-width="7"/><path d="M30 65 L50 30 L58 44 L44 44 L58 68" stroke="#FF6B00" stroke-width="7" stroke-linecap="round" stroke-linejoin="round" fill="none"/></svg></div>
          <div class="txt"><div class="name">JS GRUPO</div><div class="sub">PAINEL DE VENDAS</div></div>
        </a>
        <button class="sb-close" id="sbClose" type="button"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 6l12 12M18 6L6 18"/></svg></button>
      </div>
      <nav class="sb-nav">
        ${navItems}
      </nav>
      <div class="sb-foot">
        ${session ? `<div class="sb-meta">${esc(session.label)} · <button type="button" id="btnPagLogout" class="btn-secondary" style="padding:4px 10px;font-size:11px;margin-top:6px;width:100%">Sair e Trocar de Setor</button></div>` : ''}
        <div class="sb-meta">Liberação / Vendas</div>
      </div>
    </aside>`;
}

function bindPagMobileNav() {
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sbOverlay');
  document.getElementById('hamburger')?.addEventListener('click', () => { sidebar?.classList.add('open'); overlay?.classList.add('show'); });
  document.getElementById('sbClose')?.addEventListener('click', () => { sidebar?.classList.remove('open'); overlay?.classList.remove('show'); });
  overlay?.addEventListener('click', () => { sidebar?.classList.remove('open'); overlay?.classList.remove('show'); });
}
