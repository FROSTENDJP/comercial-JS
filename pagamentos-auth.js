const PAG_AUTH_KEY = 'jsgrupo_pag_session';

const PagAuth = {
  login(role, password) {
    const cfg = PAGAMENTOS_CONFIG.AUTH[role];
    if (!cfg || cfg.password !== password) return false;
    sessionStorage.setItem(PAG_AUTH_KEY, JSON.stringify({ role, label: cfg.label, at: Date.now() }));
    return true;
  },

  logout() {
    sessionStorage.removeItem(PAG_AUTH_KEY);
  },

  getSession() {
    try { return JSON.parse(sessionStorage.getItem(PAG_AUTH_KEY)); } catch { return null; }
  },

  isLoggedIn(role) {
    const s = this.getSession();
    return s && s.role === role;
  }
};

function showPagLogin(role, onSuccess) {
  const cfg = PAGAMENTOS_CONFIG.AUTH[role];
  const overlay = document.createElement('div');
  overlay.className = 'pag-modal-overlay';
  overlay.innerHTML = `
    <div class="pag-modal card">
      <div class="pag-modal-head">
        <h2>Acesso ${esc(cfg.label)}</h2>
        <p>Digite a senha do setor para continuar</p>
      </div>
      <form id="pagLoginForm">
        <label class="pag-field">
          <span>Senha</span>
          <input type="password" id="pagLoginPwd" required autofocus placeholder="••••••••" />
        </label>
        <p class="pag-error hidden" id="pagLoginErr">Senha incorreta</p>
        <div class="pag-modal-actions">
          <a href="${PAGAMENTOS_CONFIG.PAGES.index}" class="btn-secondary">Voltar</a>
          <button type="submit" class="btn-primary">Entrar</button>
        </div>
      </form>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector('#pagLoginForm').addEventListener('submit', e => {
    e.preventDefault();
    if (PagAuth.login(role, document.getElementById('pagLoginPwd').value)) {
      overlay.remove();
      onSuccess?.();
    } else {
      document.getElementById('pagLoginErr').classList.remove('hidden');
    }
  });
}

function bindPagLogout() {
  document.getElementById('btnPagLogout')?.addEventListener('click', () => {
    PagAuth.logout();
    window.location.href = PAGAMENTOS_CONFIG.PAGES.index;
  });
}
