const PAG_AUTH_KEY = 'jsgrupo_pag_session';

const PagAuth = {
  async login(role, usuario, password) {
    const cfg = PAGAMENTOS_CONFIG.AUTH[role];
    if (!cfg) return false;
    const base = PAGAMENTOS_CONFIG.API_URL;
    const url = `${base}${base.includes('?') ? '&' : '?'}action=authLogin&_=${Date.now()}`;
    const response = await fetch(url, {
      method: 'POST', mode: 'cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'authLogin', role, usuario, senha: password })
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const result = await response.json();
    if (result.success === false) throw new Error(result.error || 'Acesso negado.');
    sessionStorage.setItem(PAG_AUTH_KEY, JSON.stringify({ role, ...result.data, label: cfg.label, senha: password, at: Date.now() }));
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
          <span>Usuário</span>
          <input type="text" id="pagLoginUser" required autocomplete="username" />
        </label>
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
  overlay.querySelector('#pagLoginForm').addEventListener('submit', async e => {
    e.preventDefault();
    const usuario = document.getElementById('pagLoginUser')?.value || role;
    try {
      if (await PagAuth.login(role, usuario, document.getElementById('pagLoginPwd').value)) {
      overlay.remove();
      onSuccess?.();
      }
    } catch (error) {
      document.getElementById('pagLoginErr').classList.remove('hidden');
      document.getElementById('pagLoginErr').textContent = error.message;
    }
  });
}

function bindPagLogout() {
  document.getElementById('btnPagLogout')?.addEventListener('click', () => {
    PagAuth.logout();
    window.location.href = PAGAMENTOS_CONFIG.PAGES.index;
  });
}
