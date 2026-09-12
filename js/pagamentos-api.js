const PagApi = {
  async request(action, payload = {}) {
    const base = PAGAMENTOS_CONFIG.API_URL;
    if (!base) return PagDemo.handle(action, payload);

    const session = typeof PagAuth !== 'undefined' ? PagAuth.getSession() : null;
    const securedActions = [
      'listPagamentos', 'updatePagamento',
      'vendasList', 'vendasCreate', 'vendasRequestEdit', 'vendasApproveEdit',
      'colaboradoresList', 'colaboradoresCreate', 'colaboradoresSetStatus'
    ];
    const authPayload = securedActions.includes(action) && session
      ? { usuario: session.usuario, senha: session.senha }
      : {};

    const url = `${base}${base.includes('?') ? '&' : '?'}action=${action}&_=${Date.now()}`;
    const res = await fetch(url, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action, ...authPayload, ...payload })
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (data.success === false) throw new Error(data.error || 'Erro na API');
    return data;
  },

  listVendas(filters = {}) { return this.request('listPagamentos', filters); },
  updateStatus(id, status, observacao = '') { return this.request('updatePagamento', { id, status, observacao }); }
};

const PagDemo = (() => {
  let vendas = [
    { id:'V001', timestamp:'2026-09-03T17:43:23-03:00', formType:'cliente', status:'AGUARDANDO CONFIRMACAO', chassi:'LJCJCKLS5TS014061', cpf:'038.705.674-22', entrada:5997, valor:5997, cliente:'DAILSON GONÇALVES DA SILVA', vendedor:'SUE ELLEN', cidade:'João Pessoa (Matriz)', unidade:'joao-pessoa', modelo:'Az160', canal:'Loja', formaPagamento:'BV', fechouRastreador:'Não', motivoNaoFechou:'N', notaFiscalUrl:'#', observacao:'', atualizadoEm:'2026-09-03T17:43:23-03:00' },
    { id:'V002', timestamp:'2026-09-04T10:15:00-03:00', formType:'revendedor', status:'PAGAMENTO OK', chassi:'LJCJCKLS5TS014062', cpf:'123.456.789-00', entrada:8500, valor:8500, cliente:'MARIA SILVA SANTOS', vendedor:'AM', cidade:'Mamanguape', unidade:'mamanguape', modelo:'Az125', canal:'Revenda', formaPagamento:'MOTRIX', fechouRastreador:'Sim', motivoNaoFechou:'', notaFiscalUrl:'#', observacao:'Confirmado', atualizadoEm:'2026-09-04T14:30:00-03:00' },
    { id:'V003', timestamp:'2026-09-08T09:00:00-03:00', formType:'cliente', status:'FALTANDO INFORMACAO', chassi:'', cpf:'987.654.321-00', entrada:4500, valor:4500, cliente:'JOÃO PEREIRA LIMA', vendedor:'AG', cidade:'Natal', unidade:'natal', modelo:'Az160', canal:'Loja', formaPagamento:'Pan', fechouRastreador:'Não', motivoNaoFechou:'NAO QUIS', notaFiscalUrl:'', observacao:'Falta chassi', atualizadoEm:'2026-09-08T09:00:00-03:00' },
    { id:'V004', timestamp:'2026-09-09T16:20:00-03:00', formType:'cliente', status:'CONCLUIDO', chassi:'LJCJCKLS5TS014099', cpf:'111.222.333-44', entrada:7200, valor:7200, cliente:'ANA COSTA OLIVEIRA', vendedor:'SUE ELLEN', cidade:'Guarabira', unidade:'guarabira', modelo:'Az125', canal:'Loja', formaPagamento:'BV', fechouRastreador:'Sim', motivoNaoFechou:'', notaFiscalUrl:'#', observacao:'', atualizadoEm:'2026-09-10T08:00:00-03:00' }
  ];

  function filterList(f) {
    let r = [...vendas];
    if (f.unidade) r = r.filter(v => v.unidade === f.unidade);
    if (f.status) r = r.filter(v => v.status === f.status);
    if (f.formType) r = r.filter(v => v.formType === f.formType);
    if (f.periodo) {
      const now = new Date();
      r = r.filter(v => {
        const d = new Date(v.timestamp);
        if (f.periodo === 'dia') return d.toDateString() === now.toDateString();
        if (f.periodo === 'semana') { const w = new Date(now); w.setDate(now.getDate()-7); return d >= w; }
        if (f.periodo === 'mes') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        return true;
      });
    }
    return r.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
  }

  return {
    handle(action, payload) {
      if (action === 'listPagamentos') return Promise.resolve({ success:true, data: filterList(payload) });
      if (action === 'updatePagamento') {
        const i = vendas.findIndex(v => v.id === payload.id);
        if (i >= 0) { vendas[i].status = payload.status; vendas[i].observacao = payload.observacao || ''; vendas[i].atualizadoEm = new Date().toISOString(); }
        return Promise.resolve({ success:true, data: vendas[i] });
      }
      return Promise.resolve({ success:false, error:'Ação não disponível em demo' });
    }
  };
})();
