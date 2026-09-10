/**
 * Configuração — Confirmação de Pagamentos (JS GRUPO)
 * Use a mesma URL do Apps Script ou uma implantação dedicada com aba Vendas.
 */
const PAGAMENTOS_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbyLx7r1MwS75zVK6GkO1Kff4eawR8kVhJpvJIaLWDgMkGD3UFceSRQ_vmbDlM73FGvC/exec',

  AUTH: {
    analise: { password: 'ana@2026', label: 'Análise Comercial' },
    financeiro: { password: 'fin@2026', label: 'Financeiro' },
    gerentes: { password: 'ger@2026', label: 'Gerência Comercial' }
  },

  STATUS: [
    { id: 'AGUARDANDO CONFIRMACAO', label: 'AGUARDANDO CONFIRMAÇÃO', group: 'nao_iniciado' },
    { id: 'FALTANDO INFORMACAO', label: 'FALTANDO INFORMAÇÃO', group: 'nao_iniciado' },
    { id: 'PAGAMENTO OK', label: 'PAGAMENTO OK', group: 'feito' },
    { id: 'CONCLUIDO', label: 'CONCLUÍDO', group: 'fechado' }
  ],

  UNIDADES: [
    { id: 'mamanguape', nome: 'Mamanguape' },
    { id: 'guarabira', nome: 'Guarabira' },
    { id: 'patos-centro', nome: 'Patos (Centro)' },
    { id: 'patos-cd', nome: 'Patos (CD)' },
    { id: 'natal', nome: 'Natal' },
    { id: 'joao-pessoa', nome: 'João Pessoa (Matriz)' },
    { id: 'parnamirim', nome: 'Parnamirim' },
    { id: 'sape', nome: 'Sapé' },
    { id: 'pombal', nome: 'Pombal' },
    { id: 'sao-bento', nome: 'São Bento' }
  ],

  PAGES: { index: 'index.html', analise: 'analise.html', financeiro: 'financeiro.html', gerentes: 'gerentes.html' }
};

function getStatusMeta(statusId) {
  return PAGAMENTOS_CONFIG.STATUS.find(s => s.id === statusId) || PAGAMENTOS_CONFIG.STATUS[0];
}

function getUnidadeMeta(unidadeId) {
  return PAGAMENTOS_CONFIG.UNIDADES.find(u => u.id === unidadeId) || { id: unidadeId, nome: unidadeId };
}

function formatCurrency(value) {
  const num = Number(String(value).replace(/[^\d,.-]/g, '').replace(',', '.')) || 0;
  return num.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function formatDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function formatDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
}

function esc(s) {
  return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
