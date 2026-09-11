/**
 * ============================================================
 *  SCRIPT UNIFICADO — GRUPO JS
 * ============================================================
 *  Este arquivo SUBSTITUI o seu Code.gs atual no Apps Script.
 *  Ele mantém 100% do doGet() original que alimenta o painel
 *  de Análise Comercial e ADICIONA:
 *    - doPost()              → atualiza status de pagamentos
 *    - onFormSubmit()        → captura respostas dos Forms
 *    - listPagamentos()      → lê a aba BD_Pagamentos
 *    - processClienteForm()  → salva venda de cliente
 *    - processRevendedorForm() → salva compra de revendedor (lote)
 * ============================================================
 */

const SPREADSHEET_ID = '1RNoEacmCl7X3F6j39nm0Ra58AsNh8F1i8phnCQWtqGA';
const SHEET_BD       = 'BD_Pagamentos';

// ============================================================
//  1. doGet — ANÁLISE COMERCIAL (código original preservado)
//             + roteamento para listPagamentos via ?action=
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  // Rota de leitura de pagamentos (usada pelo painel Financeiro/Gerentes)
  if (action === 'listPagamentos') {
    return responseJson({ success: true, data: getPagamentos() });
  }

  return responseJson(buildAnalysisPayload());

  // ---- Código original do doGet (Análise Comercial) ----
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function readKeyValueSheet(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    const obj = {};
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      data.forEach(row => {
        const key = String(row[0]).trim();
        if (key) obj[key] = Number(row[1]) || row[1];
      });
    }
    return obj;
  }

  function readModelQtySheet(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    const obj = {};
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      data.forEach(row => {
        const key = String(row[0]).trim().toLowerCase();
        if (key) obj[key] = Number(row[1]) || 0;
      });
    }
    return obj;
  }

  function readModelBreakdownSheet(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    const result = {
      az1: 0, az125: 0, az160: 0, total: 0,
      extra: { az1: 0, az125: 0, az160: 0, total: 0 }
    };
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      data.forEach(row => {
        const modelo = String(row[0]).trim().toLowerCase();
        const qtd    = Number(row[1]) || 0;
        const extra  = Number(row[2]) || 0;
        if (result.hasOwnProperty(modelo)) {
          result[modelo]       = qtd;
          result.extra[modelo] = extra;
        }
      });
      result.total       = result.az1 + result.az125 + result.az160;
      result.extra.total = result.extra.az1 + result.extra.az125 + result.extra.az160;
    }
    return result;
  }

  function readOrigemSheet(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    const result = { marcelo: 0, lojas: 0 };
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      data.forEach(row => {
        const origem = String(row[0]).trim().toLowerCase();
        const qtd    = Number(row[1]) || 0;
        if (origem === 'marcelo')                result.marcelo = qtd;
        if (origem === 'lojas' || origem === 'loja') result.lojas = qtd;
      });
    }
    return result;
  }

  const totals = readKeyValueSheet('totals');

  const revHeaderRaw = readModelBreakdownSheet('revHeader');
  const revHeader = {
    az1: revHeaderRaw.az1, az125: revHeaderRaw.az125, az160: revHeaderRaw.az160,
    total: revHeaderRaw.total,
    previsao: Number(totals.previsao) || 0,
    meta:     Number(totals.meta)     || 0
  };

  const vendasInternas = readOrigemSheet('vendasInternasOrigem');

  const vendasInternasRaw = readModelBreakdownSheet('vendasInternas');
  const vendasInternasPorModelo = {
    az1: vendasInternasRaw.az1, az125: vendasInternasRaw.az125,
    az160: vendasInternasRaw.az160, total: vendasInternasRaw.total,
    meta: vendasInternasRaw.extra
  };

  const motosCompradas = readModelQtySheet('motosCompradas');

  const filiaisSheet = ss.getSheetByName('filiais');
  let filiais = [];
  if (filiaisSheet && filiaisSheet.getLastRow() > 1) {
    const data = filiaisSheet.getRange(2, 1, filiaisSheet.getLastRow() - 1, 7).getValues();
    filiais = data.map(row => ({
      nome:  String(row[0]).trim(),
      uf:    String(row[1]).trim().toUpperCase(),
      az1:   Number(row[2]) || 0,
      az125: Number(row[3]) || 0,
      az160: Number(row[4]) || 0,
      total: Number(row[5]) || 0,
      meta:  Number(row[6]) || 0,
      vendedores: []
    }));
  }

  const vendSheet = ss.getSheetByName('vendedores');
  let rankVendedores = [];
  if (vendSheet && vendSheet.getLastRow() > 1) {
    const data = vendSheet.getRange(2, 1, vendSheet.getLastRow() - 1, 7).getValues();
    data.forEach(row => {
      const nome  = String(row[0]).trim();
      const loja  = String(row[1]).trim();
      const uf    = String(row[2]).trim().toUpperCase();
      const az1   = Number(row[3]) || 0;
      const az125 = Number(row[4]) || 0;
      const az160 = Number(row[5]) || 0;
      const meta  = Number(row[6]) || 0;
      const total = az1 + az125 + az160;

      rankVendedores.push({ nome, loja, uf, vendas: total, qtd: total, meta });

      const lojaNorm = loja.trim().toLowerCase();
      const filial   = filiais.find(f => f.nome.trim().toLowerCase() === lojaNorm);
      if (filial) {
        filial.vendedores.push({ nome, az1, az125, az160 });
      } else {
        Logger.log('⚠️ Vendedor "' + nome + '" sem filial para "' + loja + '"');
      }
    });
  }

  const revSheet = ss.getSheetByName('revendedores');
  let revendedores = [];
  if (revSheet && revSheet.getLastRow() > 1) {
    const data = revSheet.getRange(2, 1, revSheet.getLastRow() - 1, 7).getValues();
    revendedores = data.map((row, idx) => {
      const az1   = Number(row[3]) || 0;
      const az125 = Number(row[4]) || 0;
      const az160 = Number(row[5]) || 0;
      const cota  = Number(row[2]) || 0;
      return {
        pos: idx + 1,
        nome:  String(row[0]).trim(),
        uf:    String(row[1]).trim().toUpperCase(),
        cota, meta: cota, az1, az125, az160,
        total: az1 + az125 + az160,
        status: String(row[6])
      };
    });
  }

  const rankLojasSheet = ss.getSheetByName('rankLojas');
  let rankLojas = [];
  if (rankLojasSheet && rankLojasSheet.getLastRow() > 1) {
    const data = rankLojasSheet.getRange(2, 1, rankLojasSheet.getLastRow() - 1, 6).getValues();
    rankLojas = data.map(row => ({
      posicao:    Number(row[0]) || 0,
      nome:       String(row[1]).trim(),
      loja:       String(row[1]).trim(),
      uf:         String(row[2]).trim().toUpperCase(),
      vendas:     Number(row[3]) || 0,
      qtd:        Number(row[3]) || 0,
      meta:       Number(row[4]) || 0,
      atingimento: Number(row[5]) || 0
    }));
  }

  const payload = {
    totals, filiais, rankLojas, rankVendedores,
    revHeader, revendedores, motosCompradas,
    vendasInternas, vendasInternasPorModelo
  };

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  2. doPost — ATUALIZAÇÃO DE PAGAMENTOS
//              (chamado pelo painel Financeiro ao confirmar)
// ============================================================
function doPost(e) {
  try {
    let params = {};
    if (e.postData && e.postData.contents) {
      try { params = JSON.parse(e.postData.contents); } catch (_) {}
    }
    // Permite também via query string
    if (e.parameter) Object.assign(params, e.parameter);

    const action = params.action || '';

    if (action === 'vendasLogin') {
      // Login legado do setor Vendas — mantido só por retrocompatibilidade.
      // Agora delega para authenticatePortal(), a mesma função usada por
      // Análise/Financeiro/Gerentes via action=authLogin, para garantir que
      // só usuários com perfil UNIDADE ou GERENCIA entrem em Vendas.
      return responseJson({ success: true, data: authenticatePortal(params.usuario, params.senha, 'vendas') });
    }

    if (action === 'authLogin') {
      return responseJson({ success: true, data: authenticatePortal(params.usuario, params.senha, params.role) });
    }

    if (action === 'vendasList') {
      return responseJson({ success: true, data: listVendasScoped(params) });
    }

    if (action === 'vendasCreate') {
      return responseJson({ success: true, data: createVendaVendas(params) });
    }

    if (action === 'vendasRequestEdit') {
      return responseJson({ success: true, data: requestVendaEdit(params) });
    }

    if (action === 'vendasApproveEdit') {
      return responseJson({ success: true, data: approveVendaEdit(params) });
    }

    if (action === 'listPagamentos') {
      requireVendasUser(params, ['FINANCEIRO', 'GERENCIA']);
      return responseJson({ success: true, data: getPagamentos() });
    }

    if (action === 'updatePagamento') {
      requireVendasUser(params, ['FINANCEIRO']);
      const result = updatePagamento(params.id, params.status, params.observacao);
      return responseJson({ success: true, data: result });
    }

    return responseJson({ success: false, error: 'Ação desconhecida: ' + action });
  } catch (err) {
    return responseJson({ success: false, error: err.toString() });
  }
}

// ============================================================
//  3. GATILHO DE FORMULÁRIO — roteador principal
// ============================================================
function onFormSubmit(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();
  if (sheetName === 'Form_Clientes')    processClienteForm(e);
  else if (sheetName === 'Form_Revendedores') processRevendedorForm(e);
}

// ============================================================
//  4. PROCESSADORES DOS DOIS FORMULÁRIOS
// ============================================================

/** Formulário de Unidades Próprias — detalhe individual por venda */
function processClienteForm(e) {
  const v  = e.namedValues || {};
  const db = getSheetDB();
  const id = 'C' + new Date().getTime();

  let unidade = getNamedField(v, ['Unidade', 'UNIDADE']).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '-');

  db.appendRow([
    id,
    new Date().toISOString(),
    'cliente',
    'AGUARDANDO CONFIRMACAO',
    getNamedField(v, ['Chassi']),
    getNamedField(v, ['CPF']),
    numberField(getNamedField(v, ['VALOR ENTRADA', 'Valor Entrada', 'Entrada', 'ENTRADA'])),
    numberField(getNamedField(v, ['VALOR FINANCIADO', 'Valor Financiado', 'Valor', 'VALOR'])),
    getNamedField(v, ['Nome do Cliente', 'Cliente']),
    getNamedField(v, ['Vendedor']),
    getNamedField(v, ['Cidade', 'CIDADE']),
    unidade,
    getNamedField(v, ['Modelo']),
    'Loja',
    getNamedField(v, ['Forma de Pagamento', 'FORMA DE PAGAMENTO']),
    getNamedField(v, ['Fechou Rastreador?', 'Fechou o Rastreador ?']),
    getNamedField(v, ['Motivo Não Fechou', 'MOTIVO N FECHOU']),
    '', '', // Nota Fiscal / Obs (preenchidos pelo Financeiro)
    new Date().toISOString(),
    1,
    numberField(getNamedField(v, ['VALOR FINANCIADO', 'Valor Financiado', 'Valor', 'VALOR'])),
    numberField(getNamedField(v, ['VALOR ENTRADA', 'Valor Entrada', 'Entrada', 'ENTRADA'])),
    getNamedField(v, ['TIPO ENTRADA', 'Tipo Entrada', 'FORMA DE PAGAMENTO', 'Forma de Pagamento']),
    getNamedField(v, ['BANCO FINANCIAMENTO', 'BANCO DO FINANCIAMENTO', 'Banco do Financiamento'])
  ]);

  Logger.log('✅ Cliente registrado: ' + getNamedField(v, ['Nome do Cliente', 'Cliente']));
}

/**
 * Formulário de Revendedores — compra em lote
 * Campos: Nome | Cidade | AZ1 | AZ125 | AZ160 | Observação
 * Cria um registro por modelo com quantidade > 0.
 */
function processRevendedorForm(e) {
  const v    = e.namedValues || {};
  const db   = getSheetDB();
  const nome = getNamedField(v, ['Nome', 'NOME', 'Revendedor']);
  const cidade = getNamedField(v, ['Cidade', 'CIDADE']);
  const obs  = getNamedField(v, ['Observação', 'OBSERVAÇÃO']);

  const lote = [
    { modelo: 'AZ1',   qtd: numberField(getNamedField(v, ['AZ1'])) },
    { modelo: 'AZ125', qtd: numberField(getNamedField(v, ['AZ125'])) },
    { modelo: 'AZ160', qtd: numberField(getNamedField(v, ['AZ160'])) },
  ];

  lote.forEach(m => {
    if (m.qtd <= 0) return;
    const id = 'R' + new Date().getTime() + '_' + m.modelo;
    db.appendRow([
      id, new Date().toISOString(), 'revendedor', 'AGUARDANDO CONFIRMACAO',
      '', '', 0, 0,       // Chassi / CPF / Entrada / Valor
      nome, '', cidade, '', // Cliente / Vendedor / Cidade / Unidade
      m.modelo, 'Revenda',
      '', '', '',          // Forma Pgto / Rastreador / Motivo
      '', obs,             // Nota Fiscal / Observação
      new Date().toISOString(),
      m.qtd
    ]);
    Logger.log('✅ Revendedor: ' + nome + ' | ' + m.modelo + ' x' + m.qtd);
  });
}

// ============================================================
//  5. CRUD de BD_Pagamentos
// ============================================================

function getSheetDB() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_BD);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BD);
    sheet.appendRow([
      'ID','Timestamp','FormType','Status','Chassi','CPF','Entrada','Valor',
      'Cliente','Vendedor','Cidade','Unidade','Modelo','Canal','FormaPagamento',
      'FechouRastreador','MotivoNaoFechou','NotaFiscalUrl','Observacao','AtualizadoEm','Quantidade',
      'ValorFinanciado','ValorEntrada','TipoEntrada','BancoFinanciamento',
      'CriadoPor','UnidadeOrigem','EditadoPor','EditadoEm','StatusEdicao','MotivoEdicao','SolicitadoPor','AprovadoPor'
    ]);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    ['Quantidade', 'ValorFinanciado', 'ValorEntrada', 'TipoEntrada', 'BancoFinanciamento',
      'CriadoPor','UnidadeOrigem','EditadoPor','EditadoEm','StatusEdicao','MotivoEdicao','SolicitadoPor','AprovadoPor'].forEach(header => {
      if (headers.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        headers.push(header);
      }
    });
  }
  return sheet;
}

function getPagamentos() {
  const sheet = getSheetDB();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return {
      id:             obj.ID,
      timestamp:      obj.Timestamp,
      formType:       obj.FormType,
      status:         obj.Status,
      chassi:         obj.Chassi,
      cpf:            obj.CPF,
      entrada:        obj.Entrada,
      valor:          obj.Valor,
      cliente:        obj.Cliente,
      vendedor:       obj.Vendedor,
      cidade:         obj.Cidade,
      unidade:        obj.Unidade,
      modelo:         obj.Modelo,
      canal:          obj.Canal,
      formaPagamento: obj.FormaPagamento,
      fechouRastreador: obj.FechouRastreador,
      motivoNaoFechou: obj.MotivoNaoFechou,
      notaFiscalUrl:  obj.NotaFiscalUrl,
      observacao:     obj.Observacao,
      atualizadoEm: obj.AtualizadoEm,
      quantidade: obj.Quantidade || 1,
      valorFinanciado: obj.ValorFinanciado || obj.Valor || 0,
      valorEntrada: obj.ValorEntrada || obj.Entrada || 0,
      tipoEntrada: obj.TipoEntrada || '',
      bancoFinanciamento: obj.BancoFinanciamento || '',
      criadoPor: obj.CriadoPor || '',
      unidadeOrigem: obj.UnidadeOrigem || '',
      editadoPor: obj.EditadoPor || '',
      editadoEm: obj.EditadoEm || '',
      statusEdicao: obj.StatusEdicao || '',
      motivoEdicao: obj.MotivoEdicao || '',
      solicitadoPor: obj.SolicitadoPor || '',
      aprovadoPor: obj.AprovadoPor || ''
    };
  });
}

function updatePagamento(id, status, observacao) {
  const sheet = getSheetDB();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 4).setValue(status);      // Status
      sheet.getRange(rowNum, 19).setValue(observacao); // Observacao
      sheet.getRange(rowNum, 20).setValue(new Date().toISOString()); // AtualizadoEm
      return { id, status, observacao };
    }
  }
  throw new Error('ID não encontrado: ' + id);
}

function normalizeUserProfile(profile) {
  const value = String(profile || '').trim().toUpperCase();
  if (value === 'COMERCIAL') return 'UNIDADE';
  return value;
}

function getUsuarios() {
  return readSheetRows('Usuarios');
}

function authenticateVendas(usuario, senha) {
  const wanted = String(usuario || '').trim().toUpperCase();
  const users = getUsuarios();
  const row = users.find(item => String(rowField(item, ['usuario'])).trim().toUpperCase() === wanted
    && String(rowField(item, ['senha'])).trim() === String(senha || '').trim()
    && String(rowField(item, ['ativo'])).trim().toLowerCase() === 'ativo');
  if (!row) throw new Error('Usuário ou senha inválidos.');
  return {
    usuario: String(rowField(row, ['usuario'])).trim(),
    perfil: normalizeUserProfile(rowField(row, ['perfil'])),
    unidade: String(rowField(row, ['unidade'])).trim()
  };
}

function authenticatePortal(usuario, senha, role) {
  const expectedProfiles = {
    analise: ['ANALISE'],
    financeiro: ['FINANCEIRO'],
    gerentes: ['GERENCIA'],
    vendas: ['UNIDADE', 'GERENCIA']
  };
  const profiles = expectedProfiles[String(role || '').toLowerCase()];
  if (!profiles) throw new Error('Setor inválido.');
  const user = authenticateVendas(usuario, senha);
  if (profiles.indexOf(user.perfil) === -1) throw new Error('Usuário sem acesso a este setor.');
  return user;
}

function requireVendasUser(params, profiles) {
  const user = authenticateVendas(params.usuario, params.senha);
  const normalizedProfile = normalizeUserProfile(user.perfil);
  if (profiles.indexOf(normalizedProfile) === -1) throw new Error('Perfil sem permissão para esta ação.');
  return { ...user, perfil: normalizedProfile };
}

function listVendasScoped(params) {
  const user = requireVendasUser(params, ['UNIDADE', 'GERENCIA', 'FINANCEIRO']);
  let rows = getPagamentos();
  if (user.perfil === 'UNIDADE') rows = rows.filter(row => normalizeUnit(row.unidade) === normalizeUnit(user.unidade));
  return rows;
}

function appendMappedRow(sheet, values) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(header => values[header] === undefined ? '' : values[header]));
}

function createVendaVendas(params) {
  const user = requireVendasUser(params, ['UNIDADE']);
  const unidade = user.unidade || String(params.unidade || '').trim();
  if (!unidade) throw new Error('Unidade obrigatória.');
  const sheet = getSheetDB();
  const financiado = numberField(params.valorFinanciado);
  const entrada = numberField(params.valorEntrada);
  const id = 'V' + new Date().getTime();
  appendMappedRow(sheet, {
    ID: id, Timestamp: new Date().toISOString(), FormType: 'cliente', Status: 'AGUARDANDO CONFIRMACAO',
    Chassi: String(params.chassi || '').trim(), CPF: String(params.cpf || '').trim(),
    Entrada: entrada, Valor: financiado, Cliente: String(params.cliente || '').trim(),
    Vendedor: String(params.vendedor || '').trim(), Cidade: String(params.cidade || '').trim(),
    Unidade: normalizeUnit(unidade), Modelo: String(params.modelo || '').trim(), Canal: 'Loja',
    FormaPagamento: String(params.formaPagamento || '').trim(), FechouRastreador: String(params.fechouRastreador || '').trim(),
    MotivoNaoFechou: String(params.motivoNaoFechou || '').trim(), NotaFiscalUrl: '', Observacao: '',
    AtualizadoEm: new Date().toISOString(), Quantidade: 1, ValorFinanciado: financiado,
    ValorEntrada: entrada, TipoEntrada: String(params.formaPagamento || '').trim(),
    BancoFinanciamento: String(params.bancoFinanciamento || '').trim(), CriadoPor: user.usuario,
    UnidadeOrigem: unidade, EditadoPor: '', EditadoEm: '', StatusEdicao: '', MotivoEdicao: '',
    SolicitadoPor: '', AprovadoPor: ''
  });
  return getPagamentos().find(row => String(row.id) === id);
}

function requestVendaEdit(params) {
  const user = requireVendasUser(params, ['UNIDADE']);
  const sheet = getSheetDB();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('ID');
  const statusIndex = headers.indexOf('StatusEdicao');
  const reasonIndex = headers.indexOf('MotivoEdicao');
  const requesterIndex = headers.indexOf('SolicitadoPor');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) !== String(params.id)) continue;
    if (normalizeUnit(data[i][headers.indexOf('Unidade')]) !== normalizeUnit(user.unidade)) throw new Error('Venda fora da unidade.');
    sheet.getRange(i + 1, statusIndex + 1).setValue('SOLICITACAO DE EDICAO');
    sheet.getRange(i + 1, reasonIndex + 1).setValue(String(params.motivo || '').trim());
    sheet.getRange(i + 1, requesterIndex + 1).setValue(user.usuario);
    return { id: params.id, statusEdicao: 'SOLICITACAO DE EDICAO' };
  }
  throw new Error('Venda não encontrada.');
}

function approveVendaEdit(params) {
  const user = requireVendasUser(params, ['GERENCIA']);
  const sheet = getSheetDB();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('ID');
  const statusIndex = headers.indexOf('StatusEdicao');
  const approvedIndex = headers.indexOf('AprovadoPor');
  const editIndex = headers.indexOf('EditadoPor');
  const editDateIndex = headers.indexOf('EditadoEm');
  const changes = typeof params.changes === 'string' ? JSON.parse(params.changes || '{}') : (params.changes || {});
  const allowed = ['Cliente','CPF','Chassi','Modelo','ValorFinanciado','ValorEntrada','Vendedor','Cidade','Unidade','FormaPagamento','BancoFinanciamento'];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) !== String(params.id)) continue;
    if (String(data[i][statusIndex]).trim() !== 'SOLICITACAO DE EDICAO') throw new Error('Esta venda não possui solicitação pendente.');
    allowed.forEach(field => { if (changes[field] !== undefined && headers.indexOf(field) >= 0) sheet.getRange(i + 1, headers.indexOf(field) + 1).setValue(changes[field]); });
    sheet.getRange(i + 1, statusIndex + 1).setValue('EDICAO APROVADA');
    sheet.getRange(i + 1, approvedIndex + 1).setValue(user.usuario);
    sheet.getRange(i + 1, editIndex + 1).setValue(user.usuario);
    sheet.getRange(i + 1, editDateIndex + 1).setValue(new Date().toISOString());
    return { id: params.id, statusEdicao: 'EDICAO APROVADA' };
  }
  throw new Error('Venda não encontrada.');
}

// ============================================================
//  6. HELPERS
// ============================================================
function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function getSheetByNameLoose(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const targetName = String(sheetName || '').trim();
  if (!targetName) return null;

  const exactSheet = spreadsheet.getSheetByName(targetName);
  if (exactSheet) return exactSheet;

  const targetKey = normalizeHeader(targetName);
  return spreadsheet.getSheets().find(sheet => normalizeHeader(sheet.getName()) === targetKey) || null;
}

function readSheetRows(sheetName) {
  const sheet = getSheetByNameLoose(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader);
  return values.slice(1).map(row => {
    const item = {};
    headers.forEach((header, index) => { if (header) item[header] = row[index]; });
    return item;
  });
}

function rowField(row, names) {
  for (const name of names) {
    const value = row[normalizeHeader(name)];
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '';
}

function numberField(value) {
  if (typeof value === 'number') return value;
  const text = String(value || '').replace(/R\$\s?/g, '').replace(/\./g, '').replace(',', '.');
  return Number(text.replace(/[^\d.-]/g, '')) || 0;
}

function normalizeUnit(value) {
  return String(value || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '-');
}

function inferUF(value) {
  const text = String(value || '').toLowerCase();
  return /natal|parnamirim|rn|rio grande do norte/.test(text) ? 'RN' : 'PB';
}

function buildAnalysisPayload() {
  const clients = readSheetRows('Form_Clientes').filter(row =>
    rowField(row, ['Nome do Cliente', 'Cliente']) || rowField(row, ['Modelo']));
  const resellerRows = readSheetRows('Form_Revendedores').filter(row =>
    rowField(row, ['Nome', 'NOME', 'Revendedor']));

  const filiaisMap = {};
  const vendedoresMap = {};
  const lojasMap = {};
  const clientTotals = { az1: 0, az125: 0, az160: 0 };
  const metasRows = readSheetRows('Metas');

  metasRows.forEach(row => {
    const tipo = String(rowField(row, ['tipo'])).trim().toLowerCase();
    const nome = String(rowField(row, ['nome'])).trim();
    const cidade = String(rowField(row, ['cidade'])).trim();
    const uf = String(rowField(row, ['uf'])).trim().toUpperCase() || inferUF(cidade);
    const metaAz1 = numberField(rowField(row, ['metaAz1', 'meta AZ1']));
    const metaAz125 = numberField(rowField(row, ['metaAz125', 'meta AZ125']));
    const metaAz160 = numberField(rowField(row, ['metaAz160', 'meta AZ160']));
    const metaTotal = numberField(rowField(row, ['meta'])) || metaAz1 + metaAz125 + metaAz160;
    const status = String(rowField(row, ['status'])).trim() || 'ativo';

    if (tipo === 'unidade') {
      const unidadeId = normalizeUnit(nome || cidade);
      if (!unidadeId) return;
      filiaisMap[unidadeId] = {
        nome: nome || cidade, cidade, uf, az1: 0, az125: 0, az160: 0,
        meta: metaTotal, metaAz1, metaAz125, metaAz160, status, vendedores: {}
      };
    }
  });

  clients.forEach(row => {
    const unidade = rowField(row, ['Unidade', 'Loja']) || rowField(row, ['Cidade']);
    const unidadeId = normalizeUnit(unidade);
    const uf = inferUF(unidade + ' ' + rowField(row, ['Cidade']));
    const modelo = String(rowField(row, ['Modelo'])).trim().toLowerCase();
    const key = modelo === 'az125' ? 'az125' : modelo === 'az160' ? 'az160' : modelo === 'az1' ? 'az1' : '';
    if (!key) return;

    clientTotals[key]++;
    if (!filiaisMap[unidadeId]) filiaisMap[unidadeId] = { nome: unidade || 'Sem unidade', uf, az1: 0, az125: 0, az160: 0, meta: 0, metaAz1: 0, metaAz125: 0, metaAz160: 0, status: 'ativo', vendedores: {} };
    filiaisMap[unidadeId][key]++;

    const vendedor = String(rowField(row, ['Vendedor'])).trim();
    if (vendedor) {
      const sellerKey = vendedor.toLowerCase();
      if (!vendedoresMap[sellerKey]) vendedoresMap[sellerKey] = { nome: vendedor, loja: unidade, uf, az1: 0, az125: 0, az160: 0 };
      vendedoresMap[sellerKey][key]++;
      if (!filiaisMap[unidadeId].vendedores[sellerKey]) filiaisMap[unidadeId].vendedores[sellerKey] = vendedoresMap[sellerKey];
    }

    const lojaKey = unidade || 'Sem unidade';
    if (!lojasMap[lojaKey]) lojasMap[lojaKey] = { nome: lojaKey, loja: lojaKey, uf, qtd: 0 };
    lojasMap[lojaKey].qtd++;
  });

  const resellerMap = {};
  const resellerTotals = { az1: 0, az125: 0, az160: 0 };
  metasRows.forEach(row => {
    const tipo = String(rowField(row, ['tipo'])).trim().toLowerCase();
    if (tipo !== 'revendedor') return;
    const nome = String(rowField(row, ['nome'])).trim();
    const cidade = String(rowField(row, ['cidade'])).trim();
    const key = (nome + '|' + cidade).toLowerCase();
    const metaAz1 = numberField(rowField(row, ['metaAz1', 'meta AZ1']));
    const metaAz125 = numberField(rowField(row, ['metaAz125', 'meta AZ125']));
    const metaAz160 = numberField(rowField(row, ['metaAz160', 'meta AZ160']));
    const meta = numberField(rowField(row, ['meta'])) || metaAz1 + metaAz125 + metaAz160;
    resellerMap[key] = { nome, cidade, uf: String(rowField(row, ['uf'])).trim().toUpperCase() || inferUF(cidade), cota: meta, metaAz1, metaAz125, metaAz160, az1: 0, az125: 0, az160: 0, status: String(rowField(row, ['status'])).trim() || 'ativo' };
  });
  resellerRows.forEach(row => {
    const nome = String(rowField(row, ['Nome', 'Revendedor'])).trim();
    const cidade = String(rowField(row, ['Cidade'])).trim();
    const key = (nome + '|' + cidade).toLowerCase();
    if (!resellerMap[key]) resellerMap[key] = { nome: nome || cidade, cidade, uf: inferUF(cidade), cota: 0, metaAz1: 0, metaAz125: 0, metaAz160: 0, az1: 0, az125: 0, az160: 0, status: 'ativo' };
    ['az1', 'az125', 'az160'].forEach(modelo => {
      const qtd = numberField(rowField(row, [modelo]));
      resellerMap[key][modelo] += qtd;
      resellerTotals[modelo] += qtd;
    });
  });

  const clientTotal = clientTotals.az1 + clientTotals.az125 + clientTotals.az160;
  const resellerTotal = resellerTotals.az1 + resellerTotals.az125 + resellerTotals.az160;
  const filiais = Object.values(filiaisMap).map(f => ({
    nome: f.nome, uf: f.uf, az1: f.az1, az125: f.az125, az160: f.az160,
    total: f.az1 + f.az125 + f.az160, meta: f.meta || 0,
    vendedores: Object.values(f.vendedores).map(v => ({ nome: v.nome, az1: v.az1, az125: v.az125, az160: v.az160 }))
  }));
  const rankVendedores = Object.values(vendedoresMap).map(v => ({
    nome: v.nome, loja: v.loja, uf: v.uf, az1: v.az1, az125: v.az125, az160: v.az160,
    vendas: v.az1 + v.az125 + v.az160, qtd: v.az1 + v.az125 + v.az160,
    meta: 0
  }));
  const revendedores = Object.values(resellerMap).map((r, index) => ({
    pos: index + 1, nome: r.nome, uf: r.uf,
    cota: r.cota || 0,
    meta: r.cota || 0,
    az1: r.az1, az125: r.az125, az160: r.az160, total: r.az1 + r.az125 + r.az160, status: r.status
  }));

  return {
    totals: {
      az1: clientTotals.az1, az125: clientTotals.az125, az160: clientTotals.az160,
      totalGeral: clientTotal, pb: filiais.filter(f => f.uf === 'PB').reduce((sum, f) => sum + f.total, 0),
      rn: filiais.filter(f => f.uf === 'RN').reduce((sum, f) => sum + f.total, 0),
      revendas: resellerTotal, lojas: clientTotal,
      previsao: 0, meta: filiais.reduce((sum, f) => sum + (f.meta || 0), 0)
    },
    filiais,
    rankLojas: Object.values(lojasMap).map((r, index) => ({ posicao: index + 1, ...r, vendas: r.qtd })),
    rankVendedores,
    revHeader: { ...resellerTotals, total: resellerTotal, previsao: 0, meta: revendedores.reduce((sum, r) => sum + (r.meta || 0), 0) },
    revendedores,
    motosCompradas: resellerTotals,
    vendasInternas: { marcelo: 0, lojas: clientTotal },
    vendasInternasPorModelo: { ...clientTotals, total: clientTotal, meta: { az1: 0, az125: 0, az160: 0, total: 0 } }
  };
}

function getField(namedValues, fieldName) {
  const val = namedValues[fieldName];
  return val && val[0] ? String(val[0]).trim() : '';
}

function getNamedField(namedValues, names) {
  const normalized = Object.keys(namedValues || {}).reduce((result, key) => {
    result[normalizeHeader(key)] = namedValues[key];
    return result;
  }, {});
  for (const name of names) {
    const value = normalized[normalizeHeader(name)];
    if (value && value[0] !== undefined) return String(value[0]).trim();
  }
  return '';
}
