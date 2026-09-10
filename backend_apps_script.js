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

    if (action === 'updatePagamento') {
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

  let unidade = getField(v, 'Unidade').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '-');

  db.appendRow([
    id,
    new Date().toISOString(),
    'cliente',
    'AGUARDANDO CONFIRMACAO',
    getField(v, 'Chassi'),
    getField(v, 'CPF'),
    Number(getField(v, 'Entrada'))  || 0,
    Number(getField(v, 'Valor'))    || 0,
    getField(v, 'Nome do Cliente'),
    getField(v, 'Vendedor'),
    getField(v, 'Cidade'),
    unidade,
    getField(v, 'Modelo'),
    'Loja',
    getField(v, 'Forma de Pagamento'),
    getField(v, 'Fechou Rastreador?'),
    getField(v, 'Motivo Não Fechou'),
    '', '', // Nota Fiscal / Obs (preenchidos pelo Financeiro)
    new Date().toISOString()
  ]);

  Logger.log('✅ Cliente registrado: ' + getField(v, 'Nome do Cliente'));
}

/**
 * Formulário de Revendedores — compra em lote
 * Campos: Nome | Cidade | AZ1 | AZ125 | AZ160 | Observação
 * Cria um registro por modelo com quantidade > 0.
 */
function processRevendedorForm(e) {
  const v    = e.namedValues || {};
  const db   = getSheetDB();
  const nome = getField(v, 'Nome');
  const cidade = getField(v, 'Cidade');
  const obs  = getField(v, 'Observação');

  const lote = [
    { modelo: 'AZ1',   qtd: Number(getField(v, 'AZ1'))   || 0 },
    { modelo: 'AZ125', qtd: Number(getField(v, 'AZ125')) || 0 },
    { modelo: 'AZ160', qtd: Number(getField(v, 'AZ160')) || 0 },
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
      new Date().toISOString()
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
      'FechouRastreador','MotivoNaoFechou','NotaFiscalUrl','Observacao','AtualizadoEm'
    ]);
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
      atualizadoEm:   obj.AtualizadoEm
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

// ============================================================
//  6. HELPERS
// ============================================================
function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getField(namedValues, fieldName) {
  const val = namedValues[fieldName];
  return val && val[0] ? String(val[0]).trim() : '';
}
