/**
 * TigerOffers — Backend em Google Sheets pra ferramenta de reativação de afiliados.
 *
 * COMO INSTALAR:
 * 1. Crie uma planilha nova no Google Sheets (pode ser em branco).
 * 2. Extensões → Apps Script.
 * 3. Apague o conteúdo de Code.gs e cole este arquivo inteiro.
 * 4. Rode a função `setup` uma vez (menu Executar → selecione "setup" → Executar).
 *    Isso cria as 3 abas necessárias com cabeçalho. Vai pedir autorização — aceite.
 * 5. Implantar → Nova implantação → tipo "App da Web".
 *    - Executar como: Eu (seu e-mail)
 *    - Quem tem acesso: Qualquer pessoa
 * 6. Copie a URL que aparece (algo tipo https://script.google.com/macros/s/XXXX/exec)
 * 7. Cole essa URL na constante APPS_SCRIPT_URL no topo do index.html.
 *
 * ABAS CRIADAS:
 * - AddedAffiliates: afiliados inseridos manualmente pela ferramenta
 * - ContactLog: registro de "último contato" (oferta, resumo, data) por afiliado
 * - ImportUpdates: última leitura de cada import diário, por afiliado (upsert)
 * - AffiliateMeta: nome no dash da empresa, tipo de tráfego, CPA e usuário/e-mail BuyGoods, por afiliado
 * - RevenueDaily: faturamento real por dia (relatório allProducts, aba "Detalhado por Funil"), 1 linha por
 *   afiliado com um JSON acumulado dia a dia na coluna dados_json
 * - LifetimeStats: totais vitalícios (desde o início da operação, sem data) do export "Master Affiliates"
 *   da BuyGoods — importação única, substitui a aba inteira a cada vez
 */

const SHEET_ADDED = 'AddedAffiliates';
const SHEET_CONTACT = 'ContactLog';
const SHEET_IMPORT = 'ImportUpdates';
const SHEET_META = 'AffiliateMeta';
const SHEET_REVENUE = 'RevenueDaily';
const SHEET_LIFETIME = 'LifetimeStats';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_ADDED)) {
    const sh = ss.insertSheet(SHEET_ADDED);
    sh.appendRow(['nome_norm', 'nome', 'telefone', 'telegram', 'produto_planilha', 'observacao', 'criado_em']);
  }
  if (!ss.getSheetByName(SHEET_CONTACT)) {
    const sh = ss.insertSheet(SHEET_CONTACT);
    sh.appendRow(['nome_norm', 'nome', 'oferta', 'resumo', 'data', 'atualizado_em']);
  }
  if (!ss.getSheetByName(SHEET_IMPORT)) {
    const sh = ss.insertSheet(SHEET_IMPORT);
    sh.appendRow(['nome_norm', 'nome', 'volume_ago', 'volume_set', 'net_ago', 'net_set',
                  'tem_dado_custo', 'ultima_venda', 'dias_sem_vender', 'tier',
                  'confirmado', 'total_pedidos', 'volume_total', 'produtos', 'streak_diario', 'atualizado_em']);
  }
  if (!ss.getSheetByName(SHEET_META)) {
    const sh = ss.insertSheet(SHEET_META);
    sh.appendRow(['nome_norm', 'nome', 'nome_dash', 'trafego', 'cpa', 'buygoods_user', 'atualizado_em']);
  }
  if (!ss.getSheetByName(SHEET_REVENUE)) {
    const sh = ss.insertSheet(SHEET_REVENUE);
    sh.appendRow(['nome_norm', 'nome', 'dados_json', 'atualizado_em']);
  }
  if (!ss.getSheetByName(SHEET_LIFETIME)) {
    const sh = ss.insertSheet(SHEET_LIFETIME);
    sh.appendRow(['nome_norm', 'nome', 'produtos', 'orders', 'gross', 'net', 'commissions', 'refunds', 'atualizado_em']);
  }
  // remove a aba padrão "Sheet1"/"Página1" se estiver vazia
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  Logger.log('Setup concluído. Abas: ' + ss.getSheets().map(s => s.getName()).join(', '));
}

function normName_(s) {
  if (!s) return '';
  return s.toString().trim().toLowerCase()
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim();
}

function sheetToObjects_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return {};
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    if (!obj.nome_norm) continue;
    out[obj.nome_norm] = obj; // último ganha se houver duplicata
  }
  return out;
}

function upsertRow_(sheetName, keyField, keyValue, rowObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  let headers = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  // se rowObj trouxer campos que a planilha ainda não tem como coluna, cria a coluna (schema auto-migra)
  const missing = Object.keys(rowObj).filter(k => headers.indexOf(k) === -1);
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  const data = sh.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf(keyField)] === keyValue) { foundRow = i + 1; break; }
  }
  const rowValues = headers.map(h => (h in rowObj) ? rowObj[h] : '');
  if (foundRow > 0) {
    sh.getRange(foundRow, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sh.appendRow(rowValues);
  }
}

function bulkReplaceSheet_(sheetName, headers, rows) {
  // substitui a aba inteira de uma vez (1 chamada de rede) — usado por importações grandes/únicas
  // onde upsert linha-a-linha (que relê a planilha inteira a cada chamada) seria lento demais
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear();
  const allRows = [headers].concat(rows);
  if (allRows.length > 0) {
    sh.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  }
}

function deleteRow_(sheetName, keyField, keyValue) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 1) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getDataRange().getValues();
  const keyIdx = headers.indexOf(keyField);
  // de trás pra frente pra não bagunçar os índices das linhas seguintes ao deletar
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][keyIdx] === keyValue) sh.deleteRow(i + 1);
  }
}

/**
 * GET ?action=list  -> retorna tudo (added, contacts, imports) num JSON só
 */
function doGet(e) {
  const action = e.parameter.action || 'list';
  let payload;
  if (action === 'list') {
    payload = {
      added: sheetToObjects_(SHEET_ADDED),
      contacts: sheetToObjects_(SHEET_CONTACT),
      imports: sheetToObjects_(SHEET_IMPORT),
      meta: sheetToObjects_(SHEET_META),
      revenue: sheetToObjects_(SHEET_REVENUE),
      lifetime: sheetToObjects_(SHEET_LIFETIME),
      ok: true
    };
  } else {
    payload = { ok: false, error: 'ação desconhecida' };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST body JSON: { action: 'addAffiliate' | 'deleteAffiliate' | 'logContact' | 'importBatch' | 'saveAffiliateMeta' | 'importRevenueBatch' | 'importLifetimeBatch', data: {...} }
 */
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const now = new Date().toISOString();
  let result = { ok: true };

  try {
    if (action === 'addAffiliate') {
      const d = body.data;
      const key = normName_(d.nome);
      upsertRow_(SHEET_ADDED, 'nome_norm', key, {
        nome_norm: key, nome: d.nome, telefone: d.telefone || '', telegram: d.telegram || '',
        produto_planilha: d.produto_planilha || '', observacao: d.observacao || '', criado_em: now
      });
    } else if (action === 'deleteAffiliate') {
      const d = body.data;
      const key = normName_(d.nome);
      deleteRow_(SHEET_ADDED, 'nome_norm', key);
      deleteRow_(SHEET_META, 'nome_norm', key);
    } else if (action === 'saveAffiliateMeta') {
      const d = body.data;
      const key = normName_(d.nome);
      upsertRow_(SHEET_META, 'nome_norm', key, {
        nome_norm: key, nome: d.nome, nome_dash: d.nome_dash || '', trafego: d.trafego || '',
        cpa: d.cpa || '', buygoods_user: d.buygoods_user || '', atualizado_em: now
      });
    } else if (action === 'logContact') {
      const d = body.data;
      const key = normName_(d.nome);
      upsertRow_(SHEET_CONTACT, 'nome_norm', key, {
        nome_norm: key, nome: d.nome, oferta: d.oferta || '', resumo: d.resumo || '',
        data: d.data || '', atualizado_em: now
      });
    } else if (action === 'importBatch') {
      // body.data = array de registros de import (um por afiliado)
      const arr = body.data || [];
      arr.forEach(d => {
        const key = normName_(d.nome);
        upsertRow_(SHEET_IMPORT, 'nome_norm', key, {
          nome_norm: key, nome: d.nome,
          volume_ago: d.volume_ago, volume_set: d.volume_set,
          net_ago: d.net_ago, net_set: d.net_set, tem_dado_custo: d.tem_dado_custo,
          ultima_venda: d.ultima_venda, dias_sem_vender: d.dias_sem_vender, tier: d.tier,
          confirmado: d.confirmado, total_pedidos: d.total_pedidos, volume_total: d.volume_total,
          produtos: (d.produtos || []).join(', '), streak_diario: d.streak_diario || 0,
          dados_diarios: d.dados_diarios || '{}', atualizado_em: now
        });
      });
      result.processed = arr.length;
    } else if (action === 'importRevenueBatch') {
      // body.data = array de {nome, dados_json} (JSON acumulado dia a dia, já mesclado no cliente)
      const arr = body.data || [];
      arr.forEach(d => {
        const key = normName_(d.nome);
        upsertRow_(SHEET_REVENUE, 'nome_norm', key, {
          nome_norm: key, nome: d.nome, dados_json: d.dados_json || '{}', atualizado_em: now
        });
      });
      result.processed = arr.length;
    } else if (action === 'importLifetimeBatch') {
      // body.data = array COMPLETO (substitui a aba inteira) — importação única da base histórica
      const arr = body.data || [];
      const headers = ['nome_norm', 'nome', 'produtos', 'orders', 'gross', 'net', 'commissions', 'refunds', 'atualizado_em'];
      const rows = arr.map(d => [
        d.nome_norm || normName_(d.nome), d.nome || '', d.produtos || '',
        d.orders || 0, d.gross || 0, d.net || 0, d.commissions || 0, d.refunds || 0, now
      ]);
      bulkReplaceSheet_(SHEET_LIFETIME, headers, rows);
      result.processed = arr.length;
    } else {
      result = { ok: false, error: 'ação desconhecida: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}