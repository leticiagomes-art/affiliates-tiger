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
 * - AffiliateMeta: nome no dash da empresa, tipo de tráfego, CPA, usuário/e-mail BuyGoods, e overrides
 *   de telefone/telegram (preenche ou corrige contato de qualquer afiliado, mesmo os que vêm da base
 *   original ou só de import), por afiliado
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
    sh.appendRow(['nome_norm', 'nome', 'nome_dash', 'trafego', 'cpa', 'buygoods_user', 'telefone', 'telegram', 'atualizado_em']);
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

// Espelha o mesmo mapa de apelidos do index.html \u2014 pares confirmados manualmente como sendo
// a mesma pessoa com nome escrito diferente (n\u00e3o \u00e9 fuzzy-match autom\u00e1tico).
const NAME_ALIASES = {
  'fabio sobral torrezani': 'fabio torrezani'
};
function normName_(s) {
  if (!s) return '';
  const base = s.toString().trim().toLowerCase()
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim();
  return NAME_ALIASES[base] || base;
}

function pickNonEmpty_(a, b) {
  return (b !== undefined && b !== null && b !== '') ? b : a;
}

/**
 * Junta duas linhas (chaves nome_norm diferentes) que s\u00e3o confirmadamente a mesma pessoa,
 * em todas as abas relevantes \u2014 soma o que \u00e9 aditivo (pedidos, faturamento, totais vital\u00edcios),
 * une os dias de dados_diarios/dados_json, e prefere o valor n\u00e3o-vazio de "to" pros campos de
 * contato/cadastro. Apaga a linha "from" no final. Uso pontual, sempre chamado manualmente
 * depois de confirmar que \u00e9 a mesma pessoa \u2014 nunca autom\u00e1tico.
 */
function mergeAffiliateKey_(fromKey, toKey) {
  if (!fromKey || !toKey || fromKey === toKey) return;
  const now = new Date().toISOString();

  const meta = sheetToObjects_(SHEET_META);
  if (meta[fromKey]) {
    const f = meta[fromKey], t = meta[toKey] || {};
    upsertRow_(SHEET_META, 'nome_norm', toKey, {
      nome_norm: toKey, nome: t.nome || f.nome,
      nome_dash: pickNonEmpty_(f.nome_dash, t.nome_dash), trafego: pickNonEmpty_(f.trafego, t.trafego),
      cpa: pickNonEmpty_(f.cpa, t.cpa), buygoods_user: pickNonEmpty_(f.buygoods_user, t.buygoods_user),
      telefone: pickNonEmpty_(f.telefone, t.telefone), telegram: pickNonEmpty_(f.telegram, t.telegram),
      atualizado_em: now
    });
    deleteRow_(SHEET_META, 'nome_norm', fromKey);
  }

  const imports = sheetToObjects_(SHEET_IMPORT);
  if (imports[fromKey]) {
    const f = imports[fromKey], t = imports[toKey] || {};
    let diasF = {}, diasT = {};
    try { diasF = JSON.parse(f.dados_diarios || '{}'); } catch (e) {}
    try { diasT = JSON.parse(t.dados_diarios || '{}'); } catch (e) {}
    const dias = Object.assign({}, diasF, diasT);
    let totalPedidos = 0, volumeTotal = 0;
    Object.values(dias).forEach(d => { totalPedidos += Number(d.orders) || 0; volumeTotal += Number(d.volume) || 0; });
    const ultimaF = f.ultima_venda ? String(f.ultima_venda).slice(0, 10) : null;
    const ultimaT = t.ultima_venda ? String(t.ultima_venda).slice(0, 10) : null;
    const ultimaVenda = [ultimaF, ultimaT].filter(Boolean).sort().pop() || null;
    const produtosSet = new Set((String(f.produtos || '').split(',').concat(String(t.produtos || '').split(',')))
      .map(s => s.trim()).filter(Boolean));
    upsertRow_(SHEET_IMPORT, 'nome_norm', toKey, {
      nome_norm: toKey, nome: t.nome || f.nome,
      ultima_venda: ultimaVenda, dias_sem_vender: '', tier: t.tier || f.tier || 'sem_venda_2026',
      confirmado: totalPedidos > 0, total_pedidos: totalPedidos, volume_total: Math.round(volumeTotal * 100) / 100,
      produtos: Array.from(produtosSet).join(', '),
      streak_diario: Math.max(Number(f.streak_diario) || 0, Number(t.streak_diario) || 0),
      dados_diarios: JSON.stringify(dias),
      volume_ago: pickNonEmpty_(f.volume_ago, t.volume_ago), volume_set: pickNonEmpty_(f.volume_set, t.volume_set),
      net_ago: pickNonEmpty_(f.net_ago, t.net_ago), net_set: pickNonEmpty_(f.net_set, t.net_set),
      tem_dado_custo: f.tem_dado_custo || t.tem_dado_custo || false,
      atualizado_em: now
    });
    deleteRow_(SHEET_IMPORT, 'nome_norm', fromKey);
  }

  const revenue = sheetToObjects_(SHEET_REVENUE);
  if (revenue[fromKey]) {
    const f = revenue[fromKey], t = revenue[toKey] || {};
    let recF = {}, recT = {};
    try { recF = JSON.parse(f.dados_json || '{}'); } catch (e) {}
    try { recT = JSON.parse(t.dados_json || '{}'); } catch (e) {}
    const merged = Object.assign({}, recF);
    Object.keys(recT).forEach(day => {
      if (merged[day]) {
        merged[day] = {
          vendas: (Number(merged[day].vendas) || 0) + (Number(recT[day].vendas) || 0),
          faturamento: Math.round(((Number(merged[day].faturamento) || 0) + (Number(recT[day].faturamento) || 0)) * 100) / 100,
          produtos: Array.from(new Set((merged[day].produtos || []).concat(recT[day].produtos || [])))
        };
      } else {
        merged[day] = recT[day];
      }
    });
    upsertRow_(SHEET_REVENUE, 'nome_norm', toKey, {
      nome_norm: toKey, nome: t.nome || f.nome, dados_json: JSON.stringify(merged), atualizado_em: now
    });
    deleteRow_(SHEET_REVENUE, 'nome_norm', fromKey);
  }

  const lifetime = sheetToObjects_(SHEET_LIFETIME);
  if (lifetime[fromKey]) {
    const f = lifetime[fromKey], t = lifetime[toKey] || {};
    const produtosSet = new Set((String(f.produtos || '').split(',').concat(String(t.produtos || '').split(',')))
      .map(s => s.trim()).filter(Boolean));
    upsertRow_(SHEET_LIFETIME, 'nome_norm', toKey, {
      nome_norm: toKey, nome: t.nome || f.nome, produtos: Array.from(produtosSet).join(', '),
      orders: (Number(f.orders) || 0) + (Number(t.orders) || 0),
      gross: Math.round(((Number(f.gross) || 0) + (Number(t.gross) || 0)) * 100) / 100,
      net: Math.round(((Number(f.net) || 0) + (Number(t.net) || 0)) * 100) / 100,
      commissions: Math.round(((Number(f.commissions) || 0) + (Number(t.commissions) || 0)) * 100) / 100,
      refunds: Math.round(((Number(f.refunds) || 0) + (Number(t.refunds) || 0)) * 100) / 100,
      atualizado_em: now
    });
    deleteRow_(SHEET_LIFETIME, 'nome_norm', fromKey);
  }

  const contacts = sheetToObjects_(SHEET_CONTACT);
  if (contacts[fromKey]) {
    const f = contacts[fromKey], t = contacts[toKey];
    const maisRecente = (!t || String(f.data || '') > String(t.data || '')) ? f : t;
    upsertRow_(SHEET_CONTACT, 'nome_norm', toKey, {
      nome_norm: toKey, nome: maisRecente.nome, oferta: maisRecente.oferta || '',
      resumo: maisRecente.resumo || '', data: maisRecente.data || '', atualizado_em: now
    });
    deleteRow_(SHEET_CONTACT, 'nome_norm', fromKey);
  }
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

function bulkMergeSheet_(sheetName, keyField, updates) {
  // mescla uma lista de atualizações parciais numa aba, lendo e escrevendo cada uma só 1 vez —
  // ao contrário de upsertRow_ (que relê a planilha inteira a cada chamada), isso não fica lento
  // com centenas/milhares de linhas, e ao contrário de bulkReplaceSheet_, preserva colunas/linhas
  // que as atualizações não mencionam
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  let headers = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [keyField];
  const headerSet = {};
  headers.forEach(h => headerSet[h] = true);
  updates.forEach(u => Object.keys(u).forEach(k => { if (!headerSet[k]) { headerSet[k] = true; headers.push(k); } }));

  const data = sh.getLastRow() > 0 ? sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues() : [];
  const oldHeaders = data.length ? data[0] : [];
  const rowMap = {}; // keyValue -> objeto com os campos já existentes
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    oldHeaders.forEach((h, idx) => obj[h] = data[i][idx]);
    if (obj[keyField]) rowMap[obj[keyField]] = obj;
  }
  updates.forEach(u => {
    const key = u[keyField];
    if (!key) return;
    if (!rowMap[key]) rowMap[key] = {};
    Object.keys(u).forEach(k => { rowMap[key][k] = u[k]; });
  });

  const outRows = Object.keys(rowMap).map(k => headers.map(h => (h in rowMap[k]) ? rowMap[k][h] : ''));
  sh.clear();
  if (outRows.length > 0) {
    sh.getRange(1, 1, outRows.length + 1, headers.length).setValues([headers].concat(outRows));
  } else {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
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
 * POST body JSON: { action: 'addAffiliate' | 'deleteAffiliate' | 'logContact' | 'importBatch' | 'saveAffiliateMeta' | 'importRevenueBatch' | 'importLifetimeBatch' | 'importContactsBatch' | 'mergeAffiliateKeys', data: {...} }
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
        cpa: d.cpa || '', buygoods_user: d.buygoods_user || '',
        telefone: d.telefone || '', telegram: d.telegram || '', atualizado_em: now
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
    } else if (action === 'importContactsBatch') {
      // body.data = array de atualizações parciais (nome_norm + campos a mesclar em AffiliateMeta)
      const arr = body.data || [];
      const updates = arr.map(d => {
        const u = { nome_norm: d.nome_norm || normName_(d.nome), atualizado_em: now };
        ['nome', 'telefone', 'telegram', 'nome_dash', 'trafego', 'cpa', 'buygoods_user'].forEach(f => {
          if (d[f] !== undefined && d[f] !== null && d[f] !== '') u[f] = d[f];
        });
        return u;
      });
      bulkMergeSheet_(SHEET_META, 'nome_norm', updates);
      result.processed = arr.length;
    } else if (action === 'mergeAffiliateKeys') {
      // body.data = array de {from, to} (nome_norm) — junta linhas confirmadas como a mesma pessoa
      const arr = body.data || [];
      arr.forEach(pair => { if (pair.from && pair.to) mergeAffiliateKey_(pair.from, pair.to); });
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