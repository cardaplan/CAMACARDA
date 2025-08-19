/**
 * Cardaplan CRM - Backend Adaptado para Planilha Específica (estendido)
 * Abas suportadas:
 * - Itens: SKU, item, categoria, descricao, preco, status, classificacao_adicional, observacoes, foto_url, estoque_inicial, estoque_atual, controle_estoque, usar_id_alternado
 * - Categorias: nome_categoria, titulo_exibicao, descricao, ordem, status
 * - Config: section, key, value
 * - Horários: Dia da Semana, Período 1, Período 2, Período 3
 * - Bairros: nome_bairro, valor_taxa
 * - Cupons: codigo_cupom, tipo_desconto, valor_desconto, data_inicio, data_fim, condicao_min_total, produto_sku, desconto_alternativo
 * - Analytics: data_acesso, tipo_evento, item_id, sessao_id, user_agent, ip_address
 * - Pedidos: data_hora, cliente_nome, cliente_whatsapp, tipo_entrega, endereco, itens_json, subtotal, taxa_entrega, desconto, total, forma_pagamento, observacoes, status_pedido
 * - Relatórios: ano_mes, produto_sku, produto, quantidade, valor_total
 * - Estoque_Historico: data_hora, SKU, item, movimento, quantidade, estoque_antes, estoque_depois, origem, observacao
 */

const SPREADSHEET_ID = '1oQQ6xq8OFTupdo63MTh4ZheQ13lGUMg2I4ggGMRPWnE';

const SHEET_CONFIG = {
  'Itens': {
    displayName: 'Itens',
    icon: '📦',
    fields: [
      { name: 'SKU', type: 'text', required: true },
      { name: 'item', type: 'text', required: true },
      { name: 'categoria', type: 'select', required: true, source: 'Categorias' },
      { name: 'descricao', type: 'textarea', required: false },
      { name: 'preco', type: 'text', required: true },
      { name: 'status', type: 'select', required: true, options: ['Ativo', 'Inativo', 'Esgotado'] },
      { name: 'classificacao_adicional', type: 'text', required: false },
      { name: 'observacoes', type: 'textarea', required: false },
      { name: 'foto_url', type: 'url', required: false },
      { name: 'estoque_inicial', type: 'number', required: false },
      { name: 'estoque_atual', type: 'number', required: false },
      { name: 'controle_estoque', type: 'select', required: false, options: ['Sim', 'Não'] },
      { name: 'usar_id_alternado', type: 'select', required: false, options: ['Sim', 'Não'] }
    ]
  },
  'Categorias': {
    displayName: 'Categorias',
    icon: '🏷️',
    fields: [
      { name: 'nome_categoria', type: 'text', required: true },
      { name: 'titulo_exibicao', type: 'text', required: true },
      { name: 'descricao', type: 'textarea', required: false },
      { name: 'ordem', type: 'number', required: false },
      { name: 'status', type: 'select', required: true, options: ['Ativo', 'Inativo'] }
    ]
  },
  'Config': {
    displayName: 'Configurações',
    icon: '⚙️',
    fields: [
      { name: 'section', type: 'text', required: true },
      { name: 'key', type: 'text', required: true },
      { name: 'value', type: 'text', required: true }
    ]
  },
  'Horários': {
    displayName: 'Horários',
    icon: '🕐',
    fields: [
      { name: 'Dia da Semana', type: 'select', required: true, options: ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'] },
      { name: 'Período 1', type: 'text', required: false },
      { name: 'Período 2', type: 'text', required: false },
      { name: 'Período 3', type: 'text', required: false }
    ]
  },
  'Bairros': {
    displayName: 'Bairros',
    icon: '🏘️',
    fields: [
      { name: 'nome_bairro', type: 'text', required: true },
      { name: 'valor_taxa', type: 'currency', required: true }
    ]
  },
  'Cupons': {
    displayName: 'Cupons',
    icon: '🎟️',
    fields: [
      { name: 'codigo_cupom', type: 'text', required: true },
      { name: 'tipo_desconto', type: 'select', required: true, options: ['produtos', 'total', 'frete'] },
      { name: 'valor_desconto', type: 'number%', required: true },
      { name: 'data_inicio', type: 'dd/mm/aa', required: true },
      { name: 'data_fim', type: 'dd/mm/aaaa', required: true },
      { name: 'condicao_min_total', type: 'currency', required: false },
      { name: 'produto_sku', type: 'text', required: false },
      { name: 'desconto_alternativo', type: 'text', required: false }
    ]
  },
  'Analytics': {
    displayName: 'Analytics',
    icon: '📊',
    fields: [
      { name: 'data_acesso', type: 'datetime', required: true },
      { name: 'tipo_evento', type: 'select', required: true, options: ['visualizacao_cardapio', 'visualizacao_item', 'sessao_unica'] },
      { name: 'item_id', type: 'text', required: false },
      { name: 'sessao_id', type: 'text', required: true },
      { name: 'user_agent', type: 'text', required: false },
      { name: 'ip_address', type: 'text', required: false }
    ]
  },
  'Pedidos': {
    displayName: 'Pedidos',
    icon: '🔔',
    fields: [
      { name: 'data_hora', type: 'datetime', required: true },
      { name: 'cliente_nome', type: 'text', required: true },
      { name: 'cliente_whatsapp', type: 'text', required: false },
      { name: 'tipo_entrega', type: 'text', required: true },
      { name: 'endereco', type: 'text', required: false },
      { name: 'itens_json', type: 'text', required: true },
      { name: 'subtotal', type: 'currency', required: false },
      { name: 'taxa_entrega', type: 'currency', required: false },
      { name: 'desconto', type: 'currency', required: false },
      { name: 'total', type: 'currency', required: true },
      { name: 'forma_pagamento', type: 'text', required: false },
      { name: 'observacoes', type: 'text', required: false },
      { name: 'status_pedido', type: 'select', required: true, options: ['Pendente', 'Pedido pago ✅', 'Pedido cancelado ❌'] }
    ]
  },
  'Relatórios': {
    displayName: 'Relatórios',
    icon: '📊',
    fields: [
      { name: 'ano_mes', type: 'text', required: true },
      { name: 'produto_sku', type: 'text', required: true },
      { name: 'produto', type: 'text', required: true },
      { name: 'quantidade', type: 'number', required: true },
      { name: 'valor_total', type: 'currency', required: true }
    ]
  },
  'Estoque_Historico': {
    displayName: 'Histórico de Estoque',
    icon: '📦',
    fields: [
      { name: 'data_hora', type: 'datetime', required: true },
      { name: 'SKU', type: 'text', required: true },
      { name: 'item', type: 'text', required: true },
      { name: 'movimento', type: 'select', required: true, options: ['Entrada', 'Saída', 'Ajuste'] },
      { name: 'quantidade', type: 'number', required: true },
      { name: 'estoque_antes', type: 'number', required: false },
      { name: 'estoque_depois', type: 'number', required: false },
      { name: 'origem', type: 'text', required: false },
      { name: 'observacao', type: 'text', required: false }
    ]
  }
};

const cache = { data: {}, timestamp: {}, TTL: 5 * 60 * 1000 };

function doGet(e) {
  try {
    if (e && e.parameter && (e.parameter.data || e.parameter.message)) {
      const result = handleIncomingOrderGet(e.parameter);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    const htmlOutput = HtmlService.createHtmlOutputFromFile('aplicacaoweb')
      .setTitle('Cardaplan - Gestão Inteligente')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return htmlOutput;
  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Erro ao carregar a aplicação</h1><p>Tente novamente em alguns instantes.</p>');
  }
}

function getSheetNames() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheets().map(s => s.getName());
}

function getSheetConfig(sheetName) { if (!SHEET_CONFIG[sheetName]) throw new Error('Configuração não encontrada para a aba: ' + sheetName); return SHEET_CONFIG[sheetName]; }
function getAllSheetConfigs() { return SHEET_CONFIG; }
function isCacheValid(key) { return cache.timestamp[key] && (Date.now() - cache.timestamp[key]) < cache.TTL; }

function getSheetData(sheetName) {
  if (isCacheValid(sheetName) && cache.data[sheetName]) return cache.data[sheetName];
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Aba "${sheetName}" não encontrada na planilha`);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0) return [];
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  const values = range.getValues();
  const filtered = values.filter(row => row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== ''));
  cache.data[sheetName] = filtered; cache.timestamp[sheetName] = Date.now();
  return filtered;
}

function getAllSheetsData() {
  const result = {}; const sheetNames = Object.keys(SHEET_CONFIG);
  for (const sheetName of sheetNames) {
    try { result[sheetName] = getSheetData(sheetName); } catch (_) { result[sheetName] = []; }
  }
  return result;
}

function addRow(sheetName, rowData) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  if (sheet.getLastRow() === 0) sheet.appendRow(SHEET_CONFIG[sheetName].fields.map(f => f.name));
  sheet.appendRow(rowData);
  delete cache.data[sheetName]; delete cache.timestamp[sheetName];
  return true;
}

function updateRow(sheetName, rowIndex, rowData) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Aba "${sheetName}" não encontrada`);
  const range = sheet.getRange(rowIndex, 1, 1, rowData.length);
  range.setValues([rowData]);
  delete cache.data[sheetName]; delete cache.timestamp[sheetName];
  return true;
}

function deleteRow(sheetName, rowIndex) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Aba "${sheetName}" não encontrada`);
  sheet.deleteRow(rowIndex);
  delete cache.data[sheetName]; delete cache.timestamp[sheetName];
  return true;
}

function searchAllSheets(query) {
  if (!query || query.trim() === '') return [];
  const searchQuery = query.toLowerCase();
  const results = [];
  for (const [sheetName] of Object.entries(SHEET_CONFIG)) {
    const data = getSheetData(sheetName);
    if (data.length <= 1) continue;
    const headers = data[0]; const rows = data.slice(1);
    rows.forEach((row, rowIndex) => {
      const rowText = row.map(cell => String(cell)).join(' ').toLowerCase();
      if (rowText.includes(searchQuery)) results.push({ sheet: sheetName, rowIndex: rowIndex + 2, data: row });
    });
  }
  return results;
}

function getDashboardStats() {
  const stats = { totalItens: 0, itensAtivos: 0, totalCategorias: 0, totalCupons: 0, cuponsAtivos: 0, totalBairros: 0, totalHorarios: 0, totalVisualizacoes: 0, sessesUnicas: 0, visualizacoesHoje: 0 };
  const itensData = getSheetData('Itens');
  if (itensData.length > 1) { stats.totalItens = itensData.length - 1; stats.itensAtivos = itensData.slice(1).filter(r => r[5] === 'Ativo').length; }
  const categoriasData = getSheetData('Categorias'); if (categoriasData.length > 1) stats.totalCategorias = categoriasData.length - 1;
  const cuponsData = getSheetData('Cupons'); if (cuponsData.length > 1) { stats.totalCupons = cuponsData.length - 1; const hoje = new Date(); stats.cuponsAtivos = cuponsData.slice(1).filter(row => { const [di, mi, ai] = (row[3] || '').split('/'); const [df, mf, af] = (row[4] || '').split('/'); if (di && mi && ai && df && mf && af) { const inicio = new Date(`${ai}-${mi}-${di}T00:00:00-03:00`); const fim = new Date(`${af}-${mf}-${df}T23:59:59-03:00`); return hoje >= inicio && hoje <= fim; } return false; }).length; }
  const bairrosData = getSheetData('Bairros'); if (bairrosData.length > 1) stats.totalBairros = bairrosData.length - 1;
  const horariosData = getSheetData('Horários').slice(2); if (horariosData.length > 1) stats.totalHorarios = horariosData.length - 1;
  try {
    const analyticsData = getSheetData('Analytics'); if (analyticsData.length > 1) { const rows = analyticsData.slice(1); stats.totalVisualizacoes = rows.filter(r => r[1] === 'visualizacao_cardapio').length; const sessoes = new Set(); rows.forEach(r => { if (r[3]) sessoes.add(r[3]); }); stats.sessesUnicas = sessoes.size; const hoje = new Date(); const hojeStr = hoje.toDateString(); stats.visualizacoesHoje = rows.filter(r => { if (r[0]) { const d = new Date(r[0]); return d.toDateString() === hojeStr && r[1] === 'visualizacao_cardapio'; } return false; }).length; }
  } catch (_) {}
  return stats;
}

function registrarEventoAnalytics(tipoEvento, itemId, sessaoId, userAgent, ipAddress) {
  const analyticsData = [ new Date(), tipoEvento, itemId || '', sessaoId || generateSessionId(), userAgent || '', ipAddress || '' ];
  return addRow('Analytics', analyticsData);
}

function generateSessionId() { return 'sess_' + Utilities.getUuid(); }

function getAnalyticsReports(di, df) {
  const analyticsData = getSheetData('Analytics');
  if (analyticsData.length <= 1) return { visualizacoesPorDia: [], produtosMaisVisualizados: [], sessoesPorDia: [], resumoGeral: { totalVisualizacoes: 0, sessesUnicas: 0, produtosMaisVistos: 0 } };
  const rows = analyticsData.slice(1); let filteredRows = rows;
  if (di || df) { const inicio = di ? new Date(di) : new Date('01/01/2020'); const fim = df ? new Date(df) : new Date(); filteredRows = rows.filter(row => { if (row[0]) { const dataAcesso = new Date(row[0]); return dataAcesso >= inicio && dataAcesso <= fim; } return false; }); }
  const visualizacoesPorDia = {}; const sessoesPorDia = {}; const produtosVisualizados = {}; const sessoes = new Set();
  filteredRows.forEach(row => { const dataAcesso = row[0] ? new Date(row[0]) : null; const tipoEvento = row[1]; const itemId = row[2]; const sessaoId = row[3]; if (dataAcesso) { const dataStr = dataAcesso.toISOString().split('T')[0]; if (tipoEvento === 'visualizacao_cardapio') visualizacoesPorDia[dataStr] = (visualizacoesPorDia[dataStr] || 0) + 1; if (sessaoId) { if (!sessoesPorDia[dataStr]) sessoesPorDia[dataStr] = new Set(); sessoesPorDia[dataStr].add(sessaoId); sessoes.add(sessaoId); } if (tipoEvento === 'visualizacao_item' && itemId) produtosVisualizados[itemId] = (produtosVisualizados[itemId] || 0) + 1; } });
  const sessoesPorDiaNumeros = {}; Object.keys(sessoesPorDia).forEach(data => { sessoesPorDiaNumeros[data] = sessoesPorDia[data].size; });
  const top = Object.entries(produtosVisualizados).sort(([,a],[,b]) => b-a).slice(0,10).map(([itemId, visualizacoes]) => ({ itemId, visualizacoes, nomeItem: getItemName(itemId) }));
  return { visualizacoesPorDia: Object.entries(visualizacoesPorDia).map(([data, count]) => ({ data, visualizacoes: count })), produtosMaisVisualizados: top, sessoesPorDia: Object.entries(sessoesPorDiaNumeros).map(([data, count]) => ({ data, sessoes: count })), resumoGeral: { totalVisualizacoes: Object.values(visualizacoesPorDia).reduce((a,b)=>a+b,0), sessesUnicas: sessoes.size, produtosMaisVistos: top.length } };
}

function getItemName(itemId) { try { const itensData = getSheetData('Itens'); if (itensData.length > 1) { const item = itensData.slice(1).find(row => row[0] === itemId); return item ? item[1] : itemId; } return itemId; } catch (_) { return itemId; } }

function gerarRelatorioPDF(dadosRelatorio) {
  const doc = DocumentApp.create('Relatório Cardaplan - ' + new Date().toLocaleDateString('pt-BR')); const body = doc.getBody();
  const titulo = body.appendParagraph('RELATÓRIO CARDAPLAN'); titulo.setHeading(DocumentApp.ParagraphHeading.TITLE); titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Gerado em: ' + new Date().toLocaleString('pt-BR')); body.appendParagraph('');
  body.appendParagraph('RESUMO GERAL').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Total de Visualizações: ${dadosRelatorio.resumoGeral.totalVisualizacoes}`);
  body.appendParagraph(`Sessões Únicas: ${dadosRelatorio.resumoGeral.sessesUnicas}`);
  body.appendParagraph(`Produtos Mais Vistos: ${dadosRelatorio.resumoGeral.produtosMaisVistos}`);
  body.appendParagraph('');
  if (dadosRelatorio.visualizacoesPorDia.length > 0) { body.appendParagraph('VISUALIZAÇÕES POR DIA').setHeading(DocumentApp.ParagraphHeading.HEADING1); dadosRelatorio.visualizacoesPorDia.forEach(item => { body.appendParagraph(`${item.data}: ${item.visualizacoes} visualizações`); }); body.appendParagraph(''); }
  if (dadosRelatorio.produtosMaisVisualizados.length > 0) { body.appendParagraph('PRODUTOS MAIS VISUALIZADOS').setHeading(DocumentApp.ParagraphHeading.HEADING1); dadosRelatorio.produtosMaisVisualizados.forEach((produto, index) => { body.appendParagraph(`${index + 1}. ${produto.nomeItem}: ${produto.visualizacoes} visualizações`); }); }
  doc.saveAndClose(); const file = DriveApp.getFileById(doc.getId()); const pdfBlob = file.getBlob().getAs('application/pdf'); const pdfFile = DriveApp.createFile(pdfBlob); pdfFile.setName('Relatorio_Cardaplan_' + new Date().toISOString().split('T')[0] + '.pdf'); DriveApp.getFileById(doc.getId()).setTrashed(true);
  return { success: true, fileId: pdfFile.getId(), fileName: pdfFile.getName(), downloadUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/view` };
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// ==== Config utils ====
function getConfigValue(section, key, defaultValue) {
  try { const data = getSheetData('Config'); if (data.length <= 1) return defaultValue; const rows = data.slice(1); const found = rows.find(r => String(r[0]||'').toLowerCase() === String(section).toLowerCase() && String(r[1]||'').toLowerCase() === String(key).toLowerCase()); return found ? (found[2] || defaultValue) : defaultValue; } catch (_) { return defaultValue; }
}

// ==== Pedidos / Webhook ====
function handleIncomingOrderGet(params) {
  try { const dataJson = params.data ? JSON.parse(params.data) : {}; dataJson.whatsapp_admin = params.whatsapp_admin || dataJson.whatsapp_admin || ''; dataJson.whatsapp_customer = params.whatsapp_customer || dataJson.whatsapp_customer || ''; dataJson.customer_name = params.customer_name || dataJson.customer_name || ''; const total = parseFloat(params.total_final || params.total || dataJson.total || 0) || 0; dataJson.total = total; return handleIncomingOrder(dataJson); } catch (error) { return { success: false, message: 'Erro ao processar pedido GET: ' + error.message }; }
}

function handleIncomingOrder(order) {
  const now = new Date();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pedidosSheet = ss.getSheetByName('Pedidos') || ss.insertSheet('Pedidos');
  if (pedidosSheet.getLastRow() === 0) pedidosSheet.appendRow(SHEET_CONFIG['Pedidos'].fields.map(f => f.name));
  const customer = order.customer || {};
  const items = order.items || [];
  const deliveryType = (order.customer && order.customer.delivery_type) || order.customer_delivery_type || order.delivery_type || '';
  const address = order.customer && order.customer.address ? order.customer.address : (order.address || null);
  const subtotal = (order.totals && order.totals.subtotal_products_numeric) || order.subtotal || 0;
  const deliveryFee = (order.totals && order.totals.delivery_fee_numeric) || order.delivery_fee || 0;
  const total = (order.totals && order.totals.total_final_numeric) || order.total || 0;
  const discount = (order.coupon && order.coupon.discount_numeric) || 0;
  const payment = (order.customer && order.customer.payment_method) || order.payment_method || '';
  const notes = (order.customer && order.customer.notes) || order.notes || '';
  const itensComIdAlternado = addAlternateIdsIfEnabled(items);
  pedidosSheet.appendRow([
    now,
    order.customer_name || customer.name || '',
    order.whatsapp_customer || order.customer_whatsapp || '',
    deliveryType,
    address ? JSON.stringify(address) : '',
    JSON.stringify(itensComIdAlternado),
    Number(subtotal)||0,
    Number(deliveryFee)||0,
    Number(discount)||0,
    Number(total)||0,
    payment||'',
    notes||'',
    'Pendente'
  ]);
  try { atualizarEstoqueAPartirDoPedido(items); } catch (e) { Logger.log('Erro estoque: ' + e.message); }
  try { atualizarRelatorioMensal(items); } catch (e) { Logger.log('Erro relatório: ' + e.message); }
  return { success: true, message: 'Pedido registrado', createdAt: now.toISOString() };
}

function atualizarEstoqueAPartirDoPedido(items) {
  if (!items || !items.length) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Itens'); if (!sheet) return;
  const data = sheet.getDataRange().getValues(); const headers = data[0];
  const idxSKU = headers.indexOf('SKU'); const idxItem = headers.indexOf('item'); const idxControle = headers.indexOf('controle_estoque'); const idxEstoqueAtual = headers.indexOf('estoque_atual'); const idxStatus = headers.indexOf('status');
  const historico = ss.getSheetByName('Estoque_Historico') || ss.insertSheet('Estoque_Historico'); if (historico.getLastRow() === 0) historico.appendRow(SHEET_CONFIG['Estoque_Historico'].fields.map(f => f.name));
  const map = {}; for (let i=1;i<data.length;i++){ const sku = String(data[i][idxSKU]||'').trim(); if (sku) map[sku]=i+1; }
  items.forEach(item => { const sku = String(item.sku || item.SKU || '').trim(); const qty = Number(item.quantity || item.qtd || 0); if (!sku||!qty) return; const rowIndex = map[sku]; if (!rowIndex) return; const rowValues = sheet.getRange(rowIndex,1,1,headers.length).getValues()[0]; const controle = String(rowValues[idxControle]||'').toLowerCase(); if (controle !== 'sim') return; const antes = Number(rowValues[idxEstoqueAtual]||0); const depois = Math.max(0, antes - qty); if (idxEstoqueAtual>=0) sheet.getRange(rowIndex, idxEstoqueAtual+1).setValue(depois); if (depois===0 && idxStatus>=0) sheet.getRange(rowIndex, idxStatus+1).setValue('Esgotado'); const nome = rowValues[idxItem] || sku; historico.appendRow([new Date(), sku, nome, 'Saída', qty, antes, depois, 'Venda', 'Baixa automática por pedido']); });
}

function atualizarRelatorioMensal(items) {
  if (!items || !items.length) return; const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const rel = ss.getSheetByName('Relatórios') || ss.insertSheet('Relatórios'); if (rel.getLastRow()===0) rel.appendRow(SHEET_CONFIG['Relatórios'].fields.map(f=>f.name)); const ym = new Date().toISOString().slice(0,7); const data = rel.getDataRange().getValues(); const headers = data[0]; const idxYM = headers.indexOf('ano_mes'); const idxSKU = headers.indexOf('produto_sku'); const idxNome = headers.indexOf('produto'); const idxQtd = headers.indexOf('quantidade'); const idxValor = headers.indexOf('valor_total'); const chaveToRow = {}; for(let i=1;i<data.length;i++){ const key = `${data[i][idxYM]}|${data[i][idxSKU]}`; chaveToRow[key]=i+1; } items.forEach(it=>{ const sku = it.sku||''; const nome = it.name||it.nome||sku; const qtd = Number(it.quantity||0); const valor = Number(it.subtotal_numeric||it.subtotal||0); const key = `${ym}|${sku}`; const rowIndex = chaveToRow[key]; if (rowIndex){ const qtdAtual = Number(rel.getRange(rowIndex, idxQtd+1).getValue()||0); const valAtual = Number(rel.getRange(rowIndex, idxValor+1).getValue()||0); rel.getRange(rowIndex, idxQtd+1).setValue(qtdAtual+qtd); rel.getRange(rowIndex, idxValor+1).setValue(valAtual+valor); } else { rel.appendRow([ym, sku, nome, qtd, valor]); } });
}

function addAlternateIdsIfEnabled(items) {
  if (!items || !items.length) return []; const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheet = ss.getSheetByName('Itens'); if (!sheet) return items; const data = sheet.getDataRange().getValues(); const headers = data[0]; const idxSKU = headers.indexOf('SKU'); const idxUsar = headers.indexOf('usar_id_alternado'); const usarBySku = {}; for (let i=1;i<data.length;i++){ const sku = String(data[i][idxSKU]||'').trim(); if (sku) usarBySku[sku] = String(data[i][idxUsar]||'').toLowerCase()==='sim'; } return items.map(it=>{ const sku = it.sku||''; if (sku && usarBySku[sku]) { return Object.assign({}, it, { alt_id: generateAlternateId() }); } return it; });
}

function generateAlternateId() { const props = PropertiesService.getScriptProperties(); const current = Number(props.getProperty('ALT_ID_INDEX')||'0'); const next = current + 1; props.setProperty('ALT_ID_INDEX', String(next)); const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'; const letter = letters[next % letters.length]; const number = (next * 7) % 100; return `${letter}${String(number).padStart(2,'0')}`; }

function getOrders() { const data = getSheetData('Pedidos'); if (data.length<=1) return []; const headers = data[0]; return data.slice(1).map((row, idx)=>({ rowIndex: idx+2, data: headers.reduce((acc,h,i)=>{acc[h]=row[i]; return acc;}, {}) })); }
function updateOrderStatus(rowIndex, novoStatus) { const ss = SpreadsheetApp.openById(SPREADSHEET_ID); const sheet = ss.getSheetByName('Pedidos'); if (!sheet) throw new Error('Aba Pedidos não encontrada'); const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]; const idxStatus = headers.indexOf('status_pedido'); if (idxStatus<0) throw new Error('Coluna status_pedido não encontrada'); sheet.getRange(rowIndex, idxStatus+1).setValue(novoStatus); return true; }

function getSalesReports(dataInicio, dataFim) {
  let pedidos = getOrders().filter(p => (p.data.status_pedido || '').toString().toLowerCase().indexOf('pago')>=0);
  // Filtro por período (datas no timezone local do Apps Script)
  if (dataInicio || dataFim) {
    const inicio = dataInicio ? new Date(dataInicio) : new Date('2000-01-01');
    const fim = dataFim ? new Date(dataFim) : new Date('2999-12-31T23:59:59');
    pedidos = pedidos.filter(p => {
      const dh = new Date(p.data.data_hora);
      return dh >= inicio && dh <= fim;
    });
  }
  const produtoCount = {}; const produtoNome = {}; let totalMesAtual = 0; const now = new Date(); const ymAtual = now.toISOString().slice(0,7); const pedidosDia=[]; const pedidosSemana=[]; const pedidosMes=[]; const uniqueCustomers=new Set();
  pedidos.forEach(p=>{ const dh=new Date(p.data.data_hora); const items = JSON.parse(p.data.itens_json||'[]'); const total=Number(p.data.total||0); const ym = dh.toISOString().slice(0,7); if (ym===ymAtual) totalMesAtual+=total; uniqueCustomers.add((p.data.cliente_whatsapp||'').toString().trim()); items.forEach(it=>{ const sku=it.sku||''; const nome=it.name||it.nome||sku; const qtd=Number(it.quantity||0); produtoCount[sku]=(produtoCount[sku]||0)+qtd; if(!produtoNome[sku]) produtoNome[sku]=nome; }); const reg={data:dh,pedido:p}; const hojeStr=now.toDateString(); if (dh.toDateString()===hojeStr) pedidosDia.push(reg); const diff=(now-dh)/(1000*60*60*24); if (diff<=7) pedidosSemana.push(reg); if (ym===ymAtual) pedidosMes.push(reg); });
  const produtosMaisPedidos = Object.entries(produtoCount).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([sku,qtd])=>({ sku, produto: produtoNome[sku]||sku, quantidade:qtd }));
  let produtosMaisVisualizados=[]; try { const analytics=getAnalyticsReports(dataInicio||null, dataFim||null); produtosMaisVisualizados=analytics.produtosMaisVisualizados||[]; } catch(_) {}
  return { produtosMaisVisualizados, produtosMaisPedidos, valorTotalVendidoNoMes: totalMesAtual, pedidos: { dia: pedidosDia.map(x=>x.pedido), semana: pedidosSemana.map(x=>x.pedido), mes: pedidosMes.map(x=>x.pedido) }, kpis: { totalPedidosHoje: pedidosDia.length, ticketMedio: pedidosMes.length ? (totalMesAtual / pedidosMes.length) : 0, produtoMaisVendido: produtosMaisPedidos[0] || null, totalClientesUnicos: Array.from(uniqueCustomers).filter(Boolean).length } };
}

// Gera XLSX com resumo dos relatórios
function gerarRelatorioXLSX(dadosRelatorio) {
  try {
    const ss = SpreadsheetApp.create('Relatório Cardaplan XLSX - ' + new Date().toLocaleDateString('pt-BR'));
    const sheet = ss.getActiveSheet();
    sheet.setName('Resumo');
    let row = 1;
    sheet.getRange(row++,1).setValue('RESUMO GERAL');
    sheet.getRange(row++,1,1,2).setValues([[ 'Pedidos Hoje', (dadosRelatorio.kpis && dadosRelatorio.kpis.totalPedidosHoje) || 0 ]]);
    sheet.getRange(row++,1,1,2).setValues([[ 'Ticket Médio', (dadosRelatorio.kpis && dadosRelatorio.kpis.ticketMedio) || 0 ]]);
    sheet.getRange(row++,1,1,2).setValues([[ 'Clientes Únicos', (dadosRelatorio.kpis && dadosRelatorio.kpis.totalClientesUnicos) || 0 ]]);
    row++;
    sheet.getRange(row++,1).setValue('Produtos mais visualizados');
    sheet.getRange(row++,1,1,2).setValues([[ 'Produto', 'Visualizações' ]]);
    (dadosRelatorio.produtosMaisVisualizados||[]).forEach(item => {
      sheet.getRange(row++,1,1,2).setValues([[ item.nomeItem, item.visualizacoes ]]);
    });
    row++;
    sheet.getRange(row++,1).setValue('Produtos mais pedidos');
    sheet.getRange(row++,1,1,2).setValues([[ 'Produto', 'Quantidade' ]]);
    (dadosRelatorio.produtosMaisPedidos||[]).forEach(item => {
      sheet.getRange(row++,1,1,2).setValues([[ item.produto, item.quantidade ]]);
    });
    row++;
    sheet.getRange(row++,1).setValue('Pedidos do mês');
    sheet.getRange(row++,1,1,3).setValues([[ 'Data', 'Cliente', 'Total' ]]);
    (dadosRelatorio.pedidos && dadosRelatorio.pedidos.mes || []).forEach(p => {
      const d = p.data;
      sheet.getRange(row++,1,1,3).setValues([[ new Date(d.data_hora), d.cliente_nome, Number(d.total||0) ]]);
    });
    // Converter para XLSX
    const file = DriveApp.getFileById(ss.getId());
    const xlsxBlob = file.getBlob().getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const xlsxFile = DriveApp.createFile(xlsxBlob);
    xlsxFile.setName('Relatorio_Cardaplan_'+new Date().toISOString().split('T')[0]+'.xlsx');
    // Remover planilha temporária
    file.setTrashed(true);
    return { success:true, fileId:xlsxFile.getId(), downloadUrl: `https://drive.google.com/file/d/${xlsxFile.getId()}/view` };
  } catch (error) {
    throw new Error('Erro ao gerar XLSX: '+error.message);
  }
}