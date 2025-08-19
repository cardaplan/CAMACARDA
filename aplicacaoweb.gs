/**
 * Cardaplan CRM - Backend Adaptado para Planilha Específica
 * Integração com Google Sheets ID: 1C_LUe0OdhcmPGqC5qs9AU9tFQR_PfGLndlcoeNqB6wQ
 * 
 * Abas da planilha:
 * - Itens: SKU, item, categoria, descricao, preco, status, classificacao_adicional, observacoes, foto_url
 * - Categorias: nome_categoria, titulo_exibicao, descricao, ordem, status
 * - Config: section, key, value
 * - Horários: Dia da Semana, Período 1, Período 2, Período 3
 * - Bairros: nome_bairro, valor_taxa
 * - Cupons: codigo_cupom, tipo_desconto, valor_desconto, data_inicio, data_fim
 * - Analytics: data_acesso, tipo_evento, item_id, sessao_id, user_agent, ip_address
 */

// Configuração da planilha
const SPREADSHEET_ID = '1oQQ6xq8OFTupdo63MTh4ZheQ13lGUMg2I4ggGMRPWnE';

// Mapeamento das abas e seus campos
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
      // ===== Novos campos (Controle de Estoque e ID alternado) =====
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
      { name: 'data_fim', type: 'dd/mm/aaaa', required: true }
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
  }
};

// Cache para otimização de performance
const cache = {
  data: {},
  timestamp: {},
  TTL: 5 * 60 * 1000 // 5 minutos
};

// Nomes de abas auxiliares
const SHEET_PEDIDOS = 'Pedidos';
const SHEET_RELATORIOS = 'Relatórios';
const SHEET_ESTOQUE_HIST = 'Estoque_Historico';

// Status padrão de pedidos
const PEDIDO_STATUS = {
  PENDENTE: 'Pendente',
  PAGO: 'Pedido pago ✅',
  CANCELADO: 'Pedido cancelado ❌'
};

/**
 * Obter referência de planilha por nome (cria se não existir)
 */
function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Garantir cabeçalho das abas criadas dinamicamente
 */
function ensureSheetHeaders() {
  // Pedidos
  const pedidos = getOrCreateSheet(SHEET_PEDIDOS);
  if (pedidos.getLastRow() === 0) {
    pedidos.appendRow([
      'order_id', 'timestamp', 'status_pedido', 'cliente_nome', 'whatsapp_cliente', 'delivery_tipo',
      'endereco_rua', 'endereco_numero', 'endereco_bairro', 'endereco_cidade', 'endereco_estado', 'endereco_cep', 'endereco_complemento',
      'pagamento_metodo', 'troco_para', 'observacoes',
      'subtotal_produtos', 'taxa_entrega', 'desconto', 'total_final',
      'itens_json', 'cupom_codigo', 'cupom_tipo'
    ]);
  }
  // Relatórios
  const rel = getOrCreateSheet(SHEET_RELATORIOS);
  if (rel.getLastRow() === 0) {
    rel.appendRow(['mes', 'produto_sku', 'produto_nome', 'quantidade', 'valor_total']);
  }
  // Histórico de Estoque
  const hist = getOrCreateSheet(SHEET_ESTOQUE_HIST);
  if (hist.getLastRow() === 0) {
    hist.appendRow(['timestamp', 'sku', 'item', 'movimento', 'quantidade', 'saldo', 'motivo', 'order_id']);
  }
}

/**
 * Gera um ID de pedido único
 */
function generateOrderId() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const rand = Utilities.getUuid().split('-')[0].toUpperCase();
  return `ORD-${ts}-${rand}`;
}

/**
 * Gera ID alternado rotativo por SKU (ex.: M05, A56)
 * Usa PropertiesService para manter o contador por SKU.
 */
function generateAlternateIdForSku(sku) {
  const props = PropertiesService.getScriptProperties();
  const key = `alt_id_counter::${sku}`;
  let counter = parseInt(props.getProperty(key) || '0', 10);
  counter = (counter + 1) % 100; // 00..99
  props.setProperty(key, String(counter));
  const letters = ['M', 'A', 'U', 'B', 'K', 'Z', 'R', 'T'];
  const letter = letters[counter % letters.length];
  const num = (counter < 10 ? '0' : '') + counter;
  return `${letter}${num}`;
}

/**
 * Atualiza estoque de acordo com itens do pedido e registra histórico
 */
function updateStockForOrder(orderId, items) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const itensSheet = ss.getSheetByName('Itens');
  if (!itensSheet) return;
  const lastRow = itensSheet.getLastRow();
  const lastCol = itensSheet.getLastColumn();
  if (lastRow < 2) return;
  const values = itensSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0];
  const idx = {
    sku: headers.indexOf('SKU'),
    item: headers.indexOf('item'),
    status: headers.indexOf('status'),
    estoqueInicial: headers.indexOf('estoque_inicial'),
    estoqueAtual: headers.indexOf('estoque_atual'),
    controleEstoque: headers.indexOf('controle_estoque')
  };
  const hist = getOrCreateSheet(SHEET_ESTOQUE_HIST);

  // Mapa SKU -> {qty}
  const skuToQty = {};
  (items || []).forEach(it => {
    const sku = it.sku || it.SKU;
    const q = parseFloat(it.quantity || it.qtd || 0) || 0;
    if (!sku) return;
    skuToQty[sku] = (skuToQty[sku] || 0) + q;
  });

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const sku = idx.sku >= 0 ? row[idx.sku] : '';
    if (!sku || !(sku in skuToQty)) continue;
    const wantsControl = idx.controleEstoque >= 0 ? String(row[idx.controleEstoque] || '').toLowerCase() === 'sim' : false;
    if (!wantsControl) continue;
    const qty = skuToQty[sku];
    const atual = idx.estoqueAtual >= 0 ? (parseFloat(row[idx.estoqueAtual]) || 0) : 0;
    let novo = Math.max(0, atual - qty);
    if (idx.estoqueAtual >= 0) {
      itensSheet.getRange(r + 1, idx.estoqueAtual + 1).setValue(novo);
    }
    // Se zerou, marcar como Esgotado
    if (novo === 0 && idx.status >= 0) {
      itensSheet.getRange(r + 1, idx.status + 1).setValue('Esgotado');
    }
    // Registrar histórico
    const nome = idx.item >= 0 ? row[idx.item] : '';
    hist.appendRow([new Date(), sku, nome, 'Saída', -qty, novo, 'Pedido', orderId]);
  }
  // invalidar cache da aba Itens
  delete cache.data['Itens'];
  delete cache.timestamp['Itens'];
}

/**
 * Registra pedido recebido (via webhook/front-end) na aba Pedidos e aplica regras
 * @param {Object} payload JSON com dados do pedido
 */
function registerOrderFromWebhook(payload) {
  ensureSheetHeaders();
  try {
    if (!payload || typeof payload !== 'object') {
      throw new Error('Payload inválido');
    }
    const pedidos = getOrCreateSheet(SHEET_PEDIDOS);
    const orderId = generateOrderId();

    // Extrações seguras
    const customer = payload.customer || {};
    const items = payload.items || [];
    const totals = payload.totals || {};
    const coupon = payload.coupon || null;

    // Gerar IDs alternados por item (se habilitado no cadastro do item)
    const itensData = getSheetData('Itens');
    const headers = itensData[0] || [];
    const idxMap = {
      sku: headers.indexOf('SKU'),
      usarAlt: headers.indexOf('usar_id_alternado')
    };
    const altIdBySku = {};
    if (itensData.length > 1 && idxMap.sku >= 0) {
      const map = new Map();
      itensData.slice(1).forEach(row => {
        const sku = row[idxMap.sku];
        const usarAlt = idxMap.usarAlt >= 0 ? String(row[idxMap.usarAlt] || '').toLowerCase() === 'sim' : false;
        map.set(sku, usarAlt);
      });
      (items || []).forEach(it => {
        const sku = it.sku || it.SKU;
        if (sku && map.get(sku)) {
          altIdBySku[sku] = generateAlternateIdForSku(sku);
        }
      });
    }

    // Calcular totais defensivamente
    const subtotal = Number(totals.subtotal_products_numeric || payload.subtotal || 0) || 0;
    const delivery = Number(totals.delivery_fee_numeric || payload.delivery_fee || 0) || 0;
    const discount = Number((payload.coupon_applied ? (payload.coupon?.discount_numeric || 0) : 0)) || 0;
    const total = Number(totals.total_final_numeric || payload.total || subtotal + delivery - discount) || 0;

    // Linhas para planilha Pedidos
    const row = [
      orderId,
      new Date(),
      PEDIDO_STATUS.PENDENTE,
      payload.customer_name || customer.name || '',
      payload.whatsapp_customer || '',
      customer.delivery_type || payload.delivery_type || '',
      (customer.address && customer.address.street) || '',
      (customer.address && customer.address.number) || '',
      (customer.address && customer.address.neighborhood) || '',
      (customer.address && customer.address.city) || '',
      (customer.address && customer.address.state) || '',
      (customer.address && customer.address.cep) || '',
      (customer.address && customer.address.complement) || '',
      customer.payment_method || payload.payment_method || '',
      customer.change_amount || payload.change_amount || '',
      customer.notes || payload.notes || '',
      subtotal,
      delivery,
      discount,
      total,
      JSON.stringify(items.map(it => ({
        id: it.id,
        nome: it.name || it.nome,
        sku: it.sku,
        quantity: it.quantity,
        price: it.price_numeric || it.price || 0,
        alt_id: altIdBySku[it.sku] || null
      }))),
      coupon ? (coupon.code || '') : '',
      coupon ? (coupon.type || '') : ''
    ];
    pedidos.appendRow(row);

    // Atualizar estoque e histórico
    updateStockForOrder(orderId, items);

    // Resposta
    return { success: true, orderId, altIds: altIdBySku };
  } catch (error) {
    Logger.log('[registerOrderFromWebhook] ERRO: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Lista pedidos (mais recentes primeiro)
 */
function getOrders() {
  ensureSheetHeaders();
  const sheet = getOrCreateSheet(SHEET_PEDIDOS);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0];
  const rows = values.slice(1).reverse();
  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

/**
 * Atualiza status do pedido por order_id
 */
function updateOrderStatus(orderId, newStatus) {
  ensureSheetHeaders();
  const allowed = [PEDIDO_STATUS.PENDENTE, PEDIDO_STATUS.PAGO, PEDIDO_STATUS.CANCELADO];
  if (!allowed.includes(newStatus)) throw new Error('Status inválido');
  const sheet = getOrCreateSheet(SHEET_PEDIDOS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxOrderId = headers.indexOf('order_id');
  const idxStatus = headers.indexOf('status_pedido');
  for (let r = 1; r < values.length; r++) {
    if (values[r][idxOrderId] === orderId) {
      sheet.getRange(r + 1, idxStatus + 1).setValue(newStatus);
      return true;
    }
  }
  return false;
}

/**
 * Agrega vendas apenas de pedidos pagos
 */
function getSalesReports() {
  ensureSheetHeaders();
  const pedidos = getOrCreateSheet(SHEET_PEDIDOS);
  const values = pedidos.getDataRange().getValues();
  if (values.length < 2) {
    return {
      produtosMaisPedidos: [],
      valorTotalVendidoMes: 0,
      pedidosDia: [], pedidosSemana: [], pedidosMes: []
    };
  }
  const headers = values[0];
  const idx = {
    status: headers.indexOf('status_pedido'),
    total: headers.indexOf('total_final'),
    timestamp: headers.indexOf('timestamp'),
    itensJson: headers.indexOf('itens_json')
  };
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfWeek = new Date(startOfDay); startOfWeek.setDate(startOfDay.getDate() - startOfDay.getDay());
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  let totalMonth = 0;
  const itemsCount = new Map(); // sku -> {nome, quantidade}
  const pedidosDia = [];
  const pedidosSemana = [];
  const pedidosMes = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[idx.status] !== PEDIDO_STATUS.PAGO) continue;
    const ts = new Date(row[idx.timestamp]);
    const total = Number(row[idx.total] || 0) || 0;
    if (ts >= startOfMonth) totalMonth += total;
    if (ts >= startOfDay) pedidosDia.push({ data: ts, total });
    if (ts >= startOfWeek) pedidosSemana.push({ data: ts, total });
    if (ts >= startOfMonth) pedidosMes.push({ data: ts, total });
    try {
      const itens = JSON.parse(row[idx.itensJson] || '[]');
      itens.forEach(it => {
        const key = it.sku || it.SKU || it.id || it.nome;
        const curr = itemsCount.get(key) || { nome: it.nome || it.name || key, quantidade: 0 };
        curr.quantidade += Number(it.quantity || 0) || 0;
        itemsCount.set(key, curr);
      });
    } catch (_) {}
  }
  const produtosMaisPedidos = Array.from(itemsCount.entries())
    .map(([sku, data]) => ({ sku, nome: data.nome, quantidade: data.quantidade }))
    .sort((a, b) => b.quantidade - a.quantidade)
    .slice(10);

  return {
    produtosMaisPedidos,
    valorTotalVendidoMes: totalMonth,
    pedidosDia,
    pedidosSemana,
    pedidosMes
  };
}

/**
 * Gera PDF simples de vendas a partir de getSalesReports
 */
function gerarRelatorioVendasPDF() {
  const dados = getSalesReports();
  const doc = DocumentApp.create('Relatório de Vendas - ' + new Date().toLocaleDateString('pt-BR'));
  const body = doc.getBody();
  body.appendParagraph('RELATÓRIO DE VENDAS').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Gerado em: ' + new Date().toLocaleString('pt-BR'));
  body.appendParagraph('');
  body.appendParagraph('Total vendido no mês: ' + dados.valorTotalVendidoMes.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }));
  body.appendParagraph('');
  body.appendParagraph('Produtos mais pedidos').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  dados.produtosMaisPedidos.forEach((p, idx) => body.appendParagraph(`${idx + 1}. ${p.nome} (${p.sku}) - ${p.quantidade} un`));
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getBlob().getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdf).setName('Relatorio_Vendas_' + new Date().toISOString().slice(0, 10) + '.pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return { success: true, fileId: pdfFile.getId(), downloadUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/view` };
}

/**
 * Função principal para servir a aplicação web
 */
function doGet(e) {
  try {
    Logger.log('[doGet] Iniciando aplicação web');
    // Suporte a endpoint de registro de pedidos via querystring (sem quebrar UI)
    const request = e || null;
    // Nota: em Apps Script, doGet(e) recebe parâmetros; aqui mantemos compatibilidade com doGet() antigo
    // Para ativar este modo, altere a assinatura para function doGet(e) se necessário no deploy.
    // Abaixo, fallback: tenta acessar parâmetros globais do serviço se e não estiver disponível.
    try {
      // @ts-ignore
      if (request && request.parameter) {
        const params = request.parameter || {};
        if (params.action === 'register_order') {
          const data = params.data ? JSON.parse(params.data) : null;
          const result = registerOrderFromWebhook(data);
          return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
        }
      }
    } catch (_ignore) {}
    
    const htmlOutput = HtmlService.createHtmlOutputFromFile("Index")
      .setTitle("Cardaplan - Gestão Inteligente")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    Logger.log('[doGet] Aplicação web carregada com sucesso');
    return htmlOutput;
    
  } catch (error) {
    Logger.log(`[doGet] ERRO: ${error.message}`);
    console.error("Erro ao carregar página:", error);
    return HtmlService.createHtmlOutput("<h1>Erro ao carregar a aplicação</h1><p>Tente novamente em alguns instantes.</p>");
  }
}

/**
 * Obter nomes das abas da planilha
 */
function getSheetNames() {
  try {
    Logger.log('[getSheetNames] Obtendo nomes das abas');
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    Logger.log(`[getSheetNames] Abas encontradas: ${sheetNames.join(', ')}`);
    return sheetNames;
    
  } catch (error) {
    Logger.log(`[getSheetNames] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter nomes das abas: ${error.message}`);
  }
}

/**
 * Obter configuração de uma aba específica
 */
function getSheetConfig(sheetName) {
  try {
    Logger.log(`[getSheetConfig] Obtendo configuração para: ${sheetName}`);
    
    if (!SHEET_CONFIG[sheetName]) {
      throw new Error(`Configuração não encontrada para a aba: ${sheetName}`);
    }
    
    return SHEET_CONFIG[sheetName];
    
  } catch (error) {
    Logger.log(`[getSheetConfig] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter configuração da aba: ${error.message}`);
  }
}

/**
 * Obter todas as configurações das abas
 */
function getAllSheetConfigs() {
  try {
    Logger.log('[getAllSheetConfigs] Obtendo todas as configurações');
    return SHEET_CONFIG;
  } catch (error) {
    Logger.log(`[getAllSheetConfigs] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter configurações: ${error.message}`);
  }
}

/**
 * Verificar se o cache é válido
 */
function isCacheValid(key) {
  if (!cache.timestamp[key]) return false;
  return (Date.now() - cache.timestamp[key]) < cache.TTL;
}

/**
 * Obter dados de uma aba específica
 */
function getSheetData(sheetName) {
  try {
    Logger.log(`[getSheetData] Iniciando para planilha: ${sheetName}`);
    
    // Verificar cache primeiro
    if (isCacheValid(sheetName) && cache.data[sheetName]) {
      Logger.log(`[getSheetData] Dados obtidos do cache para: ${sheetName}`);
      return cache.data[sheetName];
    }
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada na planilha`);
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0 || lastColumn === 0) {
      Logger.log(`[getSheetData] Aba "${sheetName}" está vazia`);
      return [];
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const values = range.getValues();
    
    // Filtrar linhas completamente vazias
    const filteredValues = values.filter(row => 
      row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== '')
    );
    
    // Atualizar cache
    cache.data[sheetName] = filteredValues;
    cache.timestamp[sheetName] = Date.now();
    
    Logger.log(`[getSheetData] Dados obtidos com sucesso para "${sheetName}": ${filteredValues.length} linhas`);
    return filteredValues;
    
  } catch (error) {
    Logger.log(`[getSheetData] ERRO para "${sheetName}": ${error.message}`);
    console.error(`Erro ao obter dados da planilha "${sheetName}":`, error);
    throw new Error(`Erro ao obter dados da planilha "${sheetName}": ${error.message}`);
  }
}

/**
 * Obter dados de todas as abas
 */
function getAllSheetsData() {
  try {
    Logger.log('[getAllSheetsData] Obtendo dados de todas as abas');
    
    const result = {};
    const sheetNames = Object.keys(SHEET_CONFIG);
    
    for (const sheetName of sheetNames) {
      try {
        if (sheetName === 'Horários') {
          // Para a aba 'Horários', obter dados a partir da linha 3
          result[sheetName] = getSheetData(sheetName, 2); // Assumindo que getSheetData pode receber um parâmetro de linha inicial
        } else {
          result[sheetName] = getSheetData(sheetName);
        }
        Logger.log(`[getAllSheetsData] Dados obtidos para "${sheetName}": ${result[sheetName].length} linhas`);
      } catch (error) {
        Logger.log(`[getAllSheetsData] Erro ao obter dados de "${sheetName}": ${error.message}`);
        result[sheetName] = [];
      }
    }
    
    return result;
    
  } catch (error) {
    Logger.log(`[getAllSheetsData] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter dados de todas as abas: ${error.message}`);
  }
}

/**
 * Adicionar nova linha a uma aba
 */
function addRow(sheetName, rowData) {
  try {
    Logger.log(`[addRow] Iniciando para planilha: ${sheetName}`);
    Logger.log(`[addRow] Dados recebidos: ${JSON.stringify(rowData)}`);
    
    if (!SHEET_CONFIG[sheetName]) {
      throw new Error(`Configuração não encontrada para a aba: ${sheetName}`);
    }
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada na planilha`);
    }
    
    // Validar e normalizar dados
    const config = SHEET_CONFIG[sheetName];
    const normalizedRowData = [];
    
    for (let i = 0; i < config.fields.length; i++) {
      const field = config.fields[i];
      let value = rowData[i] || '';
      
      // Validações específicas por tipo
      if (field.required && (!value || value.toString().trim() === '')) {
        throw new Error(`Campo obrigatório "${field.name}" não pode estar vazio`);
      }
      
      // Normalização por tipo
      switch (field.type) {
        case 'number':
        case 'currency':
          if (value && value !== '') {
            value = parseFloat(value.toString().replace(/[^\d.,]/g, '').replace(',', '.'));
            if (isNaN(value)) value = 0;
          }
          break;
        case 'percentage':
          if (value && value !== '') {
            // Aceita valores como "10%", "20", "15.5%"
            let numValue = value.toString().replace('%', '').replace(',', '.');
            numValue = parseFloat(numValue);
            if (isNaN(numValue)) numValue = 0;
            value = numValue + '%';
          }
          break;
        case 'date':
        case 'date':
          if (value && value !== '') {
            try {
              if (typeof value === 'string' && value.includes('T')) {
                value = new Date(value);
              } else if (typeof value === 'string') {
                const dateParts = value.split('-');
                if (dateParts.length === 3) {
                  value = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                } else {
                  value = new Date(value);
                }
              } else {
                value = new Date(value);
              }
              
              if (isNaN(value.getTime())) {
                Logger.log(`[addRow] Data inválida recebida: ${rowData[i]}`);
                value = '';
              }
            } catch (e) {
              Logger.log(`[addRow] Erro ao processar data: ${e.message}`);
              value = '';
            }
          }
          break;
        case 'datetime':
          if (value && value !== '') {
            try {
              value = new Date(value);
              if (isNaN(value.getTime())) {
                value = new Date();
              }
            } catch (e) {
              value = new Date();
            }
          } else {
            value = new Date();
          }
          break;
        case 'url':
          if (value && value !== '' && !value.startsWith('http')) {
            value = 'https://' + value;
          }
          break;
      }
      
      normalizedRowData.push(value);
    }
    
    // Adicionar linha à planilha
    sheet.appendRow(normalizedRowData);
    
    const newRowNumber = sheet.getLastRow();
    Logger.log(`[addRow] Linha adicionada com sucesso na posição: ${newRowNumber}`);
    
    // Invalidar cache
    delete cache.data[sheetName];
    delete cache.timestamp[sheetName];
    
    return true;
    
  } catch (error) {
    Logger.log(`[addRow] ERRO: ${error.message}`);
    console.error(`Erro ao adicionar linha na planilha "${sheetName}":`, error);
    throw new Error(`Erro ao adicionar linha: ${error.message}`);
  }
}

/**
 * Atualizar linha existente
 */
function updateRow(sheetName, rowIndex, rowData) {
  try {
    Logger.log(`[updateRow] Iniciando para planilha: ${sheetName}, linha: ${rowIndex}`);
    Logger.log(`[updateRow] Dados recebidos: ${JSON.stringify(rowData)}`);
    
    if (!SHEET_CONFIG[sheetName]) {
      throw new Error(`Configuração não encontrada para a aba: ${sheetName}`);
    }
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada na planilha`);
    }
    
    if (rowIndex < 1 || rowIndex > sheet.getLastRow()) {
      throw new Error(`Índice de linha inválido: ${rowIndex}`);
    }
    
    // Validar e normalizar dados (mesmo processo do addRow)
    const config = SHEET_CONFIG[sheetName];
    const normalizedRowData = [];
    
    for (let i = 0; i < config.fields.length; i++) {
      const field = config.fields[i];
      let value = rowData[i] || '';
      
      if (field.required && (!value || value.toString().trim() === '')) {
        throw new Error(`Campo obrigatório "${field.name}" não pode estar vazio`);
      }
      
      switch (field.type) {
        case 'number':
        case 'currency':
          if (value && value !== '') {
            value = parseFloat(value.toString().replace(/[^\d.,]/g, '').replace(',', '.'));
            if (isNaN(value)) value = 0;
          }
          break;
        case 'percentage':
          if (value && value !== '') {
            let numValue = value.toString().replace('%', '').replace(',', '.');
            numValue = parseFloat(numValue);
            if (isNaN(numValue)) numValue = 0;
            value = numValue + '%';
          }
          break;
        case 'date':
        case 'date':
          if (value && value !== '') {
            try {
              if (typeof value === 'string' && value.includes('T')) {
                value = new Date(value);
              } else if (typeof value === 'string') {
                const dateParts = value.split('-');
                if (dateParts.length === 3) {
                  value = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                } else {
                  value = new Date(value);
                }
              } else {
                value = new Date(value);
              }
              
              if (isNaN(value.getTime())) {
                Logger.log(`[updateRow] Data inválida recebida: ${rowData[i]}`);
                value = '';
              }
            } catch (e) {
              Logger.log(`[updateRow] Erro ao processar data: ${e.message}`);
              value = '';
            }
          }
          break;
        case 'datetime':
          if (value && value !== '') {
            try {
              value = new Date(value);
              if (isNaN(value.getTime())) {
                value = new Date();
              }
            } catch (e) {
              value = new Date();
            }
          } else {
            value = new Date();
          }
          break;
        case 'url':
          if (value && value !== '' && !value.startsWith('http')) {
            value = 'https://' + value;
          }
          break;
      }
      
      normalizedRowData.push(value);
    }
    
    // Atualizar linha na planilha
    const range = sheet.getRange(rowIndex, 1, 1, normalizedRowData.length);
    range.setValues([normalizedRowData]);
    
    Logger.log(`[updateRow] Linha ${rowIndex} atualizada com sucesso`);
    
    // Invalidar cache
    delete cache.data[sheetName];
    delete cache.timestamp[sheetName];
    
    return true;
    
  } catch (error) {
    Logger.log(`[updateRow] ERRO: ${error.message}`);
    console.error(`Erro ao atualizar linha ${rowIndex} na planilha "${sheetName}":`, error);
    throw new Error(`Erro ao atualizar linha: ${error.message}`);
  }
}

/**
 * Excluir linha
 */
function deleteRow(sheetName, rowIndex) {
  try {
    Logger.log(`[deleteRow] Iniciando para planilha: ${sheetName}, linha: ${rowIndex}`);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Aba "${sheetName}" não encontrada na planilha`);
    }
    
    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error(`Índice de linha inválido: ${rowIndex}. Não é possível excluir o cabeçalho.`);
    }
    
    sheet.deleteRow(rowIndex);
    Logger.log(`[deleteRow] Linha ${rowIndex} excluída com sucesso`);
    
    // Invalidar cache
    delete cache.data[sheetName];
    delete cache.timestamp[sheetName];
    
    return true;
    
  } catch (error) {
    Logger.log(`[deleteRow] ERRO: ${error.message}`);
    console.error(`Erro ao excluir linha ${rowIndex} na planilha "${sheetName}":`, error);
    throw new Error(`Erro ao excluir linha: ${error.message}`);
  }
}

/**
 * Buscar dados em todas as abas
 */
function searchAllSheets(query) {
  try {
    Logger.log(`[searchAllSheets] Iniciando busca por: "${query}"`);
    
    if (!query || query.trim() === '') {
      return [];
    }
    
    const searchQuery = query.toLowerCase();
    const results = [];
    
    // Buscar em todas as abas configuradas
    for (const [sheetName, config] of Object.entries(SHEET_CONFIG)) {
      try {
        const data = getSheetData(sheetName);
        
        if (data.length <= 1) continue; // Pular se não há dados além do cabeçalho
        
        const headers = data[0];
        const rows = data.slice(1);
        
        rows.forEach((row, rowIndex) => {
          const rowText = row.map(cell => String(cell)).join(' ').toLowerCase();
          if (rowText.includes(searchQuery)) {
            results.push({
              sheet: sheetName,
              rowIndex: rowIndex + 2, // +2 para considerar cabeçalho e índice 1-based
              data: row
            });
          }
        });
      } catch (error) {
        Logger.log(`[searchAllSheets] Erro ao buscar em "${sheetName}": ${error.message}`);
      }
    }
    
    Logger.log(`[searchAllSheets] Busca concluída. ${results.length} resultados encontrados.`);
    return results;
    
  } catch (error) {
    Logger.log(`[searchAllSheets] ERRO: ${error.message}`);
    console.error("Erro na busca global:", error);
    throw new Error(`Erro na busca global: ${error.message}`);
  }
}

/**
 * Obter estatísticas do Dashboard
 */
function getDashboardStats() {
  try {
    Logger.log('[getDashboardStats] Obtendo estatísticas do dashboard');
    
    const stats = {
      totalItens: 0,
      itensAtivos: 0,
      totalCategorias: 0,
      totalCupons: 0,
      cuponsAtivos: 0,
      totalBairros: 0,
      totalHorarios: 0,
      totalVisualizacoes: 0,
      sessesUnicas: 0,
      visualizacoesHoje: 0,
      // ===== KPIs adicionais
      totalPedidosHoje: 0,
      ticketMedio: 0,
      produtoMaisVendido: '',
      totalClientesUnicos: 0
    };

    // Itens
    const itensData = getSheetData('Itens');
    if (itensData.length > 1) {
      stats.totalItens = itensData.length - 1;
      stats.itensAtivos = itensData.slice(1).filter(row => row[5] === 'Ativo').length;
    }

    // Categorias
    const categoriasData = getSheetData('Categorias');
    if (categoriasData.length > 1) {
      stats.totalCategorias = categoriasData.length - 1;
    }

    // Cupons
    const cuponsData = getSheetData('Cupons');
    if (cuponsData.length > 1) {
      stats.totalCupons = cuponsData.length - 1;
      const hoje = new Date();
      stats.cuponsAtivos = cuponsData.slice(1).filter(row => {
      const [di, mi, ai] = (row[3] || '').split('/');
      const [df, mf, af] = (row[4] || '').split('/');

    if (di && mi && ai && df && mf && af) {
      const inicio = new Date(`${ai}-${mi}-${di}T00:00:00-03:00`);
      const fim = new Date(`${af}-${mf}-${df}T23:59:59-03:00`);
      return hoje >= inicio && hoje <= fim;
  }

  return false;
}).length;

    }

    // Bairros
    const bairrosData = getSheetData('Bairros');
    if (bairrosData.length > 1) {
      stats.totalBairros = bairrosData.length - 1;
    }

    // Horários
    const horariosData = getSheetData('Horários').slice(2);
    if (horariosData.length > 1) {
      stats.totalHorarios = horariosData.length - 1;
    }

    // Analytics
    try {
      const analyticsData = getSheetData('Analytics');
      if (analyticsData.length > 1) {
        const rows = analyticsData.slice(1);
        
        // Total de visualizações do cardápio
        stats.totalVisualizacoes = rows.filter(row => row[1] === 'visualizacao_cardapio').length;
        
        // Sessões únicas
        const sessoes = new Set();
        rows.forEach(row => {
          if (row[3]) sessoes.add(row[3]); // sessao_id
        });
        stats.sessesUnicas = sessoes.size;
        
        // Visualizações hoje
        const hoje = new Date();
        const hojeStr = hoje.toDateString();
        stats.visualizacoesHoje = rows.filter(row => {
          if (row[0]) {
            const dataAcesso = new Date(row[0]);
            return dataAcesso.toDateString() === hojeStr && row[1] === 'visualizacao_cardapio';
          }
          return false;
        }).length;
      }
    } catch (error) {
      Logger.log(`[getDashboardStats] Erro ao obter analytics: ${error.message}`);
    }
    
    // KPIs de Pedidos (apenas pagos)
    try {
      ensureSheetHeaders();
      const pedidos = getOrCreateSheet(SHEET_PEDIDOS);
      const values = pedidos.getDataRange().getValues();
      if (values.length > 1) {
        const headers = values[0];
        const idx = {
          status: headers.indexOf('status_pedido'),
          total: headers.indexOf('total_final'),
          ts: headers.indexOf('timestamp'),
          itens: headers.indexOf('itens_json'),
          cliente: headers.indexOf('whatsapp_cliente')
        };
        const hojeStr = new Date().toDateString();
        let somaMes = 0, qtdMes = 0;
        const inicioMes = new Date(); inicioMes.setDate(1); inicioMes.setHours(0,0,0,0);
        const mapVendas = new Map();
        const clientes = new Set();
        for (let r = 1; r < values.length; r++) {
          const row = values[r];
          if (row[idx.status] !== PEDIDO_STATUS.PAGO) continue;
          const ts = new Date(row[idx.ts]);
          const total = Number(row[idx.total] || 0) || 0;
          if (ts.toDateString() === hojeStr) stats.totalPedidosHoje++;
          if (ts >= inicioMes) { somaMes += total; qtdMes++; }
          try {
            const itens = JSON.parse(row[idx.itens] || '[]');
            itens.forEach(it => {
              const key = it.sku || it.id || it.nome;
              const prev = mapVendas.get(key) || { nome: it.nome || key, qtd: 0 };
              prev.qtd += Number(it.quantity || 0) || 0;
              mapVendas.set(key, prev);
            });
          } catch (_) {}
          const cli = row[idx.cliente]; if (cli) clientes.add(cli);
        }
        stats.ticketMedio = qtdMes > 0 ? (somaMes / qtdMes) : 0;
        const top = Array.from(mapVendas.values()).sort((a,b)=>b.qtd-a.qtd)[0];
        stats.produtoMaisVendido = top ? top.nome : '';
        stats.totalClientesUnicos = clientes.size;
      }
    } catch (e) { Logger.log('[getDashboardStats] Erro KPIs pedidos: ' + e.message); }
    
    Logger.log('[getDashboardStats] Estatísticas obtidas:', stats);
    return stats;
    
  } catch (error) {
    Logger.log(`[getDashboardStats] ERRO: ${error.message}`);
    console.error("Erro ao obter estatísticas do dashboard:", error);
    throw new Error(`Erro ao obter estatísticas do dashboard: ${error.message}`);
  }
}

/**
 * Funções específicas para Bairros
 */

/**
 * Obter lista de bairros com suas taxas
 */
function getBairrosList() {
  try {
    Logger.log('[getBairrosList] Obtendo lista de bairros');
    
    const data = getSheetData('Bairros');
    if (data.length <= 1) {
      return [];
    }
    
    const bairros = data.slice(1).map((row, index) => ({
      id: index + 2, // Row index (1-based + header)
      nome: row[0] || '',
      taxa: parseFloat(row[1] || 0)
    }));
    
    Logger.log(`[getBairrosList] ${bairros.length} bairros encontrados`);
    return bairros;
    
  } catch (error) {
    Logger.log(`[getBairrosList] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter lista de bairros: ${error.message}`);
  }
}

/**
 * Obter taxa de entrega para um bairro específico
 */
function getTaxaEntregaBairro(nomeBairro) {
  try {
    Logger.log(`[getTaxaEntregaBairro] Buscando taxa para: ${nomeBairro}`);
    
    // Verificar se há taxa única definida na Config
    try {
      const configData = getSheetData('Config');
      const headers = (configData && configData[0]) || [];
      const idxSection = headers.indexOf('section');
      const idxKey = headers.indexOf('key');
      const idxValue = headers.indexOf('value');
      if (idxSection >= 0 && idxKey >= 0 && idxValue >= 0) {
        const valorTaxaUnicaRow = configData.slice(1).find(r => String(r[idxSection]).toLowerCase() === 'checkout' && String(r[idxKey]).toLowerCase() === 'valor_taxa_unica' && String(r[idxValue]).trim() !== '');
        if (valorTaxaUnicaRow) {
          const taxaUnica = parseFloat(String(valorTaxaUnicaRow[idxValue]).replace(/[^0-9.,]/g, '').replace(',', '.'));
          if (!isNaN(taxaUnica) && taxaUnica >= 0) {
            Logger.log(`[getTaxaEntregaBairro] Aplicando taxa única: ${taxaUnica}`);
            return taxaUnica;
          }
        }
      }
    } catch (e) {
      Logger.log(`[getTaxaEntregaBairro] Aviso ao ler taxa única: ${e.message}`);
    }

    const data = getSheetData('Bairros');
    if (data.length <= 1) {
      return null;
    }
    
    const bairro = data.slice(1).find(row => 
      row[0] && row[0].toLowerCase() === nomeBairro.toLowerCase()
    );
    
    if (bairro) {
      const taxa = parseFloat(bairro[1] || 0);
      Logger.log(`[getTaxaEntregaBairro] Taxa encontrada: R$ ${taxa.toFixed(2)}`);
      return taxa;
    }
    
    Logger.log(`[getTaxaEntregaBairro] Bairro não encontrado: ${nomeBairro}`);
    return null;
    
  } catch (error) {
    Logger.log(`[getTaxaEntregaBairro] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter taxa de entrega: ${error.message}`);
  }
}

/**
 * Adicionar novo bairro
 */
function addBairro(nomeBairro, valorTaxa) {
  try {
    Logger.log(`[addBairro] Adicionando bairro: ${nomeBairro}, taxa: ${valorTaxa}`);
    
    // Validar se o bairro já existe
    const existingTaxa = getTaxaEntregaBairro(nomeBairro);
    if (existingTaxa !== null) {
      throw new Error(`Bairro "${nomeBairro}" já existe na lista`);
    }
    
    // Validar dados
    if (!nomeBairro || nomeBairro.trim() === '') {
      throw new Error('Nome do bairro é obrigatório');
    }
    
    const taxa = parseFloat(valorTaxa);
    if (isNaN(taxa) || taxa < 0) {
      throw new Error('Taxa de entrega deve ser um valor numérico válido');
    }
    
    // Adicionar à planilha
    const rowData = [nomeBairro.trim(), taxa];
    return addRow('Bairros', rowData);
    
  } catch (error) {
    Logger.log(`[addBairro] ERRO: ${error.message}`);
    throw new Error(`Erro ao adicionar bairro: ${error.message}`);
  }
}

/**
 * Atualizar bairro existente
 */
function updateBairro(rowIndex, nomeBairro, valorTaxa) {
  try {
    Logger.log(`[updateBairro] Atualizando bairro linha ${rowIndex}: ${nomeBairro}, taxa: ${valorTaxa}`);
    
    // Validar dados
    if (!nomeBairro || nomeBairro.trim() === '') {
      throw new Error('Nome do bairro é obrigatório');
    }
    
    const taxa = parseFloat(valorTaxa);
    if (isNaN(taxa) || taxa < 0) {
      throw new Error('Taxa de entrega deve ser um valor numérico válido');
    }
    
    // Atualizar na planilha
    const rowData = [nomeBairro.trim(), taxa];
    return updateRow('Bairros', rowIndex, rowData);
    
  } catch (error) {
    Logger.log(`[updateBairro] ERRO: ${error.message}`);
    throw new Error(`Erro ao atualizar bairro: ${error.message}`);
  }
}

/**
 * Excluir bairro
 */
function deleteBairro(rowIndex) {
  try {
    Logger.log(`[deleteBairro] Excluindo bairro linha: ${rowIndex}`);
    
    return deleteRow('Bairros', rowIndex);
    
  } catch (error) {
    Logger.log(`[deleteBairro] ERRO: ${error.message}`);
    throw new Error(`Erro ao excluir bairro: ${error.message}`);
  }
}

/**
 * Buscar bairros por nome
 */
function searchBairros(query) {
  try {
    Logger.log(`[searchBairros] Buscando bairros por: "${query}"`);
    
    if (!query || query.trim() === '') {
      return getBairrosList();
    }
    
    const searchQuery = query.toLowerCase();
    const data = getSheetData('Bairros');
    
    if (data.length <= 1) {
      return [];
    }
    
    const results = data.slice(1)
      .map((row, index) => ({
        id: index + 2,
        nome: row[0] || '',
        taxa: parseFloat(row[1] || 0)
      }))
      .filter(bairro => 
        bairro.nome.toLowerCase().includes(searchQuery)
      );
    
    Logger.log(`[searchBairros] ${results.length} bairros encontrados`);
    return results;
    
  } catch (error) {
    Logger.log(`[searchBairros] ERRO: ${error.message}`);
    throw new Error(`Erro ao buscar bairros: ${error.message}`);
  }
}

/**
 * Registrar evento de analytics
 */
function registrarEventoAnalytics(tipoEvento, itemId = null, sessaoId = null, userAgent = null, ipAddress = null) {
  try {
    Logger.log(`[registrarEventoAnalytics] Registrando evento: ${tipoEvento}`);
    
    const analyticsData = [
      new Date(), // data_acesso
      tipoEvento, // tipo_evento
      itemId || '', // item_id
      sessaoId || generateSessionId(), // sessao_id
      userAgent || '', // user_agent
      ipAddress || '' // ip_address
    ];
    
    return addRow('Analytics', analyticsData);
    
  } catch (error) {
    Logger.log(`[registrarEventoAnalytics] ERRO: ${error.message}`);
    // Não lançar erro para não quebrar a funcionalidade principal
    return false;
  }
}

/**
 * Gerar ID de sessão único
 */
function generateSessionId() {
  return 'sess_' + Utilities.getUuid();
}

/**
 * Obter dados de relatórios analíticos
 */
function getAnalyticsReports(dataInicio = null, dataFim = null) {
  try {
    Logger.log('[getAnalyticsReports] Obtendo relatórios analíticos');
    
    const analyticsData = getSheetData('Analytics');
    if (analyticsData.length <= 1) {
      return {
        visualizacoesPorDia: [],
        produtosMaisVisualizados: [],
        sessoesPorDia: [],
        resumoGeral: {
          totalVisualizacoes: 0,
          sessesUnicas: 0,
          produtosMaisVistos: 0
        }
      };
    }
    
    const rows = analyticsData.slice(1);
    let filteredRows = rows;
    
    // Filtrar por data se especificado
    if (dataInicio || dataFim) {
      const inicio = dataInicio ? new Date(dataInicio) : new Date('01/01/2020');
      const fim = dataFim ? new Date(dataFim) : new Date();
      
      filteredRows = rows.filter(row => {
        if (row[0]) {
          const dataAcesso = new Date(row[0]);
          return dataAcesso >= inicio && dataAcesso <= fim;
        }
        return false;
      });
    }
    
    // Visualizações por dia
    const visualizacoesPorDia = {};
    const sessoesPorDia = {};
    const produtosVisualizados = {};
    const sessoes = new Set();
    
    filteredRows.forEach(row => {
      const dataAcesso = row[0] ? new Date(row[0]) : null;
      const tipoEvento = row[1];
      const itemId = row[2];
      const sessaoId = row[3];
      
      if (dataAcesso) {
        const dataStr = dataAcesso.toISOString().split('T')[0];
        
        // Visualizações por dia
        if (tipoEvento === 'visualizacao_cardapio') {
          visualizacoesPorDia[dataStr] = (visualizacoesPorDia[dataStr] || 0) + 1;
        }
        
        // Sessões por dia
        if (sessaoId) {
          if (!sessoesPorDia[dataStr]) {
            sessoesPorDia[dataStr] = new Set();
          }
          sessoesPorDia[dataStr].add(sessaoId);
          sessoes.add(sessaoId);
        }
        
        // Produtos mais visualizados
        if (tipoEvento === 'visualizacao_item' && itemId) {
          produtosVisualizados[itemId] = (produtosVisualizados[itemId] || 0) + 1;
        }
      }
    });
    
    // Converter sessões por dia para números
    const sessoesPorDiaNumeros = {};
    Object.keys(sessoesPorDia).forEach(data => {
      sessoesPorDiaNumeros[data] = sessoesPorDia[data].size;
    });
    
    // Produtos mais visualizados (top 10)
    const produtosMaisVisualizados = Object.entries(produtosVisualizados)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10)
      .map(([itemId, visualizacoes]) => ({
        itemId,
        visualizacoes,
        nomeItem: getItemName(itemId)
      }));
    
    const resultado = {
      visualizacoesPorDia: Object.entries(visualizacoesPorDia).map(([data, count]) => ({
        data,
        visualizacoes: count
      })),
      produtosMaisVisualizados,
      sessoesPorDia: Object.entries(sessoesPorDiaNumeros).map(([data, count]) => ({
        data,
        sessoes: count
      })),
      resumoGeral: {
        totalVisualizacoes: Object.values(visualizacoesPorDia).reduce((a, b) => a + b, 0),
        sessesUnicas: sessoes.size,
        produtosMaisVistos: produtosMaisVisualizados.length
      }
    };
    
    Logger.log('[getAnalyticsReports] Relatórios obtidos com sucesso');
    return resultado;
    
  } catch (error) {
    Logger.log(`[getAnalyticsReports] ERRO: ${error.message}`);
    throw new Error(`Erro ao obter relatórios: ${error.message}`);
  }
}

/**
 * Obter nome do item pelo ID
 */
function getItemName(itemId) {
  try {
    const itensData = getSheetData('Itens');
    if (itensData.length > 1) {
      const item = itensData.slice(1).find(row => row[0] === itemId); // SKU
      return item ? item[1] : itemId; // nome do item
    }
    return itemId;
  } catch (error) {
    return itemId;
  }
}

/**
 * Gerar PDF de relatório
 */
function gerarRelatorioPDF(dadosRelatorio) {
  try {
    Logger.log('[gerarRelatorioPDF] Gerando relatório em PDF');
    
    // Criar documento temporário
    const doc = DocumentApp.create('Relatório Cardaplan - ' + new Date().toLocaleDateString('pt-BR'));
    const body = doc.getBody();
    
    // Título
    const titulo = body.appendParagraph('RELATÓRIO CARDAPLAN');
    titulo.setHeading(DocumentApp.ParagraphHeading.TITLE);
    titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Data de geração
    body.appendParagraph('Gerado em: ' + new Date().toLocaleString('pt-BR'));
    body.appendParagraph('');
    
    // Resumo geral
    body.appendParagraph('RESUMO GERAL').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Total de Visualizações: ${dadosRelatorio.resumoGeral.totalVisualizacoes}`);
    body.appendParagraph(`Sessões Únicas: ${dadosRelatorio.resumoGeral.sessesUnicas}`);
    body.appendParagraph(`Produtos Mais Vistos: ${dadosRelatorio.resumoGeral.produtosMaisVistos}`);
    body.appendParagraph('');
    
    // Visualizações por dia
    if (dadosRelatorio.visualizacoesPorDia.length > 0) {
      body.appendParagraph('VISUALIZAÇÕES POR DIA').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      dadosRelatorio.visualizacoesPorDia.forEach(item => {
        body.appendParagraph(`${item.data}: ${item.visualizacoes} visualizações`);
      });
      body.appendParagraph('');
    }
    
    // Produtos mais visualizados
    if (dadosRelatorio.produtosMaisVisualizados.length > 0) {
      body.appendParagraph('PRODUTOS MAIS VISUALIZADOS').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      dadosRelatorio.produtosMaisVisualizados.forEach((produto, index) => {
        body.appendParagraph(`${index + 1}. ${produto.nomeItem}: ${produto.visualizacoes} visualizações`);
      });
    }
    
    // Salvar e converter para PDF
    doc.saveAndClose();
    
    const file = DriveApp.getFileById(doc.getId());
    const pdfBlob = file.getBlob().getAs('application/pdf');
    
    // Criar arquivo PDF no Drive
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName('Relatorio_Cardaplan_' + new Date().toISOString().split('T')[0] + '.pdf');
    
    // Deletar documento temporário
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    
    Logger.log('[gerarRelatorioPDF] PDF gerado com sucesso');
    return {
      success: true,
      fileId: pdfFile.getId(),
      fileName: pdfFile.getName(),
      downloadUrl: `https://drive.google.com/file/d/${pdfFile.getId()}/view`
    };
    
  } catch (error) {
    Logger.log(`[gerarRelatorioPDF] ERRO: ${error.message}`);
    throw new Error(`Erro ao gerar PDF: ${error.message}`);
  }
}

// Funções auxiliares para o Google Apps Script
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Exemplo de uso para depuração (pode ser removido em produção)
function testGetSheetData() {
  try {
    const data = getSheetData('Itens');
    Logger.log('Dados de Itens:', JSON.stringify(data));
  } catch (e) {
    Logger.log('Erro ao testar getSheetData:', e.message);
  }
}

function testBairros() {
  try {
    const bairros = getBairrosList();
    Logger.log('Lista de bairros:', JSON.stringify(bairros));
    
    const taxa = getTaxaEntregaBairro('Centro');
    Logger.log('Taxa do Centro:', taxa);
  } catch (e) {
    Logger.log('Erro ao testar bairros:', e.message);
  }
}

function testAnalytics() {
  try {
    registrarEventoAnalytics('visualizacao_cardapio', null, 'test_session_123');
    const reports = getAnalyticsReports();
    Logger.log('Relatórios de analytics:', JSON.stringify(reports));
  } catch (e) {
    Logger.log('Erro ao testar analytics:', e.message);
  }
}

/**
 * 🚀 Verifica login do cliente
 * @param {string} email - Email digitado pelo cliente no login
 * @returns {Object} { sucesso: boolean, motivo?: string, nome?: string }
 */
function verificarCliente(email) {
  var ss = SpreadsheetApp.openById("1uFa4R8N2SkkVyDLtOiGcaFryq6E7dzAXujC_kA285KQ"); // coloque seu ID
  var aba = ss.getSheetByName("Clientes");
  var dados = aba.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    var emailPlanilha = (dados[i][1] || "").toString().trim().toLowerCase();
    var status = (dados[i][4] || "").toString().trim().toLowerCase(); // Coluna E

    if (emailPlanilha === email.toLowerCase()) {
      if (status === "ativo") {
        return { sucesso: true, nome: dados[i][0] }; // Coluna A = Nome
      } else {
        return { sucesso: false, motivo: "Sua assinatura está inativa." };
      }
    }
  }
  return { sucesso: false, motivo: "Email não encontrado." };
}

/**
 * 📩 Recebe webhook da Cakto e atualiza status no Sheets
 * Espera JSON com "event" e "data.customer.email"
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // ✅ Segurança: confirme o secret
    if (payload.secret !== "5afabd6b-a32c-48b4-89fa-310f70ada15d") {
      return ContentService.createTextOutput("Secret inválido")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    var ss = SpreadsheetApp.openById("1uFa4R8N2SkkVyDLtOiGcaFryq6E7dzAXujC_kA285KQ");
    var aba = ss.getSheetByName("Clientes");
    var dados = aba.getDataRange().getValues();

    var email = payload.data.customer.email;
    var nome = payload.data.customer.name;
    var telefone = payload.data.customer.phone || "";
    var doc = payload.data.customer.docNumber || "";
    var produto = payload.data.product.name || "";
    var idAssinatura = payload.data.id;
    var evento = payload.event;
    var status = (evento === "subscription_canceled") ? "Inativo" : "Ativo";

    var encontrado = false;
    for (var i = 1; i < dados.length; i++) {
      if ((dados[i][1] || "").toString().trim().toLowerCase() === email.toLowerCase()) {
        aba.getRange(i + 1, 1).setValue(nome);
        aba.getRange(i + 1, 3).setValue(telefone);
        aba.getRange(i + 1, 4).setValue(doc);
        aba.getRange(i + 1, 5).setValue(status);
        aba.getRange(i + 1, 6).setValue(new Date()); // Última atualização
        aba.getRange(i + 1, 7).setValue(evento);
        aba.getRange(i + 1, 8).setValue(produto);
        aba.getRange(i + 1, 9).setValue(idAssinatura);
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      aba.appendRow([
        nome,
        email,
        telefone,
        doc,
        status,
        new Date(),
        evento,
        produto,
        idAssinatura
      ]);
    }

    // Suporte a registro de pedido via webhook (conteúdo em JSON com campo kind='order')
    try {
      const tryJson = JSON.parse(e.postData.contents || '{}');
      if (tryJson && tryJson.kind === 'order') {
        const result = registerOrderFromWebhook(tryJson.data || tryJson);
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (_ignore) {}

    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService.createTextOutput("Erro: " + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function fazerLogin() {
    const email = document.getElementById('loginEmail').value.trim();
    const msg = document.getElementById('loginMessage');
    const btn = document.getElementById('btnEntrar');
    const originalText = btn.innerHTML;

    if (!email) {
        msg.textContent = "Digite seu e-mail.";
        return;
    }

    // Mostra spinner só no botão
    btn.innerHTML = `<div class="spinner"></div> Verificando login...`;
    btn.classList.add("btn-loading");
    msg.textContent = "";

    google.script.run
        .withSuccessHandler(function(resposta) {
            btn.innerHTML = originalText;
            btn.classList.remove("btn-loading");

            if (resposta.sucesso) {
                localStorage.setItem('usuarioEmail', email);
                document.getElementById('loginScreen').style.display = 'none';
                document.getElementById('mainApp').style.display = 'block';
                showToast("Bem-vindo, " + resposta.nome, "success");
                loadInitialData();
            } else {
                msg.textContent = resposta.motivo;
            }
        })
        .withFailureHandler(function(err) {
            btn.innerHTML = originalText;
            btn.classList.remove("btn-loading");
            msg.textContent = "Erro no login: " + err.message;
        })
        .verificarCliente(email);
}

