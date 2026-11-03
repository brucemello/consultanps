/**
 * SISTEMA NPS PROTHEUS MULTI-MÓDULO DE INTELIGÊNCIA CX
 * Importação do relatório do Portal de Insights CX (Fluig) + Dashboard Executivo em HTML
 */

const URL_FLUIG = 'https://fluig.totvs.com/portal/p/10097/portal-de-insights-cx-v2';
// Link fixo da Central NPS (Web App). É a implantação atualizada pelo deploy automático (secret DEPLOYMENT_ID).
const URL_PAINEL = 'https://script.google.com/a/macros/totvs.com.br/s/AKfycbyCUu4TFMso0d8E5VonjbOB929O0wv_MXuOhB_sXx3-f-M1-cz66at9ujLDhe7_eU-vAA/exec';
const URL_MANUAL = 'https://drive.google.com/file/d/1VD3aj2c-KN16bDuOecflMKa4YngtKRNM/view?usp=drive_link';

// Filtros padrão aplicados no Portal de Insights CX (Fluig)
const FILTROS_PADRAO = {
  tipoCliente: 'Externo',
  pais: 'BR',
  pesquisa: 'NPS Produto Protheus',
  produto: 'SIGAFIS',
  linha: 'Protheus',
  versao: 'Todos',
  cliente: 'Todos'
};

const CABECALHOS_BASE = [
  'Data da Resposta', 'Status', 'País', 'Linha de Produto', 'Módulo / Produto',
  'Cliente', 'Email', 'Nota', 'CodT', 'Tipo do Cliente', 'Razão Social', 'Comentário',
  'Data_Mes_Ano', 'Classificacao_NPS', 'Status_Aderencia'
];

// Colunas do arquivo exportado do Fluig → aceitas pelo nome (sem acento/maiúsculas)
const MAPA_COLUNAS = {
  data: ['data da resposta', 'data resposta', 'data'],
  status: ['status'],
  pais: ['pais'],
  linha: ['linha de produto', 'linha'],
  modulo: ['modulo / produto', 'modulo/produto', 'nome do produto', 'produto'],
  cliente: ['cliente', 'codigo do cliente'],
  email: ['email', 'e-mail'],
  nota: ['nota'],
  codt: ['codt', 'cod t', 'codigo t'],
  tipo: ['tipo do cliente', 'tipo cliente'],
  razao: ['razao social'],
  comentario: ['comentario', 'comentarios']
};

const COLUNAS_OBRIGATORIAS = ['data', 'status', 'nota', 'razao', 'comentario'];

const PILARES_CX = {
  '1. Desempenho & Estabilidade': {
    palavras: ['lento', 'lentidao', 'lentidão', 'travando', 'trava', 'duplica', 'horroroso', 'erros', 'lock', 'queda', 'timeout', 'demora', 'bug'],
    acao: 'Task force de engenharia para otimização de queries, índices e performance'
  },
  '2. Suporte & Atendimento': {
    palavras: ['suporte', 'atendimento', 'chamado', 'ticket', 'demora para responder', 'retorno', 'incompleto', 'csm', 'analista'],
    acao: 'Protocolo emergencial de atendimento (Red Flag) e contato direto da gestão de CS'
  },
  '3. Regras de Negócio & Funcionalidades': {
    palavras: ['rotina', 'calculo', 'cálculo', 'falha', 'divergencia', 'divergência', 'nao funciona', 'não funciona', 'campo', 'relatorio', 'relatório', 'tes', 'ibs', 'cbs', 'configurador'],
    acao: 'Análise de backlog de produto e priorização de correção de rotina'
  },
  '4. Usabilidade & Capacitação': {
    palavras: ['capacitacao', 'capacitação', 'treinamento', 'interface', 'layout', 'complexo', 'dificil', 'difícil', 'manual', 'tdn', 'nao sei usar', 'não sei usar', 'nao conheco', 'não conheço', 'chat bot', 'chatbot', 'intuitivo', 'intarativo', 'interativo'],
    acao: 'Sessão de Onboarding/Treinamento e envio de documentação técnica do TDN'
  }
};
const PILAR_OUTROS = '5. Outros / Não Classificados';
const ACAO_OUTROS = 'Acompanhamento de rotina de Customer Success';

const COMENTARIOS_VAZIOS = ['sem comentários', 'sem comentarios', 'sem comentário', 'sem comentario', 'nda', 'nada', 'n/a', 'na', '-', '--', '.'];

// ============================================================
// CONTROLE DE ACESSO
// ============================================================
// O Web App roda como o dono (quem implanta) e fica aberto para o domínio TOTVS.
// Quem não for o dono vê só o Dashboard (leitura); todas as funções que gravam,
// enviam e-mail ou mexem em configurações exigem o dono.

/** true para o dono; também true quando não há usuário ativo (gatilhos agendados). */
function ehDono_() {
  const ativo = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const dono = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  return !ativo || ativo === dono;
}

function exigirDono_() {
  if (!ehDono_()) throw new Error('Acesso restrito: somente o responsável pela Central NPS pode fazer isso.');
}

// ============================================================
// MENU, WEB APP E JANELAS
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 NPS Protheus')
    .addItem('🧭 Abrir Central NPS', 'abrirPainel')
    .addItem('🔗 Abrir Central no navegador', 'abrirPainelNavegador')
    .addSeparator()
    .addItem('📊 Dashboard', 'abrirDashboard')
    .addItem('📥 Importar relatório do Fluig', 'mostrarInterfaceUpload')
    .addItem('✉️ Envio de resumo', 'mostrarEnvioResumo')
    .addItem('👥 Times e gestores', 'mostrarCadastro')
    .addItem('⏰ Automações', 'abrirAutomacoes')
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠 Manutenção')
      .addItem('🔄 Recalcular abas de análise', 'gerarAnaliseEGraficos')
      .addItem('📸 Salvar snapshot do NPS por módulo', 'gravarSnapshotNPS')
      .addItem('🌐 Abrir Portal de Insights CX (Fluig)', 'abrirFluig'))
    .addToUi();
}

/** Web App: a mesma Central NPS (aba inicial por ?tab=dash|imp|env|cad|auto). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const leitura = !ehDono_();
  const tab = leitura ? 'dash' : (p.tab || (p.p === 'importar' ? 'imp' : 'dash'));
  return painelHtml_(tab, null, 'webapp', leitura)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function contextoPaginas_(modo, tab, periodoEscolhido, leitura) {
  const urlApp = URL_PAINEL;
  return JSON.stringify({
    modo: modo || 'dialog',
    leitura: !!leitura,
    tabInicial: tab || 'dash',
    periodoEscolhido: periodoEscolhido || null,
    urlApp: urlApp,
    urlFluig: URL_FLUIG,
    urlManual: URL_MANUAL,
    filtros: FILTROS_PADRAO,
    periodo: periodoPadrao_(new Date())
  });
}

function painelHtml_(tab, periodoEscolhido, modo, leitura) {
  const t = HtmlService.createTemplateFromFile('Painel');
  t.leitura = !!leitura;
  t.contexto = contextoPaginas_(modo || 'dialog', tab, periodoEscolhido, leitura);
  return t.evaluate().setTitle('Central NPS · Protheus');
}

/** Abre a Central NPS em tela grande dentro da planilha, já na aba pedida. */
function abrirPainel(tab, periodoEscolhido) {
  exigirDono_();
  const html = painelHtml_(typeof tab === 'string' ? tab : 'dash', periodoEscolhido || null).setWidth(1600).setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// Atalhos do menu (nomes mantidos por compatibilidade)
function abrirDashboard() {
  exigirDono_(); abrirPainel('dash'); }
function mostrarInterfaceUpload() {
  exigirDono_(); abrirPainel('imp'); }
function abrirImportadorAutomatico() {
  exigirDono_(); abrirPainel('imp'); }
function mostrarCadastro() {
  exigirDono_(); abrirPainel('cad'); }
function abrirAutomacoes() {
  exigirDono_(); abrirPainel('auto'); }
function mostrarEnvioResumo(periodoJson) {
  exigirDono_();
  abrirPainel('env', periodoJson ? periodoDeMeses_(JSON.parse(periodoJson)) : null);
}

function abrirPainelNavegador() {
  exigirDono_();
  abrirUrl_(URL_PAINEL, 'Abrindo a Central NPS...');
}

function abrirFluig() {
  exigirDono_();
  abrirUrl_(URL_FLUIG, 'Redirecionando para o Portal Fluig...');
}


function abrirUrl_(url, titulo) {
  const html = HtmlService.createHtmlOutput(
    `<script>window.open(${JSON.stringify(url)}, "_blank");google.script.host.close();</script>`
  ).setWidth(120).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, titulo);
}

// ============================================================
// UTILITÁRIOS (puros — cobertos por testes)
// ============================================================

/** Período padrão: 01/01 do ano até o último dia do mês anterior. Em janeiro, o ano anterior inteiro. */
function periodoPadrao_(hoje) {
  const fim = new Date(hoje.getFullYear(), hoje.getMonth(), 0); // último dia do mês anterior
  const inicio = new Date(fim.getFullYear(), 0, 1);
  const f = d => String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
  return { inicio: f(inicio), fim: f(fim), mesInicio: fmtMes_(inicio), mesFim: fmtMes_(fim) };
}

function fmtMes_(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function fmtDia_(d) {
  return fmtMes_(d) + '-' + String(d.getDate()).padStart(2, '0');
}

function normalizarTexto_(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Aceita Date, serial do Excel, "dd/mm/aaaa [hh:mm[:ss]]" e ISO. Retorna Date ou null. */
function parseData_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') {
    const ms = Math.round((v - 25569) * 864e5);
    const utc = new Date(ms);
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate(), utc.getUTCHours(), utc.getUTCMinutes(), utc.getUTCSeconds());
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function parseNota_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(String(v).replace(',', '.').trim());
  return isNaN(n) ? '' : n;
}

/** Encontra o índice de cada coluna conhecida a partir do cabeçalho do arquivo. */
function mapearColunas_(header) {
  const norm = header.map(normalizarTexto_);
  const idx = {};
  Object.keys(MAPA_COLUNAS).forEach(chave => {
    idx[chave] = -1;
    for (const nome of MAPA_COLUNAS[chave]) {
      const i = norm.indexOf(nome);
      if (i !== -1) { idx[chave] = i; break; }
    }
  });
  return idx;
}

function comentarioValido_(c) {
  const t = String(c == null ? '' : c).trim();
  if (t.length <= 2 && !/^[a-zA-ZÀ-ú]{2}$/.test(t)) return false;
  return COMENTARIOS_VAZIOS.indexOf(normalizarTexto_(t)) === -1;
}

function classificarComentario_(comentario) {
  const comLower = String(comentario || '').toLowerCase();
  for (const pilar in PILARES_CX) {
    if (PILARES_CX[pilar].palavras.some(termo => contemTermo_(comLower, termo))) {
      return { pilar: pilar, acao: PILARES_CX[pilar].acao };
    }
  }
  return { pilar: PILAR_OUTROS, acao: ACAO_OUTROS };
}

/** Termos curtos (≤3 letras, ex.: "tes", "ibs") só contam como palavra inteira. */
function contemTermo_(texto, termo) {
  if (termo.length > 3) return texto.indexOf(termo) !== -1;
  return new RegExp('(^|[^a-zà-ú])' + termo + '([^a-zà-ú]|$)').test(texto);
}

function classificarNota_(status, nota) {
  if (status !== 'Respondido' || nota === '') return '';
  if (nota >= 9) return 'Promotor';
  if (nota >= 7) return 'Neutro';
  return 'Detrator';
}

// ============================================================
// IMPORTAÇÃO
// ============================================================

/**
 * Recebe a planilha já lida no navegador (matriz com cabeçalho na 1ª linha,
 * datas como "dd/mm/aaaa hh:mm:ss") e consolida na Base_Historica.
 */
function importarLinhas(matriz, nomeArquivo) {
  exigirDono_();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'Outra importação está em andamento. Tente de novo em alguns segundos.' };
  }
  try {
    if (!matriz || matriz.length <= 1) {
      return { success: false, message: 'O arquivo não contém linhas de dados.' };
    }
    const idx = mapearColunas_(matriz[0]);
    const faltando = COLUNAS_OBRIGATORIAS.filter(c => idx[c] === -1);
    if (faltando.length) {
      return { success: false, message: 'Colunas obrigatórias ausentes: ' + faltando.map(c => MAPA_COLUNAS[c][0]).join(', ') };
    }
    const r = consolidarBaseHistorica(matriz);
    gerarAnaliseEGraficos();
    PropertiesService.getDocumentProperties().setProperty('ULTIMA_IMPORTACAO', JSON.stringify({
      quando: new Date().toISOString(), arquivo: nomeArquivo || '', inseridos: r.inseridos, duplicados: r.duplicados
    }));
    return {
      success: true,
      inseridos: r.inseridos,
      duplicados: r.duplicados,
      message: r.inseridos === 0
        ? 'Nenhum registro novo — tudo já estava na Base Histórica.'
        : `${r.inseridos} registros novos consolidados (${r.duplicados} já existiam).`
    };
  } catch (error) {
    return { success: false, message: 'Erro no processamento: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Mantido por compatibilidade com versões antigas da tela de upload.
function processarUpload(fileData, fileName) {
  exigirDono_();
  return { success: false, message: 'Esta tela de upload foi substituída. Feche e abra de novo pelo menu "Importar relatório do Fluig".' };
}

// 1. CONSOLIDAÇÃO DA BASE HISTÓRICA (COM CABEÇALHO PADRONIZADO: "Módulo / Produto")
function consolidarBaseHistorica(dados) {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let baseSheet = ss.getSheetByName('Base_Historica');

  if (!baseSheet) {
    baseSheet = ss.insertSheet('Base_Historica');
    baseSheet.getRange(1, 1, 1, CABECALHOS_BASE.length).setValues([CABECALHOS_BASE])
      .setBackground('#102A43').setFontColor('#FFFFFF').setFontWeight('bold');
    baseSheet.setFrozenRows(1);
  } else {
    baseSheet.getRange(1, 5).setValue('Módulo / Produto');
  }

  const dadosExistentes = baseSheet.getLastRow() > 1 ? baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, 11).getValues() : [];
  const chavesExistentes = new Set();
  dadosExistentes.forEach(row => {
    const dt = row[0] ? new Date(row[0]).getTime() : '';
    const razao = String(row[10] || '').trim().toLowerCase();
    if (dt || razao) chavesExistentes.add(`${dt}_${razao}`);
  });

  const { linhas, duplicados } = montarLinhasBase_(dados, chavesExistentes);

  if (linhas.length > 0) {
    const lastRow = baseSheet.getLastRow();
    baseSheet.getRange(lastRow + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
    baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    baseSheet.getRange(2, 13, baseSheet.getLastRow() - 1, 1).setNumberFormat('mmm/yyyy');
  }

  return { inseridos: linhas.length, duplicados: duplicados };
}

/** Converte a matriz do arquivo em linhas da Base_Historica, ignorando as já existentes. */
function montarLinhasBase_(dados, chavesExistentes) {
  const idx = mapearColunas_(dados[0]);
  const v = (row, chave) => (idx[chave] !== -1 && row[idx[chave]] != null) ? row[idx[chave]] : '';
  const linhas = [];
  let duplicados = 0;

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    if (!v(row, 'data') && !v(row, 'status')) continue;

    const dataResp = parseData_(v(row, 'data'));
    const razaoSocial = String(v(row, 'razao')).trim();
    const chave = `${dataResp ? dataResp.getTime() : ''}_${razaoSocial.toLowerCase()}`;
    if (chavesExistentes.has(chave)) { duplicados++; continue; }
    chavesExistentes.add(chave);

    const status = String(v(row, 'status')).trim();
    const nota = parseNota_(v(row, 'nota'));
    const modulo = String(v(row, 'modulo')).trim() || 'SIGAFIS';

    linhas.push([
      dataResp || '', status, v(row, 'pais'), v(row, 'linha'), modulo,
      v(row, 'cliente'), v(row, 'email'), nota, v(row, 'codt'), v(row, 'tipo'), razaoSocial, v(row, 'comentario'),
      dataResp ? new Date(dataResp.getFullYear(), dataResp.getMonth(), 1) : '',
      classificarNota_(status, nota),
      status === 'Respondido' ? 'Respondido' : 'Cancelado'
    ]);
  }
  return { linhas: linhas, duplicados: duplicados };
}

// ============================================================
// DADOS DO DASHBOARD
// ============================================================

function getDadosDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName('Base_Historica');
  const tz = Session.getScriptTimeZone();
  const linhas = [];

  if (base && base.getLastRow() > 1) {
    const data = base.getRange(2, 1, base.getLastRow() - 1, CABECALHOS_BASE.length).getValues();
    data.forEach(r => {
      const d = r[0] instanceof Date ? r[0] : parseData_(r[0]);
      if (!d) return;
      const status = String(r[1] || '').trim();
      const nota = r[7] === '' ? null : Number(r[7]);
      const com = comentarioValido_(r[11]) ? String(r[11]).trim() : '';
      const pil = com ? classificarComentario_(com).pilar : '';
      linhas.push([
        Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
        status === 'Respondido' ? 1 : 0,
        String(r[2] || ''), String(r[4] || 'SIGAFIS'), String(r[9] || ''),
        nota, String(r[13] || ''),
        com ? String(r[10] || '') : '', com, pil
      ]);
    });
  }

  const plano = [];
  const planoSheet = ss.getSheetByName('Plano_de_Acao');
  if (planoSheet && planoSheet.getLastRow() > 1) {
    planoSheet.getRange(2, 1, planoSheet.getLastRow() - 1, 9).getValues().forEach(r => {
      const mes = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM') : '';
      plano.push([String(r[0]), mes, String(r[2]), String(r[3]), String(r[4]), String(r[5]), String(r[6]), String(r[8])]);
    });
  }

  const acoes = {};
  Object.keys(PILARES_CX).forEach(p => acoes[p] = PILARES_CX[p].acao);
  acoes[PILAR_OUTROS] = ACAO_OUTROS;

  let ultima = null;
  try { ultima = JSON.parse(PropertiesService.getDocumentProperties().getProperty('ULTIMA_IMPORTACAO') || 'null'); } catch (e) { ultima = null; }

  return JSON.stringify({ linhas: linhas, plano: plano, acoes: acoes, ultimaImportacao: ultima, geradoEm: new Date().toISOString() });
}

// 2. RECONSTRÓI A ABA ANALISE_QUALI COM O MESMO NOME DE COLUNA ("Módulo / Produto")
function reconstruirAnaliseQualitativa(dataBase) {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Analise_Quali');
  if (!sheet) sheet = ss.insertSheet('Analise_Quali');
  sheet.clear();

  const headers = [['Mês Ref.', 'Módulo / Produto', 'Cliente / Razão Social', 'Nota', 'Comentário Original', 'Pilar CX / Causa Raiz', 'Ação Sugerida', 'Status Acompanhamento']];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#2C5282').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const linhasParaInserir = [];
  for (let i = 1; i < dataBase.length; i++) {
    const row = dataBase[i];
    const comentario = String(row[11] || '').trim();
    if (!comentarioValido_(comentario)) continue;
    const c = classificarComentario_(comentario);
    linhasParaInserir.push([row[12], row[4] || 'SIGAFIS', row[10] || 'N/A', row[7], comentario, c.pilar, c.acao, 'Pendente']);
  }

  if (linhasParaInserir.length > 0) {
    sheet.getRange(2, 1, linhasParaInserir.length, linhasParaInserir[0].length).setValues(linhasParaInserir);
    sheet.getRange(2, 1, linhasParaInserir.length, 1).setNumberFormat('mmm/yyyy');
  }
}

// 3. GERADOR GERAL DE DASHBOARDS E PROCESSAMENTOS
function gerarAnaliseEGraficos() {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base_Historica');
  if (!baseSheet) return;

  baseSheet.getRange(1, 5).setValue('Módulo / Produto');

  const data = baseSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  reconstruirAnaliseQualitativa(data);

  const estatisticas = {};

  for (let i = 1; i < data.length; i++) {
    const rawData = data[i][0];
    const status = data[i][1];
    const modulo = data[i][4] || 'SIGAFIS';
    const classif = data[i][13];

    if (!rawData) continue;
    let dataObj = new Date(rawData);
    if (isNaN(dataObj.getTime())) continue;

    const chaveMes = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), 'yyyy-MM');
    const chaveComposta = `${chaveMes}_${modulo}`;

    if (!estatisticas[chaveComposta]) {
      estatisticas[chaveComposta] = {
        dataInicioMes: new Date(dataObj.getFullYear(), dataObj.getMonth(), 1),
        modulo: modulo,
        abordagens: 0, respostas: 0, promotores: 0, neutros: 0, detratores: 0
      };
    }

    estatisticas[chaveComposta].abordagens++;
    if (status === 'Respondido') {
      estatisticas[chaveComposta].respostas++;
      if (classif === 'Promotor') estatisticas[chaveComposta].promotores++;
      else if (classif === 'Neutro') estatisticas[chaveComposta].neutros++;
      else if (classif === 'Detrator') estatisticas[chaveComposta].detratores++;
    }
  }

  let dashSheet = ss.getSheetByName('Dashboard_KPIs');
  if (!dashSheet) dashSheet = ss.insertSheet('Dashboard_KPIs');
  dashSheet.clear();

  const headers = [['Mês/Ano', 'Módulo / Produto', 'Abordagens', 'Respostas', 'Taxa Aderência', '% Promotores', '% Neutros', '% Detratores', 'NPS Score', 'Status NPS']];
  const rows = [];
  const chavesOrdenadas = Object.keys(estatisticas).sort();

  chavesOrdenadas.forEach(chave => {
    const st = estatisticas[chave];
    const aderencia = st.abordagens > 0 ? (st.respostas / st.abordagens) : 0;
    const pctProm = st.respostas > 0 ? (st.promotores / st.respostas) : 0;
    const pctNeu = st.respostas > 0 ? (st.neutros / st.respostas) : 0;
    const pctDet = st.respostas > 0 ? (st.detratores / st.respostas) : 0;
    const nps = (pctProm - pctDet) * 100;

    let statusNps = "Crítico";
    if (nps >= 75) statusNps = "Excelente";
    else if (nps >= 50) statusNps = "Muito Bom";
    else if (nps >= 0) statusNps = "Razoável";

    rows.push([st.dataInicioMes, st.modulo, st.abordagens, st.respostas, aderencia, pctProm, pctNeu, pctDet, nps, statusNps]);
  });

  dashSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#102A43').setFontColor('#FFFFFF').setFontWeight('bold');

  if (rows.length > 0) {
    dashSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    dashSheet.getRange(2, 1, rows.length, 1).setNumberFormat('mmm/yyyy');
    dashSheet.getRange(2, 3, rows.length, 2).setNumberFormat('#,##0');
    dashSheet.getRange(2, 5, rows.length, 4).setNumberFormat('0.0%');
    dashSheet.getRange(2, 9, rows.length, 1).setNumberFormat('0.0');
    aplicarFormatacaoCondicional(dashSheet, rows.length + 1);
  }

  gerarRankingEDecisao();
  gerarPlanoDeAcaoInteligente();
  gerarRelatorioConsolidado();
}

function aplicarFormatacaoCondicional(sheet, maxRow) {
  exigirDono_();
  sheet.clearConditionalFormatRules();
  const rangeNPS = sheet.getRange(2, 9, maxRow - 1, 1);

  const ruleVerde = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(50).setBackground('#C6F6D5').setFontColor('#22543D').setRanges([rangeNPS]).build();
  const ruleAmarelo = SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0, 49.9).setBackground('#FEFCBF').setFontColor('#744210').setRanges([rangeNPS]).build();
  const ruleVermelho = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#FED7D7').setFontColor('#742A2A').setRanges([rangeNPS]).build();

  sheet.setConditionalFormatRules([ruleVerde, ruleAmarelo, ruleVermelho]);
}

// 4. RANKING DE PROBLEMAS COM NOME DE COLUNA PADRONIZADO
function gerarRankingEDecisao() {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qualiSheet = ss.getSheetByName('Analise_Quali');
  if (!qualiSheet) return;

  const dataQuali = qualiSheet.getDataRange().getValues();
  if (dataQuali.length <= 1) return;

  const grupos = {};

  for (let i = 1; i < dataQuali.length; i++) {
    const row = dataQuali[i];
    const modulo = String(row[1] || 'SIGAFIS').trim();
    const nota = Number(row[3]);
    const pilar = String(row[5] || PILAR_OUTROS).trim();

    const chaveComposta = `${modulo}_${pilar}`;
    if (!grupos[chaveComposta]) {
      grupos[chaveComposta] = { modulo: modulo, pilar: pilar, qtd: 0, somaNota: 0 };
    }
    grupos[chaveComposta].qtd++;
    grupos[chaveComposta].somaNota += isNaN(nota) ? 0 : nota;
  }

  let rankSheet = ss.getSheetByName('Ranking_Problemas');
  if (!rankSheet) rankSheet = ss.insertSheet('Ranking_Problemas');
  rankSheet.clear();

  const headers = [['Ranking (#)', 'Módulo / Produto', 'Pilar de Causa Raiz CX', 'Qtd. Ocorrências', 'Nota Média NPS', 'Nível de Urgência', 'Ação Estratégica Recomendada']];
  const lista = Object.values(grupos).sort((a, b) => b.qtd - a.qtd);

  const linhasFinais = [];
  lista.forEach((item, index) => {
    const mediaNota = item.qtd > 0 ? (item.somaNota / item.qtd) : 0;

    let urgencia = "🟢 Baixo";
    let acao = `Acompanhamento de rotina de CS no módulo ${item.modulo}`;

    if (item.pilar.includes('Desempenho')) {
      urgencia = "🔴 Crítico";
      acao = `Task force de engenharia para otimização de queries, índices e banco de dados do módulo ${item.modulo}`;
    } else if (item.pilar.includes('Suporte')) {
      urgencia = "🔴 Crítico";
      acao = `Protocolo emergencial de atendimento (Red Flag) para chamados do módulo ${item.modulo}`;
    } else if (item.pilar.includes('Regras')) {
      urgencia = "🟡 Médio";
      acao = `Revisar rotinas e backlog de correções da equipe de produto do ${item.modulo}`;
    } else if (item.pilar.includes('Usabilidade')) {
      urgencia = "🟢 Baixo";
      acao = `Promover treinamentos, documentação no TDN e Onboarding para usuários do ${item.modulo}`;
    }

    linhasFinais.push(["#" + (index + 1), item.modulo, item.pilar, item.qtd, mediaNota, urgencia, acao]);
  });

  rankSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#1A365D').setFontColor('#FFFFFF').setFontWeight('bold');

  if (linhasFinais.length > 0) {
    rankSheet.getRange(2, 1, linhasFinais.length, linhasFinais[0].length).setValues(linhasFinais);
    rankSheet.getRange(2, 4, linhasFinais.length, 1).setNumberFormat('#,##0');
    rankSheet.getRange(2, 5, linhasFinais.length, 1).setNumberFormat('0.0');
  }
}

// 5. PLANO DE AÇÃO INTELIGENTE POR MÓDULO
function gerarPlanoDeAcaoInteligente() {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qualiSheet = ss.getSheetByName('Analise_Quali');
  let planoSheet = ss.getSheetByName('Plano_de_Acao');

  if (!qualiSheet) return;
  if (!planoSheet) planoSheet = ss.insertSheet('Plano_de_Acao');

  const dataQuali = qualiSheet.getDataRange().getValues();
  if (dataQuali.length <= 1) return;

  const cabecalhos = [['ID Ação', 'Mês Ref.', 'Módulo / Produto', 'Causa Raiz / Reclamação', 'Área Responsável', 'Plano de Ação (5W2H)', 'Prazo', 'Responsável', 'Status']];
  const linhasPlano = [];
  let cont = 1;

  for (let i = 1; i < dataQuali.length; i++) {
    const mes = dataQuali[i][0];
    const modulo = dataQuali[i][1] || 'SIGAFIS';
    const cliente = dataQuali[i][2];
    const pilar = dataQuali[i][5] || 'Outros';
    const acaoSugerida = dataQuali[i][6];

    let id = "PA-" + String(cont).padStart(3, '0');
    let area = "Equipe de Produto " + modulo;
    let causa = `${pilar} (${cliente})`;
    let prazo = "30 dias";
    let status = "A Iniciar";

    if (pilar.includes('Desempenho')) {
      area = "Engenharia & Performance";
    } else if (pilar.includes('Suporte')) {
      area = "Gestão de CS / Suporte";
    }

    linhasPlano.push([id, mes, modulo, causa, area, acaoSugerida, prazo, area, status]);
    cont++;
  }

  planoSheet.clear();
  planoSheet.getRange(1, 1, 1, cabecalhos[0].length).setValues(cabecalhos).setBackground('#276749').setFontColor('#FFFFFF').setFontWeight('bold');

  if (linhasPlano.length > 0) {
    planoSheet.getRange(2, 1, linhasPlano.length, linhasPlano[0].length).setValues(linhasPlano);
    planoSheet.getRange(2, 2, linhasPlano.length, 1).setNumberFormat('mmm/yyyy');
  }
}

// 6. RELATÓRIO CONSOLIDADO E PRINCIPAIS ACHADOS COM COLUNA MÓDULO / PRODUTO
function gerarRelatorioConsolidado() {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base_Historica');
  const rankSheet = ss.getSheetByName('Ranking_Problemas');
  if (!baseSheet) return;

  const data = baseSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const consolPorModulo = {};
  const paisesPorModulo = {};

  for (let i = 1; i < data.length; i++) {
    const status = data[i][1];
    const pais = String(data[i][2] || '').trim();
    const modulo = data[i][4] || 'SIGAFIS';
    const comentario = data[i][11];
    const classif = data[i][13];

    if (!consolPorModulo[modulo]) {
      consolPorModulo[modulo] = { respostas: 0, comentarios: 0, promotores: 0, neutros: 0, detratores: 0 };
      paisesPorModulo[modulo] = {};
    }

    if (status === 'Respondido') {
      consolPorModulo[modulo].respostas++;
      if (pais) paisesPorModulo[modulo][pais] = (paisesPorModulo[modulo][pais] || 0) + 1;

      if (classif === 'Promotor') consolPorModulo[modulo].promotores++;
      else if (classif === 'Neutro') consolPorModulo[modulo].neutros++;
      else if (classif === 'Detrator') consolPorModulo[modulo].detratores++;

      if (comentarioValido_(comentario)) consolPorModulo[modulo].comentarios++;
    }
  }

  let relSheet = ss.getSheetByName('Relatorio_Consolidado');
  if (!relSheet) relSheet = ss.insertSheet('Relatorio_Consolidado');
  relSheet.clear();

  const headersRel = [['Módulo / Produto', 'NPS Score', '% Promotores', '% Neutros', '% Detratores', 'Total Respostas', 'Total Comentários']];
  const rowsRel = [];

  for (let mod in consolPorModulo) {
    const st = consolPorModulo[mod];
    const pctProm = st.respostas > 0 ? (st.promotores / st.respostas) : 0;
    const pctNeu = st.respostas > 0 ? (st.neutros / st.respostas) : 0;
    const pctDet = st.respostas > 0 ? (st.detratores / st.respostas) : 0;
    const nps = (pctProm - pctDet) * 100;
    rowsRel.push([mod, nps, pctProm, pctNeu, pctDet, st.respostas, st.comentarios]);
  }

  relSheet.getRange(1, 1, 1, headersRel[0].length).setValues(headersRel).setBackground('#102A43').setFontColor('#FFFFFF').setFontWeight('bold');
  if (rowsRel.length > 0) {
    relSheet.getRange(2, 1, rowsRel.length, rowsRel[0].length).setValues(rowsRel);
    relSheet.getRange(2, 2, rowsRel.length, 1).setNumberFormat("0.0");
    relSheet.getRange(2, 3, rowsRel.length, 3).setNumberFormat("0.0%");
    relSheet.getRange(2, 6, rowsRel.length, 2).setNumberFormat("#,##0");
  }

  let achadosSheet = ss.getSheetByName('Principais_Achados');
  if (!achadosSheet) achadosSheet = ss.insertSheet('Principais_Achados');
  achadosSheet.clear();

  const doresPorModulo = {};
  if (rankSheet && rankSheet.getLastRow() > 1) {
    const rankData = rankSheet.getRange(2, 2, rankSheet.getLastRow() - 1, 4).getValues();
    rankData.forEach(r => {
      const mod = r[0];
      if (!doresPorModulo[mod]) doresPorModulo[mod] = [];
      doresPorModulo[mod].push(`${r[1]} (${r[2]} ocorrências)`);
    });
  }

  const headersAchados = [['Módulo / Produto', 'Item (#)', 'Principais Achados Executivos']];
  const linhasAchados = [];

  for (let mod in consolPorModulo) {
    const st = consolPorModulo[mod];
    const pctProm = st.respostas > 0 ? (st.promotores / st.respostas) : 0;
    const pctNeu = st.respostas > 0 ? (st.neutros / st.respostas) : 0;
    const pctDet = st.respostas > 0 ? (st.detratores / st.respostas) : 0;
    const nps = (pctProm - pctDet) * 100;

    const topDores = doresPorModulo[mod] ? doresPorModulo[mod].slice(0, 2).join(" e ") : "Sem dores críticas registradas";
    const paises = paisesPorModulo[mod] || {};
    const paisTop = Object.keys(paises).sort((a, b) => paises[b] - paises[a])[0] || 'N/D';

    linhasAchados.push([mod, 1, `• NPS do módulo ${mod}: ${nps.toFixed(1)} (Promotores: ${(pctProm*100).toFixed(1)}% | Neutros: ${(pctNeu*100).toFixed(1)}% | Detratores: ${(pctDet*100).toFixed(1)}%).`]);
    linhasAchados.push([mod, 2, `• Volume total: ${st.respostas} respostas consolidadas e ${st.comentarios} comentários qualificados.`]);
    linhasAchados.push([mod, 3, `• Principais fatores de atrito/detração em ${mod}: ${topDores}.`]);
    linhasAchados.push([mod, 4, `• País com maior concentração de feedbacks: ${paisTop}.`]);
  }

  achadosSheet.getRange(1, 1, 1, headersAchados[0].length).setValues(headersAchados).setBackground('#1A365D').setFontColor('#FFFFFF').setFontWeight('bold');
  if (linhasAchados.length > 0) {
    achadosSheet.getRange(2, 1, linhasAchados.length, linhasAchados[0].length).setValues(linhasAchados);
  }
}

/**
 * GRAVA O HISTÓRICO DE NPS LENDO DIRETAMENTE CADA MÊS/ANO DA ABA DASHBOARD_KPIS
 */
function gravarSnapshotNPS() {
  exigirDono_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName('Dashboard_KPIs');

  if (!dashSheet) {
    SpreadsheetApp.getUi().alert('Aba Dashboard_KPIs não encontrada. Execute "Recalcular Dashboards e Gráficos" primeiro!');
    return;
  }

  const dataDash = dashSheet.getDataRange().getValues();
  if (dataDash.length <= 1) return;

  let histSheet = ss.getSheetByName('Historico_NPS_Modulos');
  const headers = [['Data/Hora Registro', 'Mês/Ano Ref.', 'Módulo / Produto', 'NPS Score', '% Promotores', '% Neutros', '% Detratores', 'Total Respostas']];
  if (!histSheet) histSheet = ss.insertSheet('Historico_NPS_Modulos');

  const dataAtual = new Date();
  const novasLinhas = [];

  for (let i = 1; i < dataDash.length; i++) {
    const row = dataDash[i];
    const rawMes = row[0];
    const modulo = row[1];
    if (!modulo || !rawMes) continue;

    let mesFmt = rawMes;
    if (rawMes instanceof Date) mesFmt = Utilities.formatDate(rawMes, Session.getScriptTimeZone(), 'MM/yyyy');

    novasLinhas.push([dataAtual, mesFmt, modulo, row[8], row[5], row[6], row[7], row[3]]);
  }

  histSheet.clear();
  histSheet.getRange(1, 1, 1, headers[0].length).setValues(headers)
    .setBackground('#1A365D').setFontColor('#FFFFFF').setFontWeight('bold');
  histSheet.setFrozenRows(1);

  if (novasLinhas.length > 0) {
    histSheet.getRange(2, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    histSheet.getRange(2, 1, novasLinhas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    histSheet.getRange(2, 4, novasLinhas.length, 1).setNumberFormat('0.0');
    histSheet.getRange(2, 5, novasLinhas.length, 3).setNumberFormat('0.0%');
    histSheet.getRange(2, 8, novasLinhas.length, 1).setNumberFormat('#,##0');
    SpreadsheetApp.getUi().alert(`✅ Snapshot gravado com sucesso! ${novasLinhas.length} registros consolidados mês a mês.`);
  }
}

// ============================================================
// LEMBRETE MENSAL (1º DIA ÚTIL)
// ============================================================

const HANDLER_LEMBRETE = 'lembreteImportacaoMensal';

/** Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher). */
function pascoa_(ano) {
  const a = ano % 19, b = Math.floor(ano / 100), c = ano % 100, d = Math.floor(b / 4), e = b % 4;
  const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31), dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(ano, mes - 1, dia);
}

/** Feriados nacionais (e pontos facultativos usuais) no formato yyyy-MM-dd. */
function feriadosNacionais_(ano) {
  const fixos = ['01-01', '04-21', '05-01', '09-07', '10-12', '11-02', '11-15', '11-20', '12-25'].map(md => `${ano}-${md}`);
  const p = pascoa_(ano);
  const rel = dias => fmtDia_(new Date(p.getFullYear(), p.getMonth(), p.getDate() + dias));
  return new Set(fixos.concat([rel(-48), rel(-47), rel(-2), rel(60)])); // carnaval (seg/ter), sexta santa, Corpus Christi
}

function ehDiaUtil_(d) {
  const dow = d.getDay();
  return dow !== 0 && dow !== 6 && !feriadosNacionais_(d.getFullYear()).has(fmtDia_(d));
}

function primeiroDiaUtil_(ano, mes) {
  const d = new Date(ano, mes, 1);
  while (!ehDiaUtil_(d)) d.setDate(d.getDate() + 1);
  return d;
}

/** Disparado todo dia pelo gatilho; só envia no 1º dia útil do mês. */
function lembreteImportacaoMensal() {
  exigirDono_();
  const hoje = new Date();
  const pdu = primeiroDiaUtil_(hoje.getFullYear(), hoje.getMonth());
  if (fmtDia_(hoje) !== fmtDia_(pdu)) return;
  enviarLembrete_(false);
}

function enviarLembreteTeste() {
  exigirDono_();
  const para = enviarLembrete_(true);
  SpreadsheetApp.getUi().alert(`📧 Lembrete de teste enviado para ${para}.`);
}

function enviarLembrete_(teste) {
  const para = Session.getEffectiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const periodo = periodoPadrao_(new Date());
  const tz = Session.getScriptTimeZone();

  // Último mês já presente na base
  let ultimoMes = '';
  const base = ss.getSheetByName('Base_Historica');
  if (base && base.getLastRow() > 1) {
    const datas = base.getRange(2, 1, base.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(d => d instanceof Date);
    if (datas.length) ultimoMes = fmtMes_(new Date(Math.max.apply(null, datas)));
  }
  const pendente = !ultimoMes || ultimoMes < periodo.mesFim;
  const nomeMes = m => { const [y, mm] = m.split('-'); return ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'][+mm - 1] + '/' + y; };

  const urlApp = URL_PAINEL;
  const linkImportar = URL_PAINEL + '?tab=imp';
  const F = FILTROS_PADRAO;
  const linha = (k, v) => `<tr><td style="padding:6px 12px;color:#64748b;font-size:13px">${k}</td><td style="padding:6px 12px;font-weight:600;font-size:13px">${v}</td></tr>`;
  const botao = (url, txt, cor) => `<a href="${url}" style="display:inline-block;background:${cor};color:#fff;text-decoration:none;font-weight:600;padding:10px 16px;border-radius:8px;margin-right:8px;font-size:14px">${txt}</a>`;

  const html = `
  <div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;color:#0f172a">
    <div style="background:#0b1f3a;color:#fff;padding:18px 20px;border-radius:12px 12px 0 0">
      <div style="font-size:11px;letter-spacing:2px;color:#93b0e0;font-weight:700">NPS PRODUTO · PROTHEUS</div>
      <div style="font-size:19px;font-weight:700;margin-top:4px">Hora de importar o NPS de ${nomeMes(periodo.mesFim)}</div>
    </div>
    <div style="border:1px solid #e2e8f0;border-top:0;padding:18px 20px;border-radius:0 0 12px 12px">
      <p style="font-size:14px;margin:0 0 12px">${pendente
        ? `A base está atualizada até <b>${ultimoMes ? nomeMes(ultimoMes) : '—'}</b>. Extraia o relatório no Fluig com os filtros abaixo e importe o arquivo.`
        : `A base já tem dados até <b>${nomeMes(ultimoMes)}</b>. Se quiser, reimporte o período completo para pegar respostas atrasadas — duplicados são ignorados.`}</p>
      <table style="border-collapse:collapse;background:#f8fafc;border-radius:8px;width:100%;margin-bottom:16px">
        ${linha('Período', `${periodo.inicio} a ${periodo.fim}`)}
        ${linha('Tipo do Cliente', F.tipoCliente)}${linha('País', F.pais)}${linha('Pesquisa', F.pesquisa)}
        ${linha('Produto', F.produto)}${linha('Linha de Produto', F.linha)}${linha('Versão', F.versao)}${linha('Cliente', F.cliente)}
      </table>
      ${botao(URL_FLUIG, '1. Abrir o Fluig', '#13325e')}${botao(linkImportar, '2. Importar arquivo', '#2563eb')}
      ${urlApp ? `<p style="font-size:13px;margin:16px 0 0"><a href="${urlApp}" style="color:#2563eb">Abrir a Central NPS</a></p>` : ''}
    </div>
  </div>`;

  MailApp.sendEmail({
    to: para,
    subject: `${teste ? '[TESTE] ' : ''}NPS Protheus · importar relatório de ${nomeMes(periodo.mesFim)}`,
    htmlBody: html
  });
  return para;
}

function instalarLembreteMensal(silencioso) {
  exigirDono_();
  removerGatilhos_(HANDLER_LEMBRETE);
  ScriptApp.newTrigger(HANDLER_LEMBRETE).timeBased().everyDays(1).atHour(7).inTimezone(Session.getScriptTimeZone()).create();
  if (silencioso !== true) SpreadsheetApp.getUi().alert(`✅ Lembrete ativado. Você receberá um e-mail às 7h do 1º dia útil de cada mês.\nPróximo: ${proximoLembrete_()}.`);
}

function removerLembreteMensal() {
  exigirDono_();
  const n = removerGatilhos_(HANDLER_LEMBRETE);
  SpreadsheetApp.getUi().alert(n ? '🔕 Lembrete desativado.' : 'Não havia lembrete ativo.');
}

function removerGatilhos_(handler) {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === (handler || HANDLER_LEMBRETE)) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

function gatilhoAtivo_(handler) {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === handler);
}

function fmtDataBR_(d) {
  return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
}

function proximoLembrete_() {
  const hoje = new Date();
  let d = primeiroDiaUtil_(hoje.getFullYear(), hoje.getMonth());
  if (fmtDia_(d) < fmtDia_(hoje) || (fmtDia_(d) === fmtDia_(hoje) && hoje.getHours() >= 7)) d = primeiroDiaUtil_(hoje.getFullYear(), hoje.getMonth() + 1);
  return fmtDataBR_(d);
}
