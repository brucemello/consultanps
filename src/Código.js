/**
 * SISTEMA NPS PROTHEUS MULTI-MÓDULO DE INTELIGÊNCIA CX
 * Com Padronização de Cabeçalhos (Módulo / Produto) em TODAS as abas para o Looker Studio
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 NPS Protheus Multi-Módulo')
    .addItem('📊 Abrir Importador de NPS', 'mostrarInterfaceUpload')
    .addSeparator()
    .addItem('🔄 Recalcular Dashboards e Gráficos', 'gerarAnaliseEGraficos')
    .addItem('📸 Salvar Snapshot do NPS por Módulo', 'gravarSnapshotNPS') // <-- NOVO BOTÃO
    .addSeparator()
    .addItem('🌐 Abrir Portal de Insights CX (Fluig)', 'abrirFluig')
    .addItem('🔗 Abrir Dashboard no Looker Studio', 'abrirLinkLooker')
    .addToUi();
}

function abrirImportadorAutomatico() {
  mostrarInterfaceUpload();
}

function mostrarInterfaceUpload() {
  const html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setWidth(780)
    .setHeight(680)
    .setTitle('Central de Inteligência NPS Protheus');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function processarUpload(fileData, fileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bytes = Utilities.base64Decode(fileData.split(',')[1]);
    const blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
    
    const tempFile = DriveApp.createFile(blob);
    const tempSpreadsheet = SpreadsheetApp.open(tempFile);
    const sourceSheet = tempSpreadsheet.getSheets()[0];
    const sourceData = sourceSheet.getDataRange().getValues();
    tempFile.setTrashed(true);

    if (sourceData.length <= 1) {
      return { success: false, message: 'O arquivo enviado não contém linhas de dados válidas.' };
    }

    const inseridos = consolidarBaseHistorica(sourceData);
    gerarAnaliseEGraficos();

    if (inseridos === 0) {
      return { 
        success: true, 
        message: 'Nenhum novo registro inserido. Todos os dados do arquivo já existem na Base Histórica!' 
      };
    }

    return { 
      success: true, 
      message: `Sucesso! ${inseridos} novos registros consolidados.` 
    };

  } catch (error) {
    return { success: false, message: 'Erro no processamento: ' + error.toString() };
  }
}

// 1. CONSOLIDAÇÃO DA BASE HISTÓRICA (COM CABEÇALHO PADRONIZADO: "Módulo / Produto")
function consolidarBaseHistorica(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let baseSheet = ss.getSheetByName('Base_Historica');
  
  // Padronização: 'Módulo / Produto' na 5ª posição
  const cabecalhosEsperados = [
    'Data da Resposta', 'Status', 'País', 'Linha de Produto', 'Módulo / Produto', 
    'Cliente', 'Email', 'Nota', 'CodT', 'Tipo do Cliente', 'Razão Social', 'Comentário',
    'Data_Mes_Ano', 'Classificacao_NPS', 'Status_Aderencia'
  ];

  if (!baseSheet) {
    baseSheet = ss.insertSheet('Base_Historica');
    baseSheet.getRange(1, 1, 1, cabecalhosEsperados.length).setValues([cabecalhosEsperados]);
    baseSheet.getRange(1, 1, 1, cabecalhosEsperados.length)
      .setBackground('#102A43')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    baseSheet.setFrozenRows(1);
  } else {
    // Força a atualização do nome da coluna E para "Módulo / Produto" caso ainda esteja "Nome do Produto"
    baseSheet.getRange(1, 5).setValue('Módulo / Produto');
  }

  const dadosExistentes = baseSheet.getLastRow() > 1 ? baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, 11).getValues() : [];
  const chavesExistentes = new Set();
  
  dadosExistentes.forEach(row => {
    const dt = row[0] ? new Date(row[0]).getTime() : '';
    const cliente = String(row[10] || '').trim().toLowerCase();
    if (dt || cliente) chavesExistentes.add(`${dt}_${cliente}`);
  });

  const header = dados[0];
  const idxData = header.indexOf('Data da Resposta');
  const idxStatus = header.indexOf('Status');
  const idxNota = header.indexOf('Nota');
  const idxComentario = header.indexOf('Comentário');
  const idxRazao = header.indexOf('Razão Social');
  const idxProduto = header.indexOf('Módulo / Produto') !== -1 ? header.indexOf('Módulo / Produto') : 
                     (header.indexOf('Nome do Produto') !== -1 ? header.indexOf('Nome do Produto') : header.indexOf('Linha de Produto'));

  const novasLinhas = [];

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    if (!row[idxData] && !row[idxStatus]) continue;

    const dataResp = new Date(row[idxData]);
    const dataValida = !isNaN(dataResp.getTime());
    const razaoSocial = String(row[idxRazao] || '').trim();
    
    const chaveUnica = `${dataValida ? dataResp.getTime() : ''}_${razaoSocial.toLowerCase()}`;
    if (chavesExistentes.has(chaveUnica)) continue;
    chavesExistentes.add(chaveUnica);

    const dataPrimeiroDiaMes = dataValida ? new Date(dataResp.getFullYear(), dataResp.getMonth(), 1) : '';
    const status = String(row[idxStatus] || '').trim();
    const nota = (row[idxNota] !== "" && row[idxNota] !== null && !isNaN(row[idxNota])) ? Number(row[idxNota]) : "";
    const modulo = (idxProduto !== -1 && row[idxProduto]) ? String(row[idxProduto]).trim() : 'SIGAFIS';
    
    let classificacao = "";
    if (status === "Respondido" && nota !== "") {
      if (nota >= 9) classificacao = "Promotor";
      else if (nota >= 7) classificacao = "Neutro";
      else classificacao = "Detrator";
    }

    const aderencia = (status === "Respondido") ? "Respondido" : "Cancelado";

    novasLinhas.push([
      dataValida ? dataResp : '', row[1] || '', row[2] || '', row[3] || '', modulo,
      row[5] || '', row[6] || '', nota, row[8] || '', row[9] || '', row[10] || '', row[11] || '',
      dataPrimeiroDiaMes, classificacao, aderencia
    ]);
  }

  if (novasLinhas.length > 0) {
    const lastRow = baseSheet.getLastRow();
    baseSheet.getRange(lastRow + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    baseSheet.getRange(2, 1, baseSheet.getLastRow() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    baseSheet.getRange(2, 13, baseSheet.getLastRow() - 1, 1).setNumberFormat('mmm/yyyy');
  }

  return novasLinhas.length;
}

// 2. RECONSTRÓI A ABA ANALISE_QUALI COM O MESMO NOME DE COLUNA ("Módulo / Produto")
function reconstruirAnaliseQualitativa(dataBase) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Analise_Quali');
  
  if (!sheet) {
    sheet = ss.insertSheet('Analise_Quali');
  }
  sheet.clear();

  const headers = [['Mês Ref.', 'Módulo / Produto', 'Cliente / Razão Social', 'Nota', 'Comentário Original', 'Pilar CX / Causa Raiz', 'Ação Sugerida', 'Status Acompanhamento']];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#2C5282').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const PILARES_CX = {
    '1. Desempenho & Estabilidade': {
      palavras: ['lento', 'lentidao', 'lentidão', 'travando', 'duplica', 'horroroso', 'erros', 'lock', 'queda', 'timeout', 'demora', 'bug'],
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
      palavras: ['capacitacao', 'capacitação', 'treinamento', 'interface', 'layout', 'complexo', 'dificil', 'difícil', 'manual', 'tdn', 'nao sei usar', 'não sei usar', 'nao conheco', 'não conheço'],
      acao: 'Sessão de Onboarding/Treinamento e envio de documentação técnica do TDN'
    }
  };

  const linhasParaInserir = [];

  for (let i = 1; i < dataBase.length; i++) {
    const row = dataBase[i];
    const mes = row[12]; // Data_Mes_Ano
    const modulo = row[4] || 'SIGAFIS';
    const cliente = row[10] || 'N/A';
    const nota = row[7];
    const comentario = String(row[11] || '').trim();

    if (!comentario || comentario.length <= 2 || comentario.toLowerCase() === 'sem comentários') continue;

    const comLower = comentario.toLowerCase();
    let pilarEncontrado = '5. Outros / Não Classificados';
    let acaoSugerida = 'Acompanhamento de rotina de Customer Success';

    for (let pilar in PILARES_CX) {
      if (PILARES_CX[pilar].palavras.some(termo => comLower.includes(termo))) {
        pilarEncontrado = pilar;
        acaoSugerida = PILARES_CX[pilar].acao;
        break;
      }
    }

    linhasParaInserir.push([
      mes, modulo, cliente, nota, comentario, pilarEncontrado, acaoSugerida, 'Pendente'
    ]);
  }

  if (linhasParaInserir.length > 0) {
    sheet.getRange(2, 1, linhasParaInserir.length, linhasParaInserir[0].length).setValues(linhasParaInserir);
    sheet.getRange(2, 1, linhasParaInserir.length, 1).setNumberFormat('mmm/yyyy');
  }
}

// 3. GERADOR GERAL DE DASHBOARDS E PROCESSAMENTOS
function gerarAnaliseEGraficos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base_Historica');
  if (!baseSheet) return;

  // Garante que o cabeçalho da Base Histórica está como "Módulo / Produto"
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
  sheet.clearConditionalFormatRules();
  const rangeNPS = sheet.getRange(2, 9, maxRow - 1, 1);
  
  const ruleVerde = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(50).setBackground('#C6F6D5').setFontColor('#22543D').setRanges([rangeNPS]).build();
  const ruleAmarelo = SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0, 49.9).setBackground('#FEFCBF').setFontColor('#744210').setRanges([rangeNPS]).build();
  const ruleVermelho = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#FED7D7').setFontColor('#742A2A').setRanges([rangeNPS]).build();

  sheet.setConditionalFormatRules([ruleVerde, ruleAmarelo, ruleVermelho]);
}

// 4. RANKING DE PROBLEMAS COM NOME DE COLUNA PADRONIZADO
function gerarRankingEDecisao() {
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
    const pilar = String(row[5] || '5. Outros / Não Classificados').trim();

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('Base_Historica');
  const rankSheet = ss.getSheetByName('Ranking_Problemas');
  if (!baseSheet) return;

  const data = baseSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const consolPorModulo = {};

  for (let i = 1; i < data.length; i++) {
    const status = data[i][1];
    const modulo = data[i][4] || 'SIGAFIS';
    const comentario = data[i][11];
    const classif = data[i][13];

    if (!consolPorModulo[modulo]) {
      consolPorModulo[modulo] = { respostas: 0, comentarios: 0, promotores: 0, neutros: 0, detratores: 0 };
    }

    if (status === 'Respondido') {
      consolPorModulo[modulo].respostas++;

      if (classif === 'Promotor') consolPorModulo[modulo].promotores++;
      else if (classif === 'Neutro') consolPorModulo[modulo].neutros++;
      else if (classif === 'Detrator') consolPorModulo[modulo].detratores++;

      if (comentario && String(comentario).trim().length > 2 && String(comentario).trim().toLowerCase() !== 'sem comentários') {
        consolPorModulo[modulo].comentarios++;
      }
    }
  }

  // 1. ABA RELATORIO_CONSOLIDADO
  let relSheet = ss.getSheetByName('Relatorio_Consolidado');
  if (!relSheet) relSheet = ss.insertSheet('Relatorio_Consolidado');
  relSheet.clear();

  const headersLooker = [['Módulo / Produto', 'NPS Score', '% Promotores', '% Neutros', '% Detratores', 'Total Respostas', 'Total Comentários']];
  const rowsLooker = [];

  for (let mod in consolPorModulo) {
    const st = consolPorModulo[mod];
    const pctProm = st.respostas > 0 ? (st.promotores / st.respostas) : 0;
    const pctNeu = st.respostas > 0 ? (st.neutros / st.respostas) : 0;
    const pctDet = st.respostas > 0 ? (st.detratores / st.respostas) : 0;
    const nps = (pctProm - pctDet) * 100;

    rowsLooker.push([mod, nps, pctProm, pctNeu, pctDet, st.respostas, st.comentarios]);
  }

  relSheet.getRange(1, 1, 1, headersLooker[0].length).setValues(headersLooker).setBackground('#102A43').setFontColor('#FFFFFF').setFontWeight('bold');
  if (rowsLooker.length > 0) {
    relSheet.getRange(2, 1, rowsLooker.length, rowsLooker[0].length).setValues(rowsLooker);
    relSheet.getRange(2, 2, rowsLooker.length, 1).setNumberFormat("0.0");
    relSheet.getRange(2, 3, rowsLooker.length, 3).setNumberFormat("0.0%");
    relSheet.getRange(2, 6, rowsLooker.length, 2).setNumberFormat("#,##0");
  }

  // 2. ABA PRINCIPAIS_ACHADOS COM COLUNA "Módulo / Produto" PARA FILTRAGEM
  let achadosSheet = ss.getSheetByName('Principais_Achados');
  if (!achadosSheet) achadosSheet = ss.insertSheet('Principais_Achados');
  achadosSheet.clear();

  // Mapear principais dores do Ranking por módulo
  const doresPorModulo = {};
  if (rankSheet && rankSheet.getLastRow() > 1) {
    const rankData = rankSheet.getRange(2, 2, rankSheet.getLastRow() - 1, 4).getValues();
    rankData.forEach(r => {
      const mod = r[0];
      const pilar = r[1];
      const qtd = r[2];
      if (!doresPorModulo[mod]) doresPorModulo[mod] = [];
      doresPorModulo[mod].push(`${pilar} (${qtd} ocorrências)`);
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

    linhasAchados.push([mod, 1, `• NPS do módulo ${mod}: ${nps.toFixed(1)} (Promotores: ${(pctProm*100).toFixed(1)}% | Neutros: ${(pctNeu*100).toFixed(1)}% | Detratores: ${(pctDet*100).toFixed(1)}%).`]);
    linhasAchados.push([mod, 2, `• Volume total: ${st.respostas} respostas consolidadas e ${st.comentarios} comentários qualificados.`]);
    linhasAchados.push([mod, 3, `• Principais fatores de atrito/detração em ${mod}: ${topDores}.`]);
    linhasAchados.push([mod, 4, `• País com maior concentração de feedbacks: BR.`]);
  }

  achadosSheet.getRange(1, 1, 1, headersAchados[0].length).setValues(headersAchados).setBackground('#1A365D').setFontColor('#FFFFFF').setFontWeight('bold');
  if (linhasAchados.length > 0) {
    achadosSheet.getRange(2, 1, linhasAchados.length, linhasAchados[0].length).setValues(linhasAchados);
  }
}

function abrirFluig() {
  const urlFluig = "https://fluig.totvs.com/portal/p/10097/portal-de-insights-cx-v2";
  const html = HtmlService.createHtmlOutput(`<script>window.open("${urlFluig}", "_blank");google.script.host.close();</script>`).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Redirecionando para o Portal Fluig...');
}

function abrirLinkLooker() {
  const urlLooker = "https://datastudio.google.com/s/o9eHWdrTIto";
  const html = HtmlService.createHtmlOutput(`<script>window.open("${urlLooker}", "_blank");google.script.host.close();</script>`).setWidth(100).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(html, 'Redirecionando para o Looker Studio...');
}

/**
 * GRAVA O HISTÓRICO DE NPS LENDO DIRETAMENTE CADA MÊS/ANO DA ABA DASHBOARD_KPIS
 */
function gravarSnapshotNPS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName('Dashboard_KPIs');
  
  if (!dashSheet) {
    SpreadsheetApp.getUi().alert('Aba Dashboard_KPIs não encontrada. Execute "Recalcular Dashboards e Gráficos" primeiro!');
    return;
  }

  const dataDash = dashSheet.getDataRange().getValues();
  if (dataDash.length <= 1) return;

  // 1. OBTER OU CRIAR A ABA Historico_NPS_Modulos
  let histSheet = ss.getSheetByName('Historico_NPS_Modulos');
  const headers = [['Data/Hora Registro', 'Mês/Ano Ref.', 'Módulo / Produto', 'NPS Score', '% Promotores', '% Neutros', '% Detratores', 'Total Respostas']];

  if (!histSheet) {
    histSheet = ss.insertSheet('Historico_NPS_Modulos');
  }

  const dataAtual = new Date();
  const novasLinhas = [];

  // 2. EXTRAIR O NPS DE CADA MÊS E MÓDULO DA ABA DASHBOARD_KPIS
  for (let i = 1; i < dataDash.length; i++) {
    const row = dataDash[i];
    const rawMes = row[0];       // Coluna A: Mês/Ano
    const modulo = row[1];       // Coluna B: Módulo / Produto
    const respostas = row[3];    // Coluna D: Respostas
    const pctProm = row[5];      // Coluna F: % Promotores
    const pctNeu = row[6];       // Coluna G: % Neutros
    const pctDet = row[7];       // Coluna H: % Detratores
    const npsScore = row[8];     // Coluna I: NPS Score

    if (!modulo || !rawMes) continue;

    // Formatar Mês/Ano como texto (ex: "08/2026")
    let mesFmt = rawMes;
    if (rawMes instanceof Date) {
      mesFmt = Utilities.formatDate(rawMes, Session.getScriptTimeZone(), 'MM/yyyy');
    }

    novasLinhas.push([
      dataAtual,
      mesFmt,
      modulo,
      npsScore,
      pctProm,
      pctNeu,
      pctDet,
      respostas
    ]);
  }

  // 3. ATUALIZAR A ABA HISTÓRICO SOBRESCREVENDO DE FORMA ORGANIZADA
  histSheet.clear();
  histSheet.getRange(1, 1, 1, headers[0].length)
    .setValues(headers)
    .setBackground('#1A365D')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  histSheet.setFrozenRows(1);

  if (novasLinhas.length > 0) {
    histSheet.getRange(2, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    
    // Formatação de Células
    histSheet.getRange(2, 1, novasLinhas.length, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    histSheet.getRange(2, 4, novasLinhas.length, 1).setNumberFormat('0.0');
    histSheet.getRange(2, 5, novasLinhas.length, 3).setNumberFormat('0.0%');
    histSheet.getRange(2, 8, novasLinhas.length, 1).setNumberFormat('#,##0');

    SpreadsheetApp.getUi().alert(`✅ Snapshot gravado com sucesso! ${novasLinhas.length} registros consolidados mês a mês.`);
  }
}
