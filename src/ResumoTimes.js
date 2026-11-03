/**
 * CADASTRO DE TIMES / GESTORES E ENVIO DO RESUMO DE NPS POR E-MAIL
 *
 * Cadastro_Times    → cada time tem um ou mais módulos e o e-mail do TL (recebe o resumo do time)
 * Cadastro_Gestores → cada gestor acompanha alguns times (ou todos) e recebe o consolidado
 * Log_Envios        → histórico de cada e-mail enviado
 */

const ABA_TIMES = 'Cadastro_Times';
const ABA_GESTORES = 'Cadastro_Gestores';
const ABA_LOG_ENVIOS = 'Log_Envios';
const CAB_TIMES = ['Time', 'Módulos / Produtos', 'TL', 'E-mail do TL', 'Ativo'];
const CAB_GESTORES = ['Nome', 'E-mail', 'Times acompanhados (vazio = todos)', 'Ativo'];
const CAB_LOG = ['Data/Hora', 'Tipo', 'Destinatário', 'E-mail', 'Período', 'NPS', 'Assunto', 'Modo'];

const NOMES_MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const MESES_CURTOS = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

// ============================================================
// TELAS
// ============================================================

// As telas ficam no Painel (Painel.html + módulos Painel*.html); ver abrirPainel() em Código.js.

// ============================================================
// CADASTRO
// ============================================================

function abaCadastro_(nome, cab, cor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setBackground(cor || '#102A43').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function lerAba_(nome, ncol) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ncol).getValues().filter(r => r.some(c => String(c).trim() !== ''));
}

const ativo_ = v => v === true || /^(sim|s|true|1|x|ativo)$/i.test(String(v).trim());
const listaTexto_ = s => String(s == null ? '' : s).split(/[,;\n]/).map(x => x.trim()).filter(Boolean);

function lerCadastro_() {
  const times = lerAba_(ABA_TIMES, CAB_TIMES.length).map(r => ({
    time: String(r[0]).trim(), modulos: listaTexto_(r[1]), tl: String(r[2]).trim(), email: String(r[3]).trim(), ativo: ativo_(r[4])
  }));
  const gestores = lerAba_(ABA_GESTORES, CAB_GESTORES.length).map(r => ({
    nome: String(r[0]).trim(), email: String(r[1]).trim(), times: listaTexto_(r[2]), ativo: ativo_(r[3])
  }));
  return { times: times, gestores: gestores };
}

/** Dados para a tela de cadastro. */
function getCadastro() {
  const c = lerCadastro_();
  const modulos = {};
  const base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Base_Historica');
  if (base && base.getLastRow() > 1) {
    base.getRange(2, 5, base.getLastRow() - 1, 1).getValues().forEach(r => { const m = String(r[0]).trim(); if (m) modulos[m] = (modulos[m] || 0) + 1; });
  }
  if (!Object.keys(modulos).length) modulos[FILTROS_PADRAO.produto] = 0;
  return JSON.stringify({ times: c.times, gestores: c.gestores, modulos: Object.keys(modulos).sort((a, b) => modulos[b] - modulos[a]) });
}

/** Valida e-mails separados por vírgula/; — retorna os válidos e os inválidos. */
function validarEmails_(s) {
  const lista = listaTexto_(s);
  const re = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/;
  return { lista: lista, invalidos: lista.filter(e => !re.test(e)) };
}

/** Valida o cadastro vindo da tela. Retorna lista de erros (vazia = ok). */
function validarCadastro_(c) {
  const erros = [];
  const nomes = new Set();
  (c.times || []).forEach((t, i) => {
    const ref = t.time ? `Time "${escHtml_(t.time)}"` : `Time da linha ${i + 1}`;
    if (!t.time) erros.push(`${ref}: informe o nome do time.`);
    else if (nomes.has(t.time.toLowerCase())) erros.push(`${ref}: nome repetido.`);
    nomes.add(String(t.time).toLowerCase());
    if (!(t.modulos || []).length) erros.push(`${ref}: informe ao menos um módulo.`);
    const e = validarEmails_(t.email);
    if (!e.lista.length) erros.push(`${ref}: informe o e-mail do TL.`);
    if (e.invalidos.length) erros.push(`${ref}: e-mail inválido (${escHtml_(e.invalidos.join(', '))}).`);
  });
  (c.gestores || []).forEach((g, i) => {
    const ref = g.nome ? `Gestor "${escHtml_(g.nome)}"` : `Gestor da linha ${i + 1}`;
    if (!g.nome) erros.push(`${ref}: informe o nome.`);
    const e = validarEmails_(g.email);
    if (!e.lista.length) erros.push(`${ref}: informe o e-mail.`);
    if (e.invalidos.length) erros.push(`${ref}: e-mail inválido (${escHtml_(e.invalidos.join(', '))}).`);
    const desconhecidos = (g.times || []).filter(t => !nomes.has(t.toLowerCase()));
    if (desconhecidos.length) erros.push(`${ref}: time não cadastrado (${escHtml_(desconhecidos.join(', '))}).`);
  });
  return erros;
}

function salvarCadastro(json) {
  const c = JSON.parse(json);
  const erros = validarCadastro_(c);
  if (erros.length) return { success: false, message: erros.slice(0, 6).join('<br>') };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { success: false, message: 'Outra gravação em andamento. Tente de novo.' };
  try {
    const gravar = (nome, cab, cor, linhas) => {
      const sh = abaCadastro_(nome, cab, cor);
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cab.length).clearContent();
      if (linhas.length) sh.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
      sh.autoResizeColumns(1, cab.length);
    };
    gravar(ABA_TIMES, CAB_TIMES, '#13325e', c.times.map(t => [t.time, t.modulos.join(', '), t.tl || '', validarEmails_(t.email).lista.join(', '), t.ativo ? 'Sim' : 'Não']));
    gravar(ABA_GESTORES, CAB_GESTORES, '#276749', c.gestores.map(g => [g.nome, validarEmails_(g.email).lista.join(', '), (g.times || []).join(', '), g.ativo ? 'Sim' : 'Não']));
    return { success: true, message: `Cadastro salvo: ${c.times.length} time(s) e ${c.gestores.length} gestor(es).` };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// CÁLCULO DO RESUMO (funções puras — cobertas por testes)
// ============================================================

/** Converte linhas da Base_Historica (sem cabeçalho) em registros leves, aplicando tipo de cliente e país padrão. */
function registrosDeBase_(valores, filtros) {
  const out = [];
  valores.forEach(r => {
    const d = Object.prototype.toString.call(r[0]) === '[object Date]' ? r[0] : parseData_(r[0]);
    if (!d || isNaN(d.getTime())) return;
    const pais = String(r[2] || '').trim(), tipo = String(r[9] || '').trim();
    if (filtros && filtros.pais && pais && pais !== filtros.pais) return;
    if (filtros && filtros.tipoCliente && tipo && normalizarTexto_(tipo) !== normalizarTexto_(filtros.tipoCliente)) return;
    const com = comentarioValido_(r[11]) ? String(r[11]).trim() : '';
    out.push({
      dia: fmtDia_(d), mes: fmtMes_(d), resp: String(r[1]).trim() === 'Respondido',
      mod: String(r[4] || 'SIGAFIS').trim(), nota: r[7] === '' || r[7] == null ? null : Number(r[7]),
      cl: String(r[13] || ''), razao: String(r[10] || '').trim(), com: com,
      pil: com ? classificarComentario_(com).pilar : ''
    });
  });
  return out;
}

/** Monta o objeto de período a partir de {mesInicio, mesFim} (yyyy-MM). Sem argumento: padrão (jan até o mês anterior). */
function periodoDeMeses_(sel) {
  if (!sel || !/^\d{4}-\d{2}$/.test(sel.mesInicio || '') || !/^\d{4}-\d{2}$/.test(sel.mesFim || '')) return periodoPadrao_(new Date());
  let a = sel.mesInicio, b = sel.mesFim;
  if (a > b) { const x = a; a = b; b = x; }
  const [ya, ma] = a.split('-').map(Number), [yb, mb] = b.split('-').map(Number);
  const fim = new Date(yb, mb, 0);
  return { inicio: `01/${String(ma).padStart(2, '0')}/${ya}`, fim: `${String(fim.getDate()).padStart(2, '0')}/${String(mb).padStart(2, '0')}/${yb}`, mesInicio: a, mesFim: b };
}

/** "jul a out/2026", "nov/2025 a fev/2026" ou "outubro/2026" (um mês só). */
function rotuloPeriodo_(periodo) {
  const [ya, ma] = periodo.mesInicio.split('-'), [yb, mb] = periodo.mesFim.split('-');
  if (periodo.mesInicio === periodo.mesFim) return nomeMes_(periodo.mesFim);
  if (ya === yb) return `${MESES_CURTOS[+ma - 1]} a ${MESES_CURTOS[+mb - 1]}/${yb}`;
  return `${MESES_CURTOS[+ma - 1]}/${ya} a ${MESES_CURTOS[+mb - 1]}/${yb}`;
}

function mesAnterior_(m) {
  let [y, mm] = m.split('-').map(Number);
  mm--; if (mm === 0) { mm = 12; y--; }
  return `${y}-${String(mm).padStart(2, '0')}`;
}

function estatisticas_(regs) {
  const s = { abord: regs.length, resp: 0, P: 0, N: 0, D: 0, com: 0 };
  regs.forEach(r => {
    if (r.resp) { s.resp++; if (r.cl === 'Promotor') s.P++; else if (r.cl === 'Neutro') s.N++; else if (r.cl === 'Detrator') s.D++; }
    if (r.com) s.com++;
  });
  const b = s.P + s.N + s.D;
  s.base = b;
  s.nps = b ? (s.P - s.D) / b * 100 : null;
  s.pP = b ? s.P / b : 0; s.pN = b ? s.N / b : 0; s.pD = b ? s.D / b : 0;
  s.ader = s.abord ? s.resp / s.abord : 0;
  return s;
}

/** Resumo de um conjunto de módulos no período (mesIni..mesFim, formato yyyy-MM). */
function resumoModulos_(regs, modulos, mesIni, mesFim) {
  const set = new Set((modulos || []).map(m => m.toUpperCase()));
  const doTime = regs.filter(r => !set.size || set.has(r.mod.toUpperCase()));
  const periodo = doTime.filter(r => r.mes >= mesIni && r.mes <= mesFim);
  const mesAnt = mesAnterior_(mesFim);
  const s = estatisticas_(periodo);
  const ult = estatisticas_(doTime.filter(r => r.mes === mesFim));
  const ant = estatisticas_(doTime.filter(r => r.mes === mesAnt));

  // série mensal (para a mini-tendência no e-mail)
  const serie = [];
  let m = mesIni;
  while (m <= mesFim) {
    const st = estatisticas_(periodo.filter(r => r.mes === m));
    serie.push({ mes: m, nps: st.nps, base: st.base });
    let [y, mm] = m.split('-').map(Number); mm++; if (mm > 12) { mm = 1; y++; }
    m = `${y}-${String(mm).padStart(2, '0')}`;
  }
  const validos = serie.filter(x => x.nps != null && x.base >= 3);
  const melhor = validos.length ? validos.reduce((a, b) => b.nps > a.nps ? b : a) : null;
  const pior = validos.length ? validos.reduce((a, b) => b.nps < a.nps ? b : a) : null;

  // causas raiz (comentários do período)
  const g = {};
  periodo.filter(r => r.com).forEach(r => {
    const p = r.pil || PILAR_OUTROS;
    g[p] = g[p] || { pilar: p, qtd: 0, det: 0 };
    g[p].qtd++; if (r.cl === 'Detrator') g[p].det++;
  });
  const causas = Object.values(g).sort((a, b) => b.det - a.det || b.qtd - a.qtd);

  // voz do cliente: detratores do último mês primeiro, depois do período
  const vistos = new Set();
  const unico = r => { const k = normalizarTexto_(r.com); if (vistos.has(k)) return false; vistos.add(k); return true; };
  const detratores = periodo.filter(r => r.com && r.cl === 'Detrator')
    .sort((a, b) => (b.mes === mesFim) - (a.mes === mesFim) || (a.nota - b.nota) || b.dia.localeCompare(a.dia));
  const totalDetCom = detratores.length;
  const detUnicos = detratores.filter(unico);
  const elogios = periodo.filter(r => r.com && r.cl === 'Promotor').sort((a, b) => b.dia.localeCompare(a.dia));

  return {
    modulos: modulos, total: s, ultimo: Object.assign({ mes: mesFim }, ult), anterior: Object.assign({ mes: mesAnt }, ant),
    variacao: ult.nps != null && ant.nps != null ? ult.nps - ant.nps : null,
    serie: serie, melhor: melhor, pior: pior, causas: causas,
    detratores: detUnicos.slice(0, 5), totalDetratoresComentario: totalDetCom,
    elogio: elogios[0] || null
  };
}

/** Monta a lista de e-mails (um por TL e um por gestor) a partir do cadastro. */
function montarEnvios_(cadastro, regs, periodo, links) {
  const envios = [];
  const timesAtivos = cadastro.times.filter(t => t.ativo && t.time);
  const resumos = {};
  timesAtivos.forEach(t => { resumos[t.time] = resumoModulos_(regs, t.modulos, periodo.mesInicio, periodo.mesFim); });

  timesAtivos.forEach(t => {
    const r = resumos[t.time];
    envios.push({
      tipo: 'TL', nome: t.tl || t.time, time: t.time, para: validarEmails_(t.email).lista.join(','),
      assunto: `NPS ${t.time} · ${rotuloPeriodo_(periodo)}: ${fmtNps_(r.total.nps)} · ${mesCurto_(periodo.mesFim)} ${fmtNps_(r.ultimo.nps)}${setaVar_(r.variacao)}`,
      nps: r.total.nps, npsMes: r.ultimo.nps, html: htmlEmailTime_(t, r, periodo, links)
    });
  });

  cadastro.gestores.filter(g => g.ativo && g.email).forEach(g => {
    const alvo = new Set(g.times.map(x => x.toLowerCase()));
    const seus = timesAtivos.filter(t => !alvo.size || alvo.has(t.time.toLowerCase()));
    if (!seus.length) return;
    const modulos = [...new Set([].concat(...seus.map(t => t.modulos)))];
    const consolidado = resumoModulos_(regs, modulos, periodo.mesInicio, periodo.mesFim);
    envios.push({
      tipo: 'Gestor', nome: g.nome, time: seus.map(t => t.time).join(', '), para: validarEmails_(g.email).lista.join(','),
      assunto: `Resumo NPS · ${rotuloPeriodo_(periodo)} · ${seus.length} time${seus.length > 1 ? 's' : ''} · NPS ${fmtNps_(consolidado.total.nps)}`,
      nps: consolidado.total.nps, npsMes: consolidado.ultimo.nps, html: htmlEmailGestor_(g, seus.map(t => ({ t: t, r: resumos[t.time] })), consolidado, periodo, links)
    });
  });
  return envios;
}

// ============================================================
// HTML DOS E-MAILS (inline styles, compatível com Gmail/Outlook)
// ============================================================

const EM = {
  font: "font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif",
  ink: '#0f172a', ink2: '#334155', muted: '#64748b', line: '#e5e9f2', soft: '#f6f8fc',
  navy: '#0b1f3a', blue: '#2563eb', prom: '#16a34a', neu: '#f59e0b', det: '#e5484d'
};

function nomeMes_(m) { const [y, mm] = m.split('-'); return `${NOMES_MESES[+mm - 1]}/${y}`; }
function mesCurto_(m) { const [y, mm] = m.split('-'); return `${MESES_CURTOS[+mm - 1]}/${y.slice(2)}`; }
function fmtNps_(n) { return n == null ? '—' : (Math.round(n * 10) / 10).toFixed(1).replace('.', ','); }
function fmtPct_(n) { return (Math.round(n * 1000) / 10).toFixed(1).replace('.', ',') + '%'; }
function fmtInt_(n) { return String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function setaVar_(v) { return v == null ? '' : ` (${v >= 0 ? '▲' : '▼'} ${fmtNps_(Math.abs(v))})`; }
function escHtml_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

function zonaNps_(n) {
  if (n == null) return { nome: 'Sem respostas', cor: '#64748b', bg: '#f1f5f9' };
  if (n >= 75) return { nome: 'Excelente', cor: '#067647', bg: '#e3f7ea' };
  if (n >= 50) return { nome: 'Muito bom', cor: '#15803d', bg: '#e9f8ef' };
  if (n >= 0) return { nome: 'Razoável', cor: '#a15c07', bg: '#fff4dc' };
  return { nome: 'Crítico', cor: '#b42318', bg: '#fdeced' };
}

function pilarCurto_(p) { return String(p).replace(/^\d\.\s*/, ''); }

function blocoVariacao_(v, rotuloAnt) {
  if (v == null) return `<span style="color:${EM.muted};font-size:13px">sem base de comparação</span>`;
  const cor = v >= 0 ? EM.prom : EM.det, bg = v >= 0 ? '#e9f8ef' : '#fdeced';
  return `<span style="display:inline-block;background:${bg};color:${cor};font-weight:700;font-size:13px;padding:4px 10px;border-radius:999px">${v >= 0 ? '▲' : '▼'} ${fmtNps_(Math.abs(v))} pts</span>
          <span style="color:${EM.muted};font-size:12px;margin-left:4px">vs ${rotuloAnt}</span>`;
}

function barraMix_(s) {
  if (!s.base) return '';
  const w = x => Math.max(0, Math.round(x * 1000) / 10);
  const cel = (pct, cor) => pct > 0 ? `<td style="background:${cor};height:10px;width:${pct}%;font-size:0;line-height:0">&nbsp;</td>` : '';
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-radius:6px;overflow:hidden"><tr>
    ${cel(w(s.pP), EM.prom)}${cel(w(s.pN), EM.neu)}${cel(w(s.pD), EM.det)}</tr></table>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px"><tr>
    ${[['Promotores', s.pP, s.P, EM.prom], ['Neutros', s.pN, s.N, EM.neu], ['Detratores', s.pD, s.D, EM.det]].map(([n, p, q, c]) =>
      `<td style="${EM.font};font-size:12px;color:${EM.muted};width:33%"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${c};margin-right:6px"></span>${n}<br>
       <span style="font-size:18px;font-weight:700;color:${EM.ink}">${fmtPct_(p)}</span> <span style="font-size:12px">${fmtInt_(q)}</span></td>`).join('')}
  </tr></table>`;
}

/** Mini-gráfico de barras mês a mês, feito com tabela (renderiza em qualquer cliente de e-mail). */
function sparkMeses_(serie) {
  const validos = serie.filter(x => x.nps != null);
  if (validos.length < 2) return '';
  const max = Math.max(10, ...validos.map(x => Math.abs(x.nps)));
  const alturaMax = 56;
  const cols = serie.map(x => {
    const h = x.nps == null ? 0 : Math.max(3, Math.round(Math.abs(x.nps) / max * alturaMax));
    const ultimo = x === serie[serie.length - 1];
    const cor = x.nps == null ? '#e2e8f0' : x.nps < 0 ? EM.det : ultimo ? EM.blue : '#a9c1f5';
    return `<td valign="bottom" align="center" style="${EM.font};padding:0 3px;width:${Math.floor(100 / serie.length)}%">
      <div style="font-size:10px;color:${ultimo ? EM.blue : EM.ink2};font-weight:${ultimo ? 800 : 600};margin-bottom:3px">${x.nps == null ? '' : Math.round(x.nps)}</div>
      <div style="background:${cor};height:${h}px;border-radius:3px 3px 0 0"></div></td>`;
  }).join('');
  const rot = serie.map(x => `<td align="center" style="${EM.font};font-size:10px;color:${EM.muted};padding-top:4px">${MESES_CURTOS[+x.mes.split('-')[1] - 1]}</td>`).join('');
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px"><tr style="height:${alturaMax + 16}px">${cols}</tr><tr>${rot}</tr></table>`;
}

function tabelaCausas_(causas, limite) {
  if (!causas.length) return `<p style="${EM.font};font-size:13px;color:${EM.muted};margin:0">Sem comentários no período.</p>`;
  const max = Math.max(...causas.map(c => c.qtd));
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0">` + causas.slice(0, limite || 4).map(c => `
    <tr><td style="${EM.font};font-size:13px;color:${EM.ink};padding:6px 0;width:40%"><b>${escHtml_(pilarCurto_(c.pilar))}</b></td>
      <td style="padding:6px 10px;width:32%"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;border-radius:4px"><tr>
        ${c.det ? `<td style="background:${EM.det};height:8px;width:${Math.max(2, Math.round(c.det / max * 100))}%;font-size:0;border-radius:4px 0 0 4px">&nbsp;</td>` : ''}
        ${c.qtd - c.det ? `<td style="background:#c7d2fe;height:8px;width:${Math.round((c.qtd - c.det) / max * 100)}%;font-size:0">&nbsp;</td>` : ''}
        <td style="font-size:0">&nbsp;</td></tr></table></td>
      <td align="right" style="${EM.font};font-size:12px;color:${EM.muted};white-space:nowrap;padding:6px 0">${c.qtd} coment. · <span style="color:${EM.det};font-weight:600">${c.det} detr.</span></td></tr>`).join('') + `</table>`;
}

function citacao_(r, cor) {
  const txt = r.com.length > 260 ? r.com.slice(0, 257) + '…' : r.com;
  return `<tr><td style="padding:0 0 10px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${EM.soft};border-left:3px solid ${cor};border-radius:0 8px 8px 0"><tr><td style="${EM.font};padding:10px 14px">
    <div style="font-size:13.5px;color:${EM.ink};line-height:1.5">“${escHtml_(txt)}”</div>
    <div style="font-size:11.5px;color:${EM.muted};margin-top:6px">${escHtml_(r.razao || 'Cliente')} · nota <b style="color:${cor}">${r.nota}</b> · ${r.dia.split('-').reverse().join('/')}</div>
  </td></tr></table></td></tr>`;
}

function cartao_(titulo, corpo) {
  return `<tr><td style="padding:0 24px 16px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${EM.line};border-radius:12px"><tr><td style="${EM.font};padding:16px 18px">
    <div style="font-size:11px;letter-spacing:1.4px;text-transform:uppercase;color:${EM.muted};font-weight:700;margin-bottom:10px">${titulo}</div>${corpo}
  </td></tr></table></td></tr>`;
}

function moldura_(eyebrow, titulo, sub, corpo, links) {
  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#eef1f6">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef1f6;padding:24px 0"><tr><td align="center">
  <table role="presentation" width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden">
    <tr><td style="background:${EM.navy};background-image:linear-gradient(115deg,#0b1f3a 0%,#13325e 70%,#1e40af 130%);padding:22px 24px;${EM.font}">
      <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#93b0e0;font-weight:700">${eyebrow}</div>
      <div style="font-size:21px;font-weight:700;color:#ffffff;margin-top:6px">${titulo}</div>
      <div style="font-size:13px;color:#bcd0f0;margin-top:6px">${sub}</div>
    </td></tr>
    <tr><td style="height:20px"></td></tr>
    ${corpo}
    <tr><td style="padding:4px 24px 22px;${EM.font}">
      ${links && links.planilha ? `<a href="${links.planilha}" style="display:inline-block;background:${EM.blue};color:#fff;text-decoration:none;font-weight:600;font-size:14px;padding:11px 18px;border-radius:9px">Abrir dashboard completo</a>` : ''}
    </td></tr>
    <tr><td style="background:${EM.soft};padding:14px 24px;${EM.font};font-size:11.5px;color:${EM.muted};line-height:1.5">
      Fonte: Portal de Insights CX (Fluig) · pesquisa ${escHtml_(FILTROS_PADRAO.pesquisa)} · clientes ${escHtml_(FILTROS_PADRAO.tipoCliente)} · ${escHtml_(FILTROS_PADRAO.pais)}.<br>
      NPS = % promotores (9–10) − % detratores (0–6). Faixas: crítico &lt; 0 · razoável 0–49 · muito bom 50–74 · excelente ≥ 75.
    </td></tr>
  </table></td></tr></table></body></html>`;
}

function destaqueNps_(r, periodo) {
  const zu = zonaNps_(r.ultimo.nps), zt = zonaNps_(r.total.nps);
  return `<tr><td style="padding:0 24px 16px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
    <td width="50%" valign="top" style="padding-right:8px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${zu.bg};border-radius:12px"><tr><td style="${EM.font};padding:16px 18px">
      <div style="font-size:12px;color:${EM.ink2};font-weight:600">NPS de ${nomeMes_(periodo.mesFim)}</div>
      <div style="font-size:40px;line-height:1.1;font-weight:800;color:${zu.cor};margin:6px 0 8px">${fmtNps_(r.ultimo.nps)}</div>
      ${blocoVariacao_(r.variacao, mesCurto_(r.anterior.mes))}
      <div style="font-size:12px;color:${EM.muted};margin-top:8px">${fmtInt_(r.ultimo.base)} respostas com nota · ${zu.nome.toLowerCase()}</div>
    </td></tr></table></td>
    <td width="50%" valign="top" style="padding-left:8px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${EM.line};border-radius:12px"><tr><td style="${EM.font};padding:16px 18px">
      <div style="font-size:12px;color:${EM.ink2};font-weight:600">Período ${mesCurto_(periodo.mesInicio)} – ${mesCurto_(periodo.mesFim)}</div>
      <div style="font-size:40px;line-height:1.1;font-weight:800;color:${zt.cor};margin:6px 0 8px">${fmtNps_(r.total.nps)}</div>
      <div style="font-size:12px;color:${EM.muted}">${fmtInt_(r.total.resp)} respostas de ${fmtInt_(r.total.abord)} abordagens<br>aderência ${fmtPct_(r.total.ader)} · ${fmtInt_(r.total.com)} comentários</div>
    </td></tr></table></td>
  </tr></table></td></tr>`;
}

function destaquesTexto_(r) {
  const itens = [];
  if (r.melhor && r.pior && r.melhor.mes !== r.pior.mes)
    itens.push(`Melhor mês: <b>${mesCurto_(r.melhor.mes)}</b> (${fmtNps_(r.melhor.nps)}) · pior mês: <b>${mesCurto_(r.pior.mes)}</b> (${fmtNps_(r.pior.nps)}).`);
  const dor = r.causas.find(c => c.det > 0);
  if (dor) itens.push(`Principal fator de detração: <b>${escHtml_(pilarCurto_(dor.pilar))}</b> (${dor.det} detrator${dor.det > 1 ? 'es' : ''}).`);
  if (r.totalDetratoresComentario) itens.push(`<b>${r.totalDetratoresComentario}</b> detrator${r.totalDetratoresComentario > 1 ? 'es deixaram' : ' deixou'} comentário — vale contato de retorno.`);
  if (!itens.length) return '';
  return itens.map(t => `<div style="${EM.font};font-size:13px;color:${EM.ink2};line-height:1.5;padding:4px 0">• ${t}</div>`).join('');
}

function htmlEmailTime_(t, r, periodo, links) {
  const ola = t.tl ? `Olá, ${escHtml_(t.tl.split(' ')[0])}! ` : '';
  let corpo = `<tr><td style="padding:0 24px 16px;${EM.font};font-size:14px;color:${EM.ink2};line-height:1.55">${ola}Segue o resumo do NPS do time <b>${escHtml_(t.time)}</b> (${escHtml_(t.modulos.join(', '))}) no período de <b>${rotuloPeriodo_(periodo)}</b>.</td></tr>`;
  corpo += destaqueNps_(r, periodo);
  corpo += cartao_('Composição das respostas no período', barraMix_(r.total) || `<span style="${EM.font};font-size:13px;color:${EM.muted}">Sem respostas no período.</span>`);
  const sp = sparkMeses_(r.serie), dt = destaquesTexto_(r);
  if (sp || dt) corpo += cartao_('Evolução mês a mês', sp + (dt ? `<div style="margin-top:12px">${dt}</div>` : ''));
  corpo += cartao_('Causas raiz nos comentários', tabelaCausas_(r.causas, 4));
  if (r.detratores.length) corpo += cartao_('Voz do cliente · detratores', `<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${r.detratores.slice(0, 3).map(x => citacao_(x, EM.det)).join('')}</table>`);
  if (r.elogio) corpo += cartao_('Voz do cliente · elogio', `<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${citacao_(r.elogio, EM.prom)}</table>`);
  return moldura_('NPS Produto · Protheus', `NPS do time ${escHtml_(t.time)}`, `${escHtml_(t.modulos.join(', '))} · ${rotuloPeriodo_(periodo)}`, corpo, links);
}

function htmlEmailGestor_(g, itens, consolidado, periodo, links) {
  const ola = `Olá, ${escHtml_(g.nome.split(' ')[0])}! `;
  let corpo = `<tr><td style="padding:0 24px 16px;${EM.font};font-size:14px;color:${EM.ink2};line-height:1.55">${ola}Segue o consolidado do NPS de <b>${itens.length} time${itens.length > 1 ? 's' : ''}</b> no período de <b>${rotuloPeriodo_(periodo)}</b>.</td></tr>`;
  corpo += destaqueNps_(consolidado, periodo);

  const ord = itens.slice().sort((a, b) => (b.r.ultimo.nps == null ? -999 : b.r.ultimo.nps) - (a.r.ultimo.nps == null ? -999 : a.r.ultimo.nps));
  const th = (t, al) => `<td align="${al || 'left'}" style="${EM.font};font-size:10.5px;letter-spacing:.8px;text-transform:uppercase;color:${EM.muted};font-weight:700;padding:8px 6px;border-bottom:1px solid ${EM.line}">${t}</td>`;
  const linhas = ord.map(({ t, r }) => {
    const z = zonaNps_(r.ultimo.nps), v = r.variacao;
    return `<tr>
      <td style="${EM.font};font-size:13px;color:${EM.ink};padding:10px 6px;border-bottom:1px solid ${EM.line}"><b>${escHtml_(t.time)}</b><div style="font-size:11.5px;color:${EM.muted}">${escHtml_(t.tl || '')}${t.tl ? ' · ' : ''}${escHtml_(t.modulos.join(', '))}</div></td>
      <td align="center" style="padding:10px 6px;border-bottom:1px solid ${EM.line}"><span style="${EM.font};display:inline-block;background:${z.bg};color:${z.cor};font-weight:800;font-size:14px;padding:4px 10px;border-radius:8px">${fmtNps_(r.ultimo.nps)}</span></td>
      <td align="center" style="${EM.font};font-size:12.5px;font-weight:700;padding:10px 6px;border-bottom:1px solid ${EM.line};color:${v == null ? EM.muted : v >= 0 ? EM.prom : EM.det}">${v == null ? '—' : (v >= 0 ? '▲ ' : '▼ ') + fmtNps_(Math.abs(v))}</td>
      <td align="center" style="${EM.font};font-size:13px;color:${EM.ink2};padding:10px 6px;border-bottom:1px solid ${EM.line}">${fmtNps_(r.total.nps)}</td>
      <td align="center" style="${EM.font};font-size:13px;color:${EM.ink2};padding:10px 6px;border-bottom:1px solid ${EM.line}">${fmtInt_(r.total.base)}</td>
      <td align="center" style="${EM.font};font-size:13px;color:${EM.det};font-weight:600;padding:10px 6px;border-bottom:1px solid ${EM.line}">${fmtInt_(r.total.D)}</td></tr>`;
  }).join('');
  corpo += cartao_('Ranking dos times', `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>
    ${th('Time')}${th(mesCurto_(periodo.mesFim), 'center')}${th('Var.', 'center')}${th('Período', 'center')}${th('Respostas', 'center')}${th('Detratores', 'center')}</tr>${linhas}</table>`);

  const atencao = ord.filter(x => (x.r.variacao != null && x.r.variacao <= -10) || (x.r.ultimo.nps != null && x.r.ultimo.nps < 0));
  const alertas = atencao.map(({ t, r }) => {
    const dor = r.causas.find(c => c.det > 0);
    return `<div style="${EM.font};font-size:13px;color:${EM.ink2};line-height:1.5;padding:4px 0">• <b>${escHtml_(t.time)}</b>: NPS ${fmtNps_(r.ultimo.nps)}${r.variacao != null ? ` (${r.variacao >= 0 ? '+' : '−'}${fmtNps_(Math.abs(r.variacao))} pts)` : ''}${dor ? ` · principal dor: ${escHtml_(pilarCurto_(dor.pilar))}` : ''}.</div>`;
  }).join('');
  const dt = destaquesTexto_(consolidado);
  if (alertas || dt) corpo += cartao_(alertas ? 'Pontos de atenção' : 'Destaques', alertas + dt);

  corpo += cartao_('Composição das respostas · consolidado', barraMix_(consolidado.total) || '');
  corpo += cartao_('Causas raiz · consolidado', tabelaCausas_(consolidado.causas, 5));
  if (consolidado.detratores.length) corpo += cartao_('Voz do cliente · detratores', `<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${consolidado.detratores.slice(0, 2).map(x => citacao_(x, EM.det)).join('')}</table>`);
  return moldura_('NPS Produto · Protheus · Gestão', 'Resumo executivo de NPS', `${itens.length} time${itens.length > 1 ? 's' : ''} · ${rotuloPeriodo_(periodo)}`, corpo, links);
}

// ============================================================
// ENVIO
// ============================================================

function prepararEnvios_(sel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName('Base_Historica');
  const valores = base && base.getLastRow() > 1 ? base.getRange(2, 1, base.getLastRow() - 1, CABECALHOS_BASE.length).getValues() : [];
  const regs = registrosDeBase_(valores, FILTROS_PADRAO);
  const periodo = periodoDeMeses_(sel);
  const envios = montarEnvios_(lerCadastro_(), regs, periodo, { planilha: ss.getUrl() });
  const meses = [...new Set(regs.map(r => r.mes))].sort();
  return { periodo: periodo, envios: envios, meses: meses, temDados: regs.length > 0, temMesFim: regs.some(r => r.mes === periodo.mesFim) };
}

/** Prévia para a tela de envio. */
function previewResumo(periodoJson) {
  const p = prepararEnvios_(periodoJson ? JSON.parse(periodoJson) : null);
  return JSON.stringify({
    periodo: p.periodo, rotulo: rotuloPeriodo_(p.periodo), meses: p.meses, temDados: p.temDados, temMesFim: p.temMesFim, meuEmail: Session.getEffectiveUser().getEmail(),
    envios: p.envios.map(e => ({ tipo: e.tipo, nome: e.nome, time: e.time, para: e.para, assunto: e.assunto, nps: e.nps, npsMes: e.npsMes, html: e.html }))
  });
}

/**
 * modo = 'teste' → envia todos os e-mails só para você (assunto com [TESTE])
 * modo = 'real'  → envia para cada TL e gestor cadastrado
 */
function enviarResumo(modo, rotulo, periodoJson) {
  const p = prepararEnvios_(periodoJson ? JSON.parse(periodoJson) : null);
  if (!p.envios.length) return { success: false, message: 'Nenhum destinatário ativo no cadastro.' };
  const eu = Session.getEffectiveUser().getEmail();
  const teste = modo !== 'real';
  const restante = MailApp.getRemainingDailyQuota();
  const necessario = p.envios.reduce((a, e) => a + (teste ? 1 : e.para.split(',').length), 0);
  if (restante < necessario) return { success: false, message: `Cota diária de e-mail insuficiente (restam ${restante}, precisa de ${necessario}).` };

  const log = abaCadastro_(ABA_LOG_ENVIOS, CAB_LOG, '#334155');
  const agora = new Date(), linhasLog = [];
  p.envios.forEach(e => {
    MailApp.sendEmail({
      to: teste ? eu : e.para,
      subject: (teste ? `[TESTE → ${e.para}] ` : '') + e.assunto,
      htmlBody: e.html,
      name: 'NPS Protheus'
    });
    linhasLog.push([agora, e.tipo, e.nome, teste ? eu : e.para, `${p.periodo.inicio} a ${p.periodo.fim}`, e.nps == null ? '' : Math.round(e.nps * 10) / 10, e.assunto, teste ? 'Teste' : (rotulo || 'Manual')]);
  });
  log.getRange(log.getLastRow() + 1, 1, linhasLog.length, CAB_LOG.length).setValues(linhasLog);
  log.getRange(2, 1, log.getLastRow() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  return {
    success: true,
    message: teste
      ? `${p.envios.length} e-mail(s) de teste enviados para ${eu}.`
      : `${p.envios.length} e-mail(s) enviados: ${p.envios.filter(e => e.tipo === 'TL').length} para TLs e ${p.envios.filter(e => e.tipo === 'Gestor').length} para gestores.`
  };
}

// ============================================================
// ENVIO AUTOMÁTICO (2º DIA ÚTIL)
// ============================================================

const HANDLER_ENVIO_AUTO = 'envioAutomaticoResumo';
const HORA_ENVIO_AUTO = 8;

function segundoDiaUtil_(ano, mes) {
  const d = primeiroDiaUtil_(ano, mes);
  d.setDate(d.getDate() + 1);
  while (!ehDiaUtil_(d)) d.setDate(d.getDate() + 1);
  return d;
}

function proximoEnvioAuto_(hoje) {
  hoje = hoje || new Date();
  let d = segundoDiaUtil_(hoje.getFullYear(), hoje.getMonth());
  const jaFoi = PropertiesService.getDocumentProperties().getProperty('ENVIO_AUTO_MES') === periodoPadrao_(hoje).mesFim;
  if (fmtDia_(d) < fmtDia_(hoje) || (fmtDia_(d) === fmtDia_(hoje) && (hoje.getHours() >= HORA_ENVIO_AUTO + 1 || jaFoi)))
    d = segundoDiaUtil_(hoje.getFullYear(), hoje.getMonth() + 1);
  return d;
}

/**
 * Disparado todo dia pelo gatilho; só age no 2º dia útil.
 * Envia o resumo do mês fechado para TLs e gestores — uma vez por mês.
 * Se a base ainda não tiver dados do mês fechado, não envia e avisa o dono.
 */
function envioAutomaticoResumo() {
  const hoje = new Date();
  if (fmtDia_(hoje) !== fmtDia_(segundoDiaUtil_(hoje.getFullYear(), hoje.getMonth()))) return;

  const props = PropertiesService.getDocumentProperties();
  const periodo = periodoPadrao_(hoje);
  if (props.getProperty('ENVIO_AUTO_MES') === periodo.mesFim) return;

  const eu = Session.getEffectiveUser().getEmail();
  const p = prepararEnvios_();
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const avisar = (assunto, corpo) => MailApp.sendEmail({ to: eu, subject: assunto, name: 'NPS Protheus',
    htmlBody: `<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;max-width:560px">${corpo}<p style="margin-top:16px"><a href="${url}" style="color:#2563eb;font-weight:600">Abrir a planilha</a></p></div>` });

  if (!p.temMesFim) {
    avisar(`⚠️ NPS Protheus · resumo de ${nomeMes_(periodo.mesFim)} NÃO enviado`,
      `<p>O envio automático de hoje foi suspenso: a Base Histórica ainda não tem respostas de <b>${nomeMes_(periodo.mesFim)}</b>.</p>
       <p>Importe o relatório do Fluig e envie pelo menu <b>✉️ Enviar resumo do NPS por e-mail</b>.</p>`);
    return;
  }
  if (!p.envios.length) {
    avisar(`⚠️ NPS Protheus · nenhum destinatário ativo`, `<p>O envio automático não encontrou times ou gestores ativos no cadastro.</p>`);
    return;
  }

  const r = enviarResumo('real', 'Automático');
  if (r.success) {
    props.setProperty('ENVIO_AUTO_MES', periodo.mesFim);
    avisar(`✅ NPS Protheus · resumo de ${nomeMes_(periodo.mesFim)} enviado`,
      `<p>${escHtml_(r.message)}</p><ul>${p.envios.map(e => `<li><b>${escHtml_(e.nome)}</b> (${e.tipo}) · ${escHtml_(e.para)} · NPS ${fmtNps_(e.npsMes)}</li>`).join('')}</ul>`);
  } else {
    avisar(`❌ NPS Protheus · falha no envio automático`, `<p>${escHtml_(r.message)}</p>`);
  }
}

function instalarEnvioAutomatico(silencioso) {
  removerGatilhos_(HANDLER_ENVIO_AUTO);
  ScriptApp.newTrigger(HANDLER_ENVIO_AUTO).timeBased().everyDays(1).atHour(HORA_ENVIO_AUTO).inTimezone(Session.getScriptTimeZone()).create();
  if (silencioso !== true) SpreadsheetApp.getUi().alert(`✅ Envio automático ativado.\nO resumo vai para TLs e gestores às ${HORA_ENVIO_AUTO}h do 2º dia útil de cada mês.\nPróximo: ${fmtDataBR_(proximoEnvioAuto_())}.`);
}

function removerEnvioAutomatico() {
  const n = removerGatilhos_(HANDLER_ENVIO_AUTO);
  SpreadsheetApp.getUi().alert(n ? '🔕 Envio automático desativado.' : 'O envio automático não estava ativo.');
}

function ativarAutomacoes() {
  instalarLembreteMensal(true);
  instalarEnvioAutomatico(true);
  statusAutomacoes();
}

function statusAutomacoes() {
  const c = lerCadastro_();
  const tls = c.times.filter(t => t.ativo).length, gs = c.gestores.filter(g => g.ativo).length;
  const p = prepararEnvios_();
  const linhas = [
    `${gatilhoAtivo_(HANDLER_LEMBRETE) ? '✅' : '⛔'} Lembrete de importação (1º dia útil, 7h) — próximo: ${proximoLembrete_()}`,
    `${gatilhoAtivo_(HANDLER_ENVIO_AUTO) ? '✅' : '⛔'} Envio automático do resumo (2º dia útil, ${HORA_ENVIO_AUTO}h) — próximo: ${fmtDataBR_(proximoEnvioAuto_())}`,
    '',
    `Destinatários ativos: ${tls} TL(s) e ${gs} gestor(es).`,
    `Dados de ${nomeMes_(p.periodo.mesFim)} na base: ${p.temMesFim ? 'sim ✅' : 'não ⚠️ — importe antes do envio'}.`
  ];
  const ult = PropertiesService.getDocumentProperties().getProperty('ENVIO_AUTO_MES');
  if (ult) linhas.push(`Último envio automático: resumo de ${nomeMes_(ult)}.`);
  SpreadsheetApp.getUi().alert('⏰ Automações do NPS', linhas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// API DA ABA AUTOMAÇÕES (Painel)
// ============================================================

function getStatusAutomacoes() {
  const c = lerCadastro_();
  const p = prepararEnvios_();
  const tz = Session.getScriptTimeZone();
  const ult = PropertiesService.getDocumentProperties().getProperty('ENVIO_AUTO_MES');
  const log = [];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LOG_ENVIOS);
  if (sh && sh.getLastRow() > 1) {
    const n = Math.min(30, sh.getLastRow() - 1);
    sh.getRange(sh.getLastRow() - n + 1, 1, n, CAB_LOG.length).getValues().reverse().forEach(r => log.push({
      quando: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0]),
      tipo: String(r[1]), nome: String(r[2]), email: String(r[3]), periodo: String(r[4]),
      nps: r[5] === '' ? '—' : String(r[5]).replace('.', ','), modo: String(r[7] || '')
    }));
  }
  return JSON.stringify({
    lembrete: { ativo: gatilhoAtivo_(HANDLER_LEMBRETE), proximo: proximoLembrete_() },
    envio: { ativo: gatilhoAtivo_(HANDLER_ENVIO_AUTO), proximo: fmtDataBR_(proximoEnvioAuto_()) },
    destinatarios: { tls: c.times.filter(t => t.ativo).length, gestores: c.gestores.filter(g => g.ativo).length },
    temMesFim: p.temMesFim, mesFimNome: nomeMes_(p.periodo.mesFim),
    ultimoAuto: ult ? nomeMes_(ult) : '', log: log
  });
}

function alternarAutomacao(qual, ligar) {
  if (qual === 'lembrete') { if (ligar) instalarLembreteMensal(true); else removerGatilhos_(HANDLER_LEMBRETE); }
  else if (qual === 'envio') { if (ligar) instalarEnvioAutomatico(true); else removerGatilhos_(HANDLER_ENVIO_AUTO); }
  else throw new Error('Automação desconhecida: ' + qual);
  return getStatusAutomacoes();
}

function enviarLembreteTestePainel() {
  return enviarLembrete_(true);
}
