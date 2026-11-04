// Testes do cadastro de times/gestores e do resumo de NPS (src/ResumoTimes.js).
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const ctx = {};
vm.createContext(ctx);
const SRC = path.join(__dirname, '..', 'src');
fs.readdirSync(SRC).filter(f => f.endsWith('.js')).sort().forEach(f => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
const run = code => { const v = vm.runInContext(code, ctx); return v && typeof v === 'object' ? JSON.parse(JSON.stringify(v)) : v; };

let ok = 0;
function t(nome, fn) { fn(); ok++; console.log('  ✓ ' + nome); }

// base sintética: [data, status, país, linha, módulo, cliente, email, nota, codt, tipo, razão, comentário, mês, classif, aderência]
vm.runInContext(`
  var L = (d, mod, nota, com, extra) => {
    const resp = nota !== '';
    const cl = !resp ? '' : nota >= 9 ? 'Promotor' : nota >= 7 ? 'Neutro' : 'Detrator';
    return [new Date(d), resp ? 'Respondido' : 'Cancelado', (extra && extra.pais) || 'BR', 'Protheus', mod, '', '', nota, '', (extra && extra.tipo) || 'Externo', 'CLIENTE ' + d, com || '', '', cl, ''];
  };
  var BASE = [
    L('2026-09-03', 'SIGAFIS', 10), L('2026-09-04', 'SIGAFIS', 3, 'Sistema lento'), L('2026-09-05', 'SIGAFIS', 8),
    L('2026-10-02', 'SIGAFIS', 10, 'Top demais'), L('2026-10-03', 'SIGAFIS', 9), L('2026-10-04', 'SIGAFIS', 2, 'Falta suporte'), L('2026-10-05', 'SIGAFIS', ''),
    L('2026-10-06', 'SIGACOM', 10), L('2026-10-07', 'SIGACOM', 10),
    L('2026-10-08', 'SIGAFIS', 0, 'Horrível', { pais: 'MX' }), L('2026-10-09', 'SIGAFIS', 0, 'x', { tipo: 'Interno' })
  ];
  var REGS = registrosDeBase_(BASE, FILTROS_PADRAO);
  var PER = { inicio: '01/01/2026', fim: '31/10/2026', mesInicio: '2026-01', mesFim: '2026-10' };
`, ctx);

t('registros respeitam país e tipo de cliente padrão', () => {
  assert.strictEqual(run('REGS.length'), 9);
});

t('resumo calcula NPS do período, do último mês e a variação', () => {
  const r = run(`resumoModulos_(REGS, ['SIGAFIS'], '2026-01', '2026-10')`);
  assert.strictEqual(r.total.base, 6);                       // 6 respostas com nota no SIGAFIS
  assert.strictEqual(Math.round(r.total.nps * 10) / 10, 16.7); // (3P - 2D)/6
  assert.strictEqual(Math.round(r.ultimo.nps * 10) / 10, 33.3); // out: 2P 1D de 3
  assert.strictEqual(r.anterior.nps, 0);                    // set: 1P 1N 1D
  assert.strictEqual(Math.round(r.variacao * 10) / 10, 33.3);
  assert.strictEqual(r.ultimo.abord, 4);
  assert.strictEqual(r.detratores.length, 2);
  assert.strictEqual(r.detratores[0].com, 'Falta suporte');  // detrator do último mês vem primeiro
  assert.strictEqual(r.causas[0].det, 1);
  assert.strictEqual(r.elogio.com, 'Top demais');
});

t('validação do cadastro pega e-mail inválido, time repetido e gestor com time inexistente', () => {
  const erros = run(`validarCadastro_({
    times: [{ time: 'Fiscal', modulos: ['SIGAFIS'], email: 'tl@totvs.com.br' }, { time: 'fiscal', modulos: ['SIGAFIS'], email: 'tl@totvs' }],
    gestores: [{ nome: 'Ana', email: 'ana@totvs.com.br', times: ['Compras'] }]
  })`);
  assert.ok(erros.some(e => /repetido/.test(e)));
  assert.ok(erros.some(e => /inválido/.test(e)));
  assert.ok(erros.some(e => /não cadastrado/.test(e)));
  assert.strictEqual(run(`validarCadastro_({ times: [{ time: 'Fiscal', modulos: ['SIGAFIS'], email: 'a@totvs.com.br, b@totvs.com.br' }], gestores: [] }).length`), 0);
});

t('envios: um por TL ativo e um por gestor, com times corretos', () => {
  const envios = run(`montarEnvios_({
    times: [
      { time: 'Fiscal', modulos: ['SIGAFIS'], tl: 'João Silva', email: 'joao@totvs.com.br', ativo: true },
      { time: 'Compras', modulos: ['SIGACOM'], tl: 'Maria', email: 'maria@totvs.com.br', ativo: true },
      { time: 'Inativo', modulos: ['SIGAEST'], tl: 'X', email: 'x@totvs.com.br', ativo: false }
    ],
    gestores: [
      { nome: 'Ana Gestora', email: 'ana@totvs.com.br', times: [], ativo: true },
      { nome: 'Beto', email: 'beto@totvs.com.br', times: ['Compras'], ativo: true },
      { nome: 'Off', email: 'off@totvs.com.br', times: [], ativo: false }
    ]
  }, REGS, PER, { planilha: 'https://docs.google.com/x' })`);
  assert.deepStrictEqual(envios.map(e => e.tipo + ':' + e.nome), ['TL:João Silva', 'TL:Maria', 'Gestor:Ana Gestora', 'Gestor:Beto']);
  assert.strictEqual(envios[2].time, 'Fiscal, Compras');
  assert.strictEqual(envios[3].time, 'Compras');
  assert.ok(/^NPS Fiscal · jan a out\/2026: 16,7 · out\/26 33,3 \(▲ 33,3\)$/.test(envios[0].assunto), envios[0].assunto);
  assert.ok(envios[0].html.includes('Olá, João!'));
  assert.ok(envios[0].html.includes('Falta suporte'));
  assert.ok(envios[2].html.includes('Ranking dos times'));
  assert.ok(!envios[0].html.includes('undefined') && !envios[2].html.includes('NaN'));
});

t('textos do cadastro são escapados no HTML do e-mail', () => {
  const html = run(`htmlEmailTime_({ time: '<b>X</b>', modulos: ['SIGAFIS'], tl: 'Zé' }, resumoModulos_(REGS, ['SIGAFIS'], '2026-01', '2026-10'), PER, {})`);
  assert.ok(!html.includes('<b>X</b>') && html.includes('&lt;b&gt;X&lt;/b&gt;'));
});

t('2º dia útil do mês', () => {
  const f = s => run(`fmtDia_(segundoDiaUtil_(${s}))`);
  assert.strictEqual(f('2026, 10'), '2026-11-04'); // 02/11 Finados → 1º útil 03, 2º útil 04
  assert.strictEqual(f('2026, 11'), '2026-12-02');
  assert.strictEqual(f('2027, 0'), '2027-01-05');
  assert.strictEqual(f('2026, 4'), '2026-05-05');  // 01/05 feriado (sex) → 04 seg, 05 ter
});

t('período escolhido na tela (jul a out)', () => {
  const p = run(`periodoDeMeses_({ mesInicio: '2026-07', mesFim: '2026-10' })`);
  assert.deepStrictEqual([p.inicio, p.fim, p.mesInicio, p.mesFim], ['01/07/2026', '31/10/2026', '2026-07', '2026-10']);
  assert.strictEqual(run(`rotuloPeriodo_(periodoDeMeses_({ mesInicio: '2026-07', mesFim: '2026-10' }))`), 'jul a out/2026');
  assert.strictEqual(run(`rotuloPeriodo_(periodoDeMeses_({ mesInicio: '2025-11', mesFim: '2026-02' }))`), 'nov/2025 a fev/2026');
  assert.strictEqual(run(`rotuloPeriodo_(periodoDeMeses_({ mesInicio: '2026-10', mesFim: '2026-10' }))`), 'outubro/2026');
  assert.strictEqual(run(`periodoDeMeses_({ mesInicio: '2026-10', mesFim: '2026-07' }).mesInicio`), '2026-07'); // invertido
  const r = run(`resumoModulos_(REGS, ['SIGAFIS'], '2026-10', '2026-10')`);
  assert.strictEqual(r.total.base, 3); // só outubro
});

t('NPS por semestre (1º jan–jun, 2º jul–dez, parcial até o mês fechado)', () => {
  vm.runInContext(`var REGS2 = registrosDeBase_([
    L('2026-02-10', 'SIGAFIS', 10), L('2026-03-10', 'SIGAFIS', 2), L('2026-06-10', 'SIGAFIS', 10), L('2026-06-11', 'SIGAFIS', 9),
    L('2026-07-10', 'SIGAFIS', 3), L('2026-10-10', 'SIGAFIS', 10), L('2026-11-10', 'SIGAFIS', 0)
  ], FILTROS_PADRAO);`, ctx);
  const sm = run(`resumoModulos_(REGS2, ['SIGAFIS'], '2026-07', '2026-10').semestres`);
  assert.strictEqual(sm.ano, '2026');
  assert.strictEqual(sm.h1.base, 4); assert.strictEqual(sm.h1.nps, 50);        // 3P 1D de 4 — junho entra no 1º
  assert.strictEqual(sm.h2.base, 2); assert.strictEqual(sm.h2.nps, 0);         // jul+out; novembro (depois do mês fechado) fica fora
  assert.strictEqual(sm.h2.parcial, true); assert.strictEqual(sm.h2.fim, '2026-10');
  assert.strictEqual(sm.variacao, -50);
  const mar = run(`resumoModulos_(REGS2, ['SIGAFIS'], '2026-01', '2026-03').semestres`);
  assert.strictEqual(mar.h1.parcial, true); assert.strictEqual(mar.h2.futuro, true); assert.strictEqual(mar.h2.nps, null);
  const html = run(`blocoSemestres_(resumoModulos_(REGS2, ['SIGAFIS'], '2026-07', '2026-10').semestres)`);
  assert.ok(html.includes('1º semestre') && html.includes('2º semestre') && html.includes('jul a out/26 · parcial'));
});

console.log(`cadastro.test.js: ${ok} testes ok`);
