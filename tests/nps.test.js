// Testes das funções puras de src/*.js (rodam no Node, sem Apps Script).
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const ctx = {};
vm.createContext(ctx);
const SRC = path.join(__dirname, '..', 'src');
fs.readdirSync(SRC).filter(f => f.endsWith('.js')).sort().forEach(f => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
// Objetos vindos do contexto vm são de outro "realm": normaliza via JSON (datas ficam como estão).
const run = code => {
  const v = vm.runInContext(code, ctx);
  const isDate = Object.prototype.toString.call(v) === '[object Date]';
  return (v && typeof v === 'object' && !isDate) ? JSON.parse(JSON.stringify(v)) : v;
};

let ok = 0;
function t(nome, fn) { fn(); ok++; console.log('  ✓ ' + nome); }

t('período padrão: jan até o fim do mês anterior', () => {
  const p = run('periodoPadrao_(new Date(2026, 10, 2))');
  assert.deepStrictEqual([p.inicio, p.fim, p.mesInicio, p.mesFim], ['01/01/2026', '31/10/2026', '2026-01', '2026-10']);
});

t('período padrão em janeiro usa o ano anterior inteiro', () => {
  const p = run('periodoPadrao_(new Date(2027, 0, 4))');
  assert.deepStrictEqual([p.inicio, p.fim], ['01/01/2026', '31/12/2026']);
});

t('parseData_ entende dd/mm/aaaa hh:mm:ss, ISO e serial do Excel', () => {
  const d = run('parseData_("05/03/2026 14:07:09")');
  assert.deepStrictEqual([d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getSeconds()], [2026, 2, 5, 14, 9]);
  const i = run('parseData_("2026-03-05")');
  assert.deepStrictEqual([i.getMonth(), i.getDate()], [2, 5]);
  const s = run('parseData_(46086)'); // 05/03/2026
  assert.deepStrictEqual([s.getFullYear(), s.getMonth(), s.getDate()], [2026, 2, 5]);
  assert.strictEqual(run('parseData_("")'), null);
});

t('mapeia colunas por nome, sem acento e em qualquer ordem', () => {
  const idx = run('mapearColunas_(["Comentário", "Razão Social", "Nome do Produto", "NOTA", "Status", "Data da Resposta", "País"])');
  assert.strictEqual(idx.comentario, 0);
  assert.strictEqual(idx.razao, 1);
  assert.strictEqual(idx.modulo, 2);
  assert.strictEqual(idx.nota, 3);
  assert.strictEqual(idx.data, 5);
  assert.strictEqual(idx.pais, 6);
  assert.strictEqual(idx.tipo, -1);
});

t('monta linhas da base, classifica e ignora duplicados', () => {
  const r = run(`montarLinhasBase_([
    ['Data da Resposta','Status','País','Nome do Produto','Nota','Tipo do Cliente','Razão Social','Comentário'],
    ['10/02/2026 09:00:00','Respondido','BR','SIGAFIS','10','Externo','ACME','Top demais'],
    ['10/02/2026 09:00:00','Respondido','BR','SIGAFIS','10','Externo','ACME','repetido'],
    ['11/02/2026 10:00:00','Respondido','BR','SIGAFIS','7,0','Externo','BETA',''],
    ['12/02/2026 10:00:00','Respondido','BR','SIGAFIS','3','Externo','GAMA','Sistema lento e travando'],
    ['13/02/2026 10:00:00','Cancelado','BR','SIGAFIS','','Externo','DELTA','']
  ], new Set())`);
  assert.strictEqual(r.linhas.length, 4);
  assert.strictEqual(r.duplicados, 1);
  assert.deepStrictEqual(r.linhas.map(l => l[13]), ['Promotor', 'Neutro', 'Detrator', '']);
  assert.deepStrictEqual(r.linhas.map(l => l[14]), ['Respondido', 'Respondido', 'Respondido', 'Cancelado']);
  assert.strictEqual(r.linhas[0][4], 'SIGAFIS');
  assert.strictEqual(r.linhas[0][9], 'Externo');
  assert.strictEqual(r.linhas[1][7], 7);
});

t('termos curtos não casam dentro de outras palavras', () => {
  assert.strictEqual(run('classificarComentario_("Os clientes antes gostavam").pilar'), '5. Outros / Não Classificados');
  assert.strictEqual(run('classificarComentario_("Erro na TES de saída").pilar'), '3. Regras de Negócio & Funcionalidades');
  assert.strictEqual(run('classificarComentario_("Sistema lento e travando").pilar'), '1. Desempenho & Estabilidade');
  assert.strictEqual(run('classificarComentario_("Falta suporte").pilar'), '2. Suporte & Atendimento');
});

t('comentários vazios ou genéricos são ignorados', () => {
  assert.strictEqual(run('comentarioValido_("Sem comentários")'), false);
  assert.strictEqual(run('comentarioValido_("NDA")'), false);
  assert.strictEqual(run('comentarioValido_("-")'), false);
  assert.strictEqual(run('comentarioValido_("Top demais")'), true);
});


t('1º dia útil pula fim de semana e feriados nacionais', () => {
  const f = s => run(`fmtDia_(primeiroDiaUtil_(${s}))`);
  assert.strictEqual(f('2026, 10'), '2026-11-03'); // 01/11 domingo, 02/11 Finados
  assert.strictEqual(f('2027, 0'), '2027-01-04');  // 01/01 feriado, 02-03 fim de semana
  assert.strictEqual(f('2026, 4'), '2026-05-04');  // 01/05 feriado (sexta)
  assert.strictEqual(f('2026, 11'), '2026-12-01');
  assert.strictEqual(run('fmtDia_(pascoa_(2026))'), '2026-04-05');
  assert.strictEqual(run('feriadosNacionais_(2026).has("2026-02-17")'), true); // carnaval
});

console.log(`nps.test.js: ${ok} testes ok`);
