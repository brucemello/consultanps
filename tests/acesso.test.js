// Controle de acesso da Central NPS e utilitários de data.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const SRC = path.join(__dirname, '..', 'src');
const arquivos = fs.readdirSync(SRC).filter(f => f.endsWith('.js')).sort();
const ctx = { Session: null };
vm.createContext(ctx);
arquivos.forEach(f => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
const run = code => vm.runInContext(code, ctx);

let ok = 0;
function t(nome, fn) { fn(); ok++; console.log('  ✓ ' + nome); }
const sessao = (ativo, dono) => { ctx.Session = { getActiveUser: () => ({ getEmail: () => ativo }), getEffectiveUser: () => ({ getEmail: () => dono }) }; };

t('fmtDataBR_ formata dd/mm/aaaa (sem recursão)', () => {
  assert.strictEqual(run('fmtDataBR_(new Date(2026, 10, 4))'), '04/11/2026');
});

t('dono, gatilho (sem usuário ativo) e outra pessoa', () => {
  sessao('bruce.mello@totvs.com.br', 'bruce.mello@totvs.com.br'); assert.strictEqual(run('ehDono_()'), true);
  sessao('Bruce.Mello@totvs.com.br', 'bruce.mello@totvs.com.br'); assert.strictEqual(run('ehDono_()'), true);
  sessao('', 'bruce.mello@totvs.com.br'); assert.strictEqual(run('ehDono_()'), true);
  sessao('tl@totvs.com.br', 'bruce.mello@totvs.com.br'); assert.strictEqual(run('ehDono_()'), false);
  assert.throws(() => run('salvarCadastro("{}")'), /Acesso restrito/);
  assert.throws(() => run("enviarResumo('real')"), /Acesso restrito/);
  assert.throws(() => run("alternarAutomacao('envio', true)"), /Acesso restrito/);
  assert.throws(() => run('importarLinhas([[1]], "x")'), /Acesso restrito/);
});

t('toda função pública (chamável pelo navegador) está protegida, exceto as de leitura', () => {
  const LIBERADAS = new Set(['onOpen', 'doGet', 'include', 'getDadosDashboard']);
  const faltando = [];
  arquivos.forEach(f => {
    const src = fs.readFileSync(path.join(SRC, f), 'utf8');
    for (const m of src.matchAll(/^function ([A-Za-z0-9_]+)\([^)]*\) \{\n(.*)/gm)) {
      if (m[1].endsWith('_') || LIBERADAS.has(m[1])) continue;
      if (!/exigirDono_\(\)/.test(m[2])) faltando.push(`${f}:${m[1]}`);
    }
  });
  assert.deepStrictEqual(faltando, []);
});

t('Web App aberto ao domínio, executando como o dono', () => {
  const man = JSON.parse(fs.readFileSync(path.join(SRC, 'appsscript.json'), 'utf8'));
  assert.deepStrictEqual(man.webapp, { executeAs: 'USER_DEPLOYING', access: 'DOMAIN' });
});

console.log(`acesso.test.js: ${ok} testes ok`);
