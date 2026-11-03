# Consulta NPS

Código do Google Apps Script com deploy automático via GitHub Actions.

## Como funciona

| Workflow | Quando roda | O que faz |
|---|---|---|
| **Deploy Apps Script** | Todo push na `main` que mexa em `src/` (ou manualmente) | `clasp push` + atualiza a implantação `DEPLOYMENT_ID` — a URL do Web App continua a mesma |
| **Importar do Apps Script** | Manual (aba Actions) | `clasp pull` do editor do Google para `src/` e faz commit |
| **Verificar Apps Script** | Manual | Compara o código publicado com o repositório |
| **Diagnóstico de implantações** | Manual | Lista as implantações e mostra qual é o Web App atualizado pelo deploy |

> Evite editar direto no editor do Apps Script: o próximo deploy sobrescreve. Se editar, rode **Importar do Apps Script** antes.

## Central NPS (uma tela, com abas)

Abra pelo menu **🚀 NPS Protheus → 🧭 Abrir Central NPS** ou direto no navegador: <https://script.google.com/a/macros/totvs.com.br/s/AKfycbyCUu4TFMso0d8E5VonjbOB929O0wv_MXuOhB_sXx3-f-M1-cz66at9ujLDhe7_eU-vAA/exec> (aba inicial com `?tab=dash|imp|env|cad|auto`).

| Aba | Arquivo | O que faz |
|---|---|---|
| 📊 **Dashboard** | `src/PainelDashboard.html` | KPIs, NPS por mês, P/N/D, notas, causas raiz, achados, termos, comentários e plano de ação, com filtros. "Enviar resumo" leva o período do filtro |
| 📥 **Importar** | `src/PainelImportar.html` | Filtros do Fluig, prévia do arquivo e gravação na `Base_Historica` sem duplicar; atualiza o dashboard ao terminar |
| ✉️ **Envio de resumo** | `src/PainelEnvio.html` | Período De/Até, prévia de cada e-mail, teste para você e envio real (registrado em `Log_Envios`) |
| 👥 **Times e gestores** | `src/PainelCadastro.html` | Times (módulos, TL, e-mail) e gestores (times acompanhados) → abas `Cadastro_Times` e `Cadastro_Gestores` |
| ⏰ **Automações** | `src/PainelAutomacoes.html` | Liga/desliga lembrete (1º dia útil) e envio automático (2º dia útil), status e histórico de envios |

### Acesso

O Web App fica aberto para o domínio TOTVS e roda como o dono (quem implanta):

- **Você (dono)** vê a Central completa.
- **Qualquer outra pessoa da TOTVS** (TLs e gestores que clicam no botão do e-mail) vê só o **Dashboard**, em modo leitura.
- No servidor, toda função que grava, envia e-mail ou mexe em configuração chama `exigirDono_()` — o teste `tests/acesso.test.js` garante que nenhuma fica de fora.

`src/Painel.html` é a moldura (barra superior e abas) e inclui os módulos. Cada módulo isola seu CSS (`#v-…`) e seus IDs (prefixo `d-`, `i-`, `e-`, `c-`, `a-`).

### E-mails

- **TL**: NPS do último mês fechado com variação, acumulado do ano, composição P/N/D, evolução mês a mês, causas raiz, até 3 comentários de detratores e 1 elogio.
- **Gestor**: NPS consolidado, ranking dos times (mês, variação, acumulado, respostas, detratores), pontos de atenção (queda ≥ 10 pts ou NPS negativo), causas raiz e voz do cliente.

### Automações (menu ⏰ Automações)

| Quando | O quê |
|---|---|
| 1º dia útil, 7h | Lembrete para você importar o relatório do Fluig |
| 2º dia útil, 8h | Envio automático do resumo para TLs e gestores (uma vez por mês). Se a base não tiver dados do mês fechado, não envia e avisa você |

Feriados nacionais e fins de semana são pulados. Ative uma vez em *⏰ Automações → Ativar todas*.

## Testes

`for t in tests/*.test.js; do node $t; done` — roda também no deploy, antes do `clasp push`.

## Secrets necessários

- `CLASPRC_JSON` — conteúdo do `~/.clasprc.json` (login do clasp)
- `SCRIPT_ID` — ID do projeto Apps Script
- `DEPLOYMENT_ID` — ID da implantação do Web App
