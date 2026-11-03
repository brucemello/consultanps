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

## Telas

| Tela | Onde abrir | O que faz |
|---|---|---|
| **Importador** (`src/Importador.html`) | Menu da planilha → *Importar relatório do Fluig*, ou URL do Web App com `?p=importar` | Mostra os filtros do Fluig, lê o .xlsx/.csv no navegador, exibe prévia (linhas, período, NPS, alertas) e grava na `Base_Historica` sem duplicar |
| **Dashboard Executivo** (`src/Dashboard.html`) | Menu → *Abrir Dashboard Executivo* (abre em tela grande dentro da planilha), ou URL do Web App | KPIs, NPS por mês, P/N/D, distribuição de notas, causas raiz, achados, termos mais citados, comentários e plano de ação, com filtros de módulo, período, tipo e país |

| **Cadastro de times e gestores** (`src/Cadastro.html`) | Menu → *Cadastro de times e gestores* | Times (módulos, TL e e-mail do TL) e gestores (times que acompanham). Grava nas abas `Cadastro_Times` e `Cadastro_Gestores` |
| **Envio do resumo** (`src/EnvioResumo.html`) | Menu → *Enviar resumo do NPS por e-mail*, ou botão após importar | Prévia de cada e-mail, envio de teste para você e envio real. Cada envio fica registrado em `Log_Envios` |

### E-mails

- **TL**: NPS do último mês fechado com variação, acumulado do ano, composição P/N/D, evolução mês a mês, causas raiz, até 3 comentários de detratores e 1 elogio.
- **Gestor**: NPS consolidado, ranking dos times (mês, variação, acumulado, respostas, detratores), pontos de atenção (queda ≥ 10 pts ou NPS negativo), causas raiz e voz do cliente.

## Testes

`for t in tests/*.test.js; do node $t; done` — roda também no deploy, antes do `clasp push`.

## Secrets necessários

- `CLASPRC_JSON` — conteúdo do `~/.clasprc.json` (login do clasp)
- `SCRIPT_ID` — ID do projeto Apps Script
- `DEPLOYMENT_ID` — ID da implantação do Web App
