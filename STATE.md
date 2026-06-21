# STATE — Book Copa 2026
> **Fonte única de verdade.** Atualizar a cada trade ou resultado relevante.
> Última atualização manual: 21/06/2026 05:20 | Fase: Grupos (Rodada 2)

---

## 🚨 Alertas Ativos

**🔴 SENEGAL** — Jogo vs Noruega em **23/jun** é praticamente eliminatório. 10ctrl@0.75% pode ir a zero. **Considere fechar antes.**

**🔴 MARROCOS** — Jogo vs Escócia em **~22/jun** é mata-mata de fato. Escócia lidera com 3pts. Se perder → **fechar imediatamente no Call.**

**🟡 PORTUGAL×COLOMBIA** — Confronto direto em **27/jun**. Long nos dois = hedge interno. Um deles vai perder. Avaliar reduzir Portugal (10ctrl) antes.

---

## 📊 Book Atual

| Seleção | L/S | Contratos | P.Médio | Exposição bruta |
|---|---|---|---|---|
| Brasil | SHORT 🔴 | 40 | 12.625 | 505.0pp |
| Alemanha | LONG 🟢 | 30 | 9.000 | 270.0pp |
| Inglaterra | LONG 🟢 | 20 | 11.750 | 235.0pp |
| Portugal | LONG 🟢 | 10 | 12.000 | 120.0pp |
| Marrocos | LONG 🟢 | 30 | 2.420 | 72.6pp |
| Colombia | LONG 🟢 | 30 | 1.830 | 54.9pp |
| Noruega | LONG 🟢 | 10 | 3.000 | 30.0pp |
| Senegal | LONG 🟢 | 10 | 0.750 | 7.5pp |

**Exposição total:** 1295 pp×contratos

---

## 🏦 Último Trade

| Campo | Valor |
|---|---|
| ID | T009 |
| Data | 2026-06-21 |
| Seleção | Alemanha |
| Lado | VENDA |
| Contratos | 10 |
| Preço | 6.0 |
| Contraparte | Ricardo Espindola |
| Nota | Era 40C@8.25 → ficou 30C@9.00 |

> Histórico completo em `trades.json`

---

## 🗺️ Situação por Time

| Time | Grupo | Pts | Próx. Jogo | Ação | Alerta |
|---|---|---|---|---|---|
| Brasil | C | 1 | Haiti ~19/jun | MANTER SHORT | — |
| Alemanha | E | 3 | Ivory Coast ~20/jun | AGUARDAR | — |
| Inglaterra | L | 3 | Ghana/Panamá | MANTER LONG | — |
| Portugal | K | 1 | Uzbequistão 23/jun | ⚠️ VER ALERTA | 🟡 |
| Marrocos | C | 1 | Escócia ~22/jun | ⚠️ MATA-MATA | 🔴 |
| Colombia | K | 3 | DR Congo ~23/jun | MELHOR R/R | — |
| Noruega | I | 3 | Senegal 23/jun | MANTER LONG | — |
| Senegal | I | 0 | Noruega 23/jun | ⚠️ AVALIAR SAÍDA | 🔴 |

> Análise detalhada (chaveamento, R16, QF, ação completa) em `caminhos.json`

---

## 📅 Jogos Críticos Próximos

| Data | Jogo | Impacto no Portfolio |
|---|---|---|
| ~22/jun | Marrocos vs Escócia | 🔴 ALTO — quase eliminatório para Marrocos |
| 23/jun | Noruega vs Senegal | 🔴 ALTO — pode zerar posição Senegal |
| 23/jun | Colombia vs DR Congo | 🟡 MÉDIO — confirma liderança 1K |
| 23/jun | Portugal vs Uzbequistão | 🟢 BAIXO — deve vencer |
| 26/jun | Noruega vs França | 🟡 MÉDIO — jogo duro, saída ideal Noruega |
| 27/jun | Portugal vs Colombia | 🔴 ALTO — confronto direto, hedge interno |

---

## 🧠 Hedges Estruturais do Portfolio

| Estrutura | Mecanismo |
|---|---|
| Short Brasil + Long Inglaterra | Se ENG×BRA no QF: short BRA ganha + long ENG sobe |
| Long Colombia + Long Portugal | ⚠️ CONFLITO: se enfrentam em 27/jun — um cancela o outro |
| Long Noruega + Long Alemanha | R16 potencial NOr×ALE — as duas posições não podem vencer juntas nesse estágio |

---

## 📁 Mapa de Arquivos

| Arquivo | Função |
|---|---|
| `STATE.md` | **Este arquivo** — ponto único de verdade |
| `trades.json` | Ledger completo de operações (append-only) |
| `caminhos.json` | Análise estratégica por time (editável sem tocar código) |
| `opcoes_model.py` | Motor de cálculo — Greeks, cenários, Excel |
| `CONTEXTO_Call_da_Copa.md` | Tese, modelo matemático, regras (referência estática) |
| `historico_precos.csv` | Série histórica de preços Kalshi + Call da Copa |
| `kalshi_wc_markets.json` | Snapshot mais recente da Kalshi (atualizado pelo Actions) |
| `GUIA_DE_USO.md` | Operacional — rotina diária, como rodar os scripts |

---

## 🔄 Como Atualizar Este Arquivo

- **Novo trade:** atualizar tabela Book + seção Último Trade + `trades.json`
- **Resultado de jogo:** atualizar Pts/situação na tabela + alertas se necessário + `caminhos.json`
- **Preços Kalshi:** o GitHub Actions atualiza `kalshi_wc_markets.json` a cada hora automaticamente
- **No chat com Claude:** cole este arquivo no início da sessão para retomar com contexto completo
