# Guia de uso — Modelo do Pico + Kalshi ao vivo (Copa 2026)

Guia prático do dia a dia. Para os detalhes técnicos (endpoints, fórmulas, vig), veja `INTEGRACAO_KALSHI.md`. Tudo fica na pasta `C:\Users\pavan\Claude\Projects\Copa do Mundo 2026\`.

---

## Em 30 segundos

- **Ver preços e o edge agora** → abra `kalshi_dashboard.html` no navegador.
- **Operar / controlar trades** → abra `Copa_2026_Pico_Probabilidade_Campeao.xlsx` (abas **Painel**, **Kalshi Live** e **Boletagem**).
- **Puxar dados novos da Kalshi** → feche a planilha e clique duplo em `atualizar_kalshi.bat`.
- **Lançar os preços do grupo** → cole no `call_da_copa.txt` e clique duplo em `atualizar_call.bat`.

---

## Os arquivos da pasta

| Arquivo | Para quê |
|---|---|
| `Copa_2026_Pico_Probabilidade_Campeao.xlsx` | A planilha: Painel, Kalshi Live (edge), Boletagem (trades). |
| `kalshi_dashboard.html` | Painel ao vivo no navegador, com a coluna Call da Copa e edge. |
| `atualizar_kalshi.bat` | Puxa a Kalshi e atualiza tudo (planilha + dashboard + histórico). |
| `call_da_copa.txt` | Onde você cola as mensagens do WhatsApp com os preços. |
| `atualizar_call.bat` | Lê o `call_da_copa.txt` e atualiza a coluna Call da Copa. |
| `agendar_windows.bat` | Liga a atualização automática diária no Windows. |
| `historico_precos.csv` + pasta `historico/` | Histórico datado dos dois mercados (para gráficos). |

> Os demais (`kalshi_update.py`, `ler_call_da_copa.py`, `salvar_snapshot.py`, `kalshi_snapshot.json`, `call_prices.json`) são o motor — você não precisa abrir.

---

## Primeira vez (configuração — só uma vez)

1. **Confira o Python.** Prompt de Comando → `python --version`. Se aparecer 3.x, está pronto. Se não, instale pela Microsoft Store (marque "Add to PATH").
2. **Ligue a atualização automática (opcional).** Clique duplo em `agendar_windows.bat` → cria a tarefa diária das 08:00.

Não precisa instalar mais nada — o próprio `.bat` instala as bibliotecas (`requests`, `openpyxl`) na primeira execução.

---

## Uso diário

### No dashboard (consulta rápida)

Abra `kalshi_dashboard.html`. As 48 seleções aparecem ordenadas por **prob justa**, com pico esperado, teto na final e múltiplo. A coluna amarela **Call da Copa %** já vem preenchida com os preços do grupo (se você rodou o `atualizar_call.bat`) e é editável; ao lado aparecem **Edge** e o **Sinal** (▲ COMPRAR / ▼ VENDER). A barra de cima resume as maiores oportunidades. Edições feitas no navegador ficam salvas nele.

### Na planilha (operação)

- **Aba Painel** — prob justa e métricas do pico. A coluna B (odds) é preenchida pela Kalshi; você só ajusta a célula **E3** (premissa N) conforme a fase avança.
- **Aba Kalshi Live** — coluna azul **Call da Copa**, com **Edge** e **Sinal** calculados (verde COMPRAR, vermelho VENDER). Limiar na célula **S1** (padrão 1,0 pp).
- **Aba Boletagem** — onde você controla seus trades (veja a seção própria abaixo).

> **Ajuste o N (E3) conforme o torneio anda:** grupos ≈ 6 · oitavas 5 · quartas 4 · semi 3 · final 2. Quanto menor o N, mais alto o teto que a prob pode atingir.

---

## Atualizar os preços da Kalshi

1. **Manual completo (recomendado):** feche a planilha → clique duplo em `atualizar_kalshi.bat`. Atualiza planilha, dashboard, snapshot e arquiva o histórico.
2. **Só o dashboard (planilha aberta):** Prompt de Comando na pasta → `python kalshi_update.py --no-xlsx`.
3. **Automático:** a tarefa do Windows roda sozinha (08:00 por padrão).

> A atualização que mexe na planilha só funciona com o **Excel fechado**. Erro de permissão = planilha aberta; feche e rode de novo.

---

## Atualizar os preços do Call da Copa (WhatsApp)

1. No WhatsApp, **copie** as mensagens com os preços (pode copiar tudo — horário, nomes e papo são ignorados).
2. Abra `call_da_copa.txt`, **apague o conteúdo, cole** o que copiou e **salve**.
3. Feche a planilha e clique duplo em `atualizar_call.bat`.

O leitor entende o **formato do seu grupo** — bandeiras com bid/ask, ex.: `🇧🇷 12,25 / 13` (usa o meio = 12,625) e `🇫🇷 - / 18,75` (usa o lado disponível). Também aceita por nome (`Brasil 7`, `Espanha 16/17`, `Argentina: 8`). Inglaterra e Escócia são reconhecidas pela bandeira ou por "Eng"/"Esc". Ele mostra um relatório do que reconheceu e do que **não** reconheceu — se algo ficar de fora, me mande o trecho que eu ajusto o leitor.

> Preços em **pontos percentuais** (7 = 7%; quando vem bid/ask, usa o meio). Seleções não citadas mantêm o último preço.

---

## Lendo os sinais (resumo da estratégia)

- **Edge = preço do Call da Copa − prob justa da Kalshi** (pontos percentuais).
- **▲ COMPRAR** — amigos abaixo do justo (barato). Foco nas zebras de múltiplo alto (★ no dashboard).
- **▼ VENDER / FADE** — amigos acima do justo (caro). Clássico nos anfitriões e queridinhos (México, EUA, Canadá, o time da torcida) — e, pelo viés local, o **Brasil** costuma ser o maior short.
- **Timing de venda:** a prob bate o topo no *run-up antes de um jogo difícil*. Venda em camadas a cada vitória (aba Escada de Alvos). Nunca segure uma zebra **através** de um jogo contra um time mais forte.

Lembre: o edge real vem da **diferença entre o Call da Copa e a Kalshi**, não de segurar até o pico (num mercado eficiente, segurar tem lucro esperado zero).

---

## Boletagem (controle de trades e posição)

A aba **Boletagem** controla suas operações no Call da Copa (contratos que liquidam 0–100; P&L em pontos).

Registre cada execução no **LOG** (esquerda): Data, Seleção (lista), Lado (Compra/Venda), Contratos, Preço. A tabela de **POSIÇÃO** (direita) soma sozinha por seleção: **Net contr.**, **Custo líq.**, **Preço médio**, **Preço atual (Call)** (marcação pelo preço do grupo, puxado da Kalshi Live), **Valor mkt** e **P&L (pts)** com realce verde/vermelho. No topo: **P&L aberto**, **exposição** e **nº de posições abertas**.

> O LOG é **seu** — o script nunca apaga. A aba aparece na primeira vez que você rodar `atualizar_kalshi.bat` com o **Excel fechado**.

---

## Histórico datado dos preços

Toda atualização (pelos `.bat`) arquiva sozinha o estado dos dois mercados:

- `historico/kalshi_AAAA-MM-DD_HHMM.json` e `historico/call_AAAA-MM-DD_HHMM.json` — fotos completas de cada momento.
- `historico_precos.csv` — uma linha por seleção por data, com `kalshi_justa_pct`, `call_preco_pct` e `edge_pp`. É o arquivo para **plotar a evolução do prêmio** dos amigos (abra no Excel → Inserir Tabela Dinâmica/Gráfico, filtrando por seleção).

Para arquivar manualmente: `python salvar_snapshot.py`.

---

## Automático: ligar, mudar horário, desligar

```
:: ligar (todo dia 08:00) — ou só clique duplo em agendar_windows.bat
schtasks /Create /TN "Kalshi Copa 2026" /TR "C:\Users\pavan\Claude\Projects\Copa do Mundo 2026\atualizar_kalshi_silent.bat" /SC DAILY /ST 08:00 /F

:: nos dias de jogo: de hora em hora
schtasks /Create /TN "Kalshi Copa 2026" /TR "C:\Users\pavan\Claude\Projects\Copa do Mundo 2026\atualizar_kalshi_silent.bat" /SC HOURLY /F

:: desligar
schtasks /Delete /TN "Kalshi Copa 2026" /F
```

Há também uma tarefa equivalente dentro do app Claude (seção **Scheduled** na barra lateral, "kalshi-copa-2026-refresh"), que puxa os dados pelo navegador, arquiva o histórico e te manda um resumo diário.

---

## Resolvendo problemas

| Sintoma | Causa provável | Solução |
|---|---|---|
| Erro de permissão ao atualizar | Planilha aberta no Excel | Feche o Excel e rode de novo |
| "python não é reconhecido" | Python fora do PATH | Instale pela Microsoft Store com "Add to PATH" |
| Dashboard com dados velhos | Não rodou a atualização | Rode `atualizar_kalshi.bat` e dê F5 no navegador |
| Edge/Sinal em branco | Faltou o preço do Call da Copa | Rode `atualizar_call.bat` ou digite na coluna azul |
| Alguma seleção não entrou do WhatsApp | Formato fora do previsto | Veja o relatório do `atualizar_call.bat`; me mande o trecho |
| Aba Boletagem não apareceu | Rodou com o Excel aberto | Feche o Excel e rode `atualizar_kalshi.bat` |
| Nenhum dado retornado | Sem internet / exchange fechada | Cheque a conexão; o script avisa o status |

---

## Rotina sugerida num dia de jogo

1. De manhã, rode `atualizar_kalshi.bat` (ou deixe a tarefa horária ligada).
2. Cole os preços do grupo em `call_da_copa.txt` → rode `atualizar_call.bat`.
3. Abra a planilha → ajuste o **N (E3)** se a fase mudou.
4. Olhe a coluna **Sinal** (Kalshi Live ou dashboard): verde = comprar barato, vermelho = vender caro.
5. Lançou um trade? Registre na **Boletagem** para acompanhar posição e P&L.
6. Antes de um jogo difícil de uma zebra que você carrega, venda em camadas (Escada de Alvos).
