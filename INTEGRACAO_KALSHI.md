# Integração Kalshi → Modelo do Pico (Copa 2026)

Puxa os preços **ao vivo** do mercado *Men's World Cup Winner* da Kalshi e alimenta o modelo do pico de probabilidade de campeão, automaticamente.

## O que foi montado

| Arquivo | Função |
|---|---|
| `kalshi_update.py` | Motor. Checa a exchange, baixa os 48 mercados, calcula prob justa, atualiza a planilha, gera snapshot + dashboard. |
| `kalshi_dashboard.html` | Painel ao vivo (dados embutidos, regenerado a cada pull). Abra no navegador. |
| `kalshi_snapshot.json` | Último snapshot bruto (auditoria / histórico). |
| `atualizar_kalshi.bat` | Clique duplo → atualização manual completa (planilha incluída). |
| `agendar_windows.bat` | Cria a tarefa automática no Agendador do Windows. |
| `atualizar_kalshi_silent.bat` | Versão usada pela tarefa agendada (roda em segundo plano). |

## Como funciona

A Kalshi expõe uma API pública e **sem autenticação** para dados de mercado. A integração usa dois endpoints:

1. **`GET /exchange/status`** — confirma se a exchange está negociando antes de puxar preços (o link que você mandou).
2. **`GET /markets?series_ticker=KXMENWORLDCUP&status=open`** — devolve um mercado por seleção (`KXMENWORLDCUP-26-BR`, `-ES`, …). Campos usados por seleção:
   - `yes_bid_dollars` / `yes_ask_dollars` — melhor compra/venda do "Sim" (0–1 = probabilidade).
   - `last_price_dollars` — último negócio.
   - `volume_fp`, `open_interest_fp` — liquidez.

Base: `https://api.elections.kalshi.com/trade-api/v2`

### Da Kalshi para o seu modelo

Para cada seleção: **preço justo bruto** = meio do bid/ask (ou `last`/`ask` nos azarões sem bid). A soma dos preços brutos é o *overround* (a vig da Kalshi, hoje ~8,5%). O script:

- Converte o preço bruto em **odds americanas** e escreve na **coluna B do Painel** — então a planilha faz o devig por normalização exatamente como no fluxo manual, e todas as métricas de pico (`E[max] = p·(1−ln p)`, teto `p^(1/N)`, múltiplo, escada de alvos) recalculam sozinhas.
- Cria/atualiza a aba **`Kalshi Live`** com bid/ask/last, prob justa, E[max], teto, múltiplo, volume e open interest por seleção.
- Carimba a data/hora da atualização na linha 2 do Painel.

> A premissa **N** (vitórias-equivalentes até o título) continua manual na célula **E3** — ajuste por fase: grupos ≈ 6, oitavas 5, quartas 4, semi 3, final 2.

## Como rodar (na sua máquina)

Pré-requisito: **Python 3** instalado. A API só responde da sua máquina (rede aberta).

**Manual:** feche a planilha no Excel e dê **clique duplo em `atualizar_kalshi.bat`**. Ele instala as dependências (`requests`, `openpyxl`) se faltarem, puxa os dados e atualiza tudo.

Pela linha de comando, dentro da pasta do projeto:

```
python kalshi_update.py            # ao vivo, atualiza planilha + dashboard + snapshot
python kalshi_update.py --no-xlsx  # só dashboard + snapshot (não mexe na planilha)
python kalshi_update.py --dry-run  # mostra o resultado sem gravar
```

## Atualização automática (agendada)

**Clique duplo em `agendar_windows.bat`** → cria a tarefa *"Kalshi Copa 2026"* que roda **todo dia às 08:00**. Por segurança, a versão agendada usa `--no-xlsx` (atualiza só o dashboard/snapshot, sem disputar o arquivo se a planilha estiver aberta). Para também atualizar o `.xlsx` automaticamente, remova `--no-xlsx` de `atualizar_kalshi_silent.bat`.

Ajustar o horário / frequência (Prompt de Comando):

```
:: de hora em hora (útil nos dias de jogo)
schtasks /Create /TN "Kalshi Copa 2026" /TR "C:\Users\pavan\Claude\Projects\Copa do Mundo 2026\atualizar_kalshi_silent.bat" /SC HOURLY /F

:: remover a tarefa
schtasks /Delete /TN "Kalshi Copa 2026" /F
```

## Observações

- **Vig**: o overround com mid-price (~8,5%) é maior que o ~2,8% do defirate porque inclui o spread cheio do book e o `ask` dos azarões. A normalização remove tudo — a prob justa final bate com a referência.
- **Live fetch no navegador**: a API bloqueia `fetch` cross-origin (CORS/CSP), por isso o dashboard usa dados **embutidos**, regenerados pelo script — não tenta buscar do navegador.
- **Fonte**: dados de mercado, não endossados pela FIFA (texto de regras da própria Kalshi).

## Coluna "Call da Copa" (edge) no dashboard

O dashboard tem uma coluna amarela editável **Call da Copa %**: digite o preço dos amigos (em %) para cada seleção e o painel calcula, ao lado:

- **Edge (pp)** = preço do Call da Copa − prob justa da Kalshi.
- **Sinal**: ▲ **COMPRAR** quando os amigos estão abaixo do justo (barato); ▼ **VENDER/FADE** quando estão acima (caro — típico de anfitriões/queridinhos como México, EUA, Canadá).
- O **limiar** (pp) na barra de cima filtra o ruído (padrão 1,0).
- A barra de resumo destaca as maiores oportunidades de compra e venda.

Os preços digitados ficam salvos **no próprio navegador** (localStorage) e sobrevivem às atualizações do script — você só precisa digitar de novo quando o preço no grupo mudar. O botão "Limpar preços" zera tudo.

## Edge na planilha (aba Kalshi Live)

A aba **Kalshi Live** agora traz, ao lado da Prob justa, três colunas:

- **Call da Copa %** (coluna azul, editável) — digite o preço dos amigos para cada seleção, direto no Excel.
- **Edge (pp)** — fórmula `= Call − Prob justa·100`, recalcula sozinha.
- **Sinal** — fórmula que mostra **COMPRAR** (verde) quando os amigos estão baratos e **VENDER** (vermelho) quando estão caros, com realce condicional automático.

O **limiar** (pp) fica na célula **S1** (padrão 1,0) — edite ali para afrouxar/apertar os sinais. Os preços que você digitar são **preservados** quando o script roda de novo (ele lê os valores antigos por ticker antes de reconstruir a aba), então você não perde o que digitou nas atualizações.
