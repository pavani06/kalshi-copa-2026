# Contexto — Call da Copa 2026 (precificação do pico de probabilidade de campeão)

> Documento de transferência de contexto. Cole isto no início de um novo chat para retomar o trabalho sem perder nada.

## 1. Quem opera e onde

- **Operador:** monta posições de compra e venda da probabilidade de um país ser **campeão** da Copa do Mundo de 2026.
- **Dois mercados:**
  - **Call da Copa** — mercado privado dos amigos num grupo de WhatsApp. Settlement **igual Kalshi**: preço de 0 a 100, liquida em 100 se campeão e 0 se não; dá pra comprar/vender a qualquer momento. **É onde o operador prefere operar.**
  - **Kalshi** (e Polymarket) — mercado público, mais líquido e eficiente. Serve de **âncora de valor justo**.
- **Objetivo do projeto:** precificar a **maior probabilidade que o país pode chegar a ter** durante os jogos (o "topo" onde se venderia uma posição comprada), com base nas informações públicas de hoje.

## 2. A tese central (não esquecer)

A probabilidade de ser campeão é um **martingale**: o preço justo de hoje já embute o resultado esperado. **Segurar até o pico não gera lucro esperado num mercado eficiente** — o ganho esperado de qualquer estratégia de saída é zero. O **edge real** vem de:
1. **Ineficiência do Call da Copa** — amigos reagem mais devagar a gols/resultados que o Kalshi e superprecificam favoritos emocionais (anfitriões, o "queridinho").
2. **Usar o modelo do pico** para saber ONDE e QUANDO sair.

As métricas de pico **não criam valor sozinhas**; elas dimensionam risco, definem alvos de saída e revelam onde o preço dos amigos está errado.

## 3. O modelo matemático

Com `p` = probabilidade justa atual (Kalshi, sem vig):

- **Pico esperado:** `E[max] = p · (1 − ln p)` — média do máximo que a prob deve tocar no torneio. (Deriva de `P(max ≥ a) = p/a` para martingale em [0,1].)
- **Teto na final (melhor caso):** `p^(1/N)` — a maior prob que o país PODE chegar a ter (véspera da final, tendo vencido tudo). `N` = vitórias-equivalentes até o título.
- **Prob. de atingir um nível:** `P(max ≥ X) = p / X` — chance de a prob algum dia negociar em X% ou mais. Use para posicionar ordens de venda limite.
- **Premissa N** (cai conforme avança): grupos ≈ 6, oitavas = 5, quartas = 4, semi = 3, final = 2.

### Verificação de sanidade (martingale)
Buy at p, sell limit at a: prob `p/a` de ganhar `(a−p)`; senão fica abaixo. Valor esperado de saída = p ⇒ lucro esperado = 0. Confirma "sem almoço grátis". Todo o edge é mispricing entre os dois mercados.

## 4. Regras de operação

- **COMPRA (zebras de pico relativo):** preferir seleções a ~1–3%. Múltiplo pico/preço alto (5x+), custo de carregar a opção baixo. Comprar quando o Call da Copa estiver ABAIXO da "Prob justa".
- **VENDA (timing do pico):** a prob bate o topo no **run-up ANTES de um jogo difícil** (mercado precifica o próximo adversário). Vender em camadas a cada vitória (aba Escada de Alvos). **Nunca segurar uma zebra ATRAVÉS de um jogo contra time mais forte** — downside brutal (cai pra ~0).
- **FADE (anfitriões/queridinhos):** México, EUA, Canadá e o time da torcida ficam caros no Call da Copa. Vender o hype quando o preço lá passar da "Prob justa".
- **Arbitragem:** Kalshi = referência. Quando o Call da Copa divergir, montar a posição lá no lado barato; se a divergência for grande, travar o spread.

## 5. Dados usados (snapshot 13/jun/2026 — fase de grupos)

Odds **americanas agregadas Kalshi/Polymarket** (fonte: defirate.com). Vig ~2,8%. Converter: `prob bruta = 100/(odds+100)`, depois normalizar pela soma.

| Seleção | Odds | p justa | Pico esp. E[max] | Teto final |
|---|---|---|---|---|
| Espanha | +494 | 16,4% | 46,0% | 74% |
| França | +518 | 15,7% | 44,8% | 73% |
| Portugal | +844 | 10,3% | 33,7% | 68% |
| Inglaterra | +925 | 9,5% | 31,8% | 68% |
| Brasil | +1095 | 8,1% | 28,6% | 66% |
| Argentina | +1151 | 7,8% | 27,6% | 65% |
| Alemanha | +1831 | 5,0% | 20,1% | 61% |
| Holanda | +1934 | 4,8% | 19,3% | 60% |
| Noruega | +3986 | 2,4% | 11,3% | 54% |
| Bélgica | +4769 | 2,0% | 9,8% | 52% |
| Japão | +4837 | 2,0% | 9,7% | 52% |
| EUA | +5013 | 1,9% | 9,4% | 52% |
| Colômbia | +5614 | 1,7% | 8,6% | 51% |
| México | +6256 | 1,5% | 7,9% | 50% |
| Suíça | +6802 | 1,4% | 7,4% | 49% |
| Marrocos | +6830 | 1,4% | 7,4% | 49% |

Demais (Turquia/Uruguai +10426 ≈ 0,9% até Uzbequistão +199900 ≈ 0,05%) na planilha completa. Total = 48 seleções.

**Top zebras (sweet spot):** Noruega, Bélgica, Japão, EUA, Colômbia, Marrocos — múltiplo 5x+ com pico absoluto relevante.

## 6. Entregável

`Copa_2026_Pico_Probabilidade_Campeao.xlsx` (na pasta do projeto). 4 abas, todas em fórmula Excel:
- **Painel** — 48 seleções, prob justa + métricas de pico. Coluna B (odds) = input azul editável; premissa N na célula E3.
- **Escada de Alvos** — trajetória da prob a cada vitória (degraus de venda).
- **Zebras** — screen do edge com leitura tática.
- **Estratégia & Como Usar** — regras completas.

**Atualizar:** colar odd nova do Kalshi na coluna B do Painel → ajustar N (E3) → recalcular. Tudo se propaga.

## 7. Fontes

- defirate — World Cup Odds Tracker: https://defirate.com/prediction-markets/world-cup-odds/
- Kalshi — Men's World Cup Winner: https://kalshi.com/markets/kxmenworldcup/mens-world-cup-winner/kxmenworldcup-26

## 8. Log de Trades

Registro cronológico de todas as operações no Call da Copa e Kalshi.

| Data | Seleção | Lado | Contratos | Preço | Contraparte | Nota |
|---|---|---|---|---|---|---|
| 21/jun/2026 | Alemanha | V | 10 | 6.00 | Ricardo Espindola | Redução parcial do long (era 40C@8.25 → ficou 30C@9.00) |

