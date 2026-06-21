# Publicar o dashboard no GitHub Pages (atualização de hora em hora)

Este guia coloca o `kalshi_dashboard.html` no ar num link público, atualizado
automaticamente a cada hora — **sem depender do seu PC ficar ligado**. Quem faz
o trabalho é o GitHub Actions: de hora em hora ele busca as odds da Kalshi,
regenera o HTML e republica no GitHub Pages.

## O que já foi preparado nesta pasta

- `.github/workflows/deploy.yml` — a automação (cron de hora em hora + deploy).
- `requirements.txt` — dependência do script (`requests`).
- `.gitignore` — **protege seus dados privados**: a planilha de boletagem, o
  `call_prices.json`, o histórico e o `kalshi_dashboard.html` local **não** vão
  pro repositório. O site é gerado do zero no CI só com odds públicas da Kalshi,
  então os campos do "Call" aparecem em branco e são preenchidos pelo
  `localStorage` do seu navegador.

## Privacidade — leia antes

O GitHub Pages gratuito serve a partir de um **repositório público**. Graças ao
`.gitignore`, nada sensível é enviado: vão pro GitHub apenas `kalshi_update.py`,
o workflow e os arquivos de apoio. Se quiser repo privado **e** site público,
isso exige um plano pago do GitHub (Pro/Team).

---

## Passo a passo

### 1. Crie o repositório no GitHub
1. Acesse https://github.com/new (logado).
2. Nome: ex. `copa-2026-dashboard`. Deixe **Public**.
3. Não marque "Add a README". Clique em **Create repository**.

### 2. Suba os arquivos (escolha A ou B)

**A) Linha de comando** (você já tem git 2.34). No terminal, dentro desta pasta:

```bash
cd "C:\Users\pavan\Claude\Projects\Copa do Mundo 2026"
git init
git add .
git commit -m "Dashboard Kalshi + deploy automatico"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/copa-2026-dashboard.git
git push -u origin main
```

> Confira **antes do push** que nada sensível entrou:
> `git status` e `git ls-files`. A planilha, o `call_prices.json` e o histórico
> **não** podem aparecer na lista (o `.gitignore` cuida disso).

**B) GitHub Desktop** — "Add Local Repository" → aponte pra esta pasta →
"Publish repository" (mantenha **público**).

### 3. Ligue o GitHub Pages
1. No repositório: **Settings → Pages**.
2. Em **Build and deployment → Source**, selecione **GitHub Actions**.
   (Não escolha "Deploy from a branch".)

### 4. Rode a primeira publicação
- O push da etapa 2 já dispara o workflow. Veja em **Actions**.
- Se quiser forçar: aba **Actions → "Atualizar e publicar dashboard" → Run workflow**.
- Ao terminar (~1 min), o link aparece em **Settings → Pages**, no formato:

  `https://SEU_USUARIO.github.io/copa-2026-dashboard/`

Pronto — esse é o link de acesso pelo navegador, de qualquer lugar.

---

## Como funciona a atualização de hora em hora

O bloco `schedule` no workflow usa `cron: "0 * * * *"` → roda no minuto 0 de
toda hora. Pontos importantes do GitHub Actions:

- **Fuso é UTC.** "Toda hora cheia" vale igual em qualquer fuso, então não muda nada
  pro seu caso (atualiza 24x/dia).
- **Pode atrasar** alguns minutos em horários de pico do GitHub — é normal e gratuito.
- **Pausa após 60 dias sem atividade** no repositório. Se o projeto ficar parado
  meses até a Copa, o cron pode hibernar. Para reativar, basta abrir **Actions →
  Run workflow** uma vez (ou fazer qualquer commit). Como você vai mexer no
  projeto com frequência, dificilmente vai esbarrar nisso.

## Ajustes rápidos

- **Mudar a frequência:** edite o `cron` em `.github/workflows/deploy.yml`.
  Ex.: a cada 30 min → `"*/30 * * * *"`; só de 6 em 6 horas → `"0 */6 * * *"`.
- **Atualizar agora:** aba **Actions → Run workflow** (botão "Run workflow").
- **Domínio próprio:** Settings → Pages → "Custom domain".

## Limitação conhecida

Rodando sem a planilha (`--no-xlsx`), o `N` do "Teto = p^(1/N)" usa o padrão
**N = 6** embutido no script, em vez de ler o valor do seu Painel. Se quiser que
o site reflita um N diferente, me avise que ajusto o script pra fixar o N certo
no modo de deploy.
