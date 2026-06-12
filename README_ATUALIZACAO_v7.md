# Atualização v7 — Correção de fonte dos títulos + revisão do Dashboard + indicador da sidebar

## Arquivos alterados
- `src/App.jsx`
- `index.html` (carrega a fonte usada nos títulos)

## O que estava errado
- A v6 usava `'SF Pro Rounded'`/`ui-rounded` como fonte dos títulos. Essas
  fontes só existem no macOS/Safari — em Windows/Chrome/Firefox o navegador
  ignora e usa a mesma fonte padrão de sempre, por isso **nenhuma mudança
  era visível**.
- O Dashboard tinha sido apenas re-estilizado (cores/ícones), sem mudança
  real de **layout/organização** dos dados.
- A barra lateral, quando oculta, não dava nenhuma pista de que existia —
  o usuário precisava "adivinhar" que passar o mouse na borda esquerda
  revelava o menu.

## O que foi corrigido / incluído

### 1. Fonte dos títulos (agora visível em qualquer sistema)
- `index.html` agora carrega a fonte **Manrope** (Google Fonts, pesos
  500/700/800/900) — tipografia geométrica/arredondada, próxima ao espírito
  "SF Pro Rounded" usado pela Apple em dashboards e títulos.
- Aplicada em: títulos das abas (Dashboard, Pipeline, Histórico,
  Configurações), logotipo "Fluxo de Liberação" na lateral, títulos de cada
  etapa do Pipeline, números grandes dos KPIs e títulos de seção do
  Dashboard.

### 2. Dashboard — layout reorganizado (estilo relatório executivo)
- Cabeçalho "Visão Geral" maior, com data/hora de atualização.
- **Novo: "Funil do Processo"** — gráfico de barras horizontal em largura
  total, mostrando a quantidade de tubos em cada uma das 8 etapas lado a
  lado (estilo funil executivo), clicável para ir à etapa.
- "Pendências por Departamento" e "Bloqueio por Depósito" lado a lado
  (50/50), cada um com cabeçalho de seção com ícone colorido.
- "Tempo médio por etapa" em largura total, com "Por etapa" e "Por
  departamento" lado a lado internamente quando há dados de ambos.

### 3. Indicador visual da barra lateral oculta (NOVO)
- Foi adicionada uma "alça" (handle) fixa, sempre visível, na borda esquerda
  da tela (centralizada verticalmente): um pequeno retângulo com cantos
  arredondados à direita, efeito de vidro fosco, seta ">" e três pontinhos —
  sinalizando claramente que há um menu oculto ali.
- Ao passar o mouse sobre essa alça (ou na borda esquerda), a barra lateral
  desliza para fora; a alça desaparece suavemente enquanto a barra está
  aberta e reaparece quando ela se recolhe.
