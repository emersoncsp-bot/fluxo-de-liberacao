# Atualização v8 — Ajustes conforme "update_V7.docx"

## Arquivo alterado
- `src/App.jsx`

(`index.html`, `package.json`, `src/index.css` permanecem como na v7)

## Ajustes implementados

### 1. Pipeline — remoção dos títulos destacados
- Removido o título grande de cada etapa (ex: "Definição do Recurso",
  "Criação de Ordem" etc.) e o título "Importar Tubos Bloqueados" na
  Etapa 1 — em todas as etapas da Pipeline, no mesmo local.

### 2. Pop-ups — fonte do título
- O título de todos os pop-ups (confirmações de avançar/retornar/concluir
  em qualquer etapa da Pipeline, e os modais de Configurações: novo/editar
  usuário, excluir, logout) agora usa a mesma fonte do logotipo "Fluxo de
  Liberação" da barra lateral (Manrope/TITLE_FONT).

### 3. Configurações — nova aba "Faturamento"
- Nova aba ao lado de "Usuários" e "Permissões".
- Cadastro manual (Pedido, Item, Descrição) ou importação via CSV/XLS/XLSX
  (reaproveita o mesmo mapeamento de colunas da importação da Pipeline).
- Tabela com Pedido/Item, Descrição, contagem de "Tubos" (soma em todas as
  etapas da Pipeline cujo Pedido/Item corresponda) e botão Excluir.

### 4. Destaque de Faturamento na Pipeline
- Lotes cujo Pedido/Item esteja cadastrado em Faturamento aparecem com fundo
  levemente vermelho na tabela de qualquer etapa.
- Legenda "Faturamento" (mesmo estilo/fonte das etiquetas de departamento,
  ex. "Controle da Qualidade", em preto) exibida no topo da tabela quando
  houver pedidos cadastrados.

### 5. Importação de arquivo — Pedido/Item unificado
- Nas tabelas da Pipeline, pré-visualização de importação e Histórico, as
  colunas "Pedido" e "Item" foram unificadas em uma única coluna
  "Pedido/Item" (formato `pedido/item`).

### 6. Pesquisa da Pipeline agora filtra a tabela
- Antes, a busca só atualizava os contadores (badges) da barra de etapas.
  Agora os critérios informados (Pedido/Item, Lotes, IPPNs, Depósito) também
  filtram as linhas exibidas na tabela da etapa atual.

### 7. Dashboard — novo gráfico de Pareto de Motivos de Bloqueio
- Novo card "Motivos de Bloqueio", abaixo de "Tempo médio por etapa",
  baseado na coluna "Motivo Bloqueio Texto": barras ordenadas por
  quantidade, com % acumulado, exibindo apenas os motivos que somam até 80%
  do total (princípio de Pareto).

### 8. Dashboard interativo (estilo Power BI)
- Clicar em uma etapa do "Funil do Processo" agora filtra o dashboard por
  aquela etapa (clique novamente para remover o filtro). A etapa
  selecionada fica destacada e um chip "Filtro: Etapa X ✕" aparece para
  limpar o filtro.
- Com o filtro ativo, "Bloqueio por Depósito" e o novo gráfico "Motivos de
  Bloqueio" recalculam para mostrar apenas os dados da etapa selecionada.
- KPIs, "Pendências por Departamento" e "Tempo médio por etapa" continuam
  exibindo a visão geral (não filtrada).
