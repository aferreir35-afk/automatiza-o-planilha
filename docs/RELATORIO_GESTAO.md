# ZBR Importado — Relatório de Gestão (versão executiva)

Arquivo: `planilhas/ZBR_Importado_Relatorio_Gestao.xlsx`
Gerador: `scripts/gerar_relatorio.py` (lê `planilhas/origem/Analise_ZBR_importado.XLSX`)
Backup da versão anterior: `planilhas/backup/Analise_ZBR_Cruzamento_backup_2026-10-02.xlsx`

## Estrutura

| Aba | Conteúdo |
|---|---|
| MENU | Identificação do relatório, navegação com links, legenda de cores, como usar |
| 1 Dashboard | Filtros únicos (período, SKU, lote, unidade, tipo de movimentação, status), 8 KPIs, 6 quantidades, 7 gráficos |
| 2 Resumo Gerencial | Escopo, conciliação por unidade, causas, criticidade, top 5, depósito 2008, plano, conclusões por fórmula |
| 3 Base Entrada 2008 / Saídas / Estoque Físico / Movimentos | Dados SAP em tabelas estruturadas (tbEntrada2008, tbSaidas, tbEstoque, tbMovimentos) |
| 3 Base Saldo Inicial 08-09 | Estoque de 08/09 por SKU + lote + depósito (preencher). Já integrado: entra no Deveria ter do ZBR e no saldo do 2008 |
| 4 Análise ZBR | Tabela por SKU + lote seguindo os filtros do Dashboard (motivo e solução) |
| 4 Análise Depósito 2008 | Recebido, transferido, estornos, saldo e próximo passo + gráficos do 2008 |
| 5 Conciliação | Diferença, motivo, solução, componentes da explicação, categoria e criticidade (proposta) |
| 6 Plano de Ação | 61 ocorrências com status, prazos e alertas (tbPlanoAcao) |
| 7 Dicionário de Dados | Abas, campos, regras, critérios, atualização, limitações |
| 8 Atualizar Dados | Passo a passo e código das 4 consultas Power Query (também em `powerquery/*.pq`) |
| Parâmetros | Data-limite, tolerância, limites de criticidade, datas |

## Validação feita

- 12.126 fórmulas recalculadas sem erro.
- Bases com o mesmo nº de linhas e a mesma soma de quantidades da origem
  (Entrada 201 linhas / 5.631,47; Saídas 405 / −29.384,8; Estoque 228 / 63.256,456).
- KPIs do Dashboard conferidos com cálculo independente (pandas) na origem:
  entrou 65.930,28 KG, saiu 26.268,8 KG, estoque 59.270,487 KG,
  diferença +19.609,007 KG, volume 92.199,08 KG, 526 movimentações em KG.
- Filtros testados (status "Faltando", alteração de status e prazo no plano).

## Limitações

Ver aba "7 Dicionário de Dados", seção 6 (saldo inicial ausente, listas de
SKU·lote geradas pelo script, sem Tabela Dinâmica/slicers, MINIFS/MAXIFS
exigem Excel 2019/365, data exibida = data de geração).

## Recomendações implementadas

1. **Saldo inicial de 08/09**: aba `3 Base Saldo Inicial 08-09` (tabela `tbSaldoInicial`).
   Deveria ter (ZBR) = saldo inicial + entrou − saiu; saldo do 2008 = saldo inicial + recebido − transferido + estornos.
   Vazia = comportamento anterior. Testado: com 5.328,4 KG informados para 811-BR-D94 / 251643289,
   a diferença cai de 5.737,2 para 408,8 KG.
2. **Power Query**: consultas `qEntrada2008`, `qSaidas`, `qEstoque`, `qSaldoInicial` (pasta `powerquery/`),
   lendo o arquivo indicado em Parâmetros!K13 (`CaminhoSAP`) e a data-limite (`DataLimiteRecebimento`).
   Produzem as mesmas colunas das bases, já com as colunas calculadas. Não executadas aqui
   (Power Query só roda no Excel): testar na primeira carga.
3. **Visão simples**: "Resultado em 1 minuto" no MENU; colunas técnicas da conciliação recolhidas (botão "+").
