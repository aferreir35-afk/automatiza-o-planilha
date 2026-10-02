# Relatório de Gestão — ZBR Importado (versão atual)

Arquivo principal: [`planilhas/ZBR_Importado_Relatorio_Gestao.xlsx`](planilhas/ZBR_Importado_Relatorio_Gestao.xlsx)
— dashboard executivo, resumo gerencial, bases SAP estruturadas, análises,
conciliação, plano de ação e dicionário de dados. Detalhes em
[`docs/RELATORIO_GESTAO.md`](docs/RELATORIO_GESTAO.md).

**Versão simples (para quem não é da área):** [`planilhas/ZBR_Importado_Versao_Simples.xlsx`](planilhas/ZBR_Importado_Versao_Simples.xlsx)
— 5 abas (Comece Aqui, Produtos, O Que Fazer, Depósito 2008, Palavras), linguagem do dia a dia,
mesmos números do relatório completo.

Gerar de novo: `python scripts/gerar_relatorio.py [origem.xlsx] [saida.xlsx]` e depois
`python scripts/gerar_versao_simples.py` (após recalcular o relatório completo).

---

# Automação de Planilha — Volume por Turno

Planilha de controle de volume de carga por turno/divisão, com o **Painel**
("dashboard") ampliado para monitorar variância entre Planejado x Realizado
usando apenas fórmulas e formatação condicional do próprio Excel — **sem
servidor, add-in ou integração externa (e-mail/Slack/WhatsApp)**.

Arquivo: [`planilhas/Volume_por_Turno_Setembro_2026.xlsx`](planilhas/Volume_por_Turno_Setembro_2026.xlsx)

## O que já existia

- Aba **Lançamentos**: uma linha por turno/dia, com Atingimento %, Saldo (t)
  e Observação calculados automaticamente.
- Aba **Painel**: Totais do mês, Resumo por Divisão, Resumo por Turno e o
  gráfico Planejado x Realizado por Divisão.

## O que foi adicionado ao Painel

1. **Formatação condicional (vermelho/amarelo/verde)** nos percentuais de
   Atingimento já existentes (geral, por Divisão e por Turno) — mesmo padrão
   de cores já usado na coluna `Atingimento %` da aba Lançamentos
   (vermelho < 90%, amarelo 90–99%, verde ≥ 100%).

2. **Configurações de Alerta** (células amarelas, editáveis):
   - `Meta mínima de Atingimento (%)` — limite usado em todos os alertas do
     painel (padrão 90%).
   - `Limite de Saldo Semanal (t)` — limite usado no alerta de saldo da
     semana (padrão 5.000 t).

3. **Alertas por Divisão** e **Alertas por Turno** — tabelas com status
   automático: `✅ OK`, `🔴 ALERTA` (Atingimento abaixo da meta) ou
   `Sem dados` (sem volume planejado no período). A linha inteira fica
   destacada em vermelho claro quando há alerta.

4. **Maiores Desvios do Mês (Top 3 Divisões)** — ranking automático (via
   `LARGE`/`INDEX`/`MATCH`) das divisões com maior Saldo (maior défice entre
   planejado e realizado). Em caso de empate de saldo, o ranking pode repetir
   a divisão — limitação aceita para manter a fórmula simples.

5. **Resumo da Semana Atual** — calcula a semana corrente (segunda a
   domingo) a partir da última data lançada, soma Planejado/Realizado/Saldo
   da semana e mostra `🔴 ALERTA` se o saldo semanal ultrapassar o limite
   configurado.

6. **Projeção de Fechamento do Mês** — projeta Planejado, Realizado, Saldo e
   Atingimento % do mês inteiro a partir do ritmo diário observado até a
   última data lançada (`total até agora ÷ dias decorridos × dias do mês`) e
   sinaliza `🔴 ALERTA` se a projeção de Atingimento ficar abaixo da meta.

## Como usar

- Continue preenchendo apenas as células amarelas da aba **Lançamentos**
  (Data, Turno, Divisão, Volume Planejado, Volume Realizado) — todo o resto,
  incluindo os alertas do Painel, é recalculado automaticamente pelo Excel.
- Para mudar os limites de alerta, edite `Meta mínima de Atingimento (%)` e
  `Limite de Saldo Semanal (t)` na aba Painel.
- Abra a planilha no Excel/Google Sheets normalmente — não é necessário
  macro, script ou qualquer serviço externo.

## Limitações assumidas (mantendo a solução simples)

- O alerta semanal usa semana de calendário (segunda a domingo) com base na
  última data lançada, não uma janela contínua dos últimos 7 dias.
- A projeção de fechamento assume ritmo diário constante; não usa média
  móvel nem desvio-padrão.
- O ranking de maiores desvios não trata empates de forma especial (pode
  repetir uma divisão quando o Saldo é igual).

---

# Análise ZBR Importado — Entradas × Saídas × Estoque

Arquivo: [`planilhas/Analise_ZBR_Cruzamento_Entrada_Saida_Estoque.xlsx`](planilhas/Analise_ZBR_Cruzamento_Entrada_Saida_Estoque.xlsx)

Cruza as abas **Entrada** (movimentos no depósito 2008), **Saida** (movimento
601) e **Estoque atual** (depósitos 9999/0355/9991) por Material e por
Material + Lote, só com fórmulas (`SUMIFS`/`COUNTIF`).

Todas as abas trazem **SKU** e **Lote** lado a lado.

- **Painel**: filtros em lista (SKU, Lote — só os lotes do SKU escolhido —,
  Data inicial/final, Unidade, Movimentação), quadros Entrou / Saiu /
  Deveria ter / Estoque físico / Diferença, contagem de SKU·lote que
  conferem, sobram ou faltam, gráficos e tabela por SKU + lote.
- **Painel**: o cabeçalho da tabela (SKU, Lote, … Motivo, Solução) fica
  congelado e tem filtro; colunas de 1º e último lançamento e **Solução de
  ajuste (sugestão)**. Cores em tons de chocolate (paleta inspirada na
  Barry Callebaut).
- **Gráficos**: gráficos simples que seguem os filtros do Painel (Entrou ×
  Saiu × Deveria ter × Estoque físico, situação dos SKU·lote, maiores
  diferenças e movimentação por dia).
- **Depósito 2008**: consulta do depósito 2008 com filtros (SKU, lote —
  dependente do SKU —, datas, unidade, situação), quadros de Recebido /
  Transferido ao ZBR / Estornos / Saldo no 2008 / Lançamentos e tabela por
  SKU + lote (cabeçalho congelado e com filtro) com % transferido, situação
  (Aguardando transferência / Transferido 100% / Saiu saldo anterior a 09/09)
  e próximo passo.
- **Dash 2008**: gráficos que seguem os filtros da aba Depósito 2008.
- **Cruzamento**: uma linha por SKU + lote, com Diferença, Situação, **Motivo
  da diferença** e O que fazer. O motivo é calculado pelos dados: estoque
  antigo (posições sem movimento desde antes de 09/09), doca/chão/
  transferência (DCK, TRF, 922, FLR, DIF), troca de lote, saídas que não
  baixaram o estoque (inclusive pelo depósito 0355) ou saída não lançada.
  A coluna "Diferença sem explicação" mostra o que sobra para conferir.
- **Lançamentos**: cada movimento (Entrada, Saída, Recebido no 2008, Estorno).
- **Estoque físico**: cada posição no depósito.
- **Entrada** e **Saida** (cinza): dados originais do SAP.
- Abas ocultas (Resumo, Cruzamento por Material/Lote): cálculos de apoio.

Regra usada: no 2008, valores negativos são transferências para o ZBR
(entradas); positivos até 15/09/2026 são chegadas da importação e positivos
depois disso são estornos. Saldo calculado = Entradas − Saídas; Diferença =
Estoque atual − Saldo calculado. Diferenças positivas costumam indicar
estoque anterior a 09/09 (fora do período das abas).

## Painel interativo

[`painel/painel-zbr.html`](painel/painel-zbr.html) — painel (HTML, abre no
navegador) com filtros por produto, lote, período, unidade e tipo de
movimentação; indicadores de Entrou / Saiu / Deveria ter / Estoque físico /
Diferença, gráfico diário, ranking de diferenças e abas Cruzamento,
Lançamentos e Estoque físico. Botão "Depósito 2008" no topo abre a visão do
depósito 2008 (quadros, gráfico diário, aguardando transferência e tabelas).
