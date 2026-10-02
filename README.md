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

- **Análise Simples** (primeira aba, verde): uma linha por produto com
  Entrou, Saiu, Deveria ter, Tem no estoque, Diferença, Situação
  (Confere / Sobrando / Faltando) e O que fazer, em linguagem direta.
- **Resumo**: parâmetros editáveis (data-limite dos recebimentos de importação
  e tolerância), contagem OK / Sobra / Falta e totais por UM (KG, UN, CX).
- **Cruzamento por Material** e **Cruzamento por Lote**: Recebido no 2008,
  Entradas no ZBR, Saídas, Saldo calculado, Estoque atual, **Diferença**,
  Status e possível causa — ordenados da maior para a menor diferença.
- **Consulta**: filtros de Data inicial/final, Material (lista) e Lote
  (`*` = todos) com totais por UM e tabela por material no período escolhido.
- **Movimentações**: entradas e saídas numa tabela única do Excel (com
  botões de filtro) — Data de lançamento, Tipo, Material, Descrição, Lote,
  Depósito, Quantidade, Entrada/Saída ZBR, Doc., Cliente e Usuário.

Regra usada: no 2008, valores negativos são transferências para o ZBR
(entradas); positivos até 15/09/2026 são chegadas da importação e positivos
depois disso são estornos. Saldo calculado = Entradas − Saídas; Diferença =
Estoque atual − Saldo calculado. Diferenças positivas costumam indicar
estoque anterior a 09/09 (fora do período das abas).
