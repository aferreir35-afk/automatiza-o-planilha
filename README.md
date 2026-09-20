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

## Análise de Capacidade — GM (09 a 21/09/2026)

Arquivo: [`planilhas/Analise_Capacidade_GM_09_a_21.xlsx`](planilhas/Analise_Capacidade_GM_09_a_21.xlsx)

Análise do volume bruto de fornecimentos (base `Analise_GM_DO_DIA_01_AO_DIA_20.XLSX`,
enviada pelo usuário) frente à capacidade diária declarada de 35–40 t e à
equipe de separação (4 separadores + 3 operadores por turno, escala 2x2 de
12h — 2 turnos/dia cobrindo as 24h, ou seja, 8 separadores e 6 operadores em
atividade por dia).

- Aba **Dados_Filtrados**: 1.379 fornecimentos com "Data do picking" entre
  09/09 e 21/09/2026, extraídos do arquivo original.
- Aba **Dimensionamento**: premissas de equipe/escala (células amarelas,
  editáveis) e capacidade/produtividade implícita, tudo por fórmula.
- Aba **Analise_Diaria**: volume bruto por dia (`SUMIFS`/`COUNTIFS` sobre
  Dados_Filtrados), comparação com a faixa de 35–40 t, status
  (Acima/Dentro/Abaixo, com formatação condicional) e estimativa de
  separadores necessários/gap por dia.
- Aba **Painel**: KPIs do período, gráfico de volume diário x faixa de
  capacidade e uma leitura rápida dos principais achados.

**Principais achados:** o volume oscilou entre 0,6 t e 127,2 t/dia — muito
acima ou abaixo da faixa de 35–40 t. Sete dias (12, 14 a 19/09) ficaram
acima da capacidade máxima (alguns em mais de 3x, ex.: 14/09 com ~127 t),
enquanto os domingos (13 e 20/09) ficaram bem abaixo. Isso indica
concentração de "data do picking" em poucos dias, não uma demanda
uniformemente distribuída — recomenda-se validar com a operação se é pico
real de demanda ou efeito de atualização em lote do status no sistema, e
medir a produtividade real por separador (apontamento de mão de obra) para
substituir a produtividade implícita usada na estimativa de gap de equipe.

## Dashboard de Estoque LX03 — Barry Callebaut (com automação em Python)

Pasta: [`lx03/`](lx03/README.md)

Automação em Python + Excel para transformar o export da transação SAP
**LX03** (estoque por posição no depósito) num painel diário com os pontos
de atenção já destacados (vencidos, bloqueados, em quarentena, parados sem
giro, etc.), filtros por Tipo de Depósito/Centro/Alerta, gráficos e
exportação em PDF — pronto para acompanhamento operacional e apresentação à
diretoria. Rodar `python lx03/scripts/atualizar_dashboard.py <export do
dia>.xlsx` é o único passo do dia a dia.

Também inclui, todos opcionais e desligados por padrão: **agendamento
automático** (Task Scheduler/cron), **alertas por e-mail/Teams** quando um
ponto de atenção piora, **histórico de tendência** (SQLite + aba com
gráfico de evolução diária), **conexão direta ao SAP** (OData/RFC, sem
exportar manualmente) e **publicação em Google Sheets/Power BI**. Veja
[`lx03/README.md`](lx03/README.md) para os detalhes completos, o que já foi
testado (com servidores simulados, sem credenciais reais) e as limitações
assumidas (inclusive sobre a paleta de cores da marca, que não pôde ser
confirmada oficialmente).
