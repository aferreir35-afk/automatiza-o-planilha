# Automação de Planilha

Três planilhas, todas usando apenas fórmulas e formatação condicional do
próprio Excel — **sem servidor, add-in ou integração externa
(e-mail/Slack/WhatsApp)**:

1. [`planilhas/Volume_por_Turno_Setembro_2026.xlsx`](planilhas/Volume_por_Turno_Setembro_2026.xlsx)
   — controle de volume de carga por turno/divisão (Planejado x Realizado).
2. [`planilhas/Painel_Programacao_Pendente_Setembro_2026.xlsx`](planilhas/Painel_Programacao_Pendente_Setembro_2026.xlsx)
   — painel de programação pendente (Gerar/Separar/Faturar/Faturado) x
   capacidade da escala de turnos. Ver seção própria mais abaixo.
3. [`planilhas/Dados_Programacao_Correlacionados_Setembro_2026.xlsx`](planilhas/Dados_Programacao_Correlacionados_Setembro_2026.xlsx)
   — planilha de dados: a mesma correlação sqvi/GM/VL06 do item 2, só que em
   formato de tabela plana (uma linha por remessa), pronta para tabela
   dinâmica/Power BI, com dicionário de dados. Ver seção própria mais abaixo.

## 1) Volume por Turno

### O que já existia

- Aba **Lançamentos**: uma linha por turno/dia, com Atingimento %, Saldo (t)
  e Observação calculados automaticamente.
- Aba **Painel**: Totais do mês, Resumo por Divisão, Resumo por Turno e o
  gráfico Planejado x Realizado por Divisão.

### O que foi adicionado ao Painel

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

### Como usar

- Continue preenchendo apenas as células amarelas da aba **Lançamentos**
  (Data, Turno, Divisão, Volume Planejado, Volume Realizado) — todo o resto,
  incluindo os alertas do Painel, é recalculado automaticamente pelo Excel.
- Para mudar os limites de alerta, edite `Meta mínima de Atingimento (%)` e
  `Limite de Saldo Semanal (t)` na aba Painel.
- Abra a planilha no Excel/Google Sheets normalmente — não é necessário
  macro, script ou qualquer serviço externo.

### Limitações assumidas (mantendo a solução simples)

- O alerta semanal usa semana de calendário (segunda a domingo) com base na
  última data lançada, não uma janela contínua dos últimos 7 dias.
- A projeção de fechamento assume ritmo diário constante; não usa média
  móvel nem desvio-padrão.
- O ranking de maiores desvios não trata empates de forma especial (pode
  repetir uma divisão quando o Saldo é igual).

## 2) Painel de Programação Pendente (Gerar · Separar · Faturar · Faturado)

Construída a partir de um extrato SAP enviado com 3 abas (`sqvi` = faturamento,
`GM` = programação de embarque, `VL06` = monitor de remessas/WM), correlacionadas
pelo número de remessa para classificar cada pedido programado em uma etapa do
fluxo: **Gerar → Separar → Faturar → Faturado**.

### Abas do arquivo

- **Leia-me** — explica a correlação entre as abas, a regra de classificação
  de estágio e os principais achados (inclusive de qualidade de dado).
- **Painel** — dashboard executivo (estilo Power BI): KPIs do backlog por
  etapa, gráfico de capacidade x demanda ao longo do período, status dos
  dias e top 5 UF/transportadora por peso pendente.
- **Escala_Capacidade** — calendário com a escala 2x2x12h (turnos A/B/C/D,
  revezando em blocos de 2 dias) x capacidade por turno (35–40 t, ajustável
  pela presença de separadores/operadores de empilhadeira) x demanda
  pendente programada por dia.
- **Backlog_Detalhado** — lista completa e filtrável (9.953 remessas
  correlacionadas) com cliente, cidade/UF, transportadora, estágio, atraso e
  peso.
- **Qualidade_Dados** — ressalvas da correlação (pesos de remessa
  fisicamente implausíveis, reagendamentos, remessas fora da programação
  oficial, janela de confiabilidade do extrato VL06 e erros de fórmula já
  existentes no arquivo original).
- **sqvi / GM / VL06** — abas originais enviadas, mantidas para
  rastreabilidade.

### Como usar

- Todos os KPIs, o calendário de capacidade e os gráficos são fórmulas
  (`SUMIFS`/`SUMIF`/`COUNTIF`) que recalculam automaticamente ao abrir no
  Excel — não é necessário macro, script ou serviço externo.
- Edite as células amarelas na aba **Escala_Capacidade** (capacidade
  mín./máx. por turno, separadores/operadores necessários, data-base da
  escala, ou a presença real de cada turno em cada dia) para simular faltas
  ou ajustar parâmetros — o painel inteiro recalcula.
- Para atualizar com um novo extrato SAP, seria necessário reexportar as 3
  abas de origem e reprocessar a correlação (feita nesta versão via script;
  não há um botão de atualização dentro do próprio Excel).

### Premissas e limitações assumidas

- A ausência de uma remessa na `VL06` só é tratada como "Pendente Gerar"
  dentro da janela efetivamente coberta pelo extrato (30/08 a 30/10/2026);
  fora dela, é marcada como "Fora da janela analisada" e excluída dos KPIs,
  pois o extrato é um retrato pontual do que está em aberto no SAP.
- Peso de remessa acima de 20 t é tratado como provável erro de
  fórmula/lookup na origem (perfil de cliente é varejo/confeitaria de
  pequeno porte) e excluído dos totais de capacidade — sinalizado à parte
  na aba Qualidade_Dados.
- A escala 2x2x12h assume 2 turnos ativos por dia (um diurno, um noturno),
  revezando em blocos de 2 dias entre os pares informados (padrão: C/D e
  A/B) a partir de uma data-base editável.
- A capacidade efetiva de cada turno usa a média entre o mínimo e o máximo
  informados (35–40 t) escalada pelo pior entre presença de separadores e
  de operadores de empilhadeira — não modela produtividade individual.

## 3) Planilha de Dados (Correlação sqvi/GM/VL06)

Mesma correlação e mesmas regras de classificação do painel do item 2, mas
entregue como tabela plana — uma linha por remessa — em vez de dashboard.
Pensada para quem quer os dados "crus" já correlacionados para montar tabela
dinâmica, gráfico próprio ou carregar num Power BI.

### Abas do arquivo

- **Resumo** — totais por Estágio e por mês de embarque (fórmulas
  `COUNTIFS`/`SUMIFS` sobre a aba Dados).
- **Dados** — a tabela (9.953 remessas), com Estágio, Situação Temporal,
  Dias em Atraso, Peso e Peso Válido/Pendente já calculados, em formato de
  Tabela do Excel (filtro automático, cabeçalho fixo).
- **Dicionário** — o que é cada coluna e a regra de classificação do
  Estágio (mesma regra do painel do item 2).

### Como usar

- Filtrar/ordenar a aba Dados diretamente, ou copiar como intervalo para
  uma tabela dinâmica.
- Para usar num Power BI: importar a aba Dados como fonte de dados (Get
  Data → Excel) — já vem uma linha por remessa, sem células mescladas.
