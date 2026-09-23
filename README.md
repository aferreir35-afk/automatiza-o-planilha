# Automação de Planilhas — Operação Logística

## 1. Troca de Turno 2026 — Sistema de Gestão Operacional (novo)

Arquivo: [`planilhas/Troca_de_Turno_2026_Gestao_Operacional.xlsx`](planilhas/Troca_de_Turno_2026_Gestao_Operacional.xlsx)

Reconstrução da planilha "TROCA DE TURNO 2026 (VISÃO)" como ferramenta de gestão:
**simples para quem lança, completa para quem gerencia**. Só usa fórmulas nativas do
Excel (sem suplementos nem serviços externos); as macros são opcionais.

### Abas

| Aba | Para quê |
| --- | --- |
| **DASHBOARD** | 12 KPIs, comparativo A/B/C/D, status, pendências por tipo, top 10 prioridades. 12 filtros (período, turno, responsável, programação, rota, status, tipo, impacto, faturamento, embarque). |
| **GESTÃO_DO_DIA** | Números do dia por turno, "o que tratar agora" e "o que fica para o próximo turno". |
| **PASSAGEM_DE_TURNO** | Críticos, atenção, concluídos, pendências deixadas, prazos do dia, responsáveis e **resumo automático** pronto para copiar. |
| **LANÇAMENTO** | Formulário de lançamento (grava na base via macro opcional). |
| **BASE_OPERACIONAL** | Tabela fato `tbBase`: 1 linha = 1 atividade/ocorrência. 30 campos do briefing + CANCELADO?, ORIGEM + colunas automáticas. |
| **CONTROLE_PENDÊNCIAS** | Tudo o que está em aberto, ordenado por crítico → vencido → prazo mais próximo → impacto. |
| **FATURAMENTO / EMBARQUES** | Programado x realizado, por programação/turno/transportadora, alertas das regras 1 e de embarque atrasado. |
| **PRODUTIVIDADE** | Por mês, semana, dia, turno (inclui por colaborador, se informado) e responsável. |
| **GESTÃO_SEMANAL / GESTÃO_MENSAL** | Consolidações com recorrência, causas, dia de maior desvio, turno de maior volume, evolução mensal. |
| **ANÁLISE_CAUSAS** | Pendências por tipo, turno, responsável, causa e impacto; tempo médio de resolução; SLA; recorrência; matriz Frequência x Impacto. |
| **PARETO** | Causas e tipos com %, % acumulado e classe A/B/C. |
| **HISTÓRICO** | Consulta filtrada (até 500 linhas) — nada é apagado. |
| **QUALIDADE_DADOS** | % de qualidade da base e 16 indicadores (sem responsável, sem prazo, duplicados, fora do padrão…). |
| **CADASTROS / CONFIGURAÇÕES** | Listas suspensas, SLA por tipo, sequência/horário dos turnos, metas e tolerâncias. |
| **LOG_ALTERAÇÕES** | Trilha de auditoria (migração + macro). |
| **DIM_CALENDÁRIO** | Dimensão de datas (`dCalendario`) para Power BI. |
| **COMO_USAR** | Guia para usuários. |

### Regras automáticas (resumo)

- **% execução** = executado ÷ planejado (vazio se planejado vazio/0; 100% se executado ≥ planejado).
- **Status**: CANCELADO → CONCLUÍDO (nada em aberto) → CRÍTICO (prazo vencido, impacto crítico,
  embarque atrasado, sistema com impacto alto, pendência de impacto alto afetando o próximo turno)
  → ATENÇÃO (pendência aberta, embarcado sem faturar, vence hoje/em N horas, SLA estourado, erro de
  cadastro) → PENDENTE (sem execução) → EM ANDAMENTO.
- **Semáforo** 🔴🟡🟢 na coluna ALERTA + cores condicionais (ninguém pinta célula).
- **Regras de validação 1–5** do briefing (embarcado sem faturar, saldo com prazo vencido, pendência
  sem responsável/prazo, conclusão com item aberto) + extras (datas incoerentes, divergência, sem tipo).
- **Prazo**: NO PRAZO / VENCE HOJE / VENCIDO / CONCLUÍDO / SEM PRAZO; **dias em aberto**; **tempo de
  tratativa** (h); **SLA** configurável por tipo.
- Parâmetros em **CONFIGURAÇÕES** (data de referência, meta, tolerâncias, SLA, escala dos turnos).

### Migração do CSV original

`dados/TROCA_DE_TURNO_2026_original.csv` → `scripts/migrar_csv.py` → `dados/base_migrada.json`
→ `scripts/gerar_planilha.py` → planilha. Principais correções (todas no LOG_ALTERAÇÕES):

- 18 datas com dia/mês invertidos (ex.: `08/01/2026` entre 01/08 e 02/08) e 1 ano com 2 dígitos;
- números em formatos mistos (`43127`, `28.054,00`, `24,039`) → **toneladas** (valores ≥ 100 estavam em kg);
- turno/programação/medida/transportadora padronizados (ex.: EPEEDYLOG/SPEEDY → SPEEDYLOG);
- 10 linhas vazias descartadas; o bloco de totais manuais do topo (META + ACUMULADO, EQUIPAMENTO
  PARADO, NF COM ERRO…) foi substituído pelos KPIs automáticos;
- pendências identificadas pelas palavras da observação (AGUARDANDO, FALTA, PENDENTE…); causa só
  quando o texto a indica. Responsável e prazo **não** foram inventados.
- Registros migrados anteriores à **data de corte** (padrão 20/09/2026) contam nos indicadores, mas não
  geram pendência aberta. Os migrados a partir do corte aparecem como pendências reais a tratar.

Para regenerar: `pip install openpyxl && python scripts/migrar_csv.py && python scripts/gerar_planilha.py`.

### Macros opcionais (`vba/`)

`modTrocaTurno.bas` (RegistrarLancamento, LimparFormulario, CopiarResumoTurno, log) e
`EstaPastaDeTrabalho.cls` (carimbo de DATA/HORA DA ATUALIZAÇÃO e LOG em toda edição da base).
Alt+F11 → Importar o `.bas`; colar o `.cls` em *EstaPasta_de_trabalho*; salvar como `.xlsm`.

### Power BI / Power Query / Power Automate

- **Fato**: `tbBase` (filtrar `ID` não vazio; colunas `aux_*` podem ser removidas).
- **Dimensões**: `dCalendario` (DATA), e as listas de CADASTROS — turno, responsável, programação→área,
  transportadora, tipo de pendência, causa, impacto (peso). Relacionamentos 1:N pelas colunas de mesmo nome.
- **Power Query**: `Excel.Workbook(File.Contents(...))` → tabela `tbBase`; atualização agendada no
  gateway/SharePoint.
- **Power Automate** (sugestões): arquivo no OneDrive/SharePoint + "Listar linhas presentes em uma
  tabela" (`tbBase`) filtrando `STATUS = CRÍTICO` ou `SITUAÇÃO DO PRAZO = VENCIDO` → e-mail/Teams ao
  `RESPONSÁVEL PELA TRATATIVA`; na troca de turno, enviar o texto do resumo da PASSAGEM_DE_TURNO.

### Limites conhecidos

- Capacidade de 4.000 registros (≈ 1 ano no ritmo atual); ajuste `NROWS` no gerador para ampliar.
- O ID depende da posição da linha: não ordenar nem apagar linhas da base (usar CANCELADO? = SIM).
- Listas ordenadas exibem até 300 (pendências), 500 (histórico), 200 (qualidade) e 50 (alertas) itens.
- Log de alterações e carimbo de data/hora exigem a macro (Excel não registra usuário sem VBA).

---

## 2. Volume por Turno — Setembro 2026

Planilha de controle de volume de carga por turno/divisão, com o **Painel**
("dashboard") ampliado para monitorar variância entre Planejado x Realizado
usando apenas fórmulas e formatação condicional do próprio Excel — **sem
servidor, add-in ou integração externa (e-mail/Slack/WhatsApp)**.

Arquivo: [`planilhas/Volume_por_Turno_Setembro_2026.xlsx`](planilhas/Volume_por_Turno_Setembro_2026.xlsx)

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
