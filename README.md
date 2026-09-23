# Automação de Planilhas — Operação Logística

## 1. Troca de Turno 2026 (versão enxuta)

Arquivo: [`planilhas/Troca_de_Turno_2026.xlsx`](planilhas/Troca_de_Turno_2026.xlsx) — sem macros, sem senha.

| Aba | Para quê |
| --- | --- |
| **LANÇAMENTOS** | 1 linha = 1 atividade/ocorrência. Colunas de cabeçalho azul são preenchidas pelo líder; as cinzas (ID, Status, Alerta, %, Dias em aberto) são automáticas. |
| **PASSAGEM** | O que fica para o próximo turno, críticos em aberto e um resumo pronto para copiar no WhatsApp. |
| **PENDÊNCIAS** | Tudo em aberto, críticos primeiro e depois os mais antigos. Clique no ID para ir à linha. |
| **PAINEL** | Planejado x executado, em aberto, faturamento e embarques — por turno e por programação. |
| **LISTAS** | Itens das listas suspensas, ajustes e "como usar". |

**Campos de preenchimento:** Data, Turno, Responsável, Programação, Rota/Cliente, Medida, Planejado,
Executado, Faturamento, Embarque, Transportadora, Pendência, Resp. pendência, Prazo, Concluído em, Observação.

**Status automático:** CRÍTICO (prazo vencido ou embarque pendente há mais de 1 dia) · ATENÇÃO
(pendência aberta, embarcado sem faturar, vence hoje) · PENDENTE (sem execução) · EM ANDAMENTO ·
CONCLUÍDO · HISTÓRICO (item antigo em aberto, anterior à data de corte em LISTAS).

### Migração do CSV original

`dados/TROCA_DE_TURNO_2026_original.csv` → `scripts/migrar_csv.py` → `scripts/gerar_planilha.py`.
826 registros importados; correções listadas em `dados/correcoes_migracao.csv` (datas com dia/mês
invertidos, números em kg convertidos para toneladas, cadastros padronizados, linhas vazias removidas).
Itens em aberto anteriores a 20/09/2026 ficam como HISTÓRICO (não tinham responsável nem prazo).

Para regenerar: `pip install openpyxl && python scripts/migrar_csv.py && python scripts/gerar_planilha.py`.

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
