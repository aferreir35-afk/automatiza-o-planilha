# Automação de Planilha — Volume por Turno

Planilha de controle de volume de carga por turno/divisão, com o **Painel**
("dashboard") ampliado para monitorar variância entre Planejado x Realizado
usando apenas fórmulas e formatação condicional do próprio Excel, **mais um
Painel Executivo visual (estilo Power BI)** para apresentar à diretoria —
tudo **sem servidor, add-in, instalação ou integração externa**
(e-mail/Slack/WhatsApp).

- Planilha: [`planilhas/Volume_por_Turno_Setembro_2026.xlsx`](planilhas/Volume_por_Turno_Setembro_2026.xlsx)
- Painel Executivo (visual): [`dashboard/index.html`](dashboard/index.html)

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

## Painel Executivo (dashboard visual estilo Power BI)

Para apresentar os números à diretoria de forma visual — KPIs, gráficos e
alertas coloridos, em vez de uma tabela de Excel — use o arquivo
[`dashboard/index.html`](dashboard/index.html). **Não precisa saber Excel,
fórmula ou programação, nem instalar nada**: é um único arquivo HTML — dá
duplo clique e abre no navegador, é só preencher um formulário na tela.
Pode copiar esse único arquivo para o Desktop, enviar por e-mail/WhatsApp ou
guardar num pen drive; ele funciona sozinho, sem depender de mais nada.

**Rotina diária (leva menos de 1 minuto, sem Excel):**

1. Abra o arquivo `dashboard/index.html` **dando duplo clique nele** (ele
   abre no seu navegador — Chrome, Edge, etc.). Não precisa instalar nada.
2. No topo, preencha o formulário **"Lançamento do dia"**: Data, Turno,
   Divisão, Volume Planejado e Volume Realizado.
3. Clique em **"+ Adicionar lançamento"**.
4. Pronto — o painel inteiro (KPIs, gráficos, alertas, maiores desvios,
   resumo da semana, projeção de fechamento) é recalculado na hora, igual à
   aba Painel da planilha, só que em formato de apresentação.

Repita o passo 2–3 para cada turno/divisão lançado no dia. Os lançamentos já
adicionados ficam listados numa tabela logo abaixo do formulário, onde dá
para **editar** (ícone ✏️) ou **excluir** (ícone 🗑️) qualquer um deles a
qualquer momento — por exemplo, para completar o "Volume Realizado" mais
tarde, depois que o turno planejado de manhã já tiver sido lançado.

**Outros recursos do painel:**

- **Filtros de Mês, Divisão e Turno** no topo funcionam como os "filtros" do
  Power BI: escolha o mês ou clique para incluir/excluir uma Divisão/Turno,
  e todo o painel (KPIs, gráficos, tabelas) recalcula na hora.
- **"Exportar apresentação (PDF)"** abre a tela de impressão do navegador —
  escolha "Salvar como PDF" para gerar o arquivo pronto para enviar/
  apresentar.
- O painel **salva tudo automaticamente no seu navegador** a cada
  lançamento — pode fechar a aba e abrir de novo (ou reiniciar o
  computador) que os dados continuam lá, sem precisar salvar nada
  manualmente.
- **"Importar arquivo"** (opcional, no canto superior) é só para quem já tem
  uma planilha `.xlsx`/`.csv` pronta (por exemplo, a planilha Excel deste
  repositório) e quer trazer esses lançamentos de uma vez para o painel, em
  vez de digitar um por um. Pode importar quantas vezes quiser: lançamentos
  repetidos (mesma Data + Turno + Divisão) são atualizados, não duplicados.
- Os limites de alerta (`Meta mínima de Atingimento` e `Limite de Saldo
  Semanal`) são editáveis diretamente no painel.
- **Nenhum dado sai do seu computador**: todo o processamento acontece no
  navegador (não há upload para nenhum servidor/nuvem). Os dados ficam
  salvos só naquele navegador/computador — se for usar em outro computador,
  use o botão "Importar arquivo" com uma exportação da planilha para levar
  os dados junto.

## Limitações assumidas (mantendo a solução simples)

- O alerta semanal usa semana de calendário (segunda a domingo) com base na
  última data lançada, não uma janela contínua dos últimos 7 dias.
- A projeção de fechamento assume ritmo diário constante; não usa média
  móvel nem desvio-padrão.
- O ranking de maiores desvios não trata empates de forma especial (pode
  repetir uma divisão quando o Saldo é igual).
