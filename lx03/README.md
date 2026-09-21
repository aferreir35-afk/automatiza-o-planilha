# Dashboard de Estoque LX03 — Barry Callebaut

Automação em Python + Excel para transformar o export da transação SAP
**LX03** (estoque por posição no depósito) num painel de acompanhamento
diário, com os pontos que precisam de atenção já destacados — pronto para
uso operacional e para apresentação à diretoria.

## Estrutura da pasta

```
lx03/
├── scripts/
│   ├── atualizar_dashboard.py   # script principal: gera e atualiza o dashboard
│   ├── rodar_diario.py          # pega o arquivo mais recente de entrada/ e chama o principal
│   ├── agendar_tarefa.py        # configura o agendamento diário no SO (rodar 1x)
│   ├── tendencia.py             # histórico/tendência (SQLite + aba Tendencia)
│   ├── notificacoes.py          # alertas por e-mail/Teams
│   ├── conector_sap.py          # conexão direta ao SAP (OData/RFC), alternativa ao arquivo manual
│   ├── publicar_nuvem.py        # publicação no Google Sheets / Power BI
│   ├── gerar_pagina_web.py      # gera a versão web (HTML) do painel
│   ├── templates/painel.html    # modelo da página web (design/estrutura)
│   └── config.py                # carregador de config/config.ini
├── config/
│   ├── config.exemplo.ini       # modelo com todas as opções (comentado)
│   └── config.ini               # SEU arquivo real com credenciais (não vai pro Git)
├── entrada/                     # coloque aqui o export do dia (LX03 → Excel)
├── dashboard/
│   ├── Dashboard_Estoque_LX03.xlsx   # o arquivo "vivo" — sempre a última carga
│   └── Painel_LX03.html         # versão web do painel — abre só clicando, sem Excel
├── historico/                   # cópia do export + do dashboard de cada rodada, e tendencia.db
├── relatorios_pdf/              # PDF gerado a cada atualização (Painel + Pontos de Atenção)
├── GUIA_POWER_BI.md             # passo a passo para publicar no Power BI sem código
└── requirements.txt
```

## Como atualizar todos os dias

1. No SAP, rode a transação **LX03** e exporte o resultado para Excel.
2. Salve o arquivo em `lx03/entrada/` (pode sobrescrever o de ontem).
3. Rode:
   ```bash
   python lx03/scripts/atualizar_dashboard.py lx03/entrada/NOME_DO_ARQUIVO.xlsx
   ```
4. Pronto. O script:
   - confere se o arquivo tem as colunas esperadas da LX03 (avisa e para se
     não tiver — não gera um dashboard com dados errados);
   - recalcula os alertas de cada posição (vencido, bloqueado, em
     quarentena, parado sem giro, etc.);
   - reescreve `lx03/dashboard/Dashboard_Estoque_LX03.xlsx` com fórmulas do
     Excel (nada é "achatado" em números fixos — o arquivo recalcula
     sozinho se alguém mudar um filtro ou um parâmetro);
   - guarda uma cópia do export recebido e do dashboard gerado em
     `lx03/historico/` (rastreabilidade — "o que a diretoria viu no dia X
     veio de qual arquivo do SAP?");
   - gera um PDF do Painel + Pontos de Atenção em `lx03/relatorios_pdf/`;
   - gera uma versão web em `lx03/dashboard/Painel_LX03.html` — mesmos
     filtros, indicadores, gráficos e Top 10 do Painel do Excel, mas abre
     direto no navegador (duplo clique), sem precisar do Excel instalado;
   - imprime um resumo em texto no terminal com os principais alertas, para
     quem só precisa da visão rápida sem abrir o Excel.

Esse é o único comando do dia a dia — é o mesmo script tanto para a
primeira geração do dashboard quanto para todas as atualizações seguintes.

### Opções da linha de comando

```bash
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-pdf
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-historico
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-tendencia
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-notificacoes
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-nuvem
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-html
python lx03/scripts/atualizar_dashboard.py --usar-sap   # busca direto do SAP, sem arquivo (ver abaixo)
```

## Versão web (HTML)

`lx03/dashboard/Painel_LX03.html` é gerado automaticamente a cada atualização —
duplo clique e abre em qualquer navegador. É uma fotografia estática da
última carga (os filtros recalculam na hora, no navegador, mas os dados só
atualizam na próxima rodada do script) — para compartilhar, é só enviar o
arquivo `.html` (não precisa de servidor).

Os três filtros do topo (**Tipo de Depósito**, **Centro**, **Alerta**) são de
**seleção múltipla** — clique no campo pra abrir a lista de caixinhas,
marque quantos itens quiser (ou "Selecionar todos"/"Limpar"). Tudo na
página se ajusta na hora conforme o que estiver marcado:
- **Indicadores gerais** e as tabelas de **Top 10** — sempre pelos 3 filtros.
- **Pontos que carecem de monitoramento** — mostra só as categorias
  marcadas no filtro Alerta (Tipo/Centro também se aplicam); se uma
  categoria como "PARADO/SEM GIRO" for desmarcada, a linha dela some da
  lista e a tabela de Top 10 correspondente avisa que foi excluída pelo
  filtro.
- **Gráfico "Estoque por Tipo de Depósito"** — se você marcar tipos
  específicos, o gráfico mostra só esses (em vez do Top 12 padrão) —
  funciona como uma forma visual de comparar só os tipos que importam.
- **Gráfico "Posições por Alerta"** — mostra só as categorias marcadas.

### Abas — Visão Geral + uma por categoria de alerta

Logo abaixo dos filtros tem uma barra de abas: **Visão Geral** (o painel
completo de sempre — KPIs, Pontos de Atenção, gráficos, Top 10, Tendência)
e mais uma aba **para cada uma das 9 categorias de alerta** (Vencido,
Qualidade, Bloqueado, etc.). Cada aba de alerta traz:
- Uma descrição da categoria e a ação recomendada;
- KPIs só daquela categoria (posições, peso total em kg, total em UN, peso
  médio por posição);
- A lista **completa** de posições daquela categoria (não só o Top 10),
  com **Material, Lote, Posição, Tipo, Centro, UM, Quantidade, Data de
  Vencimento, Dias sem giro e Dias até vencer** — clique no cabeçalho de
  qualquer coluna para ordenar por ela (clique de novo para inverter), com
  paginação de 50 em 50 linhas.

As abas de alerta respeitam os filtros de Tipo de Depósito e Centro (para
ver só uma categoria com tudo, use a aba dela em vez do filtro de Alerta,
que é mais útil para combinar categorias na Visão Geral).

### Botões do topo

- **📤 Enviar novo relatório** — veja a seção "Atualizar sem rodar o
  script", abaixo.
- **⬇ Exportar dados filtrados (CSV)** — baixa as posições que estão sendo
  mostradas no momento (respeitando os filtros ativos) num `.csv` que abre
  direto no Excel, com acentuação, separador e casas decimais no padrão
  brasileiro (inclui a Data de Vencimento de cada posição).
- **🖶 Imprimir / Salvar PDF** — abre o diálogo de impressão do navegador
  (filtros e botões somem automaticamente do resultado); escolher "Salvar
  como PDF" no destino gera um PDF da página como está, com o filtro atual
  aplicado.

### Atualizar sem rodar o script

O botão **"📤 Enviar novo relatório"** deixa escolher um novo export da
LX03 (`.xlsx`) e processa tudo **direto no navegador** — sem instalar
Python, sem rodar comando nenhum. Em poucos segundos a página inteira
(KPIs, todas as abas, gráficos, Top 10) atualiza com os dados novos,
recalculando os mesmos alertas (vencido, bloqueado, parado/sem giro etc.)
com a data de hoje.

**O que esse botão NÃO faz** (para isso, continue rodando o script Python
normalmente):
- Não atualiza o arquivo Excel nem o `Painel_LX03.html` salvo no
  computador — o que você vê é só nesta aba do navegador, e some se
  fechar ou recarregar a página.
- Não grava nada em `lx03/historico/` nem na aba/gráfico de **Tendência**
  (ela continua mostrando o histórico salvo pelo script Python).
- Não envia e-mail/Teams, não publica no Google Sheets/Power BI.

Pense nele como uma forma rápida de espiar um export novo (ex.: conferir
antes de rodar oficialmente, ou usar num computador sem Python instalado)
— para o registro "oficial" do dia, sempre rode o `atualizar_dashboard.py`.

Para gerar a página avulsa, sem rodar a atualização inteira:

```bash
python lx03/scripts/gerar_pagina_web.py lx03/dashboard/Dashboard_Estoque_LX03.xlsx
```

## As 5 automações adicionais

Tudo abaixo é **opcional** e desligado por padrão — sem configurar nada, o
dashboard funciona exatamente como antes (arquivo manual → Excel → PDF).
Cada item liga independente dos outros, editando
`lx03/config/config.ini` (copie de `config.exemplo.ini`, que tem todos os
campos comentados).

### 1. Agendamento automático (não precisar lembrar de rodar todo dia)

Rode uma vez:

```bash
python lx03/scripts/agendar_tarefa.py --hora 07:00
```

- **Windows**: cria uma tarefa no Agendador de Tarefas do Windows.
- **Linux/Mac**: adiciona uma linha no `crontab` do usuário (se o comando
  `crontab` não existir no sistema, o script imprime a linha exata para
  você adicionar manualmente ao seu agendador).

O agendamento sempre chama `rodar_diario.py`, que por sua vez pega
**sozinho o arquivo mais recente de `lx03/entrada/`** — então o único
passo manual que sobra é salvar o export do dia nessa pasta antes do
horário agendado (ou configurar `--usar-sap`/`conector_sap.py`, item 4, para
eliminar até esse passo). Para remover o agendamento:
`python lx03/scripts/agendar_tarefa.py --remover`.

### 2. Alertas por e-mail e Microsoft Teams

Em `config.ini`, seções `[email]` e/ou `[teams]`:

```ini
[email]
ativo = true
servidor_smtp = smtp.office365.com
porta = 587
usuario = seu-usuario@barry-callebaut.com
senha = sua-senha-ou-senha-de-app
remetente = seu-usuario@barry-callebaut.com
destinatarios = gerente1@empresa.com, gerente2@empresa.com

[teams]
ativo = true
webhook_url = https://.../webhook-do-canal
```

Por padrão (`[alertas] somente_quando_piora = true`), só manda mensagem
quando algum ponto de atenção **piorou** frente à última carga (comparando
com o histórico de tendência, item 3) — e sempre manda se houver qualquer
posição **vencida**, mesmo sem piora (`sempre_alertar_se_vencido = true`).
Ambos os comportamentos são ajustáveis no `config.ini`.

Webhook do Teams: no canal desejado, `Conectores > Webhook de Entrada`
(ou, em tenants já migrados, um fluxo do Power Automate acionado por
webhook) — copie a URL para `webhook_url`.

### 3. Histórico de tendência (evolução dia a dia)

Sempre ativo por padrão (desligue com `--sem-tendencia`). A cada rodada, o
script grava um retrato do dia em `lx03/historico/tendencia.db` (SQLite) e
monta a aba **Tendencia** no dashboard, com uma tabela e um gráfico de
linha mostrando a evolução de Vencido/Qualidade/Bloqueado/Parado-sem-giro
ao longo do tempo. Roda o script mais de uma vez no mesmo dia? Ele
substitui o retrato daquele dia — a tendência é por dia, não por execução.
Começa com 1 ponto e cresce a cada atualização diária.

### 4. Conexão direta ao SAP (sem exportar manualmente)

Duas opções em `config.ini`, para usar com `--usar-sap` em vez de passar um
arquivo:

- **`[sap_odata]`** (recomendado): um serviço OData do SAP Gateway/Fiori
  que devolve as mesmas colunas da LX03. Peça ao time de Basis a URL do
  serviço e um usuário de leitura. Se os nomes dos campos do serviço forem
  diferentes dos nossos, existe um de-para em
  `[sap_odata_mapeamento]` (ver comentários no `config.exemplo.ini`).
- **`[sap_rfc]`** (alternativa): via pacote `pyrfc` + SAP NetWeaver RFC SDK
  (instalado à parte — licenciamento da SAP). O `conector_sap.py` traz um
  ponto de partida via `RFC_READ_TABLE`; para reproduzir de verdade a
  lógica de seleção da LX03, o ideal é o time de Basis expor uma função Z
  própria com a mesma saída.

```bash
python lx03/scripts/atualizar_dashboard.py --usar-sap
```

*Testado neste projeto com um serviço OData simulado (mock) — a lógica de
leitura e conversão dos dados foi validada de ponta a ponta; a conexão com
o SAP real de vocês só pode ser validada com a URL e as credenciais reais,
que o time de Basis precisa fornecer.*

### 5. Publicar em Google Sheets / Power BI

- **Google Sheets**: crie uma conta de serviço no Google Cloud Console,
  baixe o JSON de credenciais, compartilhe a planilha de destino com o
  e-mail dessa conta de serviço, e preencha `[google_sheets]` no
  `config.ini` (caminho do JSON + ID da planilha). Publica a aba
  Pontos_de_Atencao (resumo pequeno — a base bruta de 10 mil+ linhas fica
  só no Excel/SAP, para não estourar cota de API nem ficar lento).
- **Power BI**: o caminho recomendado não precisa de nenhum código — veja
  **[`GUIA_POWER_BI.md`](GUIA_POWER_BI.md)** (Power BI Desktop lendo o
  Excel direto de uma pasta do OneDrive/SharePoint, com atualização
  agendada). Para quem quiser publicar em tempo real via API (push
  dataset), o caminho avançado com App Registration no Azure AD também
  está documentado lá e já implementado em `publicar_nuvem.py`.

Ambos são independentes; ative um, outro, ou os dois em `config.ini`.

## O que tem no dashboard

- **Painel** — filtros (Tipo de Depósito, Centro, Alerta) no topo; abaixo,
  indicadores gerais, a tabela de pontos de atenção, 4 gráficos e os Top 10
  de maior estoque parado/vencido. Tudo recalcula ao trocar os filtros.
- **Pontos_de_Atencao** — lista fixa (sem filtro, todo o armazém) com cada
  ponto que precisa de acompanhamento: o que significa, quanto pesa em
  kg/%, a severidade e a ação recomendada. É a aba pensada para levar
  pronta a uma reunião de diretoria.
- **Dados_SAP** — a base bruta importada do SAP, com filtro automático do
  Excel em cada coluna (para quem quiser investigar posição por posição) e
  as colunas de apoio que o Painel usa (Alerta, dias até vencer, dias sem
  giro).
- **Parametros** — células amarelas editáveis: os dois prazos de alerta
  ("próx. vencimento" e "parado/sem giro") e uma coluna para você preencher
  o nome de negócio de cada código de Tipo de Depósito (ex.: `FM1` =
  Produto Acabado 1) — isso é específico da configuração de vocês no SAP e
  eu não tinha como confirmar com certeza, então deixei em branco para
  preencherem.
- **Tendencia** — evolução diária dos pontos de atenção (ver item 3 das
  automações adicionais, abaixo). Cresce a cada atualização.
- **Leia-me** — as mesmas instruções acima, dentro do próprio arquivo.

## Os 9 pontos de atenção monitorados

| Categoria | O que significa |
|---|---|
| VENCIDO | Material já passou da validade |
| PRÓX. VENCIMENTO | Vence dentro do prazo configurado (padrão: 60 dias) |
| QUALIDADE | Em inspeção de qualidade (tipo de estoque `Q`) |
| BLOQUEADO | Bloqueado para uso (tipo de estoque `S`) |
| RESTRITO/DEVOLUÇÃO | Estoque restrito/devolução (tipo de estoque `R`) |
| QUARENTENA | Depósito marcado como `QUAR`/`QRES` |
| DESCARTE/DEVOLUÇÃO | Está fisicamente numa posição `DESCARTE` ou `DEVOLUÇÃO` |
| SEM LOTE | Falta o número de lote (risco de rastreabilidade) |
| PARADO/SEM GIRO | Sem movimento há muitos dias (padrão: 90 dias) |

Uma posição pode se encaixar em mais de um critério ao mesmo tempo — o
script usa a ordem acima como prioridade (ex.: um item vencido aparece como
"VENCIDO" mesmo que também esteja bloqueado), para não contar a mesma
posição em mais de uma categoria e inflar os números.

## Exportar em PDF / imprimir

Nas abas **Painel** ou **Pontos_de_Atencao**: `Arquivo > Imprimir` ou
`Exportar como PDF` no Excel — a área e o layout de impressão já vêm
configurados (paisagem, ajustado à largura da página). O script também gera
esse PDF automaticamente a cada atualização.

## Cores

Paleta em tons de vinho/marrom/dourado, inspirada na identidade visual do
chocolate/Barry Callebaut. **Os códigos hexadecimais oficiais da marca não
puderam ser confirmados** — o acesso ao site institucional e a bancos de
marca de terceiros ficou bloqueado neste ambiente de execução. Se o time de
marketing tiver o manual de marca, me envie os hex exatos: eles ficam em um
único lugar do script (`scripts/atualizar_dashboard.py`, bloco "Paleta de
cores", no topo do arquivo) — troca e roda de novo, sem precisar mexer no
resto.

## Limitações assumidas

- O nome de negócio de cada código de Tipo de Depósito (aba Parametros,
  coluna F) fica em branco na primeira geração — é uma informação de
  configuração do SAP de vocês que eu não tinha como confirmar.
- A contagem de "Materiais únicos" no Painel é calculada pelo script no
  momento da carga (não é uma fórmula "ao vivo") para manter o arquivo leve
  — ela já atualiza sozinha a cada nova rodada do script.
- O PDF automático é gerado com o LibreOffice; se ele não estiver
  disponível no ambiente onde o script rodar, o dashboard ainda é gerado
  normalmente — só o PDF não sai sozinho (o aviso aparece no terminal).
- As integrações de e-mail, Teams, SAP direto, Google Sheets e Power BI
  foram implementadas e testadas com servidores simulados (mock) — não há
  como testar de ponta a ponta contra o SMTP, o SAP, o Google ou o Azure
  AD reais de vocês sem as credenciais/endereços reais. Assim que
  preencherem `config.ini`, tudo deve funcionar; se algo não bater com o
  formato específico do serviço de vocês (nomes de campo do OData,
  schema do dataset do Power BI, etc.), me avisem que eu ajusto.
- `config.ini` nunca deve ser commitado no Git (já está no `.gitignore`) —
  é onde ficam as credenciais reais.
- O botão "Enviar novo relatório" carrega a biblioteca de leitura de Excel
  (SheetJS) de um CDN público — **precisa de internet** no navegador para
  funcionar (o resto da página funciona 100% offline, já que os dados da
  última atualização vêm embutidos no próprio arquivo `.html`).
- A cor de cada linha "Crítico"/"Atenção" nos pontos de atenção segue a
  mesma prioridade do Excel (ex.: um item vencido aparece só como
  "VENCIDO", mesmo que também esteja bloqueado) — os totais das duas
  versões (Excel e web) sempre batem porque usam exatamente a mesma regra.
