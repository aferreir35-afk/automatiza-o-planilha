# Dashboard de Estoque LX03 — Barry Callebaut

Automação em Python + Excel para transformar o export da transação SAP
**LX03** (estoque por posição no depósito) num painel de acompanhamento
diário, com os pontos que precisam de atenção já destacados — pronto para
uso operacional e para apresentação à diretoria.

## Estrutura da pasta

```
lx03/
├── scripts/
│   └── atualizar_dashboard.py   # script único: gera e atualiza o dashboard
├── entrada/                     # coloque aqui o export do dia (LX03 → Excel)
├── dashboard/
│   └── Dashboard_Estoque_LX03.xlsx   # o arquivo "vivo" — sempre a última carga
├── historico/                   # cópia do export + do dashboard de cada rodada (auditoria)
└── relatorios_pdf/              # PDF gerado a cada atualização (Painel + Pontos de Atenção)
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
   - imprime um resumo em texto no terminal com os principais alertas, para
     quem só precisa da visão rápida sem abrir o Excel.

Esse é o único comando do dia a dia — é o mesmo script tanto para a
primeira geração do dashboard quanto para todas as atualizações seguintes.

### Rodar sem gerar PDF ou sem guardar histórico

```bash
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-pdf
python lx03/scripts/atualizar_dashboard.py lx03/entrada/arquivo.xlsx --sem-historico
```

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
