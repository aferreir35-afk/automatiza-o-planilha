# Publicar o Dashboard LX03 no Power BI (sem escrever código)

Este é o caminho recomendado para a maioria dos casos — usa só o Power BI
Desktop, sem precisar de App Registration no Azure nem de nenhuma API.

## Passo a passo

1. Salve a pasta `lx03/dashboard/` (ou pelo menos o arquivo
   `Dashboard_Estoque_LX03.xlsx`) numa pasta do **OneDrive/SharePoint** da
   empresa, em vez de só no computador local — é isso que permite ao
   Power BI atualizar sozinho todos os dias, mesmo com o computador
   desligado.
2. Ajuste o script `atualizar_dashboard.py` (parâmetro `--saida`) ou o
   agendamento (`agendar_tarefa.py`) para salvar o dashboard direto nessa
   pasta do OneDrive/SharePoint.
3. No **Power BI Desktop**: `Obter Dados > Excel` e aponte para o arquivo
   nessa pasta. Selecione as abas **Pontos_de_Atencao**, **Tendencia** e a
   tabela de apoio da aba **Painel**.
4. Monte os visuais que quiser (cartões de KPI, gráfico de barras por
   categoria, linha de tendência) — os dados já vêm limpos e prontos.
5. Publique no **Power BI Service** (`Página Inicial > Publicar`).
6. No Power BI Service, configure a **atualização agendada** do dataset
   (`Configurações do dataset > Atualização agendada`) apontando para o
   mesmo arquivo do OneDrive/SharePoint — o Power BI então atualiza
   sozinho, no horário que você escolher, sem precisar de gateway nem de
   nenhum código.

Esse caminho cobre bem o caso de uso de diretoria: um relatório sempre
atualizado, acessível pelo navegador ou pelo app do Power BI no celular.

## Caminho avançado: push via API (tempo real, sem Power BI Desktop)

Se a necessidade for outra — por exemplo, atualizar o Power BI no instante
em que o script Python roda, sem depender do ciclo de atualização agendada
— existe a alternativa de "push dataset" pela API REST do Power BI, já
implementada em `scripts/publicar_nuvem.py` (`publicar_power_bi`). Ela
exige:

1. Um **App Registration no Azure AD** com permissão de API
   `Power BI Service` (tipo aplicativo, não delegada) e um "client secret".
2. Um **dataset do tipo "push"** já criado no workspace do Power BI, com
   uma tabela cujo nome e colunas batem com a aba Pontos_de_Atencao
   (`Nº`, `Ponto de Atenção`, `O que significa`, `Qtd. Posições`,
   `Peso (kg)`, `% do KG total`, `Severidade`, `Ação Recomendada`) — dá
   para criar isso rapidinho pela própria API do Power BI ou com um script
   auxiliar, se precisarem, é só pedir.
3. Preencher `tenant_id`, `client_id`, `client_secret`, `workspace_id` e
   `dataset_id` na seção `[power_bi]` do `config/config.ini`.

Com isso configurado, todo `python atualizar_dashboard.py ...` sem a flag
`--sem-nuvem` já envia os pontos de atenção atualizados para o Power BI
automaticamente.
