#!/usr/bin/env python3
"""
Atualizador do Dashboard de Estoque LX03 (Barry Callebaut)
============================================================

Uso diário:
    python atualizar_dashboard.py <arquivo_lx03_exportado_do_sap.xlsx>

O que o script faz:
    1. Lê o arquivo exportado do LX03 (transação SAP de estoque por posição).
    2. Confere se as colunas esperadas existem (procura pelo NOME da coluna,
       então funciona mesmo se o SAP exportar as colunas em outra ordem).
    3. Calcula, para cada posição, um "Alerta" (vencido, em quarentena,
       bloqueado, parado sem giro, etc.) usando os parâmetros configurados
       na aba Parametros do próprio dashboard.
    4. Grava/atualiza o arquivo do dashboard (Painel, Pontos de Atenção,
       Dados_SAP, Parametros, Leia-me) com fórmulas do Excel — os números
       recalculam sozinhos sempre que o arquivo é aberto ou os filtros do
       Painel são alterados.
    5. Guarda uma cópia do arquivo enviado e do dashboard gerado na pasta
       historico/ (auditoria/rastreabilidade).
    6. Gera um PDF do Painel + Pontos de Atenção em relatorios_pdf/.
    7. Imprime um resumo simples no terminal com os principais alertas.

Este mesmo script serve tanto para a primeira geração do dashboard quanto
para as atualizações diárias seguintes — é sempre o mesmo comando.
"""

import argparse
import datetime
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Caminhos padrão (pasta lx03/ do repositório)
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent
DASHBOARD_PATH = BASE_DIR / "dashboard" / "Dashboard_Estoque_LX03.xlsx"
HISTORICO_DIR = BASE_DIR / "historico"
PDF_DIR = BASE_DIR / "relatorios_pdf"
RECALC_SCRIPT = None  # localizado em tempo de execução, ver localizar_recalc()

# ---------------------------------------------------------------------------
# Paleta de cores (inspirada na identidade visual do chocolate / Barry
# Callebaut: vinho/marrom + dourado). Os códigos oficiais da marca não
# puderam ser confirmados neste ambiente (acesso ao site institucional e a
# bancos de marca bloqueado) — troque os valores abaixo pelos códigos hex
# exatos do manual de marca assim que a equipe de marketing os enviar; é o
# único lugar do script que precisa mudar.
# ---------------------------------------------------------------------------
COR_VINHO = "6E1423"       # cor primária (títulos, cabeçalhos)
COR_VINHO_ESCURO = "4A0D18"
COR_MARROM = "3B2413"      # marrom chocolate (textos de destaque)
COR_DOURADO = "C8A24A"     # dourado/cacau (acentos, linhas de gráfico)
COR_CREME = "F7F1E8"       # fundo claro dos cartões de KPI
COR_BRANCO = "FFFFFF"
COR_CINZA_TEXTO = "5A5A5A"

COR_OK = "2E7D32"
COR_ALERTA_FUNDO_VERDE = "C6EFCE"
COR_ALERTA_FUNDO_AMARELO = "FFEB9C"
COR_ALERTA_FUNDO_VERMELHO = "FFC7CE"
COR_ALERTA_TEXTO_VERMELHO = "9C0006"
COR_ALERTA_TEXTO_AMARELO = "9C6500"
COR_ALERTA_TEXTO_VERDE = "006100"

FONTE = "Arial"

F_TITULO = Font(name=FONTE, bold=True, size=16, color=COR_VINHO)
F_SUBTITULO = Font(name=FONTE, italic=True, size=9, color=COR_CINZA_TEXTO)
F_SECAO = Font(name=FONTE, bold=True, size=11, color=COR_BRANCO)
F_LABEL = Font(name=FONTE, bold=True, size=9, color=COR_MARROM)
F_KPI = Font(name=FONTE, bold=True, size=18, color=COR_VINHO)
F_KPI_SUB = Font(name=FONTE, size=8, color=COR_CINZA_TEXTO)
F_HDR = Font(name=FONTE, bold=True, color=COR_BRANCO)
F_TXT = Font(name=FONTE, size=10, color="000000")
F_TXT_B = Font(name=FONTE, size=10, bold=True, color="000000")
F_NOTA = Font(name=FONTE, italic=True, size=9, color=COR_CINZA_TEXTO)

FILL_SECAO = PatternFill("solid", fgColor=COR_VINHO)
FILL_HDR = PatternFill("solid", fgColor=COR_VINHO)
FILL_CARD = PatternFill("solid", fgColor=COR_CREME)
FILL_INPUT = PatternFill("solid", fgColor="FFF6D9")
FILL_VERMELHO = PatternFill("solid", fgColor=COR_ALERTA_FUNDO_VERMELHO)
FILL_AMARELO = PatternFill("solid", fgColor=COR_ALERTA_FUNDO_AMARELO)
FILL_VERDE = PatternFill("solid", fgColor=COR_ALERTA_FUNDO_VERDE)

THIN = Side(style="thin", color="D9D2C7")
BORDA = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---------------------------------------------------------------------------
# Colunas esperadas no export do LX03 (nome exatamente como o SAP exporta).
# O script localiza cada uma pelo nome no cabeçalho, então a ordem das
# colunas no arquivo de origem pode mudar sem quebrar o processo.
# ---------------------------------------------------------------------------
COLUNAS_ESPERADAS = [
    "Tipo de depósito",
    "Posição no depósito",
    "Material",
    "Lote",
    "Nº de quantos",
    "Tipo de estoque",
    "Unidade de depósito",
    "Estoque disponível",
    "Último movimento",
    "Estoque total",
    "UM básica",
    "Duração",
    "Inventário ativo",
    "Depósito",
    "Data do vencimento",
    "Centro",
]

MATERIAL_VAZIO = "<< vazio >>"

DIAS_PROX_VENC_PADRAO = 60
DIAS_PARADO_PADRAO = 90

ALERTAS_ORDEM = [
    "VENCIDO",
    "PRÓX. VENCIMENTO",
    "QUALIDADE",
    "BLOQUEADO",
    "RESTRITO/DEVOLUÇÃO",
    "QUARENTENA",
    "DESCARTE/DEVOLUÇÃO",
    "SEM LOTE",
    "PARADO/SEM GIRO",
    "OK",
    "VAZIA",
]


def localizar_recalc():
    """Acha o scripts/recalc.py da skill de xlsx instalada nesta máquina."""
    candidatos = list(Path("/root/.claude/skills").glob("synced/*/xlsx/scripts/recalc.py"))
    if candidatos:
        return candidatos[0]
    return None


def ler_lx03(caminho_entrada: Path):
    """Lê o arquivo bruto do LX03 e devolve (linhas, indice_colunas)."""
    wb = openpyxl.load_workbook(caminho_entrada, data_only=True)
    ws = wb[wb.sheetnames[0]]
    cabecalho = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    indice = {nome: i + 1 for i, nome in enumerate(cabecalho) if nome}

    faltando = [c for c in COLUNAS_ESPERADAS if c not in indice]
    if faltando:
        raise SystemExit(
            "ERRO: o arquivo enviado não parece ser um export do LX03 no "
            "layout esperado. Colunas não encontradas: " + ", ".join(faltando) +
            "\nColunas encontradas no arquivo: " + ", ".join(str(c) for c in cabecalho if c)
        )

    linhas = []
    for r in range(2, ws.max_row + 1):
        valores = {nome: ws.cell(row=r, column=indice[nome]).value for nome in COLUNAS_ESPERADAS}
        # pula linhas totalmente vazias (rodapé do export, se houver)
        if all(v in (None, "") for v in valores.values()):
            continue
        linhas.append(valores)
    return linhas


def distintos_ordenados(linhas, campo):
    vistos = {}
    for linha in linhas:
        v = linha.get(campo)
        if v in (None, ""):
            continue
        vistos[str(v)] = True
    return sorted(vistos.keys())


def calcular_alerta(linha, dias_prox_venc, dias_parado, hoje):
    """Reproduz em Python a mesma regra de negócio da fórmula da coluna
    Alerta em Dados_SAP — usada só para o resumo impresso no terminal."""
    material = linha.get("Material")
    if material == MATERIAL_VAZIO:
        return "VAZIA"
    venc = linha.get("Data do vencimento")
    if isinstance(venc, datetime.datetime):
        dias_ate = (venc.date() - hoje).days
        if dias_ate < 0:
            return "VENCIDO"
        if dias_ate <= dias_prox_venc:
            return "PRÓX. VENCIMENTO"
    tipo_estoque = linha.get("Tipo de estoque")
    if tipo_estoque == "Q":
        return "QUALIDADE"
    if tipo_estoque == "S":
        return "BLOQUEADO"
    if tipo_estoque == "R":
        return "RESTRITO/DEVOLUÇÃO"
    if linha.get("Depósito") in ("QUAR", "QRES"):
        return "QUARENTENA"
    if linha.get("Posição no depósito") in ("DESCARTE", "DEVOLUÇÃO"):
        return "DESCARTE/DEVOLUÇÃO"
    if not linha.get("Lote") and linha.get("UM básica"):
        return "SEM LOTE"
    try:
        duracao = float(linha.get("Duração"))
    except (TypeError, ValueError):
        duracao = None
    if duracao is not None and duracao >= dias_parado:
        return "PARADO/SEM GIRO"
    return "OK"


def secao(ws, row, col_ini, col_fim, texto):
    ws.cell(row=row, column=col_ini, value=texto).font = F_SECAO
    for c in range(col_ini, col_fim + 1):
        ws.cell(row=row, column=c).fill = FILL_SECAO
    ws.merge_cells(start_row=row, start_column=col_ini, end_row=row, end_column=col_fim)
    ws.row_dimensions[row].height = 20


def construir_parametros(wb, linhas, nome_arquivo_origem, ultima_atualizacao, last_row_dados):
    ws = wb.create_sheet("Parametros")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 3
    ws.column_dimensions["D"].width = 3
    for col in ("E", "F", "G", "H", "I"):
        ws.column_dimensions[col].width = 20

    ws["A1"] = "Parâmetros do Dashboard — Estoque LX03"
    ws["A1"].font = F_TITULO
    ws.merge_cells("A1:I1")
    ws["A2"] = "Aba de configuração. As células amarelas podem ser editadas; o resto é gerado automaticamente pelo script a cada atualização."
    ws["A2"].font = F_SUBTITULO
    ws.merge_cells("A2:I2")

    secao(ws, 4, 1, 2, "DADOS DA ÚLTIMA CARGA")
    ws["A5"] = "Última atualização (data/hora)"
    ws["A5"].font = F_LABEL
    ws["B5"] = ultima_atualizacao
    ws["B5"].font = F_TXT
    ws["A6"] = "Arquivo de origem (SAP)"
    ws["A6"].font = F_LABEL
    ws["B6"] = nome_arquivo_origem
    ws["B6"].font = F_TXT
    ws["A7"] = "Posições carregadas"
    ws["A7"].font = F_LABEL
    c = ws["B7"]
    c.value = f"=COUNTA(Dados_SAP!$A$2:$A${last_row_dados})"
    c.font = F_TXT

    secao(ws, 9, 1, 2, "PARÂMETROS DE ALERTA (edite se necessário)")
    ws["A10"] = "Dias p/ alertar 'Próx. Vencimento'"
    ws["A10"].font = F_LABEL
    ws["B10"] = DIAS_PROX_VENC_PADRAO
    ws["B10"].font = Font(name=FONTE, color="0000FF")
    ws["B10"].fill = FILL_INPUT
    ws["B10"].comment = Comment(
        "Materiais cuja data de vencimento está a esta quantidade de dias (ou "
        "menos) da data de hoje entram no alerta 'PRÓX. VENCIMENTO'. Valor "
        "definido pelo usuário — ajuste conforme a política de qualidade.",
        "Dashboard LX03",
    )
    ws["A11"] = "Dias sem movimento p/ 'Parado/Sem Giro'"
    ws["A11"].font = F_LABEL
    ws["B11"] = DIAS_PARADO_PADRAO
    ws["B11"].font = Font(name=FONTE, color="0000FF")
    ws["B11"].fill = FILL_INPUT
    ws["B11"].comment = Comment(
        "Posições cujo campo 'Duração' (dias desde o último movimento, vindo "
        "do próprio SAP) é maior ou igual a este número entram no alerta "
        "'PARADO/SEM GIRO'. Valor definido pelo usuário.",
        "Dashboard LX03",
    )

    secao(ws, 12, 5, 9, "LISTAS PARA OS FILTROS DO PAINEL (geradas automaticamente a cada atualização)")

    ws["E13"] = "Código"
    ws["F13"] = "Descrição (preencher)"
    for cell in (ws["E13"], ws["F13"]):
        cell.font = F_HDR
        cell.fill = FILL_HDR
    tipos = distintos_ordenados(linhas, "Tipo de depósito")
    ws["E14"] = "(Todos)"
    ws["E14"].font = F_TXT_B
    for i, t in enumerate(tipos):
        ws.cell(row=15 + i, column=5, value=t).font = F_TXT
        ws.cell(row=15 + i, column=6, value="").font = F_TXT
        ws.cell(row=15 + i, column=6).fill = FILL_INPUT
    linha_fim_tipo = 14 + len(tipos)

    ws["H13"] = "Centro"
    ws["H13"].font = F_HDR
    ws["H13"].fill = FILL_HDR
    centros = [c for c in distintos_ordenados(linhas, "Centro")]
    ws["H14"] = "(Todos)"
    ws["H14"].font = F_TXT_B
    for i, cval in enumerate(centros):
        ws.cell(row=15 + i, column=8, value=cval).font = F_TXT
    linha_fim_centro = 14 + len(centros)

    col_alerta = 9  # I, ao lado do bloco de Centro para não colidir com a lista de tipos (mais longa)
    ws.cell(row=13, column=col_alerta, value="Alerta").font = F_HDR
    ws.cell(row=13, column=col_alerta).fill = FILL_HDR
    ws.cell(row=14, column=col_alerta, value="(Todos)").font = F_TXT_B
    for i, a in enumerate(ALERTAS_ORDEM):
        ws.cell(row=15 + i, column=col_alerta, value=a).font = F_TXT
    linha_fim_alerta = 14 + len(ALERTAS_ORDEM)

    linha_nota = max(linha_fim_tipo, linha_fim_centro, linha_fim_alerta) + 1
    ws.cell(row=linha_nota, column=5,
            value="↑ preencha a coluna F com o nome de negócio de cada código de Tipo de Depósito "
                  "(ex.: FM1 = Produto Acabado 1) — aparece só aqui, é sua referência.").font = F_NOTA
    ws.merge_cells(start_row=linha_nota, start_column=5, end_row=linha_nota, end_column=6)

    refs = {
        "dias_venc": "Parametros!$B$10",
        "dias_parado": "Parametros!$B$11",
        "lista_tipo": f"Parametros!$E$14:$E${linha_fim_tipo}",
        "lista_centro": f"Parametros!$H$14:$H${linha_fim_centro}",
        "lista_alerta": f"Parametros!$I$14:$I${linha_fim_alerta}",
        "tipos_para_grafico": tipos,
    }
    return ws, refs


FILTRO_TIPO = "Painel!$B$5"
FILTRO_CENTRO = "Painel!$E$5"
FILTRO_ALERTA = "Painel!$H$5"

COLUNAS_DADOS_SAP = COLUNAS_ESPERADAS + [
    "Duração (dias)",
    "Dias até vencer",
    "Alerta",
    "_chave_rank_parado",
    "_chave_rank_vencido",
]


def construir_dados_sap(wb, linhas, refs):
    ws = wb.create_sheet("Dados_SAP")
    n = len(linhas)
    last_row = n + 1

    for c, nome in enumerate(COLUNAS_DADOS_SAP, 1):
        cell = ws.cell(row=1, column=c, value=nome)
        cell.font = F_HDR
        cell.fill = FILL_HDR
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")

    for i, linha in enumerate(linhas):
        r = i + 2
        for c, nome in enumerate(COLUNAS_ESPERADAS, 1):
            valor = linha[nome]
            cell = ws.cell(row=r, column=c, value=valor)
            cell.font = F_TXT
            if nome in ("Último movimento", "Data do vencimento") and isinstance(valor, datetime.datetime):
                cell.number_format = "DD/MM/YYYY"
            if nome in ("Estoque disponível", "Estoque total"):
                cell.number_format = "#,##0.000"

        ws.cell(row=r, column=17, value=f'=IFERROR(VALUE(L{r}),"")').font = F_TXT
        ws.cell(row=r, column=18, value=f'=IF(O{r}="","",O{r}-TODAY())').font = F_TXT
        formula_alerta = (
            f'=IF(C{r}="{MATERIAL_VAZIO}","VAZIA",'
            f'IF(AND(O{r}<>"",R{r}<0),"VENCIDO",'
            f'IF(AND(O{r}<>"",R{r}<={refs["dias_venc"]}),"PRÓX. VENCIMENTO",'
            f'IF(F{r}="Q","QUALIDADE",'
            f'IF(F{r}="S","BLOQUEADO",'
            f'IF(F{r}="R","RESTRITO/DEVOLUÇÃO",'
            f'IF(OR(N{r}="QUAR",N{r}="QRES"),"QUARENTENA",'
            f'IF(OR(B{r}="DESCARTE",B{r}="DEVOLUÇÃO"),"DESCARTE/DEVOLUÇÃO",'
            f'IF(AND(D{r}="",K{r}<>""),"SEM LOTE",'
            f'IF(AND(Q{r}<>"",Q{r}>={refs["dias_parado"]}),"PARADO/SEM GIRO",'
            f'"OK")))))))))'
        )
        ws.cell(row=r, column=19, value=formula_alerta).font = F_TXT

        cond_filtro = (
            f'OR({FILTRO_TIPO}="(Todos)",A{r}={FILTRO_TIPO})*'
            f'OR({FILTRO_CENTRO}="(Todos)",P{r}={FILTRO_CENTRO})'
        )
        ws.cell(
            row=r, column=20,
            value=(f'=IF(AND($S{r}="PARADO/SEM GIRO",$K{r}="KG",{cond_filtro}),'
                   f'$J{r}+ROW()/10000000,"")'),
        ).font = F_TXT
        ws.cell(
            row=r, column=21,
            value=(f'=IF(AND($S{r}="VENCIDO",$K{r}="KG",{cond_filtro}),'
                   f'$J{r}+ROW()/10000000,"")'),
        ).font = F_TXT

    larguras = [10, 14, 20, 14, 9, 9, 10, 13, 13, 13, 8, 9, 9, 9, 13, 8, 10, 9, 18, 11, 11]
    for c, w in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"

    tabela = Table(displayName="TblLX03", ref=f"A1:U{last_row}")
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False,
        showLastColumn=False, showColumnStripes=False,
    )
    ws.add_table(tabela)

    # colunas de apoio (T, U), usadas só pelas fórmulas do Painel — ficam
    # visíveis (o LibreOffice não preserva a coluna oculta quando ela é a
    # última da planilha), mas o nome com "_" já indica que são internas.
    comentario_apoio = Comment(
        "Coluna de apoio para os rankings 'Top 10' do Painel — não precisa mexer.",
        "Dashboard LX03",
    )
    ws["T1"].comment = comentario_apoio
    ws["U1"].comment = Comment(
        "Coluna de apoio para os rankings 'Top 10' do Painel — não precisa mexer.",
        "Dashboard LX03",
    )

    return ws, last_row


ALERTAS_RISCO = [a for a in ALERTAS_ORDEM if a not in ("OK", "VAZIA")]

DESCRICAO_ALERTA = {
    "VENCIDO": "Material já venceu",
    "PRÓX. VENCIMENTO": "Vence em breve",
    "QUALIDADE": "Em inspeção de qualidade",
    "BLOQUEADO": "Bloqueado para uso",
    "RESTRITO/DEVOLUÇÃO": "Estoque restrito/devolução",
    "QUARENTENA": "Em quarentena",
    "DESCARTE/DEVOLUÇÃO": "Em posição de descarte/devolução",
    "SEM LOTE": "Sem lote (falta de rastreabilidade)",
    "PARADO/SEM GIRO": "Sem movimento há muito tempo",
    "OK": "Sem alerta",
    "VAZIA": "Posição vazia",
}


def rng(col_letra, last_row):
    return f"Dados_SAP!${col_letra}$2:${col_letra}${last_row}"


def sumproduct_filtros(condicoes_extra, last_row, incluir_alerta=True, soma_col=None):
    """Monta um SUMPRODUCT que respeita os 3 filtros do Painel.

    condicoes_extra: lista de strings, cada uma um fator array (ex.: '(Dados_SAP!$K$2:$K$100="KG")').
    soma_col: se definido, multiplica pela coluna (para somar peso); senão conta linhas.
    """
    fatores = list(condicoes_extra)
    fatores.append(f'(({FILTRO_TIPO}="(Todos)")+({rng("A", last_row)}={FILTRO_TIPO})>0)')
    fatores.append(f'(({FILTRO_CENTRO}="(Todos)")+({rng("P", last_row)}={FILTRO_CENTRO})>0)')
    if incluir_alerta:
        fatores.append(f'(({FILTRO_ALERTA}="(Todos)")+({rng("S", last_row)}={FILTRO_ALERTA})>0)')
    formula = "=SUMPRODUCT(" + "*".join(fatores)
    if soma_col:
        formula += f",{rng(soma_col, last_row)}"
    formula += ")"
    return formula


def kpi_card(ws, row, col, titulo, formula, numfmt="#,##0"):
    lbl = ws.cell(row=row, column=col, value=titulo)
    lbl.font = Font(name=FONTE, size=8, bold=True, color=COR_BRANCO)
    lbl.fill = PatternFill("solid", fgColor=COR_VINHO)
    lbl.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    val = ws.cell(row=row + 1, column=col, value=formula)
    val.font = F_KPI
    val.number_format = numfmt
    val.fill = FILL_CARD
    val.alignment = Alignment(horizontal="center")
    val.border = BORDA
    ws.row_dimensions[row].height = 28


def construir_painel(wb, refs, last_row, linhas):
    ws = wb.create_sheet("Painel")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Painel de Estoque — LX03"
    ws["A1"].font = F_TITULO
    ws.merge_cells("A1:N1")
    ws["A2"] = (
        "Barry Callebaut · Estoque por posição no depósito (fonte: SAP LX03) — "
        "os cartões, gráficos e tabelas abaixo mudam sozinhos ao trocar os filtros ou ao rodar a atualização diária."
    )
    ws["A2"].font = F_SUBTITULO
    ws.merge_cells("A2:N2")
    ws["A3"] = '=CONCATENATE("Última atualização: ",Parametros!$B$5," · Arquivo: ",Parametros!$B$6)'
    ws["A3"].font = F_NOTA
    ws.merge_cells("A3:N3")

    secao(ws, 4, 1, 14, "FILTROS  (selecione para atualizar todo o painel)")
    for col_lbl, col_val, titulo in ((1, 2, "Tipo de Depósito:"), (4, 5, "Centro:"), (7, 8, "Alerta:")):
        lbl = ws.cell(row=5, column=col_lbl, value=titulo)
        lbl.font = F_LABEL
        lbl.alignment = Alignment(horizontal="right", vertical="center")
        val = ws.cell(row=5, column=col_val, value="(Todos)")
        val.font = Font(name=FONTE, bold=True, color=COR_VINHO)
        val.fill = FILL_INPUT
        val.border = BORDA
        ws.merge_cells(start_row=5, start_column=col_val, end_row=5, end_column=col_val + 1)
    ws.row_dimensions[5].height = 22

    dv_tipo = DataValidation(type="list", formula1=refs["lista_tipo"], allow_blank=False)
    dv_centro = DataValidation(type="list", formula1=refs["lista_centro"], allow_blank=False)
    dv_alerta = DataValidation(type="list", formula1=refs["lista_alerta"], allow_blank=False)
    for dv, addr in ((dv_tipo, "B5"), (dv_centro, "E5"), (dv_alerta, "H5")):
        ws.add_data_validation(dv)
        dv.add(ws[addr])

    secao(ws, 7, 1, 14, "INDICADORES GERAIS (conforme filtro selecionado acima)")
    kpi_card(ws, 8, 1, "TOTAL EM KG",
             sumproduct_filtros([f'({rng("K", last_row)}="KG")'], last_row, soma_col="J"),
             "#,##0")
    kpi_card(ws, 8, 3, "TOTAL EM UN",
             sumproduct_filtros([f'({rng("K", last_row)}="UN")'], last_row, soma_col="J"),
             "#,##0")
    kpi_card(ws, 8, 5, "POSIÇÕES OCUPADAS",
             sumproduct_filtros([f'({rng("S", last_row)}<>"VAZIA")'], last_row),
             "#,##0")
    kpi_card(ws, 8, 7, "POSIÇÕES VAZIAS",
             sumproduct_filtros([f'({rng("S", last_row)}="VAZIA")'], last_row),
             "#,##0")
    f_ocup = sumproduct_filtros([f'({rng("S", last_row)}<>"VAZIA")'], last_row)[1:]
    f_vaz = sumproduct_filtros([f'({rng("S", last_row)}="VAZIA")'], last_row)[1:]
    kpi_card(ws, 8, 9, "TAXA DE OCUPAÇÃO",
             f'=IFERROR({f_ocup}/({f_ocup}+{f_vaz}),0)', "0%")
    n_materiais = len({l["Material"] for l in linhas if l["Material"] != MATERIAL_VAZIO})
    c = ws.cell(row=9, column=11, value=n_materiais)
    c.font = F_KPI
    c.fill = FILL_CARD
    c.alignment = Alignment(horizontal="center")
    c.border = BORDA
    c.comment = Comment(
        "Quantidade de materiais distintos no armazém (todo o armazém, sem "
        "aplicar os filtros acima) na última atualização — calculado pelo "
        "script no momento da carga para manter o arquivo leve e rápido de "
        "recalcular; atualiza sozinho a cada nova carga do SAP.",
        "Dashboard LX03",
    )
    lbl = ws.cell(row=8, column=11, value="MATERIAIS ÚNICOS (armazém)")
    lbl.font = Font(name=FONTE, size=8, bold=True, color=COR_BRANCO)
    lbl.fill = PatternFill("solid", fgColor=COR_VINHO)
    lbl.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    # --------------------------------------------------------------
    # Tabela de riscos/alertas (linha por categoria)
    # --------------------------------------------------------------
    secao(ws, 11, 1, 14, "PONTOS DE ATENÇÃO (conforme filtro de Tipo de Depósito e Centro)")
    cab_alerta = ["Categoria", "O que significa", "Qtd. Posições", "Peso (kg)", "% do KG filtrado", "Status"]
    larg_alerta = [1, 2, 6, 6, 6, 3]  # em nº de colunas mescladas por campo, ver abaixo
    row0 = 12
    col_pos = [1, 3, 8, 10, 12, 13]
    for cab, colc, cole in zip(cab_alerta, col_pos, col_pos[1:] + [14]):
        cell = ws.cell(row=row0, column=colc, value=cab)
        cell.font = F_HDR
        cell.fill = FILL_HDR
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        if cole - 1 > colc:
            ws.merge_cells(start_row=row0, start_column=colc, end_row=row0, end_column=cole - 1)

    total_kg_2f = sumproduct_filtros([f'({rng("K", last_row)}="KG")'], last_row, incluir_alerta=False, soma_col="J")[1:]
    for i, alerta in enumerate(ALERTAS_RISCO):
        r = row0 + 1 + i
        ws.cell(row=r, column=1, value=alerta).font = F_TXT_B
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(row=r, column=3, value=DESCRICAO_ALERTA[alerta]).font = F_TXT
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)

        f_qtd = sumproduct_filtros([f'({rng("S", last_row)}="{alerta}")'], last_row, incluir_alerta=False)
        ws.cell(row=r, column=8, value=f_qtd).number_format = "#,##0"
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)

        f_kg = sumproduct_filtros(
            [f'({rng("S", last_row)}="{alerta}")', f'({rng("K", last_row)}="KG")'],
            last_row, incluir_alerta=False, soma_col="J",
        )
        ws.cell(row=r, column=10, value=f_kg).number_format = "#,##0"
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)

        ws.cell(row=r, column=12,
                value=f'=IFERROR({f_kg[1:]}/{total_kg_2f},0)').number_format = "0.0%"
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=12)

        cor_status = "🟡 ATENÇÃO" if alerta == "PRÓX. VENCIMENTO" else "🔴 ALERTA"
        ws.cell(row=r, column=13, value=f'=IF(H{r}>0,"{cor_status}","✅ OK")')
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=14)
        for cc in range(1, 15):
            ws.cell(row=r, column=cc).border = BORDA
            if cc in (1, 3):
                ws.cell(row=r, column=cc).font = F_TXT if cc == 3 else F_TXT_B
            elif cc not in (13,):
                ws.cell(row=r, column=cc).font = F_TXT
                ws.cell(row=r, column=cc).alignment = Alignment(horizontal="center")

    row_status_range = f"M{row0+1}:M{row0+len(ALERTAS_RISCO)}"
    ws.conditional_formatting.add(row_status_range, CellIsRule(operator="equal", formula=['"🔴 ALERTA"'], fill=FILL_VERMELHO))
    ws.conditional_formatting.add(row_status_range, CellIsRule(operator="equal", formula=['"🟡 ATENÇÃO"'], fill=FILL_AMARELO))
    ws.conditional_formatting.add(row_status_range, CellIsRule(operator="equal", formula=['"✅ OK"'], fill=FILL_VERDE))

    row_apos_alertas = row0 + len(ALERTAS_RISCO) + 2

    # --------------------------------------------------------------
    # Área de apoio aos gráficos (fica fora da área de impressão)
    # --------------------------------------------------------------
    r_graf = row_apos_alertas
    r_graf2 = r_graf + 19
    r_top = r_graf2 + 19
    r_fim_top = r_top + 2 + 10
    apoio_row = r_fim_top + 5
    ws.cell(row=apoio_row - 1, column=1,
            value="Dados de apoio aos gráficos abaixo — não precisa editar; fora da área de impressão.").font = F_NOTA

    # 1) KG por Tipo de Depósito (top 12 da última carga), filtrado por Centro + Alerta
    tipos_top = sorted(
        refs["tipos_para_grafico"],
        key=lambda t: -sum((l.get("Estoque total") or 0) for l in linhas
                            if l.get("Tipo de depósito") == t and l.get("UM básica") == "KG"),
    )[:12]
    ws.cell(row=apoio_row, column=1, value="Tipo de Depósito").font = F_TXT_B
    ws.cell(row=apoio_row, column=2, value="KG (filtrado)").font = F_TXT_B
    for i, t in enumerate(tipos_top):
        r = apoio_row + 1 + i
        ws.cell(row=r, column=1, value=t)
        formula = sumproduct_filtros(
            [f'({rng("K", last_row)}="KG")', f'({rng("A", last_row)}="{t}")'],
            last_row, incluir_alerta=True, soma_col="J",
        )
        ws.cell(row=r, column=2, value=formula).number_format = "#,##0"
    linha_fim_tipos_grafico = apoio_row + len(tipos_top)

    # 2) Distribuição de posições por Alerta, filtrado por Tipo + Centro
    apoio2 = linha_fim_tipos_grafico + 3
    ws.cell(row=apoio2, column=1, value="Alerta").font = F_TXT_B
    ws.cell(row=apoio2, column=2, value="Qtd. Posições").font = F_TXT_B
    for i, a in enumerate(ALERTAS_ORDEM):
        r = apoio2 + 1 + i
        ws.cell(row=r, column=1, value=a)
        formula = sumproduct_filtros([f'({rng("S", last_row)}="{a}")'], last_row, incluir_alerta=False)
        ws.cell(row=r, column=2, value=formula).number_format = "#,##0"
    linha_fim_alerta_grafico = apoio2 + len(ALERTAS_ORDEM)

    # 3) Faixas de "dias sem movimento" (aging), filtrado pelos 3 filtros
    apoio3 = linha_fim_alerta_grafico + 3
    faixas_aging = [("0–30 dias", 0, 30), ("31–60 dias", 31, 60), ("61–90 dias", 61, 90),
                     ("91–180 dias", 91, 180), ("181–365 dias", 181, 365), ("> 365 dias", 366, None)]
    ws.cell(row=apoio3, column=1, value="Faixa (dias sem giro)").font = F_TXT_B
    ws.cell(row=apoio3, column=2, value="Qtd. Posições").font = F_TXT_B
    for i, (rotulo, ini, fim) in enumerate(faixas_aging):
        r = apoio3 + 1 + i
        ws.cell(row=r, column=1, value=rotulo)
        cond = [f'({rng("S", last_row)}<>"VAZIA")', f'({rng("Q", last_row)}<>"")', f'({rng("Q", last_row)}>={ini})']
        if fim is not None:
            cond.append(f'({rng("Q", last_row)}<={fim})')
        formula = sumproduct_filtros(cond, last_row, incluir_alerta=True)
        ws.cell(row=r, column=2, value=formula).number_format = "#,##0"
    linha_fim_aging = apoio3 + len(faixas_aging)

    # 4) Faixas de vencimento, filtrado pelos 3 filtros
    apoio4 = linha_fim_aging + 3
    faixas_venc = [("Vencido", None, -1), ("0–30 dias", 0, 30), ("31–60 dias", 31, 60),
                   ("61–90 dias", 61, 90), ("> 90 dias", 91, None)]
    ws.cell(row=apoio4, column=1, value="Faixa (vencimento)").font = F_TXT_B
    ws.cell(row=apoio4, column=2, value="Qtd. Posições").font = F_TXT_B
    for i, (rotulo, ini, fim) in enumerate(faixas_venc):
        r = apoio4 + 1 + i
        ws.cell(row=r, column=1, value=rotulo)
        cond = [f'({rng("O", last_row)}<>"")']
        if ini is not None:
            cond.append(f'({rng("R", last_row)}>={ini})')
        if fim is not None:
            cond.append(f'({rng("R", last_row)}<={fim})')
        formula = sumproduct_filtros(cond, last_row, incluir_alerta=True)
        ws.cell(row=r, column=2, value=formula).number_format = "#,##0"
    r = apoio4 + 1 + len(faixas_venc)
    ws.cell(row=r, column=1, value="Sem data de vencimento")
    formula = sumproduct_filtros([f'({rng("O", last_row)}="")', f'({rng("S", last_row)}<>"VAZIA")'], last_row, incluir_alerta=True)
    ws.cell(row=r, column=2, value=formula).number_format = "#,##0"
    linha_fim_venc = r

    graf_refs = {
        "tipo": (apoio_row, linha_fim_tipos_grafico),
        "alerta": (apoio2, linha_fim_alerta_grafico),
        "aging": (apoio3, linha_fim_aging),
        "venc": (apoio4, linha_fim_venc),
    }

    def grafico_barras(titulo, faixa, cor, largura=15, altura=8.5):
        r_ini, r_fim = faixa
        ch = BarChart()
        ch.type = "col"
        ch.title = titulo
        ch.style = 10
        ch.y_axis.title = None
        ch.x_axis.title = None
        ch.legend = None
        dados = Reference(ws, min_col=2, min_row=r_ini, max_row=r_fim)
        cats = Reference(ws, min_col=1, min_row=r_ini + 1, max_row=r_fim)
        ch.add_data(dados, titles_from_data=True)
        ch.set_categories(cats)
        ch.series[0].graphicalProperties.solidFill = cor
        ch.width = largura
        ch.height = altura
        return ch

    ws.cell(row=r_graf, column=1, value="Estoque (kg) por Tipo de Depósito").font = F_TXT_B
    ws.add_chart(grafico_barras("KG por Tipo de Depósito", graf_refs["tipo"], COR_VINHO), f"A{r_graf+1}")
    ws.add_chart(grafico_barras("Posições por Alerta", graf_refs["alerta"], COR_DOURADO), f"H{r_graf+1}")
    ws.add_chart(grafico_barras("Dias sem giro (faixas)", graf_refs["aging"], COR_MARROM), f"A{r_graf2}")
    ws.add_chart(grafico_barras("Vencimento (faixas)", graf_refs["venc"], COR_VINHO_ESCURO), f"H{r_graf2}")

    # --------------------------------------------------------------
    # Top 10 — maiores posições paradas / vencidas (respeitam os
    # filtros de Tipo de Depósito e Centro; ver colunas T/U em Dados_SAP)
    # --------------------------------------------------------------
    secao(ws, r_top, 1, 7, "TOP 10 — MAIOR ESTOQUE PARADO/SEM GIRO (kg)")
    secao(ws, r_top, 8, 14, "TOP 10 — MAIOR ESTOQUE VENCIDO (kg)")
    cabecalhos_top = ["Material", "Lote", "Posição", "Tipo Dep.", "kg", "Dias"]
    for base_col, chave_col, titulo_dias in ((1, "T", "Dias parado"), (8, "U", "Dias vencido")):
        for j, cab in enumerate(cabecalhos_top):
            texto = titulo_dias if cab == "Dias" else cab
            cell = ws.cell(row=r_top + 1, column=base_col + j, value=texto)
            cell.font = F_HDR
            cell.fill = FILL_HDR
        for k in range(10):
            r = r_top + 2 + k
            f_valor = f'=IFERROR(LARGE(Dados_SAP!${chave_col}$2:${chave_col}${last_row},{k+1}),"")'
            ws.cell(row=r, column=base_col + 4, value=f_valor).number_format = "#,##0"
            helper_addr = f"{get_column_letter(base_col+4)}{r}"
            f_match = f'MATCH({helper_addr},Dados_SAP!${chave_col}$2:${chave_col}${last_row},0)'
            for campo, col_dados, off in (("Material", "C", 0), ("Lote", "D", 1), ("Posição", "B", 2), ("Tipo Dep.", "A", 3)):
                f = f'=IF({helper_addr}="","",INDEX(Dados_SAP!${col_dados}$2:${col_dados}${last_row},{f_match}))'
                ws.cell(row=r, column=base_col + off, value=f).font = F_TXT
            col_dias_origem = "Q" if chave_col == "T" else "R"
            sinal = "" if chave_col == "T" else "-"
            f_dias = (f'=IF({helper_addr}="","",{sinal}INDEX(Dados_SAP!${col_dias_origem}$2:${col_dias_origem}${last_row},{f_match}))')
            ws.cell(row=r, column=base_col + 5, value=f_dias).number_format = "#,##0"
            for cc in range(base_col, base_col + 6):
                ws.cell(row=r, column=cc).border = BORDA

    ws.print_area = f"A1:N{r_fim_top}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.column_dimensions["A"].width = 18

    return ws, graf_refs, row_apos_alertas


ACAO_RECOMENDADA = {
    "VENCIDO": "Separar fisicamente, abrir processo de baixa/descarte e investigar a causa raiz (previsão de consumo x validade).",
    "PRÓX. VENCIMENTO": "Priorizar consumo/expedição por ordem de vencimento (FEFO) ou negociar saída antes do vencimento.",
    "QUALIDADE": "Acompanhar com a Qualidade o prazo de liberação; escalar inspeções paradas há muito tempo.",
    "BLOQUEADO": "Identificar o motivo do bloqueio com o responsável e definir prazo para liberação ou descarte.",
    "RESTRITO/DEVOLUÇÃO": "Definir destino (retrabalho, devolução a fornecedor/cliente ou descarte) sem deixar acumular.",
    "QUARENTENA": "Confirmar com Qualidade/Compras a liberação do lote; monitorar o tempo em quarentena.",
    "DESCARTE/DEVOLUÇÃO": "Formalizar e executar o descarte/retorno — material parado aqui trava espaço e capital de giro.",
    "SEM LOTE": "Corrigir o apontamento no SAP; lote é obrigatório para rastreabilidade em alimentos.",
    "PARADO/SEM GIRO": "Revisar previsão de demanda, avaliar redistribuição entre centros ou promover o consumo do item.",
}
SEVERIDADE_CRITICA = {"VENCIDO", "DESCARTE/DEVOLUÇÃO", "SEM LOTE"}


def construir_pontos_atencao(wb, last_row):
    ws = wb.create_sheet("Pontos_de_Atencao")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Pontos que Carecem de Monitoramento — Estoque LX03"
    ws["A1"].font = F_TITULO
    ws.merge_cells("A1:H1")
    ws["A2"] = (
        "Lista consolidada, sem filtro (todo o armazém), pensada para leitura de diretoria. "
        "Atualiza sozinha a cada nova carga do SAP — nada aqui precisa ser preenchido à mão."
    )
    ws["A2"].font = F_SUBTITULO
    ws.merge_cells("A2:H2")

    cab = ["Nº", "Ponto de Atenção", "O que significa", "Qtd. Posições", "Peso (kg)", "% do KG total", "Severidade", "Ação Recomendada"]
    for c, h in enumerate(cab, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = F_HDR
        cell.fill = FILL_HDR
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[4].height = 28

    total_kg_formula = f'SUMIFS({rng("J", last_row)},{rng("K", last_row)},"KG")'
    for i, alerta in enumerate(ALERTAS_RISCO):
        r = 5 + i
        ws.cell(row=r, column=1, value=i + 1).font = F_TXT
        ws.cell(row=r, column=2, value=alerta).font = F_TXT_B
        ws.cell(row=r, column=3, value=DESCRICAO_ALERTA[alerta]).font = F_TXT
        f_qtd = f'=COUNTIFS({rng("S", last_row)},"{alerta}")'
        ws.cell(row=r, column=4, value=f_qtd).number_format = "#,##0"
        f_kg = f'=SUMIFS({rng("J", last_row)},{rng("S", last_row)},"{alerta}",{rng("K", last_row)},"KG")'
        ws.cell(row=r, column=5, value=f_kg).number_format = "#,##0"
        ws.cell(row=r, column=6, value=f'=IFERROR(E{r}/({total_kg_formula}),0)').number_format = "0.0%"
        sev_texto = "🔴 CRÍTICO" if alerta in SEVERIDADE_CRITICA else "🟠 ATENÇÃO"
        ws.cell(row=r, column=7, value=f'=IF(D{r}=0,"✅ OK","{sev_texto}")')
        ws.cell(row=r, column=8, value=ACAO_RECOMENDADA[alerta]).font = F_TXT
        for cc in range(1, 9):
            ws.cell(row=r, column=cc).border = BORDA
            ws.cell(row=r, column=cc).alignment = Alignment(
                vertical="center", wrap_text=(cc in (2, 3, 7, 8)), horizontal=("center" if cc in (1, 4, 5, 6) else "left")
            )
        ws.row_dimensions[r].height = 32

    faixa_sev = f"G5:G{4+len(ALERTAS_RISCO)}"
    ws.conditional_formatting.add(faixa_sev, CellIsRule(operator="equal", formula=['"🔴 CRÍTICO"'], fill=FILL_VERMELHO))
    ws.conditional_formatting.add(faixa_sev, CellIsRule(operator="equal", formula=['"🟠 ATENÇÃO"'], fill=FILL_AMARELO))
    ws.conditional_formatting.add(faixa_sev, CellIsRule(operator="equal", formula=['"✅ OK"'], fill=FILL_VERDE))

    larguras = [5, 20, 24, 13, 12, 12, 13, 46]
    for c, w in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    r_fim = 4 + len(ALERTAS_RISCO)
    ws.cell(row=r_fim + 2, column=1,
            value="Para o detalhe posição-a-posição de cada ponto acima, use os filtros da aba Painel ou o filtro automático da aba Dados_SAP (coluna Alerta).").font = F_NOTA
    ws.merge_cells(start_row=r_fim + 2, start_column=1, end_row=r_fim + 2, end_column=8)

    ws.print_area = f"A1:H{r_fim+2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def construir_leiame(wb):
    ws = wb.create_sheet("Leia-me")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100

    ws["A1"] = "Dashboard de Estoque LX03 — Como Usar"
    ws["A1"].font = F_TITULO

    linhas_texto = [
        ("", ""),
        ("O que é este arquivo", (
            "Este dashboard lê o export da transação SAP LX03 (estoque por posição no depósito) "
            "e organiza tudo em indicadores, gráficos e uma lista de pontos de atenção, prontos "
            "para acompanhamento operacional e apresentação à diretoria.")),
        ("Como atualizar todos os dias", (
            "1. No SAP, rode a LX03 e exporte para Excel.\n"
            "2. Salve o arquivo em lx03/entrada/ (pode sobrescrever o de ontem).\n"
            "3. Rode: python lx03/scripts/atualizar_dashboard.py lx03/entrada/NOME_DO_ARQUIVO.xlsx\n"
            "4. Pronto — o script reescreve lx03/dashboard/Dashboard_Estoque_LX03.xlsx com os dados "
            "novos, guarda uma cópia em lx03/historico/ e gera um PDF em lx03/relatorios_pdf/.")),
        ("Aba Painel", (
            "Filtros no topo (Tipo de Depósito, Centro, Alerta) — tudo abaixo (indicadores, tabela de "
            "pontos de atenção, gráficos e os Top 10) recalcula sozinho ao trocar os filtros.")),
        ("Aba Pontos_de_Atencao", (
            "Lista fixa (sem filtro) com todos os pontos que precisam de acompanhamento, o que cada um "
            "significa, quanto pesa em kg/%, a severidade e a ação recomendada. É a aba pensada para "
            "levar pronta a uma reunião de diretoria.")),
        ("Aba Dados_SAP", (
            "Base bruta importada do SAP + colunas de apoio (Alerta, dias até vencer, dias sem giro). "
            "Tem filtro automático do Excel em cada coluna (ícone de funil no cabeçalho) para quem "
            "quiser investigar posição por posição.")),
        ("Aba Parametros", (
            "Células amarelas = pode editar: os dois prazos de alerta (dias para 'próx. vencimento' e "
            "dias para 'parado/sem giro') e a coluna de descrição dos códigos de Tipo de Depósito "
            "(preencha uma vez com o nome de negócio de cada código, ex.: FM1 = Produto Acabado 1).")),
        ("Exportar em PDF / imprimir", (
            "Na aba Painel ou Pontos_de_Atencao: Arquivo > Imprimir (ou Exportar como PDF) — a área e "
            "o layout de impressão já estão configurados (paisagem, ajustado à largura da página). O "
            "script também gera automaticamente um PDF a cada atualização, em lx03/relatorios_pdf/.")),
        ("Cores", (
            "Paleta em tons de vinho/marrom/dourado, inspirada na identidade visual do chocolate/Barry "
            "Callebaut. Os códigos hexadecimais oficiais da marca não puderam ser confirmados neste "
            "ambiente (acesso ao site institucional bloqueado) — se o time de marketing enviar o manual "
            "de marca, os códigos ficam em um único lugar do script "
            "(scripts/atualizar_dashboard.py, bloco 'Paleta de cores') e é só trocar e rodar de novo.")),
        ("Dúvidas / próximos passos", (
            "Veja a seção de ideias de evolução enviada junto com este dashboard (agendamento automático, "
            "alerta por e-mail, histórico de tendência mês a mês, integração direta com o SAP, etc.).")),
    ]
    r = 3
    for titulo, texto in linhas_texto:
        if titulo:
            ws.cell(row=r, column=1, value=titulo).font = Font(name=FONTE, bold=True, size=12, color=COR_VINHO)
            r += 1
            cell = ws.cell(row=r, column=1, value=texto)
            cell.font = F_TXT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 18 * (texto.count("\n") + 2)
            r += 2
        else:
            r += 1
    return ws


def rodar_recalc(caminho_xlsx, timeout=120):
    script = localizar_recalc()
    if not script:
        print("Aviso: não encontrei o recalc.py da skill de xlsx — o arquivo foi salvo, mas "
              "as fórmulas só recalculam quando você abrir no Excel/LibreOffice.")
        return
    resultado = subprocess.run(
        [sys.executable, str(script), str(caminho_xlsx), str(timeout)],
        capture_output=True, text=True,
    )
    print(resultado.stdout.strip())
    if resultado.returncode != 0:
        print("Aviso: recálculo automático falhou (veja acima). Abra o arquivo no Excel para recalcular.")
        if resultado.stderr:
            print(resultado.stderr.strip())


def exportar_pdf(caminho_xlsx, pasta_pdf, data_ref):
    pasta_pdf.mkdir(parents=True, exist_ok=True)
    tmp_path = pasta_pdf / f"_tmp_pdf_{data_ref}.xlsx"
    wb = openpyxl.load_workbook(caminho_xlsx)
    for nome in wb.sheetnames:
        if nome not in ("Painel", "Pontos_de_Atencao", "Tendencia"):
            wb[nome].sheet_state = "hidden"
    wb.save(tmp_path)
    destino_pdf = pasta_pdf / f"Relatorio_LX03_{data_ref}.pdf"
    try:
        subprocess.run(
            ["soffice", "--headless", "--norestore",
             f"-env:UserInstallation=file:///tmp/lo_profile_lx03_{data_ref}",
             "--convert-to", "pdf", "--outdir", str(pasta_pdf), str(tmp_path)],
            capture_output=True, text=True, timeout=120, check=False,
        )
        gerado = pasta_pdf / f"{tmp_path.stem}.pdf"
        if gerado.exists():
            gerado.replace(destino_pdf)
            print(f"PDF gerado em: {destino_pdf}")
        else:
            print("Aviso: não consegui gerar o PDF automaticamente (LibreOffice indisponível). "
                  "Use Arquivo > Exportar como PDF direto no Excel.")
    finally:
        tmp_path.unlink(missing_ok=True)


def calcular_resumo(linhas, dias_venc, dias_parado, hoje):
    """Contagem e kg por categoria de Alerta — usado no resumo do terminal,
    no histórico de tendência e na comparação para os alertas por e-mail/Teams."""
    contagem = {a: 0 for a in ALERTAS_ORDEM}
    kg = {a: 0.0 for a in ALERTAS_ORDEM}
    un = 0.0
    for linha in linhas:
        alerta = calcular_alerta(linha, dias_venc, dias_parado, hoje)
        contagem[alerta] += 1
        if linha.get("UM básica") == "KG":
            kg[alerta] += linha.get("Estoque total") or 0
        elif linha.get("UM básica") == "UN":
            un += linha.get("Estoque total") or 0
    return contagem, kg, un


def imprimir_resumo(contagem, kg):
    print("\n" + "=" * 70)
    print("RESUMO DA ATUALIZAÇÃO — pontos que carecem de monitoramento")
    print("=" * 70)
    for a in ALERTAS_RISCO:
        if contagem[a] > 0:
            print(f"  - {a:<22} {contagem[a]:>6} posições   {kg[a]:>12,.0f} kg")
    print("-" * 70)
    ocupadas = sum(v for k, v in contagem.items() if k != "VAZIA")
    print(f"  Posições ocupadas: {ocupadas} | Posições vazias: {contagem['VAZIA']} | "
          f"Sem alerta (OK): {contagem['OK']}")
    print("=" * 70 + "\n")


def main():
    ap = argparse.ArgumentParser(description="Atualiza o dashboard de estoque LX03 a partir de um novo export do SAP.")
    ap.add_argument("arquivo_sap", nargs="?", help="Caminho do arquivo .xlsx exportado da LX03 no SAP (não usar com --usar-sap)")
    ap.add_argument("--saida", default=str(DASHBOARD_PATH), help="Caminho do dashboard a gerar/atualizar")
    ap.add_argument("--sem-pdf", action="store_true", help="Não gerar o PDF automático")
    ap.add_argument("--sem-historico", action="store_true", help="Não guardar cópia em historico/")
    ap.add_argument("--sem-tendencia", action="store_true", help="Não gravar/gerar a aba e o histórico de Tendência")
    ap.add_argument("--sem-notificacoes", action="store_true", help="Não enviar e-mail/Teams mesmo se configurado")
    ap.add_argument("--sem-nuvem", action="store_true", help="Não publicar no Google Sheets/Power BI mesmo se configurado")
    ap.add_argument("--usar-sap", action="store_true", help="Buscar os dados direto do SAP (OData/RFC) em vez de um arquivo — requer config/config.ini")
    args = ap.parse_args()

    if args.usar_sap:
        import conector_sap
        print("Buscando dados diretamente do SAP (config/config.ini) ...")
        linhas = conector_sap.buscar_dados()
        nome_origem = "SAP (conexão direta)"
    else:
        if not args.arquivo_sap:
            raise SystemExit("ERRO: informe o arquivo do export da LX03, ou use --usar-sap para buscar direto do SAP.")
        caminho_entrada = Path(args.arquivo_sap)
        if not caminho_entrada.exists():
            raise SystemExit(f"ERRO: arquivo não encontrado: {caminho_entrada}")
        print(f"Lendo {caminho_entrada.name} ...")
        linhas = ler_lx03(caminho_entrada)
        nome_origem = caminho_entrada.name
    print(f"{len(linhas)} posições lidas.")

    agora = datetime.datetime.now()
    hoje = agora.date()
    data_referencia = hoje.isoformat()
    carimbo = agora.strftime("%Y-%m-%d_%H%M")
    ultima_atualizacao_txt = agora.strftime("%d/%m/%Y %H:%M")

    contagem, kg, total_un = calcular_resumo(linhas, DIAS_PROX_VENC_PADRAO, DIAS_PARADO_PADRAO, hoje)
    posicoes_ocupadas = sum(v for k, v in contagem.items() if k != "VAZIA")
    posicoes_vazias = contagem["VAZIA"]
    total_kg = sum(kg.values())
    materiais_unicos = len({l["Material"] for l in linhas if l["Material"] != MATERIAL_VAZIO})

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_dados, last_row = None, len(linhas) + 1
    ws_param, refs = construir_parametros(wb, linhas, nome_origem, ultima_atualizacao_txt, last_row)
    ws_dados, last_row = construir_dados_sap(wb, linhas, refs)
    ws_painel, graf_refs, _ = construir_painel(wb, refs, last_row, linhas)
    construir_pontos_atencao(wb, last_row)

    if not args.sem_tendencia:
        import tendencia
        conn = tendencia.conectar()
        tendencia.registrar_snapshot(
            conn, data_referencia, ultima_atualizacao_txt, nome_origem,
            contagem, kg, posicoes_ocupadas, posicoes_vazias, total_kg, total_un, materiais_unicos,
        )
        estilos_tendencia = {
            "titulo": F_TITULO, "subtitulo": F_SUBTITULO, "nota": F_NOTA,
            "cab_fonte": F_HDR, "cab_fill": FILL_HDR,
        }
        tendencia.construir_aba_tendencia(wb, conn, estilos_tendencia)
        conn.close()

    construir_leiame(wb)

    ordem = ["Painel", "Pontos_de_Atencao", "Dados_SAP", "Parametros"]
    if not args.sem_tendencia:
        ordem.append("Tendencia")
    ordem.append("Leia-me")
    wb._sheets = [wb[nome] for nome in ordem]
    wb.active = 0

    caminho_saida = Path(args.saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)
    print(f"Dashboard salvo em: {caminho_saida}")

    rodar_recalc(caminho_saida)

    if not args.sem_historico:
        HISTORICO_DIR.mkdir(parents=True, exist_ok=True)
        if not args.usar_sap:
            shutil.copy2(caminho_entrada, HISTORICO_DIR / f"lx03_origem_{carimbo}.xlsx")
        shutil.copy2(caminho_saida, HISTORICO_DIR / f"dashboard_{carimbo}.xlsx")
        print(f"Cópia de auditoria guardada em: {HISTORICO_DIR}")

    if not args.sem_pdf:
        exportar_pdf(caminho_saida, PDF_DIR, carimbo)

    imprimir_resumo(contagem, kg)

    if not args.sem_notificacoes:
        import notificacoes
        notificacoes.avaliar_e_notificar(data_referencia, contagem, kg, nome_origem, ultima_atualizacao_txt)

    if not args.sem_nuvem:
        import publicar_nuvem
        publicar_nuvem.publicar(caminho_saida)

    print("Atualização concluída com sucesso.")


if __name__ == "__main__":
    main()
