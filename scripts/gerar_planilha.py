"""Gera planilhas/Troca_de_Turno_2026.xlsx (versão enxuta, 5 abas) a partir de dados/base_migrada.json.

Uso:
    python scripts/migrar_csv.py      # 1) limpa o CSV original
    python scripts/gerar_planilha.py  # 2) monta a planilha

Abas: LANÇAMENTOS, PASSAGEM, PENDÊNCIAS, PAINEL, LISTAS.
Sem macros e sem senha. Só fórmulas clássicas do Excel (funciona no Excel 2010+ e no LibreOffice).
"""
import datetime as dt
import json
import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "dados" / "base_migrada.json"
OUT = Path(os.environ.get("TT_OUT", ROOT / "planilhas" / "Troca_de_Turno_2026.xlsx"))
NROWS = 5000                     # linhas preparadas na aba LANÇAMENTOS
R1, RN = 2, NROWS + 1
CORTE = dt.date(2026, 9, 20)     # antes desta data, itens em aberto do arquivo antigo = HISTÓRICO

FONT = "Arial"
NAVY, BLUE, GREY, LGREY, MGREY = "1F3864", "2E75B6", "595959", "F2F2F2", "D9D9D9"
GREEN, LGREEN, YEL, LYEL, RED, LRED = "2E7D32", "E2EFDA", "FFC000", "FFF2CC", "C00000", "FDE2E1"
WHITE = "FFFFFF"

L_ = "LANÇAMENTOS"


def Q(s):
    return f"'{s}'"


TURNOS = ["A", "B", "C", "D"]
RESPS = ["TATIANE", "THIAGO", "ADIB", "VINICIUS", "ADAO"]
PROGS = ["GM", "FM", "FOB", "CIF", "GRANEL", "CACAU", "BR15", "EXPORTAÇÃO", "CROSS DOCKING",
         "TRANSFERÊNCIA", "RECEBIMENTO", "FLOOR", "ABASTECIMENTO", "MOVIMENTAÇÃO", "DESCARTE",
         "INVENTÁRIO", "OUTROS"]
MEDIDAS = ["TON", "PALETE", "POSIÇÃO"]
FATS = ["PENDENTE", "PARCIAL", "FATURADO", "NÃO SE APLICA"]
EMBS = ["PENDENTE", "PARCIAL", "REALIZADO", "NÃO SE APLICA"]
PERIODOS = ["MÊS ATUAL", "ÚLTIMOS 7 DIAS", "ÚLTIMOS 30 DIAS", "TUDO"]

# (chave, cabeçalho, entrada?, largura, formato)
COLS = [
    ("id", "ID", False, 10, None),
    ("status", "STATUS", False, 13, None),
    ("alerta", "ALERTA", False, 30, None),
    ("data", "DATA", True, 11, "dd/mm/yyyy"),
    ("turno", "TURNO", True, 7, None),
    ("resp", "RESPONSÁVEL", True, 13, None),
    ("prog", "PROGRAMAÇÃO", True, 14, None),
    ("rota", "ROTA / CLIENTE", True, 26, None),
    ("med", "MEDIDA", True, 9, None),
    ("plan", "PLANEJADO", True, 11, "#,##0.000"),
    ("exec", "EXECUTADO", True, 11, "#,##0.000"),
    ("pct", "%", False, 7, "0%"),
    ("fat", "FATURAMENTO", True, 13, None),
    ("emb", "EMBARQUE", True, 13, None),
    ("transp", "TRANSPORTADORA", True, 14, None),
    ("pend", "PENDÊNCIA", True, 36, None),
    ("rpend", "RESP. PENDÊNCIA", True, 14, None),
    ("prazo", "PRAZO", True, 11, "dd/mm/yyyy"),
    ("concl", "CONCLUÍDO EM", True, 12, "dd/mm/yyyy"),
    ("obs", "OBSERVAÇÃO", True, 45, None),
    ("dias", "DIAS EM ABERTO", False, 9, "0"),
    # auxiliares (ocultas)
    ("aberto", "aux_aberto", False, 6, "0"),
    ("o_pend", "aux_ordem_pendencias", False, 12, "0"),
    ("o_pass", "aux_ordem_passagem", False, 12, "0"),
]
COL = {k: get_column_letter(i + 1) for i, (k, *_) in enumerate(COLS)}

EMB_ABERTO = 'OR({emb}="PENDENTE",{emb}="PARCIAL")'
FAT_ABERTO = 'OR({fat}="PENDENTE",{fat}="PARCIAL")'
F = {
    "id": '=IF({data}="","","TT-"&TEXT(ROW()-1,"00000"))',
    "aberto": ('=IF({data}="",0,IF(OR(AND({pend}<>"",{concl}=""),AND(N({plan})>0,N({exec})<{plan}),'
               + FAT_ABERTO + "," + EMB_ABERTO + "),1,0))"),
    "status": ('=IF({data}="","",IF({aberto}=0,"CONCLUÍDO",IF({data}<corte,"HISTÓRICO",'
               'IF(OR(AND(ISNUMBER({prazo}),{prazo}<TODAY()),AND(' + EMB_ABERTO + ',{data}<TODAY()-1)),"CRÍTICO",'
               'IF(OR(AND({pend}<>"",{concl}=""),AND({emb}="REALIZADO",{fat}="PENDENTE"),'
               'AND(ISNUMBER({prazo}),{prazo}=TODAY())),"ATENÇÃO",'
               'IF(N({exec})=0,"PENDENTE","EM ANDAMENTO"))))))'),
    "alerta": ('=IF({data}="","",IF({aberto}=0,"🟢 OK",IF({data}<corte,"⚪ HISTÓRICO (sem baixa)",'
               'IF(AND(ISNUMBER({prazo}),{prazo}<TODAY()),"🔴 PRAZO VENCIDO",'
               'IF(AND(' + EMB_ABERTO + ',{data}<TODAY()-1),"🔴 EMBARQUE ATRASADO",'
               'IF(AND({emb}="REALIZADO",{fat}="PENDENTE"),"🟡 EMBARCADO SEM FATURAR",'
               'IF(AND({pend}<>"",{concl}="",OR({rpend}="",{prazo}="")),"🟡 PENDÊNCIA SEM RESPONSÁVEL/PRAZO",'
               'IF(AND(ISNUMBER({prazo}),{prazo}=TODAY()),"🟡 VENCE HOJE",'
               'IF(AND({pend}<>"",{concl}=""),"🟡 PENDÊNCIA ABERTA",'
               'IF(AND(N({plan})>0,N({exec})<{plan}),"🟡 FALTA EXECUTAR",'
               'IF(' + FAT_ABERTO + ',"🟡 FATURAMENTO PENDENTE","🟡 EMBARQUE PENDENTE")))))))))))'),
    "pct": '=IF(OR({data}="",N({plan})=0),"",MIN(1,N({exec})/{plan}))',
    "dias": '=IF(AND({aberto}=1,{data}<>""),IF({data}>=corte,TODAY()-{data},""),"")',
    # críticos primeiro, depois os mais antigos
    "o_pend": ('=IF(AND({aberto}=1,{data}<>""),IF({data}>=corte,({status}="CRÍTICO")*1E9'
               '+(100000-{data})*10000+10000-ROW(),""),"")'),
    "o_pass": '=IF(AND({aberto}=1,{data}=pas_data,{turno}=pas_turno),10000-ROW(),"")',
}


def fx(key, r):
    return F[key].format(**{k: f"{c}{r}" for k, c in COL.items()})


# ---------------------------------------------------------------- estilo
def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


thin = Side(style="thin", color=MGREY)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def st(c, bold=False, size=10, color="000000", bg=None, h=None, wrap=False, fmt=None, italic=False, v="center"):
    c.font = Font(name=FONT, bold=bold, size=size, color=color, italic=italic)
    if bg:
        c.fill = fill(bg)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if fmt:
        c.number_format = fmt
    return c


def put(ws, ref, value, **kw):
    ws[ref] = value
    return st(ws[ref], **kw)


def box(ws, r, c, value, fmt=None, bold=False, h=None, color="000000", size=10, bg=None, wrap=False):
    x = ws.cell(row=r, column=c, value=value)
    st(x, bold=bold, fmt=fmt, h=h, color=color, size=size, bg=bg, wrap=wrap)
    x.border = BORDER
    return x


def header(ws, r, c, labels, bg=NAVY):
    for i, t in enumerate(labels):
        box(ws, r, c + i, t, bold=True, h="center", color=WHITE, size=9, bg=bg, wrap=True)
    ws.row_dimensions[r].height = 28


def title(ws, text, sub, last=12):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last)
    put(ws, "B1", text, bold=True, size=16, color=WHITE, bg=NAVY)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last)
    put(ws, "B2", sub, size=9, color=GREY, italic=True)


def section(ws, r, c, text, span, color=BLUE):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + span - 1)
    for i in range(span):
        ws.cell(row=r, column=c + i).fill = fill(color)
    st(ws.cell(row=r, column=c, value=text), bold=True, size=11, color=WHITE, bg=color)
    ws.row_dimensions[r].height = 20


def yellow(ws, ref, value=None, fmt=None):
    if value is not None:
        ws[ref] = value
    c = st(ws[ref], bold=True, bg=LYEL, h="center", fmt=fmt)
    y = Side(style="thin", color=YEL)
    c.border = Border(left=y, right=y, top=y, bottom=y)
    return c


def card(ws, r, c, label, formula, fmt="#,##0", color=NAVY):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
    ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
    st(ws.cell(row=r, column=c, value=label), bold=True, size=8, color=GREY, bg=LGREY, h="center", wrap=True)
    st(ws.cell(row=r + 1, column=c, value=formula), bold=True, size=18, color=color, bg=LGREY, h="center", fmt=fmt)
    for rr in (r, r + 1):
        for i in range(2):
            ws.cell(row=rr, column=c + i).fill = fill(LGREY)
    for i in range(2):
        ws.cell(row=r, column=c + i).border = Border(top=Side(style="medium", color=color))
    ws.row_dimensions[r].height = 22
    ws.row_dimensions[r + 1].height = 32


def dv(ws, formula, ref, strict=True, prompt=None, kind="list"):
    if kind == "list":
        d = DataValidation(type="list", formula1=formula, allow_blank=True)
        d.error, d.errorTitle = "Escolha um item da lista (aba LISTAS).", "Valor fora da lista"
    elif kind == "date":
        d = DataValidation(type="date", operator="between", formula1="45658", formula2="47848", allow_blank=True)
        d.error, d.errorTitle = "Digite uma data (dd/mm/aaaa).", "Data inválida"
    else:
        d = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
        d.error, d.errorTitle = "Digite um número (zero ou maior).", "Número inválido"
    d.errorStyle = "stop" if strict else "warning"
    if prompt:
        d.prompt, d.promptTitle = prompt, "Dica"
    ws.add_data_validation(d)
    d.add(ref)


def status_colors(ws, rng, first):
    for val, bg, fg in (("CRÍTICO", RED, WHITE), ("ATENÇÃO", YEL, "000000"), ("PENDENTE", "FFE699", "000000"),
                        ("EM ANDAMENTO", "DDEBF7", NAVY), ("CONCLUÍDO", LGREEN, GREEN), ("HISTÓRICO", LGREY, GREY)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{val}"'], fill=fill(bg),
                                                       font=Font(name=FONT, bold=True, color=fg)))


def alert_colors(ws, rng, first):
    for emo, bg, fg in (("🔴", LRED, RED), ("🟡", "FFF7D6", "7F6000"), ("🟢", LGREEN, GREEN)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},2)="{emo}"'], fill=fill(bg),
                                                       font=Font(name=FONT, color=fg)))


def name(wb, n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def dtxt(ref):
    """Data como texto dd/mm/aaaa sem depender do idioma do Excel."""
    return f'RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'


# ---------------------------------------------------------------- lista ordenada
def ranked(ws, top, n, order, cols, count=None):
    """Lista ordenada pela coluna auxiliar `order`. Coluna A (oculta) = posição do registro."""
    header(ws, top, 2, ["#"] + [c[0] for c in cols])
    for k in range(1, n + 1):
        r = top + k
        pos = f"MATCH(LARGE({order},{k}),{order},0)"
        ws.cell(row=r, column=1, value=f'=IFERROR(IF({k}>{count},"",{pos}),"")' if count else f'=IFERROR({pos},"")')
        box(ws, r, 2, f'=IF($A{r}="","",{k})', h="center", color=GREY, size=9)
        for j, (hdr, key, fmt) in enumerate(cols):
            if key == "id":
                f = f'=IF($A{r}="","",HYPERLINK("#{Q(L_)}!A"&($A{r}+1),INDEX(c_id,$A{r})))'
            elif key == "pendobs":
                f = (f'=IF($A{r}="","",IF(INDEX(c_pend,$A{r})&""<>"",INDEX(c_pend,$A{r})&"",'
                     f'INDEX(c_obs,$A{r})&""))')
            elif fmt:
                f = f'=IF($A{r}="","",INDEX(c_{key},$A{r}))'
            else:
                f = f'=IF($A{r}="","",INDEX(c_{key},$A{r})&"")'
            c = box(ws, r, 3 + j, f, fmt=fmt, size=9)
            if key == "id":
                c.font = Font(name=FONT, size=9, color=BLUE, underline="single")
    ws.column_dimensions["A"].hidden = True
    for j, (hdr, key, fmt) in enumerate(cols):
        L = get_column_letter(3 + j)
        rng = f"{L}{top + 1}:{L}{top + n}"
        if key == "status":
            status_colors(ws, rng, f"{L}{top + 1}")
        if key == "alerta":
            alert_colors(ws, rng, f"{L}{top + 1}")
    return top + 1, top + n


D = "dd/mm/yyyy;;"
LIST_COLS = [("ID", "id", None), ("DATA", "data", D), ("TURNO", "turno", None), ("STATUS", "status", None),
             ("ALERTA", "alerta", None), ("PROGRAMAÇÃO", "prog", None), ("ROTA / CLIENTE", "rota", None),
             ("PENDÊNCIA / OBSERVAÇÃO", "pendobs", None), ("RESPONSÁVEL", "rpend", None), ("PRAZO", "prazo", D)]
LIST_W = {"B": 5, "C": 10, "D": 11, "E": 7, "F": 13, "G": 30, "H": 13, "I": 22, "J": 42, "K": 13, "L": 11}


# ================================================================ abas
def build_listas(wb, ws, recs):
    title(ws, "LISTAS E AJUDA", "Itens das listas suspensas. Pode acrescentar nomes nas células amarelas.", 13)
    transps = sorted({r["transp"] for r in recs if r["transp"]})
    lists = [("turno", "TURNO", TURNOS, False), ("resp", "RESPONSÁVEIS", RESPS, True),
             ("prog", "PROGRAMAÇÃO", PROGS, True), ("med", "MEDIDA", MEDIDAS, False),
             ("fat", "FATURAMENTO", FATS, False), ("emb", "EMBARQUE", EMBS, False),
             ("transp", "TRANSPORTADORA", transps, True), ("periodo", "PERÍODO", PERIODOS, False)]
    for i, (key, t, items, dyn) in enumerate(lists):
        col = 2 + i
        L = get_column_letter(col)
        ws.column_dimensions[L].width = 16
        box(ws, 4, col, t, bold=True, h="center", color=WHITE, bg=NAVY, size=9)
        cap = 30 if dyn else len(items)
        for j in range(cap):
            c = box(ws, 5 + j, col, items[j] if j < len(items) else None)
            if dyn:
                c.fill = fill(LYEL)
        end = 4 + cap
        ref = (f"OFFSET({Q('LISTAS')}!${L}$5,0,0,MAX(1,COUNTA({Q('LISTAS')}!${L}$5:${L}${end})),1)" if dyn
               else f"{Q('LISTAS')}!${L}$5:${L}${end}")
        name(wb, f"l_{key}", ref)

    ws.column_dimensions["J"].width = 3
    ws.column_dimensions["K"].width = 30
    ws.column_dimensions["L"].width = 14
    ws.column_dimensions["M"].width = 70
    section(ws, 4, 11, "AJUSTES", 3)
    box(ws, 5, 11, "Pendências contam a partir de", bold=True)
    yellow(ws, "L5", CORTE, "dd/mm/yyyy")
    box(ws, 5, 13, "Itens em aberto do arquivo antigo com data anterior aparecem como HISTÓRICO e não entram "
                   "nas pendências (não tinham responsável nem prazo).", size=9, wrap=True)
    ws.row_dimensions[5].height = 36
    name(wb, "corte", f"{Q('LISTAS')}!$L$5")
    box(ws, 6, 11, "Turno  ➜  próximo turno", bold=True)
    for i, (a, b) in enumerate(zip(TURNOS, TURNOS[1:] + TURNOS[:1])):
        box(ws, 7 + i, 11, a, h="center")
        yellow(ws, f"L{7 + i}", b)
    name(wb, "t_prox", f"{Q('LISTAS')}!$K$7:$L$10")

    section(ws, 13, 11, "COMO USAR", 3, color=GREEN)
    ajuda = [
        ("1. Lançar", "Aba LANÇAMENTOS, primeira linha vazia. Preencha as colunas de cabeçalho AZUL. "
                      "As colunas de cabeçalho CINZA (ID, Status, Alerta, %, Dias) são automáticas."),
        ("2. Pendência", "Se algo ficou pendente: escreva em PENDÊNCIA e informe RESP. PENDÊNCIA e PRAZO."),
        ("3. Dar baixa", "Quando resolver: preencha CONCLUÍDO EM (e atualize Executado / Faturamento / Embarque)."),
        ("4. Trocar turno", "Aba PASSAGEM: o que fica para o próximo turno + resumo pronto para o WhatsApp."),
        ("5. Acompanhar", "PENDÊNCIAS lista tudo em aberto (críticos primeiro). PAINEL mostra os números."),
        ("Status", "CRÍTICO = prazo vencido ou embarque pendente há mais de 1 dia.  ATENÇÃO = pendência aberta, "
                   "embarcado sem faturar ou vence hoje.  PENDENTE = sem execução.  EM ANDAMENTO = parcial.  "
                   "CONCLUÍDO = nada em aberto."),
        ("Cuidado", "Não apague linhas nem ordene LANÇAMENTOS (o ID é a posição da linha). Para anular, "
                    "zere Planejado/Executado, marque Faturamento/Embarque como NÃO SE APLICA e escreva "
                    "CANCELADO na observação."),
    ]
    for i, (a, b) in enumerate(ajuda):
        r = 14 + i
        box(ws, r, 11, a, bold=True, color=NAVY)
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=13)
        box(ws, r, 12, b, size=9, wrap=True)
        ws.row_dimensions[r].height = 40


def build_lancamentos(wb, ws, recs):
    ws.freeze_panes = "E2"
    for i, (k, hdr, entrada, w, fmt) in enumerate(COLS):
        L = get_column_letter(i + 1)
        st(ws.cell(row=1, column=i + 1, value=hdr), bold=True, size=9, color=WHITE,
           bg=NAVY if entrada else GREY, h="center", wrap=True)
        ws.column_dimensions[L].width = w
        name(wb, f"c_{k}", f"{Q(L_)}!${L}${R1}:${L}${RN}")
        if k in ("aberto", "o_pend", "o_pass"):
            ws.column_dimensions[L].hidden = True
    ws.row_dimensions[1].height = 32
    notes = {"status": "Automático. Legenda na aba LISTAS.",
             "plan": "Na MEDIDA escolhida. TON = toneladas (28,054 = 28.054 kg).",
             "pend": "O que ficou pendente. Informe também RESP. PENDÊNCIA e PRAZO.",
             "concl": "Data em que a pendência foi resolvida."}
    for k, t in notes.items():
        ws[f"{COL[k]}1"].comment = Comment(t, "Planilha")

    fnt = Font(name=FONT, size=9)
    fnt_auto = Font(name=FONT, size=9, color="404040")
    inp = fill("FFFDF2")
    fe = {"EM PROCESSO": "PARCIAL"}
    for r in range(R1, RN + 1):
        rec = recs[r - R1] if r - R1 < len(recs) else None
        vals = {}
        if rec:
            vals = {"data": dt.date.fromisoformat(rec["data"]), "turno": rec["turno"], "resp": rec["resp"],
                    "prog": rec["prog"], "rota": rec["rota"], "med": rec["medida"], "plan": rec["plan"],
                    "exec": rec["exec"], "fat": fe.get(rec["fat"], rec["fat"]), "emb": fe.get(rec["emb"], rec["emb"]),
                    "transp": rec["transp"], "pend": rec["pend"], "obs": rec["obs"]}
        for i, (k, hdr, entrada, w, fmt) in enumerate(COLS):
            c = ws.cell(row=r, column=i + 1)
            if entrada:
                v = vals.get(k)
                c.value = v if v not in ("", None) else None
                c.fill = inp
                c.font = fnt
            else:
                c.value = fx(k, r)
                c.font = fnt_auto
            if fmt:
                c.number_format = fmt

    def rng(k):
        return f"{COL[k]}{R1}:{COL[k]}{RN}"
    dv(ws, None, rng("data"), kind="date", prompt="Data do turno (dd/mm/aaaa)")
    dv(ws, "=l_turno", rng("turno"))
    dv(ws, "=l_resp", rng("resp"), strict=False)
    dv(ws, "=l_prog", rng("prog"), strict=False)
    dv(ws, "=l_med", rng("med"), prompt="TON, PALETE ou POSIÇÃO")
    dv(ws, None, rng("plan"), kind="num")
    dv(ws, None, rng("exec"), kind="num")
    dv(ws, "=l_fat", rng("fat"))
    dv(ws, "=l_emb", rng("emb"))
    dv(ws, "=l_transp", rng("transp"), strict=False)
    dv(ws, "=l_resp", rng("rpend"), strict=False, prompt="Quem vai resolver a pendência")
    dv(ws, None, rng("prazo"), kind="date", prompt="Até quando a pendência deve ser resolvida")
    dv(ws, None, rng("concl"), kind="date", prompt="Preencha quando a pendência for resolvida")
    status_colors(ws, rng("status"), f"{COL['status']}{R1}")
    alert_colors(ws, rng("alerta"), f"{COL['alerta']}{R1}")
    for k in ("rpend", "prazo"):  # destaca o que falta quando há pendência
        ws.conditional_formatting.add(rng(k), FormulaRule(
            formula=[f'AND(${COL["pend"]}{R1}<>"",{COL[k]}{R1}="",${COL["concl"]}{R1}="",${COL["data"]}{R1}>=corte)'],
            fill=fill(YEL)))
    ws.auto_filter.ref = f"A1:{COL['dias']}{RN}"


def build_passagem(wb, ws):
    title(ws, "PASSAGEM DE TURNO",
          "O que o próximo turno precisa saber. Deixe Data e Turno vazios para ver o último turno lançado.", 12)
    put(ws, "B4", "Data:", bold=True, h="right")
    yellow(ws, "C4", None, "dd/mm/yyyy")
    dv(ws, None, "C4", kind="date")
    put(ws, "D4", "Turno:", bold=True, h="right")
    yellow(ws, "E4")
    dv(ws, "=l_turno", "E4")
    # valores efetivos (coluna O, discreta)
    ws["O4"] = '=IF(C4="",MAX(c_data),C4)'
    ws["O5"] = "=SUMPRODUCT(MAX((c_data=O4)*ROW(c_data)))"
    ws["O6"] = f'=IF(E4<>"",E4,IF(O5=0,"",INDEX({Q(L_)}!${COL["turno"]}:${COL["turno"]},O5)))'
    ws["O7"] = "=SUMPRODUCT(MAX((c_data=O4)*(c_turno=O6)*ROW(c_data)))"
    ws["O8"] = f'=IF(O7=0,"",INDEX({Q(L_)}!${COL["resp"]}:${COL["resp"]},O7))'
    ws["O9"] = '=IFERROR(VLOOKUP(O6,t_prox,2,FALSE),"")'
    for r in range(4, 10):
        st(ws[f"O{r}"], size=8, color=MGREY, fmt="dd/mm/yyyy" if r == 4 else None)
    name(wb, "pas_data", f"{Q('PASSAGEM')}!$O$4")
    name(wb, "pas_turno", f"{Q('PASSAGEM')}!$O$6")
    ws.merge_cells("B6:L6")
    ws["B6"] = '="TURNO "&O6&"  •  "&' + dtxt("O4") + '&"  •  Líder: "&O8&"      ➜  próximo turno: "&O9'
    st(ws["B6"], bold=True, size=13, color=WHITE, bg=BLUE)
    ws.row_dimensions[6].height = 26
    crit = "c_data,pas_data,c_turno,pas_turno"
    card(ws, 8, 2, "PLANEJADO (t)", f'=SUMIFS(c_plan,{crit},c_med,"TON")', "#,##0.0")
    card(ws, 8, 4, "EXECUTADO (t)", f'=SUMIFS(c_exec,{crit},c_med,"TON")', "#,##0.0")
    card(ws, 8, 6, "% EXECUÇÃO", "=IFERROR(D9/B9,0)", "0%")
    card(ws, 8, 8, "FICA PARA O PRÓXIMO", "=COUNT(c_o_pass)", "0", "7F6000")
    card(ws, 8, 10, "CRÍTICOS EM ABERTO (GERAL)", '=COUNTIF(c_status,"CRÍTICO")', "0", RED)

    section(ws, 11, 2, "📌 O QUE FICA PARA O PRÓXIMO TURNO", 11, color=NAVY)
    f1, l1 = ranked(ws, 12, 15, "c_o_pass", LIST_COLS)
    r = l1 + 2
    section(ws, r, 2, "🔴 CRÍTICOS EM ABERTO (todos os turnos)", 11, color=RED)
    ranked(ws, r + 1, 10, "c_o_pend", LIST_COLS, count='COUNTIF(c_status,"CRÍTICO")')
    for rr in range(f1, l1 + 1):  # linhas do resumo (coluna oculta N)
        ws[f"N{rr}"] = (f'=IF($A{rr}="","","• "&INDEX(c_prog,$A{rr})&" "&INDEX(c_rota,$A{rr})&" — "'
                        f'&MID(INDEX(c_alerta,$A{rr}),3,60)'
                        f'&IF(INDEX(c_rpend,$A{rr})&""="",""," (resp.: "&INDEX(c_rpend,$A{rr})&")"))')
    ws.column_dimensions["N"].hidden = True
    r += 13
    section(ws, r, 2, "📝 RESUMO PRONTO PARA COPIAR (WhatsApp / e-mail)", 11, color=GREEN)
    nl = "&CHAR(10)&"
    resumo = ('="TURNO "&O6&" — "&' + dtxt("O4") + nl + '"Líder: "&O8' + nl + '""' + nl
              + '"Planejado: "&FIXED(B9,1)&" t"' + nl + '"Executado: "&FIXED(D9,1)&" t"' + nl
              + '"Atingimento: "&ROUND(F9*100,0)&"%"' + nl + '""' + nl
              + f'"Fica para o turno "&O9&":"{nl}IF(N{f1}="","• Nada pendente",_xlfn.TEXTJOIN(CHAR(10),TRUE,N{f1}:N{l1}))'
              + nl + '""' + nl + '"Críticos em aberto (geral): "&J9')
    ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 22, end_column=10)
    ws.cell(row=r + 1, column=2, value=resumo)
    st(ws.cell(row=r + 1, column=2), size=10, wrap=True, v="top", bg="FAFAFA")
    for k, v in LIST_W.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = "A7"


def build_pendencias(wb, ws):
    title(ws, "PENDÊNCIAS EM ABERTO", "Tudo o que ainda precisa ser resolvido — críticos primeiro, depois os mais antigos. "
          "Clique no ID para ir à linha e dar baixa (CONCLUÍDO EM).", 13)
    card(ws, 4, 2, "EM ABERTO", "=COUNT(c_o_pend)", "0")
    card(ws, 4, 4, "CRÍTICOS", '=COUNTIF(c_status,"CRÍTICO")', "0", RED)
    card(ws, 4, 6, "ATENÇÃO", '=COUNTIF(c_status,"ATENÇÃO")', "0", "7F6000")
    card(ws, 4, 8, "SEM RESPONSÁVEL / PRAZO", ('=COUNTIFS(c_pend,"<>",c_concl,"",c_rpend,"",c_data,">="&corte)'
                                                '+COUNTIFS(c_pend,"<>",c_concl,"",c_rpend,"<>",c_prazo,"",c_data,">="&corte)'), "0", "7F6000")
    card(ws, 4, 10, "EMBARQUES ATRASADOS", '=COUNTIF(c_alerta,"*EMBARQUE ATRASADO*")', "0", RED)
    ranked(ws, 7, 200, "c_o_pend", LIST_COLS + [("DIAS", "dias", "0")])
    for k, v in {**LIST_W, "M": 7}.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = "A8"


def build_painel(wb, ws):
    title(ws, "PAINEL — TROCA DE TURNO 2026",
          "Números calculados automaticamente a partir de LANÇAMENTOS (volumes em toneladas).", 12)
    put(ws, "B4", "Período:", bold=True, h="right")
    ws.merge_cells("C4:D4")
    yellow(ws, "C4", "MÊS ATUAL")
    dv(ws, "=l_periodo", "C4")
    ws["K4"] = ('=IF(C4="ÚLTIMOS 7 DIAS",TODAY()-6,IF(C4="ÚLTIMOS 30 DIAS",TODAY()-29,'
                'IF(C4="TUDO",DATE(2000,1,1),DATE(YEAR(TODAY()),MONTH(TODAY()),1))))')
    ws["L4"] = '=IF(C4="TUDO",DATE(2100,1,1),IF(C4="MÊS ATUAL",EOMONTH(TODAY(),0),TODAY()))'
    for ref in ("K4", "L4"):
        st(ws[ref], size=8, color=MGREY, fmt="dd/mm/yyyy")
    name(wb, "p_ini", f"{Q('PAINEL')}!$K$4")
    name(wb, "p_fim", f"{Q('PAINEL')}!$L$4")
    P = 'c_data,">="&p_ini,c_data,"<="&p_fim'
    A = 'c_aberto,1,c_data,">="&MAX(p_ini,corte),c_data,"<="&p_fim'
    card(ws, 6, 2, "PLANEJADO (t)", f'=SUMIFS(c_plan,c_med,"TON",{P})', "#,##0.0")
    card(ws, 6, 4, "EXECUTADO (t)", f'=SUMIFS(c_exec,c_med,"TON",{P})', "#,##0.0")
    card(ws, 6, 6, "% EXECUÇÃO", "=IFERROR(D7/B7,0)", "0.0%")
    card(ws, 6, 8, "LANÇAMENTOS", f"=COUNTIFS({P})", "#,##0")
    card(ws, 6, 10, "CRÍTICOS EM ABERTO", f'=COUNTIFS(c_status,"CRÍTICO",{P})', "0", RED)
    card(ws, 9, 2, "EM ABERTO", f"=COUNTIFS({A})", "0", "7F6000")
    card(ws, 9, 4, "FATURAMENTO PENDENTE", f'=COUNTIFS(c_fat,"PENDENTE",{P})+COUNTIFS(c_fat,"PARCIAL",{P})', "0", "7F6000")
    card(ws, 9, 6, "EMBARQUES PENDENTES", f'=COUNTIFS(c_emb,"PENDENTE",{P})+COUNTIFS(c_emb,"PARCIAL",{P})', "0", "7F6000")
    card(ws, 9, 8, "EMBARQUES REALIZADOS", f'=COUNTIFS(c_emb,"REALIZADO",{P})', "0", GREEN)
    card(ws, 9, 10, "CONCLUÍDOS", f'=COUNTIFS(c_status,"CONCLUÍDO",{P})', "0", GREEN)
    ws.conditional_formatting.add("F7", FormulaRule(formula=["F7<0.95"], font=Font(name=FONT, size=18, bold=True, color=RED)))

    section(ws, 12, 2, "POR TURNO", 6)
    header(ws, 13, 2, ["TURNO", "PLANEJADO (t)", "EXECUTADO (t)", "% EXEC.", "EM ABERTO", "CRÍTICOS"])
    for i, t in enumerate(TURNOS):
        r = 14 + i
        box(ws, r, 2, t, bold=True, h="center")
        box(ws, r, 3, f'=SUMIFS(c_plan,c_turno,B{r},c_med,"TON",{P})', "#,##0.0")
        box(ws, r, 4, f'=SUMIFS(c_exec,c_turno,B{r},c_med,"TON",{P})', "#,##0.0")
        box(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', "0%", h="center")
        box(ws, r, 6, f"=COUNTIFS(c_turno,B{r},{A})", "0", h="center")
        box(ws, r, 7, f'=COUNTIFS(c_turno,B{r},c_status,"CRÍTICO",{P})', "0", h="center")

    section(ws, 20, 2, "POR PROGRAMAÇÃO", 6)
    header(ws, 21, 2, ["PROGRAMAÇÃO", "PLANEJADO (t)", "EXECUTADO (t)", "% EXEC.", "EM ABERTO", "LANÇAMENTOS"])
    for i, p in enumerate(PROGS[:10]):
        r = 22 + i
        box(ws, r, 2, p, bold=True)
        box(ws, r, 3, f'=SUMIFS(c_plan,c_prog,B{r},c_med,"TON",{P})', "#,##0.0;;-")
        box(ws, r, 4, f'=SUMIFS(c_exec,c_prog,B{r},c_med,"TON",{P})', "#,##0.0;;-")
        box(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', "0%", h="center")
        box(ws, r, 6, f"=COUNTIFS(c_prog,B{r},{A})", "0;;-", h="center")
        box(ws, r, 7, f"=COUNTIFS(c_prog,B{r},{P})", "0;;-", h="center")
    for rng in ("E14:E17", "E22:E31"):
        first = rng.split(":")[0]
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({first}),{first}<0.95)"],
                                                       font=Font(name=FONT, bold=True, color=RED)))

    ch = BarChart()
    ch.type = "col"
    ch.title = "Planejado x Executado por turno (t)"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=13, max_row=17), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=14, max_row=17))
    ch.width, ch.height = 15, 7.5
    ch.legend.position = "b"
    ch.y_axis.majorGridlines = None
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    for s, c in zip(ch.series, (MGREY, NAVY)):
        s.graphicalProperties.solidFill = c
        s.graphicalProperties.line.solidFill = c
    ws.add_chart(ch, "I12")
    for k, v in {"B": 14, "C": 13, "D": 13, "E": 10, "F": 11, "G": 12, "H": 3}.items():
        ws.column_dimensions[k].width = v


def main():
    recs = json.loads(SRC.read_text())["registros"]
    wb = Workbook()
    ws_l = wb.active
    ws_l.title = L_
    sheets = {"PASSAGEM": NAVY, "PENDÊNCIAS": RED, "PAINEL": BLUE, "LISTAS": GREY}
    ws = {n: wb.create_sheet(n) for n in sheets}
    ws_l.sheet_properties.tabColor = YEL
    for n, cor in sheets.items():
        ws[n].sheet_properties.tabColor = cor
    build_listas(wb, ws["LISTAS"], recs)
    build_lancamentos(wb, ws_l, recs)
    build_passagem(wb, ws["PASSAGEM"])
    build_pendencias(wb, ws["PENDÊNCIAS"])
    build_painel(wb, ws["PAINEL"])
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"OK -> {OUT}  ({len(recs)} registros)")


if __name__ == "__main__":
    main()
