"""Gera planilhas/Troca_de_Turno_2026_Gestao_Operacional.xlsx a partir de dados/base_migrada.json.

Uso:
    python scripts/migrar_csv.py      # 1) limpa o CSV original
    python scripts/gerar_planilha.py  # 2) monta a planilha

Somente fórmulas nativas do Excel (compatíveis com Excel 2016+ e LibreOffice):
nada de FILTER/SORT/UNIQUE/XLOOKUP. Listas ordenadas usam colunas de pontuação
(aux_sc_*) na BASE_OPERACIONAL + LARGE/MATCH/INDEX.
"""
import datetime as dt
import json
import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "dados" / "base_migrada.json"
OUT = Path(os.environ.get("TT_OUT", ROOT / "planilhas" / "Troca_de_Turno_2026_Gestao_Operacional.xlsx"))

SENHA = "tt2026"
SKIP = set(filter(None, os.environ.get("TT_SKIP", "").split(",")))  # só para testes
NROWS = int(os.environ.get("TT_NROWS", 4000))  # capacidade da BASE_OPERACIONAL
R1, RN = 2, NROWS + 1             # primeira / última linha de dados
FONT = "Arial"

# ---------------------------------------------------------------- paleta
NAVY, BLUE, LBLUE = "1F3864", "2E75B6", "DDEBF7"
GREY, MGREY, LGREY = "595959", "D9D9D9", "F2F2F2"
GREEN, LGREEN = "2E7D32", "E2EFDA"
YEL, LYEL, AMBER = "FFC000", "FFF2CC", "FFE699"
RED, LRED = "C00000", "FDE2E1"
WHITE = "FFFFFF"

S_BASE = "BASE_OPERACIONAL"
Q = lambda name: f"'{name}'"      # nome de aba sempre entre aspas


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


thin = Side(style="thin", color=MGREY)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
UNLOCK = Protection(locked=False)

# ---------------------------------------------------------------- cadastros
TURNOS = [("A", "B", dt.time(6, 0)), ("B", "C", dt.time(18, 0)),
          ("C", "D", dt.time(6, 0)), ("D", "A", dt.time(18, 0))]
RESPS = ["TATIANE", "THIAGO", "ADIB", "VINICIUS", "ADAO"]
PROGS = [("GM", "EXPEDIÇÃO"), ("FM", "EXPEDIÇÃO"), ("FOB", "EXPEDIÇÃO"), ("CIF", "EXPEDIÇÃO"),
         ("GRANEL", "EXPEDIÇÃO"), ("CACAU", "EXPEDIÇÃO"), ("BR15", "EXPEDIÇÃO"),
         ("EXPORTAÇÃO", "EXPEDIÇÃO"), ("CROSS DOCKING", "EXPEDIÇÃO"), ("TRANSFERÊNCIA", "EXPEDIÇÃO"),
         ("RECEBIMENTO", "RECEBIMENTO"), ("FLOOR", "ARMAZENAGEM"), ("ABASTECIMENTO", "ARMAZENAGEM"),
         ("MOVIMENTAÇÃO", "ARMAZENAGEM"), ("DESCARTE", "ARMAZENAGEM"), ("INVENTÁRIO", "INVENTÁRIO"),
         ("OUTROS", "OUTROS")]
ATIVS = ["EXPEDIÇÃO", "SEPARAÇÃO", "CONFERÊNCIA", "CARREGAMENTO", "RECEBIMENTO", "ARMAZENAGEM",
         "ABASTECIMENTO", "TRANSFERÊNCIA", "INVENTÁRIO", "MOVIMENTAÇÃO", "OUTROS"]
MEDIDAS = ["TON", "PALETE", "POSIÇÃO"]
STATUS = ["CRÍTICO", "ATENÇÃO", "PENDENTE", "EM ANDAMENTO", "CONCLUÍDO", "CANCELADO"]
FATS = ["PENDENTE", "EM PROCESSO", "FATURADO", "NÃO SE APLICA"]
EMBS = ["PENDENTE", "EM PROCESSO", "REALIZADO", "NÃO SE APLICA"]
IMPACTOS = [("SEM IMPACTO", 0), ("BAIXO", 1), ("MÉDIO", 2), ("ALTO", 3), ("CRÍTICO", 4)]
TIPOS = ["SEPARAÇÃO", "CONFERÊNCIA", "FATURAMENTO", "EMBARQUE", "TRANSPORTE", "SISTEMA",
         "DOCUMENTAÇÃO", "QUALIDADE", "INVENTÁRIO", "EQUIPAMENTO", "MÃO DE OBRA", "COMUNICAÇÃO",
         "PROGRAMAÇÃO", "OUTROS"]
CAUSAS = ["PESSOAS", "PROCESSO", "SISTEMA", "EQUIPAMENTO", "TRANSPORTE", "MATERIAL",
          "PLANEJAMENTO", "COMUNICAÇÃO", "FORNECEDOR", "CLIENTE", "OUTROS"]
SLA = {"SEPARAÇÃO": 4, "CONFERÊNCIA": 4, "FATURAMENTO": 2, "EMBARQUE": 1, "TRANSPORTE": 4,
       "SISTEMA": 4, "DOCUMENTAÇÃO": 2, "QUALIDADE": 8, "INVENTÁRIO": 24, "EQUIPAMENTO": 4,
       "MÃO DE OBRA": 8, "COMUNICAÇÃO": 4, "PROGRAMAÇÃO": 8, "OUTROS": 8}
PERIODOS = ["MÊS ATUAL", "SEMANA ATUAL", "ÚLTIMOS 7 DIAS", "ÚLTIMOS 30 DIAS", "HOJE",
            "TODO O HISTÓRICO", "PERSONALIZADO"]

# ---------------------------------------------------------------- BASE: colunas
# (chave, cabeçalho, tipo, largura, formato)   tipo: in = entrada | auto = fórmula | aux = técnica (oculta)
COLS = [
    ("id", "ID", "auto", 15, "@"),
    ("data", "DATA", "in", 11, "dd/mm/yyyy"),
    ("turno", "TURNO", "in", 7, "@"),
    ("resp", "RESPONSÁVEL PELO TURNO", "in", 14, "@"),
    ("area", "ÁREA", "auto", 13, "@"),
    ("ativ", "TIPO DE ATIVIDADE", "in", 14, "@"),
    ("prog", "PROGRAMAÇÃO", "in", 14, "@"),
    ("rota", "ROTA", "in", 24, "@"),
    ("med", "MEDIDA", "in", 9, "@"),
    ("plan", "PLANEJADO", "in", 11, "#,##0.000"),
    ("exec", "EXECUTADO", "in", 11, "#,##0.000"),
    ("pct", "% EXECUÇÃO", "auto", 10, "0%"),
    ("status", "STATUS", "auto", 14, "@"),
    ("alerta", "ALERTA", "auto", 30, "@"),
    ("fat", "FATURAMENTO", "in", 13, "@"),
    ("emb", "EMBARQUE", "in", 13, "@"),
    ("volfat", "VOLUME FATURADO", "in", 11, "#,##0.000"),
    ("transp", "TRANSPORTADORA", "in", 14, "@"),
    ("pend", "PENDÊNCIA", "in", 38, "@"),
    ("tipo", "TIPO DE PENDÊNCIA", "in", 15, "@"),
    ("causa", "CAUSA", "in", 13, "@"),
    ("rtrat", "RESPONSÁVEL PELA TRATATIVA", "in", 15, "@"),
    ("prazo", "PRAZO", "in", 15, "dd/mm/yyyy hh:mm"),
    ("acao", "AÇÃO", "in", 30, "@"),
    ("concl", "DATA DE CONCLUSÃO", "in", 15, "dd/mm/yyyy hh:mm"),
    ("ttrat", "TEMPO DE TRATATIVA (h)", "auto", 11, "#,##0.0"),
    ("obs", "OBSERVAÇÃO", "in", 45, "@"),
    ("imp", "IMPACTO OPERACIONAL", "in", 13, "@"),
    ("prox", "PRÓXIMO TURNO", "auto", 9, "@"),
    ("atual", "DATA/HORA DA ATUALIZAÇÃO", "in", 15, "dd/mm/yyyy hh:mm"),
    ("canc", "CANCELADO?", "in", 10, "@"),
    ("origem", "ORIGEM", "auto", 11, "@"),
    ("impacta", "IMPACTA PRÓXIMO TURNO?", "auto", 11, "@"),
    ("sitpz", "SITUAÇÃO DO PRAZO", "auto", 13, "@"),
    ("dias", "DIAS EM ABERTO", "auto", 9, "0"),
    ("sla", "SLA (h)", "auto", 8, "0"),
    ("slaest", "SLA ESTOURADO?", "auto", 10, "@"),
    ("valid", "VALIDAÇÃO (REGRAS DE NEGÓCIO)", "auto", 40, "@"),
    ("qual", "QUALIDADE DO CADASTRO", "auto", 35, "@"),
    ("saldo", "SALDO A EXECUTAR", "auto", 11, "#,##0.000"),
    # --- técnicas (agrupadas/ocultas)
    ("abert", "aux_abertura", "aux", 15, "dd/mm/yyyy hh:mm"),
    ("semana", "aux_semana", "aux", 11, "dd/mm/yyyy"),
    ("mes", "aux_mes", "aux", 11, "mm/yyyy"),
    ("ativo", "aux_ativo", "aux", 6, "0"),
    ("aberto", "aux_aberto", "aux", 6, "0"),
    ("abat", "aux_aberto_ativo", "aux", 6, "0"),
    ("venc", "aux_vencido", "aux", 6, "0"),
    ("embatr", "aux_embarque_atrasado", "aux", 6, "0"),
    ("regra1", "aux_embarcado_sem_fat", "aux", 6, "0"),
    ("impn", "aux_impacto_peso", "aux", 6, "0"),
    ("ocorr", "aux_ocorrencia", "aux", 6, "0"),
    ("atraso", "aux_atraso", "aux", 6, "0"),
    ("noprazo", "aux_concluido_no_prazo", "aux", 6, "0"),
    ("key", "aux_chave", "aux", 20, "@"),
    ("dup", "aux_duplicado", "aux", 6, "0"),
    ("fdash", "aux_filtro_dashboard", "aux", 6, "0"),
    ("fprod", "aux_filtro_produtividade", "aux", 6, "0"),
    ("sc_prio", "aux_sc_prioridade", "aux", 12, "0"),
    ("sc_agora", "aux_sc_agora", "aux", 12, "0"),
    ("sc_aten", "aux_sc_atencao", "aux", 12, "0"),
    ("sc_crit", "aux_sc_critico", "aux", 12, "0"),
    ("sc_dia", "aux_sc_dia", "aux", 12, "0"),
    ("sc_pconc", "aux_sc_pass_concl", "aux", 12, "0"),
    ("sc_ppend", "aux_sc_pass_pend", "aux", 12, "0"),
    ("sc_hoje", "aux_sc_vence_hoje", "aux", 12, "0"),
    ("sc_fat", "aux_sc_faturamento", "aux", 12, "0"),
    ("sc_emb", "aux_sc_embarque", "aux", 12, "0"),
    ("sc_hist", "aux_sc_historico", "aux", 12, "0"),
    ("sc_qual", "aux_sc_qualidade", "aux", 12, "0"),
]
COL = {k: get_column_letter(i + 1) for i, (k, *_) in enumerate(COLS)}
KIND = {k: kind for k, _, kind, *_ in COLS}
NUMERIC = {"data", "plan", "exec", "pct", "volfat", "prazo", "concl", "ttrat", "atual", "dias",
           "sla", "saldo", "abert"}


def nest(pairs, default):
    s = default
    for cond, val in reversed(pairs):
        s = f"IF({cond},{val},{s})"
    return s


# fórmulas da BASE ({chave} vira a célula da mesma linha)
_alert_pairs = [
    ('{canc}="SIM"', '"CANCELADO"'),
    ("{venc}=1", '"PRAZO VENCIDO"'),
    ("AND({impn}>=4,{aberto}=1)", '"IMPACTO CRÍTICO"'),
    ("{embatr}=1", '"EMBARQUE ATRASADO"'),
    ("{regra1}=1", '"EMBARCADO SEM FATURAMENTO"'),
    ('{slaest}="SIM"', '"SLA ESTOURADO"'),
    ('AND({aberto}=1,{sitpz}="VENCE HOJE")', '"VENCE HOJE"'),
    ('LEFT({valid},4)="ERRO"', '"ERRO DE CADASTRO"'),
    ('LEFT({valid},4)="INCO"', '"INCONSISTÊNCIA"'),
    ('LEFT({valid},4)="DIVE"', '"DIVERGÊNCIA"'),
    ("{saldo}>0", '"SALDO A EXECUTAR"'),
    ('OR({fat}="PENDENTE",{fat}="EM PROCESSO")', '"FATURAMENTO PENDENTE"'),
    ('OR({emb}="PENDENTE",{emb}="EM PROCESSO")', '"EMBARQUE PENDENTE"'),
    ('AND({pend}<>"",{concl}="")', '"PENDÊNCIA ABERTA"'),
    ("{aberto}=1", '"SEM EXECUÇÃO"'),
]
_semaforo = ('IF({status}="CRÍTICO","🔴 ",IF({status}="CONCLUÍDO","🟢 ",'
             'IF({status}="CANCELADO","⚪ ","🟡 ")))')
_status = nest([
    ('{canc}="SIM"', '"CANCELADO"'),
    ("{aberto}=0", '"CONCLUÍDO"'),
    ('OR({venc}=1,{impn}>=4,{embatr}=1,AND({tipo}="SISTEMA",{impn}>=3,{abat}=1),'
     'AND({abat}=1,{impn}>=3,{pend}<>"",{concl}=""))', '"CRÍTICO"'),
    ('OR(AND({pend}<>"",{concl}=""),{regra1}=1,{sitpz}="VENCE HOJE",'
     'AND(ISNUMBER({prazo}),({prazo}-cfg_agora)*24<=cfg_h_prox),{slaest}="SIM",{impn}>=3,{valid}<>"")',
     '"ATENÇÃO"'),
    ("N({exec})=0", '"PENDENTE"'),
], '"EM ANDAMENTO"')

F = {
    "id": '=IF({data}="","","TT-"&YEAR({data})&"-"&TEXT(ROW()-1,"00000"))',
    "area": '=IF({prog}="","",IFERROR(VLOOKUP({prog},t_prog,2,FALSE),"OUTROS"))',
    "pct": '=IF(OR({data}="",N({plan})=0),"",IF(N({exec})>={plan},1,N({exec})/{plan}))',
    "status": '=IF({data}="","",' + _status + ")",
    "alerta": '=IF({data}="","",' + _semaforo + "&" + nest(_alert_pairs, '"OK"') + ")",
    "ttrat": '=IF(AND(ISNUMBER({concl}),ISNUMBER({abert})),ROUND(MAX(0,({concl}-{abert})*24),1),"")',
    "prox": '=IF({turno}="","",IFERROR(VLOOKUP({turno},t_turno,2,FALSE),""))',
    "origem": None,  # valor fixo (MIGRAÇÃO) ou vazio = lançamento
    "impacta": '=IF({data}="","",IF({abat}=1,"SIM","NÃO"))',
    "sitpz": ('=IF({data}="","",IF({canc}="SIM","CANCELADO",IF({aberto}=0,"CONCLUÍDO",'
              'IF(NOT(ISNUMBER({prazo})),"SEM PRAZO",IF({prazo}<cfg_agora,"VENCIDO",'
              'IF(INT({prazo})=INT(cfg_agora),"VENCE HOJE","NO PRAZO"))))))'),
    "dias": ('=IF(OR({data}="",{canc}="SIM"),"",IF({aberto}=1,MAX(0,INT(cfg_agora)-{data}),'
             'IF(ISNUMBER({concl}),MAX(0,INT({concl})-{data}),0)))'),
    "sla": '=IF({tipo}="","",IFERROR(VLOOKUP({tipo},t_sla,2,FALSE),""))',
    "slaest": ('=IF(OR({sla}="",{abert}=""),"",IF(ISNUMBER({concl}),IF(({concl}-{abert})*24>{sla},"SIM","NÃO"),'
               'IF({abat}=1,IF((cfg_agora-{abert})*24>{sla},"SIM","NÃO"),"")))'),
    "valid": ('=IF({data}="","",TRIM('
              'IF(AND({ativo}=1,{pend}<>"",{rtrat}=""),"ERRO DE CADASTRO: pendência sem responsável. ","")'
              '&IF(AND({ativo}=1,{pend}<>"",NOT(ISNUMBER({prazo}))),"ERRO DE CADASTRO: pendência sem prazo. ","")'
              '&IF(AND({pend}<>"",{tipo}=""),"ERRO DE CADASTRO: pendência sem tipo. ","")'
              '&IF(AND(ISNUMBER({concl}),OR({saldo}>0,{fat}="PENDENTE",{fat}="EM PROCESSO",{emb}="PENDENTE",'
              '{emb}="EM PROCESSO")),"INCONSISTÊNCIA: conclusão informada com item ainda aberto. ","")'
              '&IF(AND(ISNUMBER({concl}),{concl}<{data}),"INCONSISTÊNCIA: conclusão anterior à data. ","")'
              '&IF(AND(ISNUMBER({prazo}),{prazo}<{data}),"INCONSISTÊNCIA: prazo anterior à data. ","")'
              '&IF(AND(N({plan})>0,N({exec})>{plan}*(1+cfg_tol_div)),"DIVERGÊNCIA: executado acima do planejado. ","")'
              "))"),
    "qual": ('=IF({data}="","",TRIM('
             'IF({turno}="","Turno vazio. ",IF(COUNTIF(l_turno,{turno})=0,"Turno fora do padrão. ",""))'
             '&IF({resp}="","Responsável vazio. ","")'
             '&IF({prog}="","Programação vazia. ",IF(COUNTIF(l_prog,{prog})=0,"Programação fora do padrão. ",""))'
             '&IF({med}="","Medida vazia. ",IF(COUNTIF(l_medida,{med})=0,"Medida fora do padrão. ",""))'
             '&IF(N({plan})=0,"Sem planejamento. ","")'
             '&IF(AND(N({plan})>0,{exec}=""),"Sem execução. ","")'
             '&IF(AND({med}=cfg_unid,MAX(N({plan}),N({exec}))>cfg_lim_val),"Valor acima do limite. ","")'
             '&IF(AND({fat}<>"",COUNTIF(l_fat,{fat})=0),"Faturamento fora do padrão. ","")'
             '&IF(AND({emb}<>"",COUNTIF(l_emb,{emb})=0),"Embarque fora do padrão. ","")'
             '&IF(AND({transp}<>"",COUNTIF(l_transp,{transp})=0),"Transportadora fora do cadastro. ","")'
             '&IF({dup}=1,"Possível duplicidade. ","")'
             "))"),
    "saldo": '=IF(N({plan})=0,0,MAX(0,{plan}-N({exec})))',
    "abert": '=IF({data}="","",{data}+IFERROR(VLOOKUP({turno},t_turno,3,FALSE),0))',
    "semana": '=IF({data}="","",{data}-WEEKDAY({data},2)+1)',
    "mes": '=IF({data}="","",DATE(YEAR({data}),MONTH({data}),1))',
    "ativo": '=IF({data}="",0,IF(AND({origem}="MIGRAÇÃO",{data}<cfg_corte),0,1))',
    "aberto": ('=IF(OR({data}="",{canc}="SIM"),0,IF(OR(AND({pend}<>"",{concl}=""),{saldo}>0,'
               '{fat}="PENDENTE",{fat}="EM PROCESSO",{emb}="PENDENTE",{emb}="EM PROCESSO",'
               'AND(N({plan})=0,N({exec})=0,{pend}="")),1,0))'),
    "abat": "={aberto}*{ativo}",
    "venc": "=IF({aberto}=1,IF(ISNUMBER({prazo}),IF({prazo}<cfg_agora,1,0),0),0)",
    "embatr": ('=IF(AND({ativo}=1,{canc}<>"SIM",OR({emb}="PENDENTE",{emb}="EM PROCESSO")),'
               'IF(ISNUMBER({prazo}),IF({prazo}<cfg_agora,1,0),IF({data}+cfg_tol_emb<INT(cfg_agora),1,0)),0)'),
    "regra1": '=IF(AND({canc}<>"SIM",{emb}="REALIZADO",{fat}="PENDENTE"),1,0)',
    "impn": '=IF({imp}="",0,IFERROR(VLOOKUP({imp},t_impacto,2,FALSE),0))',
    "ocorr": '=IF(AND({pend}<>"",{canc}<>"SIM"),1,0)',
    "atraso": ('=MAX({embatr},IF(NOT(ISNUMBER({prazo})),0,IF(ISNUMBER({concl}),IF({concl}>{prazo},1,0),'
               'IF(AND({aberto}=1,{prazo}<cfg_agora),1,0))))'),
    "noprazo": '=IF(AND(ISNUMBER({concl}),ISNUMBER({prazo})),IF({concl}<={prazo},1,0),"")',
    "key": ('=IF({data}="","",{data}&"|"&{turno}&"|"&{prog}&"|"&LEFT({rota},40)&"|"&{plan}&"|"&{exec}'
            '&"|"&{transp}&"|"&LEFT({obs},30))'),
    "dup": '=IF({key}="",0,IF(COUNTIF(b_key,{key})>1,1,0))',
    "fdash": ('=IF({data}="",0,IF(AND({data}>=d_ini,{data}<=d_fim,OR(f_turno="TODOS",{turno}=f_turno),'
              'OR(f_resp="TODOS",{resp}=f_resp),OR(f_prog="TODOS",{prog}=f_prog),'
              'OR(f_rota="",ISNUMBER(SEARCH(f_rota,{rota}))),OR(f_status="TODOS",{status}=f_status),'
              'OR(f_tipo="TODOS",{tipo}=f_tipo),OR(f_imp="TODOS",{imp}=f_imp),'
              'OR(f_fat="TODOS",{fat}=f_fat),OR(f_emb="TODOS",{emb}=f_emb)),1,0))'),
    "fprod": ('=IF({data}="",0,IF(AND({med}=pr_unid,OR(pr_turno="TODOS",{turno}=pr_turno),'
              'OR(pr_resp="TODOS",{resp}=pr_resp),OR(pr_prog="TODOS",{prog}=pr_prog)),1,0))'),
    "sc_prio": ('=IF({abat}=1,({status}="CRÍTICO")*1E12+{venc}*1E11'
                '+IF(ISNUMBER({prazo}),ROUND((60000-{prazo})*24,0)*100000,0)+{impn}*10000+10000-ROW(),"")'),
    "sc_agora": '=IF(AND({abat}=1,OR({status}="CRÍTICO",{status}="ATENÇÃO")),{sc_prio},"")',
    "sc_aten": '=IF(AND({abat}=1,{status}="ATENÇÃO"),{sc_prio},"")',
    "sc_crit": '=IF(AND({abat}=1,{status}="CRÍTICO"),{sc_prio},"")',
    "sc_dia": '=IF(AND({abat}=1,{data}=gd_data),{sc_prio},"")',
    "sc_pconc": '=IF(AND({data}=pt_data,{turno}=pt_turno,{status}="CONCLUÍDO"),10000-ROW(),"")',
    "sc_ppend": '=IF(AND({abat}=1,{data}=pt_data,{turno}=pt_turno),{sc_prio},"")',
    "sc_hoje": '=IF(AND({abat}=1,ISNUMBER({prazo})),IF(INT({prazo})=INT(cfg_agora),{sc_prio},""),"")',
    "sc_fat": '=IF(AND({regra1}=1,{ativo}=1),(60000-{data})*10000+10000-ROW(),"")',
    "sc_emb": '=IF({embatr}=1,(60000-{data})*10000+10000-ROW(),"")',
    "sc_hist": ('=IF({data}="","",IF(AND({data}>=h_ini,{data}<=h_fim,OR(h_turno="TODOS",{turno}=h_turno),'
                'OR(h_resp="TODOS",{resp}=h_resp),OR(h_prog="TODOS",{prog}=h_prog),'
                'OR(h_rota="",ISNUMBER(SEARCH(h_rota,{rota}))),OR(h_tipo="TODOS",{tipo}=h_tipo),'
                'OR(h_causa="TODOS",{causa}=h_causa),OR(h_status="TODOS",{status}=h_status),'
                'OR(h_texto="",ISNUMBER(SEARCH(h_texto,{pend}&" "&{obs}&" "&{rota})))),'
                '{data}*10000+10000-ROW(),""))'),
    "sc_qual": '=IF(OR({valid}<>"",{qual}<>""),{data}*10000+10000-ROW(),"")',
}


def base_formula(key, r):
    refs = {k: f"{COL[k]}{r}" for k in COL}
    return F[key].format(**refs)


# ---------------------------------------------------------------- helpers de estilo
def st(c, bold=False, size=10, color="000000", bg=None, h=None, v="center", wrap=False, fmt=None, italic=False):
    c.font = Font(name=FONT, bold=bold, size=size, color=color, italic=italic)
    if bg:
        c.fill = fill(bg)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if fmt:
        c.number_format = fmt
    return c


def put(ws, ref, value, **kw):
    c = ws[ref]
    c.value = value
    st(c, **kw)
    return c


def banner(ws, title, subtitle, last_col=14):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_col)
    put(ws, "B1", title, bold=True, size=16, color=WHITE, bg=NAVY, h="left")
    for c in range(2, last_col + 1):
        ws.cell(row=1, column=c).fill = fill(NAVY)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)
    put(ws, "B2", subtitle, size=9, color=GREY, italic=True, h="left")
    ws.column_dimensions["A"].width = 2


def section(ws, row, col, text, span=8, color=BLUE):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=text)
    st(c, bold=True, size=11, color=WHITE, bg=color, h="left")
    for i in range(span):
        ws.cell(row=row, column=col + i).fill = fill(color)
    ws.row_dimensions[row].height = 20


def header_row(ws, row, col, labels, bg=NAVY):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        st(c, bold=True, size=9, color=WHITE, bg=bg, h="center", wrap=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def cell(ws, row, col, value, fmt=None, bold=False, h=None, bg=None, color="000000", wrap=False, size=10):
    c = ws.cell(row=row, column=col, value=value)
    st(c, bold=bold, fmt=fmt, h=h, bg=bg, color=color, wrap=wrap, size=size)
    c.border = BORDER
    return c


def input_cell(ws, ref, value=None, fmt=None, h="center"):
    c = ws[ref]
    if value is not None:
        c.value = value
    st(c, bold=True, bg=LYEL, fmt=fmt, h=h)
    c.border = Border(left=Side(style="thin", color=YEL), right=Side(style="thin", color=YEL),
                      top=Side(style="thin", color=YEL), bottom=Side(style="thin", color=YEL))
    c.protection = UNLOCK
    return c


def kpi(ws, row, col, label, formula, fmt="#,##0", span=2, color=NAVY, name=None):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + span - 1)
    c1 = ws.cell(row=row, column=col, value=label)
    st(c1, bold=True, size=8, color=GREY, bg=LGREY, h="center", wrap=True)
    c2 = ws.cell(row=row + 1, column=col, value=formula)
    st(c2, bold=True, size=18, color=color, bg=LGREY, h="center", fmt=fmt)
    for r in (row, row + 1):
        for i in range(span):
            x = ws.cell(row=r, column=col + i)
            x.fill = fill(LGREY)
            x.border = Border(left=Side(style="thin", color=WHITE) if i == 0 else None,
                              top=Side(style="medium", color=color) if r == row else None)
    ws.row_dimensions[row].height = 24
    ws.row_dimensions[row + 1].height = 32
    return c2


def dv_list(ws, source, ref, allow_blank=True, strict=True, prompt=None):
    dv = DataValidation(type="list", formula1=f"={source}", allow_blank=allow_blank)
    dv.error = "Selecione um valor da lista (cadastros na aba CADASTROS)."
    dv.errorTitle = "Valor fora do padrão"
    dv.errorStyle = "stop" if strict else "warning"
    if prompt:
        dv.prompt, dv.promptTitle = prompt, "Preenchimento"
    ws.add_data_validation(dv)
    dv.add(ref)
    return dv


def dv_date(ws, ref, prompt=None, datetime=False):
    dv = DataValidation(type="decimal" if datetime else "date", operator="between",
                        formula1="45658", formula2="47848", allow_blank=True)  # 01/01/2025 a 31/12/2030
    dv.error = "Informe uma data válida (dd/mm/aaaa" + (" hh:mm" if datetime else "") + ") entre 2025 e 2030."
    dv.errorTitle = "Data inválida"
    if prompt:
        dv.prompt, dv.promptTitle = prompt, "Preenchimento"
    ws.add_data_validation(dv)
    dv.add(ref)


def dv_num(ws, ref, prompt=None):
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv.error, dv.errorTitle = "Informe um número maior ou igual a zero.", "Número inválido"
    if prompt:
        dv.prompt, dv.promptTitle = prompt, "Preenchimento"
    ws.add_data_validation(dv)
    dv.add(ref)


def protect(ws):
    ws.protection.sheet = True
    ws.protection.password = SENHA
    ws.protection.autoFilter = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.sort = True
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False


def status_cf(ws, rng, first_cell):
    rules = [("CRÍTICO", RED, WHITE), ("ATENÇÃO", YEL, "000000"), ("PENDENTE", AMBER, "000000"),
             ("EM ANDAMENTO", LBLUE, NAVY), ("CONCLUÍDO", LGREEN, GREEN), ("CANCELADO", MGREY, GREY)]
    for val, bg, fg in rules:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'{first_cell}="{val}"'], fill=fill(bg), font=Font(name=FONT, color=fg, bold=True)))


def semaforo_cf(ws, rng, first_cell):
    """Colore pelo texto do alerta (🔴/🟡/🟢)."""
    for emo, bg, fg in (("🔴", LRED, RED), ("🟡", "FFF7D6", "7F6000"), ("🟢", LGREEN, GREEN)):
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'LEFT({first_cell},2)="{emo}"'], fill=fill(bg), font=Font(name=FONT, color=fg)))


def prazo_cf(ws, rng, first_cell):
    for val, bg, fg in (("VENCIDO", LRED, RED), ("VENCE HOJE", "FFF7D6", "7F6000"),
                        ("NO PRAZO", LGREEN, GREEN), ("SEM PRAZO", LGREY, GREY)):
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'{first_cell}="{val}"'], fill=fill(bg), font=Font(name=FONT, color=fg, bold=True)))


def pct_cf(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first_cell}),{first_cell}<cfg_meta)'], font=Font(name=FONT, color=RED, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND(ISNUMBER({first_cell}),{first_cell}>=cfg_meta)'], font=Font(name=FONT, color=GREEN, bold=True)))


def dtxt(ref):
    """Data como texto dd/mm/aaaa sem depender do idioma do Excel (TEXT() muda com a localidade)."""
    return f'RIGHT("0"&DAY({ref}),2)&"/"&RIGHT("0"&MONTH({ref}),2)&"/"&YEAR({ref})'


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def style_chart(ch, title, w=16, h=7.5):
    ch.title = title
    ch.width, ch.height = w, h
    ch.style = 10
    ch.legend.position = "b"
    try:
        ch.x_axis.delete = False
        ch.y_axis.delete = False
    except AttributeError:
        pass
    ch.y_axis.majorGridlines = None
    return ch


def color_series(ch, colors):
    for s, c in zip(ch.series, colors):
        s.graphicalProperties.solidFill = c
        s.graphicalProperties.line.solidFill = c


# ---------------------------------------------------------------- lista ordenada genérica
def ranked_list(ws, top, n, score, cols, start_col=2, title=None, span=None, color=BLUE, extra=None):
    """Lista ordenada pela coluna de pontuação `score` (nome definido b_sc_*).

    cols: [(cabeçalho, chave_base | ('f', template com {i}), formato, largura)]
    Coluna A (oculta) guarda a posição do registro na base.
    """
    row = top
    if title:
        section(ws, row, start_col, title, span or len(cols) + 1, color)
        row += 1
    header_row(ws, row, start_col, ["#"] + [c[0] for c in cols])
    first = row + 1
    for k in range(1, n + 1):
        r = row + k
        ws.cell(row=r, column=1, value=f'=IFERROR(MATCH(LARGE({score},{k}),{score},0),"")')
        cell(ws, r, start_col, f'=IF($A{r}="","",{k})', h="center", color=GREY, size=9)
        for j, (hdr, spec, fmt, _w) in enumerate(cols):
            if isinstance(spec, tuple):
                f = "=" + spec[1].format(i=f"$A{r}", r=r)
            elif spec == "id":
                f = (f'=IF($A{r}="","",HYPERLINK("#{Q(S_BASE)}!A"&($A{r}+1),INDEX(b_id,$A{r})))')
            elif spec in NUMERIC:
                f = f'=IF($A{r}="","",INDEX(b_{spec},$A{r}))'
            else:
                f = f'=IF($A{r}="","",INDEX(b_{spec},$A{r})&"")'
            c = cell(ws, r, start_col + 1 + j, f, fmt=fmt, size=9,
                     wrap=spec in ("pend", "obs", "acao", "valid", "qual") if not isinstance(spec, tuple) else False)
            if spec == "id":
                c.font = Font(name=FONT, size=9, color=BLUE, underline="single")
    last = row + n
    ws.column_dimensions["A"].hidden = True
    # formatação condicional por coluna
    for j, (hdr, spec, fmt, _w) in enumerate(cols):
        L = get_column_letter(start_col + 1 + j)
        rng = f"{L}{first}:{L}{last}"
        if spec == "status":
            status_cf(ws, rng, f"{L}{first}")
        elif spec == "alerta":
            semaforo_cf(ws, rng, f"{L}{first}")
        elif spec == "sitpz":
            prazo_cf(ws, rng, f"{L}{first}")
    return first, last


# textos de datas/números nas listas
FMT_D = "dd/mm/yyyy;;"
FMT_DT = "dd/mm/yy hh:mm;;"
FMT_N = "#,##0.0;-#,##0.0;;"
FMT_I = "0;-0;;"

LIST_STD = [  # colunas padrão de pendências
    ("ID", "id", None, 15), ("DATA", "data", FMT_D, 11), ("TURNO", "turno", None, 7),
    ("PROGRAMAÇÃO", "prog", None, 13), ("ROTA", "rota", None, 20), ("STATUS", "status", None, 13),
    ("ALERTA", "alerta", None, 26),
    ("PENDÊNCIA / OCORRÊNCIA", ("f", 'IF({i}="","",IF(INDEX(b_pend,{i})&""="",INDEX(b_obs,{i})&"",INDEX(b_pend,{i})&""))'), None, 36),
    ("RESP. TRATATIVA", "rtrat", None, 13), ("PRAZO", "prazo", FMT_DT, 14), ("SITUAÇÃO PRAZO", "sitpz", None, 12),
]


def widths(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# ================================================================ MAIN
def main():
    data = json.loads(SRC.read_text())
    recs, mig_log = data["registros"], data["log"]
    wb = Workbook()
    order = ["DASHBOARD", "GESTÃO_DO_DIA", "PASSAGEM_DE_TURNO", "LANÇAMENTO", S_BASE,
             "CONTROLE_PENDÊNCIAS", "FATURAMENTO", "EMBARQUES", "PRODUTIVIDADE", "GESTÃO_SEMANAL",
             "GESTÃO_MENSAL", "ANÁLISE_CAUSAS", "PARETO", "HISTÓRICO", "QUALIDADE_DADOS",
             "CADASTROS", "CONFIGURAÇÕES", "LOG_ALTERAÇÕES", "DIM_CALENDÁRIO", "COMO_USAR"]
    wb.active.title = order[0]
    for name in order[1:]:
        wb.create_sheet(name)
    ws_ = {n: wb[n] for n in order}
    tabcolors = {"DASHBOARD": NAVY, "GESTÃO_DO_DIA": NAVY, "PASSAGEM_DE_TURNO": NAVY, "LANÇAMENTO": YEL,
                 S_BASE: YEL, "CONTROLE_PENDÊNCIAS": RED, "CADASTROS": GREY, "CONFIGURAÇÕES": GREY,
                 "LOG_ALTERAÇÕES": GREY, "DIM_CALENDÁRIO": GREY, "COMO_USAR": GREEN}
    for n, w in ws_.items():
        w.sheet_properties.tabColor = tabcolors.get(n, BLUE)

    if "cadastros" not in SKIP:
        build_cadastros(wb, ws_["CADASTROS"], recs)
    if "config" not in SKIP:
        build_config(wb, ws_["CONFIGURAÇÕES"])
    if "base" not in SKIP:
        build_base(wb, ws_[S_BASE], recs)
    if "dashboard" not in SKIP:
        build_dashboard(wb, ws_["DASHBOARD"])
    if "dia" not in SKIP:
        build_dia(wb, ws_["GESTÃO_DO_DIA"])
    if "passagem" not in SKIP:
        build_passagem(wb, ws_["PASSAGEM_DE_TURNO"])
    if "lancamento" not in SKIP:
        build_lancamento(wb, ws_["LANÇAMENTO"])
    if "pendencias" not in SKIP:
        build_pendencias(wb, ws_["CONTROLE_PENDÊNCIAS"])
    if "faturamento" not in SKIP:
        build_faturamento(wb, ws_["FATURAMENTO"])
    if "embarques" not in SKIP:
        build_embarques(wb, ws_["EMBARQUES"])
    if "calendario" not in SKIP:
        build_calendario(wb, ws_["DIM_CALENDÁRIO"])
    if "produtividade" not in SKIP:
        build_produtividade(wb, ws_["PRODUTIVIDADE"])
    if "semanal" not in SKIP:
        build_semanal(wb, ws_["GESTÃO_SEMANAL"])
    if "mensal" not in SKIP:
        build_mensal(wb, ws_["GESTÃO_MENSAL"])
    if "causas" not in SKIP:
        build_causas(wb, ws_["ANÁLISE_CAUSAS"])
    if "pareto" not in SKIP:
        build_pareto(wb, ws_["PARETO"])
    if "historico" not in SKIP:
        build_historico(wb, ws_["HISTÓRICO"])
    if "qualidade" not in SKIP:
        build_qualidade(wb, ws_["QUALIDADE_DADOS"])
    if "log" not in SKIP:
        build_log(wb, ws_["LOG_ALTERAÇÕES"], recs, mig_log)
    if "como_usar" not in SKIP:
        build_como_usar(wb, ws_["COMO_USAR"])

    for n, w in ws_.items():
        if n not in ("COMO_USAR",):
            protect(w)
        w.sheet_view.zoomScale = 90
    ws_["COMO_USAR"].protection.sheet = True
    ws_["COMO_USAR"].protection.password = SENHA

    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"OK -> {OUT}  ({len(recs)} registros)")


# ================================================================ CADASTROS
CAD = {}  # nome -> (coluna, n_itens, dinâmico)


def build_cadastros(wb, ws, recs):
    banner(ws, "CADASTROS — listas padronizadas",
           "Células amarelas podem ser editadas/ampliadas. Todas as listas suspensas da planilha usam estas colunas. "
           "Não renomeie itens já usados na base (os registros antigos ficariam 'fora do padrão').", 30)
    transps = sorted({r["transp"] for r in recs if r["transp"]})
    lists = [
        ("turno", "TURNO", [t[0] for t in TURNOS], False, True),
        ("resp", "RESPONSÁVEIS", RESPS, True, True),
        ("prog", "PROGRAMAÇÃO", [p[0] for p in PROGS], True, True),
        ("area", "ÁREA (da programação)", [p[1] for p in PROGS], True, False),
        ("ativ", "TIPO DE ATIVIDADE", ATIVS, True, False),
        ("medida", "MEDIDA", MEDIDAS, False, False),
        ("status", "STATUS (automático)", STATUS, False, True),
        ("fat", "FATURAMENTO", FATS, False, True),
        ("emb", "EMBARQUE", EMBS, False, True),
        ("imp", "IMPACTO OPERACIONAL", [i[0] for i in IMPACTOS], False, True),
        ("imp_peso", "PESO DO IMPACTO", [i[1] for i in IMPACTOS], False, False),
        ("tipo", "TIPO DE PENDÊNCIA", TIPOS, True, True),
        ("causa", "CAUSA DO DESVIO", CAUSAS, True, True),
        ("transp", "TRANSPORTADORA", transps, True, False),
        ("simnao", "SIM / NÃO", ["SIM", "NÃO"], False, False),
        ("periodo", "PERÍODO (filtros)", PERIODOS, False, False),
    ]
    col = 2
    for key, title, items, dyn, with_todos in lists:
        L = get_column_letter(col)
        ws.column_dimensions[L].width = max(12, min(24, max(len(str(x)) for x in items + [title]) + 2))
        c = ws.cell(row=3, column=col, value=title)
        st(c, bold=True, size=9, color=WHITE, bg=NAVY, h="center", wrap=True)
        ws.row_dimensions[3].height = 32
        if with_todos:
            put(ws, f"{L}4", "TODOS", size=9, color=GREY, italic=True, h="center")
        cap = 40 if dyn else len(items)
        for i in range(cap):
            v = items[i] if i < len(items) else None
            c = ws.cell(row=5 + i, column=col, value=v)
            st(c, size=10, bg=LYEL if dyn else None)
            c.border = BORDER
            if dyn:
                c.protection = UNLOCK
        end = 5 + cap - 1
        if dyn:
            add_name(wb, f"l_{key}", f"OFFSET({Q('CADASTROS')}!${L}$5,0,0,MAX(1,COUNTA({Q('CADASTROS')}!${L}$5:${L}${end})),1)")
            if with_todos:
                add_name(wb, f"lf_{key}", f"OFFSET({Q('CADASTROS')}!${L}$4,0,0,COUNTA({Q('CADASTROS')}!${L}$5:${L}${end})+1,1)")
        else:
            add_name(wb, f"l_{key}", f"{Q('CADASTROS')}!${L}$5:${L}${end}")
            if with_todos:
                add_name(wb, f"lf_{key}", f"{Q('CADASTROS')}!${L}$4:${L}${end}")
        CAD[key] = (L, cap, len(items))
        col += 1 if key in ("prog", "imp") else 2
    # tabelas de apoio
    Lp, La = CAD["prog"][0], CAD["area"][0]
    add_name(wb, "t_prog", f"{Q('CADASTROS')}!${Lp}$5:${La}${4 + CAD['prog'][1]}")
    Li, Lw = CAD["imp"][0], CAD["imp_peso"][0]
    add_name(wb, "t_impacto", f"{Q('CADASTROS')}!${Li}$5:${Lw}$9")
    ws.freeze_panes = "A5"
    put(ws, "B47", "Obs.: ÁREA fica ao lado da PROGRAMAÇÃO; o PESO do impacto (0–4) é usado na matriz "
        "Frequência x Impacto e na priorização.", size=9, italic=True, color=GREY)


def cad_item(key, i):
    """Referência absoluta ao i-ésimo item (0-based) de um cadastro."""
    L = CAD[key][0]
    return f"{Q('CADASTROS')}!${L}${5 + i}"


# ================================================================ CONFIGURAÇÕES
def build_config(wb, ws):
    banner(ws, "CONFIGURAÇÕES — parâmetros das regras automáticas",
           "Altere apenas as células amarelas. Todas as regras de status, alertas, prazos e SLA usam estes parâmetros.", 6)
    widths(ws, {"B": 44, "C": 20, "D": 80})
    header_row(ws, 3, 2, ["PARÂMETRO", "VALOR", "COMO É USADO"])
    params = [
        ("Data/hora de referência (vazio = agora)", None, "dd/mm/yyyy hh:mm", None,
         "Deixe VAZIO no uso normal (usa a data/hora atual). Preencha só para simular/auditar uma data passada."),
        ("Data/hora de referência efetiva", "=IF(C4=\"\",NOW(),C4)", "dd/mm/yyyy hh:mm", "cfg_agora",
         "Calculado. Base para prazos vencidos, 'vence hoje', dias em aberto e SLA."),
        ("Data de corte da migração", dt.date(2026, 9, 20), "dd/mm/yyyy", "cfg_corte",
         "Registros MIGRADOS do CSV antigo com data anterior a esta ficam como HISTÓRICO: entram nos indicadores "
         "e análises, mas não geram pendência aberta, crítico ou alerta de próximo turno (não há prazo/responsável "
         "registrados para eles). Lançamentos novos nunca são afetados."),
        ("Tolerância de embarque sem prazo (dias)", 1, "0", "cfg_tol_emb",
         "Embarque PENDENTE/EM PROCESSO sem prazo informado fica ATRASADO após DATA + N dias."),
        ("Aviso de prazo próximo (horas)", 4, "0", "cfg_h_prox",
         "Itens abertos cujo prazo vence dentro de N horas ficam em ATENÇÃO."),
        ("Meta de atingimento (% execução)", 0.95, "0%", "cfg_meta",
         "Abaixo da meta o % aparece em vermelho; acima, em verde."),
        ("Limite de volume por registro (t)", 80, "#,##0", "cfg_lim_val",
         "Registros acima deste volume aparecem em QUALIDADE_DADOS como 'Valor acima do limite' (provável erro de digitação)."),
        ("Tolerância de divergência (executado > planejado)", 0.05, "0%", "cfg_tol_div",
         "Executado acima do planejado + tolerância gera DIVERGÊNCIA."),
        ("Impacto médio considerado ALTO (0–4)", 2.5, "0.0", "cfg_imp_alto",
         "Linha de corte do eixo IMPACTO na matriz Frequência x Impacto (peso médio: 2 = MÉDIO, 3 = ALTO)."),
        ("Ocorrências mínimas para recorrência", 3, "0", "cfg_rec_min",
         "Tipo/causa com N ou mais ocorrências no período é marcado como RECORRENTE."),
        ("Unidade de volume dos indicadores", "TON", "@", "cfg_unid",
         "Planejado/Executado/Faturado dos painéis somam apenas esta medida (paletes e posições não se somam a toneladas)."),
        ("Senha de proteção das abas", SENHA, "@", None,
         "Revisar > Desproteger planilha. Evita apagar fórmulas por engano (não é segurança forte)."),
    ]
    for i, (lab, val, fmt, name, how) in enumerate(params):
        r = 4 + i
        cell(ws, r, 2, lab, bold=True)
        c = cell(ws, r, 3, val, fmt=fmt, h="center")
        if name != "cfg_agora" and lab.startswith("Senha") is False:
            input_cell(ws, f"C{r}", fmt=fmt)
        cell(ws, r, 4, how, wrap=True, size=9)
        ws.row_dimensions[r].height = 42 if len(how) > 110 else 28
        if name:
            add_name(wb, name, f"{Q('CONFIGURAÇÕES')}!$C${r}")
    dv_date(ws, "C4", "Vazio = agora", datetime=True)
    dv = DataValidation(type="list", formula1="=l_medida")
    ws.add_data_validation(dv)
    dv.add("C14")

    # SLA por tipo de pendência
    r0 = 19
    section(ws, r0 - 1, 2, "SLA POR TIPO DE PENDÊNCIA (horas para resolver)", 3)
    header_row(ws, r0, 2, ["TIPO DE PENDÊNCIA", "SLA (h)", "OBS."])
    for i, t in enumerate(TIPOS):
        r = r0 + 1 + i
        cell(ws, r, 2, f"={cad_item('tipo', i)}")
        input_cell(ws, f"C{r}", SLA[t], "0")
        cell(ws, r, 4, "Sugestão inicial do briefing (Sistema 4h, Faturamento 2h, Embarque 1h, Documentação 2h, "
             "Equipamento 4h); demais tipos: valores de partida — ajuste à realidade." if i == 0 else None, size=9, wrap=True)
    add_name(wb, "t_sla", f"{Q('CONFIGURAÇÕES')}!$B${r0 + 1}:$C${r0 + len(TIPOS)}")

    # turnos
    r1 = r0 + len(TIPOS) + 3
    section(ws, r1 - 1, 2, "SEQUÊNCIA E HORÁRIO DOS TURNOS", 3)
    header_row(ws, r1, 2, ["TURNO", "PRÓXIMO TURNO", "HORA DE INÍCIO"])
    for i, (t, nxt, h) in enumerate(TURNOS):
        r = r1 + 1 + i
        cell(ws, r, 2, f"={cad_item('turno', i)}", h="center")
        input_cell(ws, f"C{r}", nxt)
        input_cell(ws, f"D{r}", h, "hh:mm")
    cell(ws, r1 + 6, 2, "Premissa: escala 12x36 (A/C diurnos 06:00, B/D noturnos 18:00), deduzida do padrão das datas "
         "no arquivo original. Ajuste se a escala for outra — impacta abertura das pendências, SLA e 'Próximo turno'.",
         size=9, wrap=True)
    ws.merge_cells(start_row=r1 + 6, start_column=2, end_row=r1 + 6, end_column=4)
    ws.row_dimensions[r1 + 6].height = 30
    add_name(wb, "t_turno", f"{Q('CONFIGURAÇÕES')}!$B${r1 + 1}:$D${r1 + 4}")

    # headcount
    r2 = r1 + 9
    section(ws, r2 - 1, 2, "COLABORADORES POR TURNO (para produtividade por pessoa — opcional)", 3)
    header_row(ws, r2, 2, ["TURNO", "Nº COLABORADORES", "OBS."])
    for i in range(4):
        r = r2 + 1 + i
        cell(ws, r, 2, f"={cad_item('turno', i)}", h="center")
        input_cell(ws, f"C{r}", None, "0")
    cell(ws, r2 + 1, 4, "Deixe vazio se não houver a informação; a média por colaborador fica em branco.", size=9)
    add_name(wb, "t_headcount", f"{Q('CONFIGURAÇÕES')}!$B${r2 + 1}:$C${r2 + 4}")


# ================================================================ BASE
def build_base(wb, ws, recs):
    ncol = len(COLS)
    for i, (k, hdr, kind, w, fmt) in enumerate(COLS):
        L = get_column_letter(i + 1)
        c = ws.cell(row=1, column=i + 1, value=hdr)
        bg = {"in": NAVY, "auto": GREY, "aux": MGREY}[kind]
        st(c, bold=True, size=9, color=WHITE if kind != "aux" else GREY, bg=bg, h="center", wrap=True)
        ws.column_dimensions[L].width = w
        add_name(wb, f"b_{k}", f"{Q(S_BASE)}!${L}${R1}:${L}${RN}")
        if kind == "aux":
            ws.column_dimensions[L].outlineLevel = 1
            ws.column_dimensions[L].hidden = True
    ws.row_dimensions[1].height = 42
    comments = {
        "id": "Automático (TT-ano-sequência). Não digitar.",
        "data": "Data do turno (dd/mm/aaaa).",
        "plan": "Planejado na MEDIDA escolhida. TON = toneladas (ex.: 28,054 = 28.054 kg).",
        "status": "Automático: CRÍTICO / ATENÇÃO / PENDENTE / EM ANDAMENTO / CONCLUÍDO / CANCELADO. Regras em COMO_USAR.",
        "alerta": "Automático: semáforo 🔴🟡🟢 + motivo principal.",
        "pend": "Descreva o que ficou pendente. Ao preencher, informe também TIPO, RESPONSÁVEL e PRAZO.",
        "prazo": "Data e hora limite (dd/mm/aaaa hh:mm).",
        "concl": "Preencha quando a pendência for resolvida (data e hora).",
        "canc": "SIM para cancelar o registro (não apagar linhas).",
        "origem": "MIGRAÇÃO = veio do CSV antigo. Vazio = lançamento novo.",
        "atual": "Preenchida automaticamente pela macro (opcional). Sem macro: Ctrl+; e Ctrl+Shift+;",
    }
    for k, txt in comments.items():
        ws[f"{COL[k]}1"].comment = Comment(txt, "Sistema")

    # estilos por coluna (aplicados célula a célula — openpyxl não herda estilo de coluna na escrita)
    font = Font(name=FONT, size=9)
    font_auto = Font(name=FONT, size=9, color="404040")
    in_fill = fill("FFFBEA")
    for r in range(R1, RN + 1):
        rec = recs[r - R1] if r - R1 < len(recs) else None
        for i, (k, hdr, kind, w, fmt) in enumerate(COLS):
            c = ws.cell(row=r, column=i + 1)
            if kind == "in":
                if rec:
                    v = {
                        "data": dt.date.fromisoformat(rec["data"]) if rec["data"] else None,
                        "turno": rec["turno"] or None, "resp": rec["resp"] or None, "ativ": rec["tipo_ativ"] or None,
                        "prog": rec["prog"] or None, "rota": rec["rota"] or None, "med": rec["medida"] or None,
                        "plan": rec["plan"], "exec": rec["exec"], "fat": rec["fat"] or None, "emb": rec["emb"] or None,
                        "volfat": rec["vol_fat"], "transp": rec["transp"] or None, "pend": rec["pend"] or None,
                        "tipo": rec["tipo_pend"] or None, "causa": rec["causa"] or None, "obs": rec["obs"] or None,
                    }.get(k)
                    c.value = v
                c.fill = in_fill
                c.protection = UNLOCK
                c.font = font
            elif k == "origem":
                c.value = "MIGRAÇÃO" if rec else None
                c.font = font_auto
            else:
                c.value = base_formula(k, r)
                c.font = font_auto
            if fmt != "@":
                c.number_format = fmt
    # validações (uma por coluna)
    rng = lambda k: f"{COL[k]}{R1}:{COL[k]}{RN}"
    dv_date(ws, rng("data"), "Data do turno (dd/mm/aaaa)")
    dv_list(ws, "l_turno", rng("turno"), prompt="Turno A, B, C ou D")
    dv_list(ws, "l_resp", rng("resp"), strict=False, prompt="Líder do turno (cadastro em CADASTROS)")
    dv_list(ws, "l_ativ", rng("ativ"))
    dv_list(ws, "l_prog", rng("prog"))
    dv_list(ws, "l_medida", rng("med"), prompt="TON = toneladas; PALETE; POSIÇÃO")
    dv_num(ws, rng("plan"), "Volume planejado (na medida escolhida)")
    dv_num(ws, rng("exec"), "Volume executado (na medida escolhida)")
    dv_num(ws, rng("volfat"), "Volume faturado (na medida escolhida)")
    dv_list(ws, "l_fat", rng("fat"))
    dv_list(ws, "l_emb", rng("emb"))
    dv_list(ws, "l_transp", rng("transp"), strict=False)
    dv_list(ws, "l_tipo", rng("tipo"), prompt="Obrigatório quando houver PENDÊNCIA")
    dv_list(ws, "l_causa", rng("causa"))
    dv_list(ws, "l_resp", rng("rtrat"), strict=False, prompt="Obrigatório quando houver PENDÊNCIA")
    dv_date(ws, rng("prazo"), "Obrigatório quando houver PENDÊNCIA (dd/mm/aaaa hh:mm)", datetime=True)
    dv_date(ws, rng("concl"), "Data/hora em que a pendência foi resolvida", datetime=True)
    dv_date(ws, rng("atual"), None, datetime=True)
    dv_list(ws, "l_imp", rng("imp"))
    dv_list(ws, "l_simnao", rng("canc"))

    # formatação condicional
    status_cf(ws, rng("status"), f"{COL['status']}{R1}")
    semaforo_cf(ws, rng("alerta"), f"{COL['alerta']}{R1}")
    prazo_cf(ws, rng("sitpz"), f"{COL['sitpz']}{R1}")
    pct_cf(ws, rng("pct"), f"{COL['pct']}{R1}")
    a = f"{COL['impacta']}{R1}"
    ws.conditional_formatting.add(rng("impacta"), FormulaRule(
        formula=[f'AND({a}="SIM",${COL["status"]}{R1}="CRÍTICO")'], fill=fill(LRED), font=Font(name=FONT, color=RED, bold=True)))
    ws.conditional_formatting.add(rng("impacta"), FormulaRule(
        formula=[f'{a}="SIM"'], fill=fill("FFF7D6"), font=Font(name=FONT, color="7F6000", bold=True)))
    for k in ("valid", "qual"):
        ws.conditional_formatting.add(rng(k), FormulaRule(
            formula=[f'{COL[k]}{R1}<>""'], fill=fill(LRED), font=Font(name=FONT, color=RED)))
    ws.conditional_formatting.add(rng("slaest"), CellIsRule(operator="equal", formula=['"SIM"'],
                                                            font=Font(name=FONT, color=RED, bold=True)))
    # campos obrigatórios da pendência em amarelo forte quando faltarem
    for k in ("tipo", "rtrat", "prazo"):
        ws.conditional_formatting.add(rng(k), FormulaRule(
            formula=[f'AND(${COL["pend"]}{R1}<>"",{COL[k]}{R1}="",${COL["aberto"]}{R1}=1)'], fill=fill(YEL)))

    last = get_column_letter(ncol)
    tab = Table(displayName="tbBase", ref=f"A1:{last}{RN}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(tab)
    ws.freeze_panes = "C2"
    ws.sheet_properties.outlinePr.summaryRight = False


# ================================================================ DASHBOARD
def build_dashboard(wb, ws):
    banner(ws, "TROCA DE TURNO 2026 — PAINEL EXECUTIVO",
           "Indicadores 100% automáticos a partir da BASE_OPERACIONAL. Use os filtros amarelos (TODOS = sem filtro).", 17)
    ws["B2"].value = ('="Atualizado em "&' + dtxt("cfg_agora") + '&"  |  Período: "&' + dtxt("d_ini_v") +
                      '&" a "&' + dtxt("d_fim_v") + '&"  |  Filtros amarelos: TODOS = sem filtro"')
    for c in range(2, 18):
        ws.column_dimensions[get_column_letter(c)].width = 11.5
    filters = [("PERÍODO", "f_periodo", "l_periodo", "MÊS ATUAL"), ("DATA INICIAL", "f_dini", None, None),
               ("DATA FINAL", "f_dfim", None, None), ("TURNO", "f_turno", "lf_turno", "TODOS"),
               ("RESPONSÁVEL", "f_resp", "lf_resp", "TODOS"), ("PROGRAMAÇÃO", "f_prog", "lf_prog", "TODOS"),
               ("ROTA (contém)", "f_rota", None, None), ("STATUS", "f_status", "lf_status", "TODOS"),
               ("TIPO PENDÊNCIA", "f_tipo", "lf_tipo", "TODOS"), ("IMPACTO", "f_imp", "lf_imp", "TODOS"),
               ("FATURAMENTO", "f_fat", "lf_fat", "TODOS"), ("EMBARQUE", "f_emb", "lf_emb", "TODOS")]
    for i, (lab, name, lst, default) in enumerate(filters):
        col = 2 + i
        L = get_column_letter(col)
        put(ws, f"{L}4", lab, bold=True, size=8, color=GREY, h="center", wrap=True)
        input_cell(ws, f"{L}5", default, "dd/mm/yyyy" if "DATA" in lab else None)
        add_name(wb, name, f"{Q('DASHBOARD')}!${L}$5")
        if lst:
            dv_list(ws, lst, f"{L}5")
        elif "DATA" in lab:
            dv_date(ws, f"{L}5", "Usado quando PERÍODO = PERSONALIZADO")
    ws.row_dimensions[4].height = 24
    # período efetivo (células auxiliares visíveis em cinza, linha 6)
    put(ws, "N4", "Período aplicado", bold=True, size=8, color=GREY, h="center")
    ws["N5"] = ('=IF(f_periodo="PERSONALIZADO",IF(f_dini="",DATE(2000,1,1),f_dini),IF(f_periodo="MÊS ATUAL",'
                'DATE(YEAR(cfg_agora),MONTH(cfg_agora),1),IF(f_periodo="SEMANA ATUAL",INT(cfg_agora)-WEEKDAY(cfg_agora,2)+1,'
                'IF(f_periodo="ÚLTIMOS 7 DIAS",INT(cfg_agora)-6,IF(f_periodo="ÚLTIMOS 30 DIAS",INT(cfg_agora)-29,'
                'IF(f_periodo="HOJE",INT(cfg_agora),DATE(2000,1,1)))))))')
    ws["O5"] = ('=IF(f_periodo="PERSONALIZADO",IF(f_dfim="",DATE(2100,12,31),f_dfim),IF(f_periodo="MÊS ATUAL",'
                'EOMONTH(cfg_agora,0),IF(f_periodo="TODO O HISTÓRICO",DATE(2100,12,31),INT(cfg_agora))))')
    for ref in ("N5", "O5"):
        st(ws[ref], size=9, color=GREY, h="center", fmt="dd/mm/yyyy")
    add_name(wb, "d_ini", f"{Q('DASHBOARD')}!$N$5")
    add_name(wb, "d_fim", f"{Q('DASHBOARD')}!$O$5")
    ws["P5"] = '=IF(N5=DATE(2000,1,1),MIN(b_data),N5)'
    ws["Q5"] = '=IF(O5=DATE(2100,12,31),MAX(b_data),O5)'
    for ref in ("P5", "Q5"):
        st(ws[ref], size=9, color=GREY, h="center", fmt="dd/mm/yyyy")
    put(ws, "P4", "(exibição)", size=8, color=GREY, h="center")
    add_name(wb, "d_ini_v", f"{Q('DASHBOARD')}!$P$5")
    add_name(wb, "d_fim_v", f"{Q('DASHBOARD')}!$Q$5")

    U = "b_med,cfg_unid"
    kpis = [
        ("TOTAL DE ATIVIDADES", "=SUM(b_fdash)", "#,##0", NAVY),
        ("PLANEJADO (t)", f"=SUMIFS(b_plan,b_fdash,1,{U})", "#,##0.0", NAVY),
        ("EXECUTADO (t)", f"=SUMIFS(b_exec,b_fdash,1,{U})", "#,##0.0", NAVY),
        ("% EXECUÇÃO", "=IFERROR(E10/D10,0)", "0.0%", NAVY),
        ("CONCLUÍDOS", '=COUNTIFS(b_fdash,1,b_status,"CONCLUÍDO")', "#,##0", GREEN),
        ("EM ANDAMENTO", '=COUNTIFS(b_fdash,1,b_status,"EM ANDAMENTO")', "#,##0", BLUE),
        ("PENDENTES", '=COUNTIFS(b_fdash,1,b_status,"PENDENTE")', "#,##0", "7F6000"),
        ("ATENÇÃO", '=COUNTIFS(b_fdash,1,b_status,"ATENÇÃO")', "#,##0", "7F6000"),
        ("CRÍTICOS", '=COUNTIFS(b_fdash,1,b_status,"CRÍTICO")', "#,##0", RED),
        ("PENDÊNCIAS VENCIDAS", "=COUNTIFS(b_fdash,1,b_venc,1)", "#,##0", RED),
        ("FATURAMENTO PENDENTE", '=COUNTIFS(b_fdash,1,b_fat,"PENDENTE")+COUNTIFS(b_fdash,1,b_fat,"EM PROCESSO")', "#,##0", "7F6000"),
        ("EMBARQUES PENDENTES", '=COUNTIFS(b_fdash,1,b_emb,"PENDENTE")+COUNTIFS(b_fdash,1,b_emb,"EM PROCESSO")', "#,##0", "7F6000"),
    ]
    # 2 linhas de 6 cartões (cada cartão = 2 colunas + 1 de respiro não usado: 6 x 2 = 12 colunas B..M)
    kcells = {}
    for i, (lab, f, fmt, colr) in enumerate(kpis):
        row = 9 if i < 6 else 12
        col = 2 + (i % 6) * 2 + (i % 6) // 3
        kcells[lab] = kpi(ws, row, col, lab, None, fmt, 2, colr)
    # fórmulas (a de % depende das células de planejado/executado)
    for i, (lab, f, fmt, colr) in enumerate(kpis):
        c = kcells[lab]
        if lab == "% EXECUÇÃO":
            f = f"=IFERROR({kcells['EXECUTADO (t)'].coordinate}/{kcells['PLANEJADO (t)'].coordinate},0)"
        c.value = f
    pct_ref = kcells["% EXECUÇÃO"].coordinate
    pct_cf(ws, pct_ref, pct_ref)

    # comparativo por turno
    r0 = 15
    section(ws, r0, 2, "COMPARATIVO POR TURNO", 7)
    header_row(ws, r0 + 1, 2, ["INDICADOR", "", "A", "B", "C", "D", "TOTAL"])
    ws.merge_cells(start_row=r0 + 1, start_column=2, end_row=r0 + 1, end_column=3)
    rows = [
        ("Planejado (t)", "SUMIFS(b_plan,b_fdash,1,b_turno,{t},b_med,cfg_unid)", "#,##0.0"),
        ("Executado (t)", "SUMIFS(b_exec,b_fdash,1,b_turno,{t},b_med,cfg_unid)", "#,##0.0"),
        ("% execução", "IFERROR({e}/{p},\"\")", "0.0%"),
        ("Itens em aberto", "COUNTIFS(b_fdash,1,b_turno,{t},b_abat,1)", "#,##0"),
        ("Críticos", "COUNTIFS(b_fdash,1,b_turno,{t},b_status,\"CRÍTICO\")", "#,##0"),
        ("Faturado (t)", "SUMIFS(b_volfat,b_fdash,1,b_turno,{t},b_med,cfg_unid)", "#,##0.0"),
        ("Embarques realizados", "COUNTIFS(b_fdash,1,b_turno,{t},b_emb,\"REALIZADO\")", "#,##0"),
        ("Registros", "COUNTIFS(b_fdash,1,b_turno,{t})", "#,##0"),
    ]
    for j, (lab, f, fmt) in enumerate(rows):
        r = r0 + 2 + j
        cell(ws, r, 2, lab, bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        for k in range(5):
            col = 4 + k
            L = get_column_letter(col)
            if k < 4:
                t = f"{L}${r0 + 1}"
                ff = f.format(t=t, e=f"{L}{r0 + 3}", p=f"{L}{r0 + 2}")
            else:
                ff = f"SUM(D{r}:G{r})" if lab != "% execução" else f'IFERROR(H{r0 + 3}/H{r0 + 2},"")'
            cell(ws, r, col, "=" + ff, fmt=fmt, h="center", bold=(k == 4))
    pct_cf(ws, f"D{r0 + 4}:H{r0 + 4}", f"D{r0 + 4}")
    ch = BarChart()
    ch.type = "col"
    from openpyxl.chart import Series
    for rr, colr in ((r0 + 2, MGREY), (r0 + 3, NAVY)):
        s = Series(Reference(ws, min_col=4, max_col=7, min_row=rr, max_row=rr), title_from_data=False,
                   title=ws.cell(row=rr, column=2).value)
        s.graphicalProperties.solidFill = colr
        s.graphicalProperties.line.solidFill = colr
        ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=4, max_col=7, min_row=r0 + 1, max_row=r0 + 1))
    style_chart(ch, "Planejado x Executado por turno (t)", 15, 7.2)
    ws.add_chart(ch, f"J{r0}")

    # status
    r1 = r0 + 12
    section(ws, r1, 2, "STATUS DAS ATIVIDADES", 7)
    header_row(ws, r1 + 1, 2, ["STATUS", "", "QTDE", "%"])
    ws.merge_cells(start_row=r1 + 1, start_column=2, end_row=r1 + 1, end_column=3)
    for j, s in enumerate(STATUS):
        r = r1 + 2 + j
        cell(ws, r, 2, s, bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cell(ws, r, 4, f'=COUNTIFS(b_fdash,1,b_status,B{r})', fmt="#,##0", h="center")
        cell(ws, r, 5, f'=IFERROR(D{r}/SUM($D${r1 + 2}:$D${r1 + 7}),0)', fmt="0.0%", h="center")
    status_cf(ws, f"B{r1 + 2}:B{r1 + 7}", f"B{r1 + 2}")
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=4, min_row=r1 + 2, max_row=r1 + 7), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=2, min_row=r1 + 2, max_row=r1 + 7))
    style_chart(ch, "Atividades por status", 15, 6.5)
    ch.legend = None
    ch.x_axis.scaling.orientation = "maxMin"
    s = ch.series[0]
    from openpyxl.chart.marker import DataPoint
    for idx, colr in enumerate([RED, YEL, AMBER, BLUE, GREEN, "A6A6A6"]):
        pt = DataPoint(idx=idx)
        pt.graphicalProperties.solidFill = colr
        s.dPt.append(pt)
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    ws.add_chart(ch, f"J{r1}")

    # pendências por tipo
    r2 = r1 + 14
    section(ws, r2, 2, "PENDÊNCIAS (OCORRÊNCIAS) POR TIPO", 7)
    header_row(ws, r2 + 1, 2, ["TIPO", "", "OCORRÊNCIAS", "EM ABERTO", "VENCIDAS"])
    ws.merge_cells(start_row=r2 + 1, start_column=2, end_row=r2 + 1, end_column=3)
    for j in range(len(TIPOS)):
        r = r2 + 2 + j
        cell(ws, r, 2, f"={cad_item('tipo', j)}", bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cell(ws, r, 4, f"=COUNTIFS(b_fdash,1,b_tipo,B{r},b_ocorr,1)", fmt="#,##0", h="center")
        cell(ws, r, 5, f"=COUNTIFS(b_fdash,1,b_tipo,B{r},b_ocorr,1,b_abat,1)", fmt="#,##0", h="center")
        cell(ws, r, 6, f"=COUNTIFS(b_fdash,1,b_tipo,B{r},b_venc,1)", fmt="#,##0", h="center")
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=4, max_col=5, min_row=r2 + 1, max_row=r2 + 1 + len(TIPOS)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r2 + 2, max_row=r2 + 1 + len(TIPOS)))
    style_chart(ch, "Pendências por tipo", 15, 8.5)
    ch.x_axis.scaling.orientation = "maxMin"
    color_series(ch, [NAVY, RED])
    ws.add_chart(ch, f"J{r2}")

    # prioridades
    r3 = r2 + len(TIPOS) + 3
    ranked_list(ws, r3, 10, "b_sc_prio", [
        ("ID", "id", None, 0), ("DATA", "data", FMT_D, 0), ("TURNO", "turno", None, 0),
        ("PROGRAMAÇÃO", "prog", None, 0), ("STATUS", "status", None, 0), ("ALERTA", "alerta", None, 0),
        ("RESP. TRATATIVA", "rtrat", None, 0), ("PRAZO", "prazo", FMT_DT, 0),
    ], title="TOP 10 PRIORIDADES EM ABERTO (visão geral, sem filtro) — detalhes em CONTROLE_PENDÊNCIAS", span=16, color=RED)
    ws.freeze_panes = "A7"


# ================================================================ GESTÃO DO DIA
def turno_table(ws, r0, crit, title, rows_extra=None):
    """Tabela Indicador x Turno (A..D + TOTAL). crit = critérios SUMIFS/COUNTIFS extras (string)."""
    section(ws, r0, 2, title, 7)
    header_row(ws, r0 + 1, 2, ["INDICADOR", "", "A", "B", "C", "D", "TOTAL"])
    ws.merge_cells(start_row=r0 + 1, start_column=2, end_row=r0 + 1, end_column=3)
    rows = [
        ("Planejado (t)", f"SUMIFS(b_plan,{crit},b_turno,{{t}},b_med,cfg_unid)", "#,##0.0", "sum"),
        ("Executado (t)", f"SUMIFS(b_exec,{crit},b_turno,{{t}},b_med,cfg_unid)", "#,##0.0", "sum"),
        ("% execução", 'IFERROR({e}/{p},"")', "0.0%", "pct"),
        ("Registros", f"COUNTIFS({crit},b_turno,{{t}})", "#,##0", "sum"),
        ("Itens em aberto", f"COUNTIFS({crit},b_turno,{{t}},b_abat,1)", "#,##0", "sum"),
        ("Críticos", f'COUNTIFS({crit},b_turno,{{t}},b_status,"CRÍTICO")', "#,##0", "sum"),
        ("Faturado (t)", f"SUMIFS(b_volfat,{crit},b_turno,{{t}},b_med,cfg_unid)", "#,##0.0", "sum"),
        ("Faturamento pendente (itens)", f'COUNTIFS({crit},b_turno,{{t}},b_fat,"PENDENTE")+COUNTIFS({crit},b_turno,{{t}},b_fat,"EM PROCESSO")', "#,##0", "sum"),
        ("Embarques realizados", f'COUNTIFS({crit},b_turno,{{t}},b_emb,"REALIZADO")', "#,##0", "sum"),
        ("Embarques pendentes", f'COUNTIFS({crit},b_turno,{{t}},b_emb,"PENDENTE")+COUNTIFS({crit},b_turno,{{t}},b_emb,"EM PROCESSO")', "#,##0", "sum"),
    ]
    for j, (lab, f, fmt, agg) in enumerate(rows):
        r = r0 + 2 + j
        cell(ws, r, 2, lab, bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        for k in range(5):
            col = 4 + k
            L = get_column_letter(col)
            if k < 4:
                ff = f.format(t=f"{L}${r0 + 1}", e=f"{L}{r0 + 3}", p=f"{L}{r0 + 2}")
            else:
                ff = f"SUM(D{r}:G{r})" if agg == "sum" else f'IFERROR(H{r0 + 3}/H{r0 + 2},"")'
            cell(ws, r, col, "=" + ff, fmt=fmt, h="center", bold=(k == 4))
    pct_cf(ws, f"D{r0 + 4}:H{r0 + 4}", f"D{r0 + 4}")
    return r0 + 2 + len(rows)


def build_dia(wb, ws):
    banner(ws, "GESTÃO DO DIA", "O que aconteceu hoje, o que precisa ser tratado agora e o que fica para o próximo turno.", 14)
    put(ws, "B4", "Data de referência:", bold=True, h="right")
    ws.merge_cells("B4:C4")
    input_cell(ws, "D4", None, "dd/mm/yyyy")
    dv_date(ws, "D4", "Vazio = último dia com lançamentos até hoje")
    put(ws, "E4", "(vazio = último dia com lançamentos)", size=8, color=GREY, italic=True)
    put(ws, "H4", "Exibindo:", bold=True, h="right")
    ws["I4"] = '=IF(D4="",_xlfn.MAXIFS(b_data,b_data,"<="&INT(cfg_agora)),D4)'
    st(ws["I4"], bold=True, size=12, color=NAVY, fmt="dd/mm/yyyy")
    add_name(wb, "gd_data", f"{Q('GESTÃO_DO_DIA')}!$I$4")
    end = turno_table(ws, 6, "b_data,gd_data", "HOJE — POR TURNO")
    ranked_list(ws, end + 2, 15, "b_sc_agora", LIST_STD,
                title="⚠ O QUE PRECISA SER TRATADO AGORA (críticos e atenção em aberto, por prioridade)", span=12, color=RED)
    ranked_list(ws, end + 2 + 18, 15, "b_sc_dia", LIST_STD[:-1] + [("PRÓX. TURNO", "prox", None, 0)],
                title="➡ O QUE FICOU PARA O PRÓXIMO TURNO (itens do dia ainda em aberto)", span=12, color=BLUE)
    widths(ws, {"B": 5, "C": 15, "D": 11, "E": 7, "F": 13, "G": 20, "H": 13, "I": 28, "J": 40, "K": 14, "L": 14, "M": 12})
    ws.freeze_panes = "A6"


# ================================================================ PASSAGEM DE TURNO
def build_passagem(wb, ws):
    banner(ws, "PASSAGEM DE TURNO", "Tudo o que o próximo turno precisa saber — gerado automaticamente.", 14)
    put(ws, "B4", "Data:", bold=True, h="right")
    ws.merge_cells("B4:C4")
    input_cell(ws, "D4", None, "dd/mm/yyyy")
    dv_date(ws, "D4", "Vazio = último dia com lançamentos até hoje")
    put(ws, "E4", "Turno:", bold=True, h="right")
    input_cell(ws, "F4", None)
    dv_list(ws, "l_turno", "F4", prompt="Turno que está ENTREGANDO. Vazio = último turno lançado na data")
    put(ws, "G4", "(vazios = último turno lançado)", size=8, color=GREY, italic=True)
    ws["N4"] = '=IF(D4="",_xlfn.MAXIFS(b_data,b_data,"<="&INT(cfg_agora)),D4)'
    ws["N5"] = f'=SUMPRODUCT(MAX((b_data=N4)*ROW(b_data)))'
    ws["N6"] = f'=IF(F4<>"",F4,IF(N5=0,"",INDEX({Q(S_BASE)}!$C:$C,N5)))'
    ws["N7"] = f'=SUMPRODUCT(MAX((b_data=N4)*(b_turno=N6)*ROW(b_data)))'
    ws["N8"] = f'=IF(N7=0,"",INDEX({Q(S_BASE)}!${COL["resp"]}:${COL["resp"]},N7))'
    ws["N9"] = '=IFERROR(VLOOKUP(N6,t_turno,2,FALSE),"")'
    for r in range(4, 10):
        st(ws[f"N{r}"], size=8, color=MGREY, fmt="dd/mm/yyyy" if r == 4 else None)
    add_name(wb, "pt_data", f"{Q('PASSAGEM_DE_TURNO')}!$N$4")
    add_name(wb, "pt_turno", f"{Q('PASSAGEM_DE_TURNO')}!$N$6")
    ws.merge_cells("B6:L6")
    ws["B6"] = ('="TURNO "&N6&"  •  "&' + dtxt("N4") + '&"  •  Líder: "&N8&"      ➜  entrega para o TURNO "&N9')
    st(ws["B6"], bold=True, size=13, color=WHITE, bg=BLUE, h="left")
    ws.row_dimensions[6].height = 26
    crit = "b_data,pt_data,b_turno,pt_turno"
    ks = [("PLANEJADO (t)", f"=SUMIFS(b_plan,{crit},b_med,cfg_unid)", "#,##0.0", NAVY),
          ("EXECUTADO (t)", f"=SUMIFS(b_exec,{crit},b_med,cfg_unid)", "#,##0.0", NAVY),
          ("% EXECUÇÃO", "=IFERROR(D9/B9,0)", "0.0%", NAVY),
          ("PENDÊNCIAS DEIXADAS", f"=COUNTIFS({crit},b_abat,1)", "#,##0", "7F6000"),
          ("CRÍTICOS NO TURNO", f'=COUNTIFS({crit},b_status,"CRÍTICO")', "#,##0", RED),
          ("FAT. PENDENTE", f'=COUNTIFS({crit},b_fat,"PENDENTE")+COUNTIFS({crit},b_fat,"EM PROCESSO")', "#,##0", "7F6000"),
          ]
    for i, (lab, f, fmt, colr) in enumerate(ks):
        kpi(ws, 8, 2 + i * 2, lab, f, fmt, 2, colr)
    ws["F9"] = "=IFERROR(D9/B9,0)"
    pct_cf(ws, "F9", "F9")
    put(ws, "N10", f'=COUNTIFS({crit},b_emb,"PENDENTE")+COUNTIFS({crit},b_emb,"EM PROCESSO")', size=8, color=MGREY)

    cols_short = [("ID", "id", None, 0), ("DATA", "data", FMT_D, 0), ("TURNO", "turno", None, 0),
                  ("PROGRAMAÇÃO / ROTA", ("f", 'IF({i}="","",INDEX(b_prog,{i})&" • "&INDEX(b_rota,{i}))'), None, 0),
                  ("ALERTA", "alerta", None, 0),
                  ("PENDÊNCIA / OCORRÊNCIA", ("f", 'IF({i}="","",IF(INDEX(b_pend,{i})&""="",INDEX(b_obs,{i})&"",INDEX(b_pend,{i})&""))'), None, 0),
                  ("RESPONSÁVEL", "rtrat", None, 0), ("PRAZO", "prazo", FMT_DT, 0), ("AÇÃO", "acao", None, 0)]
    r = 12
    ranked_list(ws, r, 10, "b_sc_crit", cols_short, title="🔴 CRÍTICO — ação imediata (todos os itens críticos em aberto)", span=10, color=RED)
    r += 13
    ranked_list(ws, r, 10, "b_sc_aten", cols_short, title="🟡 ATENÇÃO — acompanhar", span=10, color="BF8F00")
    r += 13
    ranked_list(ws, r, 10, "b_sc_pconc", cols_short[:5] + [("OBSERVAÇÃO", "obs", None, 0), ("EXECUTADO", "exec", FMT_N, 0),
                                                            ("MEDIDA", "med", None, 0), ("TRANSPORTADORA", "transp", None, 0)],
                title="🟢 CONCLUÍDO NESTE TURNO", span=10, color=GREEN)
    r += 13
    first_p, last_p = ranked_list(ws, r, 15, "b_sc_ppend", cols_short, title="📌 PENDÊNCIAS QUE FICAM PARA O PRÓXIMO TURNO", span=10, color=NAVY)
    # descritores para o resumo (colunas ocultas P/Q)
    for rr in range(first_p, last_p + 1):
        ws[f"P{rr}"] = (f'=IF($A{rr}="","","• "&INDEX(b_prog,$A{rr})&" "&INDEX(b_rota,$A{rr})&" — "'
                        f'&MID(INDEX(b_alerta,$A{rr}),3,60)&IF(INDEX(b_rtrat,$A{rr})&""="",""," (resp.: "&INDEX(b_rtrat,$A{rr})&")"))')
        ws[f"Q{rr}"] = (f'=IF($A{rr}="","",IF(INDEX(b_rtrat,$A{rr})&""="","",IF(COUNTIF(R${first_p}:R{rr},INDEX(b_rtrat,$A{rr})&"")=1,'
                        f'INDEX(b_rtrat,$A{rr})&"","")))')
        ws[f"R{rr}"] = f'=IF($A{rr}="","",INDEX(b_rtrat,$A{rr})&"")'
    r += 18
    ranked_list(ws, r, 10, "b_sc_hoje", cols_short, title="⏰ PRAZOS QUE VENCEM HOJE", span=10, color="BF8F00")
    r += 13
    section(ws, r, 2, "👤 RESPONSÁVEIS — itens em aberto por responsável pela tratativa", 10)
    header_row(ws, r + 1, 2, ["", "RESPONSÁVEL", "EM ABERTO", "CRÍTICOS", "VENCIDOS", "VENCE HOJE"])
    for i in range(12):
        rr = r + 2 + i
        cell(ws, rr, 3, f'=IF({cad_item("resp", i)}="","",{cad_item("resp", i)})', bold=True)
        cell(ws, rr, 4, f'=IF(C{rr}="","",COUNTIFS(b_rtrat,C{rr},b_abat,1))', fmt=FMT_I, h="center")
        cell(ws, rr, 5, f'=IF(C{rr}="","",COUNTIFS(b_rtrat,C{rr},b_abat,1,b_status,"CRÍTICO"))', fmt=FMT_I, h="center")
        cell(ws, rr, 6, f'=IF(C{rr}="","",COUNTIFS(b_rtrat,C{rr},b_venc,1))', fmt=FMT_I, h="center")
        cell(ws, rr, 7, f'=IF(C{rr}="","",COUNTIFS(b_rtrat,C{rr},b_abat,1,b_sitpz,"VENCE HOJE"))', fmt=FMT_I, h="center")
    rr = r + 14
    cell(ws, rr, 3, "SEM RESPONSÁVEL", bold=True, color=RED)
    cell(ws, rr, 4, '=COUNTIFS(b_abat,1,b_ocorr,1,b_rtrat,"")', fmt="0", h="center", color=RED, bold=True)

    # resumo automático
    r += 17
    section(ws, r, 2, "📝 RESUMO AUTOMÁTICO DO TURNO (copie e cole no WhatsApp / e-mail)", 10, color=NAVY)
    # top 3 tipos de ocorrência no turno (tabela auxiliar em T:V)
    put(ws, "T11", "aux: ocorrências do turno por tipo", size=8, color=MGREY)
    for i in range(len(TIPOS)):
        rr = 12 + i
        ws[f"T{rr}"] = f"={cad_item('tipo', i)}"
        ws[f"U{rr}"] = f"=COUNTIFS({crit},b_tipo,T{rr},b_ocorr,1)*100+{len(TIPOS) - i}"
    for k in range(3):
        rr = 12 + len(TIPOS) + k
        ws[f"T{rr}"] = (f'=IFERROR(IF(LARGE($U$12:$U${11 + len(TIPOS)},{k + 1})<100,"",'
                        f'"• "&INDEX($T$12:$T${11 + len(TIPOS)},MATCH(LARGE($U$12:$U${11 + len(TIPOS)},{k + 1}),$U$12:$U${11 + len(TIPOS)},0))'
                        f'&" ("&INT(LARGE($U$12:$U${11 + len(TIPOS)},{k + 1})/100)&")"),"")')
    top_first, top_last = 12 + len(TIPOS), 14 + len(TIPOS)
    for col in ("P", "Q", "R", "T", "U"):
        ws.column_dimensions[col].hidden = True
    nl = "&CHAR(10)&"
    resumo = ('="TURNO "&N6&" — "&' + dtxt("N4") + nl + '"Líder: "&N8' + nl + '""' + nl +
              '"Planejado: "&FIXED(B9,1)&" t"' + nl + '"Executado: "&FIXED(D9,1)&" t"' + nl +
              '"Atingimento: "&ROUND(F9*100,0)&"%"' + nl + '""' + nl +
              '"Pendências: "&H9' + nl + '"Críticas: "&J9' + nl + '"Faturamento pendente: "&L9' + nl +
              '"Embarques pendentes: "&N10' + nl + '""' + nl +
              f'"Principais ocorrências:"{nl}IF(T{top_first}="","• Nenhuma ocorrência registrada",_xlfn.TEXTJOIN(CHAR(10),TRUE,T{top_first}:T{top_last})){nl}""{nl}'
              f'"Pendências para o próximo turno ("&N9&"):"{nl}IF(P{first_p}="","• Nenhuma",_xlfn.TEXTJOIN(CHAR(10),TRUE,P{first_p}:P{last_p})){nl}""{nl}'
              f'"Responsáveis: "&IF(_xlfn.TEXTJOIN(", ",TRUE,Q{first_p}:Q{last_p})="","(não atribuídos)",_xlfn.TEXTJOIN(", ",TRUE,Q{first_p}:Q{last_p}))')
    ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 30, end_column=11)
    ws.cell(row=r + 1, column=2, value=resumo)
    st(ws.cell(row=r + 1, column=2), size=10, v="top", wrap=True, bg="FAFAFA")
    widths(ws, {"B": 5, "C": 15, "D": 11, "E": 7, "F": 24, "G": 28, "H": 38, "I": 14, "J": 14, "K": 26, "L": 10})
    ws.freeze_panes = "A7"


# ================================================================ LANÇAMENTO (formulário)
FORM = [  # (rótulo, chave base, lista, obrigatório, formato)
    ("Data", "data", None, True, "dd/mm/yyyy"),
    ("Turno", "turno", "l_turno", True, None),
    ("Responsável pelo turno", "resp", "l_resp", True, None),
    ("Tipo de atividade", "ativ", "l_ativ", True, None),
    ("Programação", "prog", "l_prog", True, None),
    ("Rota / cliente", "rota", None, False, None),
    ("Medida", "med", "l_medida", True, None),
    ("Planejado", "plan", None, True, "#,##0.000"),
    ("Executado", "exec", None, True, "#,##0.000"),
    ("Faturamento", "fat", "l_fat", False, None),
    ("Embarque", "emb", "l_emb", False, None),
    ("Volume faturado", "volfat", None, False, "#,##0.000"),
    ("Transportadora", "transp", "l_transp", False, None),
    ("Pendência (o que ficou pendente)", "pend", None, False, None),
    ("Tipo de pendência", "tipo", "l_tipo", False, None),
    ("Causa", "causa", "l_causa", False, None),
    ("Responsável pela tratativa", "rtrat", "l_resp", False, None),
    ("Prazo (data e hora)", "prazo", None, False, "dd/mm/yyyy hh:mm"),
    ("Ação", "acao", None, False, None),
    ("Impacto operacional", "imp", "l_imp", False, None),
    ("Observação", "obs", None, False, None),
]


def build_lancamento(wb, ws):
    banner(ws, "LANÇAMENTO RÁPIDO", "Formulário para registrar uma atividade/ocorrência. Preencha as células amarelas.", 6)
    widths(ws, {"B": 34, "C": 42, "D": 4, "E": 60})
    header_row(ws, 4, 2, ["CAMPO", "VALOR"])
    for i, (lab, key, lst, req, fmt) in enumerate(FORM):
        r = 5 + i
        cell(ws, r, 2, lab + (" *" if req else ""), bold=req)
        input_cell(ws, f"C{r}", None, fmt, h="left")
        if lst:
            dv_list(ws, lst, f"C{r}", strict=lst not in ("l_resp", "l_transp"))
        elif key in ("data",):
            dv_date(ws, f"C{r}")
        elif key == "prazo":
            dv_date(ws, f"C{r}", datetime=True)
        elif key in ("plan", "exec", "volfat"):
            dv_num(ws, f"C{r}")
    n = len(FORM)
    r = 5 + n
    ws.merge_cells(f"B{r + 1}:C{r + 1}")
    ws[f"B{r + 1}"] = (f'=IF(COUNTBLANK(C5:C13)-COUNTBLANK(C10)>0,"⚠ Preencha os campos obrigatórios (*)",'
                       f'IF(AND(C18<>"",OR(C19="",C21="",C22="")),"⚠ Pendência exige TIPO, RESPONSÁVEL e PRAZO",'
                       f'"✔ Pronto para registrar"))')
    st(ws[f"B{r + 1}"], bold=True, size=11, color=NAVY, h="center")
    ws.conditional_formatting.add(f"B{r + 1}", FormulaRule(formula=[f'LEFT(B{r + 1},1)="⚠"'],
                                                          fill=fill(LRED), font=Font(name=FONT, color=RED, bold=True)))
    ws.conditional_formatting.add(f"B{r + 1}", FormulaRule(formula=[f'LEFT(B{r + 1},1)="✔"'],
                                                          fill=fill(LGREEN), font=Font(name=FONT, color=GREEN, bold=True)))
    texto = [
        "COMO REGISTRAR",
        "1) Com a macro instalada (arquivo .xlsm — ver COMO_USAR): pressione Alt+F8 → RegistrarLancamento "
        "(ou clique no botão, se você o criou). A linha é gravada na BASE_OPERACIONAL, com ID, data/hora de "
        "atualização e registro no LOG_ALTERAÇÕES; o formulário é limpo.",
        "2) Sem macro: digite direto na primeira linha vazia da BASE_OPERACIONAL (as mesmas listas "
        "suspensas e validações estão lá). Tudo o que é cinza é calculado sozinho.",
        "Campos automáticos (não digitar): ID, Área, % Execução, Status, Alerta, Tempo de tratativa, "
        "Próximo turno, Impacta próximo turno?, Situação do prazo, Dias em aberto, SLA, Validação, Qualidade.",
        "Regra de ouro: 1 linha = 1 atividade / ocorrência / entrega / pendência. Nunca apague linhas: use CANCELADO? = SIM.",
    ]
    for i, t in enumerate(texto):
        c = ws.cell(row=4 + i * 3 + (0 if i == 0 else 0), column=5, value=t)
        st(c, bold=(i == 0), size=10 if i else 12, color=NAVY if i == 0 else "000000", v="top", wrap=True)
        ws.merge_cells(start_row=4 + i * 3, start_column=5, end_row=6 + i * 3, end_column=5)
    add_name(wb, "frm_campos", f"{Q('LANÇAMENTO')}!$C$5:$C${4 + n}")


# ================================================================ CONTROLE DE PENDÊNCIAS
def build_pendencias(wb, ws):
    banner(ws, "CONTROLE DE PENDÊNCIAS",
           "Somente o que ainda precisa de tratativa, ordenado por: 1) crítico  2) prazo vencido  3) prazo mais próximo  4) impacto. "
           "Clique no ID para ir ao registro na base.", 22)
    ks = [("EM ABERTO", "=SUM(b_abat)", RED if False else NAVY),
          ("CRÍTICOS", '=COUNTIFS(b_abat,1,b_status,"CRÍTICO")', RED),
          ("PRAZO VENCIDO", "=COUNTIFS(b_abat,1,b_venc,1)", RED),
          ("VENCEM HOJE", '=COUNTIFS(b_abat,1,b_sitpz,"VENCE HOJE")', "7F6000"),
          ("SEM RESPONSÁVEL", '=COUNTIFS(b_abat,1,b_ocorr,1,b_rtrat,"")', "7F6000"),
          ("SEM PRAZO", '=COUNTIFS(b_abat,1,b_sitpz,"SEM PRAZO")', "7F6000"),
          ("SLA ESTOURADO", '=COUNTIFS(b_abat,1,b_slaest,"SIM")', RED),
          ("TEMPO MÉDIO DE RESOLUÇÃO (h)", '=IFERROR(AVERAGEIFS(b_ttrat,b_ocorr,1),0)', NAVY)]
    for i, (lab, f, colr) in enumerate(ks):
        kpi(ws, 4, 2 + i * 2, lab, f, "#,##0.0" if "(h)" in lab else "#,##0", 2, colr)
    cols = [("ID", "id", None, 15), ("DATA", "data", FMT_D, 11), ("TURNO", "turno", None, 7),
            ("PENDÊNCIA / OCORRÊNCIA", ("f", 'IF({i}="","",IF(INDEX(b_pend,{i})&""="",MID(INDEX(b_alerta,{i}),3,60)&" — "&INDEX(b_prog,{i})&" "&INDEX(b_rota,{i}),INDEX(b_pend,{i})&""))'), None, 42),
            ("TIPO", "tipo", None, 13), ("CAUSA", "causa", None, 12), ("RESPONSÁVEL", "rtrat", None, 13),
            ("PRAZO", "prazo", FMT_DT, 14), ("DIAS EM ABERTO", "dias", FMT_I, 8), ("IMPACTO", "imp", None, 11),
            ("STATUS", "status", None, 12), ("ALERTA", "alerta", None, 24), ("SITUAÇÃO PRAZO", "sitpz", None, 12),
            ("SLA ESTOURADO?", "slaest", None, 9), ("AÇÃO", "acao", None, 28), ("PRÓX. TURNO", "prox", None, 7),
            ("PROGRAMAÇÃO", "prog", None, 12), ("ROTA", "rota", None, 18), ("ORIGEM", "origem", None, 10)]
    first, last = ranked_list(ws, 7, 300, "b_sc_prio", cols)
    ws.column_dimensions["B"].width = 5
    for j, c in enumerate(cols):
        ws.column_dimensions[get_column_letter(3 + j)].width = c[3]
    ws.freeze_panes = "D8"
    put(ws, "B" + str(last + 2), "Lista limitada aos 300 itens de maior prioridade. Para editar uma pendência (responsável, prazo, "
        "ação, conclusão), clique no ID e altere a linha na BASE_OPERACIONAL.", size=9, italic=True, color=GREY)


# ================================================================ FATURAMENTO
def period_inputs(wb, ws, prefix, row=4, default_label="(vazio = todo o histórico)"):
    put(ws, f"B{row}", "De:", bold=True, h="right")
    input_cell(ws, f"C{row}", None, "dd/mm/yyyy")
    put(ws, f"D{row}", "Até:", bold=True, h="right")
    input_cell(ws, f"E{row}", None, "dd/mm/yyyy")
    dv_date(ws, f"C{row}")
    dv_date(ws, f"E{row}")
    put(ws, f"F{row}", default_label, size=8, color=GREY, italic=True)
    ws[f"Z{row}"] = f'=IF(C{row}="",DATE(2000,1,1),C{row})'
    ws[f"AA{row}"] = f'=IF(E{row}="",DATE(2100,12,31),E{row})'
    for ref in (f"Z{row}", f"AA{row}"):
        st(ws[ref], size=8, color=MGREY, fmt="dd/mm/yyyy")
    add_name(wb, f"{prefix}_ini", f"{Q(ws.title)}!$Z${row}")
    add_name(wb, f"{prefix}_fim", f"{Q(ws.title)}!$AA${row}")
    return f'b_data,">="&{prefix}_ini,b_data,"<="&{prefix}_fim'


def build_faturamento(wb, ws):
    banner(ws, "FATURAMENTO", "Programado x faturado e alertas de embarque realizado sem faturamento.", 14)
    P = period_inputs(wb, ws, "fa")
    A3 = '{"PENDENTE","EM PROCESSO","FATURADO"}'
    ks = [("PROGRAMADO (t)", f"=SUMPRODUCT(SUMIFS(b_plan,b_fat,{A3},b_med,cfg_unid,{P}))", "#,##0.0", NAVY),
          ("FATURADO (t)", f"=SUMIFS(b_volfat,b_fat,\"FATURADO\",b_med,cfg_unid,{P})", "#,##0.0", GREEN),
          ("% FATURADO", "=IFERROR(D7/B7,0)", "0.0%", NAVY),
          ("ITENS PENDENTES", f'=COUNTIFS(b_fat,"PENDENTE",{P})', "#,##0", "7F6000"),
          ("ITENS EM PROCESSO", f'=COUNTIFS(b_fat,"EM PROCESSO",{P})', "#,##0", BLUE),
          ("ITENS FATURADOS", f'=COUNTIFS(b_fat,"FATURADO",{P})', "#,##0", GREEN),
          ("🔴 EMBARCADO SEM FATURAR", f"=COUNTIFS(b_regra1,1,b_ativo,1,{P})", "#,##0", RED)]
    for i, (lab, f, fmt, colr) in enumerate(ks):
        kpi(ws, 6, 2 + i * 2, lab, f, fmt, 2, colr)
    pct_cf(ws, "F7", "F7")
    put(ws, "B9", "Programado = volume planejado dos itens com faturamento PENDENTE / EM PROCESSO / FATURADO (itens 'NÃO SE APLICA' "
        "ou vazios ficam fora). Faturado = soma do VOLUME FATURADO.", size=8, italic=True, color=GREY)
    r0 = 11
    section(ws, r0, 2, "POR PROGRAMAÇÃO", 7)
    header_row(ws, r0 + 1, 2, ["PROGRAMAÇÃO", "PROGRAMADO (t)", "FATURADO (t)", "% FATURADO", "PENDENTES", "EM PROCESSO", "FATURADOS"])
    for i in range(len(PROGS)):
        r = r0 + 2 + i
        cell(ws, r, 2, f"={cad_item('prog', i)}", bold=True)
        cell(ws, r, 3, f"=SUMPRODUCT(SUMIFS(b_plan,b_prog,B{r},b_fat,{A3},b_med,cfg_unid,{P}))", fmt=FMT_N)
        cell(ws, r, 4, f'=SUMIFS(b_volfat,b_prog,B{r},b_fat,"FATURADO",b_med,cfg_unid,{P})', fmt=FMT_N)
        cell(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', fmt="0%", h="center")
        for j, s in enumerate(("PENDENTE", "EM PROCESSO", "FATURADO")):
            cell(ws, r, 6 + j, f'=COUNTIFS(b_prog,B{r},b_fat,"{s}",{P})', fmt=FMT_I, h="center")
    pct_cf(ws, f"E{r0 + 2}:E{r0 + 1 + len(PROGS)}", f"E{r0 + 2}")
    r1 = r0 + len(PROGS) + 3
    section(ws, r1, 2, "POR TURNO", 7)
    header_row(ws, r1 + 1, 2, ["TURNO", "PROGRAMADO (t)", "FATURADO (t)", "% FATURADO", "PENDENTES", "EM PROCESSO", "FATURADOS"])
    for i in range(4):
        r = r1 + 2 + i
        cell(ws, r, 2, f"={cad_item('turno', i)}", bold=True, h="center")
        cell(ws, r, 3, f"=SUMPRODUCT(SUMIFS(b_plan,b_turno,B{r},b_fat,{A3},b_med,cfg_unid,{P}))", fmt=FMT_N)
        cell(ws, r, 4, f'=SUMIFS(b_volfat,b_turno,B{r},b_fat,"FATURADO",b_med,cfg_unid,{P})', fmt=FMT_N)
        cell(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', fmt="0%", h="center")
        for j, s in enumerate(("PENDENTE", "EM PROCESSO", "FATURADO")):
            cell(ws, r, 6 + j, f'=COUNTIFS(b_turno,B{r},b_fat,"{s}",{P})', fmt=FMT_I, h="center")
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=r0 + 1, max_row=r0 + 1 + len(PROGS)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 1 + len(PROGS)))
    style_chart(ch, "Programado x faturado por programação (t)", 15, 9)
    ch.x_axis.scaling.orientation = "maxMin"
    color_series(ch, [MGREY, GREEN])
    ws.add_chart(ch, f"J{r0}")
    r2 = r1 + 8
    ranked_list(ws, r2, 50, "b_sc_fat", [
        ("ID", "id", None, 0), ("DATA", "data", FMT_D, 0), ("TURNO", "turno", None, 0), ("PROGRAMAÇÃO", "prog", None, 0),
        ("ROTA", "rota", None, 0), ("TRANSPORTADORA", "transp", None, 0), ("EXECUTADO", "exec", FMT_N, 0),
        ("FATURAMENTO", "fat", None, 0), ("EMBARQUE", "emb", None, 0), ("OBSERVAÇÃO", "obs", None, 0)],
        title="🔴 ALERTA — EMBARQUE REALIZADO COM FATURAMENTO PENDENTE (regra 1)", span=11, color=RED)
    widths(ws, {"B": 14, "C": 15, "D": 13, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 36})
    ws.freeze_panes = "A5"


# ================================================================ EMBARQUES
def build_embarques(wb, ws):
    banner(ws, "EMBARQUES", "Programados x realizados, atrasos e desempenho por transportadora.", 14)
    P = period_inputs(wb, ws, "em")
    A3 = '{"PENDENTE","EM PROCESSO","REALIZADO"}'
    ks = [("PROGRAMADOS", f"=SUMPRODUCT(COUNTIFS(b_emb,{A3},{P}))", "#,##0", NAVY),
          ("REALIZADOS", f'=COUNTIFS(b_emb,"REALIZADO",{P})', "#,##0", GREEN),
          ("PENDENTES", f'=COUNTIFS(b_emb,"PENDENTE",{P})', "#,##0", "7F6000"),
          ("EM PROCESSO", f'=COUNTIFS(b_emb,"EM PROCESSO",{P})', "#,##0", BLUE),
          ("🔴 ATRASADOS", f"=COUNTIFS(b_embatr,1,{P})", "#,##0", RED),
          ("% REALIZADO", "=IFERROR(D7/B7,0)", "0.0%", NAVY),
          ("VOLUME EMBARCADO (t)", f'=SUMIFS(b_exec,b_emb,"REALIZADO",b_med,cfg_unid,{P})', "#,##0.0", NAVY)]
    for i, (lab, f, fmt, colr) in enumerate(ks):
        kpi(ws, 6, 2 + i * 2, lab, f, fmt, 2, colr)
    pct_cf(ws, "L7", "L7")
    put(ws, "B9", "ATRASADO = embarque PENDENTE/EM PROCESSO com prazo vencido ou, sem prazo, com DATA + tolerância (CONFIGURAÇÕES) "
        "já ultrapassada. Considera apenas registros ativos (lançamentos novos e migrados a partir da data de corte).",
        size=8, italic=True, color=GREY)
    r0 = 11
    section(ws, r0, 2, "POR TRANSPORTADORA", 7)
    header_row(ws, r0 + 1, 2, ["TRANSPORTADORA", "PROGRAMADOS", "REALIZADOS", "PENDENTES", "ATRASADOS", "% REALIZADO", "VOLUME (t)"])
    NT = 20
    for i in range(NT):
        r = r0 + 2 + i
        cell(ws, r, 2, f'=IF({cad_item("transp", i)}="","",{cad_item("transp", i)})', bold=True)
        cell(ws, r, 3, f'=IF(B{r}="","",SUMPRODUCT(COUNTIFS(b_transp,B{r},b_emb,{A3},{P})))', fmt=FMT_I, h="center")
        cell(ws, r, 4, f'=IF(B{r}="","",COUNTIFS(b_transp,B{r},b_emb,"REALIZADO",{P}))', fmt=FMT_I, h="center")
        cell(ws, r, 5, f'=IF(B{r}="","",COUNTIFS(b_transp,B{r},b_emb,"PENDENTE",{P})+COUNTIFS(b_transp,B{r},b_emb,"EM PROCESSO",{P}))', fmt=FMT_I, h="center")
        cell(ws, r, 6, f'=IF(B{r}="","",COUNTIFS(b_transp,B{r},b_embatr,1,{P}))', fmt=FMT_I, h="center")
        cell(ws, r, 7, f'=IF(N(C{r})=0,"",D{r}/C{r})', fmt="0%", h="center")
        cell(ws, r, 8, f'=IF(B{r}="","",SUMIFS(b_exec,b_transp,B{r},b_emb,"REALIZADO",b_med,cfg_unid,{P}))', fmt=FMT_N)
    r = r0 + 2 + NT
    cell(ws, r, 2, "(SEM TRANSPORTADORA)", bold=True, color=GREY)
    cell(ws, r, 3, f'=SUMPRODUCT(COUNTIFS(b_transp,"",b_emb,{A3},{P}))', fmt=FMT_I, h="center")
    cell(ws, r, 4, f'=COUNTIFS(b_transp,"",b_emb,"REALIZADO",{P})', fmt=FMT_I, h="center")
    pct_cf(ws, f"G{r0 + 2}:G{r0 + 1 + NT}", f"G{r0 + 2}")
    r1 = r + 2
    section(ws, r1, 2, "POR TURNO", 7)
    header_row(ws, r1 + 1, 2, ["TURNO", "PROGRAMADOS", "REALIZADOS", "PENDENTES", "ATRASADOS", "% REALIZADO", "VOLUME (t)"])
    for i in range(4):
        rr = r1 + 2 + i
        cell(ws, rr, 2, f"={cad_item('turno', i)}", bold=True, h="center")
        cell(ws, rr, 3, f"=SUMPRODUCT(COUNTIFS(b_turno,B{rr},b_emb,{A3},{P}))", fmt=FMT_I, h="center")
        cell(ws, rr, 4, f'=COUNTIFS(b_turno,B{rr},b_emb,"REALIZADO",{P})', fmt=FMT_I, h="center")
        cell(ws, rr, 5, f'=COUNTIFS(b_turno,B{rr},b_emb,"PENDENTE",{P})+COUNTIFS(b_turno,B{rr},b_emb,"EM PROCESSO",{P})', fmt=FMT_I, h="center")
        cell(ws, rr, 6, f"=COUNTIFS(b_turno,B{rr},b_embatr,1,{P})", fmt=FMT_I, h="center")
        cell(ws, rr, 7, f'=IF(N(C{rr})=0,"",D{rr}/C{rr})', fmt="0%", h="center")
        cell(ws, rr, 8, f'=SUMIFS(b_exec,b_turno,B{rr},b_emb,"REALIZADO",b_med,cfg_unid,{P})', fmt=FMT_N)
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.add_data(Reference(ws, min_col=4, max_col=5, min_row=r0 + 1, max_row=r0 + 1 + 12), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 1 + 12))
    style_chart(ch, "Embarques por transportadora", 15, 9)
    ch.x_axis.scaling.orientation = "maxMin"
    color_series(ch, [GREEN, YEL])
    ws.add_chart(ch, f"J{r0}")
    r2 = r1 + 8
    ranked_list(ws, r2, 50, "b_sc_emb", [
        ("ID", "id", None, 0), ("DATA", "data", FMT_D, 0), ("TURNO", "turno", None, 0), ("PROGRAMAÇÃO", "prog", None, 0),
        ("ROTA", "rota", None, 0), ("TRANSPORTADORA", "transp", None, 0), ("EMBARQUE", "emb", None, 0),
        ("PRAZO", "prazo", FMT_DT, 0), ("DIAS EM ABERTO", "dias", FMT_I, 0), ("OBSERVAÇÃO", "obs", None, 0)],
        title="🔴 ALERTA — EMBARQUE PENDENTE COM PROGRAMAÇÃO VENCIDA (mais antigos primeiro)", span=11, color=RED)
    widths(ws, {"B": 22, "C": 13, "D": 13, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 36})
    ws.freeze_panes = "A5"


# ================================================================ DIM CALENDÁRIO
def build_calendario(wb, ws):
    ws.sheet_view.showGridLines = False
    hdr = ["DATA", "ANO", "MÊS", "NOME_MÊS", "SEMANA_ISO", "INICIO_SEMANA", "DIA_SEMANA", "REGISTROS",
           "TURNO_A", "TURNO_B", "TURNO_C", "TURNO_D"]
    header_row(ws, 1, 1, hdr)
    meses = '"JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"'
    dias = '"SEG","TER","QUA","QUI","SEX","SÁB","DOM"'
    d0 = dt.date(2026, 1, 1)
    n = 730
    for i in range(n):
        r = 2 + i
        ws.cell(row=r, column=1, value=d0 + dt.timedelta(days=i)).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=2, value=f"=YEAR(A{r})")
        ws.cell(row=r, column=3, value=f"=MONTH(A{r})")
        ws.cell(row=r, column=4, value=f"=CHOOSE(C{r},{meses})")
        ws.cell(row=r, column=5, value=f"=WEEKNUM(A{r},21)")
        ws.cell(row=r, column=6, value=f"=A{r}-WEEKDAY(A{r},2)+1").number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=7, value=f"=CHOOSE(WEEKDAY(A{r},2),{dias})")
        ws.cell(row=r, column=8, value=f"=COUNTIFS(b_data,A{r})")
        for j, t in enumerate("ABCD"):
            ws.cell(row=r, column=9 + j, value=f'=IF(COUNTIFS(b_data,A{r},b_turno,"{t}")>0,1,0)')
    tab = Table(displayName="dCalendario", ref=f"A1:L{n + 1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    for c in "ABCDEFGHIJKL":
        ws.column_dimensions[c].width = 13
    for name, L in (("c_data", "A"), ("c_reg", "H"), ("c_A", "I"), ("c_B", "J"), ("c_C", "K"), ("c_D", "L")):
        add_name(wb, name, f"{Q('DIM_CALENDÁRIO')}!${L}$2:${L}${n + 1}")
    ws.freeze_panes = "A2"


# ================================================================ PRODUTIVIDADE
def prod_block(ws, r0, title, starts, kind, n, step_expr, label_fmt):
    """Bloco período x indicadores. starts: fórmula da 1ª data de início; step_expr: fórmula do fim a partir do início."""
    section(ws, r0, 2, title, 8)
    header_row(ws, r0 + 1, 2, ["PERÍODO", "PLANEJADO", "EXECUTADO", "% EXEC.", "DESVIO", "REGISTROS", "DIAS C/ OPERAÇÃO",
                               "MÉDIA EXEC./DIA", "ATÉ"])
    for i in range(n):
        r = r0 + 2 + i
        cell(ws, r, 2, starts if i == 0 else None, fmt=label_fmt, bold=True, h="center")
        if i > 0:
            ws.cell(row=r, column=2).value = step_expr.format(prev=f"B{r - 1}")
        end = f"J{r}"
        ws[end] = "=" + {"mes": f"EOMONTH(B{r},0)", "sem": f"B{r}+6", "dia": f"B{r}"}[kind]
        st(ws[end], size=8, color=GREY, fmt="dd/mm/yy", h="center")
        ws[end].border = BORDER
        crit = f'b_fprod,1,b_data,">="&B{r},b_data,"<="&J{r}'
        cell(ws, r, 3, f"=SUMIFS(b_plan,{crit})", fmt=FMT_N)
        cell(ws, r, 4, f"=SUMIFS(b_exec,{crit})", fmt=FMT_N)
        cell(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', fmt="0%", h="center")
        cell(ws, r, 6, f"=D{r}-C{r}", fmt="#,##0.0;[Red]-#,##0.0;-")
        cell(ws, r, 7, f"=COUNTIFS({crit})", fmt=FMT_I, h="center")
        cell(ws, r, 8, f'=COUNTIFS(c_data,">="&B{r},c_data,"<="&J{r},c_reg,">0")', fmt=FMT_I, h="center")
        cell(ws, r, 9, f'=IFERROR(D{r}/H{r},"")', fmt=FMT_N)
    pct_cf(ws, f"E{r0 + 2}:E{r0 + 1 + n}", f"E{r0 + 2}")
    return r0 + 2, r0 + 1 + n


def build_produtividade(wb, ws):
    banner(ws, "PRODUTIVIDADE", "Volume planejado x executado por dia, semana, mês, turno e responsável.", 18)
    fl = [("TURNO", "pr_turno", "lf_turno", "TODOS"), ("RESPONSÁVEL", "pr_resp", "lf_resp", "TODOS"),
          ("PROGRAMAÇÃO", "pr_prog", "lf_prog", "TODOS"), ("MEDIDA", "pr_unid", "l_medida", "TON")]
    for i, (lab, name, lst, dflt) in enumerate(fl):
        L = get_column_letter(3 + i * 2)
        put(ws, f"{L}4", lab, bold=True, size=8, color=GREY, h="center")
        input_cell(ws, f"{L}5", dflt)
        dv_list(ws, lst, f"{L}5")
        add_name(wb, name, f"{Q('PRODUTIVIDADE')}!${L}$5")
    put(ws, "K5", "Medida: TON soma toneladas; PALETE/POSIÇÃO somam unidades.", size=8, italic=True, color=GREY)
    # mensal
    m1, m2 = prod_block(ws, 7, "MENSAL (ano da data de referência)", "=DATE(YEAR(cfg_agora),1,1)", "mes", 12,
                        "=EDATE({prev},1)", "mmm/yyyy")
    # semanal: 12 semanas até a atual
    s1, s2 = prod_block(ws, m2 + 2, "SEMANAL (últimas 12 semanas, início na segunda-feira)",
                        "=INT(cfg_agora)-WEEKDAY(cfg_agora,2)+1-77", "sem", 12, "={prev}+7", "dd/mm/yyyy")
    # diário: 31 dias
    d1, d2 = prod_block(ws, s2 + 2, "DIÁRIO (últimos 31 dias)", "=INT(cfg_agora)-30", "dia", 31, "={prev}+1", "dd/mm/yyyy")
    # por turno
    t0 = d2 + 2
    section(ws, t0, 2, "POR TURNO (todo o histórico filtrado)", 8)
    header_row(ws, t0 + 1, 2, ["TURNO", "PLANEJADO", "EXECUTADO", "% EXEC.", "REGISTROS", "TURNOS TRABALHADOS",
                               "MÉDIA EXEC./TURNO", "COLABORADORES", "EXEC. POR COLABORADOR/TURNO"])
    for i, t in enumerate("ABCD"):
        r = t0 + 2 + i
        cell(ws, r, 2, f"={cad_item('turno', i)}", bold=True, h="center")
        crit = f"b_fprod,1,b_turno,B{r}"
        cell(ws, r, 3, f"=SUMIFS(b_plan,{crit})", fmt=FMT_N)
        cell(ws, r, 4, f"=SUMIFS(b_exec,{crit})", fmt=FMT_N)
        cell(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', fmt="0%", h="center")
        cell(ws, r, 6, f"=COUNTIFS({crit})", fmt=FMT_I, h="center")
        cell(ws, r, 7, f"=SUM(c_{t})", fmt=FMT_I, h="center")
        cell(ws, r, 8, f'=IFERROR(D{r}/G{r},"")', fmt=FMT_N)
        cell(ws, r, 9, f'=IFERROR(VLOOKUP(B{r},t_headcount,2,FALSE),0)', fmt=FMT_I, h="center")
        cell(ws, r, 10, f'=IF(N(I{r})=0,"",H{r}/I{r})', fmt=FMT_N)
    pct_cf(ws, f"E{t0 + 2}:E{t0 + 5}", f"E{t0 + 2}")
    # por responsável
    p0 = t0 + 8
    section(ws, p0, 2, "POR RESPONSÁVEL (líder do turno)", 8)
    header_row(ws, p0 + 1, 2, ["RESPONSÁVEL", "PLANEJADO", "EXECUTADO", "% EXEC.", "REGISTROS", "ITENS EM ABERTO", "DESVIO", "ATINGIU META?"])
    for i in range(12):
        r = p0 + 2 + i
        cell(ws, r, 2, f'=IF({cad_item("resp", i)}="","",{cad_item("resp", i)})', bold=True)
        crit = f"b_fprod,1,b_resp,B{r}"
        cell(ws, r, 3, f'=IF(B{r}="","",SUMIFS(b_plan,{crit}))', fmt=FMT_N)
        cell(ws, r, 4, f'=IF(B{r}="","",SUMIFS(b_exec,{crit}))', fmt=FMT_N)
        cell(ws, r, 5, f'=IF(N(C{r})=0,"",D{r}/C{r})', fmt="0%", h="center")
        cell(ws, r, 6, f'=IF(B{r}="","",COUNTIFS({crit}))', fmt=FMT_I, h="center")
        cell(ws, r, 7, f'=IF(B{r}="","",COUNTIFS(b_resp,B{r},b_abat,1))', fmt=FMT_I, h="center")
        cell(ws, r, 8, f'=IF(B{r}="","",D{r}-C{r})', fmt="#,##0.0;[Red]-#,##0.0;-")
        cell(ws, r, 9, f'=IF(E{r}="","",IF(E{r}>=cfg_meta,"SIM","NÃO"))', h="center")
    pct_cf(ws, f"E{p0 + 2}:E{p0 + 13}", f"E{p0 + 2}")
    ws.column_dimensions["J"].hidden = False
    widths(ws, {"B": 14, "C": 13, "D": 13, "E": 9, "F": 12, "G": 11, "H": 13, "I": 13, "J": 14})
    # gráficos
    ch = BarChart()
    ch.type = "col"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=8, max_row=m2), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=m1, max_row=m2))
    style_chart(ch, "Planejado x executado por mês", 17, 7.5)
    color_series(ch, [MGREY, NAVY])
    ch.x_axis.number_format = "mmm"
    ws.add_chart(ch, "L7")
    ch2 = LineChart()
    ch2.add_data(Reference(ws, min_col=3, max_col=4, min_row=d1 - 1, max_row=d2), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=2, min_row=d1, max_row=d2))
    style_chart(ch2, "Planejado x executado — últimos 31 dias", 17, 7.5)
    color_series(ch2, ["A6A6A6", BLUE])
    ch2.x_axis.number_format = "dd/mm"
    ws.add_chart(ch2, f"L{s1}")
    ch3 = BarChart()
    ch3.type = "col"
    ch3.add_data(Reference(ws, min_col=3, max_col=4, min_row=s1 - 1, max_row=s2), titles_from_data=True)
    ch3.set_categories(Reference(ws, min_col=2, min_row=s1, max_row=s2))
    style_chart(ch3, "Planejado x executado por semana", 17, 7.5)
    color_series(ch3, [MGREY, BLUE])
    ch3.x_axis.number_format = "dd/mm"
    ws.add_chart(ch3, f"L{d1 + 2}")
    ws.freeze_panes = "A6"


# ================================================================ GESTÃO SEMANAL / MENSAL
def count_table(ws, r0, col, title, items_key, n, crit, label="QTDE", with_rec=False, span=4):
    section(ws, r0, col, title, span)
    hdrs = ["ITEM", label, "%"] + (["RECORRENTE?"] if with_rec else [])
    header_row(ws, r0 + 1, col, hdrs)
    L = get_column_letter(col)
    Lq = get_column_letter(col + 1)
    for i in range(n):
        r = r0 + 2 + i
        cell(ws, r, col, f"={cad_item(items_key, i)}", bold=True)
        cell(ws, r, col + 1, f"=COUNTIFS({crit.format(item=f'{L}{r}')})", fmt="0", h="center")
        cell(ws, r, col + 2, f"=IFERROR({Lq}{r}/SUM(${Lq}${r0 + 2}:${Lq}${r0 + 1 + n}),0)", fmt="0%", h="center")
        if with_rec:
            cell(ws, r, col + 3, f'=IF({Lq}{r}>=cfg_rec_min,"SIM","")', h="center", color=RED, bold=True)
    return r0 + 2, r0 + 1 + n


def build_semanal(wb, ws):
    banner(ws, "GESTÃO SEMANAL", "Consolidado da semana (segunda a domingo).", 14)
    put(ws, "B4", "Qualquer data da semana:", bold=True, h="right")
    ws.merge_cells("B4:C4")
    input_cell(ws, "D4", None, "dd/mm/yyyy")
    dv_date(ws, "D4")
    put(ws, "E4", "(vazio = semana atual)", size=8, color=GREY, italic=True)
    ws["H4"] = "=IF(D4=\"\",INT(cfg_agora),D4)-WEEKDAY(IF(D4=\"\",INT(cfg_agora),D4),2)+1"
    ws["I4"] = "=H4+6"
    put(ws, "G4", "Semana:", bold=True, h="right")
    for ref in ("H4", "I4"):
        st(ws[ref], bold=True, color=NAVY, fmt="dd/mm/yyyy", h="center")
    add_name(wb, "ws_ini", f"{Q('GESTÃO_SEMANAL')}!$H$4")
    add_name(wb, "ws_fim", f"{Q('GESTÃO_SEMANAL')}!$I$4")
    P = 'b_data,">="&ws_ini,b_data,"<="&ws_fim'
    ks = [("VOLUME PLANEJADO (t)", f"=SUMIFS(b_plan,b_med,cfg_unid,{P})", "#,##0.0", NAVY),
          ("VOLUME EXECUTADO (t)", f"=SUMIFS(b_exec,b_med,cfg_unid,{P})", "#,##0.0", NAVY),
          ("EXECUÇÃO", "=IFERROR(D7/B7,0)", "0.0%", NAVY),
          ("REGISTROS", f"=COUNTIFS({P})", "#,##0", NAVY),
          ("OCORRÊNCIAS DE PENDÊNCIA", f"=COUNTIFS(b_ocorr,1,{P})", "#,##0", "7F6000"),
          ("TIPOS RECORRENTES", None, "#,##0", RED),
          ("FATURADO (t)", f'=SUMIFS(b_volfat,b_med,cfg_unid,{P})', "#,##0.0", GREEN),
          ("EMBARQUES REALIZADOS", f'=COUNTIFS(b_emb,"REALIZADO",{P})', "#,##0", GREEN),
          ("FAT. PENDENTE (itens)", f'=COUNTIFS(b_fat,"PENDENTE",{P})+COUNTIFS(b_fat,"EM PROCESSO",{P})', "#,##0", "7F6000"),
          ("EMBARQUES PENDENTES", f'=COUNTIFS(b_emb,"PENDENTE",{P})+COUNTIFS(b_emb,"EM PROCESSO",{P})', "#,##0", "7F6000")]
    for i, (lab, f, fmt, colr) in enumerate(ks):
        kpi(ws, 6 if i < 5 else 9, 2 + (i % 5) * 2, lab, f, fmt, 2, colr)
    pct_cf(ws, "F7", "F7")
    # diário
    r0 = 12
    section(ws, r0, 2, "DIA A DIA", 6)
    header_row(ws, r0 + 1, 2, ["DIA", "PLANEJADO (t)", "EXECUTADO (t)", "% EXEC.", "DESVIO (t)", "OCORRÊNCIAS"])
    for i in range(7):
        r = r0 + 2 + i
        cell(ws, r, 2, f"=ws_ini+{i}", fmt="ddd dd/mm", bold=True)
        cell(ws, r, 3, f"=SUMIFS(b_plan,b_med,cfg_unid,b_data,B{r})", fmt=FMT_N)
        cell(ws, r, 4, f"=SUMIFS(b_exec,b_med,cfg_unid,b_data,B{r})", fmt=FMT_N)
        cell(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', fmt="0%", h="center")
        cell(ws, r, 6, f"=D{r}-C{r}", fmt="#,##0.0;[Red]-#,##0.0;-")
        cell(ws, r, 7, f"=COUNTIFS(b_data,B{r},b_ocorr,1)", fmt=FMT_I, h="center")
    pct_cf(ws, f"E{r0 + 2}:E{r0 + 8}", f"E{r0 + 2}")
    r = r0 + 9
    cell(ws, r, 2, "Dia com maior desvio:", bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    cell(ws, r, 4, f'=IF(MIN(F{r0 + 2}:F{r0 + 8})>=0,"sem desvio negativo",'
         f'{dtxt(f"INDEX(B{r0 + 2}:B{r0 + 8},MATCH(MIN(F{r0 + 2}:F{r0 + 8}),F{r0 + 2}:F{r0 + 8},0))")}&" ("&FIXED(MIN(F{r0 + 2}:F{r0 + 8}),1)&" t)")',
         bold=True, color=RED)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    # turno
    t0 = r + 2
    section(ws, t0, 2, "POR TURNO", 6)
    header_row(ws, t0 + 1, 2, ["TURNO", "PLANEJADO (t)", "EXECUTADO (t)", "% EXEC.", "ITENS EM ABERTO", "OCORRÊNCIAS"])
    for i in range(4):
        rr = t0 + 2 + i
        cell(ws, rr, 2, f"={cad_item('turno', i)}", bold=True, h="center")
        cell(ws, rr, 3, f"=SUMIFS(b_plan,b_med,cfg_unid,b_turno,B{rr},{P})", fmt=FMT_N)
        cell(ws, rr, 4, f"=SUMIFS(b_exec,b_med,cfg_unid,b_turno,B{rr},{P})", fmt=FMT_N)
        cell(ws, rr, 5, f'=IFERROR(D{rr}/C{rr},"")', fmt="0%", h="center")
        cell(ws, rr, 6, f"=COUNTIFS(b_turno,B{rr},b_abat,1,{P})", fmt=FMT_I, h="center")
        cell(ws, rr, 7, f"=COUNTIFS(b_turno,B{rr},b_ocorr,1,{P})", fmt=FMT_I, h="center")
    pct_cf(ws, f"E{t0 + 2}:E{t0 + 5}", f"E{t0 + 2}")
    rr = t0 + 6
    cell(ws, rr, 2, "Turno com maior volume:", bold=True)
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    cell(ws, rr, 4, f'=IF(MAX(D{t0 + 2}:D{t0 + 5})=0,"—",INDEX(B{t0 + 2}:B{t0 + 5},MATCH(MAX(D{t0 + 2}:D{t0 + 5}),D{t0 + 2}:D{t0 + 5},0))'
         f'&" ("&FIXED(MAX(D{t0 + 2}:D{t0 + 5}),1)&" t)")', bold=True, color=NAVY)
    ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=7)
    # tipos e causas
    a1, a2 = count_table(ws, 12, 9, "OCORRÊNCIAS POR TIPO (recorrência)", "tipo", len(TIPOS),
                         f"b_tipo,{{item}},b_ocorr,1,{P}", with_rec=True)
    count_table(ws, a2 + 2, 9, "OCORRÊNCIAS POR CAUSA", "causa", len(CAUSAS), f"b_causa,{{item}},b_ocorr,1,{P}", with_rec=True)
    ws["B10"].value = f'=COUNTIF(L{a1}:L{a2},"SIM")'  # KPI "TIPOS RECORRENTES"
    widths(ws, {"B": 14, "C": 13, "D": 13, "E": 10, "F": 12, "G": 12, "H": 3, "I": 16, "J": 9, "K": 8, "L": 12})
    ch = BarChart()
    ch.type = "col"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=r0 + 1, max_row=r0 + 8), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=r0 + 8))
    style_chart(ch, "Planejado x executado na semana (t)", 15, 7)
    color_series(ch, [MGREY, NAVY])
    ch.x_axis.number_format = "ddd dd/mm"
    ws.add_chart(ch, "N6")
    ws.freeze_panes = "A5"


def build_mensal(wb, ws):
    banner(ws, "GESTÃO MENSAL", "Consolidado do mês e evolução ao longo do ano.", 14)
    put(ws, "B4", "Qualquer data do mês:", bold=True, h="right")
    ws.merge_cells("B4:C4")
    input_cell(ws, "D4", None, "dd/mm/yyyy")
    dv_date(ws, "D4")
    put(ws, "E4", "(vazio = mês atual)", size=8, color=GREY, italic=True)
    ws["H4"] = '=DATE(YEAR(IF(D4="",cfg_agora,D4)),MONTH(IF(D4="",cfg_agora,D4)),1)'
    ws["I4"] = "=EOMONTH(H4,0)"
    put(ws, "G4", "Mês:", bold=True, h="right")
    st(ws["H4"], bold=True, color=NAVY, fmt="mmmm/yyyy", h="center")
    st(ws["I4"], size=8, color=MGREY, fmt="dd/mm/yyyy")
    add_name(wb, "gm_ini", f"{Q('GESTÃO_MENSAL')}!$H$4")
    add_name(wb, "gm_fim", f"{Q('GESTÃO_MENSAL')}!$I$4")
    P = 'b_data,">="&gm_ini,b_data,"<="&gm_fim'
    ks = [("TOTAL PLANEJADO (t)", f"=SUMIFS(b_plan,b_med,cfg_unid,{P})", "#,##0.0", NAVY),
          ("TOTAL REALIZADO (t)", f"=SUMIFS(b_exec,b_med,cfg_unid,{P})", "#,##0.0", NAVY),
          ("ATINGIMENTO", "=IFERROR(D7/B7,0)", "0.0%", NAVY),
          ("REGISTROS", f"=COUNTIFS({P})", "#,##0", NAVY),
          ("TOTAL DE PENDÊNCIAS", f"=COUNTIFS(b_ocorr,1,{P})", "#,##0", "7F6000"),
          ("% DE PENDÊNCIAS", "=IFERROR(J7/H7,0)", "0.0%", "7F6000"),
          ("TOTAL DE ATRASOS", f"=COUNTIFS(b_atraso,1,{P})", "#,##0", RED),
          ("EMBARQUES REALIZADOS", f'=COUNTIFS(b_emb,"REALIZADO",{P})', "#,##0", GREEN),
          ("TOTAL FATURADO (t)", f'=SUMIFS(b_volfat,b_med,cfg_unid,{P})', "#,##0.0", GREEN),
          ("TEMPO MÉDIO RESOLUÇÃO (h)", f'=IFERROR(AVERAGEIFS(b_ttrat,b_ocorr,1,{P}),0)', "#,##0.0", NAVY)]
    for i, (lab, f, fmt, colr) in enumerate(ks):
        kpi(ws, 6 if i < 5 else 9, 2 + (i % 5) * 2, lab, f, fmt, 2, colr)
    pct_cf(ws, "F7", "F7")
    a1, a2 = count_table(ws, 12, 2, "PRINCIPAIS CAUSAS (mês)", "causa", len(CAUSAS), f"b_causa,{{item}},b_ocorr,1,{P}", with_rec=True)
    count_table(ws, 12, 7, "OCORRÊNCIAS POR TIPO (mês)", "tipo", len(TIPOS), f"b_tipo,{{item}},b_ocorr,1,{P}", with_rec=True)
    # evolução
    e0 = 12 + len(TIPOS) + 4
    section(ws, e0, 2, "EVOLUÇÃO MENSAL (ano da data de referência)", 9)
    header_row(ws, e0 + 1, 2, ["MÊS", "PLANEJADO (t)", "REALIZADO (t)", "ATINGIMENTO", "PENDÊNCIAS", "ATRASOS",
                               "EMBARQUES", "FATURADO (t)", "META?"])
    for i in range(12):
        r = e0 + 2 + i
        cell(ws, r, 2, f"=DATE(YEAR(gm_ini),{i + 1},1)", fmt="mmm/yyyy", bold=True, h="center")
        pc = f'b_data,">="&B{r},b_data,"<="&EOMONTH(B{r},0)'
        cell(ws, r, 3, f"=SUMIFS(b_plan,b_med,cfg_unid,{pc})", fmt=FMT_N)
        cell(ws, r, 4, f"=SUMIFS(b_exec,b_med,cfg_unid,{pc})", fmt=FMT_N)
        cell(ws, r, 5, f'=IFERROR(D{r}/C{r},"")', fmt="0%", h="center")
        cell(ws, r, 6, f"=COUNTIFS(b_ocorr,1,{pc})", fmt=FMT_I, h="center")
        cell(ws, r, 7, f"=COUNTIFS(b_atraso,1,{pc})", fmt=FMT_I, h="center")
        cell(ws, r, 8, f'=COUNTIFS(b_emb,"REALIZADO",{pc})', fmt=FMT_I, h="center")
        cell(ws, r, 9, f"=SUMIFS(b_volfat,b_med,cfg_unid,{pc})", fmt=FMT_N)
        cell(ws, r, 10, f'=IF(E{r}="","",IF(E{r}>=cfg_meta,"✔","✖"))', h="center")
    pct_cf(ws, f"E{e0 + 2}:E{e0 + 13}", f"E{e0 + 2}")
    widths(ws, {"B": 16, "C": 13, "D": 13, "E": 11, "F": 12, "G": 16, "H": 12, "I": 12, "J": 12, "K": 12})
    ch = BarChart()
    ch.type = "col"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=e0 + 1, max_row=e0 + 13), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=e0 + 2, max_row=e0 + 13))
    style_chart(ch, "Evolução mensal — planejado x realizado (t)", 16, 7.5)
    color_series(ch, [MGREY, NAVY])
    ch.x_axis.number_format = "mmm"
    ws.add_chart(ch, "M6")
    ch2 = BarChart()
    ch2.type = "bar"
    ch2.add_data(Reference(ws, min_col=3, min_row=a1 - 1, max_row=a2), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=2, min_row=a1, max_row=a2))
    style_chart(ch2, "Causas no mês", 16, 7.5)
    ch2.legend = None
    ch2.x_axis.scaling.orientation = "maxMin"
    color_series(ch2, [BLUE])
    ws.add_chart(ch2, f"M{e0 - 6}")
    ws.freeze_panes = "A5"


# ================================================================ ANÁLISE DE CAUSAS
def build_causas(wb, ws):
    banner(ws, "ANÁLISE DE CAUSAS E RECORRÊNCIA", "Pendências por tipo, turno, responsável, causa e impacto; matriz Frequência x Impacto.", 16)
    P = period_inputs(wb, ws, "an")
    O = f"b_ocorr,1,{P}"
    # 1) por tipo + matriz
    r0 = 6
    section(ws, r0, 2, "PENDÊNCIAS POR TIPO + MATRIZ FREQUÊNCIA x IMPACTO", 11)
    header_row(ws, r0 + 1, 2, ["TIPO", "OCORRÊNCIAS", "EM ABERTO", "CONCLUÍDAS", "TEMPO MÉDIO (h)", "SLA (h)",
                               "% DENTRO DO SLA", "IMPACTO MÉDIO (0-4)", "FREQUÊNCIA", "IMPACTO", "CLASSIFICAÇÃO"])
    n = len(TIPOS)
    f1, fl = r0 + 2, r0 + 1 + n
    for i in range(n):
        r = r0 + 2 + i
        cell(ws, r, 2, f"={cad_item('tipo', i)}", bold=True)
        cell(ws, r, 3, f"=COUNTIFS(b_tipo,B{r},{O})", fmt="0", h="center")
        cell(ws, r, 4, f"=COUNTIFS(b_tipo,B{r},b_abat,1,{O})", fmt="0", h="center")
        cell(ws, r, 5, f'=COUNTIFS(b_tipo,B{r},b_concl,">0",{O})', fmt="0", h="center")
        cell(ws, r, 6, f'=IFERROR(AVERAGEIFS(b_ttrat,b_tipo,B{r},{O}),"")', fmt="#,##0.0", h="center")
        cell(ws, r, 7, f"=IFERROR(VLOOKUP(B{r},t_sla,2,FALSE),\"\")", fmt="0", h="center")
        cell(ws, r, 8, f'=IFERROR(COUNTIFS(b_tipo,B{r},b_slaest,"NÃO",{O})/(COUNTIFS(b_tipo,B{r},b_slaest,"NÃO",{O})+COUNTIFS(b_tipo,B{r},b_slaest,"SIM",{O})),"")',
             fmt="0%", h="center")
        cell(ws, r, 9, f'=IFERROR(AVERAGEIFS(b_impn,b_tipo,B{r},b_imp,"<>",{O}),"")', fmt="0.0", h="center")
        cell(ws, r, 10, f'=IF(C{r}=0,"",IF(C{r}>=AVERAGEIF($C${f1}:$C${fl},">0"),"ALTA","BAIXA"))', h="center")
        cell(ws, r, 11, f'=IF(C{r}=0,"",IF(I{r}="","NÃO INFORMADO",IF(I{r}>=cfg_imp_alto,"ALTO","BAIXO")))', h="center")
        cell(ws, r, 12, f'=IF(OR(J{r}="",K{r}="NÃO INFORMADO"),IF(C{r}=0,"","SEM IMPACTO INFORMADO"),'
             f'IF(J{r}="ALTA",IF(K{r}="ALTO","PRIORIDADE DE ANÁLISE","OPORTUNIDADE DE MELHORIA"),'
             f'IF(K{r}="ALTO","RISCO","ACOMPANHAMENTO")))', bold=True)
    for val, bg, fg in (("PRIORIDADE DE ANÁLISE", RED, WHITE), ("RISCO", LRED, RED),
                        ("OPORTUNIDADE DE MELHORIA", "FFF7D6", "7F6000"), ("ACOMPANHAMENTO", LGREEN, GREEN)):
        ws.conditional_formatting.add(f"L{f1}:L{fl}", FormulaRule(formula=[f'L{f1}="{val}"'], fill=fill(bg),
                                                                font=Font(name=FONT, color=fg, bold=True)))
    # matriz 2x2
    m0 = fl + 2
    section(ws, m0, 2, "MATRIZ FREQUÊNCIA x IMPACTO (classificação automática pelos dados registrados)", 11, color=NAVY)
    put(ws, f"B{m0 + 1}", "Frequência ALTA = ocorrências ≥ média dos tipos com ocorrência no período.  Impacto ALTO = impacto médio "
        "registrado ≥ parâmetro em CONFIGURAÇÕES. Tipos sem impacto informado não entram na matriz.", size=8, italic=True, color=GREY)
    quads = [("PRIORIDADE DE ANÁLISE", "Alta frequência + alto impacto", RED, WHITE, m0 + 3, 4),
             ("RISCO", "Baixa frequência + alto impacto", LRED, RED, m0 + 3, 8),
             ("OPORTUNIDADE DE MELHORIA", "Alta frequência + baixo impacto", "FFF7D6", "7F6000", m0 + 7, 4),
             ("ACOMPANHAMENTO", "Baixa frequência + baixo impacto", LGREEN, GREEN, m0 + 7, 8)]
    put(ws, f"B{m0 + 4}", "IMPACTO ALTO", bold=True, color=NAVY, h="center")
    put(ws, f"B{m0 + 8}", "IMPACTO BAIXO", bold=True, color=NAVY, h="center")
    put(ws, f"D{m0 + 11}", "FREQUÊNCIA ALTA", bold=True, color=NAVY, h="center")
    put(ws, f"H{m0 + 11}", "FREQUÊNCIA BAIXA", bold=True, color=NAVY, h="center")
    for name, desc, bg, fg, rr, cc in quads:
        ws.merge_cells(start_row=rr, start_column=cc, end_row=rr, end_column=cc + 3)
        ws.merge_cells(start_row=rr + 1, start_column=cc, end_row=rr + 3, end_column=cc + 3)
        c = ws.cell(row=rr, column=cc, value=f"{name} — {desc}")
        st(c, bold=True, size=9, color=fg, bg=bg, h="center", wrap=True)
        tgt = ws.cell(row=rr + 1, column=cc)
        ref = f"{get_column_letter(cc)}{rr + 1}"
        ws[ref] = ArrayFormula(ref, f'=_xlfn.TEXTJOIN(", ",TRUE,IF($L${f1}:$L${fl}="{name}",$B${f1}:$B${fl}&" ("&$C${f1}:$C${fl}&")",""))')
        st(tgt, size=10, bg=bg if bg not in (RED,) else LRED, color="000000", h="center", v="center", wrap=True)
        for i in range(4):
            for j in range(4):
                try:
                    ws.cell(row=rr + j, column=cc + i).fill = fill(bg if j == 0 else (LRED if bg == RED else bg))
                except AttributeError:
                    pass
    # 2) por turno
    t0 = m0 + 13
    section(ws, t0, 2, "PENDÊNCIAS POR TURNO", 6)
    header_row(ws, t0 + 1, 2, ["TURNO", "OCORRÊNCIAS", "EM ABERTO", "VENCIDAS", "TEMPO MÉDIO (h)", "% NO PRAZO"])
    for i in range(4):
        r = t0 + 2 + i
        cell(ws, r, 2, f"={cad_item('turno', i)}", bold=True, h="center")
        cell(ws, r, 3, f"=COUNTIFS(b_turno,B{r},{O})", fmt="0", h="center")
        cell(ws, r, 4, f"=COUNTIFS(b_turno,B{r},b_abat,1,{O})", fmt="0", h="center")
        cell(ws, r, 5, f"=COUNTIFS(b_turno,B{r},b_venc,1,{O})", fmt="0", h="center")
        cell(ws, r, 6, f'=IFERROR(AVERAGEIFS(b_ttrat,b_turno,B{r},{O}),"")', fmt="#,##0.0", h="center")
        cell(ws, r, 7, f'=IFERROR(AVERAGEIFS(b_noprazo,b_turno,B{r},{O}),"")', fmt="0%", h="center")
    # 3) por responsável
    p0 = t0 + 7
    section(ws, p0, 2, "PENDÊNCIAS POR RESPONSÁVEL PELA TRATATIVA", 6)
    header_row(ws, p0 + 1, 2, ["RESPONSÁVEL", "OCORRÊNCIAS", "EM ABERTO", "VENCIDAS", "TEMPO MÉDIO (h)", "% NO PRAZO"])
    for i in range(12):
        r = p0 + 2 + i
        cell(ws, r, 2, f'=IF({cad_item("resp", i)}="","",{cad_item("resp", i)})', bold=True)
        cell(ws, r, 3, f'=IF(B{r}="","",COUNTIFS(b_rtrat,B{r},{O}))', fmt=FMT_I, h="center")
        cell(ws, r, 4, f'=IF(B{r}="","",COUNTIFS(b_rtrat,B{r},b_abat,1,{O}))', fmt=FMT_I, h="center")
        cell(ws, r, 5, f'=IF(B{r}="","",COUNTIFS(b_rtrat,B{r},b_venc,1,{O}))', fmt=FMT_I, h="center")
        cell(ws, r, 6, f'=IF(B{r}="","",IFERROR(AVERAGEIFS(b_ttrat,b_rtrat,B{r},{O}),""))', fmt="#,##0.0", h="center")
        cell(ws, r, 7, f'=IF(B{r}="","",IFERROR(AVERAGEIFS(b_noprazo,b_rtrat,B{r},{O}),""))', fmt="0%", h="center")
    r = p0 + 14
    cell(ws, r, 2, "(SEM RESPONSÁVEL)", bold=True, color=RED)
    cell(ws, r, 3, f'=COUNTIFS(b_rtrat,"",{O})', fmt="0", h="center", color=RED)
    # 4) recorrência por causa (ordenada)
    c0 = 6
    section(ws, c0, 14, "RECORRÊNCIA POR CAUSA (ordenado por ocorrências)", 6)
    header_row(ws, c0 + 1, 14, ["CAUSA", "OCORRÊNCIAS", "IMPACTO MÉDIO", "IMPACTO", "ÚLTIMA OCORRÊNCIA", "RECORRENTE?"])
    nc = len(CAUSAS)
    # tabela auxiliar (colunas AC:AE) com contagem + desempate
    for i in range(nc):
        r = c0 + 2 + i
        ws[f"AC{r}"] = f"={cad_item('causa', i)}"
        ws[f"AD{r}"] = f"=COUNTIFS(b_causa,AC{r},{O})"
        ws[f"AE{r}"] = f"=AD{r}*100+{nc - i}"
    rngc = f"$AE${c0 + 2}:$AE${c0 + 1 + nc}"
    for k in range(nc):
        r = c0 + 2 + k
        idx = f"MATCH(LARGE({rngc},{k + 1}),{rngc},0)"
        cell(ws, r, 14, f"=INDEX($AC${c0 + 2}:$AC${c0 + 1 + nc},{idx})", bold=True)
        cell(ws, r, 15, f"=INDEX($AD${c0 + 2}:$AD${c0 + 1 + nc},{idx})", fmt="0", h="center")
        cell(ws, r, 16, f'=IFERROR(AVERAGEIFS(b_impn,b_causa,N{r},b_imp,"<>",{O}),"")', fmt="0.0", h="center")
        cell(ws, r, 17, f'=IF(P{r}="","NÃO INFORMADO",INDEX(l_imp,ROUND(P{r},0)+1))', h="center")
        cell(ws, r, 18, f'=IF(O{r}=0,"",_xlfn.MAXIFS(b_data,b_causa,N{r},b_ocorr,1,{P}))', fmt="dd/mm/yyyy;;", h="center")
        cell(ws, r, 19, f'=IF(O{r}>=cfg_rec_min,"SIM","")', h="center", color=RED, bold=True)
    r = c0 + 2 + nc
    cell(ws, r, 14, "(CAUSA NÃO INFORMADA)", bold=True, color=GREY)
    cell(ws, r, 15, f'=COUNTIFS(b_causa,"",{O})', fmt="0", h="center", color=GREY)
    for col in ("AC", "AD", "AE"):
        ws.column_dimensions[col].hidden = True
    # 5) por impacto
    i0 = c0 + nc + 4
    section(ws, i0, 14, "PENDÊNCIAS POR IMPACTO", 6)
    header_row(ws, i0 + 1, 14, ["IMPACTO", "OCORRÊNCIAS", "EM ABERTO", "%", "", ""])
    for i in range(len(IMPACTOS)):
        r = i0 + 2 + i
        cell(ws, r, 14, f"={cad_item('imp', i)}", bold=True)
        cell(ws, r, 15, f"=COUNTIFS(b_imp,N{r},{O})", fmt="0", h="center")
        cell(ws, r, 16, f"=COUNTIFS(b_imp,N{r},b_abat,1,{O})", fmt="0", h="center")
        cell(ws, r, 17, f"=IFERROR(O{r}/SUM($O${i0 + 2}:$O${i0 + 7}),0)", fmt="0%", h="center")
    r = i0 + 2 + len(IMPACTOS)
    cell(ws, r, 14, "(NÃO INFORMADO)", bold=True, color=GREY)
    cell(ws, r, 15, f'=COUNTIFS(b_imp,"",{O})', fmt="0", h="center", color=GREY)
    cell(ws, r, 17, f"=IFERROR(O{r}/SUM($O${i0 + 2}:$O${i0 + 7}),0)", fmt="0%", h="center")
    widths(ws, {"B": 18, "C": 12, "D": 11, "E": 11, "F": 11, "G": 9, "H": 11, "I": 11, "J": 11, "K": 11, "L": 26,
                "M": 3, "N": 22, "O": 12, "P": 11, "Q": 14, "R": 13, "S": 12})
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=r0 + 1, max_row=fl), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=f1, max_row=fl))
    style_chart(ch, "Pendências por categoria", 15, 9)
    ch.x_axis.scaling.orientation = "maxMin"
    color_series(ch, [NAVY, RED])
    ws.add_chart(ch, f"N{i0 + 10}")
    ch2 = BarChart()
    ch2.type = "col"
    ch2.add_data(Reference(ws, min_col=3, min_row=t0 + 1, max_row=t0 + 5), titles_from_data=True)
    ch2.add_data(Reference(ws, min_col=4, min_row=t0 + 1, max_row=t0 + 5), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=2, min_row=t0 + 2, max_row=t0 + 5))
    style_chart(ch2, "Pendências por turno", 15, 7)
    color_series(ch2, [NAVY, RED])
    ws.add_chart(ch2, f"N{i0 + 29}")
    ws.freeze_panes = "A5"


# ================================================================ PARETO
def pareto_block(ws, r0, col, title, key, n, crit_tpl):
    section(ws, r0, col, title, 5)
    header_row(ws, r0 + 1, col, ["ITEM", "QTDE", "%", "% ACUMULADO", "CLASSE"])
    aux = get_column_letter(col + 20)
    aux2 = get_column_letter(col + 21)
    for i in range(n):
        r = r0 + 2 + i
        ws[f"{aux}{r}"] = f"={cad_item(key, i)}"
        ws[f"{aux2}{r}"] = f"=COUNTIFS({crit_tpl.format(item=f'{aux}{r}')})*100+{n - i}"
    rng = f"${aux2}${r0 + 2}:${aux2}${r0 + 1 + n}"
    names = f"${aux}${r0 + 2}:${aux}${r0 + 1 + n}"
    L, Lq, Lp, La = (get_column_letter(col + i) for i in range(4))
    for k in range(n):
        r = r0 + 2 + k
        idx = f"MATCH(LARGE({rng},{k + 1}),{rng},0)"
        cell(ws, r, col, f"=INDEX({names},{idx})", bold=True)
        cell(ws, r, col + 1, f"=INT(LARGE({rng},{k + 1})/100)", fmt="0", h="center")
        cell(ws, r, col + 2, f"=IFERROR({Lq}{r}/SUM(${Lq}${r0 + 2}:${Lq}${r0 + 1 + n}),0)", fmt="0.0%", h="center")
        cell(ws, r, col + 3, f"=SUM(${Lp}${r0 + 2}:{Lp}{r})", fmt="0.0%", h="center")
        cell(ws, r, col + 4, f'=IF({Lq}{r}=0,"",IF({La}{r}-{Lp}{r}<0.8,"A (vital)",IF({La}{r}-{Lp}{r}<0.95,"B","C")))', h="center")
    ws.column_dimensions[aux].hidden = True
    ws.column_dimensions[aux2].hidden = True
    ch = BarChart()
    ch.type = "col"
    ch.add_data(Reference(ws, min_col=col + 1, min_row=r0 + 1, max_row=r0 + 1 + n), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=col, min_row=r0 + 2, max_row=r0 + 1 + n))
    color_series(ch, [NAVY])
    ln = LineChart()
    ln.add_data(Reference(ws, min_col=col + 3, min_row=r0 + 1, max_row=r0 + 1 + n), titles_from_data=True)
    ln.y_axis.axId = 200
    ln.y_axis.number_format = "0%"
    ln.y_axis.scaling.min = 0
    ln.y_axis.scaling.max = 1
    ln.y_axis.crosses = "max"
    ln.y_axis.majorGridlines = None
    ln.series[0].graphicalProperties.line.solidFill = RED
    ln.series[0].graphicalProperties.line.width = 28000
    ln.series[0].marker.symbol = "circle"
    try:
        ln.y_axis.delete = False
    except AttributeError:
        pass
    style_chart(ch, title, 18, 9)
    ch += ln
    return ch, r0 + 1 + n


def build_pareto(wb, ws):
    banner(ws, "PARETO DE PROBLEMAS", "Quais causas/tipos concentram a maior parte das pendências (período da aba ANÁLISE_CAUSAS).", 16)
    ws["B4"] = ('="Período: "&' + dtxt("MAX(an_ini,MIN(b_data))") + '&" a "&' + dtxt("MIN(an_fim,MAX(b_data))") +
                '&"   •   Altere o período na aba ANÁLISE_CAUSAS."')
    st(ws["B4"], size=9, color=GREY, italic=True)
    O = 'b_ocorr,1,b_data,">="&an_ini,b_data,"<="&an_fim'
    ch1, e1 = pareto_block(ws, 6, 2, "TOP CAUSAS DE DESVIO", "causa", len(CAUSAS), f"b_causa,{{item}},{O}")
    ws.add_chart(ch1, "H6")
    r = e1 + 1
    ws.merge_cells(f"B{r}:F{r}")
    ws[f"B{r}"] = (f'="Ocorrências sem causa informada: "&COUNTIFS(b_causa,"",{O})&" de "&COUNTIFS({O})'
                   f'&" — preencha a CAUSA para o Pareto refletir a realidade."')
    st(ws[f"B{r}"], size=9, bold=True, color=RED, wrap=True)
    ws.row_dimensions[r].height = 28
    ch2, e2 = pareto_block(ws, e1 + 4, 2, "TOP TIPOS DE PENDÊNCIA", "tipo", len(TIPOS), f"b_tipo,{{item}},{O}")
    ws.add_chart(ch2, f"H{e1 + 4}")
    put(ws, f"B{e2 + 2}", "Classe A = itens que somam até ~80% das ocorrências (foco de ação); B até 95%; C restante.",
        size=8, italic=True, color=GREY)
    widths(ws, {"B": 20, "C": 9, "D": 9, "E": 12, "F": 10})
    ws.freeze_panes = "A5"


# ================================================================ HISTÓRICO
def build_historico(wb, ws):
    banner(ws, "HISTÓRICO — consulta", "Nenhum registro é apagado. Filtre por período, turno, responsável, programação, rota, tipo, causa, "
           "status ou texto. Resultado: até 500 linhas, mais recentes primeiro.", 22)
    fl = [("DE", "h_dini", None, None), ("ATÉ", "h_dfim", None, None), ("TURNO", "h_turno", "lf_turno", "TODOS"),
          ("RESPONSÁVEL", "h_resp", "lf_resp", "TODOS"), ("PROGRAMAÇÃO", "h_prog", "lf_prog", "TODOS"),
          ("ROTA (contém)", "h_rota", None, None), ("TIPO PENDÊNCIA", "h_tipo", "lf_tipo", "TODOS"),
          ("CAUSA", "h_causa", "lf_causa", "TODOS"), ("STATUS", "h_status", "lf_status", "TODOS"),
          ("TEXTO (contém)", "h_texto", None, None)]
    for i, (lab, name, lst, d) in enumerate(fl):
        L = get_column_letter(3 + i)
        put(ws, f"{L}4", lab, bold=True, size=8, color=GREY, h="center", wrap=True)
        input_cell(ws, f"{L}5", d, "dd/mm/yyyy" if lab in ("DE", "ATÉ") else None)
        if lst:
            dv_list(ws, lst, f"{L}5")
        elif lab in ("DE", "ATÉ"):
            dv_date(ws, f"{L}5")
        add_name(wb, name, f"{Q('HISTÓRICO')}!${L}$5")
    ws["AF5"] = '=IF(h_dini="",DATE(2000,1,1),h_dini)'
    ws["AG5"] = '=IF(h_dfim="",DATE(2100,12,31),h_dfim)'
    add_name(wb, "h_ini", f"{Q('HISTÓRICO')}!$AF$5")
    add_name(wb, "h_fim", f"{Q('HISTÓRICO')}!$AG$5")
    ws["M5"] = '="Encontrados: "&COUNT(b_sc_hist)&IF(COUNT(b_sc_hist)>500," (exibindo 500)","")'
    st(ws["M5"], bold=True, color=NAVY)
    cols = [("ID", "id", None, 15), ("DATA", "data", FMT_D, 11), ("TURNO", "turno", None, 7), ("RESPONSÁVEL", "resp", None, 12),
            ("PROGRAMAÇÃO", "prog", None, 13), ("ROTA", "rota", None, 22), ("MEDIDA", "med", None, 8),
            ("PLANEJADO", "plan", FMT_N, 10), ("EXECUTADO", "exec", FMT_N, 10), ("% EXEC.", "pct", "0%;;", 8),
            ("STATUS", "status", None, 12), ("FATURAMENTO", "fat", None, 12), ("EMBARQUE", "emb", None, 12),
            ("TRANSPORTADORA", "transp", None, 13), ("PENDÊNCIA", "pend", None, 30), ("TIPO", "tipo", None, 12),
            ("CAUSA", "causa", None, 12), ("RESP. TRATATIVA", "rtrat", None, 12), ("PRAZO", "prazo", FMT_DT, 13),
            ("CONCLUSÃO", "concl", FMT_DT, 13), ("OBSERVAÇÃO", "obs", None, 40)]
    ranked_list(ws, 7, 500, "b_sc_hist", cols)
    ws.column_dimensions["B"].width = 5
    for j, c in enumerate(cols):
        ws.column_dimensions[get_column_letter(3 + j)].width = c[3]
    ws.freeze_panes = "D8"


# ================================================================ QUALIDADE
def build_qualidade(wb, ws):
    banner(ws, "QUALIDADE DOS DADOS", "Confiabilidade da base: completude, padronização, duplicidades e regras de negócio.", 16)
    T = "COUNT(b_data)"
    ks = [("REGISTROS", f"={T}", "#,##0", NAVY),
          ("COMPLETOS", '=SUMPRODUCT((b_data<>"")*(b_valid="")*(b_qual=""))', "#,##0", GREEN),
          ("INCOMPLETOS / COM ERRO", "=B7-D7", "#,##0", RED),
          ("% QUALIDADE DA BASE", "=IFERROR(D7/B7,0)", "0.0%", NAVY)]
    for i, (lab, f, fmt, colr) in enumerate(ks):
        kpi(ws, 6, 2 + i * 2, lab, f, fmt, 2, colr)
    ws["H7"] = "=IFERROR(D7/B7,0)"
    ws["F7"] = "=B7-D7"
    ws.conditional_formatting.add("H7", FormulaRule(formula=["H7<0.9"], font=Font(name=FONT, color=RED, bold=True, size=18)))
    r0 = 10
    section(ws, r0, 2, "INDICADORES DE QUALIDADE", 5)
    header_row(ws, r0 + 1, 2, ["INDICADOR", "QTDE", "% DA BASE", "ONDE CORRIGIR", ""])
    items = [
        ("Pendências sem responsável (ativas)", '=COUNTIFS(b_valid,"*sem responsável*")', "BASE: RESPONSÁVEL PELA TRATATIVA"),
        ("Pendências sem prazo (ativas)", '=COUNTIFS(b_valid,"*sem prazo*")', "BASE: PRAZO"),
        ("Pendências sem tipo", '=COUNTIFS(b_valid,"*sem tipo*")', "BASE: TIPO DE PENDÊNCIA"),
        ("Inconsistências (conclusão/prazo)", '=COUNTIFS(b_valid,"*INCONSISTÊNCIA*")', "BASE: DATA DE CONCLUSÃO / PRAZO"),
        ("Divergências (executado > planejado)", '=COUNTIFS(b_valid,"*DIVERGÊNCIA*")', "BASE: PLANEJADO / EXECUTADO"),
        ("Atividades sem planejamento", '=COUNTIFS(b_qual,"*Sem planejamento*")', "BASE: PLANEJADO"),
        ("Atividades sem execução", '=COUNTIFS(b_qual,"*Sem execução*")', "BASE: EXECUTADO"),
        ("Turno vazio ou fora do padrão", '=COUNTIFS(b_qual,"*Turno*")', "BASE: TURNO"),
        ("Responsável do turno vazio", '=COUNTIFS(b_qual,"*Responsável vazio*")', "BASE: RESPONSÁVEL PELO TURNO"),
        ("Programação vazia ou fora do padrão", '=COUNTIFS(b_qual,"*Programação*")', "BASE: PROGRAMAÇÃO / CADASTROS"),
        ("Medida vazia ou fora do padrão", '=COUNTIFS(b_qual,"*Medida*")', "BASE: MEDIDA"),
        ("Cadastros fora do padrão (fat./emb./transp.)", '=COUNTIFS(b_qual,"*fora do*")-COUNTIFS(b_qual,"*Turno fora*")-COUNTIFS(b_qual,"*Programação fora*")-COUNTIFS(b_qual,"*Medida fora*")', "CADASTROS"),
        ("Valores acima do limite (provável erro)", '=COUNTIFS(b_qual,"*acima do limite*")', "BASE: PLANEJADO / EXECUTADO"),
        ("Registros possivelmente duplicados", "=SUM(b_dup)", "BASE (use CANCELADO? = SIM na cópia)"),
        ("Ocorrências sem causa informada", '=COUNTIFS(b_ocorr,1,b_causa,"")', "BASE: CAUSA"),
        ("Ocorrências sem impacto informado", '=COUNTIFS(b_ocorr,1,b_imp,"")', "BASE: IMPACTO OPERACIONAL"),
    ]
    for i, (lab, f, onde) in enumerate(items):
        r = r0 + 2 + i
        cell(ws, r, 2, lab, bold=True)
        cell(ws, r, 3, f, fmt="#,##0", h="center")
        cell(ws, r, 4, f"=IFERROR(C{r}/$B$7,0)", fmt="0.0%", h="center")
        cell(ws, r, 5, onde, size=9, color=GREY)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    last = r0 + 1 + len(items)
    ws.conditional_formatting.add(f"C{r0 + 2}:C{last}", CellIsRule(operator="greaterThan", formula=["0"],
                                                                  font=Font(name=FONT, color=RED, bold=True)))
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(ws, min_col=3, min_row=r0 + 1, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r0 + 2, max_row=last))
    style_chart(ch, "Problemas de qualidade por tipo", 16, 9)
    ch.legend = None
    ch.x_axis.scaling.orientation = "maxMin"
    color_series(ch, [RED])
    ws.add_chart(ch, "I10")
    ranked_list(ws, last + 3, 200, "b_sc_qual", [
        ("ID", "id", None, 0), ("DATA", "data", FMT_D, 0), ("TURNO", "turno", None, 0), ("RESPONSÁVEL", "resp", None, 0),
        ("PROGRAMAÇÃO", "prog", None, 0), ("VALIDAÇÃO (REGRAS)", "valid", None, 0), ("QUALIDADE DO CADASTRO", "qual", None, 0),
        ("ORIGEM", "origem", None, 0)], title="REGISTROS COM PROBLEMA (mais recentes primeiro — até 200)", span=9, color=RED)
    widths(ws, {"B": 38, "C": 15, "D": 11, "E": 13, "F": 12, "G": 14, "H": 44, "I": 44, "J": 11})
    ws.freeze_panes = "A5"


# ================================================================ LOG
def build_log(wb, ws, recs, mig_log):
    ws.sheet_view.showGridLines = False
    hdr = ["DATA/HORA", "USUÁRIO", "REGISTRO (ID)", "CAMPO", "VALOR ANTERIOR", "VALOR NOVO", "ORIGEM"]
    header_row(ws, 1, 1, hdr)
    line2id = {}
    for i, rec in enumerate(recs):
        row = R1 + i
        y = rec["data"][:4] if rec["data"] else "2026"
        line2id[rec["linha_csv"]] = f"TT-{y}-{row - 1:05d}"
    now = dt.datetime.now().replace(microsecond=0)
    for i, (line, campo, de, para) in enumerate(mig_log):
        r = 2 + i
        ws.cell(row=r, column=1, value=now).number_format = "dd/mm/yyyy hh:mm"
        ws.cell(row=r, column=2, value="MIGRAÇÃO (script)")
        ws.cell(row=r, column=3, value=line2id.get(line, f"CSV linha {line}" if line else "GERAL"))
        ws.cell(row=r, column=4, value=campo)
        ws.cell(row=r, column=5, value=str(de))
        ws.cell(row=r, column=6, value=str(para))
        ws.cell(row=r, column=7, value=f"CSV original, linha {line}" if line else "CSV original")
        for c in range(1, 8):
            ws.cell(row=r, column=c).font = Font(name=FONT, size=9)
    widths(ws, {"A": 17, "B": 18, "C": 16, "D": 18, "E": 40, "F": 40, "G": 24})
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{max(2, 1 + len(mig_log))}"
    # área liberada para a macro gravar
    for r in range(2 + len(mig_log), 2 + len(mig_log) + 5):
        for c in range(1, 8):
            ws.cell(row=r, column=c).protection = UNLOCK


# ================================================================ COMO USAR
def build_como_usar(wb, ws):
    banner(ws, "COMO USAR — Troca de Turno 2026", "Guia rápido para quem lança e para quem gerencia.", 3)
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 120
    blocos = [
        ("PRINCÍPIO", "UMA LINHA = UMA ATIVIDADE / OCORRÊNCIA / ENTREGA / PENDÊNCIA. A BASE_OPERACIONAL é a única fonte da verdade; "
                      "todas as outras abas são calculadas a partir dela."),
        ("CORES", "AMARELO = célula de preenchimento. CINZA = calculado automaticamente (bloqueado). Cabeçalho AZUL-ESCURO = campo de entrada; "
                  "cabeçalho CINZA = automático. 🔴 crítico / 🟡 atenção ou pendente / 🟢 concluído."),
        ("QUEM LANÇA (líder do turno)", "1) Abra LANÇAMENTO (com macro) ou vá direto à primeira linha vazia da BASE_OPERACIONAL.\n"
                                        "2) Preencha: Data, Turno, Responsável, Tipo de atividade, Programação, Rota, Medida, Planejado, Executado, "
                                        "Faturamento, Embarque, Volume faturado, Transportadora.\n"
                                        "3) Se ficou algo pendente: PENDÊNCIA + TIPO + RESPONSÁVEL PELA TRATATIVA + PRAZO (+ CAUSA, AÇÃO e IMPACTO). "
                                        "Sem responsável ou prazo a linha acusa ERRO DE CADASTRO.\n"
                                        "4) Quando resolver: preencha DATA DE CONCLUSÃO (e mude Faturamento/Embarque/Executado, se for o caso).\n"
                                        "5) Nunca apague linhas: use CANCELADO? = SIM. Não ordene a base (o ID depende da posição)."),
        ("PASSAGEM DE TURNO", "Abra PASSAGEM_DE_TURNO: por padrão mostra o último turno lançado. Traz críticos, atenção, o que foi concluído, "
                              "o que fica para o próximo turno, prazos do dia, responsáveis e um RESUMO pronto para copiar (WhatsApp/e-mail)."),
        ("STATUS AUTOMÁTICO", "CANCELADO: CANCELADO? = SIM.\n"
                              "CONCLUÍDO: nada em aberto (executado ≥ planejado, faturamento e embarque concluídos/não aplicáveis, pendência com data de conclusão).\n"
                              "CRÍTICO (item aberto): prazo vencido; impacto CRÍTICO; embarque atrasado; pendência de SISTEMA com impacto ≥ ALTO; "
                              "pendência aberta com impacto ALTO que afeta o próximo turno.\n"
                              "ATENÇÃO (item aberto): pendência aberta; embarque realizado com faturamento pendente; prazo vence hoje ou nas próximas N horas; "
                              "SLA estourado; impacto ≥ ALTO; erro de cadastro/inconsistência/divergência.\n"
                              "PENDENTE: planejado sem execução.  EM ANDAMENTO: execução parcial ou faturamento/embarque em curso."),
        ("REGRAS DE VALIDAÇÃO", "1) Embarque REALIZADO + Faturamento PENDENTE → alerta EMBARCADO SEM FATURAMENTO.\n"
                                "2) Planejado > Executado com prazo vencido → CRÍTICO (PRAZO VENCIDO).\n"
                                "3) Pendência sem responsável → ERRO DE CADASTRO.\n"
                                "4) Pendência sem prazo → ERRO DE CADASTRO.\n"
                                "5) Data de conclusão informada com item ainda aberto → INCONSISTÊNCIA.\n"
                                "Extras: conclusão/prazo anterior à data, executado acima do planejado (+tolerância), pendência sem tipo."),
        ("PRAZO E SLA", "Situação do prazo: NO PRAZO / VENCE HOJE / VENCIDO / CONCLUÍDO / SEM PRAZO. Dias em aberto = hoje − data do registro. "
                        "Tempo de tratativa = conclusão − abertura (data + hora de início do turno). SLA por tipo em CONFIGURAÇÕES."),
        ("REGISTROS MIGRADOS", "Os 826 registros do CSV antigo foram importados com ORIGEM = MIGRAÇÃO (correções no LOG_ALTERAÇÕES). "
                               "Os anteriores à DATA DE CORTE (CONFIGURAÇÕES) contam nos indicadores e análises, mas não geram pendência aberta "
                               "(não tinham responsável nem prazo). Pendências migradas foram identificadas por palavras da observação "
                               "(AGUARDANDO, FALTA, PENDENTE…); a causa só foi preenchida quando o texto a indicava. Valores TONS em kg foram "
                               "convertidos para toneladas."),
        ("PAINÉIS", "DASHBOARD (executivo, com filtros), GESTÃO_DO_DIA, GESTÃO_SEMANAL, GESTÃO_MENSAL, PRODUTIVIDADE, FATURAMENTO, "
                    "EMBARQUES, ANÁLISE_CAUSAS (recorrência e matriz), PARETO, HISTÓRICO (consulta) e QUALIDADE_DADOS. "
                    "Filtros ficam nas células amarelas no topo de cada aba (TODOS = sem filtro)."),
        ("PROTEÇÃO", f"Abas protegidas sem bloquear o uso: só as células amarelas aceitam digitação. Senha: {SENHA} "
                     "(Revisar > Desproteger planilha) — use apenas para manutenção."),
        ("CAPACIDADE", f"A base comporta {NROWS} registros (≈ 1 ano no ritmo atual). Para ampliar: gere novamente com "
                       "scripts/gerar_planilha.py alterando NROWS, ou copie a última linha para baixo e estenda os nomes b_* "
                       "(Fórmulas > Gerenciador de Nomes)."),
        ("MACRO (opcional)", "Arquivos em vba/: modTrocaTurno.bas (formulário, carimbo de data/hora, log) e EstaPastaDeTrabalho.cls. "
                             "Alt+F11 → Arquivo > Importar (o .bas) e cole o conteúdo do .cls em 'EstaPasta_de_trabalho'. Salve como .xlsm. "
                             "A macro grava DATA/HORA DA ATUALIZAÇÃO e o LOG_ALTERAÇÕES (usuário, registro, campo, antes/depois)."),
        ("POWER BI", "Fato: tabela tbBase (BASE_OPERACIONAL; filtre ID não vazio). Dimensões: dCalendario (DIM_CALENDÁRIO) e as listas de "
                     "CADASTROS (turno, responsável, programação/área, transportadora, tipo de pendência, causa, impacto). "
                     "Relacione por DATA, TURNO, PROGRAMAÇÃO, TRANSPORTADORA, TIPO DE PENDÊNCIA e CAUSA. Detalhes no README do repositório."),
    ]
    r = 4
    for t, txt in blocos:
        c1 = ws.cell(row=r, column=2, value=t)
        st(c1, bold=True, color=WHITE, bg=NAVY, v="top", wrap=True)
        c2 = ws.cell(row=r, column=3, value=txt)
        st(c2, size=10, v="top", wrap=True, bg=LGREY)
        ws.row_dimensions[r].height = max(30, 15 * (txt.count("\n") + 1 + len(txt) // 120))
        r += 1


if __name__ == "__main__":
    main()
