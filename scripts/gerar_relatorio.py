import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

import sys
from pathlib import Path

# Gerador do relatório ZBR Importado.
# Uso: python scripts/gerar_relatorio.py [arquivo_origem.xlsx] [arquivo_saida.xlsx]
RAIZ = Path(__file__).resolve().parent.parent
SRC = sys.argv[1] if len(sys.argv) > 1 else str(RAIZ / 'planilhas' / 'origem' / 'Analise_ZBR_importado.XLSX')
OUT = sys.argv[2] if len(sys.argv) > 2 else str(RAIZ / 'planilhas' / 'ZBR_Importado_Relatorio_Gestao.xlsx')
DATA_RELATORIO = pd.Timestamp('2026-10-02').to_pydatetime()  # data de geração do relatório
MAXR = 5000  # linhas cobertas pelas fórmulas nas bases (espaço para novos registros)

txt = {'Lote': str, 'Depósito': str, 'Tipo de movimento': str, 'Doc.material': str,
       'Referência': str, 'Cliente': str, 'Duração': str, 'Tipo de depósito': str,
       'Posição no depósito': str, 'Centro': str}
x = pd.read_excel(SRC, sheet_name=None, dtype=txt)
ent = x['Entrada ']
sai = x['Saida']
sai = sai[sai['Material'].notna()].reset_index(drop=True)  # remove linhas de total do rodapé
est = x['Estoque atual']
est = est.drop(columns=[c for c in est.columns if est[c].isna().all()])  # colunas 100% vazias (ex.: Inventário ativo)

F = 'Calibri'
# Identidade visual corporativa (paleta definida no projeto de padronização)
NAVY, CORP, LIGHT, WHITE_C, ZEBRA, DARK = '17365D', '245A81', 'D9EAF7', 'FFFFFF', 'F2F4F7', '404854'
OK_C, ATT_C, CRIT_C, GREY_C = '70AD47', 'FFC000', 'C00000', 'A6A6A6'
INK, INK2, INK3 = DARK, DARK, '6B7480'
SURF, SURF2, LINE, GRID, PAGE = WHITE_C, LIGHT, 'C9D3DE', ZEBRA, WHITE_C
ACCENT, GOLD, C_ENT, C_SAI, C_REC = NAVY, CORP, CORP, '9DB9D6', NAVY
G_TX, G_BG, W_TX, W_BG, W_BAR, R_TX, R_BG = '375623', 'E2EFDA', '7F6000', 'FFF2CC', ATT_C, CRIT_C, 'F8DADA'
G_BAR = OK_C
FN = 'Calibri'
TITLE = Font(name=FN, size=22, bold=True, color=NAVY)
SUB = Font(name=F, size=10, color=INK3)
H = Font(name=F, size=10, bold=True, color='FFFFFF')
HFILL = PatternFill('solid', fgColor=NAVY)
SEC = Font(name=FN, size=13, bold=True, color=CORP)
BASE = Font(name=F, size=10, color=INK)
BOLD = Font(name=F, size=10, bold=True, color=INK)
INPUT = PatternFill('solid', fgColor=LIGHT)
BLUE = Font(name=F, size=10, bold=True, color=ACCENT)
GREEN = BASE
thin = Side(style='thin', color=GRID)
BOX = Border(bottom=thin)
INBOX = Border(left=Side(style='thin', color=CORP), right=Side(style='thin', color=CORP),
               top=Side(style='thin', color=CORP), bottom=Side(style='thin', color=CORP))
NUM = '#,##0.0;-#,##0.0;0'
RED_F = PatternFill('solid', fgColor=R_BG)
YEL_F = PatternFill('solid', fgColor=W_BG)
GRN_F = PatternFill('solid', fgColor=G_BG)
WHITE = PatternFill('solid', fgColor=SURF)

wb = Workbook()


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=str(c).upper() if c else c)
        cell.font = Font(name=F, size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.border = Border(bottom=Side(style='medium', color=CORP))
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def dump(ws, df, extra_cols=()):
    cols = list(df.columns) + list(extra_cols)
    header(ws, 1, cols)
    for r, row in enumerate(df.itertuples(index=False), 2):
        for c, v in enumerate(row, 1):
            if pd.isna(v) if not isinstance(v, str) else False:
                v = None
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BASE
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = NUM
            elif hasattr(v, 'year') and hasattr(v, 'hour') and cols[c - 1].startswith('Data'):
                cell.number_format = 'DD/MM/YYYY'
    ws.freeze_panes = 'A2'
    for i, c in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(str(c)) + 4))
    return len(df) + 1


# ---------------- Resumo (primeira aba) ----------------
res = wb.active
res.title = 'Resumo'

# ---------------- Abas de dados ----------------
wsE = wb.create_sheet('Entrada')
lastE = dump(wsE, ent.rename(columns={'Material': 'SKU', 'UM básica': 'UM'}), extra_cols=('Movimentação', 'Entrou no ZBR'))
nE = len(ent.columns)
cK, cL = get_column_letter(nE + 1), get_column_letter(nE + 2)
for r in range(2, lastE + 1):
    wsE[f'{cK}{r}'] = (f'=IF(AND(G{r}>0,F{r}<=Resumo!$C$6),"Recebido no 2008",'
                       f'IF(G{r}<0,"Entrada","Estorno"))')
    wsE[f'{cL}{r}'] = f'=IF({cK}{r}="Recebido no 2008",0,-G{r})'
    wsE[f'{cK}{r}'].font = BASE
    wsE[f'{cL}{r}'].font = BASE
    wsE[f'{cL}{r}'].number_format = NUM
wsE.column_dimensions[cK].width = 30
wsE.column_dimensions[cL].width = 16
wsE.auto_filter.ref = f'A1:{cL}{lastE}'

wsS = wb.create_sheet('Saida')
lastS = dump(wsS, sai.rename(columns={'Material': 'SKU', 'Texto breve de material': 'Descrição', 'UM básica': 'UM'}), extra_cols=('Saiu',))
cSq = get_column_letter(len(sai.columns) + 1)
for r in range(2, lastS + 1):
    wsS[f'{cSq}{r}'] = f'=-H{r}'
    wsS[f'{cSq}{r}'].font = BASE
    wsS[f'{cSq}{r}'].number_format = NUM
wsS.auto_filter.ref = f'A1:{cSq}{lastS}'

wsT = wb.create_sheet('Estoque físico')
est_cols = ['Material', 'Lote', 'Tipo de depósito', 'Posição no depósito'] + [c for c in est.columns if c not in ('Material', 'Lote', 'Tipo de depósito', 'Posição no depósito')]
lastT = dump(wsT, est[est_cols].rename(columns={'Material': 'SKU', 'Estoque total': 'Estoque físico', 'UM básica': 'UM'}))
wsT.auto_filter.ref = f'A1:P{lastT}'
wsT.freeze_panes = 'C2'

# ranges
E_MAT, E_LOT, E_Q, E_CL, E_ZBR = (f"Entrada!$B$2:$B${MAXR}", f"Entrada!$C$2:$C${MAXR}",
                                   f"Entrada!$G$2:$G${MAXR}", f"Entrada!${cK}$2:${cK}${MAXR}",
                                   f"Entrada!${cL}$2:${cL}${MAXR}")
S_MAT, S_LOT, S_Q = f"Saida!$B$2:$B${MAXR}", f"Saida!$C$2:$C${MAXR}", f"Saida!${cSq}$2:${cSq}${MAXR}"
T_MAT, T_LOT, T_Q = (f"'Estoque físico'!$A$2:$A${MAXR}", f"'Estoque físico'!$B$2:$B${MAXR}",
                     f"'Estoque físico'!$J$2:$J${MAXR}")
EB = lambda col: f"Entrada!${col}$2:${col}${MAXR}"  # noqa: E731
SB = lambda col: f"Saida!${col}$2:${col}${MAXR}"  # noqa: E731

# chaves Material x Lote (união das 3 abas)
um = {}
for df in (ent, sai, est):
    for m, l, u in df[['Material', 'Lote', 'UM básica']].itertuples(index=False):
        um.setdefault((m, l), u)
keys = sorted(um)

# Ordena pela diferença absoluta (calculada só para ordenar; a planilha recalcula por fórmula)
rec = (ent['Quantidade'] > 0) & (ent['Data de lançamento'] <= '2026-09-15')
zbr = ent.assign(z=-ent['Quantidade'].where(~rec, 0))


def dif(k):
    m, l = k
    e = zbr[(zbr.Material == m) & (zbr.Lote == l)].z.sum()
    s = -sai[(sai.Material == m) & (sai.Lote == l)].Quantidade.sum()
    t = est[(est.Material == m) & (est.Lote == l)]['Estoque total'].sum()
    return t - (e - s)


keys.sort(key=lambda k: (-round(abs(dif(k)), 4), k))
mats = sorted({m for m, _ in keys})


def mdif(m):
    return sum(dif(k) for k in keys if k[0] == m)


mats.sort(key=lambda m: (-round(abs(mdif(m)), 4), m))

COLS = ['Recebido importação (2008)', 'Entradas ZBR (2008 → ZBR)', 'Saídas (601)',
        'Saldo calculado (Entradas − Saídas)', 'Estoque atual', 'Diferença (Estoque − Calculado)',
        'Status', 'Possível causa', 'Recebido no 2008 ainda não transferido (Recebido − Entradas ZBR)']


def cross_sheet(ws, title, rows, by_lot):
    ws['A1'] = title
    ws['A1'].font = TITLE
    ws['A2'] = ('Entradas = transferências do depósito 2008 para o ZBR (sinal invertido, estornos já abatidos). '
                'Saídas = movimento 601. Diferença ≠ 0 indica estoque inicial anterior ao período ou movimento '
                'não contido nas abas.')
    ws['A2'].font = SUB
    keycols = ['SKU', 'Lote', 'UM'] if by_lot else ['SKU', 'UM']
    hr = 4
    widths = ([20, 13, 6] if by_lot else [20, 6]) + [15, 15, 13, 17, 13, 17, 26, 44, 18]
    header(ws, hr, keycols + COLS, widths)
    ws.row_dimensions[hr].height = 45
    o = len(keycols)
    L = lambda i: get_column_letter(o + i)  # noqa: E731
    for r, k in enumerate(rows, hr + 1):
        if by_lot:
            m, l = k
            vals = [m, l, um[k]]
            crit_e = f'{E_MAT},$A{r},{E_LOT},$B{r}'
            crit_s = f'{S_MAT},$A{r},{S_LOT},$B{r}'
            crit_t = f'{T_MAT},$A{r},{T_LOT},$B{r}'
        else:
            m = k
            vals = [m, next(u for (mm, _), u in um.items() if mm == m)]
            crit_e, crit_s, crit_t = f'{E_MAT},$A{r}', f'{S_MAT},$A{r}', f'{T_MAT},$A{r}'
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v).font = BASE
        f = {
            1: f'=SUMIFS({E_Q},{crit_e},{E_CL},"Recebido no 2008")',
            2: f'=SUMIFS({E_ZBR},{crit_e})',
            3: f'=SUMIFS({S_Q},{crit_s})',
            4: f'={L(2)}{r}-{L(3)}{r}',
            5: f'=SUMIFS({T_Q},{crit_t})',
            6: f'=ROUND({L(5)}{r}-{L(4)}{r},3)',
            7: (f'=IF(ABS({L(6)}{r})<=Resumo!$C$7,"OK",IF({L(6)}{r}>0,'
                f'"Sobra no estoque","Falta no estoque"))'),
            8: (f'=IF({L(7)}{r}="OK","",IF(AND({L(2)}{r}=0,{L(3)}{r}=0),"Estoque sem movimento no período",'
                f'IF({L(4)}{r}<0,"Saídas maiores que entradas – estoque anterior ao período?",'
                f'IF({L(5)}{r}=0,"Saldo calculado sem estoque físico – verificar",'
                f'IF({L(6)}{r}>0,"Estoque acima do esperado – saldo anterior ou entrada não listada",'
                f'"Estoque abaixo do esperado – saída/ajuste não listado")))))'),
            9: f'=IF({L(1)}{r}=0,"",{L(1)}{r}-{L(2)}{r})',
        }
        for i, fx in f.items():
            cell = ws.cell(row=r, column=o + i, value=fx)
            cell.font = BOLD if i == 6 else BASE
            if i not in (7, 8):
                cell.number_format = NUM
        for c in range(1, o + 10):
            ws.cell(row=r, column=c).border = BOX
    last = hr + len(rows)
    tr = last + 1
    ws.cell(row=tr, column=1, value='TOTAL (unidades mistas – ver Resumo por UM)').font = BOLD
    for i in (1, 2, 3, 4, 5, 6, 9):
        c = L(i)
        cell = ws.cell(row=tr, column=o + i, value=f'=SUM({c}{hr + 1}:{c}{last})')
        cell.font, cell.number_format, cell.border = BOLD, NUM, BOX
    st = L(7)
    rng = f'A{hr + 1}:{L(9)}{last}'
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${st}{hr + 1}="Falta no estoque"'], fill=RED_F, font=Font(color=R_TX)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${st}{hr + 1}="Sobra no estoque"'], fill=YEL_F, font=Font(color=W_TX)))
    ws.conditional_formatting.add(f'{st}{hr + 1}:{st}{last}',
                                  FormulaRule(formula=[f'${st}{hr + 1}="OK"'], fill=GRN_F, font=Font(color=G_TX)))
    ws.freeze_panes = ws.cell(row=hr + 1, column=o + 1)
    ws.auto_filter.ref = f'A{hr}:{L(9)}{last}'
    return hr, last, o


wsL = wb.create_sheet('Cruzamento por Lote', 1)
hrL, lastL, oL = cross_sheet(wsL, 'Cruzamento Entrada × Saída × Estoque — por Material e Lote', keys, True)
wsM = wb.create_sheet('Cruzamento por Material', 1)
hrM, lastM, oM = cross_sheet(wsM, 'Cruzamento Entrada × Saída × Estoque — por Material', mats, False)

# ---------------- Resumo ----------------
res['A1'] = 'Análise ZBR Importado — Entradas × Saídas × Estoque atual'
res['A1'].font = TITLE
res['A2'] = ('Fonte: arquivo Analise_ZBR_importado.XLSX (abas Entrada, Saida e Estoque atual), '
             'período 09/09/2026 a 01/10/2026, centro BR01.')
res['A2'].font = SUB
res.column_dimensions['A'].width = 4
res.column_dimensions['B'].width = 46
for c in 'CDEFGH':
    res.column_dimensions[c].width = 17

res['B4'] = 'Parâmetros (células amarelas, editáveis)'
res['B4'].font = SEC
res['B6'] = 'Data limite dos recebimentos de importação no 2008'
res['C6'] = pd.Timestamp('2026-09-15').to_pydatetime()
res['C6'].number_format = 'DD/MM/YYYY'
res['C6'].comment = Comment('Lançamentos POSITIVOS no depósito 2008 até esta data são tratados como chegada da '
                            'importação (ficam no 2008, não entram no ZBR). Positivos depois desta data são '
                            'estornos de transferência. Premissa inferida dos dados (todos os positivos até '
                            '15/09 são recebimentos em lote; os posteriores anulam transferências).', 'Claude')
res['B7'] = 'Tolerância para considerar diferença OK'
res['C7'] = 0.01
res['C7'].number_format = '0.00'
for c in ('C6', 'C7'):
    res[c].fill, res[c].font, res[c].border = INPUT, BLUE, INBOX
for c in ('B6', 'B7'):
    res[c].font = BASE

res['B9'] = 'Situação dos cruzamentos'
res['B9'].font = SEC
header_cells = ['', 'Itens', 'OK', 'Sobra no estoque', 'Falta no estoque', '% OK']
for i, h in enumerate(header_cells):
    if i == 0:
        continue
    cell = res.cell(row=10, column=1 + i, value=h if i > 1 else 'Visão')
    cell.font, cell.fill, cell.border = H, HFILL, BOX
    cell.alignment = Alignment(horizontal='center')
res.cell(row=10, column=2, value='Visão')
for r, (name, sh, hr, last, o) in enumerate(
        [('Material + Lote', 'Cruzamento por Lote', hrL, lastL, oL),
         ('Material', 'Cruzamento por Material', hrM, lastM, oM)], 11):
    st = f"'{sh}'!${get_column_letter(o + 7)}${hr + 1}:${get_column_letter(o + 7)}${last}"
    res.cell(row=r, column=2, value=name)
    res.cell(row=r, column=3, value=f"=COUNTA('{sh}'!$A${hr + 1}:$A${last})")
    res.cell(row=r, column=4, value=f'=COUNTIF({st},"OK")')
    res.cell(row=r, column=5, value=f'=COUNTIF({st},"Sobra no estoque")')
    res.cell(row=r, column=6, value=f'=COUNTIF({st},"Falta no estoque")')
    res.cell(row=r, column=7, value=f'=IF(C{r}=0,0,D{r}/C{r})')
    res.cell(row=r, column=7).number_format = '0.0%'
    for c in range(2, 8):
        res.cell(row=r, column=c).border = BOX
        res.cell(row=r, column=c).font = GREEN if c in (3, 4, 5, 6) else BASE

res['B14'] = 'Totais por unidade de medida'
res['B14'].font = SEC
uh = ['UM', 'Entradas ZBR', 'Saídas (601)', 'Saldo calculado', 'Estoque atual', 'Diferença']
for i, h in enumerate(uh, 2):
    cell = res.cell(row=15, column=i, value=h)
    cell.font, cell.fill, cell.border = H, HFILL, BOX
    cell.alignment = Alignment(horizontal='center')
L_UM = f"'Cruzamento por Lote'!$C${hrL + 1}:$C${lastL}"
lc = lambda i: f"'Cruzamento por Lote'!${get_column_letter(oL + i)}${hrL + 1}:${get_column_letter(oL + i)}${lastL}"  # noqa
for r, u in enumerate(['KG', 'UN', 'CX'], 16):
    res.cell(row=r, column=2, value=u)
    for c, i in zip(range(3, 8), (2, 3, 4, 5, 6)):
        res.cell(row=r, column=c, value=f'=SUMIFS({lc(i)},{L_UM},$B{r})').number_format = NUM
    for c in range(2, 8):
        res.cell(row=r, column=c).border = BOX
        res.cell(row=r, column=c).font = GREEN if c > 2 else BASE

res['B20'] = 'Movimentação das abas de origem'
res['B20'].font = SEC
mh = ['Indicador', 'Linhas', 'Quantidade']
for i, h in enumerate(mh, 2):
    cell = res.cell(row=21, column=i, value=h)
    cell.font, cell.fill, cell.border = H, HFILL, BOX
mv = [
    ('Recebimentos de importação no 2008', f'=COUNTIF({E_CL},"Recebido no 2008")',
     f'=SUMIFS({E_Q},{E_CL},"Recebido no 2008")'),
    ('Transferências 2008 → ZBR', f'=COUNTIF({E_CL},"Entrada")',
     f'=-SUMIFS({E_Q},{E_CL},"Entrada")'),
    ('Estornos de transferência', f'=COUNTIF({E_CL},"Estorno")',
     f'=SUMIFS({E_Q},{E_CL},"Estorno")'),
    ('Saídas (601)', f'=COUNTA(Saida!$A$2:$A${MAXR})', f'=SUM({S_Q})'),
    ('Posições em estoque', f"=COUNTA('Estoque físico'!$A$2:$A${MAXR})", f'=SUM({T_Q})'),
]
for r, (a, b, c) in enumerate(mv, 22):
    res.cell(row=r, column=2, value=a).font = BASE
    res.cell(row=r, column=3, value=b).font = GREEN
    cc = res.cell(row=r, column=4, value=c)
    cc.font, cc.number_format = GREEN, NUM
    for col in (2, 3, 4):
        res.cell(row=r, column=col).border = BOX
res['B27'] = 'Quantidades somam KG, UN e CX juntos — use a tabela por UM para comparar.'
res['B27'].font = SUB

res['B29'] = 'Como ler'
res['B29'].font = SEC
notes = [
    'Depósito 2008 = recebimento da importação. Valores NEGATIVOS no 2008 são transferências para o ZBR '
    '(depósitos 9999/0355) e entram aqui como Entradas.',
    'Valores POSITIVOS no 2008 até a data-limite (C6) são chegadas da importação; depois dela, estornos que '
    'abatem as Entradas.',
    'Saldo calculado = Entradas ZBR − Saídas (601). Diferença = Estoque atual − Saldo calculado.',
    'Sobra no estoque: há mais estoque que o movimento explica (provável saldo anterior a 09/09 ou entrada não '
    'listada). Falta no estoque: há menos estoque que o esperado (saída, ajuste ou transferência não listada).',
    'Saldo em 2008: quantidade recebida da importação que ainda não foi transferida para o ZBR no período.',
    'As linhas de total do rodapé da aba Saida original foram removidas para não duplicar quantidades.',
]
for r, n in enumerate(notes, 30):
    res.cell(row=r, column=2, value='• ' + n).font = BASE
    res.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    res.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical='top')
    res.row_dimensions[r].height = 30


# ---------------- Movimentações (tabela única filtrável) ----------------
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

desc = sai.dropna(subset=['Texto breve de material']).groupby('Material')['Texto breve de material'].first().to_dict()
mov = []
for i, r in ent.iterrows():
    mov.append(('E', i + 2, r['Data de lançamento'], r['Material'], r['Lote'], r['Depósito'], r['Quantidade'],
                r['UM básica'], r['Doc.material'], None, None, None))
for i, r in sai.iterrows():
    mov.append(('S', i + 2, r['Data de lançamento'], r['Material'], r['Lote'], r['Depósito'], r['Quantidade'],
                r['UM básica'], r['Doc.material'], r['Cliente'], r['Nome 1'], r['Nome do usuário']))
mov.sort(key=lambda m: (m[2], m[3], m[4], m[0]))

wsV = wb.create_sheet('Lançamentos', 3)
MCOLS = ['Data', 'Movimentação', 'SKU', 'Lote', 'Descrição', 'Depósito', 'UM',
         'Qtd no SAP', 'Entrou', 'Saiu', 'Saldo (±)', 'Documento', 'Cód. cliente', 'Cliente', 'Usuário']
MW = [12, 17, 20, 13, 36, 9, 6, 12, 12, 12, 12, 13, 11, 36, 11]
header(wsV, 1, MCOLS, MW)
wsV.row_dimensions[1].height = 32
for n, (src, sr, dt, mat, lot, dep, q, u, doc, cli, nome, usr) in enumerate(mov, 2):
    vals = [dt.to_pydatetime(), None, mat, lot, desc.get(mat, ''), dep, u, None, None, None, None, doc, cli, nome, usr]
    for c, v in enumerate(vals, 1):
        cell = wsV.cell(row=n, column=c, value=v)
        cell.font = BASE
    if src == 'E':
        wsV.cell(row=n, column=2, value=f'=Entrada!{cK}{sr}')
        wsV.cell(row=n, column=8, value=f'=Entrada!G{sr}')
        wsV.cell(row=n, column=9, value=f'=Entrada!{cL}{sr}')
        wsV.cell(row=n, column=10, value=0)
    else:
        wsV.cell(row=n, column=2, value='Saída')
        wsV.cell(row=n, column=8, value=f'=Saida!H{sr}')
        wsV.cell(row=n, column=9, value=0)
        wsV.cell(row=n, column=10, value=f'=Saida!{cSq}{sr}')
    wsV.cell(row=n, column=11, value=f'=I{n}-J{n}')
    wsV.cell(row=n, column=1).number_format = 'DD/MM/YYYY'
    for c in range(2, 16):
        wsV.cell(row=n, column=c).font = GREEN if c in (2, 8, 9, 10) and src == 'E' or (c in (8, 10) and src == 'S') else BASE
    for c in (8, 9, 10, 11):
        wsV.cell(row=n, column=c).number_format = NUM
lastV = len(mov) + 1
tab = Table(displayName='Movimentacoes', ref=f'A1:O{lastV}')
tab.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=False)
wsV.add_table(tab)
wsV.freeze_panes = 'E2'

# ---------------- Consulta (filtros de período, material e lote) ----------------
wsC = wb.create_sheet('Consulta', 1)
wsC['A1'] = 'Consulta por período, material e lote'
wsC['A1'].font = TITLE
wsC['A2'] = 'Preencha as células amarelas. Use * para "todos". Os totais e a tabela abaixo recalculam sozinhos.'
wsC['A2'].font = SUB
wsC.column_dimensions['A'].width = 22
for c, w in zip('BCDEFGHI', [7, 15, 15, 15, 15, 15, 15, 20]):
    wsC.column_dimensions[c].width = w
inputs = [('Data inicial', pd.Timestamp('2026-09-01').to_pydatetime(), 'DD/MM/YYYY'),
          ('Data final', pd.Timestamp('2026-10-31').to_pydatetime(), 'DD/MM/YYYY'),
          ('Material (* = todos)', '*', '@'),
          ('Lote (* = todos)', '*', '@')]
for r, (lab, v, fmt) in enumerate(inputs, 4):
    wsC.cell(row=r, column=1, value=lab).font = BOLD
    cell = wsC.cell(row=r, column=3, value=v)
    cell.fill, cell.font, cell.border, cell.number_format = INPUT, BLUE, BOX, fmt
    wsC.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
wsC['C6'].comment = Comment('Escolha um material na lista ou digite * para todos. Aceita curingas: 811* traz todos que começam com 811.', 'Claude')
wsC['C7'].comment = Comment('Digite o lote exatamente como no SAP, ou * para todos.', 'Claude')
wsC['E4'] = 'Período padrão cobre todo o arquivo (09/09 a 01/10/2026).'
wsC['E4'].font = SUB
# lista de materiais para o dropdown (coluna K, oculta)
for i, m in enumerate(['*'] + sorted({m for m, _ in keys}), 1):
    wsC.cell(row=i, column=11, value=m)
wsC.column_dimensions['K'].hidden = True
nlist = len({m for m, _ in keys}) + 1
dv = DataValidation(type='list', formula1=f'=$K$1:$K${nlist}', allow_blank=False, showErrorMessage=False)
wsC.add_data_validation(dv)
dv.add('C6')

V = lambda col: f"'Lançamentos'!${col}$2:${col}${lastV}"  # noqa: E731
DATES = f'{V("A")},">="&$C$4,{V("A")},"<="&$C$5'
def crit(mat, lot):
    return f'{DATES},{V("C")},{mat},{V("D")},{lot}'

wsC['A9'] = 'Totais da seleção'
wsC['A9'].font = SEC
th = ['Indicador', '', 'KG', 'UN', 'CX']
for i, h in enumerate(th, 1):
    cell = wsC.cell(row=10, column=i, value=h)
    cell.font, cell.fill, cell.border = H, HFILL, BOX
inds = [('Recebido importação (2008)', 'REC'), ('Entradas ZBR', 'I'), ('Saídas ZBR', 'J'),
        ('Movimento líquido do período', 'K'), ('Estoque atual (independe da data)', 'EST'),
        ('Diferença (Estoque − líquido)', 'DIF'), ('Nº de lançamentos', 'CNT')]
for r, (lab, kind) in enumerate(inds, 11):
    wsC.cell(row=r, column=1, value=lab).font = BOLD if kind == 'DIF' else BASE
    for c, u in zip((3, 4, 5), ('KG', 'UN', 'CX')):
        cu = f'{V("G")},"{u}"'
        if kind == 'REC':
            f = f'=SUMIFS({V("H")},{crit("$C$6", "$C$7")},{cu},{V("B")},"Recebido no 2008")'
        elif kind in ('I', 'J', 'K'):
            f = f'=SUMIFS({V(kind)},{crit("$C$6", "$C$7")},{cu})'
        elif kind == 'EST':
            f = f"=SUMIFS({T_Q},{T_MAT},$C$6,{T_LOT},$C$7,'Estoque físico'!$K$2:$K${lastT},\"{u}\")"
        elif kind == 'DIF':
            f = f'={get_column_letter(c)}15-{get_column_letter(c)}14'
        else:
            f = f'=COUNTIFS({crit("$C$6", "$C$7")},{cu})'
        cell = wsC.cell(row=r, column=c, value=f)
        cell.font = BOLD if kind == 'DIF' else BASE
        cell.number_format = '#,##0' if kind == 'CNT' else NUM
    for c in range(1, 6):
        wsC.cell(row=r, column=c).border = BOX
wsC['A18'] = ('A diferença só é comparável com o período completo: o estoque atual é uma foto de hoje, '
              'enquanto entradas e saídas seguem o filtro de data.')
wsC['A18'].font = SUB

wsC['A20'] = 'Por material (no período e lote escolhidos)'
wsC['A20'].font = SEC
ch = ['Material', 'UM', 'Recebido 2008', 'Entradas ZBR', 'Saídas ZBR', 'Líquido período', 'Estoque atual',
      'Diferença', 'Status']
for i, h in enumerate(ch, 1):
    cell = wsC.cell(row=21, column=i, value=h)
    cell.font, cell.fill, cell.border = H, HFILL, BOX
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
mat_um = {}
for (m, _), u in um.items():
    mat_um.setdefault(m, u)
r0 = 22
for r, m in enumerate(sorted(mat_um), r0):
    wsC.cell(row=r, column=1, value=m)
    wsC.cell(row=r, column=2, value=mat_um[m])
    cr = crit(f'$A{r}', '$C$7')
    wsC.cell(row=r, column=3, value=f'=SUMIFS({V("H")},{cr},{V("B")},"Recebido no 2008")')
    wsC.cell(row=r, column=4, value=f'=SUMIFS({V("I")},{cr})')
    wsC.cell(row=r, column=5, value=f'=SUMIFS({V("J")},{cr})')
    wsC.cell(row=r, column=6, value=f'=D{r}-E{r}')
    wsC.cell(row=r, column=7, value=f'=SUMIFS({T_Q},{T_MAT},$A{r},{T_LOT},$C$7)')
    wsC.cell(row=r, column=8, value=f'=ROUND(G{r}-F{r},3)')
    wsC.cell(row=r, column=9, value=f'=IF(ABS(H{r})<=Resumo!$C$7,"OK",IF(H{r}>0,"Sobra no estoque","Falta no estoque"))')
    for c in range(1, 10):
        cell = wsC.cell(row=r, column=c)
        cell.font, cell.border = BASE, BOX
        if 3 <= c <= 8:
            cell.number_format = NUM
rL = r0 + len(mat_um) - 1
rng = f'A{r0}:I{rL}'
wsC.conditional_formatting.add(rng, FormulaRule(formula=[f'$I{r0}="Falta no estoque"'], fill=RED_F, font=Font(color=R_TX)))
wsC.conditional_formatting.add(rng, FormulaRule(formula=[f'$I{r0}="Sobra no estoque"'], fill=YEL_F, font=Font(color=W_TX)))
wsC.conditional_formatting.add(rng, FormulaRule(formula=[f'AND($C$6<>"*",$A{r0}=$C$6)'],
                                                font=Font(name=F, size=10, bold=True)))
wsC.auto_filter.ref = f'A21:I{rL}'
wsC.freeze_panes = 'A22'

# ---------------- Cruzamento por SKU e lote (leitura direta) ----------------
wsA = wb.create_sheet('Cruzamento', 0)
wsA.sheet_view.showGridLines = False
mat_um2 = {}
for (m, _), u in um.items():
    mat_um2.setdefault(m, u)
tol = 0.01
def lot_key(k):
    d = dif(k)
    if d < -tol:
        return (0, d, k)
    if d > tol:
        return (1, -d, k)
    return (2, 0, k)
lots_ord = sorted(keys, key=lot_key)
wsA['A1'] = 'Cruzamento por SKU e lote'
wsA['A1'].font = TITLE
wsA['A2'] = 'Uma linha para cada SKU + lote · movimentos de 09/09 a 01/10/2026'
wsA['A2'].font = SUB
for i, t in enumerate([
        'COMO LER',
        'ENTROU − SAIU = DEVERIA TER.  Depois compare com o ESTOQUE FÍSICO (o que tem no depósito hoje).',
        'DIFERENÇA = Estoque físico − Deveria ter.   0 = Confere  ·  positivo = Sobrando  ·  negativo = Faltando.',
        'MOTIVO = explicação calculada pelos dados: estoque antigo (posições sem movimento desde antes de 09/09), doca/transferência, troca de lote ou saída não lançada.'], 4):
    wsA.cell(row=i, column=1, value=t).font = Font(name=F, size=10, bold=(i == 4), color=INK if i == 4 else INK2)
HR = 13
CARD = [('CONFEREM', 'Confere', GRN_F, G_TX), ('SOBRANDO', 'Sobrando', YEL_F, W_TX), ('FALTANDO', 'Faltando', RED_F, R_TX)]
for j, (lab, key, fill, tx) in enumerate(CARD):
    col = 1 + j * 2
    for rr in (9, 10):
        wsA.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + 1)
        for cc in (col, col + 1):
            wsA.cell(row=rr, column=cc).fill = fill
    a = wsA.cell(row=9, column=col, value=f'SKU·LOTE {lab}')
    b = wsA.cell(row=10, column=col, value=f'=COUNTIF($J${HR + 1}:$J${HR + len(lots_ord)},"{key}")')
    a.font = Font(name=F, size=8, bold=True, color=INK2)
    b.font = Font(name=FN, size=24, bold=True, color=tx)
    for cc in (a, b):
        cc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
wsA.row_dimensions[10].height = 34
AH = ['SKU', 'Lote', 'Descrição', 'UM', 'Entrou', 'Saiu', 'Deveria ter', 'Estoque físico', 'Diferença',
      'Situação', 'Motivo da diferença', 'Solução de ajuste (sugestão)', 'Estoque antigo (antes de 09/09)',
      'Em doca / chão / transferência', 'Diferença sem explicação', 'Saídas pelo depósito 0355']
header(wsA, HR, AH, [20, 13, 30, 6, 12, 12, 13, 14, 12, 12, 66, 80, 15, 15, 15, 15])
LV = lambda col: f"'Lançamentos'!${col}$2:${col}${lastV}"  # noqa: E731
ES = lambda col: f"'Estoque físico'!${col}$2:${col}${MAXR}"  # noqa: E731
L0 = HR + 1
L1 = HR + len(lots_ord)
wsA.row_dimensions[HR].height = 24
CL = "'Cruzamento por Lote'"
cl = lambda col: f"{CL}!${col}${hrL + 1}:${col}${lastL}"  # noqa: E731
for r, (m, lt) in enumerate(lots_ord, HR + 1):
    wsA.cell(row=r, column=1, value=m)
    wsA.cell(row=r, column=2, value=lt)
    wsA.cell(row=r, column=3, value=desc.get(m, ''))
    wsA.cell(row=r, column=4, value=um[(m, lt)])
    crit = f'{cl("A")},$A{r},{cl("B")},$B{r}'
    wsA.cell(row=r, column=5, value=f'=SUMIFS({cl("E")},{crit})')
    wsA.cell(row=r, column=6, value=f'=SUMIFS({cl("F")},{crit})')
    wsA.cell(row=r, column=7, value=f'=E{r}-F{r}')
    wsA.cell(row=r, column=8, value=f'=SUMIFS({cl("H")},{crit})')
    wsA.cell(row=r, column=9, value=f'=ROUND(H{r}-G{r},2)')
    wsA.cell(row=r, column=10, value=f'=IF(ABS(I{r})<=Resumo!$C$7,"Confere",IF(I{r}>0,"Sobrando","Faltando"))')
    wsA.cell(row=r, column=13, value=f'=SUMIFS({ES("J")},{ES("A")},$A{r},{ES("B")},$B{r},{ES("I")},"<"&DATE(2026,9,9))')
    wsA.cell(row=r, column=14, value=(f'=SUMPRODUCT(({ES("A")}=$A{r})*({ES("B")}=$B{r})*({ES("I")}>=DATE(2026,9,9))'
                                      f'*ISNUMBER(MATCH({ES("C")},{{"DCK","TRF","922","FLR","DIF"}},0))*{ES("J")})'))
    troca = (f'AND(J{r}<>"Confere",SUMPRODUCT(($A${L0}:$A${L1}=$A{r})*($B${L0}:$B${L1}<>$B{r})'
             f'*(ABS($I${L0}:$I${L1}+$I{r})<=Resumo!$C$7))>0)')
    tol = 'Resumo!$C$7'
    wsA.cell(row=r, column=11, value=(
        f'=IF(J{r}="Confere","Sem diferença",'
        f'IF({troca},"Troca de lote: a mesma quantidade aparece com sinal oposto em outro lote deste SKU",'
        f'IF(AND(I{r}<0,H{r}=0,F{r}=0),"Entrou no ZBR, mas não está no estoque físico: não chegou ou foi guardado em outro lote",'
        f'IF(I{r}<0,"Estoque físico menor que o esperado: saída, ajuste ou perda não lançada no período",'
        f'IF(ABS(I{r}-M{r})<={tol},"Estoque antigo: toda a sobra está em posições sem movimento desde antes de 09/09",'
        f'IF(AND(F{r}>0,ABS(I{r}-F{r})<={tol}),"As saídas do período não baixaram o estoque físico",'
        f'IF(ABS(I{r}-N{r})<={tol},"Quantidade parada em doca, chão ou transferência: entrega ou devolução não baixada",'
        f'IF(ABS(I{r}-M{r}-N{r})<={tol},"Estoque antigo + quantidade em doca/chão/transferência",'
        f'IF(AND(F{r}>0,ABS(I{r}-M{r}-N{r}-F{r})<={tol}),IF(M{r}+N{r}>0,"Estoque antigo/doca + saídas do período que","Saídas do período que")&" não baixaram o estoque físico",'
        f'IF(AND(P{r}>0,ABS(I{r}-M{r}-N{r}-P{r})<={tol}),IF(M{r}+N{r}>0,"Estoque antigo/doca + saídas","Saídas")&" pelo depósito 0355 que não baixaram o estoque físico",'
        f'IF(M{r}+N{r}>0,"Explica em parte: estoque antigo "&FIXED(M{r},1)&" + doca/transferência "&FIXED(N{r},1)&"; faltam "&FIXED(I{r}-M{r}-N{r},1)&" a explicar",'
        f'IF(F{r}>E{r},"Saiu mais do que entrou: havia estoque antes de 09/09 que não está no arquivo",'
        f'"Estoque acima do esperado: entrada não lançada ou saldo anterior ao período"))))))))))))'))
    wsA.cell(row=r, column=16, value=f'=SUMIFS({SB(cSq)},{SB("B")},$A{r},{SB("C")},$B{r},{SB("E")},"0355")')
    q = lambda x: f'FIXED(ABS({x}),1)&" "&D{r}'  # noqa: E731
    wsA.cell(row=r, column=12, value=(
        f'=IF(J{r}="Confere","Nenhum ajuste necessário",'
        f'IF(LEFT(K{r},5)="Troca","Transferir "&{q(f"I{r}")}&" entre os lotes deste SKU (transferência lote a lote, mov. 309) para corrigir o lote",'
        f'IF(LEFT(K{r},6)="Entrou","Procurar a OT de armazenagem pendente no ZBR. Se não achar: contar e lançar ajuste de inventário de −"&{q(f"I{r}")},'
        f'IF(J{r}="Faltando","Contar o lote. Se a falta se confirmar: lançar a saída esquecida (601) ou ajuste de inventário de −"&{q(f"I{r}")},'
        f'IF(LEFT(K{r},15)="Estoque antigo:","Confirmar no SAP o saldo de 08/09 ("&{q(f"M{r}")}&"). Se estiver certo, não há ajuste",'
        f'IF(LEFT(K{r},10)="Quantidade","Confirmar ou estornar as OTs em doca/transferência ("&{q(f"N{r}")}&"): guardar na posição ou baixar a entrega",'
        f'IF(LEFT(K{r},16)="Estoque antigo +","Confirmar saldo de 08/09 ("&{q(f"M{r}")}&") e tratar as OTs em doca/transferência ("&{q(f"N{r}")}&")",'
        f'IF(ISNUMBER(SEARCH("não baixaram",K{r})),"Confirmar a OT de saída (picking) das remessas para baixar "&{q(f"I{r}-M{r}-N{r}")}&" do estoque físico"&IF(M{r}+N{r}>0,"; conferir também estoque antigo/doca",""),'
        f'IF(LEFT(K{r},7)="Explica","Confirmar saldo de 08/09 e OTs de doca; contar o lote para os "&{q(f"I{r}-M{r}-N{r}")}&" restantes e ajustar a sobra se confirmada",'
        f'"Confirmar saldo de 08/09; se não existir, contar o lote e lançar ajuste de inventário de +"&{q(f"I{r}")})))))))))'))
    wsA.cell(row=r, column=15, value=(f'=IF(OR(J{r}="Confere",LEFT(K{r},5)="Troca",ISNUMBER(SEARCH("não baixaram",K{r}))),0,'
                                      f'IF(I{r}>0,MAX(0,ROUND(I{r}-M{r}-N{r},2)),I{r}))'))
    for c in range(1, 17):
        cell = wsA.cell(row=r, column=c)
        cell.font = Font(name=F, size=10, bold=(c in (1, 2, 9)), color=INK)
        cell.border = BOX
        if 5 <= c <= 9 or 13 <= c <= 16:
            cell.number_format = '+#,##0.0;-#,##0.0;0' if c in (9, 15) else NUM
        if c == 10:
            cell.alignment = Alignment(horizontal='center')
    wsA.cell(row=r, column=3).font = Font(name=F, size=9, color=INK2)
    wsA.cell(row=r, column=11).font = Font(name=F, size=10, bold=True, color=INK)
    wsA.cell(row=r, column=12).font = Font(name=F, size=9, color=INK2)
lastA = HR + len(lots_ord)
for lab, fill, tx in (('Confere', GRN_F, G_TX), ('Sobrando', YEL_F, W_TX), ('Faltando', RED_F, R_TX)):
    wsA.conditional_formatting.add(f'J{HR + 1}:J{lastA}', FormulaRule(formula=[f'$J{HR + 1}="{lab}"'], fill=fill, font=Font(color=tx, bold=True)))
wsA.conditional_formatting.add(f'I{HR + 1}:I{lastA}', FormulaRule(formula=[f'I{HR + 1}<-Resumo!$C$7'], font=Font(color=R_TX, bold=True)))
wsA.auto_filter.ref = f'A{HR}:P{lastA}'
wsA.conditional_formatting.add(f'O{HR + 1}:O{lastA}', FormulaRule(formula=[f'ABS(O{HR + 1})>Resumo!$C$7'], font=Font(color=R_TX, bold=True)))
wsA.freeze_panes = f'C{HR + 1}'
wsA.cell(row=lastA + 2, column=1, value='Ordem: primeiro FALTANDO, depois SOBRANDO (do maior para o menor) e por último CONFERE. '
        'Use as setas do cabeçalho para filtrar por SKU ou lote.').font = SUB
wsA.cell(row=lastA + 3, column=1, value='Para ver cada movimento de um SKU/lote, use a aba "Lançamentos". '
        'Abas cinza = dados originais do SAP.').font = SUB
wb.active = 0

# ================= PAINEL (mesmo layout do painel HTML) =================
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule
del wb['Consulta']
wsP = wb.create_sheet('Painel', 0)
wsP.sheet_properties.tabColor = ACCENT
wsP.sheet_view.showGridLines = False
WID = {'A': 20, 'B': 13, 'C': 7, 'D': 12, 'E': 12, 'F': 12, 'G': 12, 'H': 13, 'I': 12, 'J': 12, 'K': 11, 'L': 12, 'M': 12, 'N': 58, 'O': 70}
for c, w in WID.items():
    wsP.column_dimensions[c].width = w
for r in range(1, 115):
    for c in range(1, 16):
        wsP.cell(row=r, column=c).fill = PatternFill('solid', fgColor=PAGE)
LBL = Font(name=F, size=8, bold=True, color=INK3)

wsP['A1'] = 'Painel ZBR Importado'
wsP['A1'].font = TITLE
wsP['A2'] = 'Entradas, saídas e estoque físico por SKU e lote · centro BR01 · movimentos de 09/09 a 01/10/2026'
wsP['A2'].font = SUB
wsP.row_dimensions[1].height = 30

# --- barra de filtros (linhas 4-6) ---
def card(r1, c1, r2, c2, fill='FCFCFB', edge=LINE, width='thin'):
    sd = Side(style=width, color=edge)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = wsP.cell(row=r, column=c)
            cell.fill = PatternFill('solid', fgColor=fill)
            cell.border = Border(left=sd if c == c1 else None, right=sd if c == c2 else None,
                                 top=sd if r == r1 else None, bottom=sd if r == r2 else None)
card(4, 1, 7, 15)
FILT = [('A', 'C', 'SKU', 'Todos'), ('D', 'E', 'LOTE', 'Todos'),
        ('F', 'G', 'DATA INICIAL', pd.Timestamp('2026-09-09').to_pydatetime()),
        ('H', 'I', 'DATA FINAL', pd.Timestamp('2026-10-01').to_pydatetime()),
        ('J', 'J', 'UNIDADE', 'KG'), ('K', 'M', 'MOVIMENTAÇÃO', 'Todas')]
LINKF = {'A': '=Dashboard!E5', 'D': '=Dashboard!H5', 'F': '=Dashboard!A5', 'H': '=Dashboard!C5',
         'J': '=Dashboard!J5', 'K': '=Dashboard!K5'}
for a, b, lab, v in FILT:
    wsP[f'{a}4'] = lab
    wsP[f'{a}4'].font = LBL
    wsP[f'{a}5'] = LINKF[a]
    if a != b:
        wsP.merge_cells(f'{a}5:{b}5')
    for col in range(ord(a), ord(b) + 1):
        cell = wsP[f'{chr(col)}5']
        cell.fill, cell.border = INPUT, INBOX
    wsP[f'{a}5'].font = Font(name=F, size=11, bold=True, color=INK)
    wsP[f'{a}5'].alignment = Alignment(horizontal='left', vertical='center', indent=1)

wsP.row_dimensions[5].height = 24
wsP['A6'] = ('Filtros definidos no Dashboard (fonte única para todas as análises). '
             'Para filtrar só esta tabela, use as setas do cabeçalho.')
wsP['A6'].hyperlink = "#'Dashboard'!A5"
wsP['A6'].font = Font(name=F, size=9, color=INK3)
wsP['A7'] = ('="Mostrando: "&IF(A5="Todos","todos os SKUs",A5)&IF(D5="Todos",""," · lote "&D5)'
             '&" · "&DAY(F5)&"/"&MONTH(F5)&"/"&YEAR(F5)&" a "&DAY(H5)&"/"&MONTH(H5)&"/"&YEAR(H5)&" · unidade "&$R$3'
             '&IF(K5="Todas",""," · movimentação "&K5)')
wsP['A7'].font = Font(name=F, size=9, bold=True, color=INK2)

# --- auxiliares ocultos (R:AT) ---
master = lots_ord  # (SKU, lote) na mesma ordem da aba Cruzamento
M0, M1 = 40, 40 + len(master) - 1
LC = lambda col: f"'Lançamentos'!${col}$2:${col}${lastV}"  # noqa: E731
wsP['R1'] = '=IF(A5="Todos","*",A5)'
wsP['R2'] = '=IF(OR(D5="Todos",D5=""),"*",D5)'
wsP['R3'] = f'=IF(A5="Todos",J5,IFERROR(INDEX($U${M0}:$U${M1},MATCH(A5,$T${M0}:$T${M1},0)),J5))'
wsP['R4'] = ('=IF(K5="Entrada","Entrada",IF(K5="Saída","Saída",'
             'IF(K5="Recebido no 2008","Recebido no 2008",IF(K5="Estorno","Estorno","*"))))')
wsP['R5'] = '=Resumo!$C$7'
wsP['R6'] = '=IF(Dashboard!$N$5="Todos","*",Dashboard!$N$5)'
CRE = f'{EB("F")},">="&$F$5,{EB("F")},"<="&$H$5'
CRS = f'{SB("G")},">="&$F$5,{SB("G")},"<="&$H$5'
for r, (m, lt) in enumerate(master, M0):
    wsP[f'T{r}'] = m
    wsP[f'AG{r}'] = lt
    wsP[f'U{r}'] = mat_um2[m]
    wsP[f'AE{r}'] = f"='Cruzamento'!K{HR + 1 + r - M0}"
    wsP[f'BD{r}'] = f"='Cruzamento'!L{HR + 1 + r - M0}"
    ce = f'{CRE},{EB("B")},$T{r},{EB("C")},$AG{r}'
    cs = f'{CRS},{SB("B")},$T{r},{SB("C")},$AG{r}'
    wsP[f'V{r}'] = f'=SUMIFS({EB("G")},{ce},{EB(cK)},"Recebido no 2008")'
    wsP[f'W{r}'] = f'=SUMIFS({EB(cL)},{ce})'
    wsP[f'X{r}'] = f'=SUMIFS({SB(cSq)},{cs})'
    wsP[f'Y{r}'] = f'=SUMIFS({T_Q},{T_MAT},$T{r},{T_LOT},$AG{r})'
    wsP[f'Z{r}'] = f'=ROUND(Y{r}-(W{r}-X{r}),3)'
    wsP[f'AA{r}'] = (f'=IF(AND(U{r}=$R$3,OR($R$1="*",T{r}=$R$1),OR($R$2="*",AG{r}=$R$2),OR($R$6="*",AC{r}=$R$6),'
                     f'ABS(V{r})+ABS(W{r})+ABS(X{r})+ABS(Y{r})>0),1,0)')
    wsP[f'BE{r}'] = f"='Cruzamento'!Q{HR + 1 + r - M0}"
    wsP[f'BF{r}'] = f"='Cruzamento'!R{HR + 1 + r - M0}"
    wsP[f'AB{r}'] = f'=IF(AA{r}=1,SUM(AA${M0}:AA{r}),"")'
    wsP[f'AC{r}'] = f'=IF(ABS(Z{r})<=$R$5,"Confere",IF(Z{r}>0,"Sobrando","Faltando"))'
    wsP[f'AD{r}'] = (f'=IF(OR($R$4="*",$R$4="Saída"),COUNTIFS({cs}),0)'
                     f'+IF($R$4<>"Saída",COUNTIFS({ce},{EB(cK)},$R$4),0)')
    wsP[f'AF{r}'] = f'=IF(AND(AA{r}=1,AC{r}<>"Confere"),ABS(Z{r})+ROW()/10000000,0)'
    wsP[f'BB{r}'] = (f'=MIN(IF(_xlfn.MINIFS({EB("F")},{ce})=0,99999,_xlfn.MINIFS({EB("F")},{ce})),'
                     f'IF(_xlfn.MINIFS({SB("G")},{cs})=0,99999,_xlfn.MINIFS({SB("G")},{cs})))')
    wsP[f'BC{r}'] = f'=MAX(_xlfn.MAXIFS({EB("F")},{ce}),_xlfn.MAXIFS({SB("G")},{cs}))'
MR = lambda col: f'${col}${M0}:${col}${M1}'  # noqa: E731
# top 12 diferenças
for k in range(1, 13):
    r = M0 + k - 1
    wsP[f'AH{r}'] = f'=LARGE({MR("AF")},{k})'
    wsP[f'AI{r}'] = f'=IF(AH{r}>0,MATCH(AH{r},{MR("AF")},0),"")'
    wsP[f'AJ{r}'] = f'=IF(AI{r}="","",INDEX({MR("T")},AI{r})&" · "&INDEX({MR("AG")},AI{r}))'
    wsP[f'AK{r}'] = f'=IF(AI{r}="",0,MAX(INDEX({MR("Z")},AI{r}),0))'
    wsP[f'AL{r}'] = f'=IF(AI{r}="",0,MIN(INDEX({MR("Z")},AI{r}),0))'
# movimento diário
days = pd.date_range('2026-09-09', '2026-10-01')
D0, D1 = M0, M0 + len(days) - 1
for r, d in enumerate(days, D0):
    wsP[f'AN{r}'] = d.strftime('%d/%m')
    wsP[f'AO{r}'] = d.to_pydatetime()
    cde = f"{EB('F')},$AO{r},{EB('B')},$R$1,{EB('C')},$R$2,{EB('H')},$R$3,{EB(cK)},$R$4"
    cds = f"{SB('G')},$AO{r},{SB('B')},$R$1,{SB('C')},$R$2,{SB('I')},$R$3"
    wsP[f'AP{r}'] = f'=IF(AND(AO{r}>=$F$5,AO{r}<=$H$5,$R$4<>"Saída"),SUMIFS({EB(cL)},{cde}),0)'
    wsP[f'AQ{r}'] = f'=IF(AND(AO{r}>=$F$5,AO{r}<=$H$5,OR($R$4="*",$R$4="Saída")),-SUMIFS({SB(cSq)},{cds}),0)'
    wsP[f'AR{r}'] = f'=SUM($AP${D0}:AP{r})+SUM($AQ${D0}:AQ{r})'
wsP['AP39'], wsP['AQ39'], wsP['AJ39'], wsP['AK39'], wsP['AL39'] = 'Entrou no ZBR', 'Saiu', 'SKU · lote', 'Sobrando', 'Faltando'
# listas dos filtros (nomes definidos, para funcionar em qualquer Excel)
from openpyxl.workbook.defined_name import DefinedName
prods = sorted(mat_um2)
for i, v in enumerate(['Todos'] + prods, 1):
    wsP[f'AS{i}'] = v
# pares produto x lote e lista de lotes que depende do produto escolhido
pairs = sorted(keys)
for i, (m, l) in enumerate(pairs, 1):
    wsP[f'AW{i}'] = m
    wsP[f'AX{i}'] = l
    wsP[f'AY{i}'] = '=IF(OR($A$5="Todos",AW1=$A$5),1,0)'
NP = len(pairs)
for i in range(2, NP + 1):
    wsP[f'AY{i}'] = f'=IF(OR($A$5="Todos",AW{i}=$A$5),MAX(AY$1:AY{i - 1})+1,0)'
wsP['AV1'] = 'Todos'
for k in range(1, NP + 1):
    wsP[f'AV{k + 1}'] = f'=IFERROR(INDEX($AX$1:$AX${NP},MATCH({k},$AY$1:$AY${NP},0)),"")'
# datas do período
for i, d in enumerate(days, 1):
    wsP[f'AU{i}'] = d.to_pydatetime()
    wsP[f'AU{i}'].number_format = 'DD/MM/YYYY'
wsP['AZ1'], wsP['AZ2'], wsP['AZ3'] = 'KG', 'UN', 'CX'
for i, v in enumerate(['Todas', 'Entrada', 'Saída', 'Recebido no 2008', 'Estorno'], 1):
    wsP[f'BA{i}'] = v
NAMES = {
    'ListaProdutos': f"Painel!$AS$1:$AS${len(prods) + 1}",
    'ListaLotes': f'OFFSET(Painel!$AV$1,0,0,1+MAX(Painel!$AY$1:$AY${NP}),1)',
    'ListaDatas': f"Painel!$AU$1:$AU${len(days)}",
    'ListaUnidades': "Painel!$AZ$1:$AZ$3",
    'ListaMovimentacao': "Painel!$BA$1:$BA$5",
}
for n, ref in NAMES.items():
    wb.defined_names[n] = DefinedName(n, attr_text=ref)
DV_DASH = []
for rng, name, msg in (('A5:C5', 'ListaProdutos', 'Clique na seta ▼ e escolha o SKU (ou Todos).'),
                       ('D5:E5', 'ListaLotes', 'Mostra só os lotes do SKU escolhido. Escolha Todos para ver todos.'),
                       ('F5:G5', 'ListaDatas', 'Escolha a data inicial.'),
                       ('H5:I5', 'ListaDatas', 'Escolha a data final.'),
                       ('J5', 'ListaUnidades', 'KG, UN ou CX. Com um SKU escolhido, vale a unidade dele.'),
                       ('K5:M5', 'ListaMovimentacao', 'Filtra o gráfico por dia e a coluna Lançamentos.')):
    DV_DASH.append((rng, name, msg))  # aplicadas no Dashboard
for a in ('F5', 'H5'):
    wsP[a].number_format = 'DD/MM/YYYY'
for a in ('A5', 'D5', 'F5', 'H5', 'J5', 'K5'):
    wsP[a].fill = PatternFill('solid', fgColor=ZEBRA)
for col in ['R', 'S'] + [get_column_letter(i) for i in range(20, 60)]:
    wsP.column_dimensions[col].hidden = True


# --- quadros (linhas 9-11) ---
TILES = [('A', 'B', '="ENTROU · "&$R$3', f'=SUMPRODUCT({MR("AA")},{MR("W")})', 'transferido do 2008 para o ZBR', C_ENT),
         ('C', 'E', '="SAIU · "&$R$3', f'=SUMPRODUCT({MR("AA")},{MR("X")})', 'vendas e entregas', C_SAI),
         ('F', 'G', '="DEVERIA TER · "&$R$3', '=A10-C10', 'entrou − saiu', None),
         ('H', 'I', '="ESTOQUE FÍSICO · "&$R$3', f'=SUMPRODUCT({MR("AA")},{MR("Y")})', 'o que tem no depósito hoje', None),
         ('J', 'K', '="DIFERENÇA · "&$R$3', '=ROUND(H10-F10,3)',
          '=IF(OR(F5<>DATE(2026,9,9),H5<>DATE(2026,10,1)),"período parcial: compare com cuidado",'
          'IF(ABS(J10)<=$R$5,"tudo confere",IF(J10>0,"sobrando no estoque","faltando no estoque")))', None)]
for a, b, lab, val, hint, mk in TILES:
    card(9, ord(a) - 64, 11, ord(b) - 64)
    for rr in (9, 10, 11):
        wsP.merge_cells(f'{a}{rr}:{b}{rr}')
    wsP[f'{a}9'] = lab
    wsP[f'{a}9'].font = Font(name=F, size=8, bold=True, color=mk or INK3)
    wsP[f'{a}10'] = val
    wsP[f'{a}10'].font = Font(name=FN, size=24, bold=True, color=INK)
    wsP[f'{a}10'].number_format = '+#,##0.0;-#,##0.0;0' if a == 'J' else NUM
    wsP[f'{a}11'] = hint
    wsP[f'{a}11'].font = Font(name=F, size=9, color=INK3)
    for rr in (9, 10, 11):
        wsP[f'{a}{rr}'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
card(9, 10, 11, 11, edge=ACCENT, width='medium')
wsP.conditional_formatting.add('J10', CellIsRule(operator='lessThan', formula=['-$R$5'], font=Font(color=R_TX)))
wsP.conditional_formatting.add('J10', CellIsRule(operator='greaterThan', formula=['$R$5'], font=Font(color=W_TX)))
wsP.conditional_formatting.add('J10', FormulaRule(formula=['ABS($J$10)<=$R$5'], font=Font(color=G_TX)))
for rr, (lab, fill, tx) in zip((9, 10, 11), (('Confere', GRN_F, G_TX), ('Sobrando', YEL_F, W_TX), ('Faltando', RED_F, R_TX))):
    wsP.merge_cells(f'L{rr}:M{rr}')
    wsP[f'L{rr}'] = f'="{lab}: "&COUNTIFS({MR("AA")},1,{MR("AC")},"{lab}")&" SKU·lote"'
    for col in 'LM':
        wsP[f'{col}{rr}'].fill = fill
    wsP[f'L{rr}'].font = Font(name=F, size=10, bold=True, color=tx)
    wsP[f'L{rr}'].alignment = Alignment(indent=1, vertical='center')
wsP['N9'] = 'Indicadores e gráficos: aba Dashboard →'
wsP['N9'].font = Font(name=F, size=9, bold=True, color=ACCENT)
wsP['N9'].hyperlink = "#'Dashboard'!A1"
for rr in (9, 10, 11):
    wsP.row_dimensions[rr].height = 22 if rr != 10 else 34

# --- tabela por SKU e lote: cabeçalho congelado e com filtro ---
T0 = 13
PH = ['SKU', 'Lote', 'UM', 'Recebido 2008', 'Entrou', 'Saiu', 'Deveria ter', 'Estoque físico', 'Diferença',
      'Situação', 'Lançamentos', '1º lançamento', 'Último lançamento', 'Motivo da diferença', 'Solução de ajuste (sugestão)']
for i, h in enumerate(PH, 1):
    cell = wsP.cell(row=T0, column=i, value=h.upper())
    cell.font = Font(name=F, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=ACCENT)
    cell.border = Border(bottom=Side(style='medium', color=GOLD))
    cell.alignment = Alignment(horizontal='right' if 4 <= i <= 9 else 'center' if i in (10, 11, 12, 13) else 'left',
                               vertical='center', wrap_text=True, indent=1 if i in (14, 15) else 0)
wsP.row_dimensions[T0].height = 30
for k in range(1, len(master) + 1):
    r = T0 + k
    idx = f'IFERROR(MATCH({k},{MR("AB")},0),0)'
    g = lambda col: f'=IF({idx}=0,"",INDEX({MR(col)},{idx}))'  # noqa: E731
    for col, src in zip('ABCDEF', ('T', 'AG', 'U', 'V', 'W', 'X')):
        wsP[f'{col}{r}'] = g(src)
    wsP[f'G{r}'] = f'=IF(A{r}="","",E{r}-F{r})'
    wsP[f'H{r}'], wsP[f'I{r}'], wsP[f'J{r}'], wsP[f'K{r}'] = g('Y'), g('Z'), g('AC'), g('AD')
    wsP[f'L{r}'] = f'=IF({idx}=0,"",IF(INDEX({MR("BB")},{idx})>=99999,"",INDEX({MR("BB")},{idx})))'
    wsP[f'M{r}'] = f'=IF({idx}=0,"",IF(INDEX({MR("BC")},{idx})=0,"",INDEX({MR("BC")},{idx})))'
    wsP[f'N{r}'] = f'=IF({idx}=0,"",INDEX({MR("AE")},{idx})&"")'
    wsP[f'O{r}'] = f'=IF({idx}=0,"",INDEX({MR("BD")},{idx})&"")'
    for c in range(1, 16):
        cell = wsP.cell(row=r, column=c)
        cell.fill = WHITE
        cell.border = Border(bottom=Side(style='thin', color=GRID))
        cell.font = Font(name=F, size=10, color=INK, bold=(c in (1, 2, 9)))
        if 4 <= c <= 9:
            cell.number_format = '+#,##0.0;-#,##0.0;0' if c == 9 else NUM
        if c in (10, 11, 12, 13):
            cell.alignment = Alignment(horizontal='center', vertical='center')
        if c in (12, 13):
            cell.number_format = 'DD/MM/YYYY'
    for c in (14, 15):
        wsP.cell(row=r, column=c).font = Font(name=F, size=9, color=INK if c == 14 else INK2, bold=(c == 14))
        wsP.cell(row=r, column=c).alignment = Alignment(indent=1, vertical='center', wrap_text=False)
TL = T0 + len(master)
for lab, fill, tx in (('Confere', GRN_F, G_TX), ('Sobrando', YEL_F, W_TX), ('Faltando', RED_F, R_TX)):
    wsP.conditional_formatting.add(f'J{T0 + 1}:J{TL}', FormulaRule(formula=[f'J{T0 + 1}="{lab}"'], fill=fill, font=Font(color=tx, bold=True)))
wsP.conditional_formatting.add(f'I{T0 + 1}:I{TL}', FormulaRule(formula=[f'AND(ISNUMBER(I{T0 + 1}),I{T0 + 1}<-$R$5)'], font=Font(color=R_TX, bold=True)))
wsP.conditional_formatting.add(f'A{T0 + 1}:O{TL}', FormulaRule(formula=[f'$A{T0 + 1}=""'], fill=PatternFill('solid', fgColor=PAGE), border=Border()))
wsP.auto_filter.ref = f'A{T0}:O{TL}'
wsP[f'A{TL + 2}'] = ('Entrou = transferido do depósito 2008 para o ZBR (estornos já descontados). Saiu = vendas e entregas. '
                     'Recebido no 2008 = importação que chegou e ainda não foi transferida.')
wsP[f'A{TL + 3}'] = ('Motivo e solução: calculados para o período completo (09/09 a 01/10). As soluções são sugestões; '
                     'confirme no SAP antes de lançar qualquer ajuste.')
for rr in (TL + 2, TL + 3):
    wsP[f'A{rr}'].font = Font(name=F, size=9, color=INK3)
wsP.freeze_panes = f'C{T0 + 1}'
wsP.page_setup.orientation = 'landscape'
wsP.page_setup.fitToWidth = 1
wsP.page_setup.fitToHeight = 0
wsP.sheet_properties.pageSetUpPr.fitToPage = True
wsP.print_area = f'A1:O{TL + 3}'
wsP.print_title_rows = f'{T0}:{T0}'

# ================= utilitários de gráfico =================
from openpyxl.chart import LineChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont


def chart_title_style(ch):
    if ch.title:
        txt = ch.title
        ch.title = txt
        try:
            ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=CharacterProperties(sz=1200, b=True, solidFill=NAVY, latin=DFont(typeface=F)))
        except Exception:
            pass


def style_chart(ch, legend=False):
    ch.visible_cells_only = False
    if not legend:
        ch.legend = None
    else:
        ch.legend.position = 'b'
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill='E3E8EF'))
    ch.y_axis.number_format = '#,##0'
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    chart_title_style(ch)


def easy_bar(title, cats, vals, colors, horizontal=False, fmt='#,##0'):
    ch = BarChart()
    ch.type = 'bar' if horizontal else 'col'
    ch.title = title
    ch.add_data(vals, titles_from_data=False)
    ch.set_categories(cats)
    ch.gapWidth = 60
    sr = ch.series[0]
    sr.graphicalProperties.line.noFill = True
    for i, col in enumerate(colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = col
        pt.graphicalProperties.line.noFill = True
        sr.dPt.append(pt)
    sr.dLbls = DataLabelList()
    sr.dLbls.showVal = True
    sr.dLbls.numFmt = fmt
    for attr in ('showSerName', 'showCatName', 'showLegendKey', 'showPercent'):
        setattr(sr.dLbls, attr, False)
    style_chart(ch)
    if horizontal:
        ch.x_axis.scaling.orientation = 'maxMin'
    return ch


# ================= DEPÓSITO 2008 (consulta + dash) =================
w8 = wb.create_sheet('Depósito 2008')
w8.sheet_view.showGridLines = False
W8 = {'A': 20, 'B': 13, 'C': 7, 'D': 13, 'E': 13, 'F': 12, 'G': 13, 'H': 12, 'I': 22, 'J': 11, 'K': 12, 'L': 12, 'M': 12, 'N': 34, 'O': 46}
for c, w in W8.items():
    w8.column_dimensions[c].width = w
for r in range(1, 112):
    for c in range(1, 16):
        w8.cell(row=r, column=c).fill = PatternFill('solid', fgColor=PAGE)

def card8(r1, c1, r2, c2, edge=LINE, width='thin'):
    sd = Side(style=width, color=edge)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = w8.cell(row=r, column=c)
            cell.fill = PatternFill('solid', fgColor=SURF)
            cell.border = Border(left=sd if c == c1 else None, right=sd if c == c2 else None,
                                 top=sd if r == r1 else None, bottom=sd if r == r2 else None)

w8['A1'] = 'Depósito 2008 — importação'
w8['A1'].font = TITLE
w8['A2'] = ('Tudo o que entrou e saiu do depósito 2008: recebimento da importação, transferências para o ZBR e estornos · '
            'movimentos de 09/09 a 30/09/2026')
w8['A2'].font = SUB
w8.row_dimensions[1].height = 30
card8(4, 1, 7, 15)
FILT8 = [('A', 'C', 'SKU', 'Todos'), ('D', 'E', 'LOTE', 'Todos'),
         ('F', 'G', 'DATA INICIAL', pd.Timestamp('2026-09-09').to_pydatetime()),
         ('H', 'I', 'DATA FINAL', pd.Timestamp('2026-10-01').to_pydatetime()),
         ('J', 'J', 'UNIDADE', 'KG'), ('K', 'M', 'SITUAÇÃO', 'Todas')]
LINK8 = {'A': '=Dashboard!E5', 'D': '=Dashboard!H5', 'F': '=Dashboard!A5', 'H': '=Dashboard!C5', 'J': '=Dashboard!J5'}
for a, b, lab, v in FILT8:
    w8[f'{a}4'] = lab
    w8[f'{a}4'].font = Font(name=F, size=8, bold=True, color=INK3)
    w8[f'{a}5'] = LINK8.get(a, v)
    if a != b:
        w8.merge_cells(f'{a}5:{b}5')
    for col in range(ord(a), ord(b) + 1):
        w8[f'{chr(col)}5'].fill, w8[f'{chr(col)}5'].border = INPUT, INBOX
    w8[f'{a}5'].font = Font(name=F, size=11, bold=True, color=INK)
    w8[f'{a}5'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
w8['K5'].number_format = '@"   ▼"'
for a in ('F5', 'H5'):
    w8[a].number_format = 'DD/MM/YYYY'
for a in ('A5', 'D5', 'F5', 'H5', 'J5'):
    w8[a].fill = PatternFill('solid', fgColor=ZEBRA)
w8.row_dimensions[5].height = 24
w8['A6'] = 'SKU, lote, datas e unidade vêm do Dashboard (fonte única). A situação ▼ é filtrada aqui mesmo.'
w8['A6'].hyperlink = "#'Dashboard'!A5"
w8['A6'].font = Font(name=F, size=9, color=INK3)
w8['A7'] = ('="Mostrando: "&IF(A5="Todos","todos os SKUs",A5)&IF(D5="Todos",""," · lote "&D5)'
            '&" · "&DAY(F5)&"/"&MONTH(F5)&"/"&YEAR(F5)&" a "&DAY(H5)&"/"&MONTH(H5)&"/"&YEAR(H5)&" · unidade "&$R$3'
            '&IF(K5="Todas",""," · "&K5)')
w8['A7'].font = Font(name=F, size=9, bold=True, color=INK2)

# --- auxiliares ocultos ---
EN = lambda col: f"Entrada!${col}$2:${col}${MAXR}"  # noqa: E731
p8 = {}
for m, l, q in ent[['Material', 'Lote', 'Quantidade']].itertuples(index=False):
    p8[(m, l)] = p8.get((m, l), 0) + q
pairs8 = sorted(p8, key=lambda k: (-round(p8[k], 3), k))
um8 = {k: um[k] for k in pairs8}
Q0, Q1 = 40, 40 + len(pairs8) - 1
w8['R1'] = '=IF(A5="Todos","*",A5)'
w8['R2'] = '=IF(OR(D5="Todos",D5=""),"*",D5)'
w8['R3'] = f'=IF(A5="Todos",J5,IFERROR(INDEX($U${Q0}:$U${Q1},MATCH(A5,$T${Q0}:$T${Q1},0)),J5))'
w8['R5'] = '=Resumo!$C$7'
CR8 = f'{EN("F")},">="&$F$5,{EN("F")},"<="&$H$5'
SIT8 = ('Aguardando transferência', 'Transferido 100%', 'Saiu saldo anterior a 09/09')
for r, (m, lt) in enumerate(pairs8, Q0):
    w8[f'T{r}'], w8[f'AG{r}'], w8[f'U{r}'] = m, lt, um8[(m, lt)]
    w8[f'AF{r}'] = desc.get(m, '')
    c = f'{CR8},{EN("B")},$T{r},{EN("C")},$AG{r}'
    w8[f'V{r}'] = f'=SUMIFS({EN("G")},{c},{EN(cK)},"Recebido no 2008")'
    w8[f'W{r}'] = f'=-SUMIFS({EN("G")},{c},{EN(cK)},"Entrada")'
    w8[f'X{r}'] = f'=SUMIFS({EN("G")},{c},{EN(cK)},"Estorno")'
    w8[f'Y{r}'] = f'=ROUND(V{r}-W{r}+X{r},3)'
    w8[f'AB{r}'] = (f'=IF(ABS(Y{r})<=$R$5,"{SIT8[1]}",IF(Y{r}>0,"{SIT8[0]}","{SIT8[2]}"))')
    w8[f'Z{r}'] = (f'=IF(AND(U{r}=$R$3,OR($R$1="*",T{r}=$R$1),OR($R$2="*",AG{r}=$R$2),'
                   f'OR($K$5="Todas",AB{r}=$K$5),ABS(V{r})+ABS(W{r})+ABS(X{r})>0),1,0)')
    w8[f'AA{r}'] = f'=IF(Z{r}=1,SUM(Z${Q0}:Z{r}),"")'
    w8[f'AC{r}'] = f'=COUNTIFS({c})'
    w8[f'AD{r}'] = f'=_xlfn.MINIFS({EN("F")},{c})'
    w8[f'AE{r}'] = f'=_xlfn.MAXIFS({EN("F")},{c})'
    w8[f'AH{r}'] = f'=IF(AND(Z{r}=1,Y{r}>$R$5),Y{r}+ROW()/10000000,0)'
Q = lambda col: f'${col}${Q0}:${col}${Q1}'  # noqa: E731
# top 10 pendentes no 2008
for k in range(1, 11):
    r = Q0 + k - 1
    w8[f'AJ{r}'] = f'=LARGE({Q("AH")},{k})'
    w8[f'AK{r}'] = f'=IF(AJ{r}>0,MATCH(AJ{r},{Q("AH")},0),"")'
    w8[f'AL{r}'] = f'=IF(AK{r}="","",INDEX({Q("T")},AK{r})&" · "&INDEX({Q("AG")},AK{r}))'
    w8[f'AM{r}'] = f'=IF(AK{r}="",0,INDEX({Q("Y")},AK{r}))'
w8[f'AL{Q0 - 1}'], w8[f'AM{Q0 - 1}'] = 'SKU · lote', 'Aguardando no 2008'
# movimento diário do 2008
for r, d in enumerate(days, Q0):
    w8[f'AO{r}'] = d.strftime('%d/%m')
    w8[f'AP{r}'] = d.to_pydatetime()
    cd = f"{EN('F')},$AP{r},{EN('B')},$R$1,{EN('C')},$R$2,{EN('H')},$R$3"
    w8[f'AQ{r}'] = f'=IF(AND(AP{r}>=$F$5,AP{r}<=$H$5),SUMIFS({EN("G")},{cd},{EN(cK)},"Recebido no 2008"),0)'
    w8[f'AR{r}'] = (f'=IF(AND(AP{r}>=$F$5,AP{r}<=$H$5),SUMIFS({EN("G")},{cd},{EN(cK)},"Entrada")'
                    f'+SUMIFS({EN("G")},{cd},{EN(cK)},"Estorno"),0)')
w8[f'AQ{Q0 - 1}'], w8[f'AR{Q0 - 1}'] = 'Recebido no 2008', 'Transferido ao ZBR (líquido)'
# listas: SKUs do 2008, lotes dependentes, situação
skus8 = sorted({m for m, _ in pairs8})
for i, v in enumerate(['Todos'] + skus8, 1):
    w8[f'AT{i}'] = v
srt8 = sorted(pairs8)
for i, (m, l) in enumerate(srt8, 1):
    w8[f'AW{i}'], w8[f'AX{i}'] = m, l
    w8[f'AY{i}'] = ('=IF(OR($A$5="Todos",AW1=$A$5),1,0)' if i == 1 else
                    f'=IF(OR($A$5="Todos",AW{i}=$A$5),MAX(AY$1:AY{i - 1})+1,0)')
w8['AV1'] = 'Todos'
for k in range(1, len(srt8) + 1):
    w8[f'AV{k + 1}'] = f'=IFERROR(INDEX($AX$1:$AX${len(srt8)},MATCH({k},$AY$1:$AY${len(srt8)},0)),"")'
for i, v in enumerate(('Todas',) + SIT8, 1):
    w8[f'AZ{i}'] = v
for n, ref in {'ListaSKU2008': f"'Depósito 2008'!$AT$1:$AT${len(skus8) + 1}",
               'ListaLotes2008': f"OFFSET('Depósito 2008'!$AV$1,0,0,1+MAX('Depósito 2008'!$AY$1:$AY${len(srt8)}),1)",
               'ListaSituacao2008': "'Depósito 2008'!$AZ$1:$AZ$4"}.items():
    wb.defined_names[n] = DefinedName(n, attr_text=ref)
for rng, name, msg in (('K5:M5', 'ListaSituacao2008', 'Filtra pela situação do SKU·lote no 2008.'),):
    dv = DataValidation(type='list', formula1=f'={name}', allow_blank=False, showErrorMessage=True,
                        errorTitle='Valor inválido', error='Escolha um item da lista.',
                        showInputMessage=True, promptTitle='Filtro', prompt=msg)
    w8.add_data_validation(dv)
    dv.add(rng)
for col in ['R', 'S'] + [get_column_letter(i) for i in range(20, 53)]:
    w8.column_dimensions[col].hidden = True

# --- quadros ---
T8 = [('A', 'B', '="RECEBIDO NO 2008 · "&$R$3', f'=SUMPRODUCT({Q("Z")},{Q("V")})', 'chegada da importação', C_REC),
      ('C', 'E', '="TRANSFERIDO AO ZBR · "&$R$3', f'=SUMPRODUCT({Q("Z")},{Q("W")})', 'saiu do 2008 para venda', C_ENT),
      ('F', 'G', '="ESTORNOS · "&$R$3', f'=SUMPRODUCT({Q("Z")},{Q("X")})', 'voltou para o 2008', None),
      ('H', 'I', '="SALDO NO 2008 · "&$R$3', '=ROUND(A10-C10+F10,3)', 'recebido − transferido + estornos', None),
      ('J', 'K', 'LANÇAMENTOS', f'=SUMPRODUCT({Q("Z")},{Q("AC")})', 'movimentos no 2008', None)]
for a, b, lab, val, hint, mk in T8:
    card8(9, ord(a) - 64, 11, ord(b) - 64)
    for rr in (9, 10, 11):
        w8.merge_cells(f'{a}{rr}:{b}{rr}')
    w8[f'{a}9'], w8[f'{a}10'], w8[f'{a}11'] = lab, val, hint
    w8[f'{a}9'].font = Font(name=F, size=8, bold=True, color=mk or INK3)
    w8[f'{a}10'].font = Font(name=FN, size=24, bold=True, color=INK)
    w8[f'{a}10'].number_format = '0' if a == 'J' else ('+#,##0.0;-#,##0.0;0' if a == 'H' else NUM)
    w8[f'{a}11'].font = Font(name=F, size=9, color=INK3)
    for rr in (9, 10, 11):
        w8[f'{a}{rr}'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
card8(9, 8, 11, 9, edge=ACCENT, width='medium')
for rr, (lab, fill, tx) in zip((9, 10, 11), ((SIT8[0], YEL_F, W_TX), (SIT8[1], GRN_F, G_TX), (SIT8[2], PatternFill('solid', fgColor=SURF2), INK2))):
    w8.merge_cells(f'L{rr}:M{rr}')
    w8[f'L{rr}'] = f'="{lab}: "&COUNTIFS({Q("Z")},1,{Q("AB")},"{lab}")'
    for col in 'LM':
        w8[f'{col}{rr}'].fill = fill
    w8[f'L{rr}'].font = Font(name=F, size=9, bold=True, color=tx)
    w8[f'L{rr}'].alignment = Alignment(indent=1, vertical='center')
w8['N9'] = 'Gráficos do 2008: abaixo da tabela ↓'
w8['N9'].font = Font(name=F, size=9, bold=True, color=ACCENT)
w8['N9'].hyperlink = "#'Dash 2008'!A1"
for rr in (9, 10, 11):
    w8.row_dimensions[rr].height = 22 if rr != 10 else 34

# --- tabela por SKU e lote (cabeçalho congelado e com filtro) ---
H8 = 13
PH8 = ['SKU', 'Lote', 'UM', 'Recebido no 2008', 'Transferido ao ZBR', 'Estornos', 'Saldo no 2008', '% transferido',
       'Situação', 'Lançamentos', '1º movimento', 'Último movimento', 'Descrição', 'Próximo passo']
for i, h in enumerate(PH8, 1):
    cell = w8.cell(row=H8, column=i, value=h.upper())
    cell.font = Font(name=F, size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor=ACCENT)
    cell.border = Border(bottom=Side(style='medium', color=GOLD))
    cell.alignment = Alignment(horizontal='right' if 4 <= i <= 8 else 'center' if i in (9, 10, 11, 12) else 'left',
                               vertical='center', wrap_text=True, indent=1 if i in (13, 14) else 0)
w8.row_dimensions[H8].height = 30
for k in range(1, len(pairs8) + 1):
    r = H8 + k
    idx = f'IFERROR(MATCH({k},{Q("AA")},0),0)'
    g8 = lambda col: f'=IF({idx}=0,"",INDEX({Q(col)},{idx}))'  # noqa: E731
    for col, src in zip('ABCDEFG', ('T', 'AG', 'U', 'V', 'W', 'X', 'Y')):
        w8[f'{col}{r}'] = g8(src)
    w8[f'H{r}'] = f'=IF(A{r}="","",IF(D{r}>0,(E{r}-F{r})/D{r},""))'
    w8[f'I{r}'], w8[f'J{r}'] = g8('AB'), g8('AC')
    w8[f'K{r}'] = f'=IF({idx}=0,"",IF(INDEX({Q("AD")},{idx})=0,"",INDEX({Q("AD")},{idx})))'
    w8[f'L{r}'] = f'=IF({idx}=0,"",IF(INDEX({Q("AE")},{idx})=0,"",INDEX({Q("AE")},{idx})))'
    w8[f'M{r}'] = f'=IF({idx}=0,"",INDEX({Q("AF")},{idx})&"")'
    w8[f'N{r}'] = (f'=IF(A{r}="","",IF(I{r}="{SIT8[0]}","Transferir "&FIXED(G{r},1)&" "&C{r}&" do 2008 para o ZBR (verificar se já há OT aberta)",'
                   f'IF(I{r}="{SIT8[1]}","Nada a fazer","Sem ação: saíram "&FIXED(-G{r},1)&" "&C{r}&" que já estavam no 2008 antes de 09/09 (conferir só se não havia saldo)")))')
    for c in range(1, 15):
        cell = w8.cell(row=r, column=c)
        cell.fill = WHITE
        cell.border = Border(bottom=Side(style='thin', color=GRID))
        cell.font = Font(name=F, size=10, color=INK, bold=(c in (1, 2, 7)))
        if 4 <= c <= 7:
            cell.number_format = '+#,##0.0;-#,##0.0;0' if c == 7 else NUM
        if c == 8:
            cell.number_format = '0%'
        if c in (9, 10, 11, 12):
            cell.alignment = Alignment(horizontal='center', vertical='center')
        if c in (11, 12):
            cell.number_format = 'DD/MM/YYYY'
    for c in (13, 14):
        w8.cell(row=r, column=c).font = Font(name=F, size=9, color=INK2 if c == 13 else INK, bold=(c == 14))
        w8.cell(row=r, column=c).alignment = Alignment(indent=1, vertical='center')
L8 = H8 + len(pairs8)
for lab, fill, tx in ((SIT8[0], YEL_F, W_TX), (SIT8[1], GRN_F, G_TX), (SIT8[2], PatternFill('solid', fgColor=SURF2), INK2)):
    w8.conditional_formatting.add(f'I{H8 + 1}:I{L8}', FormulaRule(formula=[f'I{H8 + 1}="{lab}"'], fill=fill, font=Font(color=tx, bold=True)))
w8.conditional_formatting.add(f'A{H8 + 1}:N{L8}', FormulaRule(formula=[f'$A{H8 + 1}=""'], fill=PatternFill('solid', fgColor=PAGE), border=Border()))
w8.auto_filter.ref = f'A{H8}:N{L8}'
w8.freeze_panes = f'C{H8 + 1}'
w8[f'A{L8 + 2}'] = ('Saldo no 2008 é calculado pelo movimento do período (o arquivo não traz a foto do estoque do 2008). '
                    'Saldo negativo = transferiu mercadoria que já estava no 2008 antes de 09/09.')
w8[f'A{L8 + 3}'] = 'Cada lançamento do 2008 (documento, data e hora) está na aba "Entrada": filtre por SKU e lote.'
for rr in (L8 + 2, L8 + 3):
    w8[f'A{rr}'].font = Font(name=F, size=9, color=INK3)
w8.page_setup.orientation = 'landscape'
w8.page_setup.fitToWidth = 1
w8.page_setup.fitToHeight = 0
w8.sheet_properties.pageSetUpPr.fitToPage = True
w8.print_title_rows = f'{H8}:{H8}'

# --- Gráficos do 2008 (abaixo da tabela) ---
G8 = L8 + 6
w8['N9'].hyperlink = f"#'Depósito 2008'!A{G8}"
w8[f'A{G8 - 2}'] = 'Gráficos do depósito 2008'
w8[f'A{G8 - 2}'].font = SEC
w8[f'A{G8 - 1}'] = 'Seguem os mesmos filtros desta aba.'
w8[f'A{G8 - 1}'].font = Font(name=F, size=9, color=INK3)
w8[f'A{G8}'], w8[f'B{G8}'] = 'Indicador', "=\"Quantidade (\"&$R$3&\")\""
for i, (lab, ref) in enumerate((('Recebido no 2008', 'A10'), ('Transferido ao ZBR', 'C10'), ('Estornos', 'F10'),
                                ('Saldo no 2008', 'H10')), G8 + 1):
    w8[f'A{i}'] = lab
    w8[f'B{i}'] = f'={ref}'
    w8[f'B{i}'].number_format = NUM
w8[f'D{G8}'], w8[f'E{G8}'] = 'Situação (SKU·lote)', 'Qtd'
for i, lab in enumerate(SIT8, G8 + 1):
    w8[f'D{i}'] = lab
    w8[f'E{i}'] = f'=COUNTIFS({Q("Z")},1,{Q("AB")},"{lab}")'
for cell in (f'A{G8}', f'B{G8}', f'D{G8}', f'E{G8}'):
    w8[cell].font = Font(name=F, size=10, bold=True, color='FFFFFF')
    w8[cell].fill = PatternFill('solid', fgColor=NAVY)
for r in range(G8 + 1, G8 + 5):
    for col in 'AB':
        w8[f'{col}{r}'].font, w8[f'{col}{r}'].fill = BASE, WHITE
for r in range(G8 + 1, G8 + 4):
    for col in 'DE':
        w8[f'{col}{r}'].font, w8[f'{col}{r}'].fill = BASE, WHITE
h1 = easy_bar('Recebido, transferido e saldo no 2008', Reference(w8, min_col=1, min_row=G8 + 1, max_row=G8 + 4),
              Reference(w8, min_col=2, min_row=G8 + 1, max_row=G8 + 4), [C_REC, C_ENT, GREY_C, CORP])
h1.height, h1.width = 8.5, 16
w8.add_chart(h1, f'A{G8 + 6}')
h2 = easy_bar('Situação dos SKU·lote no 2008', Reference(w8, min_col=4, min_row=G8 + 1, max_row=G8 + 3),
              Reference(w8, min_col=5, min_row=G8 + 1, max_row=G8 + 3), [W_BAR, G_BAR, GREY_C])
h2.height, h2.width = 8.5, 14
w8.add_chart(h2, f'H{G8 + 6}')
h3 = BarChart()
h3.type, h3.gapWidth = 'bar', 50
h3.title = 'Top 10 aguardando transferência no 2008'
h3.add_data(Reference(w8, min_col=39, min_row=Q0 - 1, max_row=Q0 + 9), titles_from_data=True)  # AM
h3.set_categories(Reference(w8, min_col=38, min_row=Q0, max_row=Q0 + 9))  # AL
h3.series[0].graphicalProperties.solidFill = W_BAR
h3.series[0].graphicalProperties.line.noFill = True
h3.series[0].dLbls = DataLabelList()
h3.series[0].dLbls.showVal = True
h3.series[0].dLbls.numFmt = '#,##0'
for attr in ('showSerName', 'showCatName', 'showLegendKey', 'showPercent'):
    setattr(h3.series[0].dLbls, attr, False)
style_chart(h3)
h3.x_axis.scaling.orientation = 'maxMin'
h3.height, h3.width = 9.5, 30
w8.add_chart(h3, f'A{G8 + 24}')
h4 = BarChart()
h4.type, h4.grouping, h4.overlap, h4.gapWidth = 'col', 'clustered', 100, 40
h4.title = 'Movimento por dia no 2008 (recebido acima, transferido abaixo)'
h4.add_data(Reference(w8, min_col=43, min_row=Q0 - 1, max_row=Q0 + len(days) - 1), titles_from_data=True)  # AQ
h4.add_data(Reference(w8, min_col=44, min_row=Q0 - 1, max_row=Q0 + len(days) - 1), titles_from_data=True)  # AR
h4.set_categories(Reference(w8, min_col=41, min_row=Q0, max_row=Q0 + len(days) - 1))  # AO
h4.series[0].graphicalProperties.solidFill = C_REC
h4.series[1].graphicalProperties.solidFill = C_ENT
for sr in h4.series:
    sr.graphicalProperties.line.noFill = True
style_chart(h4, legend=True)
h4.x_axis.tickLblPos = 'low'
h4.height, h4.width = 8.5, 30
w8.add_chart(h4, f'A{G8 + 44}')
for r in range(L8 + 1, G8 + 64):
    for c in range(1, 16):
        if w8.cell(row=r, column=c).fill.fgColor.rgb in (None, '00000000'):
            w8.cell(row=r, column=c).fill = PatternFill('solid', fgColor=PAGE)

# =====================================================================================
#  PADRONIZAÇÃO EXECUTIVA: parâmetros, conciliação, plano, dashboard, resumo, dicionário
# =====================================================================================
import re
from openpyxl.worksheet.table import Table as XTable, TableStyleInfo as XStyle
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

NAVF = PatternFill('solid', fgColor=NAVY)
CORPF = PatternFill('solid', fgColor=CORP)
LIGHTF = PatternFill('solid', fgColor=LIGHT)
ZEBF = PatternFill('solid', fgColor=ZEBRA)
WHF = PatternFill('solid', fgColor=WHITE_C)
TOLR = 'Resumo!$C$7'


def hdr(ws, row, col, text, fill=NAVF, color='FFFFFF', align='left', wrap=True):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=F, size=10, bold=True, color=color)
    c.fill = fill
    c.border = Border(bottom=Side(style='medium', color=CORP))
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    return c


def cell(ws, ref, value, size=10, bold=False, color=INK, fmt=None, align=None, fill=None, italic=False):
    c = ws[ref]
    c.value = value
    c.font = Font(name=F, size=size, bold=bold, color=color, italic=italic)
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=False)
    if fill:
        c.fill = fill
    return c


def title_block(ws, title, subtitle, last_col):
    cell(ws, 'A1', title, size=22, bold=True, color=NAVY)
    cell(ws, 'A2', subtitle, size=10, color=INK3)
    ws.row_dimensions[1].height = 32
    for col in range(1, last_col + 1):
        ws.cell(row=3, column=col).border = Border(top=Side(style='medium', color=NAVY))


# ---------- parâmetros adicionais (aba Resumo → "Parâmetros") ----------
res.column_dimensions['J'].width = 54
res.column_dimensions['K'].width = 16
cell(res, 'J4', 'Parâmetros de classificação e prazo (editáveis)', size=13, bold=True, color=CORP)
PARS = [('J6', 'K6', 'Falta a partir da qual a ocorrência é Crítica (na unidade do SKU)', 200, NUM),
        ('J7', 'K7', 'Sobra sem explicação a partir da qual a prioridade é Alta', 500, NUM),
        ('J8', 'K8', 'Saldo no 2008 a partir do qual a prioridade é Alta', 2000, NUM),
        ('J9', 'K9', 'Data de geração do relatório', DATA_RELATORIO, 'DD/MM/YYYY'),
        ('J10', 'K10', 'Data de referência para prazos (hoje)', '=TODAY()', 'DD/MM/YYYY'),
        ('J11', 'K11', 'Dias para alerta "vence em breve"', 3, '0')]
for lj, lk, lab, v, fmt in PARS:
    cell(res, lj, lab)
    c = cell(res, lk, v, bold=True, color=CORP, fmt=fmt, fill=INPUT)
    c.border = INBOX
res['K10'].fill = ZEBF
res['J12'] = 'K10 usa HOJE() e é a única fórmula volátil da planilha (necessária para os alertas de prazo).'
res['J12'].font = Font(name=F, size=9, color=INK3, italic=True)
P_CRIT, P_ALTA, P_2008, P_DATA, P_HOJE, P_DIAS = ('Resumo!$K$6', 'Resumo!$K$7', 'Resumo!$K$8',
                                                  'Resumo!$K$9', 'Resumo!$K$10', 'Resumo!$K$11')

# ---------- conciliação: categoria da causa, criticidade (proposta), chave e ranking ----------
CATS = ['Falta sem explicação', 'Não localizado no estoque', 'Saída sem baixa no estoque',
        'Doca/transferência pendente', 'Estoque anterior a 09/09', 'Troca de lote',
        'Sobra parcialmente explicada', 'Sobra sem explicação']
LEVELS = ['Crítica', 'Alta', 'Média', 'Baixa']
for col, txt, w in ((17, 'CATEGORIA DA CAUSA', 28), (18, 'CRITICIDADE (PROPOSTA)', 16), (19, 'CHAVE SKU|LOTE', 22),
                    (20, 'RANKING', 8)):
    hdr(wsA, HR, col, txt)
    wsA.column_dimensions[get_column_letter(col)].width = w
for r in range(L0, L1 + 1):
    wsA[f'Q{r}'] = (f'=IF(J{r}="Confere","—",IF(LEFT(K{r},5)="Troca","Troca de lote",'
                    f'IF(LEFT(K{r},6)="Entrou","Não localizado no estoque",IF(J{r}="Faltando","Falta sem explicação",'
                    f'IF(ISNUMBER(SEARCH("não baixaram",K{r})),"Saída sem baixa no estoque",'
                    f'IF(LEFT(K{r},10)="Quantidade","Doca/transferência pendente",'
                    f'IF(LEFT(K{r},14)="Estoque antigo","Estoque anterior a 09/09",'
                    f'IF(LEFT(K{r},7)="Explica","Sobra parcialmente explicada","Sobra sem explicação"))))))))')
    wsA[f'R{r}'] = (f'=IF(J{r}="Confere","—",IF(OR(LEFT(K{r},6)="Entrou",AND(I{r}<0,ABS(O{r})>={P_CRIT})),"Crítica",'
                    f'IF(OR(AND(I{r}<0,O{r}<>0),O{r}>={P_ALTA}),"Alta",IF(O{r}>0,"Média","Baixa"))))')
    wsA[f'S{r}'] = f'=A{r}&"|"&B{r}'
    wsA[f'T{r}'] = f'=IF(ABS(O{r})>{TOLR},ABS(O{r})+ROW()/10000000,0)'
    for col in 'QRST':
        wsA[f'{col}{r}'].font = Font(name=F, size=10, color=INK, bold=(col == 'R'))
        wsA[f'{col}{r}'].border = BOX
    wsA[f'R{r}'].alignment = Alignment(horizontal='center')
wsA.column_dimensions['T'].hidden = True
LEVF = {'Crítica': (R_BG, R_TX), 'Alta': (W_BG, W_TX), 'Média': (LIGHT, CORP), 'Baixa': (ZEBRA, DARK)}
for lv, (bg, tx) in LEVF.items():
    wsA.conditional_formatting.add(f'R{L0}:R{L1}', FormulaRule(formula=[f'$R{L0}="{lv}"'], fill=PatternFill('solid', fgColor=bg),
                                                               font=Font(color=tx, bold=True)))
wsA.auto_filter.ref = f'A{HR}:S{lastA}'
CC = lambda col: f"'Cruzamento'!${col}${L0}:${col}${L1}"  # noqa: E731

# ---------- PLANO DE AÇÃO ----------
wpA = wb.create_sheet('Plano')
wpA.sheet_view.showGridLines = False
title_block(wpA, 'Plano de Ação Gerencial',
            'Uma linha por ocorrência: divergências da conciliação e importação parada no depósito 2008', 17)
cell(wpA, 'A4', 'Como usar: preencha somente as colunas azul-claras (Responsável, Prazo, Conclusão, Status e Evidência). '
     'As demais são fórmulas ligadas à conciliação e às bases, e se atualizam sozinhas.', size=9, color=INK3)
cell(wpA, 'A5', 'Nada foi inventado: responsáveis, prazos e evidências estão como "Pendente de preenchimento". '
     'A prioridade é a criticidade proposta (ver Dicionário de Dados) e deve ser validada.', size=9, color=INK3)
PH_ = ['Nº', 'Origem', 'SKU', 'Lote', 'UM', 'Descrição do problema', 'Causa identificada (calculada)',
       'Ação proposta (sugestão)', 'Quantidade', 'Prioridade', 'Responsável', 'Data de abertura',
       'Prazo previsto', 'Data de conclusão', 'Status', 'Situação do prazo', 'Evidência da conclusão']
PW_ = [9, 15, 20, 13, 6, 48, 48, 56, 12, 11, 22, 12, 12, 12, 14, 18, 30]
PR0 = 7
INPUT_COLS = {11, 13, 14, 15, 17}
for i, (h, w) in enumerate(zip(PH_, PW_), 1):
    hdr(wpA, PR0, i, h.upper(), fill=CORPF if i in INPUT_COLS else NAVF, align='center' if i in (1, 5, 9, 10, 12, 13, 14, 15, 16) else 'left')
    wpA.column_dimensions[get_column_letter(i)].width = w
wpA.row_dimensions[PR0].height = 30
occ = [('Conciliação ZBR', k) for k in lots_ord if abs(dif(k)) > 0.01]
occ += [('Depósito 2008', k) for k in pairs8 if p8[k] > 0.01]
for n, (orig, (m, lt)) in enumerate(occ, 1):
    r = PR0 + n
    wpA[f'A{r}'] = f'OC-{n:03d}'
    wpA[f'B{r}'], wpA[f'C{r}'], wpA[f'D{r}'], wpA[f'E{r}'] = orig, m, lt, um[(m, lt)]
    if orig == 'Conciliação ZBR':
        mt = f'MATCH(C{r}&"|"&D{r},{CC("S")},0)'
        wpA[f'F{r}'] = (f'=IFERROR(IF(INDEX({CC("J")},{mt})="Confere","Diferença zerada na última atualização",'
                        f'INDEX({CC("J")},{mt})&" "&FIXED(ABS(INDEX({CC("I")},{mt})),1)&" "&E{r}&" no estoque físico"),'
                        f'"SKU·lote não encontrado na conciliação")')
        wpA[f'G{r}'] = f'=IFERROR(INDEX({CC("K")},{mt}),"")'
        wpA[f'H{r}'] = f'=IFERROR(INDEX({CC("L")},{mt}),"")'
        wpA[f'I{r}'] = f'=IFERROR(INDEX({CC("I")},{mt}),0)'
        wpA[f'J{r}'] = f'=IFERROR(INDEX({CC("R")},{mt}),"")'
    else:
        wpA[f'F{r}'] = f'="Importação recebida no 2008 aguardando transferência: "&FIXED(I{r},1)&" "&E{r}'
        wpA[f'G{r}'] = 'Recebido no depósito 2008 e ainda não transferido para o ZBR (saldo pelo movimento do período)'
        wpA[f'H{r}'] = f'=IF(I{r}<={TOLR},"Nada a fazer: saldo zerado","Transferir "&FIXED(I{r},1)&" "&E{r}&" do 2008 para o ZBR (verificar OT aberta)")'
        wpA[f'I{r}'] = f'=SUMIFS({EB("G")},{EB("B")},C{r},{EB("C")},D{r})'
        wpA[f'J{r}'] = f'=IF(I{r}<={TOLR},"Baixa",IF(I{r}>={P_2008},"Alta","Média"))'
    wpA[f'K{r}'] = 'Pendente de preenchimento'
    wpA[f'L{r}'] = f'={P_DATA}'
    wpA[f'O{r}'] = 'Não iniciado'
    wpA[f'P{r}'] = (f'=IF(O{r}="Concluído","Concluída",IF(O{r}="Cancelado","Cancelada",IF(M{r}="","Sem prazo definido",'
                    f'IF(M{r}<{P_HOJE},"Vencida",IF(M{r}-{P_HOJE}<={P_DIAS},"Vence em breve","No prazo")))))')
    wpA[f'Q{r}'] = 'Pendente de preenchimento'
    for c in range(1, 18):
        x = wpA.cell(row=r, column=c)
        x.font = Font(name=F, size=10, color=INK, bold=(c in (1, 10)))
        x.border = Border(bottom=Side(style='thin', color='E3E8EF'))
        x.alignment = Alignment(horizontal='center' if c in (1, 5, 10, 12, 13, 14, 15, 16) else None, vertical='center')
        if c in INPUT_COLS:
            x.fill = PatternFill('solid', fgColor='EEF5FB')
    wpA[f'I{r}'].number_format = '+#,##0.0;-#,##0.0;0'
    for c in 'LMN':
        wpA[f'{c}{r}'].number_format = 'DD/MM/YYYY'
    for c in 'FGH':
        wpA[f'{c}{r}'].font = Font(name=F, size=9, color=INK)
PL1 = PR0 + len(occ)
tb = XTable(displayName='tbPlanoAcao', ref=f'A{PR0}:Q{PL1}')
tb.tableStyleInfo = XStyle(name='TableStyleLight1', showRowStripes=False)
wpA.add_table(tb)
dvs = DataValidation(type='list', formula1='"Não iniciado,Em andamento,Concluído,Cancelado"', allow_blank=False,
                     showErrorMessage=True, error='Use: Não iniciado, Em andamento, Concluído ou Cancelado.')
wpA.add_data_validation(dvs)
dvs.add(f'O{PR0 + 1}:O{PR0 + 500}')
dvd = DataValidation(type='date', operator='greaterThan', formula1='DATE(2020,1,1)', allow_blank=True,
                     showErrorMessage=True, error='Digite uma data válida (dd/mm/aaaa).')
wpA.add_data_validation(dvd)
dvd.add(f'M{PR0 + 1}:N{PR0 + 500}')
rngP = lambda col: f'{col}{PR0 + 1}:{col}{PL1}'  # noqa: E731
for v, bg, tx in (('Vencida', R_BG, R_TX), ('Vence em breve', W_BG, W_TX), ('Concluída', G_BG, G_TX),
                  ('No prazo', LIGHT, CORP), ('Sem prazo definido', ZEBRA, INK3), ('Cancelada', ZEBRA, INK3)):
    wpA.conditional_formatting.add(rngP('P'), FormulaRule(formula=[f'$P{PR0 + 1}="{v}"'], fill=PatternFill('solid', fgColor=bg), font=Font(color=tx, bold=True)))
for v, bg, tx in (('Concluído', G_BG, G_TX), ('Em andamento', W_BG, W_TX), ('Não iniciado', ZEBRA, INK3), ('Cancelado', ZEBRA, INK3)):
    wpA.conditional_formatting.add(rngP('O'), FormulaRule(formula=[f'$O{PR0 + 1}="{v}"'], fill=PatternFill('solid', fgColor=bg), font=Font(color=tx, bold=True)))
for lv, (bg, tx) in LEVF.items():
    wpA.conditional_formatting.add(rngP('J'), FormulaRule(formula=[f'$J{PR0 + 1}="{lv}"'], fill=PatternFill('solid', fgColor=bg), font=Font(color=tx, bold=True)))
for col in 'KQ':
    wpA.conditional_formatting.add(rngP(col), FormulaRule(formula=[f'{col}{PR0 + 1}="Pendente de preenchimento"'], font=Font(color=INK3, italic=True)))
wpA.conditional_formatting.add(f'A{PR0 + 1}:Q{PL1}', FormulaRule(formula=[f'$P{PR0 + 1}="Vencida"'], font=Font(color=R_TX)))
wpA.freeze_panes = f'D{PR0 + 1}'
PLC = lambda col: f"'Plano'!${col}${PR0 + 1}:${col}${PR0 + 500}"  # noqa: E731

# ---------- DASHBOARD EXECUTIVO ----------
wd = wb.create_sheet('Dashboard')
wd.sheet_view.showGridLines = False
for i in range(1, 17):
    wd.column_dimensions[get_column_letter(i)].width = 12.5
title_block(wd, 'Dashboard Executivo — ZBR Importado',
            'Centro BR01 · movimentos de 09/09 a 01/10/2026 · fonte: exportação SAP "Analise_ZBR_importado.XLSX" · '
            'relatório gerado em 02/10/2026', 16)
PF = lambda col: f"Painel!{MR(col)}"  # noqa: E731
# filtros (fonte única para todas as análises)
for a, b, lab, v, fmt in (('A', 'B', 'DATA INICIAL', pd.Timestamp('2026-09-09').to_pydatetime(), 'DD/MM/YYYY'),
                          ('C', 'D', 'DATA FINAL', pd.Timestamp('2026-10-01').to_pydatetime(), 'DD/MM/YYYY'),
                          ('E', 'G', 'SKU (CATEGORIA)', 'Todos', '@'), ('H', 'I', 'LOTE', 'Todos', '@'),
                          ('J', 'J', 'UNIDADE', 'KG', '@'), ('K', 'M', 'TIPO DE MOVIMENTAÇÃO', 'Todas', '@'),
                          ('N', 'P', 'STATUS (SITUAÇÃO)', 'Todos', '@')):
    cell(wd, f'{a}4', lab, size=8, bold=True, color=INK3)
    cell(wd, f'{a}5', v, size=11, bold=True, color=NAVY, fmt=fmt)
    if a != b:
        wd.merge_cells(f'{a}5:{b}5')
    for col in range(ord(a), ord(b) + 1):
        wd[f'{chr(col)}5'].fill, wd[f'{chr(col)}5'].border = INPUT, INBOX
    wd[f'{a}5'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
wd.row_dimensions[5].height = 24
for a in ('E5', 'H5', 'K5', 'N5'):
    wd[a].number_format = '@"   ▼"'
wd['J5'].number_format = '@" ▼"'
for a in ('A5', 'C5'):
    wd[a].number_format = 'DD/MM/YYYY" ▼"'
for i, v in enumerate(['Todos', 'Confere', 'Sobrando', 'Faltando'], 1):
    wsP[f'BG{i}'] = v
wb.defined_names['ListaStatus'] = DefinedName('ListaStatus', attr_text="Painel!$BG$1:$BG$4")
DV_MAP = {'A5:C5': 'E5:G5', 'D5:E5': 'H5:I5', 'F5:G5': 'A5:B5', 'H5:I5': 'C5:D5', 'J5': 'J5', 'K5:M5': 'K5:M5'}
for rng, name, msg in DV_DASH + [('N5', 'ListaStatus', 'Confere, Sobrando ou Faltando (situação do SKU·lote).')]:
    dv = DataValidation(type='list', formula1=f'={name}', allow_blank=False, showErrorMessage=True,
                        errorTitle='Valor inválido', error='Escolha um item da lista.',
                        showInputMessage=True, promptTitle='Filtro', prompt=msg)
    wd.add_data_validation(dv)
    dv.add(DV_MAP.get(rng, 'N5:P5'))
wd['A6'] = ('=Painel!A7&IF(N5="Todos",""," · status "&N5)')
wd['A6'].font = Font(name=F, size=9, bold=True, color=INK2)
cell(wd, 'K6', 'Clique numa célula azul-clara e use a seta ▼. Os filtros valem para todas as análises.', size=8, color=INK3)


def kpi(ws, r, a, b, label, formula, hint, fmt=NUM, color=NAVY, accent=CORP):
    for rr in (r, r + 1, r + 2):
        ws.merge_cells(f'{a}{rr}:{b}{rr}')
    for col in range(ord(a), ord(b) + 1):
        for rr in (r, r + 1, r + 2):
            c = ws[f'{chr(col)}{rr}']
            c.fill = WHF
            c.border = Border(left=Side(style='thick', color=accent) if col == ord(a) else None,
                              right=Side(style='thin', color=LINE) if col == ord(b) else None,
                              top=Side(style='thin', color=LINE) if rr == r else None,
                              bottom=Side(style='thin', color=LINE) if rr == r + 2 else None)
    cell(ws, f'{a}{r}', label, size=8, bold=True, color=INK3)
    cell(ws, f'{a}{r + 1}', formula, size=22, bold=True, color=color, fmt=fmt)
    cell(ws, f'{a}{r + 2}', hint, size=8, color=INK3)
    for rr in (r, r + 1, r + 2):
        ws[f'{a}{rr}'].alignment = Alignment(horizontal='left', vertical='center', indent=1)


cell(wd, 'A7', 'INDICADORES PRINCIPAIS', size=11, bold=True, color=CORP)
reg = f'SUM({PF("AA")})'
kpi(wd, 8, 'A', 'B', '="VOLUME ANALISADO · "&Painel!$R$3', f'=SUMPRODUCT({PF("AA")},{PF("W")})+SUMPRODUCT({PF("AA")},{PF("X")})', 'entrou + saiu no ZBR')
kpi(wd, 8, 'C', 'D', 'SKU·LOTE ANALISADOS', f'={reg}', 'registros no filtro', fmt='#,##0')
kpi(wd, 8, 'E', 'F', 'MOVIMENTAÇÕES', f'=SUMPRODUCT({PF("AA")},{PF("AD")})', 'lançamentos no filtro', fmt='#,##0')
kpi(wd, 8, 'G', 'H', 'DIVERGÊNCIAS', f'=COUNTIFS({PF("AA")},1,{PF("AC")},"<>Confere")', 'SKU·lote com diferença', fmt='#,##0', accent=ATT_C)
kpi(wd, 8, 'I', 'J', '% CONFORMIDADE', f'=IF({reg}=0,0,COUNTIFS({PF("AA")},1,{PF("AC")},"Confere")/{reg})', 'SKU·lote que conferem', fmt='0.0%', accent=OK_C)
kpi(wd, 8, 'K', 'L', 'PENDÊNCIAS EM ABERTO', f'=COUNTIFS({PLC("A")},"OC-*",{PLC("O")},"<>Concluído",{PLC("O")},"<>Cancelado")', 'ações do plano (geral)', fmt='#,##0', accent=ATT_C)
kpi(wd, 8, 'M', 'N', 'OCORRÊNCIAS CRÍTICAS', f'=COUNTIFS({PF("AA")},1,{PF("BF")},"Crítica")', 'criticidade proposta', fmt='#,##0', color=CRIT_C, accent=CRIT_C)
kpi(wd, 8, 'O', 'P', '% CONCLUSÃO DAS AÇÕES', f'=IF(COUNTIF({PLC("A")},"OC-*")=0,0,COUNTIF({PLC("O")},"Concluído")/COUNTIF({PLC("A")},"OC-*"))', 'plano de ação (geral)', fmt='0.0%', accent=OK_C)
wd.conditional_formatting.add('I9', FormulaRule(formula=['$I$9>=0.9'], font=Font(color=OK_C)))
wd.conditional_formatting.add('I9', FormulaRule(formula=['AND($I$9>=0.7,$I$9<0.9)'], font=Font(color='BF8F00')))
wd.conditional_formatting.add('I9', FormulaRule(formula=['$I$9<0.7'], font=Font(color=CRIT_C)))
cell(wd, 'A12', '="QUANTIDADES · "&Painel!$R$3', size=11, bold=True, color=CORP)
kpi(wd, 13, 'A', 'C', 'ENTROU NO ZBR', f'=SUMPRODUCT({PF("AA")},{PF("W")})', 'transferido do 2008')
kpi(wd, 13, 'D', 'F', 'SAIU', f'=SUMPRODUCT({PF("AA")},{PF("X")})', 'vendas e entregas (601)')
kpi(wd, 13, 'G', 'I', 'DEVERIA TER', '=A14-D14', 'entrou − saiu')
kpi(wd, 13, 'J', 'L', 'ESTOQUE FÍSICO', f'=SUMPRODUCT({PF("AA")},{PF("Y")})', 'posição atual no depósito')
kpi(wd, 13, 'M', 'N', 'DIFERENÇA', '=ROUND(J14-G14,3)', 'físico − deveria ter', fmt='+#,##0.0;-#,##0.0;0', accent=ATT_C)
kpi(wd, 13, 'O', 'P', 'AGUARDANDO NO 2008',
    f"=SUMPRODUCT('Depósito 2008'!{Q('Z')}*('Depósito 2008'!{Q('Y')}>0)*'Depósito 2008'!{Q('Y')})",
    'importação não transferida', accent=ATT_C)
wd.conditional_formatting.add('M14', FormulaRule(formula=[f'ABS($M$14)<={TOLR}'], font=Font(color=OK_C)))
wd.conditional_formatting.add('M14', FormulaRule(formula=[f'$M$14<-{TOLR}'], font=Font(color=CRIT_C)))
wd.conditional_formatting.add('M14', FormulaRule(formula=[f'$M$14>{TOLR}'], font=Font(color='BF8F00')))
# dados dos gráficos (colunas ocultas T:W)
DR = 20
wd['T19'], wd['U19'] = 'Categoria da causa', 'Ocorrências'
for i, c_ in enumerate(CATS):
    wd[f'T{DR + i}'] = c_
    wd[f'U{DR + i}'] = f'=COUNTIFS({PF("AA")},1,{PF("BE")},T{DR + i})'
wd['T30'], wd['U30'] = 'Situação', 'SKU·lote'
for i, c_ in enumerate(['Confere', 'Sobrando', 'Faltando']):
    wd[f'T{31 + i}'] = c_
    wd[f'U{31 + i}'] = f'=COUNTIFS({PF("AA")},1,{PF("AC")},T{31 + i})'
wd['T36'], wd['U36'] = 'Criticidade', 'Ocorrências'
for i, c_ in enumerate(LEVELS):
    wd[f'T{37 + i}'] = c_
    wd[f'U{37 + i}'] = f'=COUNTIFS({PF("AA")},1,{PF("BF")},T{37 + i})'
wd['T43'], wd['U43'] = 'Status do plano', 'Ações'
for i, c_ in enumerate(['Não iniciado', 'Em andamento', 'Concluído', 'Cancelado']):
    wd[f'T{44 + i}'] = c_
    wd[f'U{44 + i}'] = f'=COUNTIF({PLC("O")},T{44 + i})'
for col in 'TUVW':
    wd.column_dimensions[col].hidden = True
cell(wd, 'A17', 'GRÁFICOS GERENCIAIS', size=11, bold=True, color=CORP)
CH_H, CH_W = 8.6, 17.6
k1 = BarChart()
k1.type, k1.grouping, k1.overlap, k1.gapWidth = 'col', 'clustered', 100, 40
k1.title = 'Movimentações por período (entrou × saiu por dia)'
k1.add_data(Reference(wsP, min_col=42, min_row=D0 - 1, max_row=D1), titles_from_data=True)
k1.add_data(Reference(wsP, min_col=43, min_row=D0 - 1, max_row=D1), titles_from_data=True)
k1.set_categories(Reference(wsP, min_col=40, min_row=D0, max_row=D1))
k1.series[0].graphicalProperties.solidFill = CORP
k1.series[1].graphicalProperties.solidFill = C_SAI
for sr in k1.series:
    sr.graphicalProperties.line.noFill = True
style_chart(k1, legend=True)
k1.x_axis.tickLblPos = 'low'
k1.height, k1.width = CH_H, CH_W
wd.add_chart(k1, 'A18')
wsP[f'AR{D0 - 1}'] = 'Saldo acumulado (entrou − saiu)'
k2 = LineChart()
k2.title = 'Evolução do saldo acumulado no ZBR'
k2.add_data(Reference(wsP, min_col=44, min_row=D0 - 1, max_row=D1), titles_from_data=True)
k2.set_categories(Reference(wsP, min_col=40, min_row=D0, max_row=D1))
k2.series[0].graphicalProperties.line.solidFill = NAVY
k2.series[0].graphicalProperties.line.width = 28000
k2.series[0].smooth = False
style_chart(k2)
k2.height, k2.width = CH_H, CH_W
wd.add_chart(k2, 'I18')
k3 = easy_bar('Divergências por categoria de causa', Reference(wd, min_col=20, min_row=DR, max_row=DR + len(CATS) - 1),
              Reference(wd, min_col=21, min_row=DR, max_row=DR + len(CATS) - 1),
              [CRIT_C, CRIT_C, ATT_C, ATT_C, CORP, GREY_C, CORP, ATT_C], horizontal=True)
k3.height, k3.width = CH_H, CH_W
wd.add_chart(k3, 'A36')
k4 = easy_bar('Regular × irregular (SKU·lote)', Reference(wd, min_col=20, min_row=31, max_row=33),
              Reference(wd, min_col=21, min_row=31, max_row=33), [OK_C, ATT_C, CRIT_C])
k4.height, k4.width = CH_H, CH_W
wd.add_chart(k4, 'I36')
k5 = BarChart()
k5.type, k5.grouping, k5.overlap, k5.gapWidth = 'bar', 'clustered', 100, 50
k5.title = 'Ranking: maiores diferenças por SKU · lote'
k5.add_data(Reference(wsP, min_col=37, min_row=M0 - 1, max_row=M0 + 11), titles_from_data=True)
k5.add_data(Reference(wsP, min_col=38, min_row=M0 - 1, max_row=M0 + 11), titles_from_data=True)
k5.set_categories(Reference(wsP, min_col=36, min_row=M0, max_row=M0 + 11))
k5.series[0].graphicalProperties.solidFill = ATT_C
k5.series[1].graphicalProperties.solidFill = CRIT_C
for sr in k5.series:
    sr.graphicalProperties.line.noFill = True
style_chart(k5, legend=True)
k5.x_axis.scaling.orientation = 'maxMin'
k5.x_axis.tickLblPos = 'low'
k5.height, k5.width = CH_H, CH_W
wd.add_chart(k5, 'A54')
k6 = easy_bar('Distribuição por criticidade (proposta)', Reference(wd, min_col=20, min_row=37, max_row=40),
              Reference(wd, min_col=21, min_row=37, max_row=40), [CRIT_C, ATT_C, CORP, GREY_C])
k6.height, k6.width = CH_H, CH_W
wd.add_chart(k6, 'I54')
k7 = easy_bar('Plano de ação: ocorrências por status', Reference(wd, min_col=20, min_row=44, max_row=47),
              Reference(wd, min_col=21, min_row=44, max_row=47), [GREY_C, ATT_C, OK_C, DARK])
k7.height, k7.width = CH_H, CH_W
wd.add_chart(k7, 'A72')
cell(wd, 'I73', 'Leitura rápida', size=11, bold=True, color=CORP)
LR = ['Azul = entrou no ZBR · azul-claro = saiu. Verde = regular · amarelo = atenção · vermelho = crítico.',
      'Todos os números seguem os filtros do topo, exceto Pendências e % Conclusão (plano de ação geral).',
      'Detalhe por SKU·lote: abas "4 Análise ZBR" e "5 Conciliação". Ações: aba "6 Plano de Ação".',
      'Critérios de cálculo de cada indicador: aba "7 Dicionário de Dados".']
for i, t in enumerate(LR):
    cell(wd, f'I{74 + i}', '• ' + t, size=9, color=INK)
wd.freeze_panes = 'A7'

# ---------- RESUMO GERENCIAL ----------
wr = wb.create_sheet('ResumoG')
wr.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGHI', [3, 44, 16, 16, 16, 16, 18, 18, 18]):
    wr.column_dimensions[col].width = w
title_block(wr, 'Resumo Gerencial', 'Síntese das análises, principais ocorrências e conclusões · visão geral (sem filtros)', 9)
row = 5


def sec(title):
    global row
    cell(wr, f'B{row}', title, size=13, bold=True, color=CORP)
    wr[f'B{row}'].border = Border(bottom=Side(style='thin', color=CORP))
    row += 1


def table(headers, rows, fmts=None):
    global row
    for i, h in enumerate(headers):
        hdr(wr, row, 2 + i, h, align='left' if i == 0 else 'right')
    row += 1
    for vals in rows:
        for i, v in enumerate(vals):
            c = wr.cell(row=row, column=2 + i, value=v)
            c.font = Font(name=F, size=10, color=INK, bold=(i == 0))
            c.border = Border(bottom=Side(style='thin', color='E3E8EF'))
            if fmts and fmts[i]:
                c.number_format = fmts[i]
            if i > 0:
                c.alignment = Alignment(horizontal='right')
        row += 1
    row += 1


sec('1. Escopo da análise')
table(['Item', 'Valor'], [
    ['Primeira data de movimento', f'=MIN(MIN({EB("F")}),MIN({SB("G")}))'],
    ['Última data de movimento', f'=MAX(MAX({EB("F")}),MAX({SB("G")}))'],
    ['SKU·lote conciliados', f'=COUNTA({CC("A")})'],
    ['Lançamentos no depósito 2008 (base Entrada)', f'=COUNTA({EB("A")})'],
    ['Lançamentos de saída 601 (base Saídas)', f'=COUNTA({SB("A")})'],
    ['Posições de estoque físico (base Estoque)', f"=COUNTA({ES('A')})"]],
    [None, '#,##0'])
for rr in (row - 7, row - 6):
    wr[f'C{rr}'].number_format = 'DD/MM/YYYY'
sec('2. Resultado da conciliação por unidade de medida')
r_conc = row + 1
table(['Unidade', 'SKU·lote', 'Conferem', 'Divergentes', '% conformidade', 'Diferença líquida', 'Sem explicação'],
      [[u, f'=COUNTIFS({CC("D")},"{u}")', f'=COUNTIFS({CC("D")},"{u}",{CC("J")},"Confere")',
        f'=C{r_conc + i}-D{r_conc + i}', f'=IF(C{r_conc + i}=0,0,D{r_conc + i}/C{r_conc + i})',
        f'=SUMIFS({CC("I")},{CC("D")},"{u}")', f'=SUMIFS({CC("O")},{CC("D")},"{u}")'] for i, u in enumerate(['KG', 'UN', 'CX'])],
      [None, '#,##0', '#,##0', '#,##0', '0.0%', '+#,##0.0;-#,##0.0;0', '+#,##0.0;-#,##0.0;0'])
TOT_R = row
cell(wr, f'B{row - 1}', 'Total', bold=True)
for col in 'CDE':
    cell(wr, f'{col}{row - 1}', f'=SUM({col}{r_conc}:{col}{r_conc + 2})', bold=True, fmt='#,##0', align='right')
cell(wr, f'F{row - 1}', f'=IF(C{row - 1}=0,0,D{row - 1}/C{row - 1})', bold=True, fmt='0.0%', align='right')
TOTROW = row - 1
row += 1
sec('3. Divergências por categoria de causa')
r_cat = row + 1
table(['Categoria', 'Ocorrências', 'Diferença (KG)'],
      [[c_, f'=COUNTIFS({CC("Q")},B{r_cat + i})', f'=SUMIFS({CC("I")},{CC("Q")},B{r_cat + i},{CC("D")},"KG")'] for i, c_ in enumerate(CATS)],
      [None, '#,##0', '+#,##0.0;-#,##0.0;0'])
sec('4. Ocorrências por criticidade (proposta para validação)')
r_crit = row + 1
table(['Criticidade', 'Ocorrências', 'Critério'],
      [['Crítica', f'=COUNTIF({CC("R")},"Crítica")', '="Não localizado no estoque, ou falta sem explicação ≥ "&' + P_CRIT],
       ['Alta', f'=COUNTIF({CC("R")},"Alta")', '="Outras faltas sem explicação, ou sobra sem explicação ≥ "&' + P_ALTA],
       ['Média', f'=COUNTIF({CC("R")},"Média")', 'Sobra com parte sem explicação abaixo do limite'],
       ['Baixa', f'=COUNTIF({CC("R")},"Baixa")', 'Diferença totalmente explicada pelos dados']],
      [None, '#,##0', None])
for i in range(4):
    wr[f'D{r_crit + i}'].alignment = Alignment(horizontal='left')
sec('5. Cinco maiores ocorrências (diferença sem explicação)')
r_top = row + 1
top_rows = []
for k in range(1, 6):
    mt = f'MATCH(LARGE({CC("T")},{k}),{CC("T")},0)'
    top_rows.append([f'=IFERROR(INDEX({CC("A")},{mt}),"")', f'=IFERROR(INDEX({CC("B")},{mt}),"")',
                     f'=IFERROR(INDEX({CC("D")},{mt}),"")', f'=IFERROR(INDEX({CC("I")},{mt}),"")',
                     f'=IFERROR(INDEX({CC("O")},{mt}),"")', f'=IFERROR(INDEX({CC("Q")},{mt}),"")',
                     f'=IFERROR(INDEX({CC("R")},{mt}),"")'])
table(['SKU', 'Lote', 'UM', 'Diferença', 'Sem explicação', 'Categoria', 'Criticidade'], top_rows,
      [None, None, None, '+#,##0.0;-#,##0.0;0', '+#,##0.0;-#,##0.0;0', None, None])
sec('6. Depósito 2008 (importação) · KG')
r_08 = row + 1
table(['Indicador', 'Valor'], [
    ['Recebido no 2008', f'=SUMIFS({EB("G")},{EB(cK)},"Recebido no 2008",{EB("H")},"KG")'],
    ['Transferido ao ZBR', f'=-SUMIFS({EB("G")},{EB(cK)},"Entrada",{EB("H")},"KG")'],
    ['Estornos (voltou ao 2008)', f'=SUMIFS({EB("G")},{EB(cK)},"Estorno",{EB("H")},"KG")'],
    ['Saldo do período no 2008', f'=C{r_08}-C{r_08 + 1}+C{r_08 + 2}'],
    ['Lotes aguardando transferência (todas as unidades)', f'=COUNTIF({PLC("B")},"Depósito 2008")'],
    ['Volume aguardando transferência (KG)', f'=SUMIFS({PLC("I")},{PLC("B")},"Depósito 2008",{PLC("E")},"KG")']],
    [None, '#,##0.0'])
sec('7. Plano de ação')
r_pl = row + 1
table(['Indicador', 'Valor'], [
    ['Ocorrências no plano', f'=COUNTIF({PLC("A")},"OC-*")'],
    ['Não iniciadas', f'=COUNTIF({PLC("O")},"Não iniciado")'],
    ['Em andamento', f'=COUNTIF({PLC("O")},"Em andamento")'],
    ['Concluídas', f'=COUNTIF({PLC("O")},"Concluído")'],
    ['% concluído', f'=IF(C{r_pl}=0,0,C{r_pl + 3}/C{r_pl})'],
    ['Vencidas', f'=COUNTIF({PLC("P")},"Vencida")'],
    ['Sem responsável definido', f'=COUNTIF({PLC("K")},"Pendente de preenchimento")']],
    [None, '#,##0'])
wr[f'C{r_pl + 4}'].number_format = '0.0%'
sec('8. Conclusões (geradas a partir dos números acima)')
CONC = [
    f'="Conformidade: "&D{TOTROW}&" de "&C{TOTROW}&" SKU·lote conferem ("&FIXED(F{TOTROW}*100,1)&"%); "&E{TOTROW}&" apresentam diferença."',
    f'="Principal causa: "&INDEX(B{r_cat}:B{r_cat + 7},MATCH(MAX(C{r_cat}:C{r_cat + 7}),C{r_cat}:C{r_cat + 7},0))&" ("&MAX(C{r_cat}:C{r_cat + 7})&" ocorrências)."',
    f'="Ocorrências críticas: "&C{r_crit}&" (faltas relevantes ou mercadoria não localizada) — tratar antes das sobras."',
    f'="Maior ocorrência: SKU "&B{r_top}&" lote "&C{r_top}&", "&FIXED(F{r_top},1)&" "&D{r_top}&" sem explicação ("&G{r_top}&")."',
    f'="Depósito 2008: "&FIXED(C{r_08 + 5},1)&" KG de importação aguardam transferência para o ZBR em "&C{r_08 + 4}&" lotes."',
    f'="Plano de ação: "&(C{r_pl}-C{r_pl + 3})&" ações em aberto; "&C{r_pl + 6}&" sem responsável definido."',
    'As sobras podem incluir estoque anterior a 09/09, que não consta no arquivo de origem: confirmar o saldo inicial no SAP antes de lançar ajustes.']
for t in CONC:
    c = cell(wr, f'B{row}', t, size=10, color=INK)
    wr.merge_cells(f'B{row}:I{row}')
    c.alignment = Alignment(wrap_text=True, vertical='top')
    wr.row_dimensions[row].height = 18
    row += 1
wr.freeze_panes = 'A4'

# ---------- DICIONÁRIO DE DADOS ----------
wdc = wb.create_sheet('Dicionario')
wdc.sheet_view.showGridLines = False
for col, w in zip('ABCD', [30, 34, 70, 60]):
    wdc.column_dimensions[col].width = w
title_block(wdc, 'Dicionário de Dados', 'Abas, campos, regras de negócio, critérios de cálculo, atualização e limitações', 4)
drow = 5


def dsec(title, headers, rows_):
    global drow
    cell(wdc, f'A{drow}', title, size=13, bold=True, color=CORP)
    drow += 1
    for i, h in enumerate(headers):
        hdr(wdc, drow, 1 + i, h)
    drow += 1
    for vals in rows_:
        for i, v in enumerate(vals):
            c = wdc.cell(row=drow, column=1 + i, value=v)
            c.font = Font(name=F, size=10, color=INK, bold=(i == 0))
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.border = Border(bottom=Side(style='thin', color='E3E8EF'))
            if (drow % 2) == 0:
                c.fill = ZEBF
        drow += 1
    drow += 1


SHEETS_DOC = [
    ('MENU', 'Página inicial: identificação do relatório e navegação.', '—'),
    ('1 Dashboard', 'Indicadores principais, gráficos gerenciais e FILTROS ÚNICOS (período, SKU, lote, unidade, tipo de movimentação, status).', 'Fórmulas sobre 4 Análise ZBR, 4 Análise Depósito 2008 e 6 Plano de Ação'),
    ('2 Resumo Gerencial', 'Síntese geral (sem filtros): conciliação por unidade, causas, criticidade, top 5, depósito 2008, plano e conclusões.', 'Fórmulas sobre 5 Conciliação, bases e plano'),
    ('3 Base Entrada 2008', 'Movimentos do depósito 2008 exportados do SAP (tabela tbEntrada2008) + Movimentação e Entrou no ZBR calculados.', 'Aba "Entrada" do arquivo SAP'),
    ('3 Base Saídas', 'Saídas 601 exportadas do SAP (tabela tbSaidas) + coluna Saiu. Linhas de total do rodapé original removidas.', 'Aba "Saida" do arquivo SAP'),
    ('3 Base Estoque Físico', 'Posição de estoque por posição de depósito (tabela tbEstoque). Coluna 100% vazia "Inventário ativo" removida.', 'Aba "Estoque atual" do arquivo SAP'),
    ('3 Base Movimentos', 'Visão consolidada de entradas e saídas (tabela tbMovimentos), uma linha por lançamento.', 'Fórmulas sobre 3 Base Entrada 2008 e 3 Base Saídas'),
    ('4 Análise ZBR', 'Tabela por SKU + lote segundo os filtros do Dashboard: entrou, saiu, deveria ter, estoque físico, diferença, motivo e solução.', 'Bases + 5 Conciliação'),
    ('4 Análise Depósito 2008', 'Recebido, transferido, estornos e saldo do 2008 por SKU + lote, situação e próximo passo; gráficos do 2008 abaixo da tabela.', '3 Base Entrada 2008'),
    ('5 Conciliação', 'Conciliação do período completo por SKU + lote: diferença, motivo, solução, componentes da explicação, categoria e criticidade.', 'Bases (via abas auxiliares ocultas)'),
    ('6 Plano de Ação', 'Ocorrências com causa, ação, prioridade, responsável, prazos, status e alertas de vencimento.', '5 Conciliação + 3 Base Entrada 2008 + preenchimento manual'),
    ('7 Dicionário de Dados', 'Este documento.', '—'),
    ('Parâmetros', 'Data-limite de recebimento, tolerância, limites de criticidade, data do relatório e data de referência. Conferência de totais.', 'Edição manual'),
    ('aux Conciliação Material / aux Conciliação Lote (ocultas)', 'Cálculos intermediários da conciliação (SUMIFS direto nas bases).', 'Bases')]
dsec('1. Mapa das abas', ['Aba', 'Finalidade', 'Origem dos dados'], [(a, b, c) for a, b, c in SHEETS_DOC])
FIELDS = [
    ('Todas', 'SKU', 'Código do material no SAP (campo "Material" da exportação).', 'Texto, mantido como na origem'),
    ('Todas', 'Lote', 'Lote SAP. Mantido como texto para preservar zeros e não virar número.', 'Texto (intencional)'),
    ('Bases', 'Depósito / Doc.material / Cliente / Referência', 'Códigos SAP mantidos como texto (ex.: depósito 0355).', 'Texto (intencional)'),
    ('3 Base Entrada 2008', 'Movimentação', 'Recebido no 2008 (positivo até a data-limite), Entrada (negativo = transferência para o ZBR) ou Estorno (positivo após a data-limite).', 'Fórmula; data-limite em Parâmetros!C6'),
    ('3 Base Entrada 2008', 'Entrou no ZBR', 'Quantidade que entrou no ZBR: −Quantidade, exceto recebimentos (0). Estornos entram negativos.', 'Fórmula'),
    ('3 Base Saídas', 'Saiu', 'Quantidade positiva da saída 601 (−Quantidade).', 'Fórmula'),
    ('4 Análise ZBR / 5 Conciliação', 'Entrou / Saiu / Deveria ter', 'Entrou no ZBR, saídas 601 e Deveria ter = Entrou − Saiu.', 'SUMIFS nas bases'),
    ('4 Análise ZBR / 5 Conciliação', 'Estoque físico', 'Soma do estoque total das posições do SKU + lote (inclui doca e transferência).', 'SUMIFS na base de estoque'),
    ('4 Análise ZBR / 5 Conciliação', 'Diferença', 'Estoque físico − Deveria ter. Zero (± tolerância) = Confere; positivo = Sobrando; negativo = Faltando.', 'Tolerância em Parâmetros!C7'),
    ('5 Conciliação', 'Estoque antigo (antes de 09/09)', 'Estoque em posições cujo último movimento é anterior a 09/09/2026.', 'SUMIFS por data do último movimento'),
    ('5 Conciliação', 'Em doca / chão / transferência', 'Estoque em posições DCK, TRF, 922, FLR ou DIF movimentadas a partir de 09/09.', 'SUMPRODUCT por tipo de depósito'),
    ('5 Conciliação', 'Motivo da diferença', 'Primeira regra que explica a diferença (ver seção 3).', 'Fórmula'),
    ('5 Conciliação', 'Diferença sem explicação', 'Parte da diferença que nenhuma regra explica.', 'Fórmula'),
    ('5 Conciliação', 'Categoria da causa', 'Agrupamento do motivo em 8 categorias para gráficos e resumo.', 'Fórmula'),
    ('5 Conciliação', 'Criticidade (proposta)', 'Classificação proposta para validação (seção 4).', 'Fórmula + limites em Parâmetros'),
    ('4 Análise Depósito 2008', 'Saldo no 2008', 'Recebido − transferido + estornos no período. Negativo = saiu estoque que já estava no 2008 antes de 09/09.', 'SUMIFS na base Entrada'),
    ('6 Plano de Ação', 'Situação do prazo', 'Concluída / Cancelada / Sem prazo definido / Vencida / Vence em breve / No prazo, comparando o prazo com Parâmetros!K10 (hoje).', 'Fórmula'),
    ('1 Dashboard', 'Volume analisado', 'Entrou + Saiu no ZBR no filtro (unidade escolhida).', 'SUMPRODUCT na 4 Análise ZBR'),
    ('1 Dashboard', '% conformidade', 'SKU·lote que conferem ÷ SKU·lote analisados no filtro.', 'COUNTIFS'),
    ('1 Dashboard', 'Pendências em aberto / % conclusão', 'Ações do plano com status diferente de Concluído/Cancelado; concluídas ÷ total.', 'COUNTIFS no plano (sem filtro)')]
dsec('2. Campos e cálculos', ['Aba', 'Campo', 'Descrição', 'Regra / cálculo'], FIELDS)
RULES = [
    ('Depósito 2008', 'O arquivo traz os movimentos do depósito 2008. Negativo = transferência para o ZBR (depósitos 9999/0355).', 'Inferido dos dados: 66 de 94 SKU·lote fecham exatamente com essa regra.', ''),
    ('Recebimento x estorno', 'Positivos no 2008 até 15/09 são chegada da importação; depois, estornos de transferência.', 'Data-limite editável em Parâmetros!C6.', ''),
    ('Ordem do motivo', '1 Troca de lote · 2 Não localizado · 3 Falta · 4 Estoque antigo · 5 Saídas sem baixa · 6 Doca/transferência · 7 Combinações · 8 Explica em parte · 9 Saiu mais que entrou · 10 Acima do esperado.', 'A primeira regra verdadeira define o motivo.', ''),
    ('Período', 'Os motivos e a conciliação usam o período completo do arquivo; os filtros de data afetam só as análises filtradas.', 'O estoque físico é sempre a posição atual.', '')]
dsec('3. Regras de negócio', ['Regra', 'Descrição', 'Observação', ''], RULES)
CRIT_DOC = [
    ('Crítica', 'Mercadoria que entrou no ZBR e não está no estoque, ou falta sem explicação ≥ limite crítico.', 'Parâmetros!K6 (padrão 200, na unidade do SKU)', 'PROPOSTA — validar com a gestão'),
    ('Alta', 'Demais faltas sem explicação, ou sobra sem explicação ≥ limite alto. No 2008: saldo ≥ Parâmetros!K8.', 'Parâmetros!K7 (padrão 500) e K8 (padrão 2.000)', 'PROPOSTA — validar com a gestão'),
    ('Média', 'Sobra com parte sem explicação abaixo do limite alto; no 2008, saldo aguardando abaixo do limite.', '—', 'PROPOSTA — validar com a gestão'),
    ('Baixa', 'Diferença totalmente explicada pelos dados (ex.: troca de lote, estoque antigo, doca).', '—', 'PROPOSTA — validar com a gestão')]
dsec('4. Critérios de criticidade (proposta para validação)', ['Nível', 'Critério', 'Parâmetro', 'Status'], CRIT_DOC)
UPD = [
    ('1', 'Exporte do SAP as três listas (Entrada do 2008, Saídas 601, Estoque atual) com as mesmas colunas.', '', ''),
    ('2', 'Cole os dados nas abas 3 Base (a partir da linha 2), mantendo a ordem das colunas. As fórmulas cobrem até a linha 5.000.', 'Copie também as fórmulas das colunas calculadas (Movimentação, Entrou no ZBR, Saiu) para as linhas novas.', ''),
    ('3', 'Os indicadores, a conciliação e o plano se recalculam sozinhos para os SKU·lote já existentes.', '', ''),
    ('4', 'SKU·lote novos só aparecem nas tabelas de análise depois de gerar o relatório de novo com o script scripts/gerar_relatorio.py do repositório.', 'Comando: python scripts/gerar_relatorio.py arquivo_sap.xlsx saida.xlsx', ''),
    ('5', 'Atualize os parâmetros (data-limite, limites de criticidade, data do relatório) se necessário.', '', '')]
dsec('5. Como atualizar a base', ['Passo', 'O que fazer', 'Observação', ''], UPD)
LIM = [
    ('Saldo inicial', 'O arquivo de origem não traz o estoque antes de 09/09 nem a foto do estoque do 2008. Sobras podem ser saldo inicial.', 'Recomendação: incluir a posição de estoque de 08/09 na próxima extração.', ''),
    ('Linhas de análise', 'As listas de SKU·lote das análises são geradas pelo script. Novos SKU·lote exigem gerar o relatório de novo.', 'Recomendação: migrar a importação para Power Query.', ''),
    ('Tabela/Gráfico Dinâmico', 'Não foram usados Tabelas Dinâmicas nem segmentações (slicers): a geração automática não as suporta de forma confiável. Os filtros são listas suspensas ligadas a fórmulas.', 'Os gráficos acompanham os filtros do Dashboard.', ''),
    ('Responsável', 'O filtro por responsável está disponível no cabeçalho da tabela do Plano de Ação (não há responsável nos dados do SAP).', '', ''),
    ('Armazém', 'A análise considera o centro BR01 e os depósitos ZBR (9999/0355) juntos; não há filtro por depósito.', '', ''),
    ('Data de atualização', 'A data exibida é a da geração do relatório (Parâmetros!K9). A hora da extração SAP não está no arquivo.', '', ''),
    ('Funções', 'Usa MINIFS/MAXIFS (Excel 2019/365). Em versões anteriores essas duas colunas de data mostram erro.', '', '')]
dsec('6. Limitações e recomendações', ['Tema', 'Descrição', 'Recomendação', ''], LIM)

# ---------- bases: tabelas estruturadas, zebra, formatos, larguras ----------
def autofit(ws, ncols, maxw=48):
    for i in range(1, ncols + 1):
        L = get_column_letter(i)
        m = max(len(str(ws.cell(row=r, column=i).value or '')) for r in range(1, min(ws.max_row, 400) + 1))
        ws.column_dimensions[L].width = max(10, min(maxw, m + 2))


for ws, name, lastc, lastr in ((wsE, 'tbEntrada2008', cL, lastE), (wsS, 'tbSaidas', cSq, lastS),
                               (wsT, 'tbEstoque', get_column_letter(len(est_cols)), lastT)):
    ws.auto_filter.ref = None
    t = XTable(displayName=name, ref=f'A1:{lastc}{lastr}')
    t.tableStyleInfo = XStyle(name='TableStyleLight1', showRowStripes=False)
    ws.add_table(t)
    ncol = ws.max_column
    autofit(ws, ncol)
    for col in range(1, ncol + 1):
        h = ws.cell(row=1, column=col)
        h.font = Font(name=F, size=10, bold=True, color='FFFFFF')
        h.fill = NAVF
        for r in range(2, lastr + 1):
            c = ws.cell(row=r, column=col)
            if c.font is not None and c.font.name != F:
                c.font = Font(name=F, size=10, color=INK)
            if hasattr(c.value, 'hour') and not hasattr(c.value, 'year'):
                c.number_format = 'HH:MM:SS'
    ws.row_dimensions[1].height = 30
    ws.conditional_formatting.add(f'A2:{lastc}{lastr}', FormulaRule(formula=['MOD(ROW(),2)=0'], fill=ZEBF))
wsV.conditional_formatting.add(f'A2:O{lastV}', FormulaRule(formula=['MOD(ROW(),2)=0'], fill=ZEBF))
for t in wsV.tables.values():
    t.displayName = 'tbMovimentos'
    t.name = 'tbMovimentos'
for r in range(2, lastE + 1):
    for col in (cK, cL):
        wsE[f'{col}{r}'].font = Font(name=F, size=10, color=CORP)
for r in range(2, lastS + 1):
    wsS[f'{cSq}{r}'].font = Font(name=F, size=10, color=CORP)

# ---------- renomear abas (atualiza fórmulas, nomes, validações, gráficos, links) ----------
RENAME = {'Painel': '4 Análise ZBR', 'Depósito 2008': '4 Análise Depósito 2008', 'Cruzamento': '5 Conciliação',
          'Lançamentos': '3 Base Movimentos', 'Estoque físico': '3 Base Estoque Físico', 'Entrada': '3 Base Entrada 2008',
          'Saida': '3 Base Saídas', 'Resumo': 'Parâmetros', 'Cruzamento por Material': 'aux Conciliação Material',
          'Cruzamento por Lote': 'aux Conciliação Lote', 'Plano': '6 Plano de Ação', 'Dashboard': '1 Dashboard',
          'ResumoG': '2 Resumo Gerencial', 'Dicionario': '7 Dicionário de Dados'}
_olds = sorted(RENAME, key=len, reverse=True)


def fix(text):
    if not isinstance(text, str) or '!' not in text:
        return text
    for o in _olds:
        n = RENAME[o]
        text = text.replace(f"'{o}'!", f"'{n}'!")
        if re.fullmatch(r'[A-Za-zÀ-ú0-9_]+', o):
            text = re.sub(rf"(?<![A-Za-zÀ-ú0-9_'.]){re.escape(o)}!", f"'{n}'!", text)
    return text


for ws in wb.worksheets:
    for rw in ws.iter_rows():
        for c in rw:
            if isinstance(c.value, str) and c.value.startswith('='):
                c.value = fix(c.value)
            if c.hyperlink is not None:
                loc = c.hyperlink.location or (c.hyperlink.target or '').lstrip('#')
                if loc:
                    c.hyperlink = None
                    c.hyperlink = '#' + fix(loc)
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            rule.formula = [fix(f) for f in (rule.formula or [])]
    for dv in ws.data_validations.dataValidation:
        dv.formula1 = fix(dv.formula1)
    for ch in ws._charts:
        for sr in ch.series:
            for ref in (sr.val.numRef if sr.val else None, sr.cat.numRef if sr.cat and sr.cat.numRef else None,
                        sr.cat.strRef if sr.cat and sr.cat.strRef else None, sr.tx.strRef if sr.tx and sr.tx.strRef else None):
                if ref is not None:
                    ref.f = fix(ref.f)
for nm in list(wb.defined_names.values()):
    nm.attr_text = fix(nm.attr_text)
pa = {ws.title: ws.print_area for ws in wb.worksheets if ws.print_area}
for o, n in RENAME.items():
    wb[o].title = n
for o, area in pa.items():
    ws = wb[RENAME.get(o, o)]
    ws.print_area = area.split('!')[-1] if isinstance(area, str) else [a.split('!')[-1] for a in area][0]
wsP['A1'] = 'Análise ZBR — por SKU e lote'
wsP['A2'] = 'Entradas, saídas, estoque físico, diferença, motivo e solução por SKU e lote · filtros definidos no Dashboard'
res['B4'] = 'Parâmetros (células azul-claras, editáveis)'
res['A1'] = 'Parâmetros e conferência de totais'

# ---------- MENU ----------
wm = wb.create_sheet('MENU', 0)
wm.sheet_view.showGridLines = False
for col, w in zip('ABCDE', [3, 34, 74, 14, 3]):
    wm.column_dimensions[col].width = w
for r in range(1, 45):
    for col in range(1, 6):
        wm.cell(row=r, column=col).fill = WHF
for r in range(2, 6):
    for col in range(2, 5):
        wm.cell(row=r, column=col).fill = NAVF
cell(wm, 'B2', 'ZBR IMPORTADO', size=11, bold=True, color='9DC3E6', fill=NAVF)
cell(wm, 'B3', 'Relatório de Gestão de Estoque', size=24, bold=True, color='FFFFFF', fill=NAVF)
cell(wm, 'B4', 'Entradas, saídas, estoque físico, conciliação e plano de ação · Centro BR01', size=11, color='D9EAF7', fill=NAVF)
wm.row_dimensions[3].height = 36
INFO = [('Fonte dos dados', 'Exportação SAP "Analise_ZBR_importado.XLSX" (abas Entrada, Saida, Estoque atual)'),
        ('Período dos movimentos', '09/09/2026 a 01/10/2026'),
        ('Gerado em', "='Parâmetros'!K9"),
        ('Versão anterior (backup)', 'planilhas/backup/Analise_ZBR_Cruzamento_backup_2026-10-02.xlsx')]
for i, (a, b) in enumerate(INFO, 7):
    cell(wm, f'B{i}', a, bold=True, color=INK3)
    cell(wm, f'C{i}', b, color=INK)
wm['C9'].number_format = 'DD/MM/YYYY'
wm['C9'].alignment = Alignment(horizontal='left')
cell(wm, 'B12', 'NAVEGAÇÃO', size=13, bold=True, color=CORP)
hdr(wm, 13, 2, 'ABA')
hdr(wm, 13, 3, 'PARA QUE SERVE')
hdr(wm, 13, 4, 'ABRIR', align='center')
NAV = [('1 Dashboard', 'Visão executiva: KPIs, gráficos e filtros únicos'),
       ('2 Resumo Gerencial', 'Síntese, principais ocorrências e conclusões'),
       ('3 Base Entrada 2008', 'Movimentos do depósito 2008 (SAP)'),
       ('3 Base Saídas', 'Saídas 601 (SAP)'),
       ('3 Base Estoque Físico', 'Posições de estoque (SAP)'),
       ('3 Base Movimentos', 'Entradas e saídas consolidadas'),
       ('4 Análise ZBR', 'SKU·lote: entrou, saiu, estoque, diferença, motivo, solução'),
       ('4 Análise Depósito 2008', 'Importação recebida, transferida e parada no 2008'),
       ('5 Conciliação', 'Diferenças, motivos, categorias e criticidade'),
       ('6 Plano de Ação', 'Ocorrências, responsáveis, prazos e status'),
       ('7 Dicionário de Dados', 'Campos, regras, critérios, atualização e limitações'),
       ('Parâmetros', 'Data-limite, tolerância, limites e datas')]
for i, (sh, desc_) in enumerate(NAV, 14):
    cell(wm, f'B{i}', sh, bold=True, color=NAVY)
    cell(wm, f'C{i}', desc_, color=INK)
    c = cell(wm, f'D{i}', 'Abrir →', bold=True, color=CORP, align='center')
    c.hyperlink = f"#'{sh}'!A1"
    for col in 'BCD':
        wm[f'{col}{i}'].border = Border(bottom=Side(style='thin', color='E3E8EF'))
        if i % 2 == 0:
            wm[f'{col}{i}'].fill = ZEBF
cell(wm, 'B28', 'LEGENDA DE CORES', size=13, bold=True, color=CORP)
for i, (txt, bg, tx) in enumerate((('Regular / concluído', G_BG, G_TX), ('Atenção / em andamento', W_BG, W_TX),
                                   ('Crítico / vencido / divergente', R_BG, R_TX), ('Não iniciado / sem informação', ZEBRA, INK3),
                                   ('Célula editável (filtro ou preenchimento)', LIGHT, CORP)), 29):
    cell(wm, f'B{i}', txt, bold=True, color=tx, fill=PatternFill('solid', fgColor=bg))
cell(wm, 'B35', 'COMO USAR', size=13, bold=True, color=CORP)
for i, t in enumerate(['1. Abra o 1 Dashboard e escolha período, SKU, lote, unidade, movimentação e status nas células azul-claras.',
                       '2. Veja o detalhe por SKU·lote em 4 Análise ZBR e as causas em 5 Conciliação.',
                       '3. Registre responsável, prazo, status e evidência no 6 Plano de Ação.',
                       '4. Para atualizar a base, siga o passo a passo do 7 Dicionário de Dados.'], 36):
    cell(wm, f'B{i}', t, color=INK)
    wm.merge_cells(f'B{i}:D{i}')

# ---------- ordem, cores das abas, link de volta ao menu, impressão ----------
order = ['MENU', '1 Dashboard', '2 Resumo Gerencial', '3 Base Entrada 2008', '3 Base Saídas', '3 Base Estoque Físico',
         '3 Base Movimentos', '4 Análise ZBR', '4 Análise Depósito 2008', '5 Conciliação', '6 Plano de Ação',
         '7 Dicionário de Dados', 'Parâmetros', 'aux Conciliação Material', 'aux Conciliação Lote']
wb._sheets = [wb[n] for n in order]
TABC = {'MENU': NAVY, '1 Dashboard': NAVY, '2 Resumo Gerencial': CORP, '4 Análise ZBR': CORP, '4 Análise Depósito 2008': CORP,
        '5 Conciliação': CORP, '6 Plano de Ação': ATT_C, '7 Dicionário de Dados': DARK, 'Parâmetros': DARK}
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TABC.get(ws.title, '9DB9D6')
    if ws.title.startswith('aux'):
        ws.sheet_state = 'hidden'
    else:
        ws.sheet_state = 'visible'
MENU_LINK = {'1 Dashboard': 'P1', '2 Resumo Gerencial': 'I1', '4 Análise ZBR': 'O1', '4 Análise Depósito 2008': 'O1',
             '5 Conciliação': 'L1', '6 Plano de Ação': 'Q1', '7 Dicionário de Dados': 'D1', 'Parâmetros': 'K1'}
for sh, ref in MENU_LINK.items():
    c = cell(wb[sh], ref, '◄ MENU', size=10, bold=True, color=CORP, align='right')
    c.hyperlink = "#'MENU'!A1"
for ws in wb.worksheets:
    if ws.title.startswith('aux'):
        continue
    ws.page_setup.orientation = 'portrait' if ws.title == 'MENU' else 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6, header=0.25, footer=0.25)
    ws.oddHeader.left.text = 'ZBR Importado — Relatório de Gestão'
    ws.oddHeader.left.size = 9
    ws.oddHeader.right.text = '&A'
    ws.oddHeader.right.size = 9
    ws.oddFooter.left.text = 'Fonte: exportação SAP · Centro BR01'
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = 'Página &P de &N'
    ws.oddFooter.right.size = 8
    ws.print_options.horizontalCentered = True
for sh, rows_ in (('3 Base Entrada 2008', '1:1'), ('3 Base Saídas', '1:1'), ('3 Base Estoque Físico', '1:1'),
                  ('3 Base Movimentos', '1:1'), ('5 Conciliação', f'{HR}:{HR}'), ('6 Plano de Ação', f'{PR0}:{PR0}')):
    wb[sh].print_title_rows = rows_
wb['1 Dashboard'].print_area = 'A1:P90'
from openpyxl.worksheet.pagebreak import Break
wb['1 Dashboard'].row_breaks.append(Break(id=34))
wb['1 Dashboard'].row_breaks.append(Break(id=71))
for sh in ('4 Análise ZBR', '4 Análise Depósito 2008', '5 Conciliação', '6 Plano de Ação', '3 Base Saídas'):
    wb[sh].page_setup.paperSize = wb[sh].PAPERSIZE_A3
wb.active = 0
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = (ws.title == 'MENU')

wb.save(OUT)
print('OK', OUT, len(occ), 'ocorrências')
