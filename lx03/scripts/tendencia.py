"""
Histórico de tendência do Dashboard LX03.

Guarda, a cada rodada do script, um "retrato" do dia (quantidade e peso por
categoria de alerta) num banco SQLite simples (lx03/historico/tendencia.db)
e monta a partir dele a aba "Tendencia" do dashboard, com um gráfico de
evolução diária — assim a diretoria enxerga se cada ponto de atenção está
piorando ou melhorando ao longo do tempo, não só a fotografia de hoje.

Rodar o script de atualização mais de uma vez no mesmo dia substitui o
retrato daquele dia (não duplica linhas) — a tendência é por dia, não por
execução.
"""

import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "historico" / "tendencia.db"

CATEGORIAS_TENDENCIA = [
    "VENCIDO", "PRÓX. VENCIMENTO", "QUALIDADE", "BLOQUEADO",
    "RESTRITO/DEVOLUÇÃO", "QUARENTENA", "DESCARTE/DEVOLUÇÃO",
    "SEM LOTE", "PARADO/SEM GIRO",
]


def conectar():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS totais (
            data_referencia TEXT PRIMARY KEY,
            data_hora_atualizacao TEXT,
            arquivo_origem TEXT,
            posicoes_ocupadas INTEGER,
            posicoes_vazias INTEGER,
            total_kg REAL,
            total_un REAL,
            materiais_unicos INTEGER
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS alertas (
            data_referencia TEXT,
            categoria TEXT,
            qtd INTEGER,
            kg REAL,
            PRIMARY KEY (data_referencia, categoria)
        )
    """)
    return conn


def registrar_snapshot(conn, data_referencia, data_hora_atualizacao, arquivo_origem,
                        contagem, kg, posicoes_ocupadas, posicoes_vazias,
                        total_kg, total_un, materiais_unicos):
    conn.execute(
        "INSERT OR REPLACE INTO totais VALUES (?,?,?,?,?,?,?,?)",
        (data_referencia, data_hora_atualizacao, arquivo_origem,
         posicoes_ocupadas, posicoes_vazias, total_kg, total_un, materiais_unicos),
    )
    for categoria in CATEGORIAS_TENDENCIA:
        conn.execute(
            "INSERT OR REPLACE INTO alertas VALUES (?,?,?,?)",
            (data_referencia, categoria, contagem.get(categoria, 0), kg.get(categoria, 0.0)),
        )
    conn.commit()


def obter_snapshot_anterior(conn, data_referencia_atual):
    """Retorna (contagem, kg) do dia mais recente ANTES de data_referencia_atual,
    ou (None, None) se não houver histórico anterior (primeira carga)."""
    row = conn.execute(
        "SELECT data_referencia FROM totais WHERE data_referencia < ? ORDER BY data_referencia DESC LIMIT 1",
        (data_referencia_atual,),
    ).fetchone()
    if not row:
        return None, None
    data_anterior = row[0]
    contagem, kg = {}, {}
    for categoria, qtd, peso in conn.execute(
        "SELECT categoria, qtd, kg FROM alertas WHERE data_referencia = ?", (data_anterior,)
    ):
        contagem[categoria] = qtd
        kg[categoria] = peso
    return data_anterior, (contagem, kg)


def obter_historico_completo(conn):
    """Lista de datas (ordem crescente) com dict {categoria: (qtd, kg)} e os totais do dia."""
    datas = [r[0] for r in conn.execute("SELECT data_referencia FROM totais ORDER BY data_referencia")]
    historico = []
    for data in datas:
        totais = conn.execute(
            "SELECT posicoes_ocupadas, posicoes_vazias, total_kg, total_un, materiais_unicos "
            "FROM totais WHERE data_referencia = ?", (data,)
        ).fetchone()
        alertas = {c: (q, k) for c, q, k in conn.execute(
            "SELECT categoria, qtd, kg FROM alertas WHERE data_referencia = ?", (data,)
        )}
        historico.append({"data": data, "totais": totais, "alertas": alertas})
    return historico


def construir_aba_tendencia(wb, conn, estilos):
    """Cria a aba 'Tendencia' com o histórico diário + gráfico de evolução."""
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    historico = obter_historico_completo(conn)

    ws = wb.create_sheet("Tendencia")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Tendência — Evolução Diária dos Pontos de Atenção"
    ws["A1"].font = estilos["titulo"]
    ws.merge_cells("A1:L1")
    ws["A2"] = (
        "Um retrato por dia, guardado automaticamente a cada atualização (lx03/historico/tendencia.db). "
        "Começa com poucos pontos e cresce a cada dia que o script rodar — não precisa preencher nada aqui."
    )
    ws["A2"].font = estilos["subtitulo"]
    ws.merge_cells("A2:L2")

    cab = ["Data"] + [f"{c} (kg)" for c in CATEGORIAS_TENDENCIA] + ["Total KG", "Posições Ocupadas"]
    for c, h in enumerate(cab, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = estilos["cab_fonte"]
        cell.fill = estilos["cab_fill"]
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for i, ponto in enumerate(historico):
        r = 5 + i
        ws.cell(row=r, column=1, value=ponto["data"])
        for j, categoria in enumerate(CATEGORIAS_TENDENCIA):
            _, kgv = ponto["alertas"].get(categoria, (0, 0.0))
            ws.cell(row=r, column=2 + j, value=round(kgv, 1))
        totais = ponto["totais"] or (0, 0, 0.0, 0.0, 0)
        ws.cell(row=r, column=2 + len(CATEGORIAS_TENDENCIA), value=round(totais[2], 1))
        ws.cell(row=r, column=3 + len(CATEGORIAS_TENDENCIA), value=totais[0])
    ultima_linha = 4 + max(len(historico), 1)

    for c in range(1, len(cab) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14

    if len(historico) >= 1:
        chart = LineChart()
        chart.title = "Evolução do estoque em risco (kg) — Vencido, Qualidade, Bloqueado, Parado/Sem Giro"
        chart.y_axis.title = "kg"
        chart.x_axis.title = "Data"
        chart.height = 10
        chart.width = 26
        colunas_chave = ["VENCIDO", "QUALIDADE", "BLOQUEADO", "PARADO/SEM GIRO"]
        for categoria in colunas_chave:
            idx_col = 2 + CATEGORIAS_TENDENCIA.index(categoria)
            dados = Reference(ws, min_col=idx_col, min_row=4, max_row=ultima_linha)
            chart.add_data(dados, titles_from_data=True)
        cats = Reference(ws, min_col=1, min_row=5, max_row=ultima_linha)
        chart.set_categories(cats)
        ws.add_chart(chart, f"A{ultima_linha + 3}")
        if len(historico) == 1:
            ws.cell(row=ultima_linha + 2, column=1,
                    value="Só há 1 dia de histórico até agora — o gráfico vira uma linha de verdade a partir da 2ª atualização.").font = estilos["nota"]

    ws.print_area = f"A1:L{ultima_linha + 25}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return ws
