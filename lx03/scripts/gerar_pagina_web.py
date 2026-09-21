"""
Gera a versão web (HTML estático) do Dashboard LX03 — uma fotografia da
última carga, com os mesmos filtros/indicadores/gráficos do Painel do
Excel, para quem só precisa abrir um link no navegador (sem instalar nada).

Chamado automaticamente por atualizar_dashboard.py a cada atualização
(desligue com --sem-html). Também pode ser rodado sozinho:

    python gerar_pagina_web.py caminho/Dashboard_Estoque_LX03.xlsx caminho/saida.html
"""

import datetime
import json
import sqlite3
import sys
from pathlib import Path

import openpyxl

TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "painel.html"

# ordem das chaves compactas usadas no JSON embutido na página (ver painel.html)
# — o mesmo formato é produzido pelo motor de alertas em JavaScript quando o
# usuário sobe um novo arquivo pelo botão de upload, então os dois lados
# (Python e JS) têm que gerar exatamente essas 11 colunas, nessa ordem.
CHAVES_JSON = ["tipo", "centro", "alerta", "um", "peso", "duracao", "dias_venc", "material", "lote", "posicao", "vencimento"]


def _num_ou_none(v):
    return round(float(v), 2) if isinstance(v, (int, float)) else None


def _data_ou_vazio(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d/%m/%Y")
    return ""


def extrair_dados(caminho_dashboard: Path, caminho_tendencia_db: Path):
    wb = openpyxl.load_workbook(caminho_dashboard, data_only=True)
    ws = wb["Dados_SAP"]
    param = wb["Parametros"]

    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    linhas, tipos, centros = [], set(), set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        tipo = r[idx["Tipo de depósito"]] or ""
        centro = r[idx["Centro"]] or ""
        um = r[idx["UM básica"]] or ""
        peso = r[idx["Estoque total"]] or 0
        if tipo:
            tipos.add(tipo)
        if centro:
            centros.add(centro)
        linhas.append([
            tipo, centro, r[idx["Alerta"]] or "", um,
            round(float(peso), 2) if isinstance(peso, (int, float)) else 0,
            _num_ou_none(r[idx["Duração (dias)"]]),
            _num_ou_none(r[idx["Dias até vencer"]]),
            r[idx["Material"]] or "", r[idx["Lote"]] or "", r[idx["Posição no depósito"]] or "",
            _data_ou_vazio(r[idx["Data do vencimento"]]),
        ])

    tendencia = []
    if caminho_tendencia_db.exists():
        conn = sqlite3.connect(caminho_tendencia_db)
        for data_ref, ocup, vaz, tkg, tun, mat_un in conn.execute(
            "SELECT data_referencia, posicoes_ocupadas, posicoes_vazias, total_kg, "
            "total_un, materiais_unicos FROM totais ORDER BY data_referencia"
        ):
            tendencia.append({"data": data_ref, "ocupadas": ocup, "vazias": vaz,
                               "total_kg": round(tkg, 1), "materiais_unicos": mat_un})
        conn.close()

    return {
        "meta": {
            "ultima_atualizacao": param["B5"].value,
            "arquivo_origem": param["B6"].value,
            "dias_prox_venc": param["B10"].value,
            "dias_parado": param["B11"].value,
        },
        "colunas": CHAVES_JSON,
        "linhas": linhas,
        "tipos": sorted(tipos),
        "centros": sorted(centros),
        "tendencia": tendencia,
    }


def gerar(caminho_dashboard: Path, caminho_tendencia_db: Path, caminho_saida: Path):
    dados = extrair_dados(caminho_dashboard, caminho_tendencia_db)
    dados_json = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</script", "<\\/script")

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    corpo = template.replace("__DADOS_JSON__", dados_json)

    # templates/painel.html é um fragmento (sem <html>/<head>/<body>) pensado
    # para ser publicado como Artifact, que injeta o <meta charset> sozinho.
    # Para o arquivo .html avulso (aberto direto no navegador, fora do
    # Artifact) isso não acontece — sem um <meta charset> logo no início do
    # arquivo, alguns navegadores adivinham a codificação errada e os
    # acentos saem corrompidos. Por isso embrulhamos aqui num documento
    # HTML completo; o navegador reposiciona sozinho <title>/<link>/<style>
    # para o <head> (regra padrão do algoritmo de parsing do HTML5).
    pagina = (
        "<!DOCTYPE html>\n"
        '<html lang="pt-BR">\n'
        "<head>\n"
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        "</head>\n<body>\n"
        + corpo +
        "\n</body>\n</html>\n"
    )

    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    caminho_saida.write_text(pagina, encoding="utf-8")
    return caminho_saida


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("Uso: python gerar_pagina_web.py <Dashboard_Estoque_LX03.xlsx> [saida.html]")
    entrada = Path(sys.argv[1])
    saida = Path(sys.argv[2]) if len(sys.argv) > 2 else entrada.parent / "Painel_LX03.html"
    tendencia_db = Path(__file__).resolve().parent.parent / "historico" / "tendencia.db"
    caminho = gerar(entrada, tendencia_db, saida)
    print(f"Página web gerada em: {caminho}")
