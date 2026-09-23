"""Migra o CSV original "TROCA DE TURNO 2026 (VISÃO)" para registros padronizados.

Uso:  python scripts/migrar_csv.py  (gera dados/base_migrada.json)

Correções aplicadas (todas listadas em dados/correcoes_migracao.csv):
  * datas com dia/mês invertidos (ex.: 08/01/2026 entre 01/08 e 02/08) e ano com 2 dígitos;
  * números em formatos mistos (43127 / 28.054,00 / 24,039) -> toneladas;
  * turno em minúsculo, programação/medida/transportadora fora do padrão;
  * linhas totalmente vazias descartadas;
  * pendências identificadas por palavras-chave da observação original
    (AGUARDANDO, FALTA, PENDENTE ...), sem inventar responsável ou prazo.
"""
import csv
import datetime as dt
import json
import re
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "dados" / "TROCA_DE_TURNO_2026_original.csv"
OUT = ROOT / "dados" / "base_migrada.json"

PROG_MAP = {
    "GM": "GM", "SP": "GM", "SP CAP/INT": "GM", "AGENDA SP": "GM", "AGENDA": "GM",
    "SMARTLOG": "GM", "SMT": "GM", "SMT001CA": "GM", "SMT002CA": "GM", "SMT003CA": "GM",
    "SMT004CA": "GM", "SMT001MACPAN": "GM", "SUL": "GM",
    "FM": "FM", "FOB": "FOB", "CIF": "CIF", "GRANEL": "GRANEL",
    "CACAU": "CACAU", "SEP CACAU": "CACAU", "SEP. CACAU": "CACAU", "SEPARAÇÃO CACAU": "CACAU",
    "SEP TRANSF CACAU": "CACAU", "CACAU BR15": "CACAU",
    "BR15": "BR15", "SEP. BR15": "BR15", "SEP BR 15": "BR15",
    "REC. BR15": "RECEBIMENTO", "REC CACAU": "RECEBIMENTO", "RECEB CACAU": "RECEBIMENTO",
    "RECEBIMENTO CACAU": "RECEBIMENTO", "REC. MACEIO": "RECEBIMENTO", "REC. IMPORTADO": "RECEBIMENTO",
    "REC. D2": "RECEBIMENTO", "RECEBIMENTO": "RECEBIMENTO", "RETORNO D2": "RECEBIMENTO",
    "IMPORTAÇÃO": "RECEBIMENTO", "MACEIO": "RECEBIMENTO",
    "TRANSF. MACEIO": "TRANSFERÊNCIA", "TRANSF MACEIO": "TRANSFERÊNCIA", "TRANSF": "TRANSFERÊNCIA",
    "TRANSFERENCIA": "TRANSFERÊNCIA", "TRANSF OMEGA X": "TRANSFERÊNCIA",
    "EXPORTAÇÃO": "EXPORTAÇÃO",
    "FLOOR": "FLOOR", "FLOOR / PK1": "FLOOR", "FLOOR / CACAU": "FLOOR", "FLOOR/ CACAU": "FLOOR",
    "FLOOR / RECEB. CACAU": "FLOOR",
    "ABASTECIMENTO": "ABASTECIMENTO", "ABASTECIMENTO PK1": "ABASTECIMENTO",
    "CONT. PK1": "INVENTÁRIO", "CONTAGENS PK1": "INVENTÁRIO", "INVENTÁRIO": "INVENTÁRIO",
    "MOVIMENTAÇÃO": "MOVIMENTAÇÃO", "PRODUÇÃO": "MOVIMENTAÇÃO", "DESCARTE": "DESCARTE",
    "CROSS RJ": "CROSS DOCKING", "CROSS LOGGIO": "CROSS DOCKING",
}
ATIV_MAP = {
    "GM": "EXPEDIÇÃO", "FM": "EXPEDIÇÃO", "FOB": "EXPEDIÇÃO", "CIF": "EXPEDIÇÃO",
    "GRANEL": "EXPEDIÇÃO", "CACAU": "EXPEDIÇÃO", "BR15": "EXPEDIÇÃO", "EXPORTAÇÃO": "EXPEDIÇÃO",
    "CROSS DOCKING": "EXPEDIÇÃO", "RECEBIMENTO": "RECEBIMENTO", "TRANSFERÊNCIA": "TRANSFERÊNCIA",
    "FLOOR": "ARMAZENAGEM", "ABASTECIMENTO": "ABASTECIMENTO", "INVENTÁRIO": "INVENTÁRIO",
    "MOVIMENTAÇÃO": "MOVIMENTAÇÃO", "DESCARTE": "MOVIMENTAÇÃO",
}
TRANSP_MAP = {
    "EPEEDYLOG": "SPEEDYLOG", "SPEEDY": "SPEEDYLOG", "SPEEDYLOG": "SPEEDYLOG",
    "SMARTLOG": "SMARTLOG", "SMART": "SMARTLOG", "SMART LOG": "SMARTLOG", "SMARLOG": "SMARTLOG",
    "PROPRIO": "PRÓPRIO", "PRÓPRIO": "PRÓPRIO", "OMEGA X": "OMEGA X", "OMX": "OMEGA X",
    "OMEGAX": "OMEGA X", "LOGUIO": "LOGGIO", "LOGGIO": "LOGGIO",
}
MEDIDA_MAP = {"TONS": "TON", "LIQUIDO": "TON", "PAL": "PALETE", "POSIÇÃO": "POSIÇÃO"}
FAT_MAP = {"REALIZADO": "FATURADO", "FATURADO": "FATURADO", "PENDENTE": "PENDENTE",
           "PARCIAL": "EM PROCESSO", "PARCIALMENTE": "EM PROCESSO"}
EMB_MAP = {"REALIZADO": "REALIZADO", "PENDENTE": "PENDENTE", "PARCIAL": "EM PROCESSO"}

# Palavras que indicam que a observação descreve algo ainda não resolvido.
PEND_MARK = re.compile(r"AGUARD|\bAG\.|FALTA|FATA |PENDEN|FINALIZAR|\bFATURAR\b|\bSEPARAR|CONFERIR|"
                       r"CARREGAR|DAR ENTRADA|DEPENDE|DO ZERO|EM CONFERENCIA|EM CARREGAMENTO|"
                       r"NO PATIO|NA PORTARIA|NA DOCA|- PORTARIA|PESAR")
TIPO_RULES = [  # (regex, tipo de pendência, causa ou None)
    (r"SISTEMA|RUBY|APONTAMENTO", "SISTEMA", "SISTEMA"),
    (r"LOTES TRATADOS|AVARIA|QUALIDADE", "QUALIDADE", None),
    (r"MAQUINA|MÁQUINA|EMPILHADEIRA|EQUIPAMENTO", "EQUIPAMENTO", "EQUIPAMENTO"),
    (r"CREDITO|CRÉDITO|FINANCEIRO", "FATURAMENTO", "CLIENTE"),
    (r"FATURAR|FATURAMENTO", "FATURAMENTO", None),
    (r"ENTRADA|\bNF\b|NOTAS|LACRES", "DOCUMENTAÇÃO", None),
    (r"CONFER", "CONFERÊNCIA", None),
    (r"SEPARA|SEPAR|DEPARAR|\bOTS?\b|DELIVERY|DO ZERO", "SEPARAÇÃO", None),
    (r"COLETA|CARRO|MOTORISTA|PORTARIA|PATIO|DOCA", "EMBARQUE", "TRANSPORTE"),
    (r"CARREG|EMBARQUE|PESA", "EMBARQUE", None),
]


def norm(s):
    return re.sub(r"\s+", " ", (s or "").strip()).upper()


def num(s):
    s = (s or "").strip().replace(" ", "")
    if not s:
        return None, None
    note = None
    if "," in s and "." in s:
        v = s.replace(".", "").replace(",", ".")
    elif "," in s:
        v = s.replace(",", ".")
    elif s.count(".") > 1:  # 9.012.686 -> 9012,686 (último ponto como decimal)
        parts = s.split(".")
        v = "".join(parts[:-1]) + "." + parts[-1]
        note = f"número '{s}' interpretado como {v.replace('.', ',')}"
    elif "." in s and len(s.split(".")[1]) == 3:
        v = s.replace(".", "")
    else:
        v = s
    return float(v), note


def to_ton(v):
    """Valores >= 100 estavam em kg; valores < 100 já estavam em toneladas."""
    if v is None:
        return None
    return round(v / 1000.0, 3) if abs(v) >= 100 else round(v, 3)


def parse_date(s):
    s = s.strip()
    if not s:
        return None
    d, m, y = s.split("/")
    y = int(y)
    if y < 100:
        y += 2000
    return dt.date(y, int(m), int(d))


def main():
    raw = SRC.read_bytes().decode("latin-1")
    rows = list(csv.reader(raw.splitlines(), delimiter=";"))
    data = rows[7:]
    log = [(0, "REGRA", "valores em kg (>= 100) na medida TONS",
            "convertidos para toneladas (÷1000); valores < 100 mantidos (já estavam em t)"),
           (0, "REGRA", "medida TONS / LIQUIDO / PAL", "padronizada para TON / TON / PALETE")]  # (linha csv, campo, de, para)
    recs = []

    dates = [parse_date(r[0]) if r and r[0].strip() else None for r in data]
    # datas com dia/mês invertidos: se a data trocada cabe entre as vizinhas, corrige
    fixed = list(dates)
    for i, d in enumerate(dates):
        if d is None or d.day > 12:
            continue
        prev = next((x for x in reversed(fixed[:i]) if x), None)
        if prev is None:
            continue
        sw = dt.date(d.year, d.day, d.month)
        out = abs((d - prev).days) > 20
        fits = (prev - dt.timedelta(days=3)) <= sw <= (prev + dt.timedelta(days=10))
        if out and fits:
            fixed[i] = sw

    for i, r in enumerate(data):
        line = i + 8
        r = (r + [""] * 15)[:15]
        body = [c.strip() for c in r[3:15]]
        if not any(body):
            if any(c.strip() for c in r):
                log.append((line, "LINHA", ";".join(r).strip(";"), "descartada (sem conteúdo)"))
            continue
        notes = []
        d = fixed[i]
        if dates[i] != d:
            log.append((line, "DATA", r[0], d.strftime("%d/%m/%Y")))
            notes.append(f"data corrigida de {r[0]}")
        elif r[0].strip() and len(r[0].strip().split("/")[2]) == 2:
            log.append((line, "DATA", r[0], d.strftime("%d/%m/%Y")))

        turno = norm(r[1])
        if r[1].strip() and r[1].strip() != turno:
            log.append((line, "TURNO", r[1], turno))
        resp = norm(r[2])

        prog_o = norm(r[3])
        prog = PROG_MAP.get(prog_o, prog_o) if prog_o else ""
        rota = norm(r[4])
        if not prog and rota in ("CACAU",):
            prog = "CACAU"
        if prog != prog_o:
            log.append((line, "PROGRAMAÇÃO", r[3], prog or "(vazio)"))
            if prog_o and prog_o.replace(".", "") != prog:
                notes.append(f"programação original: {prog_o}")

        med_o = norm(r[5])
        med = MEDIDA_MAP.get(med_o, med_o)
        pl, n1 = num(r[6])
        ex, n2 = num(r[7])
        vf, n3 = num(r[10])
        for n in (n1, n2, n3):
            if n:
                notes.append(n)
                log.append((line, "NÚMERO", "", n))
        if med == "POSIÇÃO" and prog in ("GM", "FM", "FOB") and (pl or 0) > 1000:
            med = "TON"
            notes.append("medida original POSIÇÃO corrigida para TON")
        if not med and pl is not None:
            med = "PALETE" if prog in ("FLOOR", "ABASTECIMENTO") else "TON"
            notes.append("medida inferida")
        if med != med_o and not (med_o, med) in (("TONS", "TON"), ("PAL", "PALETE")):
            log.append((line, "MEDIDA", r[5], med))
        if med == "TON":
            pl2, ex2, vf2 = to_ton(pl), to_ton(ex), to_ton(vf)
            pl, ex, vf = pl2, ex2, vf2

        fat_o, emb_o = norm(r[12]), norm(r[13])
        fat = FAT_MAP.get(fat_o, fat_o)
        emb = EMB_MAP.get(emb_o, emb_o)
        tr_o = norm(r[11])
        tr = TRANSP_MAP.get(tr_o, tr_o)
        if tr != tr_o:
            log.append((line, "TRANSPORTADORA", r[11], tr))

        obs = re.sub(r"\s+", " ", r[14].strip())
        obs_u = unicodedata.normalize("NFC", obs.upper())
        pend = tipo = causa = ""
        if obs and PEND_MARK.search(obs_u):
            pend = obs
            tipo = "OUTROS"
            for rx, tp, cs in TIPO_RULES:
                seg = obs_u
                if re.search(rx, seg):
                    tipo, causa = tp, cs or ""
                    break
        if notes:
            obs = (obs + " " if obs else "") + "[MIGRAÇÃO: " + "; ".join(notes) + "]"

        recs.append({
            "linha_csv": line,
            "data": d.isoformat() if d else None,
            "turno": turno, "resp": resp,
            "tipo_ativ": ATIV_MAP.get(prog, "OUTROS") if prog else "",
            "prog": prog, "rota": rota, "medida": med,
            "plan": pl, "exec": ex,
            "fat": fat, "emb": emb, "vol_fat": vf, "transp": tr,
            "pend": pend, "tipo_pend": tipo, "causa": causa,
            "obs": obs,
        })

    OUT.write_text(json.dumps({"registros": recs, "log": log}, ensure_ascii=False, indent=0))
    line2id = {rec["linha_csv"]: f"TT-{i + 1:05d}" for i, rec in enumerate(recs)}
    with open(ROOT / "dados" / "correcoes_migracao.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["LINHA DO CSV ORIGINAL", "ID NA PLANILHA", "CAMPO", "VALOR ORIGINAL", "VALOR CORRIGIDO"])
        for line, campo, de, para in log:
            w.writerow([line or "", line2id.get(line, ""), campo, de, para])
    print(f"{len(recs)} registros migrados, {len(log)} correções registradas -> {OUT}")


if __name__ == "__main__":
    main()
