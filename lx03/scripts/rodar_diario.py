#!/usr/bin/env python3
"""
Roda a atualização do dashboard usando o arquivo mais recente salvo em
lx03/entrada/ — é isto que o agendamento diário (Task Scheduler/cron)
chama; não precisa saber o nome do arquivo de hoje.
"""

import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
ENTRADA_DIR = BASE_DIR / "entrada"
SCRIPT_ATUALIZAR = Path(__file__).resolve().parent / "atualizar_dashboard.py"


def arquivo_mais_recente():
    arquivos = [p for p in ENTRADA_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
    if not arquivos:
        return None
    return max(arquivos, key=lambda p: p.stat().st_mtime)


def main():
    arq = arquivo_mais_recente()
    if not arq:
        print(f"Nenhum arquivo .xlsx encontrado em {ENTRADA_DIR} — coloque o export do dia lá antes do horário agendado.")
        sys.exit(1)
    print(f"Usando o arquivo mais recente de entrada/: {arq.name}")
    resultado = subprocess.run([sys.executable, str(SCRIPT_ATUALIZAR), str(arq)])
    sys.exit(resultado.returncode)


if __name__ == "__main__":
    main()
