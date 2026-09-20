#!/usr/bin/env python3
"""
Configura o agendamento diário do dashboard LX03 no sistema operacional —
rode este script UMA VEZ (por máquina/usuário) para não precisar mais
lembrar de rodar a atualização manualmente todos os dias.

Windows (Task Scheduler):
    python agendar_tarefa.py --hora 07:00

Linux/Mac (crontab):
    python agendar_tarefa.py --hora 07:00

Remover o agendamento:
    python agendar_tarefa.py --remover

O agendamento sempre chama "rodar_diario.py", que por sua vez pega
automaticamente o arquivo mais recente de lx03/entrada/ — então o único
passo manual que sobra no dia a dia é salvar o export da LX03 nessa pasta
antes do horário agendado (ou usar --usar-sap no dia a dia se a conexão
direta ao SAP estiver configurada, veja conector_sap.py).
"""

import argparse
import platform
import subprocess
import sys
from pathlib import Path

NOME_TAREFA = "Dashboard_LX03_Diario"
SCRIPT_DIARIO = Path(__file__).resolve().parent / "rodar_diario.py"
LOG_PATH = Path(__file__).resolve().parent.parent / "historico" / "agendamento.log"


def _validar_hora(hora_str):
    partes = hora_str.split(":")
    if len(partes) != 2 or not all(p.isdigit() for p in partes):
        raise SystemExit("ERRO: use o formato HH:MM, ex.: --hora 07:00")
    h, m = int(partes[0]), int(partes[1])
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise SystemExit("ERRO: horário inválido.")
    return h, m


def agendar_windows(hora_str):
    h, m = _validar_hora(hora_str)
    comando = (
        f'schtasks /Create /TN "{NOME_TAREFA}" '
        f'/TR "\\"{sys.executable}\\" \\"{SCRIPT_DIARIO}\\"" '
        f'/SC DAILY /ST {h:02d}:{m:02d} /F'
    )
    print("Executando:", comando)
    resultado = subprocess.run(comando, shell=True)
    if resultado.returncode == 0:
        print(f"Tarefa '{NOME_TAREFA}' agendada no Task Scheduler do Windows, todos os dias às {h:02d}:{m:02d}.")
        print("Para conferir/editar: abra o 'Agendador de Tarefas' do Windows e procure por esse nome.")
    else:
        print("Não consegui criar a tarefa automaticamente. Comando para rodar manualmente (Prompt de Comando como Administrador):")
        print(comando)


def remover_windows():
    subprocess.run(f'schtasks /Delete /TN "{NOME_TAREFA}" /F', shell=True)
    print(f"Tarefa '{NOME_TAREFA}' removida do Task Scheduler (se existia).")


def _tem_crontab():
    try:
        subprocess.run(["crontab", "-l"], capture_output=True, timeout=5)
        return True
    except FileNotFoundError:
        return False


def agendar_linux_mac(hora_str):
    h, m = _validar_hora(hora_str)
    linha = f"{m} {h} * * * {sys.executable} {SCRIPT_DIARIO} >> {LOG_PATH} 2>&1  # {NOME_TAREFA}"
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not _tem_crontab():
        print("O comando 'crontab' não está disponível neste sistema. Adicione manualmente esta linha "
              "ao seu agendador (crontab, systemd timer, launchd, etc.):\n")
        print(f"  {linha}\n")
        return

    atual = subprocess.run(["crontab", "-l"], capture_output=True, text=True)
    linhas_atuais = atual.stdout.splitlines() if atual.returncode == 0 else []
    linhas_atuais = [l for l in linhas_atuais if NOME_TAREFA not in l]
    linhas_atuais.append(linha)
    novo_crontab = "\n".join(linhas_atuais) + "\n"
    resultado = subprocess.run(["crontab", "-"], input=novo_crontab, text=True)
    if resultado.returncode == 0:
        print(f"Crontab atualizado — atualização diária às {h:02d}:{m:02d}. Log em: {LOG_PATH}")
    else:
        print("Não consegui atualizar o crontab automaticamente. Adicione manualmente:\n")
        print(f"  {linha}\n")


def remover_linux_mac():
    if not _tem_crontab():
        print(f"'crontab' não disponível — remova manualmente a linha com o comentário # {NOME_TAREFA}, se você a adicionou à mão.")
        return
    atual = subprocess.run(["crontab", "-l"], capture_output=True, text=True)
    if atual.returncode != 0:
        print("Nenhum crontab existente.")
        return
    linhas_atuais = [l for l in atual.stdout.splitlines() if NOME_TAREFA not in l]
    subprocess.run(["crontab", "-"], input="\n".join(linhas_atuais) + "\n", text=True)
    print(f"Entrada de agendamento '{NOME_TAREFA}' removida do crontab (se existia).")


def main():
    ap = argparse.ArgumentParser(description="Agenda a atualização diária automática do Dashboard LX03.")
    ap.add_argument("--hora", default="07:00", help="Horário diário no formato HH:MM (padrão: 07:00)")
    ap.add_argument("--remover", action="store_true", help="Remove o agendamento em vez de criar")
    args = ap.parse_args()

    sistema = platform.system()
    if sistema == "Windows":
        remover_windows() if args.remover else agendar_windows(args.hora)
    elif sistema in ("Linux", "Darwin"):
        remover_linux_mac() if args.remover else agendar_linux_mac(args.hora)
    else:
        print(f"Sistema '{sistema}' não reconhecido — agende manualmente a execução diária de:\n  {sys.executable} {SCRIPT_DIARIO}")


if __name__ == "__main__":
    main()
