"""
Alertas por e-mail e Microsoft Teams para o Dashboard LX03.

Compara o retrato de hoje com o do dia anterior (guardado por tendencia.py
em lx03/historico/tendencia.db) e manda um resumo por e-mail e/ou Teams
quando algo piorou — ou sempre, se preferir (veja config/config.ini).

Não faz nada (e avisa por que) se a seção correspondente não estiver
"ativo = true" e com os campos obrigatórios preenchidos em config.ini.
"""

import smtplib
import ssl
from email.mime.text import MIMEText

import config
import tendencia

ALERTAS_RISCO = [
    "VENCIDO", "PRÓX. VENCIMENTO", "QUALIDADE", "BLOQUEADO",
    "RESTRITO/DEVOLUÇÃO", "QUARENTENA", "DESCARTE/DEVOLUÇÃO",
    "SEM LOTE", "PARADO/SEM GIRO",
]


def _piorou(contagem, kg, contagem_ant, kg_ant):
    linhas_piora = []
    for cat in ALERTAS_RISCO:
        atual_qtd, atual_kg = contagem.get(cat, 0), kg.get(cat, 0.0)
        antes_qtd, antes_kg = contagem_ant.get(cat, 0), kg_ant.get(cat, 0.0)
        if atual_qtd > antes_qtd or atual_kg > antes_kg + 0.01:
            linhas_piora.append((cat, antes_qtd, atual_qtd, antes_kg, atual_kg))
    return linhas_piora


def _montar_mensagem(contagem, kg, arquivo_origem, ultima_atualizacao, linhas_piora, tem_historico):
    partes = [
        f"Dashboard de Estoque LX03 — atualização de {ultima_atualizacao}",
        f"Arquivo de origem: {arquivo_origem}",
        "",
    ]
    if linhas_piora:
        partes.append("Pontos que PIORARAM frente à última carga:")
        for cat, aq, nq, ak, nk in linhas_piora:
            partes.append(f"  - {cat}: {aq} -> {nq} posições | {ak:,.0f} -> {nk:,.0f} kg".replace(",", "."))
        partes.append("")
    elif tem_historico:
        partes.append("Nenhum ponto de atenção piorou desde a última carga.")
        partes.append("")
    partes.append("Situação atual:")
    for cat in ALERTAS_RISCO:
        if contagem.get(cat, 0) > 0:
            partes.append(f"  - {cat}: {contagem[cat]} posições | {kg.get(cat, 0):,.0f} kg".replace(",", "."))
    return "\n".join(partes)


def enviar_email(cfg, assunto, corpo):
    destinatarios = [e.strip() for e in cfg.get("email", "destinatarios", fallback="").split(",") if e.strip()]
    if not destinatarios:
        print("Aviso: seção [email] ativa mas 'destinatarios' está vazio — nada enviado.")
        return
    remetente = cfg.get("email", "remetente", fallback="") or cfg.get("email", "usuario")
    msg = MIMEText(corpo, "plain", "utf-8")
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)

    servidor = cfg.get("email", "servidor_smtp")
    porta = cfg.getint("email", "porta", fallback=587)
    usuario = cfg.get("email", "usuario")
    senha = cfg.get("email", "senha")

    try:
        contexto = ssl.create_default_context()
        with smtplib.SMTP(servidor, porta, timeout=20) as smtp:
            smtp.starttls(context=contexto)
            smtp.login(usuario, senha)
            smtp.sendmail(remetente, destinatarios, msg.as_string())
        print(f"E-mail de alerta enviado para: {', '.join(destinatarios)}")
    except Exception as exc:
        print(f"Aviso: falha ao enviar e-mail de alerta ({exc}). Confira config/config.ini.")


def enviar_teams(cfg, titulo, corpo):
    import requests

    webhook_url = cfg.get("teams", "webhook_url")
    payload = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "summary": titulo,
        "themeColor": "6E1423",
        "title": titulo,
        "text": corpo.replace("\n", "  \n"),
    }
    try:
        resp = requests.post(webhook_url, json=payload, timeout=20)
        if resp.status_code >= 300:
            print(f"Aviso: Teams respondeu {resp.status_code} ao tentar enviar o alerta.")
        else:
            print("Alerta enviado ao Teams.")
    except Exception as exc:
        print(f"Aviso: falha ao enviar alerta ao Teams ({exc}). Confira config/config.ini.")


def avaliar_e_notificar(data_referencia, contagem, kg, arquivo_origem, ultima_atualizacao):
    cfg = config.carregar()
    email_ativo = config.secao_ativa(cfg, "email", ["servidor_smtp", "usuario", "senha", "destinatarios"])
    teams_ativo = config.secao_ativa(cfg, "teams", ["webhook_url"])

    if not email_ativo and not teams_ativo:
        print("Notificações não configuradas (veja lx03/config/config.exemplo.ini, seções [email] e [teams]) — nada enviado.")
        return

    conn = tendencia.conectar()
    data_anterior, snapshot_anterior = tendencia.obter_snapshot_anterior(conn, data_referencia)
    conn.close()

    tem_historico = snapshot_anterior is not None
    linhas_piora = []
    if tem_historico:
        contagem_ant, kg_ant = snapshot_anterior
        linhas_piora = _piorou(contagem, kg, contagem_ant, kg_ant)

    sempre_se_vencido = cfg.getboolean("alertas", "sempre_alertar_se_vencido", fallback=True)
    somente_piora = cfg.getboolean("alertas", "somente_quando_piora", fallback=True)

    deve_notificar = (
        (sempre_se_vencido and contagem.get("VENCIDO", 0) > 0)
        or (not somente_piora)
        or (somente_piora and linhas_piora)
    )
    if not deve_notificar:
        print("Nenhum ponto de atenção piorou desde a última carga — notificação não enviada (veja [alertas] em config.ini).")
        return

    corpo = _montar_mensagem(contagem, kg, arquivo_origem, ultima_atualizacao, linhas_piora, tem_historico)
    assunto = "⚠️ Dashboard LX03 — pontos de atenção" + (" (piora detectada)" if linhas_piora else "")

    if email_ativo:
        enviar_email(cfg, assunto, corpo)
    if teams_ativo:
        enviar_teams(cfg, assunto, corpo)
