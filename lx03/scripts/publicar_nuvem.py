"""
Publicação do resumo do dashboard na nuvem (Google Sheets e/ou Power BI) —
para quem só precisa olhar os números pelo celular ou não tem Excel à mão.

Por desempenho e para não estourar cota de API, só publica o que é
pequeno e realmente importa para uma visão rápida: a tabela de
Pontos_de_Atencao (9 linhas) e os totais do Painel — não a base bruta de
10 mil+ linhas do Dados_SAP (essa fica só no Excel/no SAP).

Cada destino é independente e só roda se estiver configurado e ativo em
config/config.ini — sem configuração, a função avisa e não faz nada.
"""

import openpyxl

import config

NOME_ABA_DESTINO = "Pontos_de_Atencao"


def _ler_pontos_atencao(caminho_dashboard):
    wb = openpyxl.load_workbook(caminho_dashboard, data_only=True)
    ws = wb["Pontos_de_Atencao"]
    linhas = []
    cabecalho = [c.value for c in ws[4]]
    linhas.append(cabecalho)
    for row in ws.iter_rows(min_row=5, max_row=4 + 9, values_only=True):
        linhas.append(list(row))
    return linhas


def publicar_google_sheets(cfg, caminho_dashboard):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except BaseException as exc:
        # BaseException de propósito: em alguns ambientes com a lib
        # 'cryptography' do sistema em conflito, essa importação falha com
        # um erro do runtime Rust (pyo3) que não é uma Exception comum —
        # isolamos aqui para não derrubar o resto da atualização por causa
        # de uma integração opcional.
        print(f"Aviso: não consegui carregar 'gspread'/'google-auth' ({exc}) — rode "
              "'pip install gspread google-auth' num ambiente Python limpo (veja "
              "requirements.txt) para publicar no Google Sheets.")
        return

    from pathlib import Path
    caminho_credenciais = Path(cfg.get("google_sheets", "arquivo_credenciais"))
    if not caminho_credenciais.is_absolute():
        caminho_credenciais = config.BASE_DIR / caminho_credenciais
    if not caminho_credenciais.exists():
        print(f"Aviso: arquivo de credenciais do Google não encontrado em {caminho_credenciais} — "
              "veja as instruções em config/config.exemplo.ini, seção [google_sheets].")
        return

    id_planilha = cfg.get("google_sheets", "id_planilha")
    try:
        escopos = ["https://www.googleapis.com/auth/spreadsheets"]
        credenciais = Credentials.from_service_account_file(str(caminho_credenciais), scopes=escopos)
        cliente = gspread.authorize(credenciais)
        planilha = cliente.open_by_key(id_planilha)
        try:
            aba = planilha.worksheet(NOME_ABA_DESTINO)
        except gspread.WorksheetNotFound:
            aba = planilha.add_worksheet(title=NOME_ABA_DESTINO, rows=20, cols=10)
        aba.clear()
        aba.update(_ler_pontos_atencao(caminho_dashboard))
        print(f"Google Sheets atualizado: https://docs.google.com/spreadsheets/d/{id_planilha}")
    except Exception as exc:
        print(f"Aviso: falha ao publicar no Google Sheets ({exc}). Confira se a planilha foi "
              "compartilhada com o e-mail da conta de serviço (dentro do JSON de credenciais).")


def _obter_token_power_bi(cfg, url_token=None):
    import requests

    tenant_id = cfg.get("power_bi", "tenant_id")
    url = url_token or f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    dados = {
        "grant_type": "client_credentials",
        "client_id": cfg.get("power_bi", "client_id"),
        "client_secret": cfg.get("power_bi", "client_secret"),
        "scope": "https://analysis.windows.net/powerbi/api/.default",
    }
    resp = requests.post(url, data=dados, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def publicar_power_bi(cfg, caminho_dashboard, url_token=None, url_base_api=None):
    import requests

    try:
        token = _obter_token_power_bi(cfg, url_token)
    except Exception as exc:
        print(f"Aviso: falha ao autenticar no Power BI ({exc}). Confira tenant_id/client_id/"
              "client_secret em config.ini e se o App Registration tem permissão Power BI Service.")
        return

    linhas = _ler_pontos_atencao(caminho_dashboard)
    cabecalho, dados = linhas[0], linhas[1:]
    rows = [dict(zip(cabecalho, linha)) for linha in dados]

    workspace_id = cfg.get("power_bi", "workspace_id")
    dataset_id = cfg.get("power_bi", "dataset_id")
    table_name = cfg.get("power_bi", "table_name", fallback="PontosDeAtencao")
    base = url_base_api or "https://api.powerbi.com/v1.0/myorg"
    url = f"{base}/groups/{workspace_id}/datasets/{dataset_id}/tables/{table_name}/rows"

    try:
        # limpa a tabela antes de reenviar, para não acumular duplicado a cada rodada
        requests.delete(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
        resp = requests.post(url, headers={"Authorization": f"Bearer {token}"}, json={"rows": rows}, timeout=30)
        resp.raise_for_status()
        print(f"Power BI atualizado (dataset {dataset_id}, tabela {table_name}).")
    except Exception as exc:
        print(f"Aviso: falha ao publicar no Power BI ({exc}). Confira se o dataset/tabela já "
              "existem no workspace com o schema esperado (veja lx03/GUIA_POWER_BI.md).")


def publicar(caminho_dashboard):
    cfg = config.carregar()
    google_ativo = config.secao_ativa(cfg, "google_sheets", ["arquivo_credenciais", "id_planilha"])
    powerbi_ativo = config.secao_ativa(cfg, "power_bi", ["tenant_id", "client_id", "client_secret", "workspace_id", "dataset_id"])

    if not google_ativo and not powerbi_ativo:
        print("Publicação em nuvem não configurada (veja lx03/config/config.exemplo.ini, seções "
              "[google_sheets]/[power_bi], ou o caminho sem código em lx03/GUIA_POWER_BI.md).")
        return

    if google_ativo:
        publicar_google_sheets(cfg, caminho_dashboard)
    if powerbi_ativo:
        publicar_power_bi(cfg, caminho_dashboard)
