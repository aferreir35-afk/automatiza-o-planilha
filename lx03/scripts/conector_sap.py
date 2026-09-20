"""
Conexão direta ao SAP para buscar os dados da LX03 sem precisar exportar
manualmente para Excel todos os dias.

Duas formas, configuráveis em config/config.ini (veja config.exemplo.ini):

1) OData (recomendado) — um serviço OData do SAP Gateway/Fiori que devolve
   as mesmas colunas da LX03. Peça ao time de Basis/SAP para publicar (ou
   apontar um já existente) e me passar a URL + um usuário de leitura.
   O nome dos campos do serviço pode ser diferente do nosso — configure o
   de-para em [sap_odata_mapeamento] no config.ini (por padrão, assume que
   os nomes dos campos são iguais aos nossos, ver MAPEAMENTO_PADRAO).

2) RFC (alternativa) — via pacote "pyrfc" + SAP NetWeaver RFC SDK
   (instalado à parte, não incluso aqui por licenciamento da SAP). Serve
   como ponto de partida chamando RFC_READ_TABLE; para replicar a lógica
   de seleção/joins da LX03 de verdade, o ideal é o time de Basis expor
   uma RFC/BAPI própria (função Z) com a mesma saída — trocar só a função
   chamada abaixo.

Se nenhuma das duas estiver configurada como ativa, a função explica isso e
para (o comando principal cai de volta para o modo "arquivo exportado
manualmente", que é o padrão e não depende de nada disto).
"""

import datetime

import config

MATERIAL_VAZIO = "<< vazio >>"

COLUNAS_ESPERADAS = [
    "Tipo de depósito", "Posição no depósito", "Material", "Lote",
    "Nº de quantos", "Tipo de estoque", "Unidade de depósito",
    "Estoque disponível", "Último movimento", "Estoque total", "UM básica",
    "Duração", "Inventário ativo", "Depósito", "Data do vencimento", "Centro",
]

# de-para "nosso nome de coluna" -> "nome do campo no serviço OData".
# Sobrescrito por [sap_odata_mapeamento] no config.ini quando os nomes do
# serviço forem diferentes.
MAPEAMENTO_PADRAO = {c: c for c in COLUNAS_ESPERADAS}

CAMPOS_NUMERICOS = {"Nº de quantos", "Estoque disponível", "Estoque total"}
CAMPOS_DATA = {"Último movimento", "Data do vencimento"}


def _converter_data_odata(valor):
    """Aceita tanto '/Date(1699999999000)/' (OData v2 clássico do SAP
    Gateway) quanto datas ISO 8601 ('2026-09-20' ou com hora)."""
    if valor in (None, ""):
        return None
    if isinstance(valor, (int, float)):
        return datetime.datetime.utcfromtimestamp(valor / 1000)
    if isinstance(valor, str) and valor.startswith("/Date("):
        ms = int(valor.replace("/Date(", "").replace(")/", "").split("+")[0].split("-")[0])
        return datetime.datetime.utcfromtimestamp(ms / 1000)
    if isinstance(valor, str):
        try:
            return datetime.datetime.fromisoformat(valor.replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            return None
    return None


def _converter_registro(registro_bruto, mapeamento):
    linha = {}
    for nome_nosso in COLUNAS_ESPERADAS:
        campo_sap = mapeamento.get(nome_nosso, nome_nosso)
        valor = registro_bruto.get(campo_sap)
        if nome_nosso in CAMPOS_DATA:
            valor = _converter_data_odata(valor)
        elif nome_nosso in CAMPOS_NUMERICOS:
            try:
                valor = float(valor) if valor not in (None, "") else 0
            except (TypeError, ValueError):
                valor = 0
        else:
            valor = "" if valor is None else str(valor)
        linha[nome_nosso] = valor
    return linha


def _obter_mapeamento(cfg):
    mapeamento = dict(MAPEAMENTO_PADRAO)
    if cfg.has_section("sap_odata_mapeamento"):
        for nome_nosso, campo_sap in cfg.items("sap_odata_mapeamento"):
            # configparser guarda chaves em minúsculo; casamos por nome aproximado
            for c in COLUNAS_ESPERADAS:
                if c.lower().replace(" ", "_") == nome_nosso:
                    mapeamento[c] = campo_sap
    return mapeamento


def buscar_via_odata(cfg):
    import requests
    from requests.auth import HTTPBasicAuth

    url = cfg.get("sap_odata", "url_servico")
    usuario = cfg.get("sap_odata", "usuario", fallback="")
    senha = cfg.get("sap_odata", "senha", fallback="")
    verificar_ssl = cfg.getboolean("sap_odata", "verificar_ssl", fallback=True)

    headers = {"Accept": "application/json"}
    auth = HTTPBasicAuth(usuario, senha) if usuario else None
    resp = requests.get(url, headers=headers, auth=auth, verify=verificar_ssl, timeout=60)
    resp.raise_for_status()
    corpo = resp.json()

    if isinstance(corpo, dict) and "d" in corpo and "results" in corpo["d"]:
        registros = corpo["d"]["results"]           # OData v2 (SAP Gateway clássico)
    elif isinstance(corpo, dict) and "value" in corpo:
        registros = corpo["value"]                   # OData v4
    elif isinstance(corpo, list):
        registros = corpo
    else:
        raise ValueError(
            "Resposta do serviço OData em formato inesperado — esperava "
            "{'d': {'results': [...]}} (v2), {'value': [...]} (v4) ou uma lista simples."
        )

    mapeamento = _obter_mapeamento(cfg)
    return [_converter_registro(r, mapeamento) for r in registros]


def buscar_via_rfc(cfg):
    try:
        from pyrfc import Connection
    except ImportError:
        raise SystemExit(
            "ERRO: o pacote 'pyrfc' (e o SAP NetWeaver RFC SDK) não está instalado nesta "
            "máquina. Instale com 'pip install pyrfc' + o SDK da SAP (baixado do SAP "
            "Support Portal pelo time de Basis), ou use [sap_odata] em vez de [sap_rfc]."
        )

    conexao = Connection(
        ashost=cfg.get("sap_rfc", "ashost"),
        sysnr=cfg.get("sap_rfc", "sysnr", fallback="00"),
        client=cfg.get("sap_rfc", "client"),
        user=cfg.get("sap_rfc", "usuario"),
        passwd=cfg.get("sap_rfc", "senha"),
    )
    try:
        # Ponto de partida genérico via RFC_READ_TABLE (limite de ~512
        # caracteres por linha — para o layout completo da LX03, o ideal é
        # o time de Basis expor uma função Z própria com a mesma saída e
        # trocar a chamada abaixo por ela).
        resultado = conexao.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="LQUA",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": c.upper()} for c in ["LGTYP", "LGPLA", "MATNR", "CHARG", "VERME"]],
        )
        campos = [f["FIELDNAME"] for f in resultado["FIELDS"]]
        linhas = []
        for linha_rfc in resultado["DATA"]:
            valores = linha_rfc["WA"].split("|")
            registro = dict(zip(campos, valores))
            linhas.append({
                "Tipo de depósito": registro.get("LGTYP", ""),
                "Posição no depósito": registro.get("LGPLA", ""),
                "Material": registro.get("MATNR", "") or MATERIAL_VAZIO,
                "Lote": registro.get("CHARG", ""),
                "Nº de quantos": 0, "Tipo de estoque": "", "Unidade de depósito": "",
                "Estoque disponível": float(registro.get("VERME", 0) or 0),
                "Último movimento": None,
                "Estoque total": float(registro.get("VERME", 0) or 0),
                "UM básica": "", "Duração": "", "Inventário ativo": "",
                "Depósito": "", "Data do vencimento": None, "Centro": "",
            })
        return linhas
    finally:
        conexao.close()


def buscar_dados():
    cfg = config.carregar()
    odata_ativo = config.secao_ativa(cfg, "sap_odata", ["url_servico"])
    rfc_ativo = config.secao_ativa(cfg, "sap_rfc", ["ashost", "client", "usuario", "senha"])

    if odata_ativo:
        return buscar_via_odata(cfg)
    if rfc_ativo:
        return buscar_via_rfc(cfg)

    raise SystemExit(
        "ERRO: --usar-sap foi pedido, mas nem [sap_odata] nem [sap_rfc] estão configurados "
        "e ativos em lx03/config/config.ini. Preencha um dos dois (veja config.exemplo.ini) "
        "ou rode sem --usar-sap, passando o arquivo exportado manualmente da LX03."
    )
