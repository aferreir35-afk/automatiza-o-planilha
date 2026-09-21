"""
Carregador de configuração das integrações (e-mail, Teams, SAP, nuvem).

Lê lx03/config/config.ini (se existir) por cima dos valores padrão de
config.exemplo.ini, para que um "config.ini" real com credenciais nunca
precise ser criado do zero — só copiado e preenchido.
"""

import configparser
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = BASE_DIR / "config"
CONFIG_EXEMPLO = CONFIG_DIR / "config.exemplo.ini"
CONFIG_REAL = CONFIG_DIR / "config.ini"


def carregar():
    cfg = configparser.ConfigParser()
    if CONFIG_EXEMPLO.exists():
        cfg.read(CONFIG_EXEMPLO, encoding="utf-8")
    if CONFIG_REAL.exists():
        cfg.read(CONFIG_REAL, encoding="utf-8")
    return cfg


def secao_ativa(cfg, secao, campos_obrigatorios):
    """True só se a seção existir, 'ativo' for true, e os campos
    obrigatórios estiverem preenchidos (evita 'ativo=true' esquecido com
    credenciais vazias tentando enviar algo quebrado)."""
    if not cfg.has_section(secao):
        return False
    if not cfg.getboolean(secao, "ativo", fallback=False):
        return False
    for campo in campos_obrigatorios:
        if not cfg.get(secao, campo, fallback="").strip():
            return False
    return True


def caminho_config_real_existe():
    return CONFIG_REAL.exists()
