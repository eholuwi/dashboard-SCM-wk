# -*- coding: utf-8 -*-
"""
Cliente da API do SCM — somente leitura
=======================================
Endpoints validados na Fase 0 (ver FASE0_VALIDACAO_API.md). A API é anônima na
rede interna: basta ter rota até mansrvapp03:5715.

Nunca chamar endpoint de escrita a partir daqui: este módulo só expõe os GET e
os dois POST cuja semântica é de consulta (geram relatório, não alteram dado).
"""

import json
import os
import time
import urllib.error
import urllib.request

BASE_URL = os.environ.get("SCM_API_URL", "http://mansrvapp03:5715/api")

# A API não pagina e já reciclou sob varredura durante o levantamento: manter
# timeout generoso, backoff e um respiro entre chamadas em rajada.
TIMEOUT_PADRAO = 180
TENTATIVAS = 3
BACKOFF_BASE = 2.0
THROTTLE = 0.05


class SCMApiError(RuntimeError):
    pass


def _requisitar(method, path, body=None, timeout=TIMEOUT_PADRAO):
    """Executa a chamada com retry/backoff. Devolve os bytes crus da resposta."""
    url = BASE_URL + path
    headers = {"Accept": "application/json"}
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json;charset=UTF-8"

    ultimo_erro = None
    for tentativa in range(1, TENTATIVAS + 1):
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except (urllib.error.URLError, OSError) as e:
            ultimo_erro = e
            if tentativa < TENTATIVAS:
                time.sleep(BACKOFF_BASE ** tentativa)
    raise SCMApiError(f"{method} {path} falhou após {TENTATIVAS} tentativas: {ultimo_erro}")


def _json(path, timeout=TIMEOUT_PADRAO):
    """GET que devolve JSON, desembrulhando o envelope Result<T> quando houver."""
    bruto = _requisitar("GET", path, timeout=timeout)
    time.sleep(THROTTLE)
    if not bruto:
        return []
    dados = json.loads(bruto.decode("utf-8-sig"))
    if isinstance(dados, dict) and "result" in dados:
        return dados["result"]
    return dados


# =====================================================================
# Endpoints
# =====================================================================
def relatorio_solicitacoes_xlsx(cod_usuario, ini, fim):
    """Relatório "Solicitações" como .xlsx — o MESMO arquivo que era exportado à
    mão pela tela do SCM (47 colunas, 6 abas).

    `cod_usuario` define o escopo por PERMISSÃO, não por autoria: um comprador
    ou administrador enxerga todas as SCs; um solicitante comum só as próprias.
    Datas no formato yyyyMMdd.
    """
    bruto = _requisitar("POST", f"/Relatorios/GetSolicitacoesByDate/{cod_usuario}/{ini}/{fim}",
                        body={})
    if bruto[:2] != b"PK":
        raise SCMApiError(
            f"esperado .xlsx do relatório de Solicitações, veio {len(bruto)} bytes "
            f"começando com {bruto[:16]!r} (usuário {cod_usuario} tem permissão?)")
    return bruto


def pedidos_periodo(ini, fim):
    """Todos os Pedidos (P.O.) do período, nível cabeçalho, com STATUS de entrega.
    Campos: C7_NUM, C7_FILIAL, C7_FORNECE, A2_NOME, C7_EMISSAO, TOTAL, STATUS,
    DtAprovado, COMPRADOR. Datas yyyyMMdd."""
    return _json(f"/Pedidos/{ini}/{fim}")


def pedido_itens(filial, numero):
    """Itens de um Pedido (tabela SC7 do Protheus). Pedidos eliminados podem vir
    como lista vazia."""
    return _json(f"/Pedidos/ByNumero/{filial}/{numero}", timeout=60)


def centros_custo():
    """Mapa {código: descrição} dos centros de custo — é a origem do Departamento."""
    return {str(c.get("codigo", "")).strip(): str(c.get("descricao", "")).strip()
            for c in _json("/CentroCusto")}


def compradores():
    """Usuários marcados como comprador (para descobrir um código com permissão
    ampla no relatório de Solicitações)."""
    return _json("/Usuario/Compradores")


def disponivel():
    """True se a API responde agora — usado para decidir por fallback manual."""
    try:
        _requisitar("GET", "/Filial", timeout=15)
        return True
    except SCMApiError:
        return False
