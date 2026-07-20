#!/usr/bin/env python3
"""
Conciliador Contabil Inteligente - ponto de entrada.

Uso basico:
    python main.py --arquivo caminho/para/razao.xlsx

Todas as opcoes tambem podem vir de um config.yaml (veja config/config.exemplo.yaml):
    python main.py --config config/config.exemplo.yaml

Parametros de linha de comando sempre têm prioridade sobre o config.yaml.
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).parent / "src"))

from classificador import aplicar_classificacao  # noqa: E402
from excel_io import atualizar_obs_arquivo_original, carregar_razao, ler_saldo_balancete  # noqa: E402
from motor_conciliacao import ConciliadorContabil  # noqa: E402
from relatorios import gerar_excel_saida  # noqa: E402


def _configurar_log(verbose: bool) -> None:
    nivel = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=nivel,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _carregar_config(caminho: str | None) -> dict:
    if not caminho:
        return {}
    with open(caminho, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def montar_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Conciliador Contábil Inteligente - concilia entradas (crédito) "
        "com saídas (débito) em uma conta de passivo/provisão a partir de um razão em Excel."
    )
    p.add_argument("--config", help="Caminho para um config.yaml com os parâmetros abaixo")
    p.add_argument("--arquivo", help="Caminho do arquivo .xlsx de origem (o razão contábil)")
    p.add_argument("--aba", default=None, help='Nome da aba do razão (padrão: "01.Razão (2)")')
    p.add_argument("--linha-cabecalho", type=int, default=None, help="Linha do cabeçalho (padrão: 6)")
    p.add_argument("--conta", default=None, help="Código da conta contábil (para o relatório de saída)")
    p.add_argument("--tolerancia", type=float, default=None, help="Tolerância em R$ para os matches (padrão: 0.01)")
    p.add_argument("--max-grupo", type=int, default=None, help="Tamanho máximo de grupo na Etapa 3 (padrão: 6)")
    p.add_argument("--similaridade-min", type=int, default=None, help="Similaridade mínima (%%) na Etapa 2 (padrão: 80)")
    p.add_argument("--data-corte", default=None, help="Data de corte p/ aging, formato AAAA-MM-DD (padrão: data mais recente do razão)")
    p.add_argument(
        "--balancete-arquivo", default=None,
        help="Arquivo com a aba de balancete p/ a Ponte (padrão: o mesmo --arquivo)",
    )
    p.add_argument("--balancete-aba", default=None, help='Nome da aba de balancete (padrão: "00.Balancete")')
    p.add_argument("--saldo-balancete", type=float, default=None, help="Informe manualmente o saldo do balancete, se preferir não ler de uma aba")
    p.add_argument("--saida", default=None, help="Caminho do Excel de saída")
    p.add_argument(
        "--saida-original", default=None,
        help="Caminho do arquivo original atualizado (mesmas abas do --arquivo, só a "
        "coluna Obs. da aba do razão é preenchida). Padrão: <arquivo>_atualizado.xlsx",
    )
    p.add_argument("--verbose", action="store_true", help="Log detalhado (DEBUG)")
    return p


def main(argv: list[str] | None = None) -> int:
    args = montar_parser().parse_args(argv)
    cfg = _carregar_config(args.config)

    def opt(nome_cli, nome_cfg, padrao):
        valor_cli = getattr(args, nome_cli)
        if valor_cli is not None:
            return valor_cli
        return cfg.get(nome_cfg, padrao)

    _configurar_log(args.verbose or cfg.get("verbose", False))
    logger = logging.getLogger("conciliador.main")

    arquivo = opt("arquivo", "arquivo", None)
    if not arquivo:
        logger.error("Informe --arquivo ou um config.yaml com a chave 'arquivo'.")
        return 1

    aba = opt("aba", "aba", "01.Razão (2)")
    linha_cabecalho = opt("linha_cabecalho", "linha_cabecalho", 6)
    conta = opt("conta", "conta", "")
    tolerancia = opt("tolerancia", "tolerancia", 0.01)
    max_grupo = opt("max_grupo", "max_grupo", 6)
    similaridade_min = opt("similaridade_min", "similaridade_min", 80)
    data_corte_str = opt("data_corte", "data_corte", None)
    balancete_arquivo = opt("balancete_arquivo", "balancete_arquivo", arquivo)
    balancete_aba = opt("balancete_aba", "balancete_aba", "00.Balancete")
    saldo_balancete_manual = opt("saldo_balancete", "saldo_balancete", None)

    nome_base = Path(arquivo).stem
    saida_padrao = f"{nome_base}_conciliado_{datetime.now():%Y%m%d}.xlsx"
    saida = opt("saida", "saida", saida_padrao)
    saida_original_padrao = f"{nome_base}_atualizado.xlsx"
    saida_original = opt("saida_original", "saida_original", saida_original_padrao)

    logger.info("Carregando razão: %s [aba=%s]", arquivo, aba)
    carregado = carregar_razao(arquivo, aba=aba, linha_cabecalho=linha_cabecalho)

    df = aplicar_classificacao(carregado.df)

    data_corte = None
    if data_corte_str:
        from datetime import datetime as _dt
        data_corte = _dt.strptime(data_corte_str, "%Y-%m-%d")

    motor = ConciliadorContabil(
        df,
        tolerancia=tolerancia,
        max_grupo=max_grupo,
        similaridade_min=similaridade_min,
        data_corte=data_corte,
    )
    resultado = motor.rodar_cascata()
    resumo = motor.resumo_por_periodo()

    saldo_balancete = saldo_balancete_manual
    if saldo_balancete is None and conta:
        saldo_balancete = ler_saldo_balancete(balancete_arquivo, conta, aba=balancete_aba)
        if saldo_balancete is not None:
            logger.info("Saldo do balancete lido automaticamente: %.2f", saldo_balancete)
        else:
            logger.warning(
                "Não foi possível localizar o saldo do balancete automaticamente. "
                "A aba Ponte_Balancete ficará sem o item (1). Use --saldo-balancete para informar manualmente."
            )
    ponte = motor.ponte_balancete(saldo_balancete)

    periodo_ref = data_corte.strftime("%m/%Y") if data_corte else str(df["data"].max().strftime("%m/%Y"))
    gerar_excel_saida(resultado, resumo, ponte, saida, conta=conta, periodo_referencia=periodo_ref)
    atualizar_obs_arquivo_original(arquivo, resultado, saida_original, aba=aba)

    total_aberto = resultado.loc[resultado["residual_centavos"] != 0, "valor_residual"].sum()
    logger.info("Concluído. Total em aberto após a cascata: %.2f", total_aberto)
    logger.info("Arquivo gerado: %s", saida)
    logger.info("Arquivo original atualizado: %s", saida_original)
    return 0


def app(environ, start_response):
    """Stub WSGI apenas para compatibilidade com plataformas de deploy que
    exigem uma variável de nível superior 'app'/'application'/'handler'.

    Este projeto é uma ferramenta de linha de comando (CLI), não um serviço
    web. Use: python main.py --arquivo caminho/para/razao.xlsx
    """
    status = "200 OK"
    headers = [("Content-Type", "text/plain; charset=utf-8")]
    start_response(status, headers)
    mensagem = (
        "Conciliador Contabil Inteligente e uma ferramenta de linha de comando (CLI).\n"
        "Use: python main.py --arquivo caminho/para/razao.xlsx\n"
    )
    return [mensagem.encode("utf-8")]


application = app
handler = app


if __name__ == "__main__":
    raise SystemExit(main())
