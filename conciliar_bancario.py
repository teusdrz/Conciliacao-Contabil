#!/usr/bin/env python3
"""
Conciliação Bancária - 01.Razão x 02.Financeiro (ponto de entrada independente).

Diferente do main.py (que roda o motor de conciliação intra-razão em
motor_conciliacao.py), este script casa os lançamentos do razão contra o
extrato financeiro (documento suporte) de um MESMO arquivo, lançamento a
lançamento, e grava o resultado numa aba nova - sem alterar nenhuma aba,
fórmula ou dado existente no arquivo original.

Uso:
    python conciliar_bancario.py --arquivo caminho/arquivo.xlsx --saida saida.xlsx
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from conciliacao_bancaria import (  # noqa: E402
    ABA_FINANCEIRO_PADRAO,
    ABA_RAZAO_PADRAO,
    ABA_SAIDA_PADRAO,
    executar,
)


def _configurar_log(verbose: bool) -> None:
    nivel = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=nivel,
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # bibliotecas de terceiro (PIL, via openpyxl ao carregar imagens da planilha)
    # nao precisam de DEBUG mesmo com --verbose - so gerava ruido no log.
    logging.getLogger("PIL").setLevel(logging.WARNING)


def montar_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Conciliação bancária linha a linha: 01.Razão x 02.Financeiro.",
    )
    p.add_argument("--arquivo", required=True, help="Caminho do arquivo .xlsx de origem")
    p.add_argument("--aba-razao", default=ABA_RAZAO_PADRAO, help=f'Nome da aba do razão (padrão: "{ABA_RAZAO_PADRAO}")')
    p.add_argument("--aba-financeiro", default=ABA_FINANCEIRO_PADRAO, help=f'Nome da aba do extrato (padrão: "{ABA_FINANCEIRO_PADRAO}")')
    p.add_argument("--linha-cabecalho-razao", type=int, default=5, help="Linha do cabeçalho do razão (padrão: 5)")
    p.add_argument("--nome-aba-saida", default=ABA_SAIDA_PADRAO, help=f'Nome da aba nova de resultado (padrão: "{ABA_SAIDA_PADRAO}")')
    p.add_argument("--tolerancia-valor", type=float, default=0.01, help="Tolerância em R$ para casar valores (padrão: 0.01)")
    p.add_argument("--tolerancia-dias", type=int, default=5, help="Tolerância de dias entre data do razão e do extrato (padrão: 5)")
    p.add_argument("--saida", default=None, help="Caminho do Excel de saída (padrão: <arquivo>_bancario_<data>.xlsx)")
    p.add_argument("--verbose", action="store_true", help="Log detalhado (DEBUG)")
    return p


def main(argv: list[str] | None = None) -> int:
    args = montar_parser().parse_args(argv)
    _configurar_log(args.verbose)
    logger = logging.getLogger("conciliador.conciliar_bancario")

    nome_base = Path(args.arquivo).stem
    saida = args.saida or f"{nome_base}_bancario_{datetime.now():%Y%m%d}.xlsx"

    logger.info("Conciliando %s [razão=%s, financeiro=%s]", args.arquivo, args.aba_razao, args.aba_financeiro)
    resultado = executar(
        arquivo=args.arquivo,
        arquivo_saida=saida,
        aba_razao=args.aba_razao,
        aba_financeiro=args.aba_financeiro,
        linha_cabecalho_razao=args.linha_cabecalho_razao,
        nome_aba_saida=args.nome_aba_saida,
        tolerancia_valor=args.tolerancia_valor,
        tolerancia_dias=args.tolerancia_dias,
    )

    logger.info(
        "Concluído: %d conciliado(s), %d só no razão, %d só no financeiro. Diferença total: %.2f. Saída: %s",
        resultado.qtd_conciliado, resultado.qtd_so_razao, resultado.qtd_so_financeiro,
        resultado.total_diferenca, saida,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
