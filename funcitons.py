"""Compatibility wrapper for legacy module name.

Use functions.py for all business logic.
"""

import argparse

from functions import (
    ABA_PADRAO,
    ARQUIVO_TEMPLATE,
    PASTA_OUTPUT,
    processar_output_uma_vez,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Verifica a pasta Output e reaplica formulas nos arquivos Excel."
    )
    parser.add_argument(
        "--output",
        default=str(PASTA_OUTPUT),
        help="Pasta monitorada (padrao: Output)",
    )
    parser.add_argument(
        "--template",
        default=str(ARQUIVO_TEMPLATE),
        help="Arquivo template com formulas (padrao: InputTemplate.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        default=ABA_PADRAO,
        help="Nome da aba para copiar formulas (padrao: Monthly CF)",
    )

    args = parser.parse_args()
    processar_output_uma_vez(
        path_output=args.output,
        arquivo_template=args.template,
        nome_aba=args.sheet,
    )


if __name__ == "__main__":
    main()
