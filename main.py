"""Gera relatórios de contas a receber e comissões para cada parceiro comercial"""

__author__ = "Mauricio Luan"

from config import logging
from src.load import load
from src.transform import transform
from src.save import save


def main():
    try:
        contas_a_receber, comissoes = load()
        contas_a_receber_final = transform(contas_a_receber, comissoes)
        save(contas_a_receber_final)

    except Exception as e:
        logging.error(e)


if __name__ == "__main__":
    main()
