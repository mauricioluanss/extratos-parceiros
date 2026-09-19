import pandas as pd
from pathlib import Path
from openpyxl.utils.exceptions import InvalidFileException
from config import logging, COLUNAS, ARQ_COMISSOES, ARQ_CONTASARECEBER

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def load_file(file_path: Path, cols: list = None) -> pd.DataFrame:
    """
    Carrega um arquivo Excel em um DataFrame, mantendo apenas as colunas especificadas
    (se fornecidas), ajustando o separador decimal para vírgula e removendo espaços
    extras dos nomes das colunas.
    """
    if file_path is None:
        raise FileNotFoundError("Arquivo fonte não encontrado")

    df = pd.read_excel(file_path, usecols=cols, decimal=",")
    df.columns = df.columns.str.strip()
    return df


def load() -> tuple[pd.DataFrame]:
    """
    Carrega e retorna os arquivos de contas a receber (com colunas específicas) e de
    comissões como DataFrames.
    """

    try:
        logging.info(f"[load] -> arquivo de comissoes: {ARQ_COMISSOES}")
        logging.info(f"[load] -> arquivo de contas a receber: {ARQ_CONTASARECEBER}")

        df_contas_a_receber = load_file(ARQ_CONTASARECEBER, COLUNAS.values())
        df_comissoes = load_file(ARQ_COMISSOES)

        return df_contas_a_receber, df_comissoes

    except FileNotFoundError as e:
        raise ValueError(f"[load] -> {e}") from e

    except ValueError as e:
        raise ValueError(f"[load] -> Arquivo vazio -> {e}") from e

    except InvalidFileException as e:
        raise ValueError(f"[load] -> Tipo de arquivo inválido -> {e}") from e

    except Exception as e:
        raise ValueError(f"[load] -> Erro ao carregar planilhas -> {e}") from e


if __name__ == "__main__":
    load()
