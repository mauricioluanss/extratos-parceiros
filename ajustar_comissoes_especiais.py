import numpy as np
import pandas as pd
from pathlib import Path
from src.style import apply_style
from config import (
    logging,
    DIR_OUTPUT,
    ARQ_APPEL,
    ARQ_DATALAN,
    ARQ_MAXISOFT,
    COLUNAS,
    COLUNAS_CRIADAS
)

def ajustar_repasse_por_lookup(
    caminho_arquivo_seller: Path,
    df_lookup: pd.DataFrame,
    col_lookup_valor: str,
    col_lookup_tarifa: str,
) -> None:
    """
    Ajusta o repasse de um seller aplicando regras de tarifa baseadas em um lookup (faixas de valor).
    """

    try:
        df_seller = pd.read_excel(caminho_arquivo_seller)
    except FileNotFoundError:
        logging.error(f"Arquivo do seller não encontrado: {caminho_arquivo_seller}")
        return

    # Remove a linha do total se já existir (para evitar duplicação ou erro de tipo)
    if "TOTAL" in df_seller[df_seller.columns[0]].values:
        df_seller = df_seller[df_seller[df_seller.columns[0]] != "TOTAL"]

    col_base_calculo = COLUNAS["VALOR_RECEBIDO_PARCELA"]
    # col_base_calculo = COLUNAS["VALOR_TOTAL_RECEBIDO_PARCELA"]

    df_seller[col_base_calculo] = df_seller[col_base_calculo].astype("float64")
    df_lookup[col_lookup_valor] = df_lookup[col_lookup_valor].astype("float64")

    df_seller_sorted = df_seller.sort_values(by=col_base_calculo)
    df_lookup_sorted = df_lookup.sort_values(by=col_lookup_valor)

    df_com_tarifa = pd.merge_asof(
        df_seller_sorted,
        df_lookup_sorted,
        left_on=col_base_calculo,
        right_on=col_lookup_valor,
        direction="backward",
    )

    df_com_tarifa = df_com_tarifa.set_index(df_seller_sorted.index)

    df_seller["TARIFA (R$)"] = df_com_tarifa[col_lookup_tarifa]
    df_seller["TARIFA (R$)"] = np.trunc(df_seller["TARIFA (R$)"] * 100) / 100

    df_seller[COLUNAS_CRIADAS["TOTAL_REPASSE"]] = (
        df_seller[col_base_calculo] - df_seller["TARIFA (R$)"]
    )

    colunas_para_limpar = [
        "VALOR",
        "DESCONTO",
        "TARIFA",
        "VALOR_DA_COMISSAO",
        "COMISSÃO",
        col_lookup_valor,  # Remove as colunas usadas no merge
        col_lookup_tarifa,
    ]
    for col in colunas_para_limpar:
        if col in df_seller.columns:
            df_seller = df_seller.drop(columns=col)

    # Reordenação final das colunas
    ordem_final_das_colunas = [
        COLUNAS["IDENTIFICADOR_CLIENTE"],
        COLUNAS["NOME_CLIENTE"],
        COLUNAS_CRIADAS["SELLERS"],
        COLUNAS["DATA_COMPETENCIA"],
        COLUNAS["DATA_VENCIMENTO"],
        COLUNAS["VALOR_ORIGINAL_PARCELA"],
        COLUNAS["VALOR_RECEBIDO_PARCELA"],
        #COLUNAS["VALOR_TOTAL_RECEBIDO_PARCELA"],
        COLUNAS["DATA_ULTIMO_PAGAMENTO"],
        "TARIFA (R$)",
        COLUNAS_CRIADAS["TOTAL_REPASSE"],
    ]

    colunas_existentes_na_ordem = [
        col for col in ordem_final_das_colunas if col in df_seller.columns
    ]
    df_seller = df_seller[colunas_existentes_na_ordem]

    soma_repasse = df_seller[COLUNAS_CRIADAS["TOTAL_REPASSE"]].sum()
    linha_total = pd.DataFrame(
        [
            {
                df_seller.columns[0]: "TOTAL",
                COLUNAS_CRIADAS["TOTAL_REPASSE"]: soma_repasse,
            }
        ]
    )

    df_final_com_total = pd.concat([df_seller, linha_total], ignore_index=True)

    with pd.ExcelWriter(caminho_arquivo_seller, engine="openpyxl") as writer:
        sheet_name = "Extrato"
        df_final_com_total.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        apply_style(worksheet, nome_coluna_valor=COLUNAS_CRIADAS["TOTAL_REPASSE"])


def main():
    ajustes_config = [
        {
            "arquivo_seller": DIR_OUTPUT / "APPEL.xlsx",
            "lookup_file": ARQ_APPEL,
            "col_valor": "VALOR",
            "col_tarifa": "DESCONTO",
            "ativo": True,
        },
        {
            "arquivo_seller": DIR_OUTPUT / "DATALAN.xlsx",
            "lookup_file": ARQ_DATALAN,
            "col_valor": "VALOR",
            "col_tarifa": "TARIFA",
            "ativo": True,
        },
        {
            "arquivo_seller": DIR_OUTPUT / "MAXISOFT.xlsx",
            "lookup_file": ARQ_MAXISOFT,
            "col_valor": "VALOR",
            "col_tarifa": "DESCONTO",
            "ativo": True,
        },
    ]

    for config in ajustes_config:
        if not config["ativo"]:
            continue

        arquivo_seller = config["arquivo_seller"]

        if arquivo_seller.exists():
            try:
                df_lookup = pd.read_excel(config["lookup_file"])

                ajustar_repasse_por_lookup(
                    arquivo_seller,
                    df_lookup,
                    config["col_valor"],
                    config["col_tarifa"],
                )
                logging.info(f"{arquivo_seller.stem} ajustado com sucesso.")

            except FileNotFoundError:
                logging.error(
                    f"Arquivo de lookup não encontrado: {config['lookup_file']}"
                )
            except Exception as e:
                logging.error(f"Erro ao processar {arquivo_seller.stem}: {e}")
        else:
            logging.warning(f"Arquivo para ajuste não encontrado: {arquivo_seller}")


if __name__ == "__main__":
    main()
