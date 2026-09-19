import pandas as pd
from .style import apply_style
from config import logging, COLUNAS, COLUNAS_CRIADAS, DIR_OUTPUT


def save(contas_a_receber: pd.DataFrame) -> None:
    """
    Salva relatórios individuais em Excel a partir do DataFrame final de contas a
    receber.
    Remove colunas auxiliares, agrupa os dados por seller e, para cada grupo, cria um
    arquivo .xlsx contendo as linhas do seller e uma linha extra com o total de repasses.
    """

    logging.info("[save] -> salvando dataframes...")

    try:
        DIR_OUTPUT.mkdir(parents=True, exist_ok=True)

        contas_a_receber = contas_a_receber.drop(
            columns=[
                COLUNAS_CRIADAS["RESSELLERS"],
                COLUNAS_CRIADAS["NOME_RESSELLER"],
            ]
        )

        grupos_de_sellers = contas_a_receber.groupby(COLUNAS_CRIADAS["SELLERS"])

        for seller, planilha_seller in grupos_de_sellers:
            logging.info(f"Processando relatório para o seller: '{seller}.xlsx')")

            planilha_seller.sort_values(by=COLUNAS["NOME_CLIENTE"], inplace=True)
            soma_dos_repasses = planilha_seller[COLUNAS_CRIADAS["TOTAL_REPASSE"]].sum()

            # cria um dataframe com a soma de repasses
            linha_total = pd.DataFrame(
                [
                    {
                        planilha_seller.columns[0]: "TOTAL",
                        COLUNAS_CRIADAS["TOTAL_REPASSE"]: soma_dos_repasses,
                    }
                ]
            )

            nome_arquivo = f"{seller}.xlsx"
            caminho_completo_do_arquivo = DIR_OUTPUT / nome_arquivo

            # junta o dataframe com a linha de total com o contas_a_receber
            planilha_com_total = pd.concat(
                [planilha_seller, linha_total], ignore_index=True
            )

            with pd.ExcelWriter(
                caminho_completo_do_arquivo, engine="openpyxl"
            ) as writer:
                planilha_com_total.to_excel(
                    writer, sheet_name=f"Extrato {seller}", index=False
                )
                worksheet = writer.sheets[f"Extrato {seller}"]
                apply_style(worksheet, nome_coluna_valor="TOTAL DO REPASSE (R$)")

            logging.info(f"Relatório para o seller '{seller}.xlsx' salvo com estilos.")
        
        logging.info("[save] -> Finalizado!!!!!!!")
    except Exception as e:
        raise FileNotFoundError(f"[save] -> Erro ao salvar arquivos: {e}") from e
