from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from config import logging


def apply_style(worksheet, nome_coluna_valor: str):
    """
    Aplica estilos a uma planilha de extrato.

    - Cabeçalho: Fundo preto, fonte branca, negrito e centralizado.
    - Coluna de Valor:
        - Fundo verde para valores > 0.
        - Fundo amarelo para valores <= 0 ou não numéricos.
    - Ajusta a largura de todas as colunas.
    """

    estilo_cabecalho = PatternFill(
        start_color="000000", end_color="000000", fill_type="solid"
    )
    fonte_cabecalho = Font(color="FFFFFF", bold=True)
    alinhamento_central = Alignment(horizontal="center", vertical="center")

    preenchimento_verde = PatternFill(
        start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
    )
    preenchimento_amarelo = PatternFill(
        start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"
    )

    letra_coluna_valor = None

    for cell in worksheet[1]:
        cell.fill = estilo_cabecalho
        cell.font = fonte_cabecalho
        cell.alignment = alinhamento_central
        if cell.value == nome_coluna_valor:
            letra_coluna_valor = get_column_letter(cell.column)

    if not letra_coluna_valor:
        logging.error(
            f"Aviso: A coluna '{nome_coluna_valor}' não foi encontrada. A formatação condicional não será aplicada."
        )
        return

    for cell in worksheet[letra_coluna_valor][1:]:
        # Ignora a última linha que contém o texto "TOTAL"
        if isinstance(cell.value, (int, float)):
            if cell.value > 0:
                cell.fill = preenchimento_verde
            else:
                cell.fill = preenchimento_amarelo
        else:
            primeira_celula_da_linha = worksheet.cell(row=cell.row, column=1).value
            if primeira_celula_da_linha != "TOTAL":
                cell.fill = preenchimento_amarelo

    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = max_length + 2
        worksheet.column_dimensions[column_letter].width = adjusted_width
