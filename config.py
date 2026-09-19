import os
import logging
from pathlib import Path
from dotenv import load_dotenv

# log config
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# diretorios
load_dotenv()
RAIZ = Path(__file__).resolve().parent # extratos-sellers/
DIR_OUTPUT = RAIZ / "output"
DIR_DEBUG = RAIZ / "debug"
DIR_DOWNLOADS = Path.home() / "Downloads"
DIR_ARQUIVOS_FONTE = Path(os.getenv("ARQUIVOS_FONTE"))

# arquivos fonte
ARQ_CONTASARECEBER = next(DIR_DOWNLOADS.glob("visao_contas_a_receber*"), None)
ARQ_APPEL = DIR_ARQUIVOS_FONTE / "appel.xlsx"
ARQ_DATALAN = DIR_ARQUIVOS_FONTE / "datalan.xlsx"
ARQ_MAXISOFT = DIR_ARQUIVOS_FONTE / "maxisoft.xlsx"
ARQ_COMISSOES = DIR_ARQUIVOS_FONTE / "comissoes.xlsx"

# mapeamento de colunas
COLUNAS = {
    "IDENTIFICADOR_CLIENTE": "Identificador do cliente",
    "NOME_CLIENTE": "Nome do cliente",
    "DATA_COMPETENCIA": "Data de competência",
    "DATA_VENCIMENTO": "Data de vencimento",
    "VALOR_ORIGINAL_PARCELA": "Valor original da parcela (R$)",
    "VALOR_RECEBIDO_PARCELA": "Valor recebido da parcela (R$)",
    "DATA_ULTIMO_PAGAMENTO": "Data do último pagamento",
    "VALOR_TOTAL_RECEBIDO_PARCELA": "Valor total recebido da parcela (R$)"
}

COLUNAS_CRIADAS = {
    "NOME_RESSELLER": "nome_resseller",
    "SELLERS": "SELLERS",
    "RESSELLERS": "RESSELLERS",
    "COMISSAO": "COMISSÃO",
    "TOTAL_REPASSE": "TOTAL DO REPASSE (R$)",
}