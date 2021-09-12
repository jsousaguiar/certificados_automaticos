# IMPORTAÇÕES
import pandas as pd
import jinja2
from docxtpl import DocxTemplate
from docx2pdf import convert
from datetime import date
import logging
import os
import hjson
import os.path
import numpy as np

from cpf_cnpj import formatar_cpf_cnpj

###################################################################################################
# CRIA PASTA PARA SALVAR OS CERTIFICADOS

os.chdir(os.path.dirname(globals()["__file__"]))  # RFSA

pasta = r"./certificados"
if not os.path.isdir(pasta):  # verifica e cria a pasta "certificados", caso não exista
    os.mkdir(pasta)

###################################################################################################
# CONFIGURAÇÃO DE LOGS

logging.basicConfig(
    level=logging.INFO,
    # para formato mais simples, semelhante ao print(), pode-se usar:
    # format="%(message)s",
    format="%(asctime)s [%(name)s] [%(levelname)s]: %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S",
)
LOGGER = logging.getLogger("GERADOR DE CERTIFICADOS")

###################################################################################################
# IMPORTANDO CONFIGURAÇÕES DO ARQUIVO config.hjson

config_data = {}
if os.path.exists("config.hjson"):
    with open("config.hjson") as config_file:
        try:
            config_data = hjson.load(config_file)
        except hjson.HjsonDecodeError:
            LOGGER.critical(
                "Erro ao decodificar arquivo de configuração:", exc_info=True
            )
            exit(1)


###################################################################################################
# VARIÁVEIS PARA O CERTIFICADO
meses = (
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
)

nome_curso = config_data.get("nome_do_curso", "INSERIR NOME DO CURSO EM config.hjson")
nome_instituicao = config_data.get(
    "nome_instituicao", "INSERIR NOME DA INSTITUIÇÃO EM config.hjson"
)
carga_horaria = config_data.get(
    "carga_horaria", "INSERIR CARGA HORÁRIA EM config.hjson"
)
cidade = config_data.get("cidade")
dia_inicio = int(config_data.get("dia_inicio", 1))
mes_inicio = int(config_data.get("mes_inicio", 1))
ano_inicio = int(config_data.get("ano_inicio", 9999))
dia_fim = int(config_data.get("dia_fim", 1))
mes_fim = int(config_data.get("mes_fim", 2))
ano_fim = int(config_data.get("ano_fim", 9999))

data_inicio = date(year=ano_inicio, month=mes_inicio, day=dia_inicio)
data_fim = date(year=ano_fim, month=mes_fim, day=dia_fim)
dias_de_curso = (data_fim - data_inicio).days
mes_extenso = meses[(mes_fim) - 1]

###################################################################################################
# Formatar CPF/CNPJ


def formatar_cpf_cnpj_se_presente(cpf_cnpj: np.int64) -> str:
    if cpf_cnpj == 0 or cpf_cnpj == "":
        return ""
    return f"{formatar_cpf_cnpj(cpf_cnpj)}"


###################################################################################################
# IMPORTAR LISTA DE INSCRITOS
arquivo_inscricoes = config_data.get("arquivo_inscricoes")
path = f"./{arquivo_inscricoes}"
inscritos = pd.read_excel(path)
inscritos = inscritos.fillna("")
inscritos.to_dict(orient="records")
nome_coluna_nome = config_data.get("nome_coluna_nome")
nome_coluna_cpf = config_data.get("nome_coluna_cpf")


###################################################################################################
# RENDERIZAÇÃO
def gerar_certificado(nome: str, cpf: str) -> bool:
    LOGGER.info(f"Gerando certificado de {nome}, CPF {cpf}...")

    # Formatar o CPF
    cpf = formatar_cpf_cnpj_se_presente(cpf)

    arquivo_template = r"./modelo_certificado.docx"
    arquivo_destino = f"./certificados/certificado_{nome}.docx"
    template = DocxTemplate(arquivo_template)

    jinja_env = jinja2.Environment()

    context = {
        "nome_pessoa": nome,
        "cpf": cpf,
        "nome_instituicao": nome_instituicao,
        "curso": nome_curso,
        "carga_horaria": carga_horaria,
        "cidade": cidade,
        "data_inicio": f"{data_inicio.day}/{data_inicio.month}/{data_inicio.year}",
        "data_fim": f"{data_fim.day}/{data_fim.month}/{data_fim.year}",
        "dias_de_curso": dias_de_curso,
        "dia": f"{data_fim.day}",
        "mes_extenso": mes_extenso,
        "ano": f"{data_fim.year}",
    }

    template.render(context, jinja_env)
    template.save(arquivo_destino)
    template = None
    convert(
        f"./certificados/certificado_{nome}.docx",
        f"./certificados/certificado_{nome}.pdf",
    )
    os.remove(f"./certificados/certificado_{nome}.docx")
    return True


certificados = 0
print()
for (indice, inscrito) in inscritos.iterrows():
    nome = inscrito[nome_coluna_nome]
    cpf = formatar_cpf_cnpj_se_presente(inscrito[nome_coluna_cpf])
    print(f"Gerando certificado para {nome}, CPF {cpf}...")
    certificados += 1 if gerar_certificado(nome, cpf) else 0
print()


if certificados == 0:
    mensagem = "Não foram gerados certificados."
elif certificados == 1:
    mensagem = "Foi gerado apenas um certificado."
else:
    mensagem = f"Foram gerados {certificados} certificados."

LOGGER.info(mensagem)
print()
