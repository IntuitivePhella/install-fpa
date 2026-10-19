"""Constantes de negócio. Fonte: instrucoes_relatorios_ base.md."""
import datetime as dt

# Regras e dados de cada produtora (Localidade, Núcleo do RRA, modelo de comissões, linha-mestre da
# fórmula de DESPESAS!AP, endereços do Esferas) ficam em produtoras/<id>.json (veja produtoras.py).

PESQUISA = "PESQUISA E TRATAMENTO"

# Primeira linha de dados e linha de cabeçalho no DESTINO (modelo Excel).
DEST = {
    "rra": {"sheet": "BASE", "title_row": 1, "header_row": 3, "first_row": 5},
    "cp": {"sheet": "DESPESAS", "header_row": 2, "first_row": 3},
    "cr": {"sheet": "RECEITA", "header_row": 1, "first_row": 2},
    "comissoes": {"sheet": "REAL X ORÇADO COMISSÕES", "header_row": 1, "first_row": 2},
}

# Motivos de remoção (usados em `exclusoes` no banco e no relatório).
R_BLANK = "linha_vazia"
R_FOOTER = "rodape_total"
R_JOB = "job_preenchido_sem_pesquisa"
R_LOC = "fora_localidade_untitled"
R_NUCLEO = "nucleo_diferente_de_untitled"
R_PESQ = "pesquisa_e_tratamento"
R_MODELO = "modelo_diferente"
R_H = "comissao_excluida_coluna_h"

# Tipo de cada coluna, por nome de cabeçalho. O que não estiver aqui é texto.
# text | date | datetime | money (estrito) | money_or_text
CP_KINDS = {
    **{n: "date" for n in (
        "Vcto", "Vcto alterado", "Data prevista", "Data entrada", "Emissão",
        "Competência", "Compensado", "Data documento", "Comp.")},
    "Criado em": "datetime",
    "Valor": "money",
    "Deduções": "money_or_text",
    "Alterado": "money_or_text",
}
CR_KINDS = {
    **{n: "date" for n in ("Emissão", "Vcto", "Competência")},
    **{n: "money" for n in ("Vl Documento", "Vl cobrado", "Vl recebido", "Dif. receber")},
    **{n: "money_or_text" for n in ("Desconto", "Outras ded.", "Multa", "Juros", "IDA", "Atraso")},
}
# O CR tem duas colunas "Recebido": a de índice 5 é data; a de índice 24 é Sim/Não (texto).
CR_INDEX_KINDS = {5: "date"}

# Rótulos dos rodapés de TOTAL. Só servem para reconhecer a linha como técnica
# (R_FOOTER): os valores dos rodapés nunca são lidos.
CP_FOOTER = (4, "VALOR TOTAL")                       # (índice da coluna, rótulo exato)
CR_FOOTER = (6, ("TOTAL PÁGINA", "TOTAL GERAL"))
RRA_FOOTER = "TOTAL"                                 # coluna A
COMISSOES_FOOTER = "TOTAL GERAL"                     # coluna H, com a coluna A vazia

# Fuso de trabalho (Brasília, sem horário de verão desde 2019). Datas sem fuso valem neste fuso.
FUSO_BRASILIA = dt.timezone(dt.timedelta(hours=-3))

# Feriados além dos nacionais (ISO AAAA-MM-DD). A Papaki acrescenta os locais aqui.
FERIADOS_EXTRAS: tuple[str, ...] = ()

# Nomes das subpastas sob PAPAKI_DADOS_DIR (a instalação pode trocá-los por nomes já usados na Papaki).
NOME_PASTA_EXTRACOES = "extracoes"
NOME_PASTA_APURACOES = "apuracoes"
