"""Extrações sintéticas para os testes: arquivos inventados, sem dados reais."""
import datetime as dt
import io
import zipfile
from pathlib import Path

from papaki_dre.extracao import RELATORIOS, iniciar

OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64


def _xlsx_minimo() -> bytes:
    """Um zip com a estrutura mínima de um .xlsx (só o que a conferência de assinatura olha)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
    return buf.getvalue()


ZIP = _xlsx_minimo()
CONTEUDO = {
    "cp": ("Contas a pagar.csv", b"\xef\xbb\xbfCod;Valor\n1;10\n"),
    "cr": ("Contas a receber.csv", b"Emissao;Valor\n01/01/2026;10\n"),
    "rra": ("Relatorio Rentabilidade.xls", OLE),
    "comissoes": ("Comissoes.xlsx", ZIP),
}
_BASE = dict(data_inicio="2026-01-01", data_fim="2026-09-30", status="Todos")
TELA = {
    "cp": dict(localidade="Todas", criterio="Data competência", total=37030,
               pesquisar_em="Individuais e agrupadas", **_BASE),
    "cr": dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA", criterio="Data competência", total=483, **_BASE),
    "rra": dict(localidade="Todas", criterio="Aprovação", total=1410, **_BASE),
    "comissoes": dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA", criterio="Aprovação", total=88, **_BASE),
}


def tela(rel, **mudar):
    return {**TELA[rel], **mudar}


def png(pasta: Path, nome="tela.png") -> Path:
    p = Path(pasta) / nome
    p.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 16)
    return p


def mensagem_rra(ex, id_="msg1"):
    quando = dt.datetime.fromisoformat(ex.estado["relatorios"]["rra"]["disparo"])
    return {"id": id_, "remetente": "Papaki <notas@papaki.com.br>",
            "assunto": "Relatório de rentabilidade gerado com sucesso",
            "corpo": f"solicitado em {quando.strftime('%d/%m/%Y %H:%M:%S')}"}


def extracao_fechada(tmp_path, *, produtora="untitled", competencia="2026-09", periodo_fim="2026-09-30", agora=None):
    """Faz o caminho completo (tela -> disparo -> arquivo) dos 4 relatórios e fecha. Devolve a Extracao."""
    base = Path(tmp_path) / "dados" / "extracoes" / competencia / produtora
    ex = iniciar(base, produtora, competencia, periodo_fim, agora=agora)
    print_ = png(tmp_path)
    downloads = Path(tmp_path) / "Downloads"
    downloads.mkdir(exist_ok=True)
    ex.rra_ids_antes(["velha"])
    for rel in RELATORIOS:
        ex.registrar_tela(rel, tela(rel), print_)
        ex.disparar(rel)
        if rel == "rra":
            ex.escolher_rra([mensagem_rra(ex)])
        nome, dados = CONTEUDO[rel]
        arquivo = downloads / nome
        arquivo.write_bytes(dados)           # criado DEPOIS do disparo
        ex.registrar_arquivo(rel, arquivo)
    ex.fechar()
    return ex
