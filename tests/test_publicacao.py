import subprocess
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ / "scripts"))

from auditar_publicacao import Achado, auditar, git_arquivos  # noqa: E402

ARROBA = "@"


def _arquivo(tmp_path, nome, texto):
    destino = tmp_path / nome
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(texto, encoding="utf-8")
    return nome


def _tipos(achados):
    return sorted({a.tipo for a in achados})


def test_arquivo_limpo_nao_gera_achado(tmp_path):
    nome = _arquivo(tmp_path, "ok.md", f"fale com notas{ARROBA}papaki.com.br ou ana{ARROBA}exemplo.com\nC:\\Users\\<usuário>\\x\n")
    assert auditar(tmp_path, [nome]) == []


def test_email_fora_da_lista(tmp_path):
    nome = _arquivo(tmp_path, "a.md", f"linha 1\nescreva para fulano{ARROBA}gmail.com\n")
    assert auditar(tmp_path, [nome]) == [Achado("a.md", 2, "e-mail fora da lista permitida")]


def test_documento_formatado_so_vale_em_tests(tmp_path):
    cpf = "123." + "456." + "789-" + "09"
    cnpj = "12." + "345." + "678/" + "0001-" + "95"
    a = _arquivo(tmp_path, "docs/x.md", f"{cpf} e {cnpj}\n")
    b = _arquivo(tmp_path, "tests/t.py", f"{cpf} e {cnpj}\n")
    achados = auditar(tmp_path, [a, b])
    assert [(x.arquivo, x.linha, x.tipo) for x in achados] == [("docs/x.md", 1, "CPF ou CNPJ formatado")]


def test_caminho_de_usuario_do_windows(tmp_path):
    ruim = "C:" + "\\Users\\" + "Fulano\\Dados"
    ruim2 = "D:/Users/" + "Fulano/x"
    ok = "C:\\Users\\<usuário>\\Dados"
    nome = _arquivo(tmp_path, "c.md", f"{ruim}\n{ruim2}\n{ok}\n")
    assert [(a.linha, a.tipo) for a in auditar(tmp_path, [nome])] == [
        (1, "caminho de usuário do Windows"), (2, "caminho de usuário do Windows")]


def test_chaves_e_codigos_do_supabase(tmp_path):
    ref = "a" * 20
    texto = "\n".join([
        f"https://{ref}.supabase.co",
        f"project_ref={ref}",
        f"o ref `{ref}` do projeto",
        "eyJ" + "x" * 30,
        "sb_secret_" + "y" * 12,
        "sbp_" + "0" * 24,
        "texto normal com a palavra refatoração e ref curto",
    ])
    nome = _arquivo(tmp_path, "s.md", texto)
    achados = auditar(tmp_path, [nome])
    assert [a.linha for a in achados] == [1, 2, 3, 4, 5, 6]


def test_arquivos_proibidos_pelo_nome(tmp_path):
    nomes = [_arquivo(tmp_path, n, "x") for n in ("a.xlsx", "b.csv", "dir/c.XLS", ".env", "ok.env.example", "HANDOFF-x.md")]
    achados = auditar(tmp_path, nomes)
    assert sorted(a.arquivo for a in achados) == sorted(["a.xlsx", "b.csv", "dir/c.XLS", ".env", "HANDOFF-x.md"])
    assert all(a.linha == 0 for a in achados)


def test_o_relatorio_nunca_mostra_o_trecho_achado(tmp_path, capsys):
    from auditar_publicacao import main
    _arquivo(tmp_path, "a.md", f"fulano{ARROBA}gmail.com\n")
    subprocess.run(["git", "init", "-q"], cwd=tmp_path, check=True)
    subprocess.run(["git", "add", "a.md"], cwd=tmp_path, check=True)
    assert main([str(tmp_path)]) == 1
    saida = capsys.readouterr()
    assert "a.md:1" in saida.out and "fulano" not in saida.out + saida.err


def test_o_repositorio_passa_na_auditoria():
    if not (RAIZ / ".git").exists():
        pytest.skip("sem .git")
    try:
        arquivos = git_arquivos(RAIZ)
    except (OSError, subprocess.CalledProcessError):
        pytest.skip("git indisponível")
    assert auditar(RAIZ, arquivos) == []
