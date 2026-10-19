import pytest

from papaki_dre import ambiente


def test_cria_o_arquivo_com_as_linhas_dos_segredos_vazias(tmp_path):
    arq = tmp_path / ".env"
    nomes = ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": r"C:\Dados\Papaki"})
    assert nomes == ["PAPAKI_DADOS_DIR"]
    linhas = arq.read_text(encoding="utf-8").splitlines()
    assert "SUPABASE_URL=" in linhas and "SUPABASE_SERVICE_KEY=" in linhas and "SUPABASE_ACCESS_TOKEN=" in linhas
    assert r"PAPAKI_DADOS_DIR=C:\Dados\Papaki" in linhas


def test_atualiza_a_chave_existente_e_preserva_o_resto(tmp_path):
    arq = tmp_path / ".env"
    arq.write_text("# comentário\nSUPABASE_URL=https://abc.supabase.co\nPAPAKI_DADOS_DIR=C:\\Velha\nOUTRA=1\n",
                   encoding="utf-8")
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": r"C:\Nova", "PAPAKI_USUARIO_ESFERAS": "ana@exemplo.com"})
    assert arq.read_text(encoding="utf-8").splitlines() == [
        "# comentário", "SUPABASE_URL=https://abc.supabase.co", "PAPAKI_DADOS_DIR=C:\\Nova", "OUTRA=1",
        "PAPAKI_USUARIO_ESFERAS=ana@exemplo.com"]


def test_aceita_modelo_por_produtora_e_nomes_de_pasta(tmp_path):
    arq = tmp_path / ".env"
    ambiente.gravar_env(arq, {"PAPAKI_MODELO_UNTITLED": r"C:\M\Modelo.xlsx", "PAPAKI_PASTA_EXTRACOES": "Brutos",
                              "PAPAKI_PASTA_APURACOES": "DRE"})
    texto = arq.read_text(encoding="utf-8")
    for linha in (r"PAPAKI_MODELO_UNTITLED=C:\M\Modelo.xlsx", "PAPAKI_PASTA_EXTRACOES=Brutos",
                  "PAPAKI_PASTA_APURACOES=DRE"):
        assert linha in texto


@pytest.mark.parametrize("chave", ["SUPABASE_URL", "SUPABASE_SERVICE_KEY", "SUPABASE_ACCESS_TOKEN", "PATH",
                                   "PAPAKI_MODELO_", "PAPAKI_MODELO_x", "PAPAKI_OUTRA", "papaki_dados_dir"])
def test_recusa_chaves_fora_da_lista(tmp_path, chave):
    arq = tmp_path / ".env"
    with pytest.raises(ValueError, match="não pode ser gravada"):
        ambiente.gravar_env(arq, {chave: "x"})
    assert not arq.exists()


@pytest.mark.parametrize("valor", ["", " x", "x ", "a\nb", "a\rb", "a #b"])
def test_recusa_valores_ruins(tmp_path, valor):
    with pytest.raises(ValueError, match="valor inválido"):
        ambiente.gravar_env(tmp_path / ".env", {"PAPAKI_DADOS_DIR": valor})


def test_arquivo_que_nao_e_utf8_e_recusado_sem_alterar(tmp_path):
    arq = tmp_path / ".env"
    bruto = "# acentuação\nA=1\n".encode("cp1252")
    arq.write_bytes(bruto)
    with pytest.raises(ValueError, match="UTF-8"):
        ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\X"})
    assert arq.read_bytes() == bruto


def test_ler_arquivo_ignora_vazios(tmp_path):
    arq = tmp_path / ".env"
    arq.write_text("A=1\nB=\n", encoding="utf-8")
    assert ambiente.ler_arquivo(arq) == {"A": "1"}
    assert ambiente.ler_arquivo(tmp_path / "nao_existe") == {}


def test_arquivo_env_padrao_fica_na_raiz_do_projeto(tmp_path, monkeypatch):
    monkeypatch.setattr(ambiente, "RAIZ_PROJETO", tmp_path)
    assert ambiente.arquivo_env() == tmp_path / ".env"


# ---- regressões da revisão de segurança ----

SURROGATE = "a" + chr(0xD800) + "b"


@pytest.mark.parametrize("valor", ["a\u2028b", "a\u2029b", "a\x85b", "a\x0bb", "a\x0cb", "a\x1cb", "a\x1db",
                                   "a\x1eb", "a\tb", "a${X}b", "'abc'", '"abc"', "'abc", 'abc"', SURROGATE])
def test_recusa_valores_que_o_dotenv_leria_diferente_ou_quebram_linha(tmp_path, valor):
    arq = tmp_path / ".env"
    arq.write_text("SUPABASE_URL=https://abc.supabase.co\n", encoding="utf-8")
    antes = arq.read_bytes()
    with pytest.raises(ValueError, match="valor inválido"):
        ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": valor})
    assert arq.read_bytes() == antes


def test_ataque_em_duas_etapas_nao_injeta_supabase_url(tmp_path):
    from dotenv import dotenv_values
    arq = tmp_path / ".env"
    arq.write_text("SUPABASE_URL=https://abc.supabase.co\n", encoding="utf-8")
    for sep in ("\u2028", "\u2029", "\x85", "\x0b", "\x0c", "\x1c"):
        malicioso = "C:\\X" + sep + "SUPABASE_URL=https://evil.example"
        with pytest.raises(ValueError):
            ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": malicioso})
        ambiente.gravar_env(arq, {"PAPAKI_USUARIO_ESFERAS": "ana@exemplo.com"})
        assert dotenv_values(arq)["SUPABASE_URL"] == "https://abc.supabase.co"
    assert "evil" not in arq.read_text(encoding="utf-8")


def test_arquivo_com_separador_unicode_preexistente_nao_e_quebrado(tmp_path):
    arq = tmp_path / ".env"
    arq.write_bytes("# nota\u2028SUPABASE_URL=x\nPAPAKI_DADOS_DIR=C:\\Velha\n".encode("utf-8"))
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert arq.read_bytes() == "# nota\u2028SUPABASE_URL=x\nPAPAKI_DADOS_DIR=C:\\Nova\n".encode("utf-8")


def test_preserva_crlf_e_bytes_do_resto(tmp_path):
    arq = tmp_path / ".env"
    arq.write_bytes(b"# c\r\nA=1\r\nPAPAKI_DADOS_DIR=C:\\Velha\r\nB=2\r\n")
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova", "PAPAKI_USUARIO_ESFERAS": "ana@exemplo.com"})
    assert arq.read_bytes() == (b"# c\r\nA=1\r\nPAPAKI_DADOS_DIR=C:\\Nova\r\nB=2\r\n"
                                b"PAPAKI_USUARIO_ESFERAS=ana@exemplo.com\r\n")


def test_arquivo_sem_quebra_final_recebe_chave_em_linha_propria(tmp_path):
    arq = tmp_path / ".env"
    arq.write_bytes(b"A=1")
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\X"})
    assert arq.read_bytes() == b"A=1\nPAPAKI_DADOS_DIR=C:\\X\n"


def test_surrogate_solto_nao_zera_o_arquivo(tmp_path):
    arq = tmp_path / ".env"
    bruto = b"A=1\nPAPAKI_DADOS_DIR=C:\\Velha\n"
    arq.write_bytes(bruto)
    with pytest.raises(ValueError):
        ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\x" + chr(0xD800)})
    assert arq.read_bytes() == bruto


def test_falha_na_troca_mantem_original_e_remove_temporario(tmp_path, monkeypatch):
    arq = tmp_path / ".env"
    bruto = b"A=1\nPAPAKI_DADOS_DIR=C:\\Velha\n"
    arq.write_bytes(bruto)

    def falha(origem, destino):
        raise PermissionError(13, "negado", str(destino))

    monkeypatch.setattr(ambiente.os, "replace", falha)
    with pytest.raises(ValueError) as e:
        ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert str(tmp_path) not in str(e.value) and "Nova" not in str(e.value)
    assert arq.read_bytes() == bruto
    assert [p.name for p in tmp_path.iterdir()] == [".env"]


@pytest.mark.parametrize("onde", ["fsync", "replace"])
def test_interrupcao_do_usuario_mantem_original_e_remove_temporario(tmp_path, monkeypatch, onde):
    arq = tmp_path / ".env"
    bruto = b"A=1\nPAPAKI_DADOS_DIR=C:\\Velha\n"
    arq.write_bytes(bruto)

    def interrompe(*args, **kw):
        raise KeyboardInterrupt

    monkeypatch.setattr(ambiente.os, onde, interrompe)
    with pytest.raises(KeyboardInterrupt):
        ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert arq.read_bytes() == bruto
    assert [p.name for p in tmp_path.iterdir()] == [".env"]


def test_forca_a_gravacao_em_disco_antes_da_troca(tmp_path, monkeypatch):
    arq = tmp_path / ".env"
    ordem = []
    fsync, replace = ambiente.os.fsync, ambiente.os.replace
    monkeypatch.setattr(ambiente.os, "fsync", lambda fd: ordem.append("fsync") or fsync(fd))
    monkeypatch.setattr(ambiente.os, "replace", lambda a, b: ordem.append("replace") or replace(a, b))
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert ordem == ["fsync", "replace"]


def test_atualiza_todas_as_ocorrencias_da_chave(tmp_path):
    from dotenv import dotenv_values
    arq = tmp_path / ".env"
    arq.write_text("PAPAKI_DADOS_DIR=C:\\Um\nX=1\nPAPAKI_DADOS_DIR=C:\\Dois\n", encoding="utf-8")
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert arq.read_text(encoding="utf-8") == "PAPAKI_DADOS_DIR=C:\\Nova\nX=1\nPAPAKI_DADOS_DIR=C:\\Nova\n"
    assert dotenv_values(arq)["PAPAKI_DADOS_DIR"] == "C:\\Nova"


def test_reconhece_export_e_mantem_o_prefixo(tmp_path):
    arq = tmp_path / ".env"
    arq.write_text("export PAPAKI_DADOS_DIR=C:\\Velha\n", encoding="utf-8")
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert arq.read_text(encoding="utf-8") == "export PAPAKI_DADOS_DIR=C:\\Nova\n"


def test_reconhece_chave_na_primeira_linha_com_bom(tmp_path):
    arq = tmp_path / ".env"
    arq.write_bytes("\ufeffPAPAKI_DADOS_DIR=C:\\Velha\nA=1\n".encode("utf-8"))
    ambiente.gravar_env(arq, {"PAPAKI_DADOS_DIR": "C:\\Nova"})
    assert arq.read_bytes() == "\ufeffPAPAKI_DADOS_DIR=C:\\Nova\nA=1\n".encode("utf-8")
