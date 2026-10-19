from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
BRUTO = ROOT / "amostras" / "2026-09" / "bruto"
TEMPLATE = ROOT / "Modelo DRE - Untitled - base.xlsx"

REAL_FILES = {
    "cp": BRUTO / "CP_papaki_consolidado_2026-01-01_a_2026-09-30.csv",
    "cr": BRUTO / "CR_untitled_UTLD_2026-01-01_a_2026-09-30.csv",
    "comissoes": BRUTO / "COMISSOES_untitled_UTLD_2026-01-01_a_2026-09-30.xlsx",
    "rra": BRUTO / "Relatório Rentabilidade.xls",
}


def _real(name):
    path = REAL_FILES[name]
    if not path.exists():
        pytest.skip(f"bruto de referência ausente: {path}")
    return path


@pytest.fixture
def cp_real():
    return _real("cp")


@pytest.fixture
def cr_real():
    return _real("cr")


@pytest.fixture
def comissoes_real():
    return _real("comissoes")


@pytest.fixture
def rra_real():
    return _real("rra")


@pytest.fixture
def template_real():
    if not TEMPLATE.exists():
        pytest.skip(f"modelo ausente: {TEMPLATE}")
    return TEMPLATE


@pytest.fixture(autouse=True)
def _ambiente_isolado(monkeypatch, tmp_path_factory):
    """Nenhum teste lê o `.env` nem as variáveis PAPAKI_*/SUPABASE_* reais da máquina (o Plano 5 passou a
    ler nomes de pastas do ambiente). Fica de fora o teste ao vivo (PAPAKI_LIVE=1), que usa o `.env` real."""
    import os

    if os.environ.get("PAPAKI_LIVE") == "1":
        return
    from papaki_dre import ambiente

    for nome in list(os.environ):
        if nome.startswith(("PAPAKI_", "SUPABASE_")) and nome not in ("PAPAKI_DEBUG",):
            monkeypatch.delenv(nome)
    monkeypatch.setattr(ambiente, "RAIZ_PROJETO", tmp_path_factory.mktemp("sem_env"))
