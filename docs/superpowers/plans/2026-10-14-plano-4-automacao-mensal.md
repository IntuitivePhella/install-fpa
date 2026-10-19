# Plano 4 — Automação mensal da extração Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rodar todo mês (2º dia útil) e sob demanda o caminho Esferas (Chrome) → arquivos brutos → `importar`, com o Claude dirigindo o Chrome por roteiros e o código impondo as travas: um disparo por relatório, Localidade/datas/Status conferidos antes de exportar, e-mail do RRA escolhido sem ambiguidade, banco pausado acordado sozinho, pastas definidas pelo ambiente.

**Architecture:** Peças pequenas e testáveis (`calendario`, `ambiente`/`caminhos`, `supabase_admin`, `rra_email`, `extracao`) mais um módulo de CLI (`cli_extracao`: `agenda` e `extracao ...`). O `importar` ganha `--extracao <pasta>` (lê o registro fechado da extração), `--restaurar-banco`, `--modelo`/`--saida` opcionais, data de extração com fuso e limpeza de rodadas órfãs. Dois roteiros em `runbooks/` (extração e rotina mensal) são o texto que o Claude segue; a tarefa agendada só manda ele seguir o roteiro mensal.

**Tech Stack:** Python ≥ 3.11, httpx (API de gerenciamento do Supabase), supabase-py, python-dotenv, pytest. Windows + Chrome (extensão do Claude) para o que é visual.

**Spec:** `docs/superpowers/specs/2026-10-14-plano-4-automacao-mensal-design.md` (aprovada em 14/10/2026). Decisões 1–13 do `docs/superpowers/HANDOFF-execucao-planos.md` continuam valendo.

**Regras inegociáveis (valem em todas as tasks):**
- Comandos Python: `py -m pytest ...` e `py -m papaki_dre ...` (o `python` do PATH não funciona nesta máquina).
- Dados sensíveis (CPF/CNPJ/dados bancários): nunca em respostas, logs ou testes novos; só dados inventados.
- O `.env` **nunca** é aberto nem impresso. O token (`SUPABASE_ACCESS_TOKEN`) e a chave de serviço nunca aparecem em mensagem, log ou teste (nos testes entra um valor falso).
- Excel: nunca encerrar `EXCEL.EXE` que você não iniciou; os testes `real`/`excel` usam `com_retry`.
- Gravar no Supabase **com dados reais** só com o "pode importar" do usuário. Acordar o projeto pausado de verdade, criar a tarefa agendada e capturar a tela do usuário só com autorização dele.
- Mensagem de erro nova passa por `mensagem_segura` (sem URL, ref, chave nem linha de dados).
- Commits terminam com a linha `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>` (use um segundo `-m`).
- Ao fim de cada task, diga ao usuário em 2–3 linhas, em português simples: o que passou, o que não passou, o que mudou para ele.

**Linha de base (verifique antes de começar):** `py -m pytest tests -q -m "not real and not excel"` → 401 passam. (`py -m pytest tests/test_cli_real.py tests/test_build_real.py` → 20 passam em ~4 min, usa Excel; rode só na Task 11.)

---

## Mapa de arquivos

| Arquivo | Ação | Responsabilidade |
|---|---|---|
| `src/papaki_dre/config.py` | modificar | `FUSO_BRASILIA`, `FERIADOS_EXTRAS`, `NOME_PASTA_EXTRACOES`, `NOME_PASTA_APURACOES` |
| `src/papaki_dre/calendario.py` | criar | dias úteis, 2º dia útil, competência alvo, `deve_rodar`, `competencia_concluida`, `periodo_da_competencia` |
| `src/papaki_dre/ambiente.py` | criar | lê variáveis do Windows + `.env` (o `.env` prevalece) |
| `src/papaki_dre/caminhos.py` | criar | raiz de dados, pastas de extração e apuração, modelo por produtora |
| `src/papaki_dre/supabase_admin.py` | criar | estado e restauração do projeto pela API de gerenciamento |
| `src/papaki_dre/rra_email.py` | criar | escolha da mensagem do RRA (função pura) |
| `src/papaki_dre/extracao.py` | criar | registro da extração, travas, manifesto |
| `src/papaki_dre/cli_extracao.py` | criar | subcomandos `agenda` e `extracao ...` |
| `src/papaki_dre/cli.py` | modificar | `importar --extracao`, `--restaurar-banco`, opcionais, fuso, órfãs, RESULTADO.md |
| `src/papaki_dre/db.py` | modificar | `_secrets` inclui o token; `Repository.encerrar_orfas` |
| `src/papaki_dre/report.py` | modificar | `render_resultado` |
| `scripts/capturar_tela.ps1` | criar | print da tela inteira para a prova de Localidade |
| `runbooks/extracao-esferas.md`, `runbooks/rotina-mensal.md` | criar | roteiros para o Claude |
| `tests/extracao_factory.py` | criar | extração fechada sintética para os testes |
| `tests/test_calendario.py`, `test_ambiente_caminhos.py`, `test_supabase_admin.py`, `test_rra_email.py`, `test_extracao.py`, `test_cli_extracao.py`, `test_cli_rotina.py` | criar | testes de cada peça |
| `tests/test_cli.py`, `tests/test_db.py` | modificar | ajustes pontuais |
| `pyproject.toml`, `.env.example`, `README.md`, `docs/superpowers/HANDOFF-execucao-planos.md` | modificar | dependência, exemplos, documentação |

---

### Task 1: Calendário (2º dia útil) e constantes de configuração

**Files:**
- Modify: `src/papaki_dre/config.py`
- Create: `src/papaki_dre/calendario.py`
- Test: `tests/test_calendario.py`

- [ ] **Step 1: Constantes novas em `config.py`**

No topo de `src/papaki_dre/config.py`, depois do docstring de uma linha, acrescente `import datetime as dt`; no fim do arquivo acrescente:

```python
# Fuso de trabalho (Brasília, sem horário de verão desde 2019). Datas sem fuso valem neste fuso.
FUSO_BRASILIA = dt.timezone(dt.timedelta(hours=-3))

# Feriados além dos nacionais (ISO AAAA-MM-DD). A Papaki acrescenta os locais aqui.
FERIADOS_EXTRAS: tuple[str, ...] = ()

# Nomes das subpastas sob PAPAKI_DADOS_DIR (a instalação pode trocá-los por nomes já usados na Papaki).
NOME_PASTA_EXTRACOES = "extracoes"
NOME_PASTA_APURACOES = "apuracoes"
```

- [ ] **Step 2: Escrever os testes que falham**

Crie `tests/test_calendario.py`:

```python
import datetime as dt
import json

import pytest

from papaki_dre import calendario as c

D = dt.date


def test_pascoa_de_anos_conhecidos():
    assert c.pascoa(2025) == D(2025, 4, 20)
    assert c.pascoa(2026) == D(2026, 4, 5)
    assert c.pascoa(2027) == D(2027, 3, 28)


def test_sexta_feira_santa_e_feriado():
    assert D(2026, 4, 3) in c.feriados(2026)
    assert not c.eh_dia_util(D(2026, 4, 3))


def test_consciencia_negra_so_a_partir_de_2024():
    assert D(2024, 11, 20) in c.feriados(2024)
    assert D(2023, 11, 20) not in c.feriados(2023)


def test_carnaval_e_corpus_christi_contam_como_dia_util():
    assert c.eh_dia_util(D(2026, 2, 17))
    assert c.eh_dia_util(D(2026, 6, 4))


def test_feriado_extra():
    assert c.eh_dia_util(D(2026, 7, 9))
    assert not c.eh_dia_util(D(2026, 7, 9), extras=("2026-07-09",))


def test_novembro_de_2026_o_segundo_dia_util_e_dia_4():
    # dia 1 é domingo, dia 2 é Finados
    assert c.dias_uteis(2026, 11)[:3] == [D(2026, 11, 3), D(2026, 11, 4), D(2026, 11, 5)]
    assert c.enesimo_dia_util(2026, 11, 2) == D(2026, 11, 4)


def test_janeiro_de_2027_e_a_competencia_de_dezembro():
    assert c.enesimo_dia_util(2027, 1, 1) == D(2027, 1, 4)
    assert c.competencia_alvo(D(2027, 1, 5)) == "2026-12"


def test_competencia_alvo_e_o_mes_anterior():
    assert c.competencia_alvo(D(2026, 11, 4)) == "2026-10"


def test_deve_rodar_no_segundo_dia_util():
    assert c.deve_rodar(D(2026, 11, 4), ja_importada=False) == (True, "")


def test_nao_roda_no_primeiro_dia_util():
    rodar, motivo = c.deve_rodar(D(2026, 11, 3), ja_importada=False)
    assert not rodar and "2º dia útil" in motivo


def test_roda_ate_o_quinto_dia_util_e_depois_nao():
    # novembro/2026: 3, 4, 5, 6, 9 são os cinco primeiros dias úteis
    assert c.deve_rodar(D(2026, 11, 9), ja_importada=False)[0] is True
    rodar, motivo = c.deve_rodar(D(2026, 11, 10), ja_importada=False)
    assert not rodar and "5º dia útil" in motivo


def test_fim_de_semana_e_feriado_nao_rodam():
    assert c.deve_rodar(D(2026, 11, 7), ja_importada=False)[0] is False
    assert c.deve_rodar(D(2026, 11, 2), ja_importada=False)[0] is False


def test_competencia_ja_importada_nao_roda():
    rodar, motivo = c.deve_rodar(D(2026, 11, 4), ja_importada=True)
    assert not rodar and "já importada" in motivo


def test_periodo_da_competencia():
    assert c.periodo_da_competencia("2026-09") == ("2026-09", D(2026, 1, 1), D(2026, 9, 30))
    assert c.periodo_da_competencia("2024-02")[2] == D(2024, 2, 29)
    for ruim in ("2026-13", "2026/09", "setembro", "", None):
        with pytest.raises(ValueError):
            c.periodo_da_competencia(ruim)


def _conf(tmp_path, **kw):
    conf = {"status": "ok", "meta": {"competencia": "2026-09", "periodo": "2026-01-01 a 2026-09-30",
                                       "importacao_id": "abc"}}
    conf.update(kw.pop("topo", {}))
    conf["meta"].update(kw)
    p = tmp_path / "conferencia.json"
    p.write_text(json.dumps(conf), encoding="utf-8")
    return p


def test_competencia_concluida(tmp_path):
    assert c.competencia_concluida(_conf(tmp_path), "2026-09") is True
    assert c.competencia_concluida(_conf(tmp_path, topo={"status": "com_alertas"}), "2026-09") is True


@pytest.mark.parametrize("mudanca", [
    {"topo": {"status": "bloqueada"}},
    {"topo": {"status": "falhou"}},
    {"importacao_id": None},                       # rodada sem banco
    {"competencia": "2026-08"},
    {"periodo": "2026-01-01 a 2026-09-15"},        # demonstrativo parcial
])
def test_competencia_nao_concluida(tmp_path, mudanca):
    assert c.competencia_concluida(_conf(tmp_path, **mudanca), "2026-09") is False


def test_competencia_concluida_sem_arquivo_ou_ilegivel(tmp_path):
    assert c.competencia_concluida(tmp_path / "nao-existe.json", "2026-09") is False
    ruim = tmp_path / "x.json"
    ruim.write_text("{não é json", encoding="utf-8")
    assert c.competencia_concluida(ruim, "2026-09") is False
```

- [ ] **Step 3: Rodar e ver falhar**

Run: `py -m pytest tests/test_calendario.py -q`
Expected: FAIL (`ImportError` / `cannot import name 'calendario'`).

- [ ] **Step 4: Implementar `src/papaki_dre/calendario.py`**

```python
"""Calendário da rotina mensal: dias úteis (segunda a sexta, sem feriados nacionais), 2º dia útil,
competência alvo e a decisão "hoje é dia de rodar?"."""
from __future__ import annotations

import calendar
import datetime as dt
import json
import re
from pathlib import Path

# (mês, dia) dos feriados nacionais fixos. Carnaval e Corpus Christi NÃO entram (decisão do usuário).
FERIADOS_FIXOS = ((1, 1), (4, 21), (5, 1), (9, 7), (10, 12), (11, 2), (11, 15), (12, 25))
CONSCIENCIA_NEGRA = (11, 20)        # feriado nacional desde 2024 (Lei 14.759/2023)
PRIMEIRO_DIA_UTIL = 2
ULTIMO_DIA_UTIL = 5                 # depois disso a rotina não roda sozinha (refazer é manual)


def pascoa(ano: int) -> dt.date:
    """Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return dt.date(ano, mes, dia)


def feriados(ano: int, extras=()) -> set[dt.date]:
    out = {dt.date(ano, m, d) for m, d in FERIADOS_FIXOS}
    if ano >= 2024:
        out.add(dt.date(ano, *CONSCIENCIA_NEGRA))
    out.add(pascoa(ano) - dt.timedelta(days=2))                 # Sexta-feira Santa
    out |= {dt.date.fromisoformat(x) for x in extras if x.startswith(f"{ano}-")}
    return out


def eh_dia_util(dia: dt.date, extras=()) -> bool:
    return dia.weekday() < 5 and dia not in feriados(dia.year, extras)


def dias_uteis(ano: int, mes: int, extras=()) -> list[dt.date]:
    ultimo = calendar.monthrange(ano, mes)[1]
    return [d for d in (dt.date(ano, mes, n) for n in range(1, ultimo + 1)) if eh_dia_util(d, extras)]


def enesimo_dia_util(ano: int, mes: int, n: int, extras=()) -> dt.date:
    uteis = dias_uteis(ano, mes, extras)
    if not 1 <= n <= len(uteis):
        raise ValueError(f"o mês {ano}-{mes:02d} não tem {n}º dia útil")
    return uteis[n - 1]


def competencia_alvo(hoje: dt.date) -> str:
    """Competência que a rotina de `hoje` apura: o mês anterior (AAAA-MM)."""
    anterior = hoje.replace(day=1) - dt.timedelta(days=1)
    return f"{anterior.year}-{anterior.month:02d}"


def deve_rodar(hoje: dt.date, ja_importada: bool, extras=()) -> tuple[bool, str]:
    """(rodar?, motivo se não). Roda do 2º ao 5º dia útil do mês, e só se a competência anterior ainda
    não foi importada (assim uma falha é refeita no dia útil seguinte)."""
    if not eh_dia_util(hoje, extras):
        return False, "hoje não é dia útil"
    posicao = dias_uteis(hoje.year, hoje.month, extras).index(hoje) + 1
    if posicao < PRIMEIRO_DIA_UTIL:
        return False, f"ainda não é o {PRIMEIRO_DIA_UTIL}º dia útil"
    if posicao > ULTIMO_DIA_UTIL:
        return False, f"já passou do {ULTIMO_DIA_UTIL}º dia útil (rode manualmente se preciso)"
    if ja_importada:
        return False, "competência já importada"
    return True, ""


def periodo_da_competencia(texto) -> tuple[str, dt.date, dt.date]:
    """('AAAA-MM', 1º de janeiro, último dia do mês). ValueError se o texto não for AAAA-MM."""
    m = re.fullmatch(r"(\d{4})-(0[1-9]|1[0-2])", texto or "")
    if not m:
        raise ValueError(f"--competencia deve ser AAAA-MM; recebido {texto!r}")
    ano, mes = int(m.group(1)), int(m.group(2))
    return texto, dt.date(ano, 1, 1), dt.date(ano, mes, calendar.monthrange(ano, mes)[1])


def competencia_concluida(caminho_conferencia, competencia: str) -> bool:
    """A pasta de apuração já tem a conferência de uma importação REAL (com id no banco) `ok` ou
    `com_alertas` da competência, com o período até o último dia do mês."""
    try:
        conf = json.loads(Path(caminho_conferencia).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    if not isinstance(conf, dict):
        return False
    meta = conf.get("meta") or {}
    if conf.get("status") not in ("ok", "com_alertas") or not meta.get("importacao_id"):
        return False
    if meta.get("competencia") != competencia:
        return False
    _, _, fim = periodo_da_competencia(competencia)
    return str(meta.get("periodo", "")).endswith(fim.isoformat())
```

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests/test_calendario.py -q`
Expected: todos passam.

- [ ] **Step 6: `cli.parse_competencia` passa a usar o calendário (sem duplicar a regra)**

Em `src/papaki_dre/cli.py`, acrescente `from . import calendario` junto dos imports e troque o corpo de `parse_competencia` (e remova os imports `calendar` e `re` se ficarem sem uso):

```python
def parse_competencia(text: str) -> tuple[str, dt.date, dt.date]:
    try:
        return calendario.periodo_da_competencia(text)
    except ValueError as e:
        raise UsageError(str(e)) from None
```

Run: `py -m pytest tests -q -m "not real and not excel"`
Expected: 401 + os novos passam.

- [ ] **Step 7: Commit**

```bash
git add src/papaki_dre/config.py src/papaki_dre/calendario.py src/papaki_dre/cli.py tests/test_calendario.py
git commit -m "feat: calendário da rotina mensal (2º dia útil, competência alvo, deve_rodar)" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Ambiente e caminhos (pastas definidas pelo ambiente)

**Files:**
- Create: `src/papaki_dre/ambiente.py`, `src/papaki_dre/caminhos.py`
- Test: `tests/test_ambiente_caminhos.py`

- [ ] **Step 1: Escrever os testes que falham**

Crie `tests/test_ambiente_caminhos.py`:

```python
from pathlib import Path

import pytest

from papaki_dre import ambiente, caminhos


def test_resolver_devolve_o_dicionario_dado():
    env = {"A": "1"}
    assert ambiente.resolver(env) is env


def test_ler_env_o_dotenv_prevalece_e_vazio_nao_conta(monkeypatch):
    import dotenv
    monkeypatch.setenv("A", "do-windows")
    monkeypatch.setenv("B", "do-windows")
    monkeypatch.setattr(dotenv, "dotenv_values", lambda *a, **k: {"A": "do-env", "B": "", "C": None})
    env = ambiente.ler_env()
    assert env["A"] == "do-env" and env["B"] == "do-windows" and "C" not in env


def test_raiz_dados_exige_a_variavel():
    with pytest.raises(caminhos.CaminhosErro, match="PAPAKI_DADOS_DIR"):
        caminhos.raiz_dados({})


def test_raiz_dados_nao_pode_ficar_dentro_do_projeto():
    dentro = caminhos.RAIZ_PROJETO / "dados"
    with pytest.raises(caminhos.CaminhosErro, match="dentro da pasta do projeto"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(dentro)})
    with pytest.raises(caminhos.CaminhosErro, match="dentro da pasta do projeto"):
        caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(caminhos.RAIZ_PROJETO)})


def test_raiz_dados_fora_do_projeto(tmp_path):
    assert caminhos.raiz_dados({"PAPAKI_DADOS_DIR": str(tmp_path)}) == tmp_path.resolve()


def test_pastas_padrao(tmp_path):
    raiz = tmp_path
    assert caminhos.pasta_extracoes(raiz, "2026-09", "untitled") == raiz / "extracoes" / "2026-09" / "untitled"
    assert caminhos.pasta_apuracao(raiz, "untitled", "2026-09") == raiz / "apuracoes" / "untitled" / "2026-09"
    assert caminhos.rel_extracao("2026-09", "untitled", "20261104-0930") == \
        "extracoes/2026-09/untitled/20261104-0930"


def test_modelo_da_produtora(tmp_path):
    modelo = tmp_path / "Modelo DRE - Untitled - 2026.xlsx"
    modelo.write_bytes(b"x")
    assert caminhos.modelo_da_produtora("untitled", {"PAPAKI_MODELO_UNTITLED": str(modelo)}) == modelo
    with pytest.raises(caminhos.CaminhosErro, match="PAPAKI_MODELO_UNTITLED"):
        caminhos.modelo_da_produtora("untitled", {})
    with pytest.raises(caminhos.CaminhosErro, match="não existe"):
        caminhos.modelo_da_produtora("untitled", {"PAPAKI_MODELO_UNTITLED": str(tmp_path / "x.xlsx")})
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_ambiente_caminhos.py -q`
Expected: FAIL (`ImportError`).

- [ ] **Step 3: Implementar**

`src/papaki_dre/ambiente.py`:

```python
"""Variáveis de ambiente do sistema: as do Windows mais as do `.env` do projeto. O `.env` prevalece
(mesma regra de `db.make_client`: nunca mistura projetos) e valor vazio no `.env` conta como ausente."""
from __future__ import annotations

import os


def ler_env() -> dict[str, str]:
    from dotenv import dotenv_values

    env = dict(os.environ)
    for chave, valor in dotenv_values().items():
        if valor:
            env[chave] = valor
    return env


def resolver(env):
    """`env` dado (testes) ou o ambiente real."""
    return env if env is not None else ler_env()
```

`src/papaki_dre/caminhos.py`:

```python
"""Onde ficam os arquivos. Uma só peça conhece a estrutura de pastas; o resto do código pede o caminho
a ela. A raiz vem de PAPAKI_DADOS_DIR (definida na instalação, fora do projeto) e o modelo base de cada
produtora de PAPAKI_MODELO_<PRODUTORA> (até o Plano 5 trazer um arquivo de configuração)."""
from __future__ import annotations

from pathlib import Path

from .config import NOME_PASTA_APURACOES, NOME_PASTA_EXTRACOES

RAIZ_PROJETO = Path(__file__).resolve().parents[2]


class CaminhosErro(ValueError):
    """Configuração de pastas ausente ou inválida."""


def raiz_dados(env, projeto: Path | None = None) -> Path:
    valor = (env or {}).get("PAPAKI_DADOS_DIR")
    if not valor:
        raise CaminhosErro("defina PAPAKI_DADOS_DIR no .env (pasta de dados, fora do projeto; veja .env.example)")
    raiz = Path(valor).expanduser().resolve()
    proj = (projeto or RAIZ_PROJETO).resolve()
    if raiz == proj or proj in raiz.parents:
        raise CaminhosErro("PAPAKI_DADOS_DIR não pode ficar dentro da pasta do projeto: "
                           "os dados reais nunca vão para o repositório")
    return raiz


def pasta_extracoes(raiz: Path, competencia: str, produtora: str) -> Path:
    return Path(raiz) / NOME_PASTA_EXTRACOES / competencia / produtora


def pasta_apuracao(raiz: Path, produtora: str, competencia: str) -> Path:
    return Path(raiz) / NOME_PASTA_APURACOES / produtora / competencia


def rel_extracao(competencia: str, produtora: str, nome_pasta: str) -> str:
    """Caminho da pasta de uma extração relativo à raiz de dados (o que vai para `prova_caminho`)."""
    return f"{NOME_PASTA_EXTRACOES}/{competencia}/{produtora}/{nome_pasta}"


def modelo_da_produtora(produtora: str, env) -> Path:
    var = f"PAPAKI_MODELO_{produtora.upper()}"
    valor = (env or {}).get(var)
    if not valor:
        raise CaminhosErro(f"defina {var} no .env (caminho do Modelo DRE base anual da produtora)")
    caminho = Path(valor).expanduser()
    if not caminho.is_file():
        raise CaminhosErro(f"o arquivo de {var} não existe: {caminho.name}")
    return caminho
```

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_ambiente_caminhos.py -q`
Expected: todos passam.

- [ ] **Step 5: Exemplos no `.env.example`**

Acrescente ao fim de `.env.example` (sem valores reais):

```
# Pasta de dados (modelos, extrações, resultados). Fora do projeto; definida na instalação.
PAPAKI_DADOS_DIR=<pasta de dados — nunca dentro do projeto>
# Modelo DRE base anual de cada produtora (o FP&A mantém o arquivo no mesmo local).
PAPAKI_MODELO_UNTITLED=<caminho do Modelo DRE - Untitled - base do ano>
```

- [ ] **Step 6: Commit**

```bash
git add src/papaki_dre/ambiente.py src/papaki_dre/caminhos.py tests/test_ambiente_caminhos.py .env.example
git commit -m "feat: ambiente e caminhos (raiz de dados e modelo por produtora definidos pelo ambiente)" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Acordar o projeto Supabase pausado (`supabase_admin`)

**Files:**
- Create: `src/papaki_dre/supabase_admin.py`
- Modify: `src/papaki_dre/db.py` (`_secrets`), `pyproject.toml`
- Test: `tests/test_supabase_admin.py`

- [ ] **Step 1: Escrever os testes que falham**

Crie `tests/test_supabase_admin.py` (token e ref são valores inventados):

```python
import pytest

from papaki_dre.db import mensagem_segura
from papaki_dre.supabase_admin import AdminErro, AdminSupabase, ref_do_projeto

TOKEN = "sbp_valor_falso_para_teste_123"
REF = "refinventadoabcdefgh"


class Resp:
    def __init__(self, status=200, corpo=None):
        self.status_code = status
        self._corpo = corpo or {}

    def json(self):
        return self._corpo


class FakeHttp:
    """GET devolve os estados em sequência (o último se repete); POST devolve `post_status`."""

    def __init__(self, estados, post_status=200, get_status=200, erro_rede=None):
        self.estados, self.post_status, self.get_status, self.erro_rede = list(estados), post_status, get_status, erro_rede
        self.chamadas = []

    def request(self, metodo, caminho, headers=None):
        self.chamadas.append((metodo, caminho, headers))
        if self.erro_rede:
            raise self.erro_rede
        if metodo == "GET":
            estado = self.estados.pop(0) if len(self.estados) > 1 else self.estados[0]
            return Resp(self.get_status, {"status": estado})
        return Resp(self.post_status, {})


class Relogio:
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        return self.t

    def dormir(self, s):
        self.t += s


def _admin(http, relogio=None):
    relogio = relogio or Relogio()
    return AdminSupabase(TOKEN, REF, http=http, dormir=relogio.dormir, relogio=relogio), relogio


def _posts(http):
    return [c for c in http.chamadas if c[0] == "POST"]


def test_ref_do_projeto():
    assert ref_do_projeto(f"https://{REF}.supabase.co") == REF
    with pytest.raises(AdminErro):
        ref_do_projeto("")


def test_ja_ativo_nao_restaura():
    http = FakeHttp(["ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "ja_ativo"
    assert _posts(http) == []


def test_pausado_restaura_uma_vez_e_espera_ficar_ativo():
    http = FakeHttp(["INACTIVE", "INACTIVE", "RESTORING", "COMING_UP", "ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "restaurado"
    assert len(_posts(http)) == 1
    assert _posts(http)[0][1] == f"/v1/projects/{REF}/restore"
    assert http.chamadas[0][2] == {"Authorization": f"Bearer {TOKEN}"}


def test_restaurando_so_espera_sem_pedir_de_novo():
    http = FakeHttp(["RESTORING", "ACTIVE_HEALTHY"])
    admin, _ = _admin(http)
    assert admin.acordar() == "aguardado"
    assert _posts(http) == []


def test_tempo_esgotado():
    http = FakeHttp(["INACTIVE", "RESTORING"])
    admin, relogio = _admin(http)
    with pytest.raises(AdminErro, match="não ficou ativo em 10 min"):
        admin.acordar(espera_max=600, intervalo=15)
    assert relogio.t >= 600


@pytest.mark.parametrize("estado", ["RESTORE_FAILED", "INIT_FAILED", "REMOVED"])
def test_estados_de_falha_param_na_hora(estado):
    http = FakeHttp(["INACTIVE", estado])
    admin, _ = _admin(http)
    with pytest.raises(AdminErro, match=estado):
        admin.acordar()


@pytest.mark.parametrize("status", [401, 403])
def test_chave_recusada_nao_vaza_nada(status):
    admin, _ = _admin(FakeHttp(["INACTIVE"], get_status=status))
    with pytest.raises(AdminErro) as e:
        admin.acordar()
    assert str(status) in str(e.value) and TOKEN not in str(e.value) and REF not in str(e.value)


def test_erro_de_rede_so_diz_o_tipo():
    admin, _ = _admin(FakeHttp(["x"], erro_rede=ConnectionError(f"falhou em https://{REF}.supabase.co {TOKEN}")))
    with pytest.raises(AdminErro) as e:
        admin.estado()
    assert "ConnectionError" in str(e.value) and TOKEN not in str(e.value) and REF not in str(e.value)


def test_erro_http_do_restore():
    admin, _ = _admin(FakeHttp(["INACTIVE"], post_status=500))
    with pytest.raises(AdminErro, match="HTTP 500"):
        admin.acordar()


def test_token_nunca_aparece_em_mensagem_segura(monkeypatch):
    monkeypatch.setenv("SUPABASE_ACCESS_TOKEN", TOKEN)
    assert TOKEN not in mensagem_segura(RuntimeError(f"erro com {TOKEN} no meio"))
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_supabase_admin.py -q`
Expected: FAIL (`ModuleNotFoundError: papaki_dre.supabase_admin`).

- [ ] **Step 3: Implementar `src/papaki_dre/supabase_admin.py`**

```python
"""Estado e restauração do projeto Supabase pela API de gerenciamento (plano gratuito: o projeto pausa
após ~7 dias sem uso). Usa SUPABASE_ACCESS_TOKEN (chave pessoal, só no .env). As mensagens de erro nunca
levam o token, o ref do projeto nem a URL."""
from __future__ import annotations

import time
from urllib.parse import urlparse

import httpx

API = "https://api.supabase.com"
ATIVO = "ACTIVE_HEALTHY"
PAUSADO = {"INACTIVE", "PAUSED"}
FALHA = {"RESTORE_FAILED", "INIT_FAILED", "PAUSE_FAILED", "REMOVED"}


class AdminErro(RuntimeError):
    """Falha ao consultar ou restaurar o projeto (mensagem já limpa)."""


def ref_do_projeto(url: str) -> str:
    ref = (urlparse(url or "").hostname or "").split(".")[0]
    if not ref:
        raise AdminErro("SUPABASE_URL inválida: não foi possível identificar o projeto")
    return ref


class AdminSupabase:
    def __init__(self, token: str, ref: str, *, http=None, dormir=time.sleep, relogio=time.monotonic):
        self.token, self.ref = token, ref
        self.http = http or httpx.Client(base_url=API, timeout=30)
        self.dormir, self.relogio = dormir, relogio

    def _chamar(self, metodo: str, caminho: str):
        try:
            r = self.http.request(metodo, caminho, headers={"Authorization": f"Bearer {self.token}"})
        except Exception as e:  # noqa: BLE001 - só o tipo: a mensagem do httpx traz URL
            raise AdminErro(f"API do Supabase inacessível ({type(e).__name__})") from None
        if r.status_code in (401, 403):
            raise AdminErro(f"a chave pessoal do Supabase foi recusada (HTTP {r.status_code})")
        if r.status_code >= 400:
            raise AdminErro(f"a API do Supabase respondeu HTTP {r.status_code}")
        return r

    def estado(self) -> str:
        corpo = self._chamar("GET", f"/v1/projects/{self.ref}").json()
        return str(corpo.get("status", "UNKNOWN"))

    def restaurar(self) -> None:
        self._chamar("POST", f"/v1/projects/{self.ref}/restore")

    def acordar(self, espera_max: int = 600, intervalo: int = 15) -> str:
        """'ja_ativo' (nada a fazer), 'restaurado' (pediu a restauração e esperou) ou 'aguardado'
        (já estava subindo). Só pede restauração se a API confirmar que o projeto está pausado."""
        estado = self.estado()
        if estado == ATIVO:
            return "ja_ativo"
        fim = self.relogio() + espera_max
        acao = "aguardado"
        if estado in PAUSADO:
            self.restaurar()
            acao = "restaurado"
        while True:
            estado = self.estado()
            if estado == ATIVO:
                return acao
            if estado in FALHA:
                raise AdminErro(f"o Supabase informou {estado}; restaure o projeto pelo painel")
            if self.relogio() >= fim:
                raise AdminErro(f"o projeto não ficou ativo em {espera_max // 60} min (último estado: {estado})")
            self.dormir(intervalo)
```

- [ ] **Step 4: O token entra na lista de segredos de `db.py`**

Em `src/papaki_dre/db.py`, na função `_secrets`, troque o bloco da chave:

```python
    key = os.environ.get("SUPABASE_SERVICE_KEY") or ""
    if key:
        out.append(key)
    return sorted((s for s in out if s), key=len, reverse=True)
```

por:

```python
    for nome in ("SUPABASE_SERVICE_KEY", "SUPABASE_ACCESS_TOKEN"):
        valor = os.environ.get(nome) or ""
        if valor:
            out.append(valor)
    return sorted((s for s in out if s), key=len, reverse=True)
```

- [ ] **Step 5: Dependência explícita**

Em `pyproject.toml`, na lista `dependencies`, acrescente `"httpx>=0.27",` (já vem junto do supabase-py; aqui fica explícita porque agora é usada direto).

- [ ] **Step 6: Rodar e ver passar**

Run: `py -m pytest tests/test_supabase_admin.py tests/test_db.py -q`
Expected: todos passam.

- [ ] **Step 7: Commit**

```bash
git add src/papaki_dre/supabase_admin.py src/papaki_dre/db.py pyproject.toml tests/test_supabase_admin.py
git commit -m "feat: acordar o projeto Supabase pausado pela API de gerenciamento" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Escolha da mensagem do RRA (`rra_email`)

**Files:**
- Modify: `src/papaki_dre/config.py` (nada; o fuso já existe)
- Create: `src/papaki_dre/rra_email.py`
- Test: `tests/test_rra_email.py`

- [ ] **Step 1: Escrever os testes que falham**

Crie `tests/test_rra_email.py`:

```python
import datetime as dt

import pytest

from papaki_dre.config import FUSO_BRASILIA
from papaki_dre.rra_email import RraEmailErro, SemMensagemNova, escolher_mensagem, horario_solicitado

DISPARO = dt.datetime(2026, 11, 4, 9, 30, 0, tzinfo=FUSO_BRASILIA)


def msg(id_, quando="04/11/2026 09:30:20", remetente="Papaki <notas@papaki.com.br>",
        assunto="Relatório de rentabilidade gerado com sucesso"):
    return {"id": id_, "remetente": remetente, "assunto": assunto,
            "corpo": f"O relatório solicitado em {quando} está anexo."}


def test_horario_solicitado():
    assert horario_solicitado("... solicitado em 04/11/2026 09:31:07 ...") == \
        dt.datetime(2026, 11, 4, 9, 31, 7, tzinfo=FUSO_BRASILIA)
    assert horario_solicitado("sem horário") is None
    assert horario_solicitado("solicitado em 31/02/2026 09:31:07") is None


def test_escolhe_a_unica_mensagem_nova_e_no_horario():
    assert escolher_mensagem([msg("a1")], ids_antes=[], disparo=DISPARO) == "a1"


def test_ignora_mensagem_que_ja_existia_antes_do_disparo():
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([msg("velha")], ids_antes=["velha"], disparo=DISPARO)


def test_ignora_remetente_e_assunto_errados():
    outras = [msg("b1", remetente="alguem@exemplo.com"), msg("b2", assunto="Outro assunto")]
    with pytest.raises(SemMensagemNova):
        escolher_mensagem(outras, ids_antes=[], disparo=DISPARO)


@pytest.mark.parametrize("quando", ["04/11/2026 09:29:50", "04/11/2026 09:30:00", "04/11/2026 09:33:00"])
def test_horarios_dentro_da_janela(quando):
    assert escolher_mensagem([msg("c1", quando)], ids_antes=[], disparo=DISPARO) == "c1"


@pytest.mark.parametrize("quando", ["04/11/2026 09:29:49", "04/11/2026 09:33:01", "03/11/2026 09:30:00"])
def test_horarios_fora_da_janela(quando):
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([msg("d1", quando)], ids_antes=[], disparo=DISPARO)


def test_duas_candidatas_validas_param_sem_escolher():
    with pytest.raises(RraEmailErro, match="mais de uma") as e:
        escolher_mensagem([msg("e1"), msg("e2", "04/11/2026 09:30:40")], ids_antes=[], disparo=DISPARO)
    assert not isinstance(e.value, SemMensagemNova)


def test_mensagem_sem_horario_no_corpo_nao_serve():
    sem = {"id": "f1", "remetente": "notas@papaki.com.br",
           "assunto": "Relatório de rentabilidade gerado com sucesso", "corpo": "sem horário"}
    with pytest.raises(SemMensagemNova):
        escolher_mensagem([sem], ids_antes=[], disparo=DISPARO)
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_rra_email.py -q`
Expected: FAIL (`ModuleNotFoundError`).

- [ ] **Step 3: Implementar `src/papaki_dre/rra_email.py`**

```python
"""Escolha da mensagem do RRA entre as que o Claude leu no Gmail. Função pura: nada de rede."""
from __future__ import annotations

import datetime as dt
import re

from .config import FUSO_BRASILIA

REMETENTE = "notas@papaki.com.br"
ASSUNTO = "Relatório de rentabilidade gerado com sucesso"
ANTES_S = 10          # o "solicitado em" do Esferas pode ficar até 10 s antes do disparo registrado
DEPOIS_S = 180        # o disparo é gravado ANTES do clique; o pedido real vem até 3 min depois
_SOLICITADO = re.compile(r"solicitad[oa] em\s*(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", re.I)


class RraEmailErro(RuntimeError):
    """Não foi possível escolher a mensagem com segurança (ambígua, por exemplo)."""


class SemMensagemNova(RraEmailErro):
    """Ainda não há mensagem nova que combine: o chamador espera e tenta de novo."""


def horario_solicitado(corpo) -> dt.datetime | None:
    m = _SOLICITADO.search(corpo or "")
    if not m:
        return None
    dia, mes, ano, hh, mm, ss = (int(x) for x in m.groups())
    try:
        return dt.datetime(ano, mes, dia, hh, mm, ss, tzinfo=FUSO_BRASILIA)
    except ValueError:
        return None


def escolher_mensagem(candidatas, ids_antes, disparo: dt.datetime, *,
                      antes_s: int = ANTES_S, depois_s: int = DEPOIS_S) -> str:
    """Id da única mensagem que: não existia antes do disparo, vem de `notas@papaki.com.br`, tem o
    assunto esperado e cujo "solicitado em" cai entre `antes_s` antes e `depois_s` depois do disparo.
    Nenhuma -> SemMensagemNova; mais de uma -> RraEmailErro (nunca escolhe no palpite)."""
    ja_existiam = set(ids_antes)
    validas = []
    for c in candidatas:
        if c.get("id") in ja_existiam:
            continue
        if REMETENTE not in str(c.get("remetente", "")).casefold():
            continue
        if ASSUNTO.casefold() not in str(c.get("assunto", "")).casefold():
            continue
        quando = horario_solicitado(c.get("corpo"))
        if quando is None:
            continue
        if -antes_s <= (quando - disparo).total_seconds() <= depois_s:
            validas.append(c["id"])
    if not validas:
        raise SemMensagemNova("ainda não há mensagem nova do RRA que combine com o disparo")
    if len(set(validas)) > 1:
        raise RraEmailErro(f"mais de uma mensagem combina com o disparo do RRA (ids: {', '.join(sorted(set(validas)))}); "
                           "pare e peça ajuda ao usuário")
    return validas[0]
```

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_rra_email.py -q`
Expected: todos passam.

- [ ] **Step 5: Commit**

```bash
git add src/papaki_dre/rra_email.py tests/test_rra_email.py
git commit -m "feat: escolha da mensagem do RRA (remetente, assunto, id novo, horário do pedido)" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Registro da extração e travas (`extracao`)

**Files:**
- Create: `src/papaki_dre/extracao.py`, `tests/extracao_factory.py`
- Test: `tests/test_extracao.py`

- [ ] **Step 1: Fábrica de extração sintética**

Crie `tests/extracao_factory.py`:

```python
"""Extrações sintéticas para os testes: arquivos inventados, sem dados reais."""
import datetime as dt
from pathlib import Path

from papaki_dre.extracao import RELATORIOS, iniciar

OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64
ZIP = b"PK\x03\x04" + b"\x00" * 64
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


def extracao_fechada(tmp_path, *, produtora="untitled", competencia="2026-09", periodo_fim="2026-09-30"):
    """Faz o caminho completo (tela -> disparo -> arquivo) dos 4 relatórios e fecha. Devolve a Extracao."""
    base = Path(tmp_path) / "dados" / "extracoes" / competencia / produtora
    ex = iniciar(base, produtora, competencia, periodo_fim)
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
```

- [ ] **Step 2: Escrever os testes que falham**

Crie `tests/test_extracao.py`:

```python
import json
import os
import time
from pathlib import Path

import pytest

from extracao_factory import CONTEUDO, OLE, extracao_fechada, mensagem_rra, png, tela
from papaki_dre.extracao import (ARQ_MANIFESTO, RELATORIOS, Extracao, ExtracaoErro, carregar_manifesto, iniciar)


@pytest.fixture
def ex(tmp_path):
    return iniciar(tmp_path / "dados" / "extracoes" / "2026-09" / "untitled", "untitled", "2026-09", "2026-09-30")


@pytest.fixture
def print_(tmp_path):
    return png(tmp_path)


def _arquivo_novo(tmp_path, rel):
    nome, dados = CONTEUDO[rel]
    p = tmp_path / nome
    p.write_bytes(dados)
    return p


# --- iniciar / retomar ---------------------------------------------------------------------------

def test_iniciar_cria_a_pasta_com_estado(ex):
    assert (ex.pasta / "estado.json").is_file()
    assert ex.estado["periodo_inicio"] == "2026-01-01" and ex.estado["periodo_fim"] == "2026-09-30"


def test_produtora_desconhecida(tmp_path):
    with pytest.raises(ExtracaoErro, match="produtora"):
        iniciar(tmp_path / "x", "outra", "2026-09", "2026-09-30")


def test_extracao_aberta_exige_retomar(ex):
    with pytest.raises(ExtracaoErro, match="--retomar"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30")
    de_novo = iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)
    assert de_novo.pasta == ex.pasta


def test_retomar_sem_nada_aberto(tmp_path):
    with pytest.raises(ExtracaoErro, match="nada para retomar|não há extração aberta"):
        iniciar(tmp_path / "x", "untitled", "2026-09", "2026-09-30", retomar=True)


def test_nova_nao_e_permitida_com_disparo_sem_arquivo(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    with pytest.raises(ExtracaoErro, match="disparo sem arquivo"):
        iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", nova=True)


def test_retomar_continua_de_onde_parou(ex, print_, tmp_path):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    ex.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    de_novo = iniciar(ex.pasta.parent, "untitled", "2026-09", "2026-09-30", retomar=True)
    assert de_novo.estado["relatorios"]["cp"]["arquivo"]["nome"] == "Contas a pagar.csv"
    with pytest.raises(ExtracaoErro, match="já disparado"):
        de_novo.disparar("cp")


def test_extracao_fechada_nao_conta_como_aberta(tmp_path):
    import datetime as dt
    from papaki_dre.config import FUSO_BRASILIA
    fechada = extracao_fechada(tmp_path)
    depois = lambda: dt.datetime.now(FUSO_BRASILIA) + dt.timedelta(minutes=5)   # outro nome de pasta
    nova = iniciar(fechada.pasta.parent, "untitled", "2026-09", "2026-09-30", agora=depois)
    assert (fechada.pasta / ARQ_MANIFESTO).is_file()
    assert nova.pasta != fechada.pasta and (nova.pasta / "estado.json").is_file()


# --- travas da tela ------------------------------------------------------------------------------

@pytest.mark.parametrize("rel,mudanca,trecho", [
    ("cr", dict(localidade="UTLD PRODUCAO"), "Localidade"),                 # sem cedilha/til: proibida
    ("cr", dict(localidade="Todas"), "Localidade"),
    ("comissoes", dict(localidade="Todas as localidades"), "Localidade"),
    ("cp", dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA"), "Localidade"),
    ("rra", dict(localidade="UTLD PRODUÇÃO DE FILMES LTDA"), "Localidade"),
    ("cp", dict(criterio="Venc. alterado"), "Critério"),
    ("rra", dict(criterio="Data competência"), "Critério"),
    ("cp", dict(data_inicio="2026-02-01"), "Data inicial"),
    ("cp", dict(data_fim="2026-09-15"), "Data final"),
    ("cp", dict(data_fim="amanhã"), "Data final"),
    ("cr", dict(status="Pendentes"), "Status"),
    ("cp", dict(pesquisar_em="Individuais"), "Pesquisar em"),
    ("cp", dict(total=-1), "Total"),
    ("cp", dict(total="muitos"), "Total"),
])
def test_tela_invalida_e_recusada(ex, print_, rel, mudanca, trecho):
    with pytest.raises(ExtracaoErro, match=trecho):
        ex.registrar_tela(rel, tela(rel, **mudanca), print_)
    assert "tela" not in ex.estado["relatorios"].get(rel, {})


def test_tela_aceita_variantes_legitimas(ex, print_):
    ex.registrar_tela("cp", tela("cp", localidade="Todas as localidades", data_inicio="01/01/2026",
                                data_fim="30/09/2026", status=" todos "), print_)
    t = ex.estado["relatorios"]["cp"]["tela"]
    assert t["data_inicio"] == "2026-01-01" and t["data_fim"] == "2026-09-30" and t["total"] == 37030


def test_tela_sem_print_e_recusada(ex, tmp_path):
    with pytest.raises(ExtracaoErro, match="print"):
        ex.registrar_tela("cp", tela("cp"), tmp_path / "nao-existe.png")


def test_print_e_copiado_para_a_pasta(ex, print_):
    ex.registrar_tela("cr", tela("cr"), print_)
    assert (ex.pasta / "prints" / "cr.png").is_file()
    assert ex.estado["relatorios"]["cr"]["prova"] == "prints/cr.png"


def test_relatorio_desconhecido(ex, print_):
    with pytest.raises(ExtracaoErro, match="relatório"):
        ex.registrar_tela("xx", tela("cp"), print_)


# --- disparo único -------------------------------------------------------------------------------

def test_nao_dispara_sem_tela_validada(ex):
    with pytest.raises(ExtracaoErro, match="registre a tela"):
        ex.disparar("cp")


def test_um_disparo_por_relatorio_e_o_registro_vai_para_o_disco(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    gravado = json.loads((ex.pasta / "estado.json").read_text(encoding="utf-8"))
    assert gravado["relatorios"]["cp"]["disparo"]
    with pytest.raises(ExtracaoErro, match="NÃO dispare de novo"):
        ex.disparar("cp")


def test_tela_nao_muda_depois_do_disparo(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    with pytest.raises(ExtracaoErro, match="já foi disparado"):
        ex.registrar_tela("cp", tela("cp", total=1), print_)


def test_rra_exige_ids_antes_registrados(ex, print_):
    ex.registrar_tela("rra", tela("rra"), print_)
    with pytest.raises(ExtracaoErro, match="rra-ids-antes"):
        ex.disparar("rra")
    ex.rra_ids_antes([])
    ex.disparar("rra")
    with pytest.raises(ExtracaoErro, match="depois do disparo"):
        ex.rra_ids_antes(["x"])


# --- arquivos ------------------------------------------------------------------------------------

def test_arquivo_exige_disparo(ex, tmp_path):
    with pytest.raises(ExtracaoErro, match="disparo"):
        ex.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))


@pytest.fixture
def disparado(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    return ex


def test_arquivo_registrado_com_hash(disparado, tmp_path):
    disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))
    a = disparado.estado["relatorios"]["cp"]["arquivo"]
    assert (disparado.pasta / a["nome"]).is_file() and len(a["sha256"]) == 64 and a["tamanho"] > 0
    with pytest.raises(ExtracaoErro, match="já tem arquivo"):
        disparado.registrar_arquivo("cp", _arquivo_novo(tmp_path, "cp"))


def test_arquivo_vazio(disparado, tmp_path):
    p = tmp_path / "vazio.csv"
    p.write_bytes(b"")
    with pytest.raises(ExtracaoErro, match="vazio"):
        disparado.registrar_arquivo("cp", p)


def test_extensao_errada(disparado, tmp_path):
    p = tmp_path / "x.xlsx"
    p.write_bytes(b"a;b\n")
    with pytest.raises(ExtracaoErro, match="extensão"):
        disparado.registrar_arquivo("cp", p)


def test_assinatura_errada(disparado, tmp_path):
    p = tmp_path / "x.csv"
    p.write_bytes(b"\x00\x01\x02 sem ponto e virgula")
    with pytest.raises(ExtracaoErro, match="não parece"):
        disparado.registrar_arquivo("cp", p)


def test_arquivo_anterior_ao_disparo_e_recusado(disparado, tmp_path):
    p = _arquivo_novo(tmp_path, "cp")
    velho = time.time() - 3600
    os.utime(p, (velho, velho))
    with pytest.raises(ExtracaoErro, match="anterior ao disparo"):
        disparado.registrar_arquivo("cp", p)


def test_rra_exige_mensagem_escolhida_e_assinatura_ole2(ex, print_, tmp_path):
    ex.registrar_tela("rra", tela("rra"), print_)
    ex.rra_ids_antes([])
    ex.disparar("rra")
    p = tmp_path / "Relatorio.xls"
    p.write_bytes(OLE)
    with pytest.raises(ExtracaoErro, match="mensagem"):
        ex.registrar_arquivo("rra", p)
    ex.escolher_rra([mensagem_rra(ex)])
    ex.registrar_arquivo("rra", p)
    ruim = tmp_path / "Outro.xls"
    ruim.write_bytes(b"isto nao e um xls de verdade")
    ex2 = iniciar(tmp_path / "d2", "untitled", "2026-09", "2026-09-30")
    ex2.registrar_tela("rra", tela("rra"), print_)
    ex2.rra_ids_antes([])
    ex2.disparar("rra")
    ex2.escolher_rra([mensagem_rra(ex2)])
    with pytest.raises(ExtracaoErro, match="não parece"):
        ex2.registrar_arquivo("rra", ruim)


def test_escolher_rra_e_idempotente(ex, print_):
    ex.registrar_tela("rra", tela("rra"), print_)
    ex.rra_ids_antes(["velha"])
    ex.disparar("rra")
    assert ex.escolher_rra([mensagem_rra(ex, "m1")]) == "m1"
    assert ex.escolher_rra([mensagem_rra(ex, "m1"), mensagem_rra(ex, "m2")]) == "m1"


# --- fechar e manifesto --------------------------------------------------------------------------

def test_fechar_incompleto_lista_o_que_falta(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    with pytest.raises(ExtracaoErro) as e:
        ex.fechar()
    for rel in RELATORIOS:
        assert rel in str(e.value)


def test_fechar_gera_manifesto_valido(tmp_path):
    ex = extracao_fechada(tmp_path)
    m = carregar_manifesto(ex.pasta)
    assert m["produtora"] == "untitled" and m["competencia"] == "2026-09" and m["periodo_fim"] == "2026-09-30"
    assert set(m["relatorios"]) == set(RELATORIOS)
    cp = m["relatorios"]["cp"]
    assert cp["total_tela"] == 37030 and cp["prova"] == "prints/cp.png"
    assert cp["filtros_tela"]["localidade"] == "Todas" and "total" not in cp["filtros_tela"]
    assert m["data_extracao"] == min(r["disparo"] for r in m["relatorios"].values())
    assert m["origem"] == "Esferas (Chrome)"


def test_manifesto_com_arquivo_adulterado_e_recusado(tmp_path):
    ex = extracao_fechada(tmp_path)
    (ex.pasta / "Contas a pagar.csv").write_bytes(b"\xef\xbb\xbfoutro;conteudo\n")
    with pytest.raises(ExtracaoErro, match="não confere"):
        carregar_manifesto(ex.pasta)


def test_manifesto_com_arquivo_sumido(tmp_path):
    ex = extracao_fechada(tmp_path)
    (ex.pasta / "Comissoes.xlsx").unlink()
    with pytest.raises(ExtracaoErro, match="não encontrado"):
        carregar_manifesto(ex.pasta)


def test_pasta_sem_manifesto_ou_ilegivel(tmp_path):
    with pytest.raises(ExtracaoErro, match="não está fechada"):
        carregar_manifesto(tmp_path)
    (tmp_path / ARQ_MANIFESTO).write_text("{ruim", encoding="utf-8")
    with pytest.raises(ExtracaoErro, match="ilegível"):
        carregar_manifesto(tmp_path)


def test_abrir_pasta_invalida(tmp_path):
    with pytest.raises(ExtracaoErro, match="inválida"):
        Extracao.abrir(tmp_path)


def test_status_resume_o_que_ja_foi_feito(ex, print_):
    ex.registrar_tela("cp", tela("cp"), print_)
    ex.disparar("cp")
    s = ex.status()
    assert s["relatorios"]["cp"] == {"tela": True, "disparo": True, "arquivo": False}
    assert s["relatorios"]["cr"] == {"tela": False, "disparo": False, "arquivo": False}
    assert s["rra_ids_antes"] is False and s["rra_mensagem"] is False
```

- [ ] **Step 3: Rodar e ver falhar**

Run: `py -m pytest tests/test_extracao.py -q`
Expected: FAIL (`ModuleNotFoundError: papaki_dre.extracao`).

- [ ] **Step 4: Implementar `src/papaki_dre/extracao.py`**

```python
"""Registro de cada extração do Esferas: o que a tela mostrou, o disparo único de cada relatório e os
arquivos baixados. O Claude chama estes passos pelo CLI (`extracao ...`); as travas valem AQUI, não só
no roteiro. Nada neste módulo lê o conteúdo dos relatórios (só assinatura de formato, tamanho e hash)."""
from __future__ import annotations

import datetime as dt
import json
import os
import shutil
from pathlib import Path

from . import rra_email
from .config import FUSO_BRASILIA, PRODUTORAS
from .db import sha256_file

RELATORIOS = ("cp", "cr", "rra", "comissoes")
ARQ_ESTADO = "estado.json"
ARQ_MANIFESTO = "extracao.json"
VERSAO = 1
ORIGEM = "Esferas (Chrome)"
CRITERIO = {"cp": "Data competência", "cr": "Data competência", "rra": "Aprovação", "comissoes": "Aprovação"}
EXTENSAO = {"cp": ".csv", "cr": ".csv", "rra": ".xls", "comissoes": ".xlsx"}
EXT_PRINT = {".png", ".jpg", ".jpeg"}
PESQUISAR_EM_CP = "Individuais e agrupadas"
OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
TOLERANCIA_MTIME_S = 5


class ExtracaoErro(ValueError):
    """Trava violada ou uso fora de ordem. A mensagem não leva dados dos relatórios."""


def _agora() -> dt.datetime:
    return dt.datetime.now(FUSO_BRASILIA)


def _norm(x) -> str:
    return " ".join(str("" if x is None else x).split()).casefold()


def _data(texto) -> dt.date | None:
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(str(texto or "").strip(), fmt).date()
        except ValueError:
            pass
    return None


def localidade_esperada(rel: str, produtora: str) -> str | None:
    """None = 'Todas'. CR e Comissões são exatamente a Localidade da produtora (nunca 'UTLD PRODUCAO')."""
    return None if rel in ("cp", "rra") else PRODUTORAS[produtora]["localidade"]


def validar_tela(rel: str, estado: dict, tela: dict) -> list[str]:
    erros = []
    esperado = localidade_esperada(rel, estado["produtora"])
    visto = tela.get("localidade")
    if esperado is None:
        if not _norm(visto).startswith("todas"):
            erros.append(f"Localidade deve ser 'Todas'; a tela mostra {visto!r}")
    elif _norm(visto) != _norm(esperado):
        erros.append(f"Localidade deve ser exatamente {esperado!r}; a tela mostra {visto!r}")
    if _norm(tela.get("criterio")) != _norm(CRITERIO[rel]):
        erros.append(f"Critério de data deve ser {CRITERIO[rel]!r}; a tela mostra {tela.get('criterio')!r}")
    ini, fim = _data(tela.get("data_inicio")), _data(tela.get("data_fim"))
    if ini != dt.date.fromisoformat(estado["periodo_inicio"]):
        erros.append(f"Data inicial deve ser {estado['periodo_inicio']}; a tela mostra {tela.get('data_inicio')!r}")
    if fim != dt.date.fromisoformat(estado["periodo_fim"]):
        erros.append(f"Data final deve ser {estado['periodo_fim']}; a tela mostra {tela.get('data_fim')!r}")
    if _norm(tela.get("status")) != "todos":
        erros.append(f"Status/aba deve ser 'Todos'; a tela mostra {tela.get('status')!r}")
    if rel == "cp" and _norm(tela.get("pesquisar_em")) != _norm(PESQUISAR_EM_CP):
        erros.append(f"Pesquisar em deve ser {PESQUISAR_EM_CP!r}; a tela mostra {tela.get('pesquisar_em')!r}")
    total = tela.get("total")
    if isinstance(total, bool) or not isinstance(total, int) or total < 0:
        erros.append(f"Total de registros deve ser um inteiro >= 0; recebido {total!r}")
    return erros


def _assinatura_ok(rel: str, caminho: Path) -> bool:
    with open(caminho, "rb") as f:
        cab = f.read(4096)
    if rel == "rra":
        return cab.startswith(OLE2)
    if rel == "comissoes":
        return cab.startswith(b"PK\x03\x04")
    return b"\x00" not in cab and b";" in cab          # CSV (UTF-8 com BOM ou Windows-1252)


def _gravar_json(caminho: Path, dados: dict) -> None:
    tmp = caminho.with_name(caminho.name + ".tmp")
    tmp.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, caminho)


class Extracao:
    def __init__(self, pasta, estado: dict, agora=None):
        self.pasta, self.estado, self._agora = Path(pasta), estado, agora or _agora

    @classmethod
    def abrir(cls, pasta, agora=None) -> "Extracao":
        try:
            estado = json.loads((Path(pasta) / ARQ_ESTADO).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raise ExtracaoErro(f"pasta de extração inválida (sem {ARQ_ESTADO} legível): {Path(pasta).name}") from None
        return cls(pasta, estado, agora)

    def _salvar(self) -> None:
        _gravar_json(self.pasta / ARQ_ESTADO, self.estado)

    def _rel(self, rel: str) -> dict:
        if rel not in RELATORIOS:
            raise ExtracaoErro(f"relatório desconhecido: {rel!r} (use {', '.join(RELATORIOS)})")
        return self.estado["relatorios"].setdefault(rel, {})

    # --- tela -----------------------------------------------------------------------------------
    def registrar_tela(self, rel: str, tela: dict, print_path) -> None:
        r = self._rel(rel)
        if r.get("disparo"):
            raise ExtracaoErro(f"{rel}: já foi disparado; a tela registrada não pode mais mudar")
        erros = validar_tela(rel, self.estado, tela)
        print_path = Path(print_path)
        if not print_path.is_file():
            erros.append("print da tela não encontrado")
        elif print_path.suffix.casefold() not in EXT_PRINT:
            erros.append("o print deve ser .png ou .jpg")
        if erros:
            raise ExtracaoErro(f"{rel}: trava da tela falhou - " + "; ".join(erros) + ". NÃO exporte.")
        destino = self.pasta / "prints" / f"{rel}{print_path.suffix.casefold()}"
        destino.parent.mkdir(exist_ok=True)
        shutil.copyfile(print_path, destino)
        r["tela"] = {
            "localidade": str(tela["localidade"]).strip(), "criterio": str(tela["criterio"]).strip(),
            "data_inicio": _data(tela["data_inicio"]).isoformat(), "data_fim": _data(tela["data_fim"]).isoformat(),
            "status": str(tela["status"]).strip(), "total": tela["total"],
            **({"pesquisar_em": str(tela["pesquisar_em"]).strip()} if tela.get("pesquisar_em") else {}),
            "vista_em": self._agora().isoformat()}
        r["prova"] = f"prints/{destino.name}"
        self._salvar()

    # --- disparo --------------------------------------------------------------------------------
    def rra_ids_antes(self, ids) -> None:
        if self._rel("rra").get("disparo"):
            raise ExtracaoErro("rra: os ids anteriores só podem ser registrados antes do disparo; "
                               "já existe disparo, não dá mais para registrar depois do disparo")
        self.estado["rra_ids_antes"] = sorted({str(i) for i in ids})
        self._salvar()

    def disparar(self, rel: str) -> str:
        r = self._rel(rel)
        if r.get("disparo"):
            raise ExtracaoErro(f"{rel}: já disparado em {r['disparo']}: NÃO dispare de novo")
        if not r.get("tela"):
            raise ExtracaoErro(f"{rel}: registre a tela (extracao tela) antes de disparar")
        if rel == "rra" and self.estado.get("rra_ids_antes") is None:
            raise ExtracaoErro("rra: registre os ids dos e-mails existentes (extracao rra-ids-antes) antes de disparar")
        r["disparo"] = self._agora().isoformat()
        self._salvar()                                   # gravado ANTES do clique
        return r["disparo"]

    def escolher_rra(self, candidatas) -> str:
        r = self._rel("rra")
        if self.estado.get("rra_mensagem"):
            return self.estado["rra_mensagem"]           # idempotente: a primeira escolha vale
        if not r.get("disparo"):
            raise ExtracaoErro("rra: ainda não foi disparado")
        escolhida = rra_email.escolher_mensagem(
            candidatas, self.estado.get("rra_ids_antes") or [], dt.datetime.fromisoformat(r["disparo"]))
        self.estado["rra_mensagem"] = escolhida
        self._salvar()
        return escolhida

    # --- arquivo --------------------------------------------------------------------------------
    def registrar_arquivo(self, rel: str, caminho) -> dict:
        r = self._rel(rel)
        if not r.get("disparo"):
            raise ExtracaoErro(f"{rel}: sem disparo registrado; não há o que baixar")
        if r.get("arquivo"):
            raise ExtracaoErro(f"{rel}: já tem arquivo registrado")
        if rel == "rra" and not self.estado.get("rra_mensagem"):
            raise ExtracaoErro("rra: escolha antes a mensagem do e-mail (extracao rra-escolher)")
        src = Path(caminho)
        if not src.is_file():
            raise ExtracaoErro(f"{rel}: arquivo não encontrado: {src.name}")
        tamanho = src.stat().st_size
        if tamanho == 0:
            raise ExtracaoErro(f"{rel}: o arquivo está vazio")
        if src.suffix.casefold() != EXTENSAO[rel]:
            raise ExtracaoErro(f"{rel}: extensão {src.suffix!r} inesperada (esperada {EXTENSAO[rel]})")
        if not _assinatura_ok(rel, src):
            raise ExtracaoErro(f"{rel}: o arquivo não parece um {EXTENSAO[rel]} válido")
        disparo = dt.datetime.fromisoformat(r["disparo"])
        modificado = dt.datetime.fromtimestamp(src.stat().st_mtime, FUSO_BRASILIA)
        if modificado < disparo - dt.timedelta(seconds=TOLERANCIA_MTIME_S):
            raise ExtracaoErro(f"{rel}: o arquivo é anterior ao disparo (provavelmente um download antigo)")
        destino = self.pasta / src.name
        if destino.exists():
            raise ExtracaoErro(f"{rel}: já existe {src.name} nesta extração")
        shutil.copyfile(src, destino)
        r["arquivo"] = {"nome": destino.name, "sha256": sha256_file(destino), "tamanho": tamanho,
                        "registrado_em": self._agora().isoformat()}
        self._salvar()
        return r["arquivo"]

    # --- fechamento -----------------------------------------------------------------------------
    def status(self) -> dict:
        rels = {rel: {"tela": bool(self.estado["relatorios"].get(rel, {}).get("tela")),
                      "disparo": bool(self.estado["relatorios"].get(rel, {}).get("disparo")),
                      "arquivo": bool(self.estado["relatorios"].get(rel, {}).get("arquivo"))}
                for rel in RELATORIOS}
        return {"pasta": self.pasta.name, "competencia": self.estado["competencia"], "relatorios": rels,
                "rra_ids_antes": self.estado.get("rra_ids_antes") is not None,
                "rra_mensagem": bool(self.estado.get("rra_mensagem")),
                "fechada": (self.pasta / ARQ_MANIFESTO).is_file()}

    def fechar(self) -> dict:
        faltam = []
        for rel in RELATORIOS:
            r = self.estado["relatorios"].get(rel, {})
            for chave, texto in (("tela", "tela validada"), ("disparo", "disparo"), ("arquivo", "arquivo")):
                if not r.get(chave):
                    faltam.append(f"{rel}: falta {texto}")
        if not self.estado.get("rra_mensagem"):
            faltam.append("rra: falta a mensagem do e-mail escolhida")
        if faltam:
            raise ExtracaoErro("extração incompleta - " + "; ".join(faltam))
        rels = {}
        for rel in RELATORIOS:
            r = self.estado["relatorios"][rel]
            rels[rel] = {"arquivo": r["arquivo"]["nome"], "sha256": r["arquivo"]["sha256"],
                         "tamanho": r["arquivo"]["tamanho"], "total_tela": r["tela"]["total"],
                         "filtros_tela": {k: v for k, v in r["tela"].items() if k != "total"},
                         "prova": r["prova"], "disparo": r["disparo"]}
        manifesto = {"versao": VERSAO, "produtora": self.estado["produtora"],
                     "competencia": self.estado["competencia"], "periodo_inicio": self.estado["periodo_inicio"],
                     "periodo_fim": self.estado["periodo_fim"],
                     "data_extracao": min(r["disparo"] for r in rels.values()), "origem": ORIGEM,
                     "rra_mensagem": self.estado["rra_mensagem"], "relatorios": rels}
        _gravar_json(self.pasta / ARQ_MANIFESTO, manifesto)
        return manifesto


def _abertas(base: Path) -> list[Path]:
    if not base.is_dir():
        return []
    return sorted(p for p in base.iterdir() if (p / ARQ_ESTADO).is_file() and not (p / ARQ_MANIFESTO).is_file())


def _tem_disparo_sem_arquivo(pasta: Path) -> bool:
    try:
        rels = Extracao.abrir(pasta).estado["relatorios"]
    except ExtracaoErro:
        return False
    return any(r.get("disparo") and not r.get("arquivo") for r in rels.values())


def iniciar(base, produtora: str, competencia: str, periodo_fim: str, *, retomar: bool = False,
            nova: bool = False, agora=None) -> Extracao:
    """Cria a pasta de uma extração em `base` (= .../extracoes/<competencia>/<produtora>) ou, com
    `retomar`, reabre a aberta. Uma extração aberta com disparo sem arquivo nunca é abandonada."""
    if produtora not in PRODUTORAS:
        raise ExtracaoErro(f"produtora desconhecida: {produtora!r}")
    if retomar and nova:
        raise ExtracaoErro("use --retomar ou --nova, não os dois")
    base = Path(base)
    abertas = _abertas(base)
    if retomar:
        if not abertas:
            raise ExtracaoErro("não há extração aberta para retomar")
        ex = Extracao.abrir(abertas[-1], agora)
        if ex.estado["periodo_fim"] != periodo_fim or ex.estado["competencia"] != competencia:
            raise ExtracaoErro("a extração aberta é de outra competência/período")
        return ex
    if abertas:
        pendente = [p.name for p in abertas if _tem_disparo_sem_arquivo(p)]
        if pendente:
            raise ExtracaoErro(f"há extração aberta com disparo sem arquivo ({pendente[-1]}): use --retomar; "
                               "não é seguro começar outra (dispararia o relatório de novo)")
        if not nova:
            raise ExtracaoErro(f"há extração aberta ({abertas[-1].name}): use --retomar (ou --nova para ignorá-la)")
    momento = (agora or _agora)()
    pasta = base / momento.strftime("%Y%m%d-%H%M")
    if pasta.exists():
        raise ExtracaoErro(f"a pasta {pasta.name} já existe")
    pasta.mkdir(parents=True)
    estado = {"versao": VERSAO, "produtora": produtora, "competencia": competencia,
              "periodo_inicio": f"{competencia[:4]}-01-01", "periodo_fim": periodo_fim,
              "criada_em": momento.isoformat(), "rra_ids_antes": None, "rra_mensagem": None, "relatorios": {}}
    ex = Extracao(pasta, estado, agora)
    ex._salvar()
    return ex


def carregar_manifesto(pasta) -> dict:
    """Lê e confere `extracao.json`: versão, os 4 relatórios e cada arquivo (existência, tamanho, hash)."""
    pasta = Path(pasta)
    caminho = pasta / ARQ_MANIFESTO
    if not caminho.is_file():
        raise ExtracaoErro(f"a extração {pasta.name} não está fechada (sem {ARQ_MANIFESTO})")
    try:
        m = json.loads(caminho.read_text(encoding="utf-8"))
    except ValueError:
        raise ExtracaoErro(f"{ARQ_MANIFESTO} ilegível em {pasta.name}") from None
    if m.get("versao") != VERSAO or set(m.get("relatorios", {})) != set(RELATORIOS):
        raise ExtracaoErro(f"{ARQ_MANIFESTO} de {pasta.name} com versão ou relatórios inesperados")
    for rel in RELATORIOS:
        r = m["relatorios"][rel]
        arq = pasta / r["arquivo"]
        if not arq.is_file():
            raise ExtracaoErro(f"{rel}: arquivo não encontrado na extração: {r['arquivo']}")
        if arq.stat().st_size != r["tamanho"] or sha256_file(arq) != r["sha256"]:
            raise ExtracaoErro(f"{rel}: o arquivo {r['arquivo']} não confere com o registrado (foi alterado)")
    return m
```

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests/test_extracao.py -q`
Expected: todos passam. Se algum `match=` não casar por diferença de acento/redação, ajuste o **teste** ao texto da mensagem (a mensagem é a fonte), não o contrário, desde que o teste continue provando a mesma trava.

- [ ] **Step 6: Commit**

```bash
git add src/papaki_dre/extracao.py tests/extracao_factory.py tests/test_extracao.py
git commit -m "feat: registro da extração (travas de tela, disparo único, arquivos e manifesto)" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Subcomandos `agenda` e `extracao ...` (CLI)

**Files:**
- Create: `src/papaki_dre/cli_extracao.py`
- Modify: `src/papaki_dre/cli.py` (parser, `main`)
- Test: `tests/test_cli_extracao.py`

- [ ] **Step 1: Escrever os testes que falham**

Crie `tests/test_cli_extracao.py`:

```python
import json

import pytest

from extracao_factory import CONTEUDO, mensagem_rra, png, tela
from papaki_dre import cli
from papaki_dre.extracao import Extracao


@pytest.fixture
def env(tmp_path):
    return {"PAPAKI_DADOS_DIR": str(tmp_path / "dados")}


def _rodar(capsys, argv, env):
    codigo = cli.main(argv, env=env)
    saida = capsys.readouterr()
    return codigo, saida.out, saida.err


def _iniciar(capsys, env, *extra):
    codigo, out, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                       "--competencia", "2026-09", *extra], env)
    assert codigo == 0, err
    return json.loads(out)["pasta"]


def _tela_args(rel, pasta, print_, **mudar):
    t = tela(rel, **mudar)
    args = ["extracao", "tela", "--pasta", pasta, "--relatorio", rel, "--localidade", t["localidade"],
            "--criterio", t["criterio"], "--data-inicio", t["data_inicio"], "--data-fim", t["data_fim"],
            "--status", t["status"], "--total", str(t["total"]), "--print", str(print_)]
    if t.get("pesquisar_em"):
        args += ["--pesquisar-em", t["pesquisar_em"]]
    return args


# --- agenda --------------------------------------------------------------------------------------

def test_agenda_no_segundo_dia_util(capsys, env):
    codigo, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], env)
    assert codigo == 0
    assert json.loads(out) == {"rodar": True, "competencia": "2026-10", "motivo": ""}


def test_agenda_nao_roda_no_feriado(capsys, env):
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-02"], env)
    assert json.loads(out)["rodar"] is False


def test_agenda_nao_roda_se_a_competencia_ja_foi_importada(capsys, env, tmp_path):
    apuracao = tmp_path / "dados" / "apuracoes" / "untitled" / "2026-10"
    apuracao.mkdir(parents=True)
    (apuracao / "conferencia.json").write_text(json.dumps(
        {"status": "ok", "meta": {"competencia": "2026-10", "periodo": "2026-01-01 a 2026-10-31",
                                  "importacao_id": "abc"}}), encoding="utf-8")
    _, out, _ = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], env)
    assert json.loads(out)["rodar"] is False and "já importada" in json.loads(out)["motivo"]


def test_agenda_exige_a_pasta_de_dados(capsys):
    codigo, _, err = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "2026-11-04"], {})
    assert codigo == 3 and "PAPAKI_DADOS_DIR" in err


def test_agenda_data_invalida(capsys, env):
    codigo, _, err = _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", "ontem"], env)
    assert codigo == 3 and "ERRO" in err


# --- extracao: fluxo completo --------------------------------------------------------------------

def test_fluxo_completo_pelo_cli(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    print_ = png(tmp_path)
    downloads = tmp_path / "Downloads"
    downloads.mkdir()
    assert _rodar(capsys, ["extracao", "rra-ids-antes", "--pasta", pasta, "--ids", "velha1,velha2"], env)[0] == 0
    for rel in ("cp", "cr", "rra", "comissoes"):
        assert _rodar(capsys, _tela_args(rel, pasta, print_), env)[0] == 0
        codigo, out, err = _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", rel], env)
        assert codigo == 0, err
        if rel == "rra":
            ex = Extracao.abrir(pasta)
            cand = tmp_path / "cand.json"
            cand.write_text(json.dumps([mensagem_rra(ex, "novo1")]), encoding="utf-8")
            codigo, out, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta,
                                                "--candidatas", str(cand)], env)
            assert codigo == 0 and json.loads(out)["mensagem"] == "novo1"
        nome, dados = CONTEUDO[rel]
        (downloads / nome).write_bytes(dados)
        codigo, _, err = _rodar(capsys, ["extracao", "arquivo", "--pasta", pasta, "--relatorio", rel,
                                         "--caminho", str(downloads / nome)], env)
        assert codigo == 0, err
    codigo, out, err = _rodar(capsys, ["extracao", "fechar", "--pasta", pasta], env)
    assert codigo == 0, err
    assert json.loads(out)["relatorios"] == ["cp", "cr", "rra", "comissoes"]


def test_tela_invalida_devolve_codigo_3_e_a_mensagem(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    codigo, _, err = _rodar(capsys, _tela_args("cr", pasta, png(tmp_path), localidade="UTLD PRODUCAO"), env)
    assert codigo == 3 and "Localidade" in err and "NÃO exporte" in err


def test_segundo_disparo_e_recusado(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    assert _rodar(capsys, _tela_args("cp", pasta, png(tmp_path)), env)[0] == 0
    assert _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "cp"], env)[0] == 0
    codigo, _, err = _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "cp"], env)
    assert codigo == 3 and "NÃO dispare de novo" in err


def test_rra_escolher_sem_mensagem_devolve_4(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    assert _rodar(capsys, _tela_args("rra", pasta, png(tmp_path)), env)[0] == 0
    _rodar(capsys, ["extracao", "rra-ids-antes", "--pasta", pasta, "--ids", ""], env)
    assert _rodar(capsys, ["extracao", "disparar", "--pasta", pasta, "--relatorio", "rra"], env)[0] == 0
    cand = tmp_path / "vazio.json"
    cand.write_text("[]", encoding="utf-8")
    codigo, _, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta, "--candidatas", str(cand)], env)
    assert codigo == 4 and "ainda" in err


def test_rra_escolher_candidatas_ilegiveis(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    ruim = tmp_path / "ruim.json"
    ruim.write_text("{não", encoding="utf-8")
    codigo, _, err = _rodar(capsys, ["extracao", "rra-escolher", "--pasta", pasta, "--candidatas", str(ruim)], env)
    assert codigo == 3 and "ERRO" in err


def test_iniciar_de_novo_pede_retomar_e_estado_mostra_o_que_falta(capsys, env, tmp_path):
    pasta = _iniciar(capsys, env)
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09"], env)
    assert codigo == 3 and "--retomar" in err
    assert _iniciar(capsys, env, "--retomar") == pasta
    codigo, out, _ = _rodar(capsys, ["extracao", "estado", "--pasta", pasta], env)
    assert codigo == 0 and json.loads(out)["relatorios"]["cp"] == {"tela": False, "disparo": False, "arquivo": False}


def test_iniciar_sem_pasta_de_dados(capsys):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-09"], {})
    assert codigo == 3 and "PAPAKI_DADOS_DIR" in err


def test_iniciar_competencia_invalida(capsys, env):
    codigo, _, err = _rodar(capsys, ["extracao", "iniciar", "--produtora", "untitled",
                                     "--competencia", "2026-13"], env)
    assert codigo == 3 and "AAAA-MM" in err
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_cli_extracao.py -q`
Expected: FAIL (o `main` ainda não aceita `env` nem conhece os subcomandos).

- [ ] **Step 3: Implementar `src/papaki_dre/cli_extracao.py`**

```python
"""Subcomandos de apoio à rotina mensal: `agenda` e `extracao ...`. A saída que o Claude lê é sempre
JSON em ASCII (caminhos com acento chegam escapados e voltam intactos ao decodificar)."""
from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

from . import ambiente, calendario, caminhos
from .config import FERIADOS_EXTRAS, PRODUTORAS
from .extracao import RELATORIOS, Extracao, ExtracaoErro, iniciar
from .rra_email import RraEmailErro, SemMensagemNova

EXIT_OK, EXIT_FALHA, EXIT_AGUARDAR = 0, 3, 4


def registrar(sub) -> None:
    ag = sub.add_parser("agenda", help="diz se hoje é dia de rodar a rotina mensal")
    ag.add_argument("--produtora", required=True)
    ag.add_argument("--hoje", help="AAAA-MM-DD; padrão = hoje")

    ex = sub.add_parser("extracao", help="registro da extração do Esferas (usado pelo Claude)")
    acoes = ex.add_subparsers(dest="acao", required=True)

    p = acoes.add_parser("iniciar", help="cria (ou, com --retomar, reabre) a pasta de uma extração")
    p.add_argument("--produtora", required=True)
    p.add_argument("--competencia", required=True, help="AAAA-MM")
    p.add_argument("--periodo-fim", help="AAAA-MM-DD; padrão = último dia da competência")
    p.add_argument("--retomar", action="store_true")
    p.add_argument("--nova", action="store_true")

    def com_pasta(nome, ajuda, relatorio=False):
        q = acoes.add_parser(nome, help=ajuda)
        q.add_argument("--pasta", required=True, type=Path)
        if relatorio:
            q.add_argument("--relatorio", required=True, choices=RELATORIOS)
        return q

    t = com_pasta("tela", "registra e valida o que a tela mostra ANTES de exportar", relatorio=True)
    t.add_argument("--localidade", required=True)
    t.add_argument("--criterio", required=True, help="Data competência (CP, CR) ou Aprovação (RRA, Comissões)")
    t.add_argument("--data-inicio", required=True)
    t.add_argument("--data-fim", required=True)
    t.add_argument("--status", required=True, help="Todos")
    t.add_argument("--total", required=True, type=int, help="registros que a tela mostra")
    t.add_argument("--pesquisar-em", help="só no CP: Individuais e agrupadas")
    t.add_argument("--print", dest="print_", required=True, type=Path, help="imagem .png/.jpg da tela")
    com_pasta("disparar", "registra o disparo (antes do clique em Exportar/Relatórios)", relatorio=True)
    a = com_pasta("arquivo", "registra o arquivo baixado", relatorio=True)
    a.add_argument("--caminho", required=True, type=Path)
    r = com_pasta("rra-ids-antes", "registra os ids de e-mails já existentes (antes de disparar o RRA)")
    r.add_argument("--ids", required=True, help="ids separados por vírgula (vazio se não há nenhum)")
    c = com_pasta("rra-escolher", "escolhe a mensagem do RRA entre as candidatas (JSON)")
    c.add_argument("--candidatas", required=True, type=Path)
    com_pasta("fechar", "fecha a extração e gera o registro para o importar")
    com_pasta("estado", "mostra o que já foi feito (para retomar)")


def _json(dados) -> None:
    print(json.dumps(dados, ensure_ascii=True))


def executar(a, env=None) -> int:
    try:
        return _agenda(a, env) if a.cmd == "agenda" else _extracao(a, env)
    except SemMensagemNova as e:
        print(f"AGUARDE: {e}", file=sys.stderr)
        return EXIT_AGUARDAR
    except (ExtracaoErro, caminhos.CaminhosErro, RraEmailErro, ValueError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return EXIT_FALHA


def _agenda(a, env) -> int:
    if a.produtora not in PRODUTORAS:
        raise ExtracaoErro(f"produtora desconhecida: {a.produtora!r}")
    hoje = dt.date.fromisoformat(a.hoje) if a.hoje else dt.date.today()
    competencia = calendario.competencia_alvo(hoje)
    raiz = caminhos.raiz_dados(ambiente.resolver(env))
    conferencia = caminhos.pasta_apuracao(raiz, a.produtora, competencia) / "conferencia.json"
    ja = calendario.competencia_concluida(conferencia, competencia)
    rodar, motivo = calendario.deve_rodar(hoje, ja, extras=FERIADOS_EXTRAS)
    _json({"rodar": rodar, "competencia": competencia, "motivo": motivo})
    return EXIT_OK


def _extracao(a, env) -> int:
    if a.acao == "iniciar":
        competencia, _, fim = calendario.periodo_da_competencia(a.competencia)
        if a.periodo_fim:
            fim = dt.date.fromisoformat(a.periodo_fim)
        raiz = caminhos.raiz_dados(ambiente.resolver(env))
        ex = iniciar(caminhos.pasta_extracoes(raiz, competencia, a.produtora), a.produtora, competencia,
                     fim.isoformat(), retomar=a.retomar, nova=a.nova)
        _json({"pasta": str(ex.pasta)})
        return EXIT_OK
    ex = Extracao.abrir(a.pasta)
    if a.acao == "tela":
        ex.registrar_tela(a.relatorio, {
            "localidade": a.localidade, "criterio": a.criterio, "data_inicio": a.data_inicio,
            "data_fim": a.data_fim, "status": a.status, "total": a.total, "pesquisar_em": a.pesquisar_em}, a.print_)
        _json({"relatorio": a.relatorio, "tela": "validada"})
    elif a.acao == "disparar":
        _json({"relatorio": a.relatorio, "disparo": ex.disparar(a.relatorio)})
    elif a.acao == "arquivo":
        arq = ex.registrar_arquivo(a.relatorio, a.caminho)
        _json({"relatorio": a.relatorio, "arquivo": arq["nome"], "sha256": arq["sha256"][:12]})
    elif a.acao == "rra-ids-antes":
        ex.rra_ids_antes([i.strip() for i in a.ids.split(",") if i.strip()])
        _json({"rra_ids_antes": "registrado"})
    elif a.acao == "rra-escolher":
        try:
            candidatas = json.loads(Path(a.candidatas).read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raise ExtracaoErro("arquivo de candidatas ilegível (esperado: lista JSON de mensagens)") from None
        if not isinstance(candidatas, list):
            raise ExtracaoErro("arquivo de candidatas ilegível (esperado: lista JSON de mensagens)")
        _json({"mensagem": ex.escolher_rra(candidatas)})
    elif a.acao == "fechar":
        m = ex.fechar()
        _json({"fechada": True, "pasta": str(ex.pasta), "relatorios": list(m["relatorios"])})
    elif a.acao == "estado":
        _json(ex.status())
    return EXIT_OK
```

- [ ] **Step 4: Ligar no `cli.py`**

Em `src/papaki_dre/cli.py`:
1. Acrescente `from . import cli_extracao` aos imports.
2. Em `_parser()`, antes de `return p`, acrescente `cli_extracao.registrar(sub)`.
3. Troque `main` por:

```python
def main(argv=None, client=None, admin=None, env=None) -> int:
    _tolerar_console()
    try:
        args = _parser().parse_args(argv)
        if args.cmd == "importar":
            return importar(args, client=client, admin=admin, env=env)
        return cli_extracao.executar(args, env)
    except UsageError as e:
        print(f"ERRO DE USO: {e}", file=sys.stderr)
        return EXIT_FAILED
    except Exception as e:  # noqa: BLE001 - qualquer falha vira código 3 (a importação já foi marcada 'falhou')
        if os.environ.get("PAPAKI_DEBUG") == "1":
            traceback.print_exc()
        print(f"FALHA: {mensagem_segura(e)}", file=sys.stderr)
        return EXIT_FAILED
```

4. Em `importar`, mude só a assinatura para `def importar(a, client=None, admin=None, env=None) -> int:` (os parâmetros novos são usados nas Tasks 7 e 8).

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests/test_cli_extracao.py tests/test_cli.py -q`
Expected: todos passam (os testes do `importar` não mudaram de comportamento).

- [ ] **Step 6: Commit**

```bash
git add src/papaki_dre/cli_extracao.py src/papaki_dre/cli.py tests/test_cli_extracao.py
git commit -m "feat: subcomandos agenda e extracao (registro, travas e escolha do e-mail do RRA)" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 7: `importar --extracao`, opções opcionais, fuso e RESULTADO.md

**Files:**
- Modify: `src/papaki_dre/cli.py`, `src/papaki_dre/report.py`
- Modify: `tests/test_cli.py` (um teste), `tests/test_report.py` (se o arquivo existir, acrescente o teste de `render_resultado`; senão crie `tests/test_resultado.py`)
- Test: `tests/test_cli_rotina.py`

- [ ] **Step 1: Escrever os testes que falham**

Crie `tests/test_cli_rotina.py`:

```python
"""importar com --extracao, modelo/saída opcionais, fuso e RESULTADO.md. Dublês do test_cli (sanitizadores,
build e contar_nd trocados); nenhum teste toca o banco real nem o Excel."""
import json

import pytest

from extracao_factory import extracao_fechada
from fake_supabase import FakeClient
from papaki_dre import cli
from test_cli import OUT_OK, env  # noqa: F401  (fixture `env` e nome do arquivo de saída)


def _argv_ext(env, pasta, *extra, modelo=True, saida=True):
    a = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--extracao", str(pasta)]
    if modelo:
        a += ["--modelo", str(env["model"])]
    if saida:
        a += ["--saida", str(env["out"])]
    return a + list(extra)


@pytest.fixture
def extracao(env):
    return extracao_fechada(env["out"].parent)


def test_importar_com_extracao_grava_proveniencia_e_resultado(env, extracao):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, "--gatilho", "agendado"), client=client) == 0
    brutos = {r["relatorio"]: r for r in client.tables["arquivos_brutos"]}
    assert set(brutos) == {"cp", "cr", "rra", "comissoes"}
    cp = brutos["cp"]
    assert cp["total_tela"] == 37030
    assert cp["filtros_tela"]["localidade"] == "Todas" and cp["filtros_tela"]["status"] == "Todos"
    assert cp["prova_caminho"] == f"extracoes/2026-09/untitled/{extracao.pasta.name}/prints/cp.png"
    assert brutos["cr"]["filtros_tela"]["localidade"] == "UTLD PRODUÇÃO DE FILMES LTDA"
    imp = client.tables["importacoes"][0]
    assert imp["gatilho"] == "agendado" and imp["origem"] == "Esferas (Chrome)" and imp["vigente"] is True
    assert imp["data_extracao"].endswith("-03:00")
    resultado = (env["out"] / "RESULTADO.md").read_text(encoding="utf-8")
    assert "Competência: 2026-09" in resultado and "foto vigente: sim" in resultado


def test_a_tela_do_registro_alimenta_a_conferencia(env, extracao):
    client = FakeClient()
    cli.main(_argv_ext(env, extracao.pasta), client=client)
    conf = json.loads((env["out"] / "conferencia.json").read_text(encoding="utf-8"))
    assert any("tela mostrava 37030" in a for a in conf["alertas"])


@pytest.mark.parametrize("opcao", [("--cp", "x"), ("--tela", "cp=1"), ("--data-extracao", "2026-10-08T09:00:00"),
                                   ("--origem", "outra")])
def test_extracao_nao_combina_com_arquivos_soltos(env, extracao, opcao, capsys):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta, *opcao), client=client) == 3
    assert "não combina" in capsys.readouterr().err and client.calls == []


def test_extracao_de_outra_competencia(env, extracao, capsys):
    client = FakeClient()
    argv = _argv_ext(env, extracao.pasta)
    argv[argv.index("--competencia") + 1] = "2026-08"
    assert cli.main(argv, client=client) == 3
    assert "outra produtora/competência" in capsys.readouterr().err and client.calls == []


def test_extracao_adulterada_nao_grava_nada(env, extracao, capsys):
    (extracao.pasta / "Contas a pagar.csv").write_bytes(b"\xef\xbb\xbfoutro;conteudo\n")
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 3
    assert "não confere" in capsys.readouterr().err and client.calls == []


def test_extracao_nao_fechada(env, tmp_path, capsys):
    pasta = tmp_path / "aberta"
    pasta.mkdir()
    assert cli.main(_argv_ext(env, pasta), client=FakeClient()) == 3
    assert "não está fechada" in capsys.readouterr().err


def test_periodo_fim_diferente_do_registro_e_erro(env, extracao, capsys):
    assert cli.main(_argv_ext(env, extracao.pasta, "--periodo-fim", "2026-09-15"), client=FakeClient()) == 3
    assert "período" in capsys.readouterr().err


def test_sem_extracao_faltam_arquivos(env, capsys):
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--modelo", str(env["model"]),
            "--saida", str(env["out"]), "--sem-banco"]
    assert cli.main(argv) == 3
    assert "faltam: --rra, --cp, --cr, --comissoes" in capsys.readouterr().err


def test_modelo_e_saida_vem_do_ambiente(env, extracao, tmp_path):
    dados = tmp_path / "dados"
    envs = {"PAPAKI_DADOS_DIR": str(dados), "PAPAKI_MODELO_UNTITLED": str(env["model"])}
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", modelo=False, saida=False)
    assert cli.main(argv, env=envs) == 0
    assert (dados / "apuracoes" / "untitled" / "2026-09" / OUT_OK).is_file()


def test_sem_modelo_nem_ambiente_e_erro_de_uso(env, extracao, capsys):
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", modelo=False)
    assert cli.main(argv, env={}) == 3
    assert "PAPAKI_MODELO_UNTITLED" in capsys.readouterr().err


def test_raiz_de_dados_dentro_do_projeto_e_recusada(env, extracao, capsys):
    from papaki_dre.caminhos import RAIZ_PROJETO
    argv = _argv_ext(env, extracao.pasta, "--sem-banco", saida=False)
    assert cli.main(argv, env={"PAPAKI_DADOS_DIR": str(RAIZ_PROJETO / "dados"),
                               "PAPAKI_MODELO_UNTITLED": str(env["model"])}) == 3
    assert "dentro da pasta do projeto" in capsys.readouterr().err


def test_data_extracao_sem_fuso_vale_brasilia(env):
    f = env["files"]
    argv = ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--cp", str(f["cp"]),
            "--cr", str(f["cr"]), "--rra", str(f["rra"]), "--comissoes", str(f["comissoes"]),
            "--modelo", str(env["model"]), "--saida", str(env["out"]), "--data-extracao", "2026-10-08T09:00:00"]
    client = FakeClient()
    assert cli.main(argv, client=client) == 0
    assert client.tables["importacoes"][0]["data_extracao"] == "2026-10-08T09:00:00-03:00"
```

Acrescente também o teste do resumo (em `tests/test_report.py`, ao fim; se o arquivo não existir, crie `tests/test_resultado.py` com este conteúdo):

```python
from papaki_dre.report import render_resultado


def test_render_resultado_resume_em_portugues_simples():
    conf = {"status": "com_alertas", "alertas": ["algo para olhar"], "bloqueios": [],
            "meta": {"produtora": "untitled", "competencia": "2026-09", "periodo": "2026-01-01 a 2026-09-30",
                     "importacao_id": "abc"}}
    texto = render_resultado(conf)
    assert "Status: **com_alertas**" in texto and "foto vigente: sim" in texto
    assert "## Alertas (1)" in texto and "- algo para olhar" in texto and "## Bloqueios (0)" in texto


def test_render_resultado_bloqueada_ou_sem_banco_nao_e_vigente():
    base = {"alertas": [], "bloqueios": ["x"], "meta": {"produtora": "untitled", "competencia": "2026-09"}}
    assert "foto vigente: não" in render_resultado({"status": "bloqueada", **base})
    assert "nenhuma (rodada sem banco)" in render_resultado({"status": "ok", **base})
```

No `tests/test_cli.py`, no teste `test_missing_required_argument_is_a_usage_error`, troque a chamada `assert cli.main(argv) == 3` por `assert cli.main(argv, env={}) == 3` (sem `--modelo` o programa agora procura `PAPAKI_MODELO_UNTITLED`; com `env={}` o resultado não depende do `.env` real da máquina).

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_cli_rotina.py tests/test_cli.py -q`
Expected: FAIL (`--extracao` ainda não existe).

- [ ] **Step 3: `render_resultado` em `report.py`**

Acrescente ao fim de `src/papaki_dre/report.py`:

```python
def render_resultado(conf: dict) -> str:
    """Resumo curto, em português simples, para quem não vai ler a conferência inteira."""
    meta = conf.get("meta") or {}
    status = conf.get("status", "?")
    vigente = status in ("ok", "com_alertas") and bool(meta.get("importacao_id"))
    linhas = [
        "# Resultado da apuração", "",
        f"- Produtora: {meta.get('produtora', '?')}",
        f"- Competência: {meta.get('competencia', '?')} (período {meta.get('periodo', '?')})",
        f"- Status: **{status}**",
        f"- Importação no banco: {meta.get('importacao_id') or 'nenhuma (rodada sem banco)'}",
        f"- É a foto vigente: {'sim' if vigente else 'não'}", ""]
    for titulo, chave in (("Bloqueios", "bloqueios"), ("Alertas", "alertas")):
        itens = conf.get(chave) or []
        linhas.append(f"## {titulo} ({len(itens)})")
        linhas.extend([f"- {i}" for i in itens] or ["- nenhum"])
        linhas.append("")
    linhas.append("Detalhes: conferencia.md.")
    return "\n".join(linhas) + "\n"
```

- [ ] **Step 4: `cli.py` — imports, parser e funções de apoio**

Em `src/papaki_dre/cli.py`:

1. Imports (junto dos existentes): `from . import ambiente, caminhos`, `from .config import DEST, FUSO_BRASILIA, PRODUTORAS` (troque o import de `config`), `from .extracao import ExtracaoErro, carregar_manifesto`, e em `from .report import ...` acrescente `render_resultado`.
2. Em `_parser()`, troque os argumentos do `importar`:

```python
    i.add_argument("--produtora", required=True)
    i.add_argument("--competencia", required=True, help="AAAA-MM")
    for r in REPORTS:
        i.add_argument(f"--{r}", type=Path, help="bruto do relatório (ou use --extracao)")
    i.add_argument("--extracao", type=Path,
                   help="pasta de uma extração fechada (no lugar de --cp/--cr/--rra/--comissoes, --tela, "
                        "--data-extracao e --origem)")
    i.add_argument("--modelo", type=Path, help="modelo anual vazio (.xlsx); padrão: PAPAKI_MODELO_<PRODUTORA>")
    i.add_argument("--saida", type=Path, help="pasta de saída; padrão: <PAPAKI_DADOS_DIR>/apuracoes/<produtora>/<competencia>")
```

   e `--origem` passa a `default=None` (o padrão "arquivos locais" é aplicado em `_entradas`). Mantenha os demais argumentos.
3. Acrescente, antes de `importar`, estas funções:

```python
def _iso_com_fuso(texto: str | None) -> str:
    """ISO-8601 com fuso. Sem fuso, vale o horário de Brasília."""
    if not texto:
        return dt.datetime.now(dt.timezone.utc).isoformat()
    try:
        valor = dt.datetime.fromisoformat(texto)
    except ValueError:
        raise UsageError(f"--data-extracao deve ser ISO-8601; recebido {texto!r}") from None
    if valor.tzinfo is None:
        valor = valor.replace(tzinfo=FUSO_BRASILIA)
    return valor.isoformat()


def _entradas(a, competencia: str, env) -> dict:
    """Arquivos, modelo, saída e proveniência da rodada: do registro da extração (--extracao) ou das
    opções soltas. Qualquer inconsistência é UsageError (nada gravado)."""
    ent = {"tela": {}, "filtros": {}, "provas": {}, "data_extracao": None, "periodo_fim": None,
           "origem": a.origem or "arquivos locais"}
    if a.extracao:
        soltas = [n for n in (*REPORTS, "tela", "data_extracao", "origem") if getattr(a, n)]
        if soltas:
            raise UsageError("--extracao não combina com: " + ", ".join("--" + n.replace("_", "-") for n in soltas))
        try:
            m = carregar_manifesto(a.extracao)
        except ExtracaoErro as e:
            raise UsageError(f"--extracao: {e}") from None
        if m["produtora"] != a.produtora or m["competencia"] != competencia:
            raise UsageError("--extracao é de outra produtora/competência")
        if a.periodo_fim and a.periodo_fim != m["periodo_fim"]:
            raise UsageError(f"--periodo-fim difere do período da extração ({m['periodo_fim']})")
        pasta = Path(a.extracao)
        rel_pasta = caminhos.rel_extracao(competencia, a.produtora, pasta.name)
        rels = m["relatorios"]
        ent.update(
            files={k: pasta / rels[k]["arquivo"] for k in REPORTS},
            tela={k: rels[k]["total_tela"] for k in REPORTS},
            filtros={k: rels[k]["filtros_tela"] for k in REPORTS},
            provas={k: f"{rel_pasta}/{rels[k]['prova']}" for k in REPORTS},
            data_extracao=m["data_extracao"], periodo_fim=m["periodo_fim"], origem=m["origem"])
    else:
        faltam = [f"--{k}" for k in REPORTS if not getattr(a, k)]
        if faltam:
            raise UsageError("faltam: " + ", ".join(faltam) + " (ou use --extracao)")
        ent.update(files={k: Path(getattr(a, k)) for k in REPORTS}, tela=parse_tela(a.tela),
                   data_extracao=a.data_extracao)
    try:
        ent["modelo"] = Path(a.modelo) if a.modelo else caminhos.modelo_da_produtora(
            a.produtora, ambiente.resolver(env))
        ent["saida"] = Path(a.saida) if a.saida else caminhos.pasta_apuracao(
            caminhos.raiz_dados(ambiente.resolver(env)), a.produtora, competencia)
    except caminhos.CaminhosErro as e:
        raise UsageError(str(e)) from None
    return ent


def _escrever_resultado(out_dir: Path, conf: dict) -> None:
    """RESULTADO.md só nas rodadas por --extracao. Nunca transforma uma rodada encerrada em falha."""
    try:
        (out_dir / "RESULTADO.md").write_text(render_resultado(conf), encoding="utf-8")
    except OSError as e:
        print(f"AVISO: não foi possível gravar RESULTADO.md: {mensagem_segura(e)}", file=sys.stderr)
```

- [ ] **Step 5: `cli.py` — o começo de `importar`**

Troque o trecho de `importar` desde o comentário `# 1. argumentos` até o fim do passo 2 (inclui o aviso de "Arquivos idênticos") por:

```python
    # 1. argumentos (UsageError -> código 3, nada gravado)
    if a.produtora not in PRODUTORAS:
        raise UsageError(f"produtora desconhecida: {a.produtora!r} (cadastradas: {', '.join(PRODUTORAS)})")
    competencia, inicio, fim = parse_competencia(a.competencia)
    ent = _entradas(a, competencia, env)
    periodo_fim = a.periodo_fim or ent["periodo_fim"]
    if periodo_fim:
        try:
            fim = dt.date.fromisoformat(periodo_fim)
        except ValueError:
            raise UsageError(f"--periodo-fim deve ser AAAA-MM-DD; recebido {periodo_fim!r}") from None
        if fim < inicio:
            raise UsageError(f"--periodo-fim {fim.isoformat()} é anterior ao início do período "
                             f"({inicio.isoformat()})")
    extracao = _iso_com_fuso(ent["data_extracao"])
    tela, files, modelo_path, out_dir = ent["tela"], ent["files"], ent["modelo"], ent["saida"]
    nome = PRODUTORAS[a.produtora]["nome"]
    faltando = [str(p) for p in [*files.values(), modelo_path] if not p.is_file()]
    if faltando:
        raise UsageError(f"arquivo(s) não encontrado(s): {', '.join(faltando)}")
    destinos = _destinos(out_dir / _out_name(nome, competencia))
    modelo = modelo_path.resolve()
    if any(d.resolve() == modelo for d in destinos):
        raise UsageError(f"a saída sobrescreveria o modelo {modelo_path.name}; use outra pasta em --saida")
    if a.sem_banco and _saida_de_rodada_real(out_dir):
        raise UsageError(f"{out_dir} tem a saída de uma importação registrada no banco; --sem-banco não a "
                         "sobrescreve: use outra pasta em --saida")

    # 2. banco: acesso, dedupe por hash (nada é gravado até aqui)
    hashes = {k: sha256_file(p) for k, p in files.items()}
    repo = None
    if not a.sem_banco:
        repo = Repository(client if client is not None else make_client())
        try:
            repo.verificar_acesso()            # plano gratuito: pausa após ~7 dias sem uso
        except BancoIndisponivel as e:         # a mensagem já vem sem URL/chave (mensagem_segura)
            print(f"ERRO: banco inacessível ({e}).\n{MSG_PAUSADO} "
                  "(painel do Supabase ou MCP restore_project). Nada foi gravado.", file=sys.stderr)
            return EXIT_FAILED
        same = None if a.forcar else repo.find_identical(a.produtora, competencia, hashes)
        if same:
            print(f"Arquivos idênticos já importados (importação {same}); nada a fazer e nenhum arquivo "
                  "foi gerado nesta pasta. Use --forcar para repetir.")
            return EXIT_OK
```

(O `verificar_acesso` ganha a restauração automática na Task 8.)

- [ ] **Step 6: `cli.py` — o resto de `importar`**

No restante da função, aplique estas trocas pontuais:
- `build_workbook(Path(a.modelo), preparo, results)` → `build_workbook(modelo_path, preparo, results)`.
- Em `repo.criar_importacao(...)`: `origem=a.origem` → `origem=ent["origem"]`.
- No laço `for k in REPORTS:` do passo 4, a chamada de `registrar_arquivo` passa a ser:

```python
                repo.registrar_arquivo(imp_id, k, nome=files[k].name, sha256=hashes[k],
                                       tamanho=files[k].stat().st_size, total_tela=tela.get(k),
                                       filtros_tela=ent["filtros"].get(k), prova_caminho=ent["provas"].get(k))
```
- No ramo do `ContractError` (bloqueio por contrato), depois de `publicado = True`, acrescente `if a.extracao: _escrever_resultado(out_dir, blocked_conf)`.
- No caminho de sucesso (ok, com alertas ou bloqueada pela conferência), logo depois da linha `publicado = True` que já existe e **antes** do `_status(f"Status: ...")`, acrescente `if a.extracao: _escrever_resultado(out_dir, conf)`. (Tem de vir depois de `publicado = True`: assim, mesmo que gravar o resumo falhe, a rodada já promovida não é regravada como "falhou".)
- O `sanitize_rra(..., work_dir=out_dir / "work")` e o restante continuam usando `out_dir` (agora um `Path` vindo de `ent`).

- [ ] **Step 7: Rodar e ver passar**

Run: `py -m pytest tests/test_cli_rotina.py tests/test_cli.py tests/test_report.py -q`
Expected: todos passam. Se algum teste antigo do `test_cli.py` conferia o `data_extracao` gravado sem fuso ou contava exatamente os round-trips do `client.calls`, ajuste a **expectativa** do teste (agora `-03:00` para data sem fuso) e diga ao usuário qual teste mudou e por quê.

- [ ] **Step 8: Suíte rápida inteira**

Run: `py -m pytest tests -q -m "not real and not excel"`
Expected: tudo passa.

- [ ] **Step 9: Commit**

```bash
git add src/papaki_dre/cli.py src/papaki_dre/report.py tests/test_cli_rotina.py tests/test_cli.py tests/test_report.py
git commit -m "feat: importar --extracao (proveniência no banco), modelo/saída pelo ambiente, fuso e RESULTADO.md" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

(Se criou `tests/test_resultado.py` em vez de editar `test_report.py`, adicione-o no `git add`.)

---

### Task 8: `importar --restaurar-banco`

**Files:**
- Modify: `src/papaki_dre/cli.py`
- Test: `tests/test_cli_rotina.py` (acrescentar)

- [ ] **Step 1: Escrever os testes que falham**

Acrescente a `tests/test_cli_rotina.py`:

```python
from papaki_dre.supabase_admin import AdminErro


class AdminFalso:
    """Acordar 'liga' o cliente falso (que estava pausado), como a restauração real faria."""

    def __init__(self, client, resultado="restaurado", erro=None):
        self.client, self.resultado, self.erro, self.chamadas = client, resultado, erro, 0

    def acordar(self, *a, **k):
        self.chamadas += 1
        if self.erro:
            raise self.erro
        self.client.fail_with = None
        return self.resultado


@pytest.fixture
def sem_espera(monkeypatch):
    monkeypatch.setattr(cli, "_dormir", lambda s: None)


def _com_extracao(env, extracao, *extra):
    return _argv_ext(env, extracao.pasta, *extra)


def test_banco_pausado_com_restaurar_acorda_e_importa(env, extracao, sem_espera):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin) == 0
    assert admin.chamadas == 1 and client.tables["importacoes"][0]["vigente"] is True


def test_sem_a_opcao_o_banco_pausado_continua_so_avisando(env, extracao, capsys):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client)
    assert cli.main(_com_extracao(env, extracao), client=client, admin=admin) == 3
    assert admin.chamadas == 0 and "banco pausado: restaure o projeto e rode de novo" in capsys.readouterr().err


def test_restaurar_sem_chave_e_erro_de_uso_antes_de_tudo(env, extracao, capsys):
    client = FakeClient()
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env={}) == 3
    err = capsys.readouterr().err
    assert "SUPABASE_ACCESS_TOKEN" in err and client.calls == []


def test_falha_ao_acordar_vira_codigo_3_com_aviso(env, extracao, sem_espera, capsys):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    admin = AdminFalso(client, erro=AdminErro("o projeto não ficou ativo em 10 min (último estado: RESTORING)"))
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin) == 3
    err = capsys.readouterr().err
    assert "não ficou ativo" in err and "banco pausado" in err and "importacoes" not in client.tables


def test_ativo_mas_banco_ainda_nao_responde_desiste_depois_das_tentativas(env, extracao, sem_espera, capsys):
    client = FakeClient(fail_with=ConnectionError("fora do ar"))

    class NaoLiga(AdminFalso):
        def acordar(self, *a, **k):
            self.chamadas += 1
            return "ja_ativo"                      # a API diz ativo, mas o banco segue sem responder

    admin = NaoLiga(client)
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, admin=admin) == 3
    assert admin.chamadas == 1 and "banco pausado" in capsys.readouterr().err


def test_restaurar_com_token_no_ambiente_cria_o_admin_real(env, extracao, monkeypatch):
    criados = {}

    class AdminCriado(AdminFalso):
        def __init__(self, token, ref):
            criados["token"], criados["ref"] = token, ref
            super().__init__(client)

    client = FakeClient(fail_with=ConnectionError("project paused"))
    monkeypatch.setattr(cli, "AdminSupabase", AdminCriado)
    monkeypatch.setattr(cli, "_dormir", lambda s: None)
    envs = {"SUPABASE_ACCESS_TOKEN": "sbp_falso", "SUPABASE_URL": "https://refinventado.supabase.co"}
    assert cli.main(_com_extracao(env, extracao, "--restaurar-banco"), client=client, env=envs) == 0
    assert criados == {"token": "sbp_falso", "ref": "refinventado"}
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_cli_rotina.py -q`
Expected: FAIL (`--restaurar-banco` não existe).

- [ ] **Step 3: Implementar em `cli.py`**

1. Imports: `import time` e `from .supabase_admin import AdminErro, AdminSupabase, ref_do_projeto`.
2. Constantes e função de espera (junto das constantes do topo):

```python
_dormir = time.sleep                      # trocável nos testes
TENTATIVAS_APOS_ACORDAR = 10
INTERVALO_APOS_ACORDAR = 6                # segundos entre as tentativas de falar com o banco
```

3. No parser do `importar`, acrescente:

```python
    i.add_argument("--restaurar-banco", action="store_true",
                   help="se o banco estiver pausado, acorda o projeto (exige SUPABASE_ACCESS_TOKEN no .env)")
```

4. Novas funções (antes de `importar`):

```python
def _criar_admin(env) -> AdminSupabase:
    e = ambiente.resolver(env)
    token, url = e.get("SUPABASE_ACCESS_TOKEN"), e.get("SUPABASE_URL")
    if not token or not url:
        raise UsageError("--restaurar-banco exige SUPABASE_ACCESS_TOKEN e SUPABASE_URL no .env")
    try:
        return AdminSupabase(token, ref_do_projeto(url))
    except AdminErro as exc:
        raise UsageError(str(exc)) from None


def _garantir_banco(repo, admin) -> None:
    """Testa o banco; se estiver inacessível e há `admin`, confirma na API que o projeto está pausado,
    restaura, espera e testa de novo algumas vezes. Falha final -> BancoIndisponivel."""
    try:
        repo.verificar_acesso()
        return
    except BancoIndisponivel as primeiro:
        erro = primeiro
    if admin is None:
        raise erro
    print("Banco inacessível; verificando o projeto no Supabase...", file=sys.stderr)
    try:
        admin.acordar()
    except AdminErro as e:
        raise BancoIndisponivel(f"{erro}; {e}") from None
    for tentativa in range(TENTATIVAS_APOS_ACORDAR):
        try:
            repo.verificar_acesso()
            return
        except BancoIndisponivel as e:
            erro = e
            if tentativa < TENTATIVAS_APOS_ACORDAR - 1:
                _dormir(INTERVALO_APOS_ACORDAR)
    raise erro
```

5. Em `importar`, logo depois de `nome = PRODUTORAS[...]["nome"]` (no bloco novo da Task 7), acrescente:

```python
    if not a.restaurar_banco:
        admin = None                            # só a opção explícita autoriza restaurar o projeto
    elif admin is None and not a.sem_banco:
        admin = _criar_admin(env)               # sem chave: erro de uso ANTES de qualquer gravação
```

   e troque `repo.verificar_acesso()` (no `try` do passo 2) por `_garantir_banco(repo, admin)`.

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_cli_rotina.py tests/test_cli.py -q`
Expected: todos passam (inclusive `test_paused_database_is_detected_and_nothing_is_written`).

- [ ] **Step 5: Commit**

```bash
git add src/papaki_dre/cli.py tests/test_cli_rotina.py
git commit -m "feat: importar --restaurar-banco acorda o projeto pausado antes de gravar" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 9: Rodadas órfãs (`em_andamento` esquecidas)

**Files:**
- Modify: `src/papaki_dre/db.py`, `src/papaki_dre/cli.py`
- Test: `tests/test_db.py` (acrescentar), `tests/test_cli_rotina.py` (acrescentar)

- [ ] **Step 1: Escrever os testes que falham**

Acrescente a `tests/test_db.py` (usa o `FakeClient` já importado lá; se o arquivo importar de outro jeito, siga o padrão dele):

```python
import datetime as dt

from fake_supabase import FakeClient
from papaki_dre.db import Repository


def _imp(client, status="em_andamento", criado_em=None, produtora="untitled", competencia="2026-09"):
    client.tables.setdefault("importacoes", []).append({
        "id": f"id-{len(client.tables.get('importacoes', []))}", "produtora_id": produtora,
        "competencia": competencia, "status": status, "vigente": False, "gatilho": "manual",
        "criado_em": criado_em})
    return client.tables["importacoes"][-1]["id"]


def test_encerrar_orfas_marca_so_as_antigas_em_andamento():
    agora = dt.datetime(2026, 11, 4, 12, 0, tzinfo=dt.timezone.utc)
    c = FakeClient()
    velha = _imp(c, criado_em="2026-11-04T05:00:00+00:00")           # 7 h
    recente = _imp(c, criado_em="2026-11-04T09:00:00+00:00")         # 3 h
    concluida = _imp(c, status="concluida", criado_em="2026-10-01T00:00:00+00:00")
    outra_comp = _imp(c, criado_em="2026-10-01T00:00:00+00:00", competencia="2026-08")
    sem_data = _imp(c, criado_em=None)
    assert Repository(c).encerrar_orfas("untitled", "2026-09", horas=6, agora=agora) == 1
    por_id = {r["id"]: r for r in c.tables["importacoes"]}
    assert por_id[velha]["status"] == "falhou" and "interrompida" in por_id[velha]["motivo_falha"]
    assert por_id[recente]["status"] == "em_andamento"
    assert por_id[concluida]["status"] == "concluida"
    assert por_id[outra_comp]["status"] == "em_andamento" and por_id[sem_data]["status"] == "em_andamento"


def test_encerrar_orfas_aceita_z_no_fim():
    agora = dt.datetime(2026, 11, 4, 12, 0, tzinfo=dt.timezone.utc)
    c = FakeClient()
    _imp(c, criado_em="2026-11-04T01:00:00Z")
    assert Repository(c).encerrar_orfas("untitled", "2026-09", agora=agora) == 1
```

Acrescente a `tests/test_cli_rotina.py`:

```python
def test_importar_encerra_orfas_antigas_da_mesma_competencia(env, extracao):
    client = FakeClient()
    client.tables["importacoes"] = [{"id": "orfa", "produtora_id": "untitled", "competencia": "2026-09",
                                     "status": "em_andamento", "vigente": False, "gatilho": "manual",
                                     "criado_em": "2026-01-01T00:00:00+00:00"}]
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0
    por_id = {r["id"]: r for r in client.tables["importacoes"]}
    assert por_id["orfa"]["status"] == "falhou"
    assert len([r for r in client.tables["importacoes"] if r["vigente"]]) == 1


def test_arquivos_identicos_ja_importados_nao_mexem_em_nada(env, extracao):
    client = FakeClient()
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0
    antes = [dict(r) for r in client.tables["importacoes"]]
    assert cli.main(_argv_ext(env, extracao.pasta), client=client) == 0          # duplicado
    assert client.tables["importacoes"] == antes
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_db.py tests/test_cli_rotina.py -q`
Expected: FAIL (`encerrar_orfas` não existe).

- [ ] **Step 3: `Repository.encerrar_orfas`**

Em `src/papaki_dre/db.py`, dentro da classe `Repository` (depois de `marcar`):

```python
    def encerrar_orfas(self, produtora, competencia, *, horas=6, agora=None) -> int:
        """Marca 'falhou' as importações 'em_andamento' da produtora e competência criadas há mais de
        `horas` (o processo morreu no meio). Só lê e marca; nunca toca em foto vigente. Devolve quantas."""
        agora = agora or dt.datetime.now(dt.timezone.utc)
        linhas = (self.client.table("importacoes").select("id,criado_em")
                  .eq("produtora_id", produtora).eq("competencia", competencia)
                  .eq("status", "em_andamento").execute().data)
        encerradas = 0
        for r in linhas:
            if not r.get("criado_em"):
                continue
            criado = dt.datetime.fromisoformat(r["criado_em"])
            if criado.tzinfo is None:
                criado = criado.replace(tzinfo=dt.timezone.utc)
            if agora - criado > dt.timedelta(hours=horas):
                if self.marcar(r["id"], "falhou", motivo=f"rodada interrompida (sem desfecho há mais de {horas} h)"):
                    encerradas += 1
        return encerradas
```

- [ ] **Step 4: Chamar no `importar`**

Em `src/papaki_dre/cli.py`, no passo 2, **depois** do bloco "Arquivos idênticos" (isto é, só quando a rodada vai de fato acontecer) e antes de `out_dir.mkdir(...)`, acrescente (ainda dentro do `if not a.sem_banco:`):

```python
        repo.encerrar_orfas(a.produtora, competencia)
```

Se algum teste do `test_cli.py` conferir a lista exata de chamadas (`client.calls`) de uma rodada normal, acrescente a esperada `("importacoes", "select", None)` na posição certa e avise o usuário.

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests -q -m "not real and not excel"`
Expected: tudo passa.

- [ ] **Step 6: Commit**

```bash
git add src/papaki_dre/db.py src/papaki_dre/cli.py tests/test_db.py tests/test_cli_rotina.py
git commit -m "feat: encerrar rodadas órfãs em_andamento antes de uma nova rodada" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Roteiros para o Claude e script de print

**Files:**
- Create: `runbooks/extracao-esferas.md`, `runbooks/rotina-mensal.md`, `scripts/capturar_tela.ps1`
- Test: `tests/test_runbooks.py`

- [ ] **Step 1: Teste que guarda os roteiros contra o código**

Os roteiros citam comandos do programa. Este teste impede que o texto fique desatualizado. Crie `tests/test_runbooks.py`:

```python
import re
from pathlib import Path

from papaki_dre import cli
from papaki_dre.extracao import RELATORIOS

RAIZ = Path(__file__).resolve().parents[1]
ROTEIROS = [RAIZ / "runbooks" / "extracao-esferas.md", RAIZ / "runbooks" / "rotina-mensal.md"]


def _comandos(texto):
    return re.findall(r"py -m papaki_dre ([a-z-]+)(?: ([a-z-]+))?", texto)


def test_roteiros_existem_e_so_citam_comandos_que_existem():
    acoes = {"iniciar", "tela", "disparar", "arquivo", "rra-ids-antes", "rra-escolher", "fechar", "estado"}
    for roteiro in ROTEIROS:
        texto = roteiro.read_text(encoding="utf-8")
        assert texto.strip()
        for cmd, sub in _comandos(texto):
            assert cmd in {"agenda", "extracao", "importar"}, (roteiro.name, cmd)
            if cmd == "extracao":
                assert sub in acoes, (roteiro.name, sub)


def test_opcoes_citadas_do_importar_existem():
    texto = (RAIZ / "runbooks" / "rotina-mensal.md").read_text(encoding="utf-8")
    ajuda = cli._parser().parse_args(
        ["importar", "--produtora", "untitled", "--competencia", "2026-09", "--extracao", "x",
         "--gatilho", "agendado", "--restaurar-banco"])
    assert ajuda.restaurar_banco and ajuda.extracao
    assert "--extracao" in texto and "--restaurar-banco" in texto and "--gatilho" in texto


def test_roteiro_de_extracao_cobre_os_quatro_relatorios_e_as_armadilhas():
    texto = (RAIZ / "runbooks" / "extracao-esferas.md").read_text(encoding="utf-8")
    for rel in RELATORIOS:
        assert f"--relatorio {rel}" in texto
    for armadilha in ("UTLD PRODUCAO", "seta", "Localidade", "NÃO exporte", "uma única vez"):
        assert armadilha in texto
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_runbooks.py -q`
Expected: FAIL (arquivos não existem).

- [ ] **Step 3: `scripts/capturar_tela.ps1`**

```powershell
# Salva um print da tela inteira (todos os monitores) em PNG. Usado como prova de Localidade/filtros.
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -File scripts\capturar_tela.ps1 -Destino C:\pasta\cp.png
param([Parameter(Mandatory = $true)][string]$Destino)
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$tela = [System.Windows.Forms.SystemInformation]::VirtualScreen
$bmp = New-Object System.Drawing.Bitmap $tela.Width, $tela.Height
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($tela.Left, $tela.Top, 0, 0, $bmp.Size)
$bmp.Save($Destino, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
```

- [ ] **Step 4: `runbooks/extracao-esferas.md`**

````markdown
# Roteiro: extrair os 4 relatórios do Esferas

Para o Claude executar sem depender de nenhuma conversa anterior. Fonte das regras: `instrucoes_relatorios_ base.md`.
Use o Chrome já logado no Esferas (ferramentas do Chrome do Claude). Comandos `py -m papaki_dre ...` rodam na
pasta do projeto. Toda saída do programa é uma linha JSON; erro vem como `ERRO: ...` com código 3.

## Regras que valem sempre

1. **Um disparo por relatório.** Antes do clique em Exportar/Relatórios, rode `extracao disparar`. Se ele recusar
   ("NÃO dispare de novo"), **não clique**. Nunca reexporte por demora, aviso de "mais de 200 registros" ou
   falha de download: confirme o aviso e espere.
2. **Trava antes de exportar.** Depois de **Pesquisar**, leia na tela e registre com `extracao tela`. Se o programa
   recusar (Localidade, datas, critério, Status), **NÃO exporte**: pare, escreva o motivo no `RESULTADO.md` da
   pasta de apuração e avise o usuário.
3. **Localidade abre pré-preenchida** com a última usada. Confira sempre; nunca confie no valor inicial.
4. Os formulários podem **recarregar sem as datas** depois de Pesquisar: releia as datas na tela após pesquisar.
5. Nunca copie linhas, CPF/CNPJ ou dados bancários para a conversa, para logs ou para o `RESULTADO.md`. Só
   contagens, nomes de arquivo e os motivos das travas.
6. Período: **acumulado de 01/01 até o fim do período** da extração (padrão: último dia da competência).
7. Print: após Pesquisar, com os filtros visíveis, rode
   `powershell -NoProfile -ExecutionPolicy Bypass -File scripts\capturar_tela.ps1 -Destino <pasta da extração>\<relatorio>-tela.png`
   e passe esse arquivo em `--print`. A tela do computador precisa estar desbloqueada com o Chrome à frente.

## Antes de tudo

`py -m papaki_dre extracao iniciar --produtora <p> --competencia <AAAA-MM>` devolve `{"pasta": ...}`. Guarde a pasta.
Se disser que há extração aberta, rode `extracao estado --pasta <pasta>`: o que já tem arquivo **não se refaz**; o que
tem disparo sem arquivo **não se dispara de novo** (procure o arquivo em Downloads; no RRA, no e-mail; se não achar,
pare e avise). Use `iniciar ... --retomar` para continuar.

## Ordem

RRA primeiro (o relatório chega por e-mail e demora), depois CP, CR e Comissões, e então o e-mail do RRA.

## RRA — papaki.esferas.com.br

1. **Consolidado das produtoras → Jobs.** A listagem abre em nova aba.
2. Competência = **Aprovação**; período acumulado; Localidade = **Todas**; **Pesquisar**.
3. Aba **Todos**; selecione todos; confira filtros e total de registros.
4. Registre a tela: `py -m papaki_dre extracao tela --pasta <pasta> --relatorio rra --localidade "Todas" --criterio "Aprovação" --data-inicio <AAAA-MM-DD> --data-fim <AAAA-MM-DD> --status "Todos" --total <N> --print <png>`.
5. **Antes do disparo**, abra o Gmail do mesmo Chrome e liste os ids das mensagens de `notas@papaki.com.br` que já
   existem; registre com `extracao rra-ids-antes --pasta <pasta> --ids <id1,id2,...>` (`""` se nenhuma).
6. `extracao disparar --relatorio rra`. Só então: **seta à direita de "Relatórios"** (clicar no texto pode dar 404)
   → **Rentabilidade analítica**, **uma única vez**. Confirme o aviso de envio por e-mail.
7. Siga para os outros relatórios; volte ao e-mail depois.

## CP — papaki.esferas.com.br

1. **Consolidado das produtoras → Contas a pagar** (abre em nova aba).
2. **Data competência** com o período acumulado; **Status = Todos**; **Pesquisar em = Individuais e agrupadas**;
   Localidade = **Todas**; **Pesquisar**.
3. `extracao tela --relatorio cp ... --pesquisar-em "Individuais e agrupadas" ...`.
4. Veja o que já existe em Downloads (anote). `extracao disparar --relatorio cp`.
5. **Seta à direita de Exportar → "Excel"**, uma única vez (apesar do nome, baixa um CSV). Aviso de mais de 200
   registros ou de demora: só confirme e espere.
6. Quando o novo arquivo aparecer em Downloads: `extracao arquivo --relatorio cp --caminho <arquivo>`.

## CR — papaki.esferas.com.br

1. **Consolidado das produtoras → Contas a receber.**
2. Troque **Venc. alterado** por **Data competência**; período acumulado, sem datas em outros campos de competência;
   **Status = Todos**.
3. Localidade = exatamente **`UTLD PRODUÇÃO DE FILMES LTDA`**. **Nunca** `UTLD PRODUCAO` (sem cedilha e til).
4. **Pesquisar**, conferir, `extracao tela --relatorio cr ...`, anotar Downloads, `extracao disparar --relatorio cr`.
5. **Exportar** uma única vez (use a seta se houver menu); confirme o aviso; `extracao arquivo --relatorio cr ...`.
   O CSV do CR é Windows-1252 (o programa sabe; não reabra nem converta o arquivo).

## Comissões — untitled.esferas.com.br

1. **JOBS → Projetos → Projetos.**
2. Competência = **Aprovação**; período acumulado; Localidade = exatamente **`UTLD PRODUÇÃO DE FILMES LTDA`**;
   **Pesquisar**; aba **Todos**.
3. `extracao tela --relatorio comissoes ...`, anotar Downloads, `extracao disparar --relatorio comissoes`.
4. **Relatórios → Real x Orçado Comissões**, uma única vez. Baixa um XLSX direto (sem e-mail).
5. `extracao arquivo --relatorio comissoes --caminho <arquivo>`.

## E-mail do RRA

1. Abra o Gmail e leia as mensagens de `notas@papaki.com.br` com o assunto "Relatório de rentabilidade gerado com
   sucesso" (id, remetente, assunto e o texto "solicitado em dd/mm/aaaa hh:mm:ss" do corpo). Grave a lista em um
   JSON `[{"id","remetente","assunto","corpo"}, ...]` na pasta da extração.
2. `py -m papaki_dre extracao rra-escolher --pasta <pasta> --candidatas <json>`. Código `4` ("AGUARDE") = ainda não
   chegou: espere 1 minuto e repita, **até 20 minutos no total, sem novo disparo**. Código `3` com "mais de uma
   mensagem" = pare e avise o usuário. Código `0` devolve o id escolhido.
3. Abra **essa** mensagem, baixe o anexo original (não abra nem reconverta), e registre:
   `extracao arquivo --relatorio rra --caminho <arquivo .xls>`.

## Fechar

`py -m papaki_dre extracao fechar --pasta <pasta>`. Só passa com os 4 relatórios completos. O resultado é o registro
que o `importar --extracao` lê.

## Se algo der errado

| Situação | O que fazer |
|---|---|
| Sessão do Esferas expirada | Pare; peça ao usuário para logar de novo; retome com `iniciar --retomar`. |
| Trava da tela recusou | NÃO exporte; escreva o motivo; avise. |
| Disparo feito, download não apareceu | Procure em Downloads por até 5 min; não redispare; se não achar, pare e avise. |
| Arquivo recusado (extensão, formato, anterior ao disparo) | Não renomeie nem edite; pare e avise. |
| Duas mensagens candidatas do RRA | Pare e avise. |
````

- [ ] **Step 5: `runbooks/rotina-mensal.md`**

````markdown
# Roteiro: rotina mensal da apuração (texto da tarefa agendada)

Para o Claude executar sem contexto anterior. Pasta do projeto: a pasta onde está este arquivo (`runbooks/..`).
Comandos `py -m papaki_dre ...` rodam nela. Produtora: `untitled`. Regras de dados sensíveis do roteiro
`extracao-esferas.md` valem aqui: nada de linhas dos relatórios, CPF/CNPJ ou chaves em conversa, log ou arquivo.

## Duas formas de começar

- **Agendada (todo dia útil de manhã):** rode `py -m papaki_dre agenda --produtora untitled`. Leia o JSON. Se
  `"rodar": false`, **termine em silêncio** (sem aviso). Se `true`, a competência é `"competencia"`.
- **Manual (o usuário pediu "rode o mês AAAA-MM"):** pule o `agenda` e use essa competência. Para um mês ainda em
  andamento, o usuário informa o último dia (`--periodo-fim`).

## Passos

1. **Extração.** Siga `runbooks/extracao-esferas.md` até `extracao fechar` devolver a pasta da extração.
2. **Importar.**
   - Rodada **agendada**: está autorizada pela criação da tarefa. Rode:
     `py -m papaki_dre importar --produtora untitled --competencia <AAAA-MM> --extracao <pasta> --gatilho agendado --restaurar-banco`
   - Rodada **manual**: antes de gravar no Supabase com dados reais, pergunte ao usuário "posso importar no
     Supabase?" e espere o "pode importar". Sem o sim, rode com `--sem-banco --saida <pasta de ensaio>`. Com o sim:
     `py -m papaki_dre importar --produtora untitled --competencia <AAAA-MM> --extracao <pasta> --gatilho manual --restaurar-banco`
   (`--modelo` e `--saida` saem do ambiente: `PAPAKI_MODELO_UNTITLED` e `PAPAKI_DADOS_DIR`.)
3. **Desfecho pelo código de saída:**

| Código | Significado | Aviso ao usuário (português simples, 2–3 linhas) |
|---|---|---|
| 0 | ok ou com alertas (ou arquivos idênticos já importados) | "Apuração de <competência> importada. Arquivo: <nome>. Alertas: <n>." |
| 2 | bloqueada | "A apuração de <competência> foi bloqueada: <motivo curto>. Nada virou vigente." |
| 3 | falha, banco ou uso | "A rotina de <competência> falhou: <motivo>. Pode rodar de novo." |

4. **Resumo.** Leia `RESULTADO.md` na pasta de apuração (`<PAPAKI_DADOS_DIR>/apuracoes/untitled/<competência>`) e
   use só ele para o aviso. Avise o usuário (notificação do Claude, se disponível, e a mensagem final da sessão).
   **O sistema não envia e-mail.**

## Em qualquer parada antes do `importar`

Se uma trava recusou, o login expirou ou um relatório não chegou: escreva em
`<PAPAKI_DADOS_DIR>/apuracoes/untitled/<competência>/RESULTADO.md` (crie a pasta se preciso) um texto curto com a
etapa, o motivo e o que fazer ("pasta de extração aberta: use `extracao iniciar ... --retomar`"), e avise o usuário.
Não importe nada. A tarefa agendada tentará de novo no dia útil seguinte (até o 5º), retomando a extração aberta.

## O que esta rotina nunca faz

Reexportar um relatório já disparado; escolher no palpite entre duas mensagens do RRA; gravar no banco em rodada
manual sem o "pode importar"; enviar e-mail; alterar o modelo-base; mexer em abas fora das 4 de base.
````

- [ ] **Step 6: Rodar e ver passar**

Run: `py -m pytest tests/test_runbooks.py -q`
Expected: passam. (Se o teste reclamar de um comando citado, corrija o **roteiro**.)

- [ ] **Step 7: Commit**

```bash
git add runbooks scripts/capturar_tela.ps1 tests/test_runbooks.py
git commit -m "docs: roteiros da extração e da rotina mensal para o Claude, e script de print" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Documentação e regressão completa

**Files:**
- Modify: `README.md`, `docs/superpowers/HANDOFF-execucao-planos.md`, `docs/superpowers/specs/2026-10-09-papaki-esferas-mvp-design.md`

- [ ] **Step 1: README**

Em `README.md`:
- Em "Banco pausado", troque o último período ("A restauração automática só vem no Plano 4.") por:
  "Com `--restaurar-banco` e `SUPABASE_ACCESS_TOKEN` no `.env` (chave pessoal do Supabase: Conta → Access Tokens), o comando confirma na API que o projeto está pausado, pede a restauração, espera até 10 minutos e segue. Sem a opção, continua só avisando."
- Na tabela de opções, acrescente as linhas `--extracao`, `--restaurar-banco`, e troque `--modelo`/`--saida` para "opcional quando `PAPAKI_MODELO_<PRODUTORA>` e `PAPAKI_DADOS_DIR` estão no `.env`".
- Acrescente a seção **"Rotina mensal"**: resumo de 6 linhas (agenda no 2º dia útil, `extracao ...` registra tela/disparo/arquivos, roteiros em `runbooks/`, `importar --extracao`), a estrutura de pastas (`<PAPAKI_DADOS_DIR>/extracoes/<AAAA-MM>/<produtora>/<data-hora>/` e `<PAPAKI_DADOS_DIR>/apuracoes/<produtora>/<AAAA-MM>/`) e o aviso "a pasta de dados fica fora do projeto".
- Em "Códigos de saída", acrescente a observação: os subcomandos `extracao` usam `0` ok, `3` erro e `4` "aguarde" (só no `rra-escolher`).

- [ ] **Step 2: HANDOFF**

Em `docs/superpowers/HANDOFF-execucao-planos.md`, na seção "Estado atual", acrescente no topo o item: "**Plano 4 ESCRITO e aprovado (14/10/2026); código das Tasks 1–11 pronto e testado; ensaio assistido (Task 12) e tarefa agendada (Task 13) pendentes.** Módulos novos: `calendario`, `ambiente`, `caminhos`, `supabase_admin`, `rra_email`, `extracao`, `cli_extracao`; roteiros em `runbooks/`." (O resultado real do ensaio só é escrito na Task 12, Step 6, trocando este item por "**Plano 4 EXECUTADO (data)**" com as contagens e o desfecho observados.)
Acrescente nas "Decisões já tomadas": "14. **Plano 4 (14/10/2026):** banco pausado acordado com `--restaurar-banco` + `SUPABASE_ACCESS_TOKEN`; pastas fora do projeto (`PAPAKI_DADOS_DIR`), modelo por `PAPAKI_MODELO_<PRODUTORA>`; dia útil = seg–sex sem feriados nacionais (Carnaval e Corpus Christi contam); roda do 2º ao 5º dia útil até a competência ficar `ok`/`com_alertas`; o sistema não envia e-mail."

- [ ] **Step 3: Spec do MVP**

Na spec `2026-10-09-papaki-esferas-mvp-design.md`, §10, na linha "**Supabase pausado** (confirmado em 09/10/2026)", acrescente ao fim: "A restauração automática (opção `--restaurar-banco`, com a chave pessoal) foi construída no Plano 4."

- [ ] **Step 4: Regressão completa**

Run (rápidos): `py -m pytest tests -q -m "not real and not excel"`
Expected: 399 originais + novos, todos passam.

Run (reais, usa o Excel, ~4 min; não encerre `EXCEL.EXE` que não seja seu): `py -m pytest tests/test_cli_real.py tests/test_build_real.py`
Expected: 20 passam. Se algum falhar, investigue antes de seguir (provável causa: opções do `importar` que viraram opcionais ou o fuso).

- [ ] **Step 5: Commit**

```bash
git add README.md docs/superpowers
git commit -m "docs: README, HANDOFF e spec com a rotina mensal do Plano 4" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Preparar a pasta de dados e fazer o ensaio assistido (com o usuário)

Esta task **não é automática**: cada passo marcado **[PEDIR O OK]** só acontece depois do "sim" do usuário no chat.

- [ ] **Step 1: Pasta de dados e modelo (fora do projeto)**

Sugerir ao usuário `<pasta de dados>` (ou outro local à escolha dele). Com o OK: criar `<pasta>\modelos\untitled\` e copiar `Modelo DRE - Untitled - base.xlsx` para lá (renomeie para `Modelo DRE - Untitled - base 2026.xlsx`). **Acrescentar** ao `.env`, sem abrir o arquivo, as linhas (anexando com `printf >>`; confira antes com `grep -c "^PAPAKI_DADOS_DIR"` que não existem):

```
PAPAKI_DADOS_DIR=<pasta de dados>
PAPAKI_MODELO_UNTITLED=<pasta de dados>\modelos\untitled\Modelo DRE - Untitled - base 2026.xlsx
```

- [ ] **Step 2: Teste de leitura da API do Supabase [PEDIR O OK]**

O usuário confirma que colou o `SUPABASE_ACCESS_TOKEN` no `.env`. Com o OK, rode **só leitura** do estado do projeto (sem imprimir token nem ref):

```bash
py -c "from papaki_dre import ambiente, supabase_admin as s; e=ambiente.ler_env(); a=s.AdminSupabase(e['SUPABASE_ACCESS_TOKEN'], s.ref_do_projeto(e['SUPABASE_URL'])); print('estado:', a.estado())"
```

Expected: `estado: ACTIVE_HEALTHY`. Se vier "chave recusada", peça ao usuário um token novo. **Não** pause o projeto para testar a restauração.

- [ ] **Step 3: Teste do print [PEDIR O OK]**

Avisar que o print captura **a tela inteira**. Com o OK e o Chrome à frente numa página sem dados sensíveis (por exemplo, a tela inicial do Esferas), rode `scripts\capturar_tela.ps1` para um PNG de teste dentro de `papaki_dados\teste\` e confirme que o arquivo abre. Pergunte ao usuário se a **tela bloqueada** na hora da tarefa agendada é um problema real; se for, decidir com ele uma prova alternativa (por exemplo, texto lido da página) **antes** da Task 13 e abrir uma task nova para isso.

- [ ] **Step 4: Simulação do calendário**

```bash
py -m papaki_dre agenda --produtora untitled --hoje 2026-11-04
py -m papaki_dre agenda --produtora untitled --hoje 2026-11-02
```

Expected: o primeiro `{"rodar": true, "competencia": "2026-10", ...}`; o segundo `"rodar": false`.

- [ ] **Step 5: Ensaio supervisionado da extração [PEDIR O OK]**

Com o usuário olhando o Chrome, o Claude segue `runbooks/extracao-esferas.md` para a competência `2026-09` (período até 30/09/2026), com **um** disparo por relatório. Decidir com o usuário, antes de começar, se o ensaio usa uma **competência de ensaio** (ex.: `2026-08`) ou `2026-09`: como setembro já está no banco, uma extração nova com arquivos diferentes viraria a foto vigente. Importar **sempre** primeiro com `--sem-banco --saida <pasta de ensaio>`:

```bash
py -m papaki_dre importar --produtora untitled --competencia 2026-09 --extracao <pasta> --sem-banco --saida <pasta_apuracao>\ensaios\<AAAA-MM-DD>
```

Comparar as contagens (BASE/DESPESAS/RECEITA/Comissões) com as de setembro (72/1.110/484/649 na 1ª extração; 72/1.101/481/649 na 2ª) e relatar ao usuário: o que passou, o que não passou, o que mudou para ele. Registrar o que o Chrome fez de diferente do roteiro e **corrigir o roteiro** (commit `docs:` à parte).

- [ ] **Step 6: Gravar no banco [PEDIR O "pode importar"]**

Só com o "pode importar" explícito do usuário, repetir o `importar` sem `--sem-banco`. Relatar o resultado (importação, vigente, alertas) e atualizar a linha do Plano 4 no HANDOFF com o resultado real. Commit `docs: HANDOFF com o ensaio assistido do Plano 4`.

---

### Task 13: Criar a tarefa agendada (só com autorização)

- [ ] **Step 1: Teste curto e supervisionado [PEDIR O OK]**

Antes de criar a tarefa, testar com o usuário olhando que uma sessão agendada **consegue**: abrir `papaki.esferas.com.br`, `untitled.esferas.com.br` e o Gmail no Chrome logado, sem pedir permissão de site a cada passo, e rodar `py -m papaki_dre agenda ...`. Sem exportar nada. Se pedir permissões, o usuário as concede e elas ficam registradas; se não funcionar sem alguém na tela, relatar e parar (plano B: disparo manual com o mesmo roteiro).

- [ ] **Step 2: Criar a tarefa [PEDIR O OK]**

Com o OK, criar a tarefa agendada do Claude (ferramenta de tarefas agendadas do aplicativo) com:
- agenda: dias úteis (segunda a sexta), de manhã (por exemplo 08:00 no horário local; o `agenda` decide se hoje é dia de rodar);
- texto: "Siga `runbooks/rotina-mensal.md` na pasta do projeto `<pasta do projeto>`, forma agendada. Se `agenda` disser que não é para rodar, termine em silêncio.";
- permissões de site e de comando já concedidas no Step 1.

Mostrar ao usuário o nome, o horário e como desligar a tarefa. O disparo manual continua sendo "rode o mês AAAA-MM" numa sessão do Claude na pasta do projeto.

- [ ] **Step 3: Fechar o plano**

Atualizar o HANDOFF (Plano 4 concluído; Plano 5 é o próximo e **não** começa sem pedido), commit `docs: HANDOFF com a tarefa agendada do Plano 4`. Dizer ao usuário em 2–3 linhas o que passou, o que não passou e o que mudou para ele.

---

## Auto-revisão contra a spec

| Seção da spec | Onde está no plano |
|---|---|
| §2.1 banco pausado automático com chave | Tasks 3 e 8 (+ leitura real na Task 12.2) |
| §2.2 agendamento + disparo manual | Tasks 1, 6, 10, 13 |
| §2.3 / §4.4 e-mail do RRA pelo Gmail do Chrome | Tasks 4, 5 (`escolher_rra`), 10 |
| §2.4 sistema não envia e-mail | Task 10 (`rotina-mensal.md`: aviso, nunca e-mail) |
| §2.5 dia útil | Task 1 |
| §2.6 / §4.7 pastas definidas pelo ambiente | Task 2, Task 7 (`_entradas`) |
| §2.7 modelo base por produtora no mesmo local | Task 2 (`modelo_da_produtora`) |
| §3 fluxo mensal (agenda 2º–5º dia, refazer no dia seguinte) | Tasks 1, 6, 10 |
| §4.3 travas e disparo único, retomada | Task 5 e Task 6 |
| §4.5 `importar --extracao`, fuso, órfãs, `RESULTADO.md` | Tasks 7 e 9 |
| §4.6 roteiros | Task 10 |
| §5 segurança (token fora de logs, raiz fora do projeto) | Tasks 2 e 3 |
| §6 testes só com dados inventados | todas as tasks |
| §7 rodada assistida, teste da tarefa agendada, restauração real não forçada | Tasks 12 e 13 |

**Pontos que a spec deixava em aberto e este plano fecha (avise o usuário):** o disparo é gravado **antes** do clique e o RRA tem janela de 10 s antes a 180 s depois (spec ajustada); `--restaurar-banco` é opção explícita (spec ajustada); a pasta da extração inclui a produtora (spec ajustada); `prova_caminho` guarda o caminho relativo à raiz de dados; a prova é um print da **tela inteira** (risco: tela bloqueada, tratado na Task 12.3).
