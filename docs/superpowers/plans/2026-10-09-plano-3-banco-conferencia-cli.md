# Plano 3 — Conferência, Banco (Supabase) e Comando `importar` Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ligar tudo de ponta a ponta: um comando `python -m papaki_dre importar ...` que sanitiza os 4 brutos, grava as bases tratadas no Supabase como uma **foto** (importação) vigente, gera a cópia do Modelo DRE e emite o relatório de conferência.

**Architecture:** `report.py` consolida o resultado em um dict (status `ok` / `com_alertas` / `bloqueada`) e em Markdown. `db.py` isola o acesso ao Supabase atrás de uma classe `Repository` que só usa uma interface mínima do cliente (`table().insert/select/update/eq/in_().execute()` e `rpc()`), testável com um cliente falso. A migração SQL cria o esquema; a promoção da foto vigente é uma função SQL atômica. `cli.py` orquestra: validar argumentos → hashes → (banco: verificar acesso — pausado/inacessível sai com código 3 e aviso; ignorar se idêntico; criar importação) → sanitizar → carregar Excel → conferência → (banco: arquivos, linhas, exclusões, cargas) → promover se não bloqueada. Um `ContractError` na sanitização ou na carga do Excel vira importação **bloqueada** (nada publicado); qualquer outra exceção vira **falhou**; argumento inválido (`UsageError`) não grava nada.

**Tech Stack:** Python ≥ 3.11, supabase-py, python-dotenv, pytest, Supabase (Postgres 15+).

**Pré-requisitos:** Planos 1 e 2 concluídos. Um **projeto Supabase dedicado no plano gratuito** (decisão do usuário de 09/10/2026; criar o projeto exige confirmar nome e organização com o usuário — ver Task 2).

**Interface consumida do Plano 1 (código final, 138 testes):** `Sanitized` (`report`, `header`, `rows` — valores tipados com `Formula`/`XlError`, só para o Excel —, `db_rows` — só valores, lista distinta, contrato de entrada do banco —, `source_rows` — linha de origem de cada linha mantida —, `raw_count`, `technical` e `exclusions` — listas separadas de `Removal(source_row, reason)` —, `warnings`, `title`, `kept`, `balanced()`); `ContractError`, `Formula`, `XlError`, `Removal` (`types.py`); `DEST` e os motivos `R_BLANK`, `R_FOOTER`, `R_JOB`, `R_LOC`, `R_NUCLEO`, `R_PESQ`, `R_MODELO`, `R_H` (`config.py`); `sanitize_cp(path)`, `sanitize_cr(path)`, `sanitize_comissoes(path)`, `sanitize_rra(xls_path, *, work_dir)` (`sanitize/`). **Não existem** totais de controle (`ControlTotal`, `Sanitized.controls`, `CR_CONTROL_COLUMNS`, `CR_FOOTER_SHIFT`, `TOLERANCIA_CONTROLE` foram eliminados em 09/10/2026) nem `R_ODD` (registro com nº de colunas inesperado levanta `ContractError`). `sanitize_rra` converte o XLS com `convert_xls_to_xlsx`, que grava `<work_dir>/<stem do bruto>-<hash12 do conteúdo>.xlsx` (`xlsx_cache_path`) e devolve esse caminho; nada neste plano depende de um nome fixo do XLSX convertido.

**Interface consumida do Plano 2:** `build_workbook(template, out_path, results) -> BuildReport`, `BuildReport(sheets: dict[str, SheetReport], parts: dict)`, `SheetReport(sheet, first_row, last_row, rows, missing_formulas, orphan_formula_rows, date_style_warnings, total_row=None)` (`build.py`). `build_workbook` levanta `ContractError` quando o modelo não tem a fórmula de classificação a estender (`AP:AQ`/`AC` — "classificação ausente") ou quando a região de destino não está vazia. `missing_formulas` é a cobertura da spec §8 lida do arquivo gravado (`formula_coverage`): células sem fórmula em `AP:AQ` (DESPESAS), `AC` (RECEITA), `L:N` e `O:Q` (Comissões) em qualquer linha carregada, e lacunas na linha `TOTAL` da `BASE` (rótulo em `A` ou alguma das 35 somas); lista vazia = cobertura completa. `total_row` (só na `BASE`) tem padrão `None`, então os construtores posicionais deste plano continuam válidos.

**Atualização de 13/10/2026 (Plano 2 executado):** `SheetReport` ganhou, no fim e com padrão, `overwritten_constants: int = 0` (constantes digitadas — e, desde a revisão final de 13/10, fórmulas de outra família, como um resumo `=COUNTA(...)` ou `<f/>` vazio — em `DESPESAS!AP:AQ` / `RECEITA!AC` dentro das linhas carregadas que foram substituídas pela fórmula do modelo — decisão "opção a"; 439 no modelo de setembro) e `autofilter_extended: bool = False` (o `<autoFilter>` da aba foi estendido até a última linha carregada). O relatório de conferência deve **exibir** os dois (como informação/alerta) e **não** bloquear por eles. Também como **alerta** (nunca bloqueio, decisão 8): `#N/D` em `AP:AQ` = classificação sem correspondência no De-Para (327 células no arquivo de setembro, incluindo linhas órfãs); só aparece com os valores recalculados pelo Excel, já que a cópia gravada tem `fullCalcOnLoad` e caches antigos. `missing_formulas` passou a contar também `<f/>` vazio e filha compartilhada sem mestre.

**AJUSTES DO PRÉ-VOO (13/10/2026) — valem sobre o código dos blocos abaixo, que continuam como base; onde houver conflito, vale esta seção:**
- **Comando:** use `py -m pytest ...` e `py -m papaki_dre ...` (o `python` do PATH não funciona nesta máquina).
- **Novos campos da conferência (decisões 10–12 do HANDOFF), só informativos, nunca bloqueiam e não mudam o status:** `conference()` guarda em `abas[<aba>]` as chaves `constantes_substituidas` (`SheetReport.overwritten_constants`), `formulas_unificadas` (`unified_formula_rows`) e `autofiltro_estendido` (`autofilter_extended`), e uma lista `informacoes` (texto) que o Markdown mostra numa seção "Informações" e na tabela das abas. Só contagens, sem listar linhas.
- **Alerta `#N/D` em `DESPESAS!AP:AQ` (decisão 11 + pré-voo b, aprovado pelo usuário):** novo módulo `src/papaki_dre/recalc.py` com `contar_nd(xlsx, aba, colunas, primeira, ultima) -> int`: copia o arquivo gerado para uma **pasta temporária** (nunca recalcula o entregue), abre numa instância **própria** do Excel por COM (`DispatchEx`, invisível, sem alertas, `com_retry` em toda chamada), faz `CalculateFull`, lê `Range(AP{primeira}:AQ{ultima}).Value2` e conta os inteiros iguais a `-2146826246` (código COM de `#N/D`); fecha sem salvar, encerra **só a instância que ele iniciou** e apaga a cópia. `conference(..., na_count=None, na_error=None)`: `na_count > 0` vira alerta ("N células com #N/D em AP:AQ: classificação sem correspondência no De-Para; a classificação é regra da Papaki; só alerta"); `na_error` (o COM falhou) vira alerta de aviso ("não foi possível recalcular no Excel para contar #N/D: <motivo>"). Nunca bloqueia. `cli.py` chama `contar_nd` depois do `build_workbook` (importável como `cli.contar_nd` para o teste trocar por dublê) dentro de `try/except Exception`; opção `--sem-recalculo` pula (então nem contagem nem aviso, só uma linha em `informacoes`: "contagem de #N/D não solicitada"). Erros de célula via COM chegam como inteiros < -2146820000; só `#N/D` é contado.
- **`build_workbook`** agora aceita `master_formula_rows=` (padrão = `config.MASTER_FORMULA_ROWS`); o `cli.py` não precisa passar. `SheetReport` tem os 3 campos novos acima, com padrão.
- **Números esperados com os brutos reais (Plano 2, build real):** BASE 72, DESPESAS 1.110, RECEITA 484, Comissões 649; órfãs 47 (DESPESAS) / 179 (RECEITA); `overwritten_constants` 439 e `unified_formula_rows` 846 em DESPESAS; `#N/D` em AP:AQ ≈ 231 (linhas 3–1112). Divergência deles na T5 = parar e investigar.
- **Segurança de segredos:** o `.env` nunca é aberto nem impresso; testes de conexão só informam "conectou" ou o tipo do erro, sem URL/chave. O projeto do Supabase é o ref `<ref do projeto>`; as ferramentas MCP do servidor `supabase` (escopo do projeto) apontam para ele; o conector geral do Supabase só lista outro projeto e **não** deve ser usado.
- `.gitignore` ganha `.mcp.json` e `.claude/` (configuração local, pode conter credenciais).

**Decisões confirmadas pelo usuário (09/10/2026):**
1. **Rodada parada por erro de contrato é registrada.** A importação é criada no banco **antes** de sanitizar; se ocorrer `ContractError` (sanitização ou montagem do Excel), ela fica `bloqueada` com o motivo, **nenhuma** linha é gravada, nem `arquivos_brutos`/proveniência, e o comando sai com código 2 (T4).
2. **Rodada bloqueada pela conferência** (contagens que não fecham, mantidas ≠ carregadas, linha carregada sem fórmula do modelo): linhas e exclusões **ficam gravadas** no banco, junto com `arquivos_brutos` e `cargas_excel`, **sem** se tornarem vigentes; a importação fica `bloqueada`, a saída vira `*_BLOQUEADO.xlsx` e o código é 2 (T4).
3. **Fórmulas de classificação/cálculo em todas as linhas carregadas.** Só se detecta classificação **ausente** (modelo sem fórmula para estender → `ContractError`; linha carregada sem fórmula → `missing_formulas`); ambas bloqueiam. Nenhum outro critério de classificação entra no MVP.
4. **Supabase pausado** (plano gratuito, ~7 dias sem uso): o `importar` apenas **detecta** banco inacessível/pausado antes de gravar qualquer coisa, avisa `banco pausado: restaure o projeto e rode de novo` e sai com código 3. A restauração automática do projeto é do **Plano 4** (automação mensal) (T3 `verificar_acesso`, T4).

**Spec:** `docs/superpowers/specs/2026-10-09-papaki-esferas-mvp-design.md` (§4 dados, §5 falhas, §6 sanitização, §8 conferência, §9 aceite, §10 decisões de 09/10/2026).

**Decisão de simplificação (confirmada em 09/10/2026; spec §10, "Decisões"):** a spec §4 diz "dado tratado e tipado (49, 41, 28 e 14 colunas)". Para o MVP, cada uma das quatro tabelas guarda as colunas do relatório num `jsonb` com **valores tipados** (datas em ISO-8601, números como número) mais uma **chave promovida** (`cod` no CP, `num_job` no RRA e nas Comissões) e colunas de rastreio. Motivo: o RRA tem 49 colunas de tipos que só a planilha conhece; colunas tipadas fixas podem ser geradas depois, como views, sem migrar dado. A Task 7 atualiza a spec. `data_extracao` e `origem` ficam em `importacoes` (uma vez por foto), não repetidas em cada linha; as tabelas de dados levam `importacao_id`, `produtora_id` e `competencia`.

**Decisão sobre `exclusoes`:** a tabela guarda **as duas listas** do `Sanitized` — remoções técnicas (`technical`: linha vazia, rodapé de TOTAL) com `tipo = 'tecnica'` e exclusões de negócio (`exclusions`) com `tipo = 'negocio'` —, cada uma com relatório, linha de origem, `motivo` (o código `R_*` de `config.py`) e `regra` (texto da regra da spec §6). Assim a contagem bruto → técnicas → exclusões → mantidas pode ser refeita só com o banco.

---

## File Structure

```
src/papaki_dre/report.py             # conference(), render_markdown()
src/papaki_dre/db.py                 # sha256_file, json_value, row_to_json, Repository, make_client
src/papaki_dre/cli.py                # main(), importar()
src/papaki_dre/__main__.py
supabase/migrations/0001_schema.sql
.env.example
README.md
tests/fake_supabase.py               # cliente falso
tests/test_report.py
tests/test_db.py
tests/test_cli.py
tests/test_cli_real.py
```

---

### Task 1: Relatório de conferência

**Files:**
- Create: `src/papaki_dre/report.py`
- Test: `tests/test_report.py`

**Não há totais de controle** (eliminados em 09/10/2026, spec §8): as linhas de TOTAL (`R_FOOTER`) são só contadas entre as remoções técnicas; seus valores nunca são lidos, comparados nem usados para alertar ou bloquear. A conferência é feita pelas **contagens** (bruto → técnicas → exclusões por motivo → mantidas → carregadas na aba), com o invariante `Sanitized.balanced()` e a igualdade mantidas = carregadas.

Regras de status:
- **bloqueada** se as contagens de um relatório não fecham (`not s.balanced()`, ou linhas mantidas ≠ linhas carregadas na aba), se alguma linha carregada ficou sem a fórmula exigida (`SheetReport.missing_formulas`, spec §8: classificação ausente em `AC`/`AP:AQ`, `L:N`/`O:Q` das comissões ou lacuna na linha `TOTAL` da `BASE`), ou se ocorreu erro de contrato (`ContractError` do sanitizador ou do `build_workbook`, p.ex. "classificação ausente" — relatório gerado por `contract_blocked`).
- **com_alertas** se houver diferença de contagem tela × arquivo (sempre alerta, nunca erro), avisos dos sanitizadores (`Sanitized.warnings`), fórmulas órfãs abaixo dos dados ou datas gravadas em células sem formato de data.
- caso contrário **ok**.

- [ ] **Step 1: Escrever os testes que falham**

```python
from papaki_dre.build import BuildReport, SheetReport
from papaki_dre.config import R_BLANK, R_FOOTER, R_JOB, R_LOC
from papaki_dre.report import contract_blocked, conference, render_markdown
from papaki_dre.types import Removal, Sanitized


def _s(report, kept=2, technical=1, exclusions=(), warnings=()):
    return Sanitized(
        report=report, header=["a"], rows=[["x"]] * kept, db_rows=[["x"]] * kept,
        source_rows=list(range(kept)), raw_count=kept + technical + len(exclusions),
        technical=[Removal(i, R_FOOTER) for i in range(technical)],
        exclusions=list(exclusions), warnings=list(warnings))


def _results(**over):
    base = {k: _s(k) for k in ("rra", "cp", "cr", "comissoes")}
    base.update(over)
    return base


def _build(**over):
    sheets = {
        "BASE": SheetReport("BASE", 5, 6, 2),
        "DESPESAS": SheetReport("DESPESAS", 3, 4, 2),
        "RECEITA": SheetReport("RECEITA", 2, 3, 2),
        "REAL X ORÇADO COMISSÕES": SheetReport("REAL X ORÇADO COMISSÕES", 2, 3, 2),
    }
    sheets.update(over)
    return BuildReport(sheets=sheets, parts={"changed": ["a"], "removed": ["xl/calcChain.xml"], "added": []})


def test_clean_run_is_ok():
    conf = conference(_results(), _build())
    assert conf["status"] == "ok" and not conf["alertas"] and not conf["bloqueios"]


def test_footer_rows_are_only_counted_never_checked():
    # 3 rodapés de TOTAL removidos como técnicos: entram na contagem e em mais nada
    cp = _s("cp", technical=3)
    conf = conference(_results(cp=cp), _build())
    assert conf["status"] == "ok"
    assert conf["relatorios"]["cp"]["tecnicas_por_motivo"] == {R_FOOTER: 3}
    assert "controles" not in conf["relatorios"]["cp"]


def test_count_flow_is_reported():
    ex = [Removal(10, R_JOB), Removal(11, R_LOC), Removal(12, R_LOC)]
    cp = _s("cp", kept=2, technical=1, exclusions=ex)
    cp.technical.append(Removal(99, R_BLANK))
    cp.raw_count += 1
    r = conference(_results(cp=cp), _build())["relatorios"]["cp"]
    assert r["registros_bruto"] == 7
    assert r["removidas_tecnicas"] == 2 and r["tecnicas_por_motivo"] == {R_FOOTER: 1, R_BLANK: 1}
    assert r["registros_de_dados"] == 5
    assert r["excluidas"] == {R_JOB: 1, R_LOC: 2}
    assert r["mantidas"] == 2 and r["carregadas"] == 2


def test_screen_count_difference_is_an_alert_not_an_error():
    cp = _s("cp", kept=3, technical=2)              # 3 registros de dados
    conf = conference(_results(cp=cp), _build(DESPESAS=SheetReport("DESPESAS", 3, 5, 3)),
                      screen_counts={"cp": 5})
    assert conf["status"] == "com_alertas" and not conf["bloqueios"]
    assert any("tela mostrava 5" in a and "tem 3" in a for a in conf["alertas"])


def test_matching_screen_count_is_silent():
    cp = _s("cp", kept=3, technical=2)
    conf = conference(_results(cp=cp), _build(DESPESAS=SheetReport("DESPESAS", 3, 5, 3)),
                      screen_counts={"cp": 3})
    assert conf["status"] == "ok"


def test_unbalanced_counts_block():
    s = _s("rra")
    s.raw_count += 5
    conf = conference(_results(rra=s), _build())
    assert conf["status"] == "bloqueada" and "não fecham" in conf["bloqueios"][0]


def test_kept_rows_must_all_be_loaded():
    conf = conference(_results(), _build(RECEITA=SheetReport("RECEITA", 2, 2, 1)))
    assert conf["status"] == "bloqueada"
    assert "RECEITA" in conf["bloqueios"][0] and "carregadas" in conf["bloqueios"][0]


def test_missing_model_formulas_block():
    sr = SheetReport("DESPESAS", 3, 4, 2, missing_formulas=["AP4", "AQ4"])
    conf = conference(_results(), _build(DESPESAS=sr))
    assert conf["status"] == "bloqueada" and "DESPESAS" in conf["bloqueios"][0]


def test_coverage_gaps_in_commissions_and_base_total_block():
    com = SheetReport("REAL X ORÇADO COMISSÕES", 2, 3, 2, missing_formulas=["L3", "M3", "N3"])
    base = SheetReport("BASE", 5, 6, 2, missing_formulas=["A8", "J8"], total_row=8)
    conf = conference(_results(), _build(BASE=base, **{"REAL X ORÇADO COMISSÕES": com}))
    assert conf["status"] == "bloqueada"
    text = "\n".join(conf["bloqueios"])
    assert "BASE" in text and "J8" in text and "COMISSÕES" in text and "L3" in text


def test_orphans_dates_and_warnings_are_alerts():
    sr = SheetReport("RECEITA", 2, 3, 2, orphan_formula_rows=179, date_style_warnings={"B": 2})
    warn = "comissoes: 3 fórmula(s) de linhas mantidas sem valor em cache (ex.: L9)"
    conf = conference(_results(comissoes=_s("comissoes", warnings=[warn])), _build(RECEITA=sr))
    assert conf["status"] == "com_alertas" and not conf["bloqueios"]
    text = "\n".join(conf["alertas"])
    assert "179" in text and "sem valor em cache" in text and "formato de data" in text


def test_contract_error_is_a_blocked_conference():
    conf = contract_blocked("DESPESAS!AP: classificação ausente", meta={"competencia": "2026-09"})
    assert conf["status"] == "bloqueada"
    assert conf["bloqueios"] == ["erro de contrato: DESPESAS!AP: classificação ausente"]
    md = render_markdown(conf)
    assert "bloqueada" in md and "classificação ausente" in md and "2026-09" in md


def test_markdown_mentions_status_and_every_report():
    md = render_markdown(conference(_results(), _build(), meta={"produtora": "untitled", "competencia": "2026-09"}))
    assert "# Conferência da importação" in md and "untitled" in md and "2026-09" in md
    for label in ("RRA", "CP", "CR", "Comissões"):
        assert label in md
    assert "**Status: ok**" in md and "Carregadas" in md
    assert "Controles" not in md
```

- [ ] **Step 2: Rodar e ver falhar**

Run: `py -m pytest tests/test_report.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.report'`

- [ ] **Step 3: Implementar `report.py`**

```python
from __future__ import annotations

from collections import Counter

from .build import BuildReport
from .config import DEST
from .types import Sanitized

ORDER = ("rra", "cp", "cr", "comissoes")
LABEL = {
    "rra": "RRA (BASE)",
    "cp": "CP (DESPESAS)",
    "cr": "CR (RECEITA)",
    "comissoes": "Comissões (REAL X ORÇADO COMISSÕES)",
}


def conference(results: dict[str, Sanitized], build: BuildReport, *,
               screen_counts: dict[str, int] | None = None, meta: dict | None = None) -> dict:
    screen_counts = screen_counts or {}
    alertas: list[str] = []
    bloqueios: list[str] = []
    relatorios = {}

    for key in ORDER:
        s = results[key]
        label = LABEL[key]
        aba = DEST[key]["sheet"]
        sr = build.sheets.get(aba)
        carregadas = sr.rows if sr is not None else None
        registros = s.raw_count - len(s.technical)          # o que a tela contaria
        tela = screen_counts.get(key)
        relatorios[key] = {
            "registros_bruto": s.raw_count,
            "removidas_tecnicas": len(s.technical),
            "tecnicas_por_motivo": dict(Counter(t.reason for t in s.technical)),
            "registros_de_dados": registros,
            "excluidas": dict(Counter(e.reason for e in s.exclusions)),
            "mantidas": s.kept,
            "carregadas": carregadas,
            "avisos": list(s.warnings),
            "tela": tela,
        }
        if tela is not None and tela != registros:
            alertas.append(f"{label}: a tela mostrava {tela} registros; o arquivo tem {registros} "
                           f"(alerta, não é erro)")
        alertas.extend(f"{label}: {w}" for w in s.warnings)
        if not s.balanced():
            bloqueios.append(f"{label}: as contagens não fecham (bruto {s.raw_count} ≠ mantidas {s.kept} "
                             f"+ técnicas {len(s.technical)} + exclusões {len(s.exclusions)})")
        if carregadas != s.kept:
            bloqueios.append(f"{label}: {s.kept} linhas mantidas, mas {carregadas} carregadas na aba {aba}")

    abas = {}
    for name, sr in build.sheets.items():
        abas[name] = {"linhas": sr.rows, "primeira": sr.first_row, "ultima": sr.last_row,
                      "sem_formula": len(sr.missing_formulas), "formulas_orfas": sr.orphan_formula_rows,
                      "datas_sem_formato": sr.date_style_warnings}
        if sr.missing_formulas:
            bloqueios.append(f"{name}: {len(sr.missing_formulas)} linhas carregadas sem a fórmula do "
                             f"modelo (ex.: {', '.join(sr.missing_formulas[:3])})")
        if sr.orphan_formula_rows:
            alertas.append(f"{name}: {sr.orphan_formula_rows} linhas com fórmulas do modelo abaixo "
                           f"dos dados carregados (mantidas)")
        if sr.date_style_warnings:
            alertas.append(f"{name}: datas gravadas em células sem formato de data: {sr.date_style_warnings}")

    status = "bloqueada" if bloqueios else ("com_alertas" if alertas else "ok")
    return {"status": status, "meta": meta or {}, "relatorios": relatorios, "abas": abas,
            "partes_alteradas": build.parts, "alertas": alertas, "bloqueios": bloqueios}


def contract_blocked(error: str, *, meta: dict | None = None) -> dict:
    """Conferência de uma rodada interrompida por ContractError (cabeçalho/colunas divergentes,
    registro com nº de colunas inesperado, erro inesperado de célula, classificação ausente,
    modelo sujo). Nada é publicado."""
    return {"status": "bloqueada", "meta": meta or {}, "relatorios": {}, "abas": {},
            "partes_alteradas": {}, "alertas": [], "bloqueios": [f"erro de contrato: {error}"]}


def render_markdown(conf: dict) -> str:
    meta = conf["meta"]
    out = ["# Conferência da importação", ""]
    for k in ("produtora", "competencia", "periodo", "data_extracao", "importacao_id"):
        if k in meta:
            out.append(f"- **{k}:** {meta[k]}")
    out += ["", f"**Status: {conf['status']}**", ""]
    if conf["bloqueios"]:
        out += ["## Bloqueios (a rodada não é publicada)", *[f"- {b}" for b in conf["bloqueios"]], ""]
    if conf["alertas"]:
        out += ["## Alertas", *[f"- {a}" for a in conf["alertas"]], ""]
    keys = [k for k in ORDER if k in conf["relatorios"]]
    if keys:
        out += ["## Contagens", "",
                "| Relatório | Bruto | Técnicas | Excluídas | Mantidas | Carregadas | Tela |",
                "|---|---:|---:|---:|---:|---:|---:|"]
        for key in keys:
            r = conf["relatorios"][key]
            tela = "" if r["tela"] is None else r["tela"]
            out.append(f"| {LABEL[key]} | {r['registros_bruto']} | {r['removidas_tecnicas']} | "
                       f"{sum(r['excluidas'].values())} | {r['mantidas']} | {r['carregadas']} | {tela} |")
        out += ["", "### Remoções por motivo", ""]
        for key in keys:
            r = conf["relatorios"][key]
            for motivo, n in r["tecnicas_por_motivo"].items():
                out.append(f"- {LABEL[key]}: técnica `{motivo}` — {n}")
            for motivo, n in r["excluidas"].items():
                out.append(f"- {LABEL[key]}: exclusão `{motivo}` — {n}")
    if conf["abas"]:
        out += ["", "## Abas carregadas", "",
                "| Aba | Linhas | Faixa | Sem fórmula | Fórmulas órfãs |", "|---|---:|---|---:|---:|"]
        for name, a in conf["abas"].items():
            out.append(f"| {name} | {a['linhas']} | {a['primeira']}–{a['ultima']} | "
                       f"{a['sem_formula']} | {a['formulas_orfas']} |")
    if conf["partes_alteradas"]:
        out += ["", "## Partes do arquivo alteradas", ""]
        out += [f"- `{p}`" for p in conf["partes_alteradas"].get("changed", [])]
        for p in conf["partes_alteradas"].get("removed", []):
            out.append(f"- removida: `{p}`")
    out.append("")
    return "\n".join(out)
```

- [ ] **Step 4: Rodar e ver passar**

Run: `py -m pytest tests/test_report.py -v`
Expected: 12 passed

- [ ] **Step 5: Commit**

```bash
git add src/papaki_dre/report.py tests/test_report.py && git commit -m "feat: relatório de conferência por contagens com status ok/com_alertas/bloqueada"
```

---

### Task 2: Esquema do Supabase (migração SQL)

**Files:**
- Create: `supabase/migrations/0001_schema.sql`

- [ ] **Step 1: Criar a migração**

```sql
-- MVP Papaki: importação das bases do Esferas. Todas as tabelas com RLS ligado e SEM políticas:
-- só a chave de serviço (service_role) acessa. Dados reais contêm CPF/CNPJ e dados bancários.

create table public.produtoras (
  id         text primary key,
  nome       text not null,
  localidade text not null
);
insert into public.produtoras (id, nome, localidade)
values ('untitled', 'Untitled', 'UTLD PRODUÇÃO DE FILMES LTDA');

create table public.importacoes (
  id             uuid primary key default gen_random_uuid(),
  produtora_id   text not null references public.produtoras (id),
  competencia    text not null check (competencia ~ '^\d{4}-(0[1-9]|1[0-2])$'),
  periodo_inicio date not null,
  periodo_fim    date not null,
  data_extracao  timestamptz not null,
  origem         text not null,
  gatilho        text not null check (gatilho in ('manual', 'agendado')),
  status         text not null default 'em_andamento'
                 check (status in ('em_andamento', 'concluida', 'bloqueada', 'falhou')),
  vigente        boolean not null default false,
  motivo_falha   text,          -- motivo da falha ('falhou') ou do erro de contrato ('bloqueada')
  conferencia    jsonb,
  criado_em      timestamptz not null default now()
);
-- no máximo uma foto vigente por produtora e competência
create unique index importacoes_uma_vigente on public.importacoes (produtora_id, competencia) where vigente;

create table public.arquivos_brutos (
  id             uuid primary key default gen_random_uuid(),
  importacao_id  uuid not null references public.importacoes (id) on delete cascade,
  relatorio      text not null check (relatorio in ('rra', 'cp', 'cr', 'comissoes')),
  nome_arquivo   text not null,
  sha256         text not null,
  tamanho_bytes  bigint not null,
  total_tela     integer,
  filtros_tela   jsonb,
  prova_caminho  text,
  unique (importacao_id, relatorio)
);
create index arquivos_brutos_sha256 on public.arquivos_brutos (sha256);

create table public.rra_base (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null,
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  num_job       text,
  dados         jsonb not null
);
create index rra_base_importacao on public.rra_base (importacao_id);

create table public.cp_despesas (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null,
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  cod           text,
  dados         jsonb not null
);
create index cp_despesas_importacao on public.cp_despesas (importacao_id);

create table public.cr_receita (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null,
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  dados         jsonb not null
);
create index cr_receita_importacao on public.cr_receita (importacao_id);

create table public.comissoes (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null,
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  num_job       text,
  dados         jsonb not null
);
create index comissoes_importacao on public.comissoes (importacao_id);

create table public.exclusoes (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  relatorio     text not null check (relatorio in ('rra', 'cp', 'cr', 'comissoes')),
  linha_origem  integer not null,
  tipo          text not null check (tipo in ('tecnica', 'negocio')),   -- Sanitized.technical | .exclusions
  motivo        text not null,                                          -- código R_* de config.py
  regra         text not null                                           -- texto da regra (spec §6)
);
create index exclusoes_importacao on public.exclusoes (importacao_id, relatorio);

create table public.cargas_excel (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  aba           text not null,
  linha_inicial integer not null,
  linha_final   integer not null,
  linhas        integer not null,
  arquivo_saida text not null,
  sha256_saida  text not null
);

-- Promove uma importação a foto vigente de forma atômica: a anterior continua guardada, mas deixa de ser vigente.
create or replace function public.promover_importacao(p_id uuid) returns void
language plpgsql
set search_path = public
as $$
declare
  v_prod text;
  v_comp text;
begin
  select produtora_id, competencia into v_prod, v_comp
    from public.importacoes where id = p_id and status = 'em_andamento' for update;
  if not found then
    raise exception 'importação % não está em andamento', p_id;
  end if;
  update public.importacoes set vigente = false
    where produtora_id = v_prod and competencia = v_comp and vigente;
  update public.importacoes set vigente = true, status = 'concluida' where id = p_id;
end;
$$;
revoke all on function public.promover_importacao(uuid) from public, anon, authenticated;

alter table public.produtoras      enable row level security;
alter table public.importacoes     enable row level security;
alter table public.arquivos_brutos enable row level security;
alter table public.rra_base        enable row level security;
alter table public.cp_despesas     enable row level security;
alter table public.cr_receita      enable row level security;
alter table public.comissoes       enable row level security;
alter table public.exclusoes       enable row level security;
alter table public.cargas_excel    enable row level security;
```

- [ ] **Step 2: Escolher o projeto Supabase (decisão já tomada: plano gratuito)**

O usuário decidiu usar o **plano gratuito** do Supabase (09/10/2026). Use a ferramenta MCP `list_projects` para ver se já existe um projeto dedicado; se não, crie um no plano gratuito (`get_organization`/`list_organizations` para a organização, e confirme com o usuário o nome e a organização antes de criar). Registre o `project_id`.

Limites do plano gratuito que importam aqui: banco de 500 MB (as 4 bases de uma competência ocupam poucos MB, mas `exclusoes` guarda também as remoções técnicas e cada foto repete tudo: acompanhe o crescimento se importar muitas fotos) e **pausa automática após cerca de 7 dias sem atividade**. Como a rotina real roda 1 vez por mês, o projeto estará pausado na maior parte do tempo. **Decisão confirmada (09/10/2026):** no MVP o comando `importar` apenas **detecta** o banco inacessível/pausado (`Repository.verificar_acesso`, T3), antes de gravar qualquer coisa, avisa `banco pausado: restaure o projeto e rode de novo` e sai com código 3; ele **não** restaura o projeto (exigiria um token da API de gerenciamento, que não guardamos). Nas execuções manuais deste plano, quem executa restaura pelo painel ou pelo MCP (`get_project`; se pausado, `restore_project` e aguardar ficar ativo) e roda de novo. A restauração automática antes de importar é do **Plano 4** (automação mensal). A decisão de migrar para um plano pago fica para o ambiente Papaki.

- [ ] **Step 3: Aplicar a migração**

Com o `project_id` aprovado, use `apply_migration` (nome `0001_schema`, conteúdo = o SQL acima). Em seguida rode `get_advisors` (tipo `security`) e `list_tables` para confirmar: 9 tabelas, RLS ativo em todas, função `promover_importacao` presente.
Expected: nenhum alerta de segurança de "RLS desativado" ou "função exposta". Avisos informativos de "RLS ativo sem políticas" são esperados e desejados (só a chave de serviço acessa; ela fica no `.env`, fora do código e do git).

- [ ] **Step 4: Commit**

```bash
git add supabase/migrations/0001_schema.sql && git commit -m "feat: esquema do Supabase (importações, bases tratadas, remoções técnicas e exclusões, cargas) com RLS"
```

---

### Task 3: Acesso ao banco (`db.py`) com cliente falso

**Files:**
- Create: `tests/fake_supabase.py`, `src/papaki_dre/db.py`
- Test: `tests/test_db.py`

- [ ] **Step 1: Criar o cliente falso `tests/fake_supabase.py`**

```python
"""Cliente Supabase falso, em memória, com a interface mínima que o Repository usa."""
import copy
import uuid


class _Resp:
    def __init__(self, data):
        self.data = data


class _Query:
    def __init__(self, client, name):
        self.client, self.name = client, name
        self._op = None
        self._filters = []

    def insert(self, rows):
        self._op = ("insert", rows if isinstance(rows, list) else [rows])
        return self

    def select(self, cols="*"):
        self._op = ("select",)
        return self

    def update(self, values):
        self._op = ("update", values)
        return self

    def eq(self, col, val):
        self._filters.append(lambda r: r.get(col) == val)
        return self

    def in_(self, col, vals):
        self._filters.append(lambda r: r.get(col) in vals)
        return self

    def execute(self):
        rows = self.client.tables.setdefault(self.name, [])
        kind = self._op[0]
        if kind == "insert":
            out = []
            for r in self._op[1]:
                r = copy.deepcopy(r)
                if self.name in ("importacoes", "arquivos_brutos"):
                    r.setdefault("id", str(uuid.uuid4()))
                if self.name == "importacoes":
                    r.setdefault("status", "em_andamento")
                    r.setdefault("vigente", False)
                rows.append(r)
                out.append(copy.deepcopy(r))
            return _Resp(out)
        match = [r for r in rows if all(f(r) for f in self._filters)]
        if kind == "update":
            for r in match:
                r.update(self._op[1])
        return _Resp(copy.deepcopy(match))


class _Rpc:
    def __init__(self, client, name, params):
        self.client, self.name, self.params = client, name, params

    def execute(self):
        assert self.name == "promover_importacao"
        rows = self.client.tables["importacoes"]
        target = next(r for r in rows if r["id"] == self.params["p_id"])
        if target["status"] != "em_andamento":
            raise RuntimeError("importação não está em andamento")
        for r in rows:
            if (r["produtora_id"], r["competencia"]) == (target["produtora_id"], target["competencia"]) and r["vigente"]:
                r["vigente"] = False
        target["vigente"], target["status"] = True, "concluida"
        return _Resp(None)


class FakeClient:
    def __init__(self, fail_with: Exception | None = None):
        self.tables = {}
        self.fail_with = fail_with          # simula projeto pausado/inacessível: toda chamada falha

    def table(self, name):
        if self.fail_with is not None:
            raise self.fail_with
        return _Query(self, name)

    def rpc(self, name, params):
        if self.fail_with is not None:
            raise self.fail_with
        return _Rpc(self, name, params)
```

- [ ] **Step 2: Escrever os testes que falham**

```python
import datetime as dt
import hashlib

import pytest

from papaki_dre.config import R_BLANK, R_FOOTER, R_H, R_JOB, R_LOC, R_MODELO, R_NUCLEO, R_PESQ
from papaki_dre.db import BancoIndisponivel, REGRAS, Repository, json_value, row_to_json, sha256_file
from papaki_dre.types import Formula, Removal, Sanitized, XlError
from fake_supabase import FakeClient


def test_sha256_file(tmp_path):
    p = tmp_path / "a.bin"
    p.write_bytes(b"abc")
    assert sha256_file(p) == hashlib.sha256(b"abc").hexdigest()


def test_json_value_types():
    assert json_value(dt.datetime(2026, 1, 5)) == "2026-01-05T00:00:00"
    assert json_value(dt.date(2026, 1, 5)) == "2026-01-05"
    assert json_value(1.5) == 1.5 and json_value(3) == 3 and json_value("x") == "x"
    assert json_value(True) is True and json_value(None) is None
    assert json_value(float("nan")) is None
    assert json_value(XlError("#DIV/0!")) is None       # mesma política do Plano 1 (RowCopier): erro -> nulo no banco


def test_json_value_rejects_formulas():
    # db_rows é só valores; uma fórmula aqui significa que `rows` foi passado no lugar de `db_rows`
    with pytest.raises(TypeError, match="db_rows"):
        json_value(Formula("=SUM(A1:A2)"))


def test_every_removal_reason_has_a_rule():
    for r in (R_BLANK, R_FOOTER, R_JOB, R_LOC, R_NUCLEO, R_PESQ, R_MODELO, R_H):
        assert REGRAS[r]


def test_row_to_json_disambiguates_duplicate_headers():
    d = row_to_json(["Recebido", "Valor", "Recebido"], [dt.datetime(2026, 1, 6), 10.0, "Sim"])
    assert d == {"Recebido": "2026-01-06T00:00:00", "Valor": 10.0, "Recebido (2)": "Sim"}


def _san(report, rows, header=("Cód.", "Valor"), exclusions=(), technical=()):
    return Sanitized(report=report, header=list(header), rows=rows, db_rows=rows,
                     source_rows=[10 + i for i in range(len(rows))], raw_count=len(rows),
                     exclusions=list(exclusions), technical=list(technical))


def _repo():
    return Repository(FakeClient())


def _new(repo, comp="2026-09"):
    return repo.criar_importacao(produtora="untitled", competencia=comp, periodo_inicio="2026-01-01",
                                 periodo_fim="2026-09-30", data_extracao="2026-10-08T16:34:00+00:00",
                                 origem="teste", gatilho="manual")


def test_criar_importacao_starts_in_progress_and_not_current():
    repo = _repo()
    imp = _new(repo)
    row = repo.client.tables["importacoes"][0]
    assert row["id"] == imp and row["status"] == "em_andamento" and row["vigente"] is False


def test_inserir_linhas_stores_key_destination_row_and_json():
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [["untitled-1", 10.0], ["untitled-2", 20.0]])
    repo.inserir_linhas(imp, "untitled", "2026-09", "cp", s)
    rows = repo.client.tables["cp_despesas"]
    assert [r["cod"] for r in rows] == ["untitled-1", "untitled-2"]
    assert [r["linha_origem"] for r in rows] == [10, 11]
    assert [r["linha_destino"] for r in rows] == [3, 4]          # DEST["cp"]["first_row"] = 3
    assert rows[0]["dados"] == {"Cód.": "untitled-1", "Valor": 10.0}


def test_inserir_linhas_uses_db_rows_not_excel_rows():
    repo = _repo()
    imp = _new(repo)
    s = Sanitized(report="comissoes", header=["Núm. JOB", "Real"], rows=[["75131", Formula("=B2*2")]],
                  db_rows=[["75131", 40.0]], source_rows=[4], raw_count=1)
    repo.inserir_linhas(imp, "untitled", "2026-09", "comissoes", s)
    row = repo.client.tables["comissoes"][0]
    assert row["dados"] == {"Núm. JOB": "75131", "Real": 40.0} and row["num_job"] == "75131"
    assert row["linha_origem"] == 4 and row["linha_destino"] == 2      # DEST["comissoes"]["first_row"] = 2


def test_inserir_linhas_batches_large_inputs():
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [[f"c{i}", 1.0] for i in range(1201)])
    repo.inserir_linhas(imp, "untitled", "2026-09", "cp", s)
    assert len(repo.client.tables["cp_despesas"]) == 1201


def test_cr_has_no_promoted_key():
    repo = _repo()
    imp = _new(repo)
    repo.inserir_linhas(imp, "untitled", "2026-09", "cr", _san("cr", [["x", 1.0]], header=("Cliente", "Vl")))
    assert "cod" not in repo.client.tables["cr_receita"][0]


def test_inserir_exclusoes_records_technical_and_business():
    repo = _repo()
    imp = _new(repo)
    s = _san("cp", [], exclusions=[Removal(5, R_LOC)], technical=[Removal(2, R_FOOTER)])
    repo.inserir_exclusoes(imp, "cp", s)
    got = sorted((r["tipo"], r["motivo"], r["linha_origem"]) for r in repo.client.tables["exclusoes"])
    assert got == [("negocio", R_LOC, 5), ("tecnica", R_FOOTER, 2)]
    assert {r["regra"] for r in repo.client.tables["exclusoes"]} == {REGRAS[R_LOC], REGRAS[R_FOOTER]}


def test_registrar_arquivo_and_find_identical():
    repo = _repo()
    imp = _new(repo)
    hashes = {"rra": "h1", "cp": "h2", "cr": "h3", "comissoes": "h4"}
    for rel, h in hashes.items():
        repo.registrar_arquivo(imp, rel, nome=f"{rel}.x", sha256=h, tamanho=10, total_tela=None)
    assert repo.find_identical("untitled", "2026-09", hashes) is None     # ainda em andamento
    repo.promover(imp)
    assert repo.find_identical("untitled", "2026-09", hashes) == imp
    assert repo.find_identical("untitled", "2026-09", {**hashes, "cp": "OUTRO"}) is None
    assert repo.find_identical("untitled", "2026-10", hashes) is None


def test_promover_replaces_the_previous_current_one_without_deleting_it():
    repo = _repo()
    a, b = _new(repo), _new(repo)
    repo.promover(a)
    repo.promover(b)
    by_id = {r["id"]: r for r in repo.client.tables["importacoes"]}
    assert by_id[a]["vigente"] is False and by_id[a]["status"] == "concluida"
    assert by_id[b]["vigente"] is True


def test_other_competencia_keeps_its_own_current():
    repo = _repo()
    a, b = _new(repo, "2026-08"), _new(repo, "2026-09")
    repo.promover(a)
    repo.promover(b)
    assert {r["competencia"]: r["vigente"] for r in repo.client.tables["importacoes"]} == {
        "2026-08": True, "2026-09": True}


def test_marcar_falhou_and_bloqueada():
    repo = _repo()
    a, b = _new(repo), _new(repo)
    repo.marcar(a, "falhou", motivo="erro X")
    repo.marcar(b, "bloqueada", motivo="ContractError: classificação ausente",
                conferencia={"status": "bloqueada"})
    rows = {r["id"]: r for r in repo.client.tables["importacoes"]}
    assert rows[a]["status"] == "falhou" and rows[a]["motivo_falha"] == "erro X"
    assert rows[b]["status"] == "bloqueada" and rows[b]["vigente"] is False
    assert "classificação ausente" in rows[b]["motivo_falha"]


def test_registrar_carga():
    repo = _repo()
    imp = _new(repo)
    repo.registrar_carga(imp, aba="BASE", primeira=5, ultima=76, arquivo="m.xlsx", sha256="zz")
    row = repo.client.tables["cargas_excel"][0]
    assert row["linhas"] == 72 and row["aba"] == "BASE"


def test_verificar_acesso_detects_paused_or_unreachable_database():
    Repository(FakeClient()).verificar_acesso()                     # banco ativo: não levanta
    repo = Repository(FakeClient(fail_with=ConnectionError("name resolution failed")))
    with pytest.raises(BancoIndisponivel, match="name resolution failed"):
        repo.verificar_acesso()
```

- [ ] **Step 3: Rodar e ver falhar**

Run: `py -m pytest tests/test_db.py -v`
Expected: FAIL com `ModuleNotFoundError: No module named 'papaki_dre.db'`

- [ ] **Step 4: Implementar `db.py`**

```python
from __future__ import annotations

import datetime as dt
import hashlib
import math
import os
from pathlib import Path

from .config import (COMISSOES_EXCLUIR_H, DEST, MODELO_COMISSOES, NUCLEO_RRA, PESQUISA, UTLD,
                     R_BLANK, R_FOOTER, R_H, R_JOB, R_LOC, R_MODELO, R_NUCLEO, R_PESQ)
from .types import Formula, Sanitized, XlError

TABLES = {"rra": "rra_base", "cp": "cp_despesas", "cr": "cr_receita", "comissoes": "comissoes"}
# coluna promovida: (nome no banco, cabeçalho do relatório)
KEY_COLUMN = {"rra": ("num_job", "Núm. JOB"), "cp": ("cod", "Cód."), "comissoes": ("num_job", "Núm. JOB")}
BATCH = 500

# Texto da regra de cada motivo de remoção (coluna `exclusoes.regra`; spec §6).
REGRAS = {
    R_BLANK: "técnica: linha vazia",
    R_FOOTER: "técnica: rodapé de TOTAL reconhecido por rótulo e posição (valores não lidos)",
    R_JOB: f"CP: Nº job preenchido e Job sem '{PESQUISA}'",
    R_LOC: f"CP: Localidade diferente de '{UTLD}'",
    R_NUCLEO: f"RRA: Núcleo diferente de '{NUCLEO_RRA}'",
    R_PESQ: f"RRA/Comissões: job contém '{PESQUISA}'",
    R_MODELO: f"Comissões: Modelo de Orçamento diferente de '{MODELO_COMISSOES}'",
    R_H: "Comissões: coluna H em " + ", ".join(COMISSOES_EXCLUIR_H),
}


def sha256_file(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def json_value(v):
    if isinstance(v, Formula):
        raise TypeError("fórmula encontrada nos valores do banco; use Sanitized.db_rows, não .rows")
    if v is None or isinstance(v, XlError):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    if isinstance(v, float):
        return v if math.isfinite(v) else None
    if isinstance(v, (int, str)):
        return v
    return str(v)


def row_to_json(header, values) -> dict:
    out, seen = {}, {}
    for name, v in zip(header, values):
        seen[name] = seen.get(name, 0) + 1
        key = name if seen[name] == 1 else f"{name} ({seen[name]})"
        out[key] = json_value(v)
    return out


def make_client():
    """Cria o cliente a partir de SUPABASE_URL e SUPABASE_SERVICE_KEY (variáveis de ambiente ou .env)."""
    from dotenv import load_dotenv
    from supabase import create_client

    load_dotenv()
    url, key = os.environ.get("SUPABASE_URL"), os.environ.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        raise RuntimeError("defina SUPABASE_URL e SUPABASE_SERVICE_KEY (veja .env.example)")
    return create_client(url, key)


class BancoIndisponivel(RuntimeError):
    """O Supabase não respondeu (projeto pausado no plano gratuito, rede, URL errada)."""


class Repository:
    def __init__(self, client):
        self.client = client

    def verificar_acesso(self) -> None:
        """Consulta leve antes de qualquer gravação. Qualquer falha vira BancoIndisponivel:
        no plano gratuito, a causa mais comum é o projeto pausado após ~7 dias sem uso.
        Não restaura o projeto (isso é do Plano 4)."""
        try:
            self.client.table("produtoras").select("id").execute()
        except Exception as e:  # noqa: BLE001 - o cliente levanta tipos variados (httpx, postgrest)
            raise BancoIndisponivel(f"{type(e).__name__}: {e}") from e

    def _insert(self, table: str, rows: list[dict]):
        for i in range(0, len(rows), BATCH):
            self.client.table(table).insert(rows[i:i + BATCH]).execute()

    def criar_importacao(self, *, produtora, competencia, periodo_inicio, periodo_fim,
                         data_extracao, origem, gatilho) -> str:
        res = self.client.table("importacoes").insert({
            "produtora_id": produtora, "competencia": competencia, "periodo_inicio": periodo_inicio,
            "periodo_fim": periodo_fim, "data_extracao": data_extracao, "origem": origem,
            "gatilho": gatilho}).execute()
        return res.data[0]["id"]

    def registrar_arquivo(self, importacao_id, relatorio, *, nome, sha256, tamanho,
                          total_tela=None, filtros_tela=None, prova_caminho=None) -> None:
        self.client.table("arquivos_brutos").insert({
            "importacao_id": importacao_id, "relatorio": relatorio, "nome_arquivo": nome,
            "sha256": sha256, "tamanho_bytes": tamanho, "total_tela": total_tela,
            "filtros_tela": filtros_tela, "prova_caminho": prova_caminho}).execute()

    def inserir_linhas(self, importacao_id, produtora, competencia, relatorio, s: Sanitized) -> None:
        first = DEST[relatorio]["first_row"]
        key = KEY_COLUMN.get(relatorio)
        key_idx = s.header.index(key[1]) if key else None
        rows = []
        for i, (values, origem) in enumerate(zip(s.db_rows, s.source_rows)):
            row = {"importacao_id": importacao_id, "produtora_id": produtora, "competencia": competencia,
                   "linha_origem": origem, "linha_destino": first + i,
                   "dados": row_to_json(s.header, values)}
            if key:
                v = json_value(values[key_idx])
                row[key[0]] = None if v is None else str(v)
            rows.append(row)
        self._insert(TABLES[relatorio], rows)

    def inserir_exclusoes(self, importacao_id, relatorio, s: Sanitized) -> None:
        """Grava as duas listas: remoções técnicas (tipo 'tecnica') e exclusões de negócio ('negocio')."""
        rows = [{"importacao_id": importacao_id, "relatorio": relatorio, "linha_origem": r.source_row,
                 "tipo": tipo, "motivo": r.reason, "regra": REGRAS.get(r.reason, r.reason)}
                for tipo, removals in (("tecnica", s.technical), ("negocio", s.exclusions))
                for r in removals]
        self._insert("exclusoes", rows)

    def registrar_carga(self, importacao_id, *, aba, primeira, ultima, arquivo, sha256) -> None:
        self.client.table("cargas_excel").insert({
            "importacao_id": importacao_id, "aba": aba, "linha_inicial": primeira, "linha_final": ultima,
            "linhas": max(ultima - primeira + 1, 0), "arquivo_saida": arquivo, "sha256_saida": sha256}).execute()

    def find_identical(self, produtora, competencia, hashes: dict[str, str]) -> str | None:
        """Id de uma importação concluída com exatamente estes arquivos (mesmos hashes)."""
        found = self.client.table("arquivos_brutos").select("importacao_id,relatorio,sha256") \
            .in_("sha256", list(hashes.values())).execute().data
        by_imp: dict[str, dict[str, str]] = {}
        for r in found:
            by_imp.setdefault(r["importacao_id"], {})[r["relatorio"]] = r["sha256"]
        for imp_id, got in by_imp.items():
            if got != hashes:
                continue
            imp = self.client.table("importacoes").select("id,produtora_id,competencia,status") \
                .eq("id", imp_id).execute().data[0]
            if (imp["produtora_id"], imp["competencia"], imp["status"]) == (produtora, competencia, "concluida"):
                return imp_id
        return None

    def promover(self, importacao_id) -> None:
        self.client.rpc("promover_importacao", {"p_id": importacao_id}).execute()

    def marcar(self, importacao_id, status, *, motivo=None, conferencia=None) -> None:
        values = {"status": status, "vigente": False}
        if motivo is not None:
            values["motivo_falha"] = motivo
        if conferencia is not None:
            values["conferencia"] = conferencia
        self.client.table("importacoes").update(values).eq("id", importacao_id).execute()

    def salvar_conferencia(self, importacao_id, conferencia) -> None:
        self.client.table("importacoes").update({"conferencia": conferencia}).eq("id", importacao_id).execute()
```

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests/test_db.py -v`
Expected: 17 passed

- [ ] **Step 6: Commit**

```bash
git add src/papaki_dre/db.py tests/fake_supabase.py tests/test_db.py && git commit -m "feat: Repository do Supabase com fotos vigentes, dedupe por hash e cliente falso para testes"
```

---

### Task 4: Comando `importar` (orquestração)

**Files:**
- Create: `src/papaki_dre/cli.py`, `src/papaki_dre/__main__.py`
- Modify: `src/papaki_dre/config.py` (cadastro de produtoras)
- Test: `tests/test_cli.py`

Fluxo de `importar`: validar argumentos (erro de uso → código 3, nada é gravado) → hashes dos 4 brutos → (banco) `verificar_acesso`: banco pausado/inacessível → aviso `banco pausado: restaure o projeto e rode de novo`, código 3, nada gravado (a restauração automática é do Plano 4) → (banco) se já houver importação **concluída** com os mesmos 4 hashes e não houver `--forcar`, informar e sair com 0 (spec §4: mesmo hash = ignorado) → (banco) criar importação `em_andamento` → sanitizar os 4 → carregar Excel → conferência → gravar `conferencia.json`/`.md` → (banco) arquivos, linhas (`db_rows`), remoções técnicas e exclusões, cargas, conferência → se `bloqueada`, marcar e renomear a saída para `*_BLOQUEADO.xlsx`; senão promover (foto vigente por produtora + competência; a anterior deixa de ser vigente e nada é apagado).

Erros: `ContractError` na sanitização ou na carga do Excel, com ou sem banco (sanitizador: cabeçalho/colunas divergentes, registro com nº de colunas inesperado, erro inesperado de célula; `build_workbook`: classificação ausente, modelo sujo) → conferência `bloqueada` via `contract_blocked`, importação marcada `bloqueada` com o motivo, nenhuma linha gravada, código 2. Qualquer outra exceção → importação marcada `falhou` com o motivo (spec §5), código 3. Códigos de saída: 0 ok/com alertas ou duplicado; 2 bloqueada; 3 falha, banco pausado/inacessível ou erro de uso.

A importação é criada **antes** de sanitizar (decisão confirmada pelo usuário em 09/10/2026) para que toda rodada bloqueada por contrato fique registrada no banco com o motivo; o dedupe só precisa dos hashes. Bloqueio por contrato não grava linhas nem `arquivos_brutos`/proveniência; bloqueio pela conferência grava linhas, exclusões, arquivos e cargas sem torná-los vigentes (decisões confirmadas, ver o topo do plano).

- [ ] **Step 1: Cadastrar a produtora na config**

Acrescente ao final de `src/papaki_dre/config.py`:

```python
# Produtoras do fluxo. MVP: só a Untitled. Chave = produtora_id no banco.
PRODUTORAS = {
    "untitled": {"nome": "Untitled", "localidade": UTLD},
}
```

- [ ] **Step 2: Escrever os testes que falham**

Os testes substituem os sanitizadores e o `build_workbook` por dublês (monkeypatch), para exercitar só a orquestração.

```python
import datetime as dt

import pytest

from papaki_dre import cli
from papaki_dre.build import BuildReport, SheetReport
from papaki_dre.config import R_FOOTER
from papaki_dre.types import ContractError, Removal, Sanitized
from fake_supabase import FakeClient


def _san(report, kept=2):
    header = ["Núm. JOB", "Valor"] if report in ("rra", "comissoes") else ["Cód.", "Valor"]
    return Sanitized(report=report, header=header, rows=[["a", 1.0]] * kept,
                     db_rows=[["a", 1.0]] * kept, source_rows=list(range(10, 10 + kept)),
                     raw_count=kept + 1, technical=[Removal(2, R_FOOTER)])


def _build_report(missing=()):
    names = {"BASE": (5, "rra"), "DESPESAS": (3, "cp"), "RECEITA": (2, "cr"),
             "REAL X ORÇADO COMISSÕES": (2, "comissoes")}
    sheets = {n: SheetReport(n, f, f + 1, 2, missing_formulas=list(missing) if n == "DESPESAS" else [])
              for n, (f, _) in names.items()}
    return BuildReport(sheets=sheets, parts={"changed": ["x"], "removed": [], "added": []})


@pytest.fixture
def env(tmp_path, monkeypatch):
    files = {}
    for k in ("cp", "cr", "rra", "comissoes"):
        p = tmp_path / f"{k}.dat"
        p.write_bytes(k.encode())
        files[k] = p
    model = tmp_path / "modelo.xlsx"
    model.write_bytes(b"modelo")
    state = {"missing": ()}

    def fake_san(key):
        def inner(*a, **kw):
            return _san(key)
        return inner

    monkeypatch.setattr(cli, "sanitize_cp", fake_san("cp"))
    monkeypatch.setattr(cli, "sanitize_cr", fake_san("cr"))
    monkeypatch.setattr(cli, "sanitize_comissoes", fake_san("comissoes"))
    monkeypatch.setattr(cli, "sanitize_rra", fake_san("rra"))

    def fake_build(template, out, results):
        out.write_bytes(b"xlsx-" + str(len(results)).encode())
        return _build_report(missing=state.get("missing", ()))

    monkeypatch.setattr(cli, "build_workbook", fake_build)
    return {"files": files, "model": model, "out": tmp_path / "saida", "state": state}


def _argv(env, *extra):
    f = env["files"]
    return ["importar", "--produtora", "untitled", "--competencia", "2026-09",
            "--cp", str(f["cp"]), "--cr", str(f["cr"]), "--rra", str(f["rra"]),
            "--comissoes", str(f["comissoes"]), "--modelo", str(env["model"]),
            "--saida", str(env["out"]), *extra]


def test_happy_path_without_database(env):
    code = cli.main(_argv(env, "--sem-banco"))
    assert code == 0
    assert (env["out"] / "Modelo DRE - Untitled - 09_2026.xlsx").exists()
    assert (env["out"] / "conferencia.md").exists()
    assert (env["out"] / "conferencia.json").exists()


def test_happy_path_with_database_promotes_the_snapshot(env):
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "concluida" and imp["vigente"] is True
    assert imp["periodo_inicio"] == "2026-01-01" and imp["periodo_fim"] == "2026-09-30"
    assert len(client.tables["arquivos_brutos"]) == 4
    assert len(client.tables["cp_despesas"]) == 2
    assert len(client.tables["cargas_excel"]) == 4
    assert imp["conferencia"]["status"] in ("ok", "com_alertas")


def test_same_files_again_are_ignored(env):
    client = FakeClient()
    cli.main(_argv(env), client=client)
    assert cli.main(_argv(env), client=client) == 0
    assert len(client.tables["importacoes"]) == 1            # nada novo foi criado


def test_forcar_creates_a_new_snapshot_and_supersedes_the_old_one(env):
    client = FakeClient()
    cli.main(_argv(env), client=client)
    assert cli.main(_argv(env, "--forcar"), client=client) == 0
    rows = client.tables["importacoes"]
    assert len(rows) == 2 and [r["vigente"] for r in rows] == [False, True]


def test_rows_without_model_formula_block_and_do_not_promote(env):
    env["state"]["missing"] = ["AP3", "AQ3"]
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 2
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "bloqueada" and imp["vigente"] is False
    assert imp["conferencia"]["status"] == "bloqueada"
    assert (env["out"] / "Modelo DRE - Untitled - 09_2026_BLOQUEADO.xlsx").exists()
    assert not (env["out"] / "Modelo DRE - Untitled - 09_2026.xlsx").exists()


def test_contract_error_in_the_build_blocks_and_is_recorded(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("DESPESAS!AP: classificação ausente")
    monkeypatch.setattr(cli, "build_workbook", boom)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 2
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "bloqueada" and imp["vigente"] is False
    assert "classificação ausente" in imp["motivo_falha"]
    assert imp["conferencia"]["bloqueios"] == ["erro de contrato: DESPESAS!AP: classificação ausente"]
    assert client.tables.get("cp_despesas", []) == []          # nada carregado
    assert "bloqueada" in (env["out"] / "conferencia.md").read_text(encoding="utf-8")


def test_sanitizer_contract_error_blocks_and_is_recorded(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("cr: coluna 3 é 'X'; esperado 'Y'")
    monkeypatch.setattr(cli, "sanitize_cr", boom)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 2
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "bloqueada" and "coluna 3" in imp["motivo_falha"]
    assert client.tables.get("arquivos_brutos", []) == [] and client.tables.get("cr_receita", []) == []


def test_contract_error_without_database_still_writes_the_conference(env, monkeypatch):
    def boom(*a, **kw):
        raise ContractError("rra: registro com nº de colunas inesperado")
    monkeypatch.setattr(cli, "sanitize_rra", boom)
    assert cli.main(_argv(env, "--sem-banco")) == 2
    conf = (env["out"] / "conferencia.json").read_text(encoding="utf-8")
    assert "bloqueada" in conf and "colunas inesperado" in conf


def test_unexpected_error_marks_failure(env, monkeypatch):
    def boom(*a, **kw):
        raise OSError("disco cheio")
    monkeypatch.setattr(cli, "build_workbook", boom)
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 3
    imp = client.tables["importacoes"][0]
    assert imp["status"] == "falhou" and imp["vigente"] is False
    assert "disco cheio" in imp["motivo_falha"]


def test_paused_database_is_detected_and_nothing_is_written(env, capsys):
    client = FakeClient(fail_with=ConnectionError("project paused"))
    assert cli.main(_argv(env), client=client) == 3
    err = capsys.readouterr().err
    assert "banco pausado: restaure o projeto e rode de novo" in err
    assert client.tables == {}
    assert not (env["out"] / "conferencia.json").exists()
    assert not (env["out"] / "Modelo DRE - Untitled - 09_2026.xlsx").exists()


def test_partial_period_end_and_screen_counts(env):
    client = FakeClient()
    cli.main(_argv(env, "--periodo-fim", "2026-09-15", "--tela", "cp=9,cr=8"), client=client)
    assert client.tables["importacoes"][0]["periodo_fim"] == "2026-09-15"
    by_rel = {r["relatorio"]: r for r in client.tables["arquivos_brutos"]}
    assert by_rel["cp"]["total_tela"] == 9 and by_rel["rra"]["total_tela"] is None
    conf = client.tables["importacoes"][0]["conferencia"]
    assert any("tela mostrava 9" in a for a in conf["alertas"])


@pytest.mark.parametrize("bad", ["2026-13", "2026/09", "setembro", "26-09"])
def test_invalid_competencia_is_rejected(env, bad):
    argv = _argv(env, "--sem-banco")
    argv[argv.index("--competencia") + 1] = bad
    assert cli.main(argv) == 3


def test_unknown_produtora_is_rejected(env):
    argv = _argv(env, "--sem-banco")
    argv[argv.index("--produtora") + 1] = "outra"
    assert cli.main(argv) == 3


def test_parse_tela():
    assert cli.parse_tela("cp=37030,cr=483") == {"cp": 37030, "cr": 483}
    with pytest.raises(cli.UsageError):
        cli.parse_tela("xx=1")
    with pytest.raises(cli.UsageError):
        cli.parse_tela("cp=abc")


def test_usage_errors_touch_nothing(env):
    client = FakeClient()
    argv = _argv(env)
    argv[argv.index("--competencia") + 1] = "2026-13"
    assert cli.main(argv, client=client) == 3
    assert client.tables == {} and not (env["out"] / "conferencia.json").exists()
```

- [ ] **Step 3: Rodar e ver falhar**

Run: `py -m pytest tests/test_cli.py -v`
Expected: FAIL com `ImportError: cannot import name 'cli'`

- [ ] **Step 4: Implementar `cli.py`**

```python
from __future__ import annotations

import argparse
import calendar
import datetime as dt
import json
import re
import sys
import traceback
from pathlib import Path

from .build import build_workbook
from .config import DEST, PRODUTORAS
from .db import BancoIndisponivel, Repository, make_client, sha256_file
from .report import contract_blocked, conference, render_markdown
from .sanitize import sanitize_comissoes, sanitize_cp, sanitize_cr, sanitize_rra
from .types import ContractError

REPORTS = ("rra", "cp", "cr", "comissoes")
EXIT_OK, EXIT_BLOCKED, EXIT_FAILED = 0, 2, 3


class UsageError(ValueError):
    """Argumento inválido do comando. Nada é gravado (nem banco nem arquivos)."""


def parse_tela(text: str | None) -> dict[str, int]:
    out: dict[str, int] = {}
    if not text:
        return out
    for part in text.split(","):
        key, _, val = part.partition("=")
        key = key.strip()
        if key not in REPORTS:
            raise UsageError(f"--tela: relatório desconhecido {key!r} (use {', '.join(REPORTS)})")
        try:
            out[key] = int(val)
        except ValueError:
            raise UsageError(f"--tela: valor inválido para {key}: {val!r}") from None
    return out


def parse_competencia(text: str) -> tuple[str, dt.date, dt.date]:
    m = re.fullmatch(r"(\d{4})-(0[1-9]|1[0-2])", text or "")
    if not m:
        raise UsageError(f"--competencia deve ser AAAA-MM; recebido {text!r}")
    year, month = int(m.group(1)), int(m.group(2))
    return text, dt.date(year, 1, 1), dt.date(year, month, calendar.monthrange(year, month)[1])


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="papaki_dre")
    sub = p.add_subparsers(dest="cmd", required=True)
    i = sub.add_parser("importar", help="sanitiza os 4 brutos, grava no Supabase e gera o Modelo DRE")
    i.add_argument("--produtora", required=True)
    i.add_argument("--competencia", required=True, help="AAAA-MM")
    for r in REPORTS:
        i.add_argument(f"--{r}", required=True, type=Path)
    i.add_argument("--modelo", required=True, type=Path, help="modelo anual vazio (.xlsx)")
    i.add_argument("--saida", required=True, type=Path)
    i.add_argument("--periodo-fim", help="AAAA-MM-DD; padrão = último dia da competência")
    i.add_argument("--gatilho", choices=["manual", "agendado"], default="manual")
    i.add_argument("--origem", default="arquivos locais")
    i.add_argument("--data-extracao", help="ISO-8601; padrão = agora (UTC)")
    i.add_argument("--tela", help="totais vistos na tela, ex.: cp=37030,cr=483,rra=1410")
    i.add_argument("--sem-banco", action="store_true")
    i.add_argument("--forcar", action="store_true", help="importa mesmo que os arquivos já tenham sido importados")
    return p


def _out_name(nome: str, competencia: str) -> str:
    year, month = competencia.split("-")
    return f"Modelo DRE - {nome} - {month}_{year}.xlsx"


def _write_conference(out_dir: Path, conf: dict) -> None:
    (out_dir / "conferencia.json").write_text(json.dumps(conf, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "conferencia.md").write_text(render_markdown(conf), encoding="utf-8")


def importar(a, client=None) -> int:
    # 1. argumentos (UsageError -> código 3, nada gravado)
    if a.produtora not in PRODUTORAS:
        raise UsageError(f"produtora desconhecida: {a.produtora!r} (cadastradas: {', '.join(PRODUTORAS)})")
    competencia, inicio, fim = parse_competencia(a.competencia)
    if a.periodo_fim:
        try:
            fim = dt.date.fromisoformat(a.periodo_fim)
        except ValueError:
            raise UsageError(f"--periodo-fim deve ser AAAA-MM-DD; recebido {a.periodo_fim!r}") from None
    tela = parse_tela(a.tela)
    nome = PRODUTORAS[a.produtora]["nome"]
    files = {k: Path(getattr(a, k)) for k in REPORTS}
    faltando = [str(p) for p in [*files.values(), Path(a.modelo)] if not p.is_file()]
    if faltando:
        raise UsageError(f"arquivo(s) não encontrado(s): {', '.join(faltando)}")
    out_dir = Path(a.saida)
    out_dir.mkdir(parents=True, exist_ok=True)
    extracao = a.data_extracao or dt.datetime.now(dt.timezone.utc).isoformat()

    # 2. dedupe por hash e criação da importação (antes de sanitizar: toda rodada fica registrada)
    hashes = {k: sha256_file(p) for k, p in files.items()}
    repo, imp_id = None, None
    if not a.sem_banco:
        repo = Repository(client if client is not None else make_client())
        try:
            repo.verificar_acesso()            # plano gratuito: pausa após ~7 dias sem uso
        except BancoIndisponivel as e:
            print(f"ERRO: banco inacessível ({e}).\n"
                  "banco pausado: restaure o projeto e rode de novo "
                  "(painel do Supabase ou MCP restore_project). Nada foi gravado.", file=sys.stderr)
            return EXIT_FAILED
        same = None if a.forcar else repo.find_identical(a.produtora, competencia, hashes)
        if same:
            print(f"Arquivos idênticos já importados (importação {same}); nada a fazer. Use --forcar para repetir.")
            return EXIT_OK
        imp_id = repo.criar_importacao(
            produtora=a.produtora, competencia=competencia, periodo_inicio=inicio.isoformat(),
            periodo_fim=fim.isoformat(), data_extracao=extracao, origem=a.origem, gatilho=a.gatilho)
    meta = {"produtora": a.produtora, "competencia": competencia,
            "periodo": f"{inicio.isoformat()} a {fim.isoformat()}", "data_extracao": extracao,
            **({"importacao_id": imp_id} if imp_id else {})}

    try:
        # 3. sanitizar, carregar o Excel e conferir
        try:
            results = {
                "cp": sanitize_cp(files["cp"]),
                "cr": sanitize_cr(files["cr"]),
                "comissoes": sanitize_comissoes(files["comissoes"]),
                # o XLSX convertido fica em work/<stem>-<hash12>.xlsx (xlsx_cache_path) e é reaproveitado
                "rra": sanitize_rra(files["rra"], work_dir=out_dir / "work"),
            }
            out_xlsx = out_dir / _out_name(nome, competencia)
            build = build_workbook(Path(a.modelo), out_xlsx, results)
        except ContractError as e:
            conf = contract_blocked(str(e), meta=meta)
            _write_conference(out_dir, conf)
            if repo:
                repo.marcar(imp_id, "bloqueada", motivo=f"ContractError: {e}", conferencia=conf)
            print(f"Status: bloqueada — erro de contrato: {e}", file=sys.stderr)
            return EXIT_BLOCKED

        conf = conference(results, build, screen_counts=tela, meta=meta)
        blocked = conf["status"] == "bloqueada"
        if blocked:
            final = out_xlsx.with_name(out_xlsx.stem + "_BLOQUEADO.xlsx")
            out_xlsx.replace(final)
            out_xlsx = final
        _write_conference(out_dir, conf)

        # 4. gravar no banco e promover (ou marcar bloqueada)
        if repo:
            for k in REPORTS:
                repo.registrar_arquivo(imp_id, k, nome=files[k].name, sha256=hashes[k],
                                       tamanho=files[k].stat().st_size, total_tela=tela.get(k))
                repo.inserir_linhas(imp_id, a.produtora, competencia, k, results[k])
                repo.inserir_exclusoes(imp_id, k, results[k])
            out_hash = sha256_file(out_xlsx)
            for k in REPORTS:
                sr = build.sheets[DEST[k]["sheet"]]
                repo.registrar_carga(imp_id, aba=sr.sheet, primeira=sr.first_row, ultima=sr.last_row,
                                     arquivo=out_xlsx.name, sha256=out_hash)
            if blocked:
                repo.marcar(imp_id, "bloqueada", conferencia=conf)
            else:
                repo.salvar_conferencia(imp_id, conf)
                repo.promover(imp_id)
        print(f"Status: {conf['status']} — arquivo: {out_xlsx}")
        return EXIT_BLOCKED if blocked else EXIT_OK
    except Exception as e:
        if repo and imp_id:
            repo.marcar(imp_id, "falhou", motivo=f"{type(e).__name__}: {e}")
        raise


def main(argv=None, client=None) -> int:
    args = _parser().parse_args(argv)
    try:
        return importar(args, client=client)
    except UsageError as e:
        print(f"ERRO DE USO: {e}", file=sys.stderr)
        return EXIT_FAILED
    except Exception as e:  # noqa: BLE001 - qualquer falha vira código 3 com rastro (já marcada 'falhou')
        traceback.print_exc()
        print(f"FALHA: {type(e).__name__}: {e}", file=sys.stderr)
        return EXIT_FAILED
```

`src/papaki_dre/__main__.py`:

```python
from .cli import main

raise SystemExit(main())
```

- [ ] **Step 5: Rodar e ver passar**

Run: `py -m pytest tests/test_cli.py -v`
Expected: 18 passed (o teste parametrizado de competência conta como 4).

- [ ] **Step 6: Commit**

```bash
git add src/papaki_dre/cli.py src/papaki_dre/__main__.py src/papaki_dre/config.py tests/test_cli.py && git commit -m "feat: comando importar (sanitização, banco, carga no Excel, conferência, fotos vigentes)"
```

---

### Task 5: Ponta a ponta com os brutos reais

**Files:**
- Test: `tests/test_cli_real.py`

- [ ] **Step 1: Escrever o teste**

```python
import json

import openpyxl
import pytest

from papaki_dre import cli
from papaki_dre.config import R_BLANK, R_FOOTER, R_H, R_JOB, R_LOC, R_MODELO, R_NUCLEO, R_PESQ
from conftest import REAL_FILES, TEMPLATE
from fake_supabase import FakeClient

pytestmark = [pytest.mark.real, pytest.mark.excel]

# Contagens finais do Plano 1 (extração de 08/10/2026): bruto = mantidas + técnicas + exclusões.
EXPECTED = {
    "rra": {"bruto": 2408, "tecnicas": 998, "mantidas": 72, "excluidas": {R_NUCLEO: 1329, R_PESQ: 9}},
    "cp": {"bruto": 36999, "tecnicas": 2, "mantidas": 1110},          # exclusões: 21.505 + 14.382 (R_JOB/R_LOC)
    "cr": {"bruto": 486, "tecnicas": 2, "mantidas": 484, "excluidas": {}},
    "comissoes": {"bruto": 1222, "tecnicas": 2, "mantidas": 649,
                  "excluidas": {R_MODELO: 81, R_PESQ: 126, R_H: 364}},
}


def _argv(out, *extra):
    return ["importar", "--produtora", "untitled", "--competencia", "2026-09",
            "--cp", str(REAL_FILES["cp"]), "--cr", str(REAL_FILES["cr"]),
            "--rra", str(REAL_FILES["rra"]), "--comissoes", str(REAL_FILES["comissoes"]),
            "--modelo", str(TEMPLATE), "--saida", str(out),
            "--tela", "cp=37030,cr=483,rra=1410", *extra]


@pytest.fixture(scope="module")
def run(tmp_path_factory):
    if not all(p.exists() for p in [TEMPLATE, *REAL_FILES.values()]):
        pytest.skip("modelo ou brutos de referência ausentes")
    out = tmp_path_factory.mktemp("e2e") / "saida"
    client = FakeClient()
    code = cli.main(_argv(out), client=client)
    return code, out, client


def _conf(out):
    return json.loads((out / "conferencia.json").read_text(encoding="utf-8"))


def test_run_is_published_with_alerts_only(run):
    code, out, _ = run
    conf = _conf(out)
    assert code == 0, (code, conf["bloqueios"])
    assert conf["status"] == "com_alertas" and conf["bloqueios"] == []


def test_september_count_flow_closes(run):
    _, out, _ = run
    rel = _conf(out)["relatorios"]
    for key, exp in EXPECTED.items():
        r = rel[key]
        assert r["registros_bruto"] == exp["bruto"], key
        assert r["removidas_tecnicas"] == exp["tecnicas"], key
        assert r["mantidas"] == r["carregadas"] == exp["mantidas"], key
        assert r["registros_bruto"] == r["mantidas"] + r["removidas_tecnicas"] + sum(r["excluidas"].values()), key
        if "excluidas" in exp:
            assert r["excluidas"] == exp["excluidas"], key
    assert set(rel["cp"]["excluidas"]) == {R_JOB, R_LOC}
    assert sorted(rel["cp"]["excluidas"].values()) == [14382, 21505]
    assert sorted(rel["cp"]["tecnicas_por_motivo"]) == sorted([R_BLANK, R_FOOTER])


def test_screen_count_differences_are_alerts_only(run):
    _, out, _ = run
    conf = _conf(out)
    alertas = "\n".join(conf["alertas"])
    assert "tela mostrava 37030" in alertas and "tem 36997" in alertas      # CP: risco 3 da spec
    assert "tela mostrava 483" in alertas and "tem 484" in alertas          # CR
    assert not any("RRA" in a and "tela" in a for a in conf["alertas"])     # RRA: 2408 - 998 = 1410
    assert not any("tela" in b for b in conf["bloqueios"])


def test_database_received_one_current_snapshot(run):
    _, _, client = run
    imps = client.tables["importacoes"]
    assert len(imps) == 1
    assert imps[0]["status"] == "concluida" and imps[0]["vigente"] is True
    assert len(client.tables["rra_base"]) == 72
    assert len(client.tables["cp_despesas"]) == 1110
    assert len(client.tables["cr_receita"]) == 484
    assert len(client.tables["comissoes"]) == 649
    by_type = {}
    for e in client.tables["exclusoes"]:
        by_type[e["tipo"]] = by_type.get(e["tipo"], 0) + 1
    tecnicas = sum(EXPECTED[k]["tecnicas"] for k in EXPECTED)                 # 998 + 2 + 2 + 2
    negocio = sum(EXPECTED[k]["bruto"] - EXPECTED[k]["tecnicas"] - EXPECTED[k]["mantidas"] for k in EXPECTED)
    assert by_type == {"tecnica": tecnicas, "negocio": negocio}


def test_output_workbook_has_the_data(run):
    _, out, _ = run
    candidates = list(out.glob("Modelo DRE - Untitled - 09_2026*.xlsx"))
    assert len(candidates) == 1
    ws = openpyxl.load_workbook(candidates[0])["DESPESAS"]
    assert ws["A3"].value is not None and ws["A1112"].value is not None and ws["A1113"].value is None


def test_second_run_with_same_files_is_ignored(run, tmp_path):
    _, _, client = run
    assert client.tables["importacoes"][0]["status"] == "concluida"
    assert cli.main(_argv(tmp_path / "again"), client=client) == 0
    assert len(client.tables["importacoes"]) == 1
```

- [ ] **Step 2: Rodar**

Run: `py -m pytest tests/test_cli_real.py -v`
Expected: 6 passed. Leva alguns minutos (conversão do XLS no Excel — reaproveitada em `work/<stem>-<hash12>.xlsx` se já existir —, carga e 37 mil registros de CP).

**Se `test_run_is_published_with_alerts_only` falhar com bloqueio:** leia `conf["bloqueios"]`. Contagens que não fecham ou linhas sem fórmula do modelo são defeito a investigar (Planos 1 e 2), não motivo para afrouxar a regra. Os alertas esperados são as diferenças tela × arquivo do CP e do CR, as fórmulas órfãs de `DESPESAS`/`RECEITA` (Plano 2, Task 5) e eventuais avisos dos sanitizadores.

**Se `test_september_count_flow_closes` divergir**: os números são os finais do Plano 1 para a extração de 08/10/2026 (RRA 2.408 → 998 técnicas → 1.329 núcleo + 9 PESQUISA → 72; CP 36.999 → 2 → 21.505 + 14.382 → 1.110; CR 486 → 2 → 484; Comissões 1.222 → 2 → 81 modelo + 126 PESQUISA + 364 coluna H → 649). Investigue antes de mudar qualquer número.

- [ ] **Step 3: Commit**

```bash
git add tests/test_cli_real.py && git commit -m "test: ponta a ponta com os brutos reais de setembro e cliente falso do banco"
```

---

### Task 6: Execução real contra o Supabase e documentação de uso

**Files:**
- Create: `.env.example`, `README.md`

- [ ] **Step 1: Criar `.env.example`**

```dotenv
# Copie para .env (que NÃO vai para o git) e preencha com o projeto Supabase dedicado.
SUPABASE_URL=https://<project-ref>.supabase.co
SUPABASE_SERVICE_KEY=<service_role key — nunca no código nem no git>
```

- [ ] **Step 2: Criar `README.md`**

```markdown
# papaki-dre

Importa as bases do Esferas (RRA, CP, CR e Real x Orçado Comissões) de uma produtora, sanitiza,
grava no Supabase e gera uma **cópia** do Modelo DRE preenchida só nas abas de base.
MVP: produtora **Untitled**.

## Requisitos
Windows com Microsoft Excel instalado (usado para converter o XLS do RRA), Python ≥ 3.11 e um
projeto Supabase dedicado (plano gratuito) com a migração `supabase/migrations/0001_schema.sql` aplicada.

**Supabase pausado:** no plano gratuito o projeto pausa após ~7 dias sem uso. O comando detecta o
banco inacessível antes de gravar qualquer coisa, avisa "banco pausado: restaure o projeto e rode de
novo" e termina com código 3. Restaure pelo painel (ou pelo MCP `restore_project`), aguarde ficar
ativo e rode de novo. A restauração automática virá com a rotina mensal (Plano 4).

## Instalação
    python -m pip install -e ".[dev]"
    copy .env.example .env      # e preencha SUPABASE_URL e SUPABASE_SERVICE_KEY

## Uso
    python -m papaki_dre importar --produtora untitled --competencia 2026-09 ^
      --cp "CP.csv" --cr "CR.csv" --rra "RRA.xls" --comissoes "Comissoes.xlsx" ^
      --modelo "Modelo DRE - Untitled - base.xlsx" --saida out\2026-09 ^
      --tela cp=37030,cr=483,rra=1410

Opções: `--periodo-fim AAAA-MM-DD` (demonstrativo parcial), `--sem-banco`, `--forcar`,
`--gatilho agendado`, `--origem`, `--data-extracao`.

## Saídas
`Modelo DRE - <Produtora> - MM_AAAA.xlsx` (ou `..._BLOQUEADO.xlsx`), `conferencia.md`, `conferencia.json`.

## Códigos de saída
`0` ok/com alertas (ou arquivos idênticos já importados) · `2` bloqueada (não publicada: contagens
que não fecham, linha sem fórmula do modelo/classificação ausente, lacuna no TOTAL da BASE, ou erro
de contrato) · `3` falha, banco pausado/inacessível ou argumento inválido.

## Conferência
Feita por contagens: bruto → remoções técnicas (linhas vazias e rodapés de TOTAL) → exclusões de
negócio por motivo → mantidas → carregadas. Os valores das linhas de TOTAL dos relatórios não são
lidos nem comparados. Diferença entre o total visto na tela (`--tela`) e o arquivo é só alerta.

## Regras
Ver `instrucoes_relatorios_ base.md` e `docs/superpowers/specs/`. O sistema só escreve nas abas
`BASE`, `DESPESAS`, `RECEITA` e `REAL X ORÇADO COMISSÕES`; as demais abas não são tocadas.
As tabelas dinâmicas do modelo não são atualizadas: isso é do FP&A.
```

- [ ] **Step 3: Rodar uma importação real no Supabase (com autorização do usuário)**

Pré-condições: Task 2 concluída, o `.env` preenchido pelo usuário (**nunca** peça nem imprima a chave de serviço) e o projeto Supabase **ativo** (MCP `get_project`; se pausado, `restore_project` e aguardar — restaurar o projeto do usuário é ação dele ou com a autorização dele). Se o comando responder "banco pausado: restaure o projeto e rode de novo" (código 3), restaure e repita; nada terá sido gravado. Execute com os brutos de setembro:

```bash
python -m papaki_dre importar --produtora untitled --competencia 2026-09 \
  --cp "amostras/2026-09/bruto/CP_papaki_consolidado_2026-01-01_a_2026-09-30.csv" \
  --cr "amostras/2026-09/bruto/CR_untitled_UTLD_2026-01-01_a_2026-09-30.csv" \
  --rra "amostras/2026-09/bruto/Relatório Rentabilidade.xls" \
  --comissoes "amostras/2026-09/bruto/COMISSOES_untitled_UTLD_2026-01-01_a_2026-09-30.xlsx" \
  --modelo "Modelo DRE - Untitled - base.xlsx" --saida out/2026-09 --tela cp=37030,cr=483,rra=1410
```

Expected: `Status: com_alertas` (diferenças tela × arquivo no CP e no CR), código 0. Confira no Supabase (MCP `execute_sql`): `select status, vigente from importacoes;` → `concluida | true`; `select count(*) from cp_despesas;` → 1110; `select relatorio, tipo, count(*) from exclusoes group by 1, 2;` → técnicas rra 998, cp 2, cr 2, comissoes 2 e negócio rra 1338, cp 35887, comissoes 571; repita o comando e confirme a mensagem "Arquivos idênticos já importados". Para testar a foto vigente (aceite (e) da spec §9), importe a 2ª extração de `amostras/2026-09/bruto_papaki/` (nomes no padrão Papaki; contagens próprias, ainda não calculadas) e confira que ela vira `vigente` e a primeira continua guardada com `vigente = false`.

Os dados reais contêm CPF/CNPJ e dados bancários: confirme com o usuário antes de gravar no banco e não cole linhas desses dados em respostas ou logs.

- [ ] **Step 4: Commit**

```bash
git add .env.example README.md && git commit -m "docs: README de uso e exemplo de configuração do Supabase"
```

---

### Task 7: Alinhar a spec com o que foi construído

**Files:**
- Modify: `docs/superpowers/specs/2026-10-09-papaki-esferas-mvp-design.md`

- [ ] **Step 1: Ajustar a linha de dados tratados da §4**

Substitua, na seção 4, a linha que começa com `` - `rra_base`, `cp_despesas`, `cr_receita`, `comissoes` `` por:

```markdown
- `rra_base`, `cp_despesas`, `cr_receita`, `comissoes` — dado tratado: todas as colunas do relatório (49, 41, 28 e 14) em um `jsonb` com valores tipados (datas ISO-8601, números como número, vazio = `null`), mais chave promovida (`cod` no CP; `num_job` no RRA e nas Comissões) e rastreio (`linha_origem`, `linha_destino`, `importacao_id`). Colunas tipadas fixas podem ser geradas depois como views.
```

Na mesma seção, substitua a linha que começa com `` - `exclusoes` `` por:

```markdown
- `exclusoes` — relatório, linha de origem, `tipo` (`tecnica` = linha vazia ou rodapé de TOTAL; `negocio` = regra da §6), `motivo` (código) e `regra` (texto). As duas listas são gravadas, para refazer bruto → técnicas → exclusões → mantidas só com o banco.
```

E, logo abaixo do título da §4, troque a linha que começa com `Todas as tabelas:` (hoje: ``Todas as tabelas: `produtora_id`, `competencia` (ex.: `2026-09`), `data_extracao`, `origem`.``) por:

```markdown
Todas as tabelas de dados levam `importacao_id`, `produtora_id` e `competencia` (ex.: `2026-09`); `data_extracao`, `origem`, período e gatilho ficam em `importacoes` (uma vez por foto).
```

- [ ] **Step 2: Registrar o que mudou desde a spec**

Na §8, troque a linha `- Contagens: bruto → após remoções técnicas → após exclusões → carregadas; lista de exclusões com motivo.` por:

```markdown
- Contagens: bruto → após remoções técnicas → após exclusões (por motivo) → mantidas → carregadas na aba; bruto = mantidas + técnicas + exclusões e mantidas = carregadas, senão **bloqueia**. Erro de contrato (cabeçalho/colunas divergentes, registro com nº de colunas inesperado, erro inesperado de célula, classificação ausente) também deixa a importação **bloqueada**, registrada com o motivo; outras exceções a deixam **falhou**.
```

Acrescente ao final da lista numerada da §10 (riscos e pontos em aberto), sem repetir o que já está em "Decisões de 09/10/2026" (cabeçalhos escritos pelo sistema; `O:Q` do modelo de setembro; cobertura de fórmulas em todas as linhas carregadas; Supabase pausado detectado com código 3 e restaurado pelo Plano 4):

```markdown
7. **Tabelas dinâmicas** do modelo mantêm o cache antigo até o FP&A atualizar.
```

- [ ] **Step 3: Commit**

```bash
git add docs/superpowers/specs/2026-10-09-papaki-esferas-mvp-design.md && git commit -m "docs: alinha a spec com o desenho de dados, a conferência por contagens e os achados do modelo"
```

---

## Self-Review

**Cobertura da spec:** §4 tabelas, foto vigente por produtora + competência e RLS sem políticas (T2, T3), chave de serviço só no `.env` (T3 `make_client`, T6), dedupe por hash — mesmo hash = ignorado (T3, T4), histórico preservado (T2 função, T3 teste), `exclusoes` com técnicas e negócio, motivo e regra (T2, T3); §5 falha marcada `falhou` com motivo (T4); §6 contrato estrutural (`ContractError`) bloqueia a rodada (T4); §8 conferência por contagens bruto → técnicas → exclusões por motivo → mantidas → carregadas, cobertura de fórmulas (`AP:AQ`, `AC`, `L:N`, `O:Q` e TOTAL da `BASE`, vinda do Plano 2) em todas as linhas carregadas, classificação ausente bloqueia, alertas tela × arquivo nunca erro (T1, T4) — **sem totais de controle** (eliminados em 09/10/2026; linhas de TOTAL só contadas como técnicas); status bloqueada não publica (T1, T4); §9 aceite: (a) importa sem intervenção (T5), (b) conferência fecha nas contagens (T1 invariante, T5 números reais), (c) abas fora das bases intactas (Plano 2), (d) reimportar não duplica (T4, T5), (e) segunda extração vira vigente sem apagar a primeira (T3 promover, T4 `--forcar`, T6 com `bruto_papaki`); competência, data de extração e origem registradas (T3, T4); funcionamento independente da data (`--periodo-fim`, `--data-extracao`, T4); Supabase no plano gratuito com pausa detectada (`verificar_acesso`, código 3; T2, T3, T4, T6).

**Decisões confirmadas em 09/10/2026 (antes "em aberto"):** rodada parada por `ContractError` registrada como `bloqueada` com o motivo, sem linhas nem `arquivos_brutos`, código 2 (T4); rodada bloqueada pela conferência grava linhas e exclusões sem torná-las vigentes (T4); só classificação ausente é detectada, e ela bloqueia (T1, T4); Supabase pausado: detectar, avisar e sair com código 3, restauração automática no Plano 4 (T3, T4).

**Fora deste plano:** Plano 4 (automação do Chrome: extração, e-mail do RRA, prova de Localidade em `arquivos_brutos.prova_caminho`/`filtros_tela`, restauração do projeto Supabase pausado, agendamento no 2º dia útil).

**Riscos assumidos:** criar projeto no Supabase exige confirmar nome e organização com o usuário (T2); pausa do plano gratuito: no MVP o `importar` só detecta e avisa (código 3), quem executa restaura e roda de novo até o Plano 4 automatizar (T2, T4, T6); `verificar_acesso` trata qualquer falha da consulta leve como banco indisponível (a mensagem traz o erro original, para distinguir pausa de chave/URL errada) (T3).

**Consistência de nomes:** `conference`/`contract_blocked`/`render_markdown` (T1); `BancoIndisponivel`, `Repository` (`verificar_acesso`, `criar_importacao`, `registrar_arquivo`, `inserir_linhas`, `inserir_exclusoes`, `registrar_carga`, `find_identical`, `promover`, `marcar`, `salvar_conferencia`), `sha256_file`, `json_value`, `row_to_json`, `make_client`, `REGRAS` (T3); `main`, `importar`, `parse_tela`, `parse_competencia`, `UsageError` (T4); `PRODUTORAS` (T4, `config.py`). Do Plano 1, só nomes que existem no código final: `Sanitized` (sem `controls`), `Removal`, `ContractError`, `Formula`, `XlError`, `DEST`, `R_*` (sem `R_ODD`), `sanitize_*`; do Plano 2: `build_workbook`, `BuildReport`, `SheetReport` (com `total_row` opcional; `missing_formulas` inclui `L:N` e o TOTAL da `BASE`).
