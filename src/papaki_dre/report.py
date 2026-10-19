from __future__ import annotations

from collections import Counter

from .build import BuildReport
from .config import DEST
from .types import ContractError, Sanitized

ORDER = ("rra", "cp", "cr", "comissoes")
LABEL = {
    "rra": "RRA (BASE)",
    "cp": "CP (DESPESAS)",
    "cr": "CR (RECEITA)",
    "comissoes": "Comissões (REAL X ORÇADO COMISSÕES)",
}


def conference(results: dict[str, Sanitized], build: BuildReport, *,
               screen_counts: dict[str, int] | None = None, meta: dict | None = None,
               na_count: int | None = None, na_error: str | None = None) -> dict:
    """Consolida a rodada. `na_count`: células com #N/D em DESPESAS!AP:AQ contadas pelo Excel
    (`recalc.contar_nd`); `na_error`: motivo de a contagem ter falhado. Os dois só geram ALERTAS,
    nunca bloqueio (a classificação é regra da Papaki). `informacoes` é só informativa: não muda o
    status."""
    absent = [k for k in ORDER if k not in results]
    if absent:
        raise ContractError(f"faltam resultados para: {', '.join(absent)}")
    screen_counts = screen_counts or {}
    alertas: list[str] = []
    bloqueios: list[str] = []
    informacoes: list[str] = []
    relatorios = {}

    for key in ORDER:
        s = results[key]
        label = LABEL[key]
        aba = DEST[key]["sheet"]
        sr = build.sheets.get(aba)
        if sr is None:
            bloqueios.append(f"{label}: aba {aba} ausente no build")
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
        if sr is not None and carregadas != s.kept:
            bloqueios.append(f"{label}: {s.kept} linhas mantidas, mas {carregadas} carregadas na aba {aba}")

    abas = {}
    for name, sr in build.sheets.items():
        abas[name] = {"linhas": sr.rows, "primeira": sr.first_row, "ultima": sr.last_row,
                      "sem_formula": len(sr.missing_formulas), "formulas_orfas": sr.orphan_formula_rows,
                      "datas_sem_formato": sr.date_style_warnings,
                      "constantes_substituidas": sr.overwritten_constants,
                      "formulas_unificadas": sr.unified_formula_rows,
                      "autofiltro_estendido": sr.autofilter_extended}
        if sr.missing_formulas:
            bloqueios.append(f"{name}: {len(sr.missing_formulas)} linhas carregadas sem a fórmula do "
                             f"modelo (ex.: {', '.join(sr.missing_formulas[:3])})")
        if sr.orphan_formula_rows:
            alertas.append(f"{name}: {sr.orphan_formula_rows} linhas com fórmulas do modelo abaixo "
                           f"dos dados carregados (mantidas)")
        if sr.date_style_warnings:
            alertas.append(f"{name}: datas gravadas em células sem formato de data: {sr.date_style_warnings}")
        # informativo (decisões 10-12 do HANDOFF): só contagens, nunca alerta nem bloqueio
        if sr.overwritten_constants:
            informacoes.append(f"{name}: {sr.overwritten_constants} constante(s) ou fórmula(s) de outra "
                               f"família substituída(s) pela fórmula do modelo")
        if sr.unified_formula_rows:
            informacoes.append(f"{name}: {sr.unified_formula_rows} linha(s) com a fórmula de dados "
                               f"unificada para a fórmula completa do modelo")
        if sr.autofilter_extended:
            informacoes.append(f"{name}: autofiltro estendido até a última linha carregada")

    if na_count:
        alertas.append(f"{na_count} células com #N/D em AP:AQ: classificação sem correspondência no "
                       f"De-Para; a classificação é regra da Papaki; só alerta")
    if na_error:
        alertas.append(f"não foi possível recalcular no Excel para contar #N/D: {na_error}")

    if na_count is None and na_error is None:
        informacoes.append("contagem de #N/D não executada")

    status = "bloqueada" if bloqueios else ("com_alertas" if alertas else "ok")
    return {"status": status, "meta": meta or {}, "relatorios": relatorios, "abas": abas,
            "partes_alteradas": build.parts, "alertas": alertas, "bloqueios": bloqueios,
            "informacoes": informacoes}


def contract_blocked(error: str, *, meta: dict | None = None) -> dict:
    """Conferência de uma rodada interrompida por ContractError (cabeçalho/colunas divergentes,
    registro com nº de colunas inesperado, erro inesperado de célula, classificação ausente,
    modelo sujo). Nada é publicado."""
    return {"status": "bloqueada", "meta": meta or {}, "relatorios": {}, "abas": {},
            "partes_alteradas": {}, "alertas": [], "bloqueios": [f"erro de contrato: {error}"],
            "informacoes": []}


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
    if conf.get("informacoes"):
        out += ["## Informações", *[f"- {i}" for i in conf["informacoes"]], ""]
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
                "| Aba | Linhas | Faixa | Sem fórmula | Fórmulas órfãs | Constantes substituídas "
                "| Fórmulas unificadas | Autofiltro estendido |",
                "|---|---:|---|---:|---:|---:|---:|---|"]
        for name, a in conf["abas"].items():
            out.append(f"| {name} | {a['linhas']} | {a['primeira']}–{a['ultima']} | "
                       f"{a['sem_formula']} | {a['formulas_orfas']} | {a['constantes_substituidas']} | "
                       f"{a['formulas_unificadas']} | {'sim' if a['autofiltro_estendido'] else 'não'} |")
    if conf["partes_alteradas"]:
        out += ["", "## Partes do arquivo alteradas", ""]
        out += [f"- `{p}`" for p in conf["partes_alteradas"].get("changed", [])]
        for p in conf["partes_alteradas"].get("removed", []):
            out.append(f"- removida: `{p}`")
    out.append("")
    return "\n".join(out)


def render_resultado(conf: dict) -> str:
    """Resumo curto (RESULTADO.md), em português simples, para quem não vai ler a conferência inteira."""
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
