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
    if not isinstance(meta, dict):
        return False
    if conf.get("status") not in ("ok", "com_alertas") or not meta.get("importacao_id"):
        return False
    if meta.get("competencia") != competencia:
        return False
    _, _, fim = periodo_da_competencia(competencia)
    return str(meta.get("periodo", "")).endswith(fim.isoformat())
