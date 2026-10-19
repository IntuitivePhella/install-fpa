"""Cliente Supabase falso, em memória, com a interface mínima que o Repository usa.

Imita só as regras do banco (supabase/migrations/0001_schema.sql) que importam aos testes:
checks de `status`/`gatilho`/`relatorio`/`tipo`, a constraint `vigente_so_se_concluida`, o índice
único parcial de vigente por (produtora_id, competencia), o unique (importacao_id, relatorio) de
`arquivos_brutos`, a FK `importacao_id` com `on delete cascade`, a função `promover_importacao`
(só para 'em_andamento'), o limite padrão de linhas por consulta do PostgREST (1000), a projeção de
colunas do `select`, `returning=minimal` (data vazia) e a serialização JSON do cliente real (todo
payload passa por `json.dumps(allow_nan=False)`; o Postgres recusa texto com o caractere NUL).
Uma violação levanta `FakeAPIError` e não altera nada (comando atômico).
"""
import copy
import json
import uuid

MAX_ROWS = 1000  # limite padrão do PostgREST por consulta

CHECKS = {
    "importacoes": {"status": {"em_andamento", "concluida", "bloqueada", "falhou"},
                    "gatilho": {"manual", "agendado"}},
    "arquivos_brutos": {"relatorio": {"rra", "cp", "cr", "comissoes"}},
    "exclusoes": {"relatorio": {"rra", "cp", "cr", "comissoes"}, "tipo": {"tecnica", "negocio"}},
}
CHILDREN = ("arquivos_brutos", "rra_base", "cp_despesas", "cr_receita", "comissoes", "exclusoes",
            "cargas_excel")


class FakeAPIError(Exception):
    """Equivale ao postgrest.APIError: violação de constraint, check ou regra da função."""


def _wire(payload):
    """O que o cliente real manda pela rede: JSON estrito, sem NaN; texto com NUL é recusado."""
    try:
        data = json.loads(json.dumps(payload, allow_nan=False))
    except (TypeError, ValueError) as e:
        raise FakeAPIError(f"payload não serializável em JSON: {e}") from e

    def walk(v):
        if isinstance(v, str) and "\x00" in v:
            raise FakeAPIError("unsupported Unicode escape sequence (\\u0000 cannot be converted to text)")
        if isinstance(v, dict):
            for k, x in v.items():
                walk(k)
                walk(x)
        elif isinstance(v, list):
            for x in v:
                walk(x)
    walk(data)
    return data


class _Resp:
    def __init__(self, data):
        self.data = data


def _validate(name, rows, all_tables):
    for col, allowed in CHECKS.get(name, {}).items():
        for r in rows:
            if r.get(col) not in allowed:
                raise FakeAPIError(f"check violado em {name}.{col}: {r.get(col)!r}")
    if name == "importacoes":
        vigentes = set()
        for r in rows:
            if r["vigente"] and r["status"] != "concluida":
                raise FakeAPIError("violates check constraint vigente_so_se_concluida")
            if r["vigente"]:
                k = (r["produtora_id"], r["competencia"])
                if k in vigentes:
                    raise FakeAPIError("duplicate key value violates unique constraint importacoes_uma_vigente")
                vigentes.add(k)
    if name == "arquivos_brutos":
        keys = [(r["importacao_id"], r["relatorio"]) for r in rows]
        if len(keys) != len(set(keys)):
            raise FakeAPIError("duplicate key value violates unique constraint (importacao_id, relatorio)")
    if name in CHILDREN:
        ids = {r["id"] for r in all_tables.get("importacoes", [])}
        for r in rows:
            if r.get("importacao_id") not in ids:
                raise FakeAPIError(f"violates foreign key constraint em {name}.importacao_id")


class _Query:
    def __init__(self, client, name):
        self.client, self.name = client, name
        self._op = None
        self._filters = []
        self._cols = None                  # None = todas as colunas
        self._minimal = False
        self._limit = None

    def insert(self, rows, *, returning=None):
        rows = _wire(rows if isinstance(rows, list) else [rows])
        self._op = ("insert", rows)
        self._minimal = getattr(returning, "value", returning) == "minimal"
        self.client.returning.setdefault(self.name, []).append(returning)
        return self

    def select(self, *cols):
        names = [c.strip() for part in cols for c in part.split(",") if c.strip()]
        self._cols = None if not names or "*" in names else names
        self._op = ("select",)
        return self

    def update(self, values):
        self._op = ("update", _wire(values))
        return self

    def delete(self):
        self._op = ("delete",)
        return self

    def limit(self, n):
        self._limit = n
        return self

    def eq(self, col, val):
        self._filters.append(lambda r: r.get(col) == val)
        return self

    def in_(self, col, vals):
        vals = list(vals)
        self._filters.append(lambda r: r.get(col) in vals)
        return self

    def execute(self):
        tables = self.client.tables
        rows = tables.setdefault(self.name, [])
        kind = self._op[0]
        n = len(self._op[1]) if kind == "insert" else None
        self.client.calls.append((self.name, kind, n))
        if kind == "insert":
            new = []
            for r in self._op[1]:
                r = copy.deepcopy(r)
                if self.name in ("importacoes", "arquivos_brutos"):
                    r.setdefault("id", str(uuid.uuid4()))
                else:
                    self.client.seq += 1
                    r.setdefault("id", self.client.seq)
                if self.name == "importacoes":
                    r.setdefault("status", "em_andamento")
                    r.setdefault("vigente", False)
                new.append(r)
            _validate(self.name, rows + new, tables)
            rows.extend(new)
            return _Resp([] if self._minimal else copy.deepcopy(new))
        match = [r for r in rows if all(f(r) for f in self._filters)]
        if kind == "select":
            teto = self.client.max_rows if self._limit is None else min(self._limit, self.client.max_rows)
            out = match[:teto]
            if self._cols is not None:
                out = [{c: r.get(c) for c in self._cols} for r in out]
            return _Resp(copy.deepcopy(out))
        if kind == "update":
            ids = {id(r) for r in match}
            candidate = [{**r, **self._op[1]} if id(r) in ids else r for r in rows]
            _validate(self.name, candidate, tables)
            for r in match:
                r.update(self._op[1])
            return _Resp(copy.deepcopy(match))
        # delete (com cascade das tabelas filhas quando apaga importações)
        ids = {id(r) for r in match}
        tables[self.name] = [r for r in rows if id(r) not in ids]
        if self.name == "importacoes":
            gone = {r["id"] for r in match}
            for child in CHILDREN:
                if child in tables:
                    tables[child] = [r for r in tables[child] if r.get("importacao_id") not in gone]
        return _Resp(copy.deepcopy(match))


class _Rpc:
    def __init__(self, client, name, params):
        self.client, self.name, self.params = client, name, _wire(params)

    def execute(self):
        assert self.name == "promover_importacao"
        self.client.calls.append(("rpc", self.name, None))
        rows = self.client.tables.get("importacoes", [])
        target = next((r for r in rows if r["id"] == self.params["p_id"]), None)
        if target is None or target["status"] != "em_andamento":
            raise FakeAPIError(f"importação {self.params['p_id']} não está em andamento")
        key = (target["produtora_id"], target["competencia"])
        candidate = []
        for r in rows:
            r2 = dict(r)
            if (r2["produtora_id"], r2["competencia"]) == key and r2["vigente"]:
                r2["vigente"] = False
            if r2["id"] == target["id"]:
                r2["vigente"], r2["status"] = True, "concluida"
            candidate.append(r2)
        _validate("importacoes", candidate, self.client.tables)
        self.client.tables["importacoes"] = candidate
        return _Resp(None)


class FakeClient:
    def __init__(self, fail_with: Exception | None = None, max_rows: int = MAX_ROWS):
        self.tables = {"produtoras": [{"id": "untitled", "nome": "Untitled",
                                       "localidade": "UTLD PRODUÇÃO DE FILMES LTDA"}]}
        self.fail_with = fail_with          # simula projeto pausado/inacessível: toda chamada falha
        self.max_rows = max_rows
        self.calls: list[tuple] = []        # (tabela, operação, nº de linhas inseridas) — round-trips
        self.seq = 0
        self.returning: dict[str, list] = {}  # tabela -> `returning` de cada insert

    def table(self, name):
        if self.fail_with is not None:
            raise self.fail_with
        return _Query(self, name)

    def rpc(self, name, params):
        if self.fail_with is not None:
            raise self.fail_with
        return _Rpc(self, name, params)
