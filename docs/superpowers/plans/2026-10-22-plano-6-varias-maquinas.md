# Plano 6 — Várias produtoras e várias máquinas: Plano de Implementação

> **Para quem for executar (agente):** SUB-SKILL OBRIGATÓRIA: use superpowers:subagent-driven-development (recomendado) ou superpowers:executing-plans para executar este plano tarefa por tarefa. Os passos usam caixas `- [ ]` para acompanhamento.
>
> **Regras do repositório que valem aqui (CLAUDE.md):** fale em português simples com o usuário; nunca abra nem imprima o `.env`; nunca digite senha; dados de teste sempre inventados; o repositório é **público**, então rode `.venv/Scripts/python.exe scripts/auditar_publicacao.py` antes de qualquer commit; **só faça `git commit` quando o usuário pedir** (os passos de commit abaixo são o ponto de parada sugerido, não uma ordem); gravar no Supabase real, criar projeto/branch de teste ou criar tarefa agendada só com o "sim" do usuário no chat.

**Objetivo:** deixar o sistema pronto para as 18 produtoras da Papaki rodarem em mais de uma máquina (uma principal com todas, mais máquinas de grupo), sem duplicar código, banco nem rotina, e sem duas máquinas extraírem a mesma produtora e competência ao mesmo tempo.

**Arquitetura:** o código e o esquema do banco continuam únicos. A defesa contra extração em duplicidade passa a ter duas camadas: horários escalonados (só texto) e uma **reserva atômica** no Supabase (tabela `reservas_rodada` + funções SQL `reservar_rodada` e `liberar_rodada`), pedida pelo `agenda --reservar` e liberada pelo `importar`. Cada máquina se identifica por `PAPAKI_MAQUINA` no `.env`, e a importação grava essa máquina. Um comando de leitura, `produtora situacao`, mostra as 18 produtoras de uma vez.

**Tecnologias:** Python 3.11+ (argparse, pytest), Supabase/PostgREST (RPC com funções `plpgsql`), Markdown.

**Desenho aprovado:** `docs/superpowers/specs/2026-10-22-plano-6-varias-maquinas-design.md` (decisões do usuário em 22/10/2026: pasta de dados compartilhada; rodada agendada segue e avisa com o banco indisponível; reserva de 6 h; intervalo entre tarefas a medir no ensaio; validação da função SQL num banco de teste autorizada).

---

## Mapa de arquivos

| Arquivo | O que muda |
|---|---|
| `INSTALAR.md`, `runbooks/tarefa-agendada.md`, `NOVA-PRODUTORA.md` | Regra de várias máquinas, horários escalonados, pasta compartilhada (Tarefa 1); depois reserva e `PAPAKI_MAQUINA` (Tarefa 10) |
| `supabase/migrations/0002_reservas.sql` (novo) | Coluna `importacoes.maquina`, tabela `reservas_rodada`, funções `reservar_rodada` e `liberar_rodada` |
| `tests/fake_supabase.py` | O cliente falso passa a entender as duas funções novas e a ter um relógio trocável |
| `src/papaki_dre/db.py` | `Repository.reservar_rodada`, `liberar_rodada`, `situacao`; `criar_importacao(maquina=)` |
| `src/papaki_dre/ambiente.py` | `PAPAKI_MAQUINA` entra nas chaves graváveis |
| `src/papaki_dre/cli_instalacao.py` | `configurar --maquina`; `produtora situacao` |
| `src/papaki_dre/cli_extracao.py` | `agenda --reservar`; comando `reserva obter` |
| `src/papaki_dre/cli.py` | `_reservar_rodada`, `_liberar_reserva`; `importar` grava a máquina e libera a reserva; ligação no `main` |
| `src/papaki_dre/verificacao.py` | Avisos de `PAPAKI_MAQUINA` e da tabela de reservas |
| `.env.example`, `README.md`, `runbooks/rotina-mensal.md` | Documentação das chaves e dos comandos novos |
| `tests/test_reservas_db.py`, `tests/test_reserva_cli.py`, `tests/test_produtora_situacao.py`, `tests/test_migracao_0002.py` (novos); `tests/test_cli.py`, `tests/test_cli_instalacao.py`, `tests/test_verificacao.py`, `tests/test_docs_instalacao.py` (ampliados) | Testes de cada parte |

Comando de teste usado em todo o plano: `.venv/Scripts/python.exe -m pytest <arquivo> -q`. A suíte rápida inteira é `.venv/Scripts/python.exe -m pytest tests -q -m "not real and not excel"`.

---

## Parte A — só texto

### Tarefa 1: Regra de várias máquinas e horários escalonados nos documentos

**Arquivos:**
- Modificar: `INSTALAR.md` (pergunta 1 do Passo 3, pergunta 4 do Passo 3, Passo 9)
- Modificar: `runbooks/tarefa-agendada.md` (parágrafo inicial)
- Modificar: `NOVA-PRODUTORA.md` (item 3 da seção 6)
- Teste: `tests/test_docs_instalacao.py`

- [ ] **Passo 1: escrever os testes que falham**

Em `tests/test_docs_instalacao.py`, troque a lista de trechos de `test_instalar_cobre_os_passos_e_as_travas` (substitua `"duas máquinas"` por `"máquina principal"`) e acrescente a asserção da regra antiga ao final dessa função:

```python
def test_instalar_cobre_os_passos_e_as_travas():
    t = _ler("INSTALAR.md")
    for trecho in ("Nunca abra o `.env`", "winget", "py -0p", ".venv/Scripts/python.exe -m pip install",
                   "papaki_dre configurar", "papaki_dre verificar", "notepad .env", "pode importar",
                   "Checklist da primeira execução", "tarefa agendada", "runbooks/tarefa-agendada.md",
                   "máquina principal", "extracoes", "apuracoes", "Modelo DRE base", "--sem-banco",
                   "git pull", "NOVA-PRODUTORA.md"):
        assert trecho in t, trecho
    assert t.index("Passo 5") < t.index("Passo 6") < t.index("Passo 7") < t.index("Passo 9")
    assert "Só **uma** máquina por produtora" not in t and "Por isso, uma máquina por produtora" not in t
```

Troque `test_modelo_da_tarefa_agendada` e `test_nova_produtora_nao_adivinha`:

```python
def test_modelo_da_tarefa_agendada():
    t = _ler("runbooks/tarefa-agendada.md")
    for trecho in ("0 8 2-9 * 1-5", "<PASTA_DO_PROJETO>", "<EMAIL_DO_USUARIO>", "<PRODUTORA>",
                   "agenda --produtora <PRODUTORA> --restaurar-banco", "Nunca digite senha",
                   "máquina principal", "máquinas de grupo", "Horários escalonados"):
        assert trecho in t, trecho
    assert "uma máquina por produtora" not in t


def test_nova_produtora_nao_adivinha():
    t = _ler("NOVA-PRODUTORA.md")
    for trecho in ("adivinhe nada", "produtoras/<id>.json", "linha_mestre", "nucleo_rra", "modelo_comissoes",
                   "comissoes_excluir_h", "plano próprio", "produtora mostrar", "públic", "em quais máquinas"):
        assert trecho in t, trecho
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_docs_instalacao.py -q`
Esperado: FALHA em `test_instalar_cobre_os_passos_e_as_travas`, `test_modelo_da_tarefa_agendada` e `test_nova_produtora_nao_adivinha`.

- [ ] **Passo 3: editar `runbooks/tarefa-agendada.md`**

Substitua o parágrafo das linhas 3 a 5:

```
Para o Claude criar a tarefa no **Claude desta máquina** (ferramenta de tarefas agendadas do aplicativo), **só com o OK do
usuário** e só se esta for a máquina responsável por rodar a produtora: **uma máquina por produtora**. Duas máquinas
agendando a mesma produtora podem extrair juntas no mesmo dia antes de qualquer importação gravada.
```

por:

```
Para o Claude criar a tarefa no **Claude desta máquina** (ferramenta de tarefas agendadas do aplicativo), **só com o OK do
usuário** e só para as produtoras que esta máquina deve rodar. A Papaki tem uma **máquina principal**, que roda todas as
produtoras, e pode ter **máquinas de grupo**, que rodam só algumas. Uma mesma produtora pode, portanto, ter tarefa em
mais de uma máquina: isso é intencional. Cada produtora tem a **própria tarefa** em cada máquina.

## Horários escalonados (várias máquinas e várias produtoras)

- **Máquina de grupo:** agenda mais cedo (por exemplo, `0 7 2-9 * 1-5`).
- **Máquina principal:** agenda mais tarde (por exemplo, `0 10 2-9 * 1-5`). Se o grupo já importou, o `agenda` da principal
  vê a competência importada no banco central e termina em silêncio; se o grupo falhou, a principal assume sozinha.
- **Produtoras diferentes na mesma máquina:** horários diferentes, para o Chrome e o Excel não serem usados ao mesmo tempo
  (comece com 30 minutos entre uma e outra e ajuste pelo tempo real de uma extração).
- **Limite:** o horário só reduz o risco. Uma máquina que ainda esteja extraindo quando a outra acorda só aparece no banco
  depois da importação.

## Modelo da tarefa
```

(O título `## Modelo da tarefa` fecha a seção nova e vem logo antes do trecho "Troque os campos entre `<>`:", que continua como está.)

- [ ] **Passo 4: editar `INSTALAR.md`**

(a) Na pergunta 1 do Passo 3, substitua

```
   somam dezenas de MB por mês) e, de preferência, fora de pasta sincronizada em tempo real (o OneDrive pode travar
   arquivo aberto). O sistema cria, dentro dela,
```

por

```
   somam dezenas de MB por mês) e, como as máquinas da Papaki **compartilham** esta pasta, em um lugar com sincronização
   confiável (disco de rede ou serviço de sincronização): cada máquina aponta o próprio caminho para a **mesma** pasta, e
   ninguém abre as cópias do Excel enquanto a rotina roda. Quem decide se a rotina roda é o banco central, nunca só a
   pasta. O sistema cria, dentro dela,
```

(b) Substitua a pergunta 4 inteira (as duas linhas que começam em `4. **Rotina mensal automática:**`):

```
4. **Rotina mensal automática:** "Esta máquina será a **principal** (roda todas as produtoras) ou uma máquina de
   **grupo** (roda só algumas)? Quais produtoras? Pode haver mais de uma máquina para a mesma produtora: a de grupo
   roda mais cedo e a principal mais tarde." (Se nada será agendado aqui, pule o passo 9: nesta máquina só o disparo manual.)
```

(c) No título do Passo 9, troque `## Passo 9: tarefa agendada (só na máquina responsável e com o OK do usuário)` por
`## Passo 9: tarefa agendada (só para as produtoras que esta máquina deve rodar e com o OK do usuário)`.

(d) No fim do primeiro parágrafo do Passo 9, substitua a partir de `**Risco conhecido:**` até o fim do parágrafo por:

```
**Várias máquinas:** a máquina principal roda todas as produtoras e as de grupo rodam algumas; use
horários diferentes (grupo mais cedo, principal mais tarde), como em `runbooks/tarefa-agendada.md`. **Limite conhecido:**
só o horário evita que duas máquinas extraiam juntas; se uma ainda estiver extraindo quando a outra acordar, o banco só
percebe depois da importação.
```

- [ ] **Passo 5: editar `NOVA-PRODUTORA.md`**

Substitua o item 3 da seção 6:

```
3. Tarefa agendada **própria** da produtora (Passo 9, `runbooks/tarefa-agendada.md`), só com o OK do usuário e se esta for a
   máquina responsável por ela.
```

por:

```
3. Pergunte **em quais máquinas** a produtora vai rodar (a principal roda todas; pode haver máquinas de grupo) e crie a
   tarefa agendada **própria** da produtora em cada uma (Passo 9, `runbooks/tarefa-agendada.md`, com horários escalonados),
   só com o OK do usuário. Quem usa cada máquina repete o `configurar --modelo` dela.
```

- [ ] **Passo 6: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_docs_instalacao.py tests/test_runbooks.py -q`
Esperado: tudo passa.

- [ ] **Passo 7: auditoria e commit (só se o usuário pedir)**

```bash
.venv/Scripts/python.exe scripts/auditar_publicacao.py
git add INSTALAR.md runbooks/tarefa-agendada.md NOVA-PRODUTORA.md tests/test_docs_instalacao.py
git commit -m "docs: várias máquinas, máquina principal e horários escalonados" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

## Parte B — reserva atômica e identidade da máquina

### Tarefa 2: Migração `0002_reservas.sql`

**Arquivos:**
- Criar: `supabase/migrations/0002_reservas.sql`
- Teste: `tests/test_migracao_0002.py` (novo)

- [ ] **Passo 1: escrever o teste que falha**

`tests/test_migracao_0002.py`:

```python
"""A lógica SQL só roda no Postgres (ensaio no banco de teste, Tarefa 11); aqui só se garante que o arquivo da
migração existe e traz as proteções combinadas no desenho."""
from pathlib import Path

SQL = (Path(__file__).resolve().parents[1] / "supabase" / "migrations" / "0002_reservas.sql").read_text(
    encoding="utf-8")


def test_cria_a_tabela_a_coluna_e_o_indice_unico_parcial():
    assert "alter table public.importacoes add column maquina text" in SQL
    assert "create table public.reservas_rodada" in SQL
    assert "create unique index reservas_rodada_uma_ativa" in SQL and "where status = 'ativa'" in SQL


def test_rls_ligado_sem_politicas_e_sem_acesso_publico():
    assert "alter table public.reservas_rodada enable row level security" in SQL
    assert "create policy" not in SQL.lower()
    assert "revoke all on table public.reservas_rodada from anon, authenticated" in SQL
    for funcao in ("reservar_rodada(text, text, text, integer)", "liberar_rodada(text, text, text)"):
        assert f"revoke all on function public.{funcao} from public, anon, authenticated" in SQL
        assert f"grant execute on function public.{funcao} to service_role" in SQL


def test_as_funcoes_travam_por_produtora_e_competencia_e_usam_o_relogio_do_banco():
    assert SQL.count("pg_advisory_xact_lock") == 2
    assert "now()" in SQL and "p_horas < 1 or p_horas > 24" in SQL
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_migracao_0002.py -q`
Esperado: FALHA (`FileNotFoundError`: o arquivo não existe).

- [ ] **Passo 3: criar `supabase/migrations/0002_reservas.sql`**

```sql
-- Plano 6 (várias máquinas): reserva da rodada (extração + importação) por produtora e competência, para duas
-- máquinas não extraírem a mesma produtora ao mesmo tempo. Mesmas regras de segurança da 0001: RLS ligado e SEM
-- políticas; só a chave de serviço (service_role) acessa. Só guarda ids de produtora, competência e o nome da máquina.

alter table public.importacoes add column maquina text;   -- quem importou (PAPAKI_MAQUINA); nulo nas rodadas antigas

create table public.reservas_rodada (
  id           uuid primary key default gen_random_uuid(),
  produtora_id text not null,   -- sem chave estrangeira: a linha de `produtoras` só nasce na primeira importação
  competencia  text not null check (competencia ~ '^\d{4}-(0[1-9]|1[0-2])$'),
  maquina      text not null check (char_length(maquina) between 1 and 40),
  criada_em    timestamptz not null default now(),
  expira_em    timestamptz not null,
  status       text not null default 'ativa' check (status in ('ativa', 'liberada', 'expirada'))
);
alter table public.reservas_rodada enable row level security;
-- no máximo uma reserva ativa por produtora e competência
create unique index reservas_rodada_uma_ativa on public.reservas_rodada (produtora_id, competencia)
  where status = 'ativa';

-- Reserva a rodada para a máquina. Devolve {"reservada": bool, "maquina": quem a tem, "expira_em": instante}.
-- A mesma máquina renova; outra máquina só consegue depois que a reserva expira ou é liberada.
-- O tempo é o do banco (now()), nunca o relógio das máquinas.
create or replace function public.reservar_rodada(p_produtora text, p_competencia text, p_maquina text,
                                                  p_horas integer default 6) returns jsonb
language plpgsql
set search_path = public
as $$
declare
  v_atual  public.reservas_rodada%rowtype;
  v_expira timestamptz;
begin
  if p_horas is null or p_horas < 1 or p_horas > 24 then
    raise exception 'p_horas deve estar entre 1 e 24';
  end if;
  -- duas reservas da mesma produtora e competência não correm ao mesmo tempo
  perform pg_advisory_xact_lock(hashtextextended('reserva|' || p_produtora || '|' || p_competencia, 0));
  update public.reservas_rodada set status = 'expirada'
    where produtora_id = p_produtora and competencia = p_competencia and status = 'ativa' and expira_em <= now();
  select * into v_atual from public.reservas_rodada
    where produtora_id = p_produtora and competencia = p_competencia and status = 'ativa';
  v_expira := now() + make_interval(hours => p_horas);
  if found then
    if v_atual.maquina = p_maquina then
      update public.reservas_rodada set expira_em = v_expira where id = v_atual.id;
      return jsonb_build_object('reservada', true, 'maquina', p_maquina, 'expira_em', v_expira);
    end if;
    return jsonb_build_object('reservada', false, 'maquina', v_atual.maquina, 'expira_em', v_atual.expira_em);
  end if;
  insert into public.reservas_rodada (produtora_id, competencia, maquina, expira_em)
    values (p_produtora, p_competencia, p_maquina, v_expira);
  return jsonb_build_object('reservada', true, 'maquina', p_maquina, 'expira_em', v_expira);
end;
$$;
revoke all on function public.reservar_rodada(text, text, text, integer) from public, anon, authenticated;
grant execute on function public.reservar_rodada(text, text, text, integer) to service_role;

-- Libera a reserva ativa da própria máquina. Devolve true se havia reserva dela para liberar.
create or replace function public.liberar_rodada(p_produtora text, p_competencia text, p_maquina text)
returns boolean
language plpgsql
set search_path = public
as $$
declare
  v_n integer;
begin
  perform pg_advisory_xact_lock(hashtextextended('reserva|' || p_produtora || '|' || p_competencia, 0));
  update public.reservas_rodada set status = 'liberada'
    where produtora_id = p_produtora and competencia = p_competencia and maquina = p_maquina
      and status = 'ativa';
  get diagnostics v_n = row_count;
  return v_n > 0;
end;
$$;
revoke all on function public.liberar_rodada(text, text, text) from public, anon, authenticated;
grant execute on function public.liberar_rodada(text, text, text) to service_role;

-- Defesa em profundidade (igual à 0001): anon e authenticated ficam sem privilégio na tabela nova.
revoke all on table public.reservas_rodada from anon, authenticated;
```

- [ ] **Passo 4: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_migracao_0002.py -q`
Esperado: 3 passam.

- [ ] **Passo 5: commit (só se o usuário pedir)**

```bash
git add supabase/migrations/0002_reservas.sql tests/test_migracao_0002.py
git commit -m "feat(banco): migração 0002 com reservas de rodada e coluna da máquina" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 3: Cliente falso e `Repository.reservar_rodada` / `liberar_rodada`

**Arquivos:**
- Modificar: `tests/fake_supabase.py` (`import`, `_Rpc`, `FakeClient.__init__`)
- Modificar: `src/papaki_dre/db.py` (dois métodos novos em `Repository`)
- Teste: `tests/test_reservas_db.py` (novo)

- [ ] **Passo 1: escrever os testes que falham**

`tests/test_reservas_db.py`:

```python
import datetime as dt

from fake_supabase import FakeClient
from papaki_dre.db import Repository

T0 = dt.datetime(2026, 11, 3, 8, 0, tzinfo=dt.timezone.utc)


def _banco():
    """(repo, client, relogio): o relogio é o 'now()' do banco falso; os testes o adiantam."""
    relogio = {"agora": T0}
    client = FakeClient()
    client.agora = lambda: relogio["agora"]
    return Repository(client), client, relogio


def test_primeira_reserva_e_obtida():
    repo, client, _ = _banco()
    r = repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    assert r["reservada"] is True and r["maquina"] == "principal"
    assert dt.datetime.fromisoformat(r["expira_em"]) == T0 + dt.timedelta(hours=6)
    assert [x["status"] for x in client.tables["reservas_rodada"]] == ["ativa"]


def test_outra_maquina_nao_consegue_e_ve_quem_tem():
    repo, _, _ = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    r = repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    assert r["reservada"] is False and r["maquina"] == "grupo-a"


def test_a_mesma_maquina_renova_o_prazo():
    repo, _, relogio = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    relogio["agora"] = T0 + dt.timedelta(hours=2)
    r = repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    assert r["reservada"] is True
    assert dt.datetime.fromisoformat(r["expira_em"]) == T0 + dt.timedelta(hours=8)


def test_reserva_vencida_deixa_outra_maquina_assumir():
    repo, client, relogio = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    relogio["agora"] = T0 + dt.timedelta(hours=6, minutes=1)
    r = repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    assert r["reservada"] is True and r["maquina"] == "principal"
    assert sorted(x["status"] for x in client.tables["reservas_rodada"]) == ["ativa", "expirada"]


def test_liberar_deixa_outra_maquina_reservar():
    repo, _, _ = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    assert repo.liberar_rodada("untitled", "2026-10", "grupo-a") is True
    assert repo.reservar_rodada("untitled", "2026-10", "principal", 6)["reservada"] is True


def test_liberar_so_vale_para_a_propria_maquina():
    repo, _, _ = _banco()
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    assert repo.liberar_rodada("untitled", "2026-10", "principal") is False
    assert repo.reservar_rodada("untitled", "2026-10", "principal", 6)["reservada"] is False


def test_liberar_sem_reserva_devolve_falso():
    repo, _, _ = _banco()
    assert repo.liberar_rodada("untitled", "2026-10", "principal") is False


def test_produtoras_e_competencias_diferentes_sao_independentes():
    repo, _, _ = _banco()
    assert repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)["reservada"] is True
    assert repo.reservar_rodada("outra", "2026-10", "principal", 6)["reservada"] is True
    assert repo.reservar_rodada("untitled", "2026-11", "principal", 6)["reservada"] is True
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_reservas_db.py -q`
Esperado: FALHA (`AttributeError: 'Repository' object has no attribute 'reservar_rodada'`).

- [ ] **Passo 3: ensinar o cliente falso**

Em `tests/fake_supabase.py`, logo abaixo de `import copy`, acrescente `import datetime as dt` (a ordem fica `copy`, `datetime`, `json`, `uuid`):

```python
import copy
import datetime as dt
import json
import uuid
```

Troque a classe `_Rpc` inteira por:

```python
class _Rpc:
    def __init__(self, client, name, params):
        self.client, self.name, self.params = client, name, _wire(params)

    def execute(self):
        assert self.name in ("promover_importacao", "reservar_rodada", "liberar_rodada"), self.name
        self.client.calls.append(("rpc", self.name, None))
        return getattr(self, "_" + self.name)()

    def _promover_importacao(self):
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

    def _reservar_rodada(self):
        """Espelha supabase/migrations/0002_reservas.sql: expira as vencidas, renova a da mesma máquina, nega a
        de outra e senão cria. O tempo é o de `client.agora()` (o `now()` do banco)."""
        p = self.params
        agora = self.client.agora()
        expira = (agora + dt.timedelta(hours=p["p_horas"])).isoformat()
        reservas = self.client.tables.setdefault("reservas_rodada", [])
        da_rodada = [r for r in reservas if r["produtora_id"] == p["p_produtora"]
                     and r["competencia"] == p["p_competencia"] and r["status"] == "ativa"]
        for r in da_rodada:
            if dt.datetime.fromisoformat(r["expira_em"]) <= agora:
                r["status"] = "expirada"
        atual = next((r for r in da_rodada if r["status"] == "ativa"), None)
        if atual is not None and atual["maquina"] == p["p_maquina"]:
            atual["expira_em"] = expira
            return _Resp({"reservada": True, "maquina": p["p_maquina"], "expira_em": expira})
        if atual is not None:
            return _Resp({"reservada": False, "maquina": atual["maquina"], "expira_em": atual["expira_em"]})
        reservas.append({"id": str(uuid.uuid4()), "produtora_id": p["p_produtora"],
                         "competencia": p["p_competencia"], "maquina": p["p_maquina"],
                         "criada_em": agora.isoformat(), "expira_em": expira, "status": "ativa"})
        return _Resp({"reservada": True, "maquina": p["p_maquina"], "expira_em": expira})

    def _liberar_rodada(self):
        p = self.params
        achou = False
        for r in self.client.tables.get("reservas_rodada", []):
            if (r["produtora_id"], r["competencia"], r["maquina"], r["status"]) == (
                    p["p_produtora"], p["p_competencia"], p["p_maquina"], "ativa"):
                r["status"] = "liberada"
                achou = True
        return _Resp(achou)
```

Em `FakeClient.__init__`, depois de `self.returning: dict[str, list] = {}  # ...`, acrescente:

```python
        self.agora = lambda: dt.datetime.now(dt.timezone.utc)   # o "now()" do banco; os testes o trocam
```

- [ ] **Passo 4: implementar em `src/papaki_dre/db.py`**

Em `Repository`, logo antes de `def salvar_conferencia(self, importacao_id, conferencia) -> None:`, acrescente:

```python
    def reservar_rodada(self, produtora, competencia, maquina, horas) -> dict:
        """Reserva a rodada (extração + importação) para `maquina` por `horas` horas (função SQL atômica; o
        tempo é o do banco). A mesma máquina renova. Devolve {"reservada": bool, "maquina": quem a tem,
        "expira_em": ISO}."""
        res = self.client.rpc("reservar_rodada", {"p_produtora": produtora, "p_competencia": competencia,
                                                  "p_maquina": maquina, "p_horas": horas}).execute()
        return res.data

    def liberar_rodada(self, produtora, competencia, maquina) -> bool:
        """Libera a reserva ativa da própria `maquina`. True se havia uma reserva dela para liberar."""
        res = self.client.rpc("liberar_rodada", {"p_produtora": produtora, "p_competencia": competencia,
                                                 "p_maquina": maquina}).execute()
        return bool(res.data)

```

- [ ] **Passo 5: rodar e ver passar (e conferir que nada quebrou)**

Run: `.venv/Scripts/python.exe -m pytest tests/test_reservas_db.py tests/test_cli.py tests/test_db.py tests/test_agenda_banco.py -q`
Esperado: tudo passa (a promoção continua funcionando pelo cliente falso reorganizado).

- [ ] **Passo 6: commit (só se o usuário pedir)**

```bash
git add tests/fake_supabase.py tests/test_reservas_db.py src/papaki_dre/db.py
git commit -m "feat(db): reservar_rodada e liberar_rodada no Repository" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 4: `criar_importacao` grava a máquina

**Arquivos:**
- Modificar: `src/papaki_dre/db.py` (`criar_importacao`)
- Teste: `tests/test_reservas_db.py` (acrescentar)

- [ ] **Passo 1: escrever os testes que falham** (acrescente ao fim de `tests/test_reservas_db.py`)

```python
def _criar(repo, **extra):
    return repo.criar_importacao(produtora="untitled", competencia="2026-10", periodo_inicio="2026-01-01",
                                 periodo_fim="2026-10-31", data_extracao="2026-11-03T10:00:00-03:00",
                                 origem="teste", gatilho="agendado", **extra)


def test_criar_importacao_grava_a_maquina():
    repo, client, _ = _banco()
    _criar(repo, maquina="grupo-a")
    assert client.tables["importacoes"][0]["maquina"] == "grupo-a"


def test_criar_importacao_sem_maquina_nao_manda_a_coluna():
    """Banco ainda sem a migração 0002: a coluna nem pode ser enviada."""
    repo, client, _ = _banco()
    _criar(repo)
    assert "maquina" not in client.tables["importacoes"][0]
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_reservas_db.py -q`
Esperado: FALHA em `test_criar_importacao_grava_a_maquina` (`TypeError: ... unexpected keyword argument 'maquina'`).

- [ ] **Passo 3: implementar** — troque `criar_importacao` inteira em `src/papaki_dre/db.py` por:

```python
    def criar_importacao(self, *, produtora, competencia, periodo_inicio, periodo_fim,
                         data_extracao, origem, gatilho, maquina=None) -> str:
        linha = {
            "produtora_id": produtora, "competencia": competencia, "periodo_inicio": periodo_inicio,
            "periodo_fim": periodo_fim, "data_extracao": data_extracao, "origem": origem,
            "gatilho": gatilho}
        if maquina:                       # coluna da migração 0002: sem nome da máquina, nada a gravar
            linha["maquina"] = maquina
        res = self.client.table("importacoes").insert(linha).execute()
        return res.data[0]["id"]
```

- [ ] **Passo 4: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_reservas_db.py tests/test_db.py tests/test_cli.py -q`
Esperado: tudo passa.

- [ ] **Passo 5: commit (só se o usuário pedir)**

```bash
git add src/papaki_dre/db.py tests/test_reservas_db.py
git commit -m "feat(db): criar_importacao registra a máquina quando informada" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 5: `PAPAKI_MAQUINA` e `configurar --maquina`

**Arquivos:**
- Modificar: `src/papaki_dre/ambiente.py` (`CHAVES_GRAVAVEIS`)
- Modificar: `src/papaki_dre/cli_instalacao.py` (opção e validação)
- Modificar: `.env.example`
- Teste: `tests/test_cli_instalacao.py`, `tests/test_docs_instalacao.py`

- [ ] **Passo 1: escrever os testes que falham**

Acrescente ao fim de `tests/test_cli_instalacao.py`:

```python
def test_configurar_grava_o_nome_da_maquina(capsys, tmp_path):
    arq_env = tmp_path / ".env"
    codigo, out, err = _rodar(capsys, ["configurar", "--maquina", "grupo-a", "--arquivo-env", str(arq_env)])
    assert codigo == 0, err
    assert json.loads(out)["gravadas"] == ["PAPAKI_MAQUINA"]
    assert "PAPAKI_MAQUINA=grupo-a" in arq_env.read_text(encoding="utf-8")


@pytest.mark.parametrize("nome", ["com espaço", "-comeca-com-traco", "a" * 41, "acentuação"])
def test_configurar_recusa_nome_de_maquina_invalido(capsys, tmp_path, nome):
    codigo, out, err = _rodar(capsys, ["configurar", "--maquina", nome, "--arquivo-env", str(tmp_path / ".env")])
    assert codigo == 3 and "--maquina" in err and out == ""
    assert not (tmp_path / ".env").exists()
```

Em `tests/test_docs_instalacao.py`, dentro de `test_opcoes_do_configurar_e_do_agenda_citadas_existem`, acrescente antes da última linha de `for opcao in ...` a asserção de parse e inclua `"--maquina"` na tupla de opções exigidas em `INSTALAR.md` (isso será satisfeito na Tarefa 10; aqui só a parte do parser):

```python
    assert p.parse_args(["configurar", "--maquina", "principal"]).maquina == "principal"
```

E em `test_env_example_lista_as_chaves_sem_valores_reais`, inclua `"PAPAKI_MAQUINA"` na tupla de chaves e no `assert all(ambiente.chave_gravavel(...))`.

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_cli_instalacao.py tests/test_docs_instalacao.py -q`
Esperado: FALHA (opção desconhecida; chave ausente do `.env.example`).

- [ ] **Passo 3: implementar**

`src/papaki_dre/ambiente.py` — troque `CHAVES_GRAVAVEIS`:

```python
CHAVES_GRAVAVEIS = ("PAPAKI_DADOS_DIR", "PAPAKI_PASTA_EXTRACOES", "PAPAKI_PASTA_APURACOES",
                    "PAPAKI_USUARIO_ESFERAS", "PAPAKI_MAQUINA")
```

`src/papaki_dre/cli_instalacao.py`:

1. Logo abaixo de `_EMAIL = re.compile(...)`, acrescente:

```python
_NOME_MAQUINA = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{0,39}")
```

2. Em `registrar`, depois da linha de `--usuario-esferas`, acrescente:

```python
    c.add_argument("--maquina", help="nome curto desta máquina (ex.: principal, grupo-a); vai para as reservas e "
                                     "para a importação")
```

3. Em `_configurar`, depois do bloco de `a.usuario_esferas`, acrescente:

```python
    if a.maquina:
        if not _NOME_MAQUINA.fullmatch(a.maquina.strip()):
            raise ValueError("--maquina: use até 40 letras sem acento, números, '_' ou '-', começando por letra "
                             "ou número")
        pares["PAPAKI_MAQUINA"] = a.maquina.strip()
```

`.env.example` — acrescente ao final:

```
# Nome curto desta máquina (ex.: principal, grupo-a). Identifica quem reservou e quem importou cada produtora.
PAPAKI_MAQUINA=<nome curto desta máquina>
```

- [ ] **Passo 4: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_cli_instalacao.py tests/test_ambiente_gravar.py tests/test_docs_instalacao.py -q`
Esperado: tudo passa, exceto a exigência de `"--maquina"` em `INSTALAR.md`, que só entra na Tarefa 10 (se você já a incluiu no teste, deixe esse trecho para a Tarefa 10 e rode de novo lá).

- [ ] **Passo 5: commit (só se o usuário pedir)**

```bash
git add src/papaki_dre/ambiente.py src/papaki_dre/cli_instalacao.py .env.example tests/test_cli_instalacao.py tests/test_docs_instalacao.py
git commit -m "feat(config): PAPAKI_MAQUINA e configurar --maquina" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 6: `agenda --reservar` e `reserva obter`

**Arquivos:**
- Modificar: `src/papaki_dre/cli.py` (constante, `_reservar_rodada`, `main`)
- Modificar: `src/papaki_dre/cli_extracao.py` (`registrar`, `executar`, `_agenda`, `_reserva`)
- Teste: `tests/test_reserva_cli.py` (novo)

- [ ] **Passo 1: escrever os testes que falham**

`tests/test_reserva_cli.py`:

```python
import json

import pytest

from fake_supabase import FakeClient
from papaki_dre import cli
from papaki_dre.db import Repository


@pytest.fixture
def env(tmp_path):
    return {"PAPAKI_DADOS_DIR": str(tmp_path / "dados"), "PAPAKI_MAQUINA": "principal"}


def _rodar(capsys, argv, env, client):
    codigo = cli.main(argv, env=env, client=client)
    saida = capsys.readouterr()
    return codigo, (json.loads(saida.out) if saida.out.strip() else None), saida.err


def _agenda(capsys, env, client, *extra, hoje="2026-11-04"):
    return _rodar(capsys, ["agenda", "--produtora", "untitled", "--hoje", hoje, *extra], env, client)


def test_sem_reservar_o_agenda_nao_mexe_nas_reservas(capsys, env):
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client)
    assert d["reserva"] == "nao_pedida" and d["rodar"] is True
    assert client.tables.get("reservas_rodada", []) == []


def test_reservar_com_a_rodada_livre_obtem_a_reserva(capsys, env):
    client = FakeClient()
    codigo, d, _ = _agenda(capsys, env, client, "--reservar")
    assert codigo == 0 and d["rodar"] is True and d["reserva"] == "obtida"
    reservas = client.tables["reservas_rodada"]
    assert [(r["produtora_id"], r["competencia"], r["maquina"], r["status"]) for r in reservas] == [
        ("untitled", "2026-10", "principal", "ativa")]


def test_reserva_de_outra_maquina_manda_nao_rodar(capsys, env):
    client = FakeClient()
    Repository(client).reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    codigo, d, _ = _agenda(capsys, env, client, "--reservar")
    assert codigo == 0 and d["rodar"] is False and d["reserva"] == "de_outra_maquina"
    assert "outra máquina" in d["motivo"]


def test_a_propria_reserva_deixa_rodar_de_novo(capsys, env):
    client = FakeClient()
    _agenda(capsys, env, client, "--reservar")
    _, d, _ = _agenda(capsys, env, client, "--reservar")
    assert d["rodar"] is True and d["reserva"] == "obtida"


def test_sem_nome_da_maquina_segue_e_avisa(capsys, tmp_path):
    env = {"PAPAKI_DADOS_DIR": str(tmp_path / "dados")}
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client, "--reservar")
    assert d["rodar"] is True and d["reserva"] == "sem_nome_da_maquina"
    assert client.tables.get("reservas_rodada", []) == []


def test_banco_indisponivel_segue_e_avisa(capsys, env):
    _, d, err = _agenda(capsys, env, FakeClient(fail_with=ConnectionError("sem rede")), "--reservar")
    assert d["rodar"] is True and d["banco"] == "indisponivel" and d["reserva"] == "nao_consultada"
    assert "sem rede" not in err


def test_sem_banco_nao_reserva(capsys, env):
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client, "--reservar", "--sem-banco")
    assert d["banco"] == "pulado" and d["reserva"] == "nao_consultada" and d["rodar"] is True


def test_competencia_ja_importada_nao_reserva(capsys, env):
    client = FakeClient()
    repo = Repository(client)
    imp = repo.criar_importacao(produtora="untitled", competencia="2026-10", periodo_inicio="2026-01-01",
                                periodo_fim="2026-10-31", data_extracao="2026-11-03T10:00:00-03:00",
                                origem="teste", gatilho="agendado")
    repo.promover(imp)
    _, d, _ = _agenda(capsys, env, client, "--reservar")
    assert d["rodar"] is False and d["reserva"] == "nao_pedida"
    assert client.tables.get("reservas_rodada", []) == []


def test_fora_do_dia_de_rodar_nao_reserva(capsys, env):
    client = FakeClient()
    _, d, _ = _agenda(capsys, env, client, "--reservar", hoje="2026-11-02")      # feriado
    assert d["rodar"] is False and d["reserva"] == "nao_pedida"
    assert client.calls == []


def test_reserva_obter_para_a_competencia(capsys, env):
    client = FakeClient()
    codigo, d, _ = _rodar(capsys, ["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-10"],
                          env, client)
    assert codigo == 0 and d["estado"] == "obtida" and d["maquina"] == "principal"
    assert d["produtora"] == "untitled" and d["competencia"] == "2026-10" and d["expira_em"]


def test_reserva_obter_mostra_quem_tem_a_reserva(capsys, env):
    client = FakeClient()
    Repository(client).reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    _, d, _ = _rodar(capsys, ["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-10"],
                     env, client)
    assert d["estado"] == "de_outra_maquina" and d["maquina"] == "grupo-a"


@pytest.mark.parametrize("argv, trecho", [
    (["reserva", "obter", "--produtora", "naoexiste", "--competencia", "2026-10"], "produtora desconhecida"),
    (["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-13"], "2026-13"),
])
def test_reserva_obter_erros_de_uso(capsys, env, argv, trecho):
    codigo, d, err = _rodar(capsys, argv, env, FakeClient())
    assert codigo == 3 and d is None and trecho in err
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_reserva_cli.py -q`
Esperado: FALHA (`ERRO DE USO` por argumento desconhecido `--reservar` / comando `reserva`).

- [ ] **Passo 3: implementar em `src/papaki_dre/cli.py`**

1. Junto das constantes (depois de `INTERVALO_APOS_ACORDAR = 6 ...`), acrescente:

```python
HORAS_RESERVA = 6                         # validade da reserva da rodada (renovada pela mesma máquina)
```

2. Logo acima de `def importar(a, client=None, admin=None, env=None) -> int:`, acrescente:

```python
def _nome_da_maquina(env=None) -> str | None:
    """`PAPAKI_MAQUINA` do `.env`/ambiente, ou None se não foi configurada."""
    return (ambiente.resolver(env).get("PAPAKI_MAQUINA") or "").strip() or None


def _reservar_rodada(produtora: str, competencia: str, *, env=None, client=None) -> dict:
    """Para o `agenda --reservar` e o `reserva obter`: tenta reservar a rodada para esta máquina. Nunca levanta e
    nenhum texto de exceção sai daqui. estado = obtida | de_outra_maquina | sem_nome_da_maquina | nao_consultada
    (sem banco, a defesa é só o escalonamento dos horários)."""
    maquina = _nome_da_maquina(env)
    if not maquina:
        return {"estado": "sem_nome_da_maquina"}
    try:
        repo = Repository(client if client is not None else make_client(env))
        r = repo.reservar_rodada(produtora, competencia, maquina, HORAS_RESERVA)
    except Exception:  # noqa: BLE001 - BancoIndisponivel, rede, migração 0002 ainda não aplicada...
        return {"estado": "nao_consultada"}
    return {"estado": "obtida" if r.get("reservada") else "de_outra_maquina",
            "maquina": r.get("maquina"), "expira_em": r.get("expira_em")}
```

3. Em `main`, troque as duas últimas linhas do bloco `try` (a partir de `return cli_extracao.executar(...)`) por:

```python
        reservar = lambda produtora, competencia: _reservar_rodada(  # noqa: E731
            produtora, competencia, env=env, client=client)
        return cli_extracao.executar(args, env, consulta_banco=consulta, reservar_banco=reservar)
```

- [ ] **Passo 4: implementar em `src/papaki_dre/cli_extracao.py`**

1. Em `registrar`, troque o trecho

```python
    ag.add_argument("--restaurar-banco", action="store_true",
                    help="se o banco estiver pausado, acorda o projeto antes de consultar (exige "
                         "SUPABASE_ACCESS_TOKEN no .env)")

    ex = sub.add_parser("extracao", help="registro da extração do Esferas (usado pelo Claude)")
```

por

```python
    ag.add_argument("--restaurar-banco", action="store_true",
                    help="se o banco estiver pausado, acorda o projeto antes de consultar (exige "
                         "SUPABASE_ACCESS_TOKEN no .env)")
    ag.add_argument("--reservar", action="store_true",
                    help="se for dia de rodar, reserva a rodada no banco central para esta máquina (PAPAKI_MAQUINA); "
                         "se outra máquina já a reservou, responde 'não rodar'")

    rv = sub.add_parser("reserva", help="reserva da rodada de uma produtora e competência (usado pelo Claude)")
    rvs = rv.add_subparsers(dest="acao", required=True)
    o = rvs.add_parser("obter", help="tenta reservar a rodada para esta máquina e diz quem a tem")
    o.add_argument("--produtora", required=True)
    o.add_argument("--competencia", required=True, help="AAAA-MM")

    ex = sub.add_parser("extracao", help="registro da extração do Esferas (usado pelo Claude)")
```

2. Troque `executar`:

```python
def executar(a, env=None, consulta_banco=None, reservar_banco=None) -> int:
    try:
        if a.cmd == "agenda":
            return _agenda(a, env, consulta_banco, reservar_banco)
        if a.cmd == "reserva":
            return _reserva(a, reservar_banco)
        return _extracao(a, env)
```

(o restante de `executar`, os `except`, não muda).

3. Troque a assinatura `def _agenda(a, env, consulta_banco=None) -> int:` por `def _agenda(a, env, consulta_banco=None, reservar_banco=None) -> int:` e, no mesmo corpo, troque

```python
        if ja_no_banco:
            rodar, motivo = False, "competência já importada (banco central)"
    _json({"rodar": rodar, "competencia": competencia, "motivo": motivo, "banco": banco,
           "pasta_apuracao": str(apuracao),
```

por

```python
        if ja_no_banco:
            rodar, motivo = False, "competência já importada (banco central)"
    reserva = "nao_pedida"
    if rodar and a.reservar:
        reserva = "nao_consultada"
        if banco == "consultado" and reservar_banco is not None:
            reserva = reservar_banco(a.produtora, competencia)["estado"]
            if reserva == "de_outra_maquina":
                rodar, motivo = False, "rodada em andamento em outra máquina"
    _json({"rodar": rodar, "competencia": competencia, "motivo": motivo, "banco": banco, "reserva": reserva,
           "pasta_apuracao": str(apuracao),
```

4. Logo abaixo de `_agenda` (antes de `def _extracao`), acrescente:

```python
def _reserva(a, reservar_banco) -> int:
    if not produtoras.existe(a.produtora):
        raise ExtracaoErro(f"produtora desconhecida: {a.produtora!r}")
    competencia, _, _ = calendario.periodo_da_competencia(a.competencia)   # ValueError -> ERRO (código 3)
    _json({"produtora": a.produtora, "competencia": competencia, **reservar_banco(a.produtora, competencia)})
    return EXIT_OK
```

- [ ] **Passo 5: rodar e ver passar (e a suíte de agenda/extração inteira)**

Run: `.venv/Scripts/python.exe -m pytest tests/test_reserva_cli.py tests/test_agenda_banco.py tests/test_cli_extracao.py tests/test_cli_rotina.py -q`
Esperado: tudo passa.

- [ ] **Passo 6: commit (só se o usuário pedir)**

```bash
git add src/papaki_dre/cli.py src/papaki_dre/cli_extracao.py tests/test_reserva_cli.py
git commit -m "feat(agenda): agenda --reservar e comando reserva obter" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 7: `importar` grava a máquina e libera a reserva

**Arquivos:**
- Modificar: `src/papaki_dre/cli.py` (`_liberar_reserva`, `importar`)
- Teste: `tests/test_cli.py` (acrescentar no fim, usando o fixture `env` e o `_argv` que já existem ali)

- [ ] **Passo 1: escrever os testes que falham** (fim de `tests/test_cli.py`)

```python
# --- várias máquinas: máquina gravada e reserva liberada ------------------------------------------

def _reservar(client, maquina):
    from papaki_dre.db import Repository
    return Repository(client).reservar_rodada("untitled", "2026-09", maquina, 6)


def _status_reservas(client):
    return [r["status"] for r in client.tables.get("reservas_rodada", [])]


def test_importar_grava_a_maquina_e_libera_a_reserva(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    _reservar(client, "principal")
    assert cli.main(_argv(env), client=client) == 0
    assert client.tables["importacoes"][0]["maquina"] == "principal"
    assert _status_reservas(client) == ["liberada"]


def test_importar_sem_nome_da_maquina_nao_manda_a_coluna_e_nao_mexe_na_reserva(env):
    client = FakeClient()
    _reservar(client, "grupo-a")
    assert cli.main(_argv(env), client=client) == 0
    assert "maquina" not in client.tables["importacoes"][0]
    assert _status_reservas(client) == ["ativa"]


def test_importar_nao_libera_a_reserva_de_outra_maquina(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    _reservar(client, "grupo-a")
    assert cli.main(_argv(env), client=client) == 0
    assert _status_reservas(client) == ["ativa"]


def test_importar_bloqueada_tambem_libera_a_reserva(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")

    def bloqueia(*a, **kw):
        raise ContractError("cabeçalho diferente")

    monkeypatch.setattr(cli, "sanitize_cp", bloqueia)
    client = FakeClient()
    _reservar(client, "principal")
    assert cli.main(_argv(env), client=client) == 2
    assert _status_reservas(client) == ["liberada"]


def test_importar_identico_tambem_libera_a_reserva(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    assert cli.main(_argv(env), client=client) == 0
    _reservar(client, "principal")
    assert cli.main(_argv(env), client=client) == 0           # mesmos arquivos: "Arquivos idênticos já importados"
    assert len(client.tables["importacoes"]) == 1
    assert _status_reservas(client) == ["liberada"]


def test_falha_ao_liberar_nao_derruba_a_importacao(env, monkeypatch):
    monkeypatch.setenv("PAPAKI_MAQUINA", "principal")
    client = FakeClient()
    monkeypatch.setattr("papaki_dre.db.Repository.liberar_rodada",
                        lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("banco caiu")))
    assert cli.main(_argv(env), client=client) == 0
    assert client.tables["importacoes"][0]["status"] == "concluida"
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_cli.py -q -k "maquina or reserva"`
Esperado: FALHA (`KeyError: 'maquina'` e reservas continuam `ativa`).

- [ ] **Passo 3: implementar em `src/papaki_dre/cli.py`**

1. Logo abaixo de `_reservar_rodada` (acrescentada na Tarefa 6), acrescente:

```python
def _liberar_reserva(repo, produtora: str, competencia: str, maquina: str | None) -> None:
    """Libera a reserva desta máquina ao fim do `importar` (importada, bloqueada, idêntica ou falha). Nunca
    derruba a rodada: se não conseguir, a reserva expira sozinha."""
    if repo is None or not maquina:
        return
    try:
        repo.liberar_rodada(produtora, competencia, maquina)
    except Exception:  # noqa: BLE001 - o cliente levanta tipos variados
        pass
```

2. Em `importar`, logo acima de `repo = None` (a linha `    repo = None` que vem depois do bloco `hashes = ...`), acrescente:

```python
    maquina = _nome_da_maquina(env)
```

3. No retorno dos arquivos idênticos, troque

```python
                        "nesta pasta; use --forcar para voltar a ela.")
            return EXIT_OK
```

por

```python
                        "nesta pasta; use --forcar para voltar a ela.")
            _liberar_reserva(repo, a.produtora, competencia, maquina)
            return EXIT_OK
```

4. Na criação da importação, troque

```python
                periodo_fim=fim.isoformat(), data_extracao=extracao, origem=ent["origem"], gatilho=a.gatilho)
```

por

```python
                periodo_fim=fim.isoformat(), data_extracao=extracao, origem=ent["origem"], gatilho=a.gatilho,
                maquina=maquina)
```

5. No fim do bloco `try` da rodada, troque

```python
        if substituido and not publicado and mudou is not False and conf is not None:
            _regravar_falha(out_dir, conf, motivo)
        raise
```

por

```python
        if substituido and not publicado and mudou is not False and conf is not None:
            _regravar_falha(out_dir, conf, motivo)
        raise
    finally:                                   # todo desfecho (ok, bloqueada, falha, Ctrl+C) solta a reserva
        _liberar_reserva(repo, a.produtora, competencia, maquina)
```

- [ ] **Passo 4: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_cli.py tests/test_cli_rotina.py tests/test_reserva_cli.py -q`
Esperado: tudo passa (os testes antigos do `importar` continuam passando: sem `PAPAKI_MAQUINA` nada muda).

- [ ] **Passo 5: commit (só se o usuário pedir)**

```bash
git add src/papaki_dre/cli.py tests/test_cli.py
git commit -m "feat(importar): grava a máquina e libera a reserva ao terminar" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

## Parte C — visão do grupo

### Tarefa 8: `produtora situacao`

**Arquivos:**
- Modificar: `src/papaki_dre/db.py` (`_no_futuro`, `Repository.situacao`)
- Modificar: `src/papaki_dre/cli_instalacao.py` (`registrar`, `executar`, `_produtora`, `_situacao`)
- Teste: `tests/test_produtora_situacao.py` (novo)

- [ ] **Passo 1: escrever os testes que falham**

`tests/test_produtora_situacao.py`:

```python
import datetime as dt
import json

from fake_supabase import FakeClient
from papaki_dre import cli, produtoras
from papaki_dre.db import Repository


def _situacao(capsys, client, *extra):
    codigo = cli.main(["produtora", "situacao", "--competencia", "2026-10", *extra], client=client)
    saida = capsys.readouterr()
    return codigo, (json.loads(saida.out) if saida.out.strip() else None), saida.err


def _importar(repo, maquina="grupo-a", periodo_fim="2026-10-31"):
    imp = repo.criar_importacao(produtora="untitled", competencia="2026-10", periodo_inicio="2026-01-01",
                                periodo_fim=periodo_fim, data_extracao="2026-11-03T10:00:00-03:00",
                                origem="teste", gatilho="agendado", maquina=maquina)
    repo.promover(imp)


def _item(d, produtora="untitled"):
    return next(p for p in d["produtoras"] if p["produtora"] == produtora)


def test_lista_todas_as_produtoras_cadastradas(capsys):
    codigo, d, _ = _situacao(capsys, FakeClient())
    assert codigo == 0 and d["competencia"] == "2026-10"
    assert [p["produtora"] for p in d["produtoras"]] == produtoras.ids()
    assert _item(d) == {"produtora": "untitled", "importada": False, "periodo_fim": None, "maquina": None,
                        "reserva": None}


def test_mostra_a_foto_vigente_e_quem_importou(capsys):
    client = FakeClient()
    _importar(Repository(client))
    _, d, _ = _situacao(capsys, client)
    assert _item(d) == {"produtora": "untitled", "importada": True, "periodo_fim": "2026-10-31",
                        "maquina": "grupo-a", "reserva": None}


def test_importacao_antiga_sem_maquina_mostra_nulo(capsys):
    client = FakeClient()
    _importar(Repository(client), maquina=None)
    _, d, _ = _situacao(capsys, client)
    assert _item(d)["importada"] is True and _item(d)["maquina"] is None


def test_mostra_a_reserva_ativa(capsys):
    client = FakeClient()
    Repository(client).reservar_rodada("untitled", "2026-10", "principal", 6)
    _, d, _ = _situacao(capsys, client)
    reserva = _item(d)["reserva"]
    assert reserva["maquina"] == "principal" and reserva["expira_em"]


def test_ignora_reserva_vencida_e_reserva_liberada(capsys):
    client = FakeClient()
    client.agora = lambda: dt.datetime(2020, 1, 1, tzinfo=dt.timezone.utc)     # reserva criada "no passado"
    repo = Repository(client)
    repo.reservar_rodada("untitled", "2026-10", "grupo-a", 6)
    _, d, _ = _situacao(capsys, client)
    assert _item(d)["reserva"] is None
    client.agora = lambda: dt.datetime.now(dt.timezone.utc)
    repo.reservar_rodada("untitled", "2026-10", "principal", 6)
    repo.liberar_rodada("untitled", "2026-10", "principal")
    _, d, _ = _situacao(capsys, client)
    assert _item(d)["reserva"] is None


def test_competencia_invalida(capsys):
    codigo = cli.main(["produtora", "situacao", "--competencia", "2026-13"], client=FakeClient())
    assert codigo == 3 and "2026-13" in capsys.readouterr().err
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_produtora_situacao.py -q`
Esperado: FALHA (`ERRO DE USO` por `situacao` inválido).

- [ ] **Passo 3: implementar em `src/papaki_dre/db.py`**

1. Logo acima de `class Repository:`, acrescente:

```python
def _no_futuro(texto, agora: dt.datetime) -> bool:
    """`texto` (ISO do PostgREST; sem fuso vale UTC) é depois de `agora`? Ilegível ou ausente = não."""
    try:
        instante = dt.datetime.fromisoformat(str(texto))
    except ValueError:
        return False
    if instante.tzinfo is None:
        instante = instante.replace(tzinfo=dt.timezone.utc)
    return instante > agora


```

2. Em `Repository`, logo antes de `def salvar_conferencia(...)`, acrescente:

```python
    def situacao(self, produtoras_ids, competencia, agora=None) -> list[dict]:
        """Uma linha por produtora, só leitura: se a competência tem foto vigente (e até que dia, e qual máquina a
        importou) e se há reserva ativa e não vencida (de qual máquina, até quando). Exige a migração 0002."""
        agora = agora or dt.datetime.now(dt.timezone.utc)
        vigentes = {r["produtora_id"]: r for r in self.client.table("importacoes")
                    .select("produtora_id,periodo_fim,maquina").eq("competencia", competencia)
                    .eq("vigente", True).execute().data}
        reservas = {r["produtora_id"]: r for r in self.client.table("reservas_rodada")
                    .select("produtora_id,maquina,expira_em").eq("competencia", competencia)
                    .eq("status", "ativa").execute().data if _no_futuro(r.get("expira_em"), agora)}
        saida = []
        for pid in produtoras_ids:
            v, r = vigentes.get(pid), reservas.get(pid)
            saida.append({
                "produtora": pid, "importada": v is not None,
                "periodo_fim": str(v["periodo_fim"]) if v else None,
                "maquina": v.get("maquina") if v else None,
                "reserva": {"maquina": r["maquina"], "expira_em": str(r["expira_em"])} if r else None})
        return saida

```

- [ ] **Passo 4: implementar em `src/papaki_dre/cli_instalacao.py`**

1. Imports: troque `import argparse` / `import json` / `import re` / `import sys` por

```python
import argparse
import datetime as dt
import json
import re
import sys
```

e a linha `from . import ambiente, caminhos, produtoras` por

```python
from . import ambiente, calendario, caminhos, produtoras
from .config import FUSO_BRASILIA
```

2. Em `registrar`, depois de `m.add_argument("--produtora", required=True)`, acrescente:

```python
    s = ps.add_parser("situacao", help="para cada produtora: competência importada, máquina e reserva (banco central)")
    s.add_argument("--competencia", help="AAAA-MM; padrão = a competência que a rotina está tratando agora")
```

3. Em `executar`, troque `return _produtora(a)` por `return _produtora(a, env, client)`.

4. Troque o início de `_produtora`:

```python
def _produtora(a, env=None, client=None) -> int:
    if a.acao == "listar":
        _json({"produtoras": produtoras.ids()})
        return EXIT_OK
    if a.acao == "situacao":
        return _situacao(a, env, client)
    p = produtoras.carregar(a.produtora)
```

(o resto de `_produtora` não muda) e, logo abaixo dela, acrescente:

```python
def _situacao(a, env, client) -> int:
    from .db import Repository, make_client
    competencia = (calendario.periodo_da_competencia(a.competencia)[0] if a.competencia
                   else calendario.competencia_alvo(dt.datetime.now(FUSO_BRASILIA).date()))
    repo = Repository(client if client is not None else make_client(env))
    _json({"competencia": competencia, "produtoras": repo.situacao(produtoras.ids(), competencia)})
    return EXIT_OK
```

- [ ] **Passo 5: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_produtora_situacao.py tests/test_cli_instalacao.py -q`
Esperado: tudo passa.

- [ ] **Passo 6: commit (só se o usuário pedir)**

```bash
git add src/papaki_dre/db.py src/papaki_dre/cli_instalacao.py tests/test_produtora_situacao.py
git commit -m "feat(produtora): comando situacao com a visão do grupo" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 9: `verificar` avisa do nome da máquina e da tabela de reservas

**Arquivos:**
- Modificar: `src/papaki_dre/verificacao.py` (`_variaveis`, `_banco`)
- Teste: `tests/test_verificacao.py`

- [ ] **Passo 1: escrever os testes que falham**

Em `tests/test_verificacao.py`, no `_env`, inclua `"PAPAKI_MAQUINA": "principal",` no dicionário padrão (depois de `"PAPAKI_USUARIO_ESFERAS": "ana@exemplo.com",`) para o `test_tudo_certo` continuar todo OK. Acrescente ao fim do arquivo:

```python
def test_sem_nome_da_maquina_avisa(tmp_path, modelo, projeto):
    env = _env(tmp_path, modelo)
    del env["PAPAKI_MAQUINA"]
    item = _por_nome(_rodar(tmp_path, modelo, projeto, env=env))["PAPAKI_MAQUINA"]
    assert item.nivel == AVISO and "--maquina" in item.detalhe and "principal" not in item.detalhe


def test_tabela_de_reservas_presente_e_ok(tmp_path, modelo, projeto):
    assert _por_nome(_rodar(tmp_path, modelo, projeto))["reservas"].nivel == OK


def test_tabela_de_reservas_ausente_e_so_aviso(tmp_path, modelo, projeto):
    class SemReservas(FakeClient):
        def table(self, name):
            if name == "reservas_rodada":
                e = Exception("relation does not exist")
                e.code = "42P01"
                raise e
            return super().table(name)

    itens = _por_nome(_rodar(tmp_path, modelo, projeto, client=SemReservas()))
    assert itens["reservas"].nivel == AVISO and "0002_reservas.sql" in itens["reservas"].detalhe
    assert itens["banco"].nivel == OK                      # as 9 tabelas da 0001 continuam o que decide o banco
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_verificacao.py -q`
Esperado: FALHA (`KeyError: 'PAPAKI_MAQUINA'` e `'reservas'`).

- [ ] **Passo 3: implementar em `src/papaki_dre/verificacao.py`**

1. Em `_variaveis`, depois do bloco de `PAPAKI_USUARIO_ESFERAS`, acrescente:

```python
    itens.append(Item(OK, "PAPAKI_MAQUINA", "definida") if _tem(env, "PAPAKI_MAQUINA")
                 else Item(AVISO, "PAPAKI_MAQUINA", "não definida: sem o nome da máquina a rotina agendada não "
                                                    "reserva a produtora (configurar --maquina)"))
```

2. Em `_banco`, troque o item final (o `itens.append(Item(OK, "banco", "responde e as 9 tabelas existem") ...`) por:

```python
    itens.append(Item(OK, "banco", "responde e as 9 tabelas existem") if not ausentes
                 else Item(FALHA, "banco", "faltam tabelas: " + ", ".join(ausentes)
                           + " (aplique supabase/migrations/0001_schema.sql no projeto)"))
    if not ausentes:
        _reservas(cliente, itens)


def _reservas(cliente, itens: list[Item]) -> None:
    """Migração 0002 (várias máquinas): sem ela o sistema funciona, mas sem a proteção da reserva."""
    try:
        cliente.table("reservas_rodada").select("*").limit(1).execute()
    except Exception as e:  # noqa: BLE001 - o erro do PostgREST/rede tem tipos variados
        if _tabela_inexistente(e):
            itens.append(Item(AVISO, "reservas", "falta a tabela reservas_rodada (aplique "
                                                 "supabase/migrations/0002_reservas.sql): sem ela duas máquinas "
                                                 "podem extrair a mesma produtora ao mesmo tempo"))
        else:
            itens.append(Item(AVISO, "reservas", f"erro ao consultar as reservas ({type(e).__name__})"))
        return
    itens.append(Item(OK, "reservas", "tabela de reservas existe"))
```

- [ ] **Passo 4: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_verificacao.py tests/test_cli_instalacao.py -q`
Esperado: tudo passa.

- [ ] **Passo 5: commit (só se o usuário pedir)**

```bash
git add src/papaki_dre/verificacao.py tests/test_verificacao.py
git commit -m "feat(verificar): avisa do nome da máquina e da tabela de reservas" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Tarefa 10: Documentação da reserva, de `PAPAKI_MAQUINA` e de `produtora situacao`

**Arquivos:**
- Modificar: `README.md`, `INSTALAR.md`, `NOVA-PRODUTORA.md`, `runbooks/rotina-mensal.md`, `runbooks/tarefa-agendada.md`
- Teste: `tests/test_docs_instalacao.py`, `tests/test_runbooks.py` (conferir)

- [ ] **Passo 1: escrever os testes que falham** (`tests/test_docs_instalacao.py`)

Em `test_comandos_citados_existem`, inclua `"reserva"` no conjunto de comandos aceitos e acrescente as regras de subcomando:

```python
            assert cmd in {"agenda", "extracao", "importar", "configurar", "verificar", "produtora",
                           "instalacao", "reserva"}, (nome, cmd)
            if cmd == "extracao":
                assert sub in ACOES_EXTRACAO, (nome, sub)
            if cmd == "produtora":
                assert sub in {"listar", "mostrar", "situacao"}, (nome, sub)
            if cmd == "reserva":
                assert sub == "obter", (nome, sub)
```

Em `test_opcoes_do_configurar_e_do_agenda_citadas_existem`, acrescente:

```python
    assert p.parse_args(["agenda", "--produtora", "untitled", "--reservar"]).reservar
    assert p.parse_args(["reserva", "obter", "--produtora", "untitled", "--competencia", "2026-10"]).acao == "obter"
    assert p.parse_args(["produtora", "situacao"]).acao == "situacao"
```

e inclua `"--maquina"` e `"0002_reservas.sql"` na tupla/verificação de `INSTALAR.md` e do `README.md`:

```python
    for opcao in ("--dados-dir", "--criar-pastas", "--modelo", "--usuario-esferas", "--sem-banco",
                  "--pasta-extracoes", "--pasta-apuracoes", "--maquina"):
        assert opcao in texto, opcao
    assert "0002_reservas.sql" in texto and "0002_reservas.sql" in _ler("README.md")
```

Acrescente um teste novo:

```python
def test_documentos_da_reserva():
    tarefa = _ler("runbooks/tarefa-agendada.md")
    assert "agenda --produtora <PRODUTORA> --restaurar-banco --reservar" in tarefa
    assert "PAPAKI_MAQUINA" in _ler("README.md") and "produtora situacao" in _ler("README.md")
    mensal = _ler("runbooks/rotina-mensal.md")
    assert "--reservar" in mensal and '"reserva"' in mensal and "reserva obter" in mensal
    assert "de_outra_maquina" in mensal and "sem_nome_da_maquina" in mensal
```

- [ ] **Passo 2: rodar e ver falhar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_docs_instalacao.py -q`
Esperado: FALHA em `test_opcoes_...` e `test_documentos_da_reserva`.

- [ ] **Passo 3: `runbooks/tarefa-agendada.md`**

(a) No bloco "Troque os campos entre `<>`", nada muda nos campos. No **Texto da tarefa**, troque o trecho

```
Primeiro rode `.venv/Scripts/python.exe -m papaki_dre agenda --produtora <PRODUTORA> --restaurar-banco`: se `rodar` for false, termine em silêncio, sem mais nada. Se for true,
```

por

```
Primeiro rode `.venv/Scripts/python.exe -m papaki_dre agenda --produtora <PRODUTORA> --restaurar-banco --reservar`: se `rodar` for false, termine em silêncio, sem mais nada. Se for true,
```

(esse trecho está dentro do bloco de código do texto da tarefa; confira o texto exato na linha do bloco antes de editar.)

(b) Na seção "Horários escalonados", troque o item **Limite:** por:

```
- **Reserva no banco:** com `--reservar`, a máquina que chega primeiro reserva a produtora e a competência no banco
  central; a outra recebe "não rodar" (`reserva: de_outra_maquina`) e termina em silêncio. A reserva vale 6 horas, é
  renovada pela mesma máquina e é liberada pelo `importar`. Sem o nome da máquina (`PAPAKI_MAQUINA`) ou sem banco, a
  rodada segue e o Claude **avisa o usuário**; nesse caso só o horário protege.
```

(c) Em "Avisos ao usuário ao criar", acrescente o item 5: `5. Cada máquina precisa do próprio nome curto (`configurar --maquina`) e do banco com a migração `0002_reservas.sql` aplicada (`verificar` avisa se faltar).`

- [ ] **Passo 4: `runbooks/rotina-mensal.md`**

(a) No bullet "Agendada", troque `agenda --produtora untitled --restaurar-banco` por `agenda --produtora untitled --restaurar-banco --reservar` e, depois da frase sobre o campo `"banco"`, acrescente:

```
  O campo `"reserva"` diz se esta máquina conseguiu reservar a rodada: `obtida` (siga), `de_outra_maquina` (o `agenda`
  já devolve `rodar: false`: termine em silêncio), `sem_nome_da_maquina` ou `nao_consultada` (siga, mas **avise o usuário**
  no resumo final: só o horário protege contra outra máquina) ou `nao_pedida`.
```

(b) No bullet "Manual", acrescente ao final:

```
  Antes de extrair, rode `.venv/Scripts/python.exe -m papaki_dre reserva obter --produtora untitled --competencia <AAAA-MM>`
  e leia `estado`: se for `de_outra_maquina`, **avise o usuário** (o JSON traz `maquina` e `expira_em`) e só continue se ele
  mandar; nunca use `--nova` por causa disso.
```

(c) Na lista "Em qualquer parada antes do `importar`", acrescente uma frase ao final do primeiro parágrafo: `A reserva da rodada **não** é liberada (ela expira em 6 h e a mesma máquina a renova ao retomar), para outra máquina não extrair do zero uma produtora cuja extração já foi disparada.`

- [ ] **Passo 5: `README.md`**

(a) Em "Banco de dados (Supabase)", troque o item 2 por:

```
2. Aplique as migrações, nesta ordem: `supabase/migrations/0001_schema.sql` e `supabase/migrations/0002_reservas.sql` (painel, SQL Editor, ou CLI do Supabase). A 0002 serve a várias máquinas: guarda as reservas de rodada e o nome da máquina que importou.
```

(b) Em "Comandos novos", depois do parágrafo do `agenda`, acrescente:

```
O `agenda` também aceita `--reservar`: se for dia de rodar e o banco responder, reserva a produtora e a competência para esta máquina (nome em `PAPAKI_MAQUINA`); se outra máquina já reservou, responde "não rodar". A reserva vale 6 horas e o `importar` a libera. **`reserva obter --produtora <id> --competencia AAAA-MM`** faz o mesmo à mão e diz quem tem a reserva (usado nas rodadas manuais).

**`produtora situacao [--competencia AAAA-MM]`** mostra, para cada produtora cadastrada, se a competência já foi importada (até que dia e por qual máquina) e se há reserva ativa.
```

(c) Na tabela de variáveis do `.env`, depois da linha de `PAPAKI_USUARIO_ESFERAS`, acrescente:

```
| `PAPAKI_MAQUINA` | Nome curto desta máquina (ex.: `principal`, `grupo-a`). Vai para as reservas de rodada e para cada importação. Gravado por `configurar --maquina`. |
```

(d) Na tabela de comandos de instalação, troque a linha de `produtora listar` por:

```
| `produtora listar` \| `produtora mostrar --produtora <id>` \| `produtora situacao` | Lista as produtoras, mostra as regras de uma (arquivo `produtoras/<id>.json`) ou mostra a situação de todas no banco central. |
```

- [ ] **Passo 6: `INSTALAR.md` e `NOVA-PRODUTORA.md`**

`INSTALAR.md`: no Passo 3, acrescente à pergunta 4 uma frase final: `Peça também o **nome curto** desta máquina (ex.: principal, grupo-a).` No comando do Passo 4, acrescente `--maquina "<nome curto>"` ao final da linha do `configurar` e, abaixo do comando, uma frase: `Banco novo ou já existente: confirme com o administrador que a migração `supabase/migrations/0002_reservas.sql` também foi aplicada (o `verificar` avisa se faltar).` No Passo 9, troque a frase **Limite conhecido** por: `**Reserva:** a tarefa usa `agenda --reservar`, então a máquina que chega primeiro reserva a produtora no banco central e a outra termina em silêncio; só falta essa proteção se o banco ou o nome da máquina não estiverem configurados (o `verificar` avisa).`

`NOVA-PRODUTORA.md`: no item 3 da seção 6, acrescente no fim: `Depois do primeiro mês, rode `.venv/Scripts/python.exe -m papaki_dre produtora situacao` para conferir a produtora entre as demais.`

- [ ] **Passo 7: rodar e ver passar**

Run: `.venv/Scripts/python.exe -m pytest tests/test_docs_instalacao.py tests/test_runbooks.py -q`
Esperado: tudo passa. Se `tests/test_runbooks.py` exigir um trecho que o novo texto removeu, restaure o trecho no documento (não afrouxe o teste).

- [ ] **Passo 8: commit (só se o usuário pedir)**

```bash
.venv/Scripts/python.exe scripts/auditar_publicacao.py
git add README.md INSTALAR.md NOVA-PRODUTORA.md runbooks/rotina-mensal.md runbooks/tarefa-agendada.md tests/test_docs_instalacao.py
git commit -m "docs: reserva de rodada, PAPAKI_MAQUINA e produtora situacao" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

## Fechamento

### Tarefa 11: Ensaio da função SQL num banco de teste (só com o "sim" do usuário no momento de executar)

O usuário já autorizou, em 22/10/2026, validar a função SQL num banco de teste. **Mesmo assim, confirme no chat qual banco será usado antes de criar ou gravar qualquer coisa. Nunca use o banco com dados reais da Papaki para este ensaio.**

**Arquivos:** nenhum (ensaio). Ferramentas: o MCP do Supabase (`apply_migration`, `execute_sql`, `create_branch`) ou o painel/SQL Editor.

- [ ] **Passo 1: escolher o banco de teste.** Opções (o usuário decide): (a) um **branch de desenvolvimento** do projeto da Papaki (tem o esquema, não os dados reais); (b) um **projeto Supabase novo e vazio**, onde se aplica primeiro `0001_schema.sql`. Registre qual foi escolhido.
- [ ] **Passo 2: aplicar `0002_reservas.sql`** no banco de teste (`apply_migration` com o conteúdo do arquivo, ou SQL Editor).
- [ ] **Passo 3: rodar os cenários** (cada `select` numa chamada separada de `execute_sql`) e conferir o resultado esperado:

```sql
select public.reservar_rodada('ensaio', '2099-01', 'maq-a', 6);   -- reservada: true, maquina: maq-a
select public.reservar_rodada('ensaio', '2099-01', 'maq-b', 6);   -- reservada: false, maquina: maq-a
select public.reservar_rodada('ensaio', '2099-01', 'maq-a', 6);   -- reservada: true (renovou: expira_em avançou)
select public.liberar_rodada('ensaio', '2099-01', 'maq-b');       -- false (não é a dona)
select public.reservar_rodada('ensaio', '2099-01', 'maq-a', 0);   -- ERRO: p_horas deve estar entre 1 e 24
update public.reservas_rodada set expira_em = now() - interval '1 minute'
  where produtora_id = 'ensaio' and status = 'ativa';
select public.reservar_rodada('ensaio', '2099-01', 'maq-b', 6);   -- reservada: true (a de maq-a expirou)
select produtora_id, maquina, status from public.reservas_rodada where produtora_id = 'ensaio' order by criada_em;
                                                                  -- maq-a expirada, maq-b ativa
select public.liberar_rodada('ensaio', '2099-01', 'maq-b');       -- true
select public.reservar_rodada('ensaio', '2099-01', 'maq-a', 6);   -- reservada: true (liberada, outra pode pegar)
insert into public.reservas_rodada (produtora_id, competencia, maquina, expira_em)
  values ('ensaio', '2099-01', 'maq-c', now() + interval '1 hour');
                                                                  -- ERRO: duplicate key ... reservas_rodada_uma_ativa
select has_function_privilege('anon', 'public.reservar_rodada(text,text,text,integer)', 'execute'),
       has_table_privilege('anon', 'public.reservas_rodada', 'select');     -- false, false
```

- [ ] **Passo 4: limpar o ensaio** no banco de teste: `delete from public.reservas_rodada where produtora_id = 'ensaio';`
- [ ] **Passo 5: registrar o resultado** (o que passou e o que não passou) em 2 ou 3 linhas para o usuário. Se algum cenário falhar, **pare**, corrija `0002_reservas.sql` e o espelho em `tests/fake_supabase.py` (`_reservar_rodada`/`_liberar_rodada`) e refaça o ensaio desde o Passo 2 num banco limpo.
- [ ] **Passo 6: apagar o branch ou projeto de teste** se foi criado só para isso (com o "sim" do usuário).

### Tarefa 12: Suíte completa, auditoria e entrada em produção

- [ ] **Passo 1: suíte rápida inteira**

Run: `.venv/Scripts/python.exe -m pytest tests -q -m "not real and not excel"`
Esperado: tudo passa (os 908+ testes anteriores e os novos). Os testes `real` e `excel` ficam de fora; se a máquina tiver o Excel e as amostras, rode também `.venv/Scripts/python.exe -m pytest tests -q -m "real or excel"` para provar que a Untitled gera o mesmo resultado.

- [ ] **Passo 2: auditoria de publicação**

Run: `.venv/Scripts/python.exe scripts/auditar_publicacao.py`
Esperado: `nada encontrado em N arquivos`.

- [ ] **Passo 3: aplicar a migração `0002_reservas.sql` no banco real da Papaki** — **só com o "sim" do usuário no chat** (é gravação no Supabase real). Depois: `.venv/Scripts/python.exe -m papaki_dre verificar` deve mostrar `OK reservas` e, na máquina que ainda não tem o nome, `AVISO PAPAKI_MAQUINA`.

- [ ] **Passo 4: cada máquina recebe o nome** (`.venv/Scripts/python.exe -m papaki_dre configurar --maquina "<nome curto>"`), o `git pull` do projeto e o novo texto da tarefa agendada (com `--reservar`), **só com o OK do usuário** para alterar tarefa agendada existente.

- [ ] **Passo 5: ensaio de ponta a ponta com o usuário** (com o "sim" dele, pois grava no Supabase real): rodar `reserva obter --produtora untitled --competencia <AAAA-MM>` em duas máquinas e ver `obtida` e `de_outra_maquina`; rodar `produtora situacao`; medir quanto tempo leva uma extração real e confirmar se 6 h de reserva e 30 minutos entre tarefas servem (decisões c e d do desenho). Anotar os valores finais em `runbooks/tarefa-agendada.md`.

- [ ] **Passo 6: fechar** — resumir ao usuário em 2 ou 3 linhas o que passou, o que não passou e o que mudou, e atualizar `docs/superpowers/HANDOFF-execucao-planos.md` se o usuário quiser manter esse registro.

---

## Autorrevisão (feita ao escrever o plano)

**Cobertura do desenho (seções 4 e 8):**
- 4.1 Parte A (escalonamento + reescrita da regra + NOVA-PRODUTORA 6.3) → Tarefa 1; reforçado na Tarefa 10.
- 4.2 reserva atômica (tabela, índice parcial, função, tempo do banco, renovação, expiração, liberação) → Tarefas 2, 3, 6, 7; identidade `PAPAKI_MAQUINA` + `importacoes.maquina` → Tarefas 2, 4, 5, 7; `agenda --reservar` e rodada manual (`reserva obter`) → Tarefa 6; parada não libera → Tarefa 10 (texto); banco indisponível segue e avisa → Tarefa 6 (`nao_consultada`) e Tarefa 10 (texto).
- 4.3 `produtora situacao` → Tarefa 8; `verificar` → Tarefa 9. **Divergência consciente do desenho:** o desenho dizia que o `verificar` passaria a conferir cada produtora que a máquina tem; isso **já existe** em `verificacao._modelos` (conferência de `PAPAKI_MODELO_<ID>` por produtora), então a Tarefa 9 só acrescenta os avisos novos.
- 8a pasta compartilhada → Tarefa 1 (INSTALAR, pergunta 1); 8b → Tarefas 6 e 10; 8c/8d → medidos na Tarefa 12, Passo 5; 8e (ensaio SQL) → Tarefa 11.

**Varredura de lacunas de texto:** os passos de documentação dizem exatamente o que trocar e qual texto entra; os trechos de código estão completos. Um ponto pede conferência do executor ao editar (sinalizado no passo): o texto exato do bloco da tarefa agendada em `runbooks/tarefa-agendada.md` (Tarefa 10, Passo 3a). Na Tarefa 5, a exigência de `"--maquina"` em `INSTALAR.md` só passa depois da Tarefa 10 (o passo avisa).

**Consistência de nomes:** `Repository.reservar_rodada(produtora, competencia, maquina, horas)`, `Repository.liberar_rodada(produtora, competencia, maquina)`, `Repository.situacao(produtoras_ids, competencia, agora=None)`, `criar_importacao(..., maquina=None)`, `cli._nome_da_maquina(env)`, `cli._reservar_rodada(produtora, competencia, *, env, client)` → dict com `estado`, `cli._liberar_reserva(repo, produtora, competencia, maquina)`, `cli_extracao._agenda(a, env, consulta_banco, reservar_banco)`, `cli_extracao._reserva(a, reservar_banco)`, `FakeClient.agora`, JSON do `agenda` com a chave `reserva` (`obtida`, `de_outra_maquina`, `sem_nome_da_maquina`, `nao_consultada`, `nao_pedida`). Funções SQL: `reservar_rodada(text, text, text, integer)` e `liberar_rodada(text, text, text)`, com os mesmos nomes de parâmetro (`p_produtora`, `p_competencia`, `p_maquina`, `p_horas`) no SQL, no cliente falso e no `Repository`.

**Riscos que o plano não elimina:** o espelho em `tests/fake_supabase.py` prova o comportamento do Python, não do Postgres (por isso a Tarefa 11); sem banco ou sem `PAPAKI_MAQUINA` a defesa volta a ser só o horário; a sincronização da pasta compartilhada pode atrasar arquivos entre máquinas (a decisão de rodar vem do banco, não da pasta).
