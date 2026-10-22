"""Acesso ao Supabase. O Repository só usa uma interface mínima do cliente
(`table().insert/select/update/eq/in_().execute()` e `rpc()`), testável com um cliente falso."""
from __future__ import annotations

import datetime as dt
import hashlib
import math
import os
import re
from urllib.parse import urlparse

from postgrest.types import ReturnMethod

from .config import DEST, PESQUISA, R_BLANK, R_FOOTER, R_H, R_JOB, R_LOC, R_MODELO, R_NUCLEO, R_PESQ
from .produtoras import Regras, regras_padrao
from .types import Formula, Sanitized, XlError

TABLES = {"rra": "rra_base", "cp": "cp_despesas", "cr": "cr_receita", "comissoes": "comissoes"}
# coluna promovida: (nome no banco, cabeçalho do relatório)
KEY_COLUMN = {"rra": ("num_job", "Núm. JOB"), "cp": ("cod", "Cód."), "comissoes": ("num_job", "Núm. JOB")}
BATCH = 500
ENV_VARS = ("SUPABASE_URL", "SUPABASE_SERVICE_KEY")
# Estados finais que `marcar` pode gravar; 'concluida' só por `promover` (função SQL atômica).
MARCAVEIS = ("falhou", "bloqueada")
MAX_MOTIVO = 500
# Detalhes do Postgres que reproduzem a linha recusada (podem conter CPF/CNPJ): corta daí até o fim.
_DETALHE_DE_LINHA = re.compile(r"(?is)\b(failing row contains|details?\s*:|key\s*\().*$")

def regras_texto(regras: Regras) -> dict[str, str]:
    """Texto da regra de cada motivo de remoção (coluna `exclusoes.regra`; spec §6), com os valores da produtora."""
    return {
        R_BLANK: "técnica: linha vazia",
        R_FOOTER: "técnica: rodapé de TOTAL reconhecido por rótulo e posição (valores não lidos)",
        R_JOB: f"CP: Nº job preenchido e Job sem '{PESQUISA}'",
        R_LOC: f"CP: Localidade diferente de '{regras.localidade_cp}'",
        R_NUCLEO: f"RRA: Núcleo diferente de '{regras.nucleo_rra}'",
        R_PESQ: f"RRA/Comissões: job contém '{PESQUISA}'",
        R_MODELO: f"Comissões: Modelo de Orçamento diferente de '{regras.modelo_comissoes}'",
        R_H: "Comissões: coluna H em " + ", ".join(regras.comissoes_excluir_h),
    }


REGRAS = regras_texto(regras_padrao())      # texto da Untitled: testes e uso avulso


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
    if isinstance(v, str):
        return v.replace("\x00", "")          # o Postgres recusa NUL em text/jsonb
    if isinstance(v, int):
        return v
    return str(v).replace("\x00", "")


def row_to_json(header, values) -> dict:
    out, seen = {}, {}
    for name, v in zip(header, values):
        seen[name] = seen.get(name, 0) + 1
        key = name if seen[name] == 1 else f"{name} ({seen[name]})"
        key = str(key).replace("\x00", "")
        out[key] = json_value(v)
    return out


def _secrets() -> list[str]:
    """Valores que nunca podem aparecer em mensagens: URL, host/ref do projeto, chave de serviço e
    token pessoal (API de gerenciamento)."""
    out = []
    url = os.environ.get("SUPABASE_URL") or ""
    if url:
        out.append(url)
        host = urlparse(url).hostname or ""
        if host:
            out.append(host)
            out.append(host.split(".")[0])          # ref do projeto
    for nome in ("SUPABASE_SERVICE_KEY", "SUPABASE_ACCESS_TOKEN"):
        valor = os.environ.get(nome) or ""
        if valor:
            out.append(valor)
    return sorted((s for s in out if s), key=len, reverse=True)


def _scrub(msg: str) -> str:
    for s in _secrets():
        msg = msg.replace(s, "***")
    return msg


def _seguro(texto: str) -> str:
    """Sem URL/host/chave, sem detalhes de linha recusada, no máximo MAX_MOTIVO caracteres."""
    texto = _DETALHE_DE_LINHA.sub("", _scrub(str(texto))).strip()
    return texto if len(texto) <= MAX_MOTIVO else texto[:MAX_MOTIVO - 1] + "…"


def mensagem_segura(exc: BaseException) -> str:
    """Texto de uma exceção que pode ir para `motivo_falha` e para o stderr: tipo + código + mensagem.
    Nunca leva `details`/`hint` do PostgREST (a linha recusada pode conter CPF/CNPJ), nem URL, host
    ou chave do Supabase. No máximo MAX_MOTIVO caracteres."""
    nome = type(exc).__name__
    message = getattr(exc, "message", None)
    if isinstance(message, str) and message:
        code = getattr(exc, "code", None)
        texto = f"{nome}: {code} {message}" if code else f"{nome}: {message}"
    else:
        texto = f"{nome}: {exc}"
    return _seguro(texto)


def make_client(env=None):
    """Cria o cliente a partir de SUPABASE_URL e SUPABASE_SERVICE_KEY lidos por `ambiente` (o MESMO
    lugar de onde `importar --restaurar-banco` lê o token e o ref: o .env da raiz do projeto, que vale
    sobre as variáveis do Windows; valor vazio conta como ausente). `env` dado = testes.
    Os valores usados são copiados para os.environ para `mensagem_segura` continuar escondendo-os.
    Mensagens de erro dizem só QUAL variável falta, nunca valores."""
    import supabase

    from . import ambiente

    e = ambiente.resolver(env)
    missing = [v for v in ENV_VARS if not (e.get(v) or "").strip()]
    if missing:
        raise RuntimeError(f"defina {' e '.join(missing)} (veja .env.example)")
    for nome in (*ENV_VARS, "SUPABASE_ACCESS_TOKEN"):
        if (e.get(nome) or "").strip():
            os.environ[nome] = e[nome]
    return supabase.create_client(e["SUPABASE_URL"], e["SUPABASE_SERVICE_KEY"])


class BancoIndisponivel(RuntimeError):
    """O Supabase não respondeu (projeto pausado no plano gratuito, rede, URL errada)."""


def _no_futuro(texto, agora: dt.datetime) -> bool:
    """`texto` (ISO do PostgREST; sem fuso vale UTC) é depois de `agora`? Ilegível ou ausente = não."""
    try:
        instante = dt.datetime.fromisoformat(str(texto))
    except ValueError:
        return False
    if instante.tzinfo is None:
        instante = instante.replace(tzinfo=dt.timezone.utc)
    return instante > agora


class Repository:
    def __init__(self, client):
        self.client = client

    def verificar_acesso(self) -> None:
        """Consulta leve antes de qualquer gravação. Qualquer falha vira BancoIndisponivel:
        no plano gratuito, a causa mais comum é o projeto pausado após ~7 dias sem uso.
        Não restaura o projeto (isso é do Plano 4). A mensagem não expõe URL nem chave."""
        motivo = None
        try:
            self.client.table("produtoras").select("id").execute()
        except Exception as e:  # noqa: BLE001 - o cliente levanta tipos variados (httpx, postgrest)
            motivo = mensagem_segura(e)
        if motivo is not None:          # fora do except: o __context__ não guarda a exceção com URL/headers
            raise BancoIndisponivel(motivo)

    def _insert(self, table: str, rows: list[dict]) -> None:
        """Insere em lotes de BATCH linhas (um round-trip por lote, nunca por linha)."""
        for i in range(0, len(rows), BATCH):
            self.client.table(table).insert(rows[i:i + BATCH], returning=ReturnMethod.minimal).execute()

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

    def registrar_arquivo(self, importacao_id, relatorio, *, nome, sha256, tamanho,
                          total_tela=None, filtros_tela=None, prova_caminho=None) -> None:
        self.client.table("arquivos_brutos").insert({
            "importacao_id": importacao_id, "relatorio": relatorio, "nome_arquivo": nome,
            "sha256": sha256, "tamanho_bytes": tamanho, "total_tela": total_tela,
            "filtros_tela": filtros_tela, "prova_caminho": prova_caminho},
            returning=ReturnMethod.minimal).execute()

    def inserir_linhas(self, importacao_id, produtora, competencia, relatorio, s: Sanitized) -> None:
        first = DEST[relatorio]["first_row"]
        key = KEY_COLUMN.get(relatorio)
        key_idx = s.header.index(key[1]) if key else None
        rows = []
        for i, (values, origem) in enumerate(zip(s.db_rows, s.source_rows, strict=True)):
            row = {"importacao_id": importacao_id, "produtora_id": produtora, "competencia": competencia,
                   "linha_origem": origem, "linha_destino": first + i,
                   "dados": row_to_json(s.header, values)}
            if key:
                v = json_value(values[key_idx])
                row[key[0]] = None if v is None else str(v)
            rows.append(row)
        self._insert(TABLES[relatorio], rows)

    def inserir_exclusoes(self, importacao_id, relatorio, s: Sanitized, regras_txt: dict[str, str] | None = None) -> None:
        """Grava as duas listas: remoções técnicas (tipo 'tecnica') e exclusões de negócio ('negocio').
        `regras_txt` = `regras_texto(...)` da produtora (padrão: a da Untitled)."""
        textos = regras_txt or REGRAS
        rows = [{"importacao_id": importacao_id, "relatorio": relatorio, "linha_origem": r.source_row,
                 "tipo": tipo, "motivo": r.reason, "regra": textos.get(r.reason, r.reason)}
                for tipo, removals in (("tecnica", s.technical), ("negocio", s.exclusions))
                for r in removals]
        self._insert("exclusoes", rows)

    def registrar_carga(self, importacao_id, *, aba, primeira, ultima, arquivo, sha256) -> None:
        self.client.table("cargas_excel").insert({
            "importacao_id": importacao_id, "aba": aba, "linha_inicial": primeira, "linha_final": ultima,
            "linhas": max(ultima - primeira + 1, 0), "arquivo_saida": arquivo, "sha256_saida": sha256},
            returning=ReturnMethod.minimal).execute()

    def find_identical(self, produtora, competencia, hashes: dict[str, str]) -> str | None:
        """Id da importação idêntica (ver `identica`), ou None."""
        achada = self.identica(produtora, competencia, hashes)
        return achada["id"] if achada else None

    def vigente(self, produtora, competencia) -> str | None:
        """Id da foto vigente da produtora/competência, ou None."""
        rows = (self.client.table("importacoes").select("id").eq("produtora_id", produtora)
                .eq("competencia", competencia).eq("vigente", True).execute().data)
        return rows[0]["id"] if rows else None

    def vigente_completa(self, produtora, competencia, periodo_fim) -> str | None:
        """Id da foto vigente cujo período vai até `periodo_fim` (AAAA-MM-DD) ou além; None se não há.
        É o que o `agenda` usa para saber se a competência já foi importada por QUALQUER máquina."""
        rows = (self.client.table("importacoes").select("id,periodo_fim").eq("produtora_id", produtora)
                .eq("competencia", competencia).eq("vigente", True).execute().data)
        return rows[0]["id"] if rows and str(rows[0]["periodo_fim"]) >= periodo_fim else None

    def garantir_produtora(self, produtora_id, nome, localidade) -> bool:
        """Garante a linha da produtora em `produtoras` (as outras tabelas apontam para ela). True se criou."""
        achou = self.client.table("produtoras").select("id").eq("id", produtora_id).execute().data
        if achou:
            return False
        self.client.table("produtoras").insert(
            {"id": produtora_id, "nome": nome, "localidade": localidade}, returning=ReturnMethod.minimal).execute()
        return True

    def identica(self, produtora, competencia, hashes: dict[str, str]) -> dict | None:
        """{"id", "vigente"} de uma importação concluída (vigente ou não; spec §4: mesmo hash = ignorado) da
        produtora e competência cujos arquivos são EXATAMENTE `hashes`: o mesmo conjunto de relatórios, cada
        um com o mesmo hash. Prefere a vigente; `vigente` diz se a achada é a foto vigente.

        Parte das importações concluídas da produtora/competência (poucas) e só então lê os arquivos
        delas (4 por importação), então nenhuma consulta esbarra no limite de 1000 linhas do PostgREST,
        por mais rodadas falhadas/bloqueadas com os mesmos arquivos que existam."""
        if not hashes:
            return None
        imps = (self.client.table("importacoes").select("id,vigente")
                .eq("produtora_id", produtora).eq("competencia", competencia)
                .eq("status", "concluida").execute().data)
        if not imps:
            return None
        found = (self.client.table("arquivos_brutos").select("importacao_id,relatorio,sha256")
                 .in_("importacao_id", [i["id"] for i in imps]).execute().data)
        by_imp: dict[str, dict[str, str]] = {}
        for r in found:
            by_imp.setdefault(r["importacao_id"], {})[r["relatorio"]] = r["sha256"]
        matches = [i for i in imps if by_imp.get(i["id"]) == hashes]
        if not matches:
            return None
        achada = next((i for i in matches if i.get("vigente")), matches[0])
        return {"id": achada["id"], "vigente": bool(achada.get("vigente"))}

    def promover(self, importacao_id) -> None:
        """Torna a importação a foto vigente (função SQL atômica; só para 'em_andamento')."""
        self.client.rpc("promover_importacao", {"p_id": importacao_id}).execute()

    def marcar(self, importacao_id, status, *, motivo=None, conferencia=None) -> bool:
        """Grava 'falhou' ou 'bloqueada'. Só altera importações ainda 'em_andamento': uma exceção
        depois de `promover()` nunca desfaz uma foto vigente. Devolve se alguma linha mudou.
        O `motivo` é limpo (sem URL/chave/linha recusada) e truncado; para exceções, passe
        `mensagem_segura(e)`."""
        if status not in MARCAVEIS:
            raise ValueError(f"status inválido para marcar: {status!r} (use {', '.join(MARCAVEIS)})")
        values = {"status": status, "vigente": False}
        if motivo is not None:
            values["motivo_falha"] = _seguro(motivo)
        if conferencia is not None:
            values["conferencia"] = conferencia
        res = self.client.table("importacoes").update(values) \
            .eq("id", importacao_id).eq("status", "em_andamento").execute()
        return bool(res.data)

    def encerrar_orfas(self, produtora, competencia, *, horas=6, agora=None) -> int:
        """Marca 'falhou' (motivo 'rodada interrompida') as importações 'em_andamento' da MESMA produtora e
        competência criadas há mais de `horas` (o processo morreu no meio). Só marca: nunca apaga linhas,
        nunca toca em outra produtora/competência, em rodada recente, 'concluida' ou vigente (`marcar` só
        altera 'em_andamento'). Devolve quantas foram marcadas.

        Filtra produtora/competência/status no servidor e compara a data aqui (`criado_em` é timestamptz:
        o PostgREST devolve ISO com fuso; sem fuso vale UTC; ilegível ou ausente = não encerra). Limite
        conhecido: uma rodada legítima com mais de `horas` horas seria marcada 'falhou' por engano; as
        rodadas reais levam minutos, então 6 h é folga larga."""
        agora = agora or dt.datetime.now(dt.timezone.utc)
        if agora.tzinfo is None:
            agora = agora.replace(tzinfo=dt.timezone.utc)
        linhas = (self.client.table("importacoes").select("id,criado_em")
                  .eq("produtora_id", produtora).eq("competencia", competencia)
                  .eq("status", "em_andamento").execute().data)
        limite = dt.timedelta(hours=horas)
        encerradas = 0
        for r in linhas:
            try:
                criado = dt.datetime.fromisoformat(str(r.get("criado_em")))
            except ValueError:
                continue
            if criado.tzinfo is None:
                criado = criado.replace(tzinfo=dt.timezone.utc)
            if agora - criado > limite and self.marcar(
                    r["id"], "falhou", motivo=f"rodada interrompida (sem desfecho há mais de {horas} h)"):
                encerradas += 1
        return encerradas

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

    def salvar_conferencia(self, importacao_id, conferencia) -> None:
        self.client.table("importacoes").update({"conferencia": conferencia}).eq("id", importacao_id).execute()
