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
