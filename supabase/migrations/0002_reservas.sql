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
