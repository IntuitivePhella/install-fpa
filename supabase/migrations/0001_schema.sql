-- MVP Papaki: importação das bases do Esferas. Todas as tabelas com RLS ligado e SEM políticas:
-- só a chave de serviço (service_role) acessa. Dados reais contêm CPF/CNPJ e dados bancários.

create table public.produtoras (
  id         text primary key,
  nome       text not null,
  localidade text not null
);
alter table public.produtoras enable row level security;
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
  criado_em      timestamptz not null default now(),
  -- só uma importação concluída pode ser a foto vigente
  constraint vigente_so_se_concluida check (not vigente or status = 'concluida'),
  constraint periodo_ordenado check (periodo_inicio <= periodo_fim)
);
alter table public.importacoes enable row level security;
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
alter table public.arquivos_brutos enable row level security;
create index arquivos_brutos_sha256 on public.arquivos_brutos (sha256);

create table public.rra_base (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null references public.produtoras (id),
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  num_job       text,
  dados         jsonb not null
);
alter table public.rra_base enable row level security;
create index rra_base_importacao on public.rra_base (importacao_id);

create table public.cp_despesas (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null references public.produtoras (id),
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  cod           text,
  dados         jsonb not null
);
alter table public.cp_despesas enable row level security;
create index cp_despesas_importacao on public.cp_despesas (importacao_id);

create table public.cr_receita (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null references public.produtoras (id),
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  dados         jsonb not null
);
alter table public.cr_receita enable row level security;
create index cr_receita_importacao on public.cr_receita (importacao_id);

create table public.comissoes (
  id            bigint generated always as identity primary key,
  importacao_id uuid not null references public.importacoes (id) on delete cascade,
  produtora_id  text not null references public.produtoras (id),
  competencia   text not null,
  linha_origem  integer not null,
  linha_destino integer not null,
  num_job       text,
  dados         jsonb not null
);
alter table public.comissoes enable row level security;
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
alter table public.exclusoes enable row level security;
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
alter table public.cargas_excel enable row level security;
create index cargas_excel_importacao on public.cargas_excel (importacao_id);

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
  -- duas promoções da mesma produtora e competência não correm ao mesmo tempo
  perform pg_advisory_xact_lock(hashtextextended(v_prod || '|' || v_comp, 0));
  update public.importacoes set vigente = false
    where produtora_id = v_prod and competencia = v_comp and vigente;
  update public.importacoes set vigente = true, status = 'concluida' where id = p_id;
end;
$$;
revoke all on function public.promover_importacao(uuid) from public, anon, authenticated;
grant execute on function public.promover_importacao(uuid) to service_role;

-- Defesa em profundidade: o Supabase dá privilégios de tabela/sequence a anon e authenticated por
-- padrão; com RLS sem políticas eles já não leem nada, mas aqui também ficam sem o privilégio.
revoke all on table public.produtoras, public.importacoes, public.arquivos_brutos, public.rra_base,
  public.cp_despesas, public.cr_receita, public.comissoes, public.exclusoes, public.cargas_excel
  from anon, authenticated;
revoke all on all sequences in schema public from anon, authenticated;
