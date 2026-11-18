-- EngSafe | Autenticação, aprovação de acesso, perfis, RLS e auditoria
-- Execute no SQL Editor do Supabase com um usuário proprietário do projeto.
-- ANTES DO PRIMEIRO CADASTRO, troque os dois valores SEU_* no final deste arquivo.

begin;

create extension if not exists pgcrypto with schema extensions;

create table if not exists public.dominios_corporativos (
  id uuid primary key default gen_random_uuid(),
  dominio text not null,
  ativo boolean not null default true,
  created_at timestamptz not null default now(),
  created_by uuid references auth.users(id)
);

create table if not exists public.usuarios (
  id uuid primary key references auth.users(id) on delete cascade,
  nome text not null,
  email text not null,
  perfil text not null default 'leitor' check (perfil in ('leitor','colaborador','admin')),
  gerencia_acessos boolean not null default false,
  status_acesso text not null default 'pendente' check (status_acesso in ('pendente','aprovado','rejeitado')),
  ativo boolean not null default true,
  aprovado_por uuid references auth.users(id),
  aprovado_em timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.usuarios add column if not exists gerencia_acessos boolean not null default false;

create table if not exists public.logs_auditoria (
  id bigint generated always as identity primary key,
  usuario_id uuid references auth.users(id) on delete set null,
  usuario_email text,
  tabela text not null,
  operacao text not null check (operacao in ('INSERT','UPDATE','DELETE','ACESSO')),
  registro_id text,
  dados_anteriores jsonb,
  dados_novos jsonb,
  criado_em timestamptz not null default now()
);

create index if not exists usuarios_status_idx on public.usuarios(status_acesso,ativo);
create index if not exists usuarios_perfil_idx on public.usuarios(perfil);
create unique index if not exists dominios_corporativos_lower_uidx on public.dominios_corporativos(lower(dominio));
create unique index if not exists usuarios_email_lower_uidx on public.usuarios(lower(email));
create index if not exists logs_auditoria_usuario_idx on public.logs_auditoria(usuario_id,criado_em desc);
create index if not exists logs_auditoria_tabela_idx on public.logs_auditoria(tabela,criado_em desc);

create or replace function public.engsafe_perfil()
returns text language sql stable security definer set search_path = ''
as $$ select u.perfil from public.usuarios u where u.id = auth.uid() and u.status_acesso='aprovado' and u.ativo=true $$;

create or replace function public.engsafe_tem_acesso()
returns boolean language sql stable security definer set search_path = ''
as $$ select exists(select 1 from public.usuarios u where u.id=auth.uid() and u.status_acesso='aprovado' and u.ativo=true) $$;

create or replace function public.engsafe_pode_gerir()
returns boolean language sql stable security definer set search_path = ''
as $$ select coalesce(public.engsafe_perfil() in ('colaborador','admin'),false) $$;

create or replace function public.engsafe_e_admin()
returns boolean language sql stable security definer set search_path = ''
as $$ select coalesce(public.engsafe_perfil()='admin',false) $$;

create or replace function public.engsafe_gerencia_acessos()
returns boolean language sql stable security definer set search_path = ''
as $$
  select exists(
    select 1 from public.usuarios u
    where u.id=auth.uid() and u.perfil='admin' and u.gerencia_acessos=true
      and u.status_acesso='aprovado' and u.ativo=true
  )
$$;

create or replace function public.engsafe_novo_usuario()
returns trigger language plpgsql security definer set search_path = ''
as $$
declare v_dominio text;
begin
  v_dominio := lower(split_part(new.email,'@',2));
  if not exists(select 1 from public.dominios_corporativos d where lower(d.dominio::text)=v_dominio and d.ativo) then
    raise exception 'Utilize um e-mail corporativo autorizado';
  end if;
  insert into public.usuarios(id,nome,email,perfil,status_acesso,ativo)
  values(new.id,coalesce(nullif(trim(new.raw_user_meta_data->>'nome'),''),split_part(new.email,'@',1)),new.email,'leitor','pendente',true)
  on conflict(id) do nothing;
  return new;
end $$;

drop trigger if exists engsafe_auth_novo_usuario on auth.users;
create trigger engsafe_auth_novo_usuario after insert on auth.users
for each row execute function public.engsafe_novo_usuario();

-- Inclui contas Auth que já existiam antes desta instalação.
insert into public.usuarios(id,nome,email)
select a.id,coalesce(nullif(trim(a.raw_user_meta_data->>'nome'),''),split_part(a.email,'@',1)),a.email
from auth.users a where a.email is not null on conflict(id) do nothing;

create or replace function public.admin_atualizar_usuario(
  p_usuario_id uuid, p_status text, p_perfil text, p_ativo boolean
) returns public.usuarios language plpgsql security definer set search_path = ''
as $$
declare v_result public.usuarios;
begin
  if not public.engsafe_gerencia_acessos() then raise exception 'Gestão de acessos não autorizada para esta conta'; end if;
  if p_status not in ('pendente','aprovado','rejeitado') then raise exception 'Status inválido'; end if;
  if p_perfil not in ('leitor','colaborador','admin') then raise exception 'Perfil inválido'; end if;
  if p_usuario_id=auth.uid() and (p_perfil<>'admin' or p_status<>'aprovado' or p_ativo=false) then
    raise exception 'O administrador não pode remover o próprio acesso';
  end if;
  update public.usuarios set
    status_acesso=p_status, perfil=p_perfil, ativo=p_ativo,
    aprovado_por=case when p_status='aprovado' then auth.uid() else aprovado_por end,
    aprovado_em=case when p_status='aprovado' then now() else aprovado_em end,
    updated_at=now()
  where id=p_usuario_id returning * into v_result;
  if v_result.id is null then raise exception 'Usuário não encontrado'; end if;
  insert into public.logs_auditoria(usuario_id,usuario_email,tabela,operacao,registro_id,dados_novos)
  select auth.uid(),u.email,'usuarios','ACESSO',p_usuario_id::text,to_jsonb(v_result) from public.usuarios u where u.id=auth.uid();
  return v_result;
end $$;

revoke all on function public.admin_atualizar_usuario(uuid,text,text,boolean) from public,anon;
grant execute on function public.admin_atualizar_usuario(uuid,text,text,boolean) to authenticated;
grant execute on function public.engsafe_perfil(),public.engsafe_tem_acesso(),public.engsafe_pode_gerir(),public.engsafe_e_admin(),public.engsafe_gerencia_acessos() to authenticated;

alter table public.dominios_corporativos enable row level security;
alter table public.usuarios enable row level security;
alter table public.logs_auditoria enable row level security;

drop policy if exists engensafe_usuario_proprio on public.usuarios;
drop policy if exists engsafe_usuario_proprio on public.usuarios;
drop policy if exists engsafe_admin_usuarios on public.usuarios;
create policy engsafe_usuario_proprio on public.usuarios for select to authenticated using(id=auth.uid());
create policy engsafe_admin_usuarios on public.usuarios for select to authenticated using(public.engsafe_gerencia_acessos());

drop policy if exists engsafe_admin_dominios on public.dominios_corporativos;
create policy engsafe_admin_dominios on public.dominios_corporativos for all to authenticated using(public.engsafe_gerencia_acessos()) with check(public.engsafe_gerencia_acessos());
drop policy if exists engsafe_admin_auditoria on public.logs_auditoria;
create policy engsafe_admin_auditoria on public.logs_auditoria for select to authenticated using(public.engsafe_gerencia_acessos());

revoke all on public.usuarios,public.dominios_corporativos,public.logs_auditoria from anon;
grant select on public.usuarios to authenticated;
grant select,insert,update,delete on public.dominios_corporativos to authenticated;
grant select on public.logs_auditoria to authenticated;

-- Políticas para as tabelas operacionais já criadas no EngSafe.
do $$
declare t text; p record;
begin
  foreach t in array array['ocorrencias','atividades_sst','contratadas','autoridades_contratada'] loop
    if to_regclass('public.'||t) is not null then
      execute format('alter table public.%I enable row level security',t);
      -- Remove políticas antigas permissivas (inclusive as queries iniciais que liberavam anon).
      for p in select policyname from pg_policies where schemaname='public' and tablename=t loop
        execute format('drop policy if exists %I on public.%I',p.policyname,t);
      end loop;
      execute format('create policy engsafe_ler on public.%I for select to authenticated using (public.engsafe_tem_acesso())',t);
      execute format('create policy engsafe_inserir on public.%I for insert to authenticated with check (public.engsafe_pode_gerir())',t);
      execute format('create policy engsafe_editar on public.%I for update to authenticated using (public.engsafe_pode_gerir()) with check (public.engsafe_pode_gerir())',t);
      execute format('create policy engsafe_excluir on public.%I for delete to authenticated using (public.engsafe_pode_gerir())',t);
      execute format('revoke all on public.%I from anon',t);
      execute format('grant select,insert,update,delete on public.%I to authenticated',t);
    end if;
  end loop;
end $$;

-- O histórico é lido por usuários aprovados. A migration 003 libera o registro manual para Colaborador/Admin.
alter table if exists public.historico_emails enable row level security;
alter table if exists public.historico_emails alter column ocorrencia_id drop not null;
alter table if exists public.historico_emails alter column contratada_id drop not null;
alter table if exists public.historico_emails add column if not exists tipo_envio text not null default 'notificacao';
alter table if exists public.historico_emails add column if not exists filtros text[];
do $$ begin
  if to_regclass('public.historico_emails') is not null then
    drop policy if exists engsafe_historico_select on public.historico_emails;
    create policy engsafe_historico_select on public.historico_emails for select to authenticated using(public.engsafe_tem_acesso());
    revoke all on public.historico_emails from anon,authenticated;
    grant select on public.historico_emails to authenticated;
  end if;
end $$;

-- Storage: leitor baixa/visualiza; colaborador/admin envia e remove evidências.
update storage.buckets set public=false where id in ('evidencias','atividades');
drop policy if exists engsafe_storage_ler on storage.objects;
drop policy if exists engsafe_storage_inserir on storage.objects;
drop policy if exists engsafe_storage_editar on storage.objects;
drop policy if exists engsafe_storage_excluir on storage.objects;
create policy engsafe_storage_ler on storage.objects for select to authenticated
using(bucket_id in ('evidencias','atividades') and public.engsafe_tem_acesso());
create policy engsafe_storage_inserir on storage.objects for insert to authenticated
with check(bucket_id in ('evidencias','atividades') and public.engsafe_pode_gerir());
create policy engsafe_storage_editar on storage.objects for update to authenticated
using(bucket_id in ('evidencias','atividades') and public.engsafe_pode_gerir())
with check(bucket_id in ('evidencias','atividades') and public.engsafe_pode_gerir());
create policy engsafe_storage_excluir on storage.objects for delete to authenticated
using(bucket_id in ('evidencias','atividades') and public.engsafe_pode_gerir());

create or replace function public.engsafe_auditar_alteracao()
returns trigger language plpgsql security definer set search_path = ''
as $$
declare v_old jsonb; v_new jsonb; v_id text; v_email text;
begin
  v_old:=case when tg_op in ('UPDATE','DELETE') then to_jsonb(old) else null end;
  v_new:=case when tg_op in ('INSERT','UPDATE') then to_jsonb(new) else null end;
  v_id:=coalesce(v_new->>'id',v_old->>'id');
  select u.email into v_email from public.usuarios u where u.id=auth.uid();
  insert into public.logs_auditoria(usuario_id,usuario_email,tabela,operacao,registro_id,dados_anteriores,dados_novos)
  values(auth.uid(),v_email,tg_table_name,tg_op,v_id,v_old,v_new);
  return case when tg_op='DELETE' then old else new end;
end $$;

do $$
declare t text;
begin
  foreach t in array array['ocorrencias','atividades_sst','contratadas','autoridades_contratada'] loop
    if to_regclass('public.'||t) is not null then
      execute format('drop trigger if exists engsafe_auditoria on public.%I',t);
      execute format('create trigger engsafe_auditoria after insert or update or delete on public.%I for each row execute function public.engsafe_auditar_alteracao()',t);
    end if;
  end loop;
end $$;

commit;

-- =============================================================
-- PERSONALIZE E EXECUTE ESTAS DUAS INSTRUÇÕES APÓS O BLOCO ACIMA
-- =============================================================
-- 1. Autorize o domínio corporativo antes de novos cadastros:
-- insert into public.dominios_corporativos(dominio) values ('SEU-DOMINIO.com.br') on conflict do nothing;

-- 2. Depois que você solicitar acesso pelo HTML, torne sua conta a primeira administradora:
-- update public.usuarios set perfil='admin',status_acesso='aprovado',ativo=true,aprovado_em=now()
-- where email='SEU_EMAIL_CORPORATIVO';

-- 3. Permita a Gestão de Acessos somente nas duas contas pessoais da responsável:
-- update public.usuarios set gerencia_acessos=true
-- where lower(email) in ('gabriella.santos@aegea.com.br','engenharia.ser@ambientalserra.com.br');
