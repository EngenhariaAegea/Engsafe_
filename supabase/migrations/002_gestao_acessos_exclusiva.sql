-- EngSafe | Migração para Gestão de Acessos exclusiva
-- Execute esta query agora no SQL Editor do Supabase.

begin;

alter table public.usuarios
add column if not exists gerencia_acessos boolean not null default false;

-- As duas contas são da mesma responsável e continuam administradoras.
update public.usuarios
set perfil='admin',
    status_acesso='aprovado',
    ativo=true,
    gerencia_acessos=true,
    aprovado_em=coalesce(aprovado_em,now()),
    updated_at=now()
where lower(email) in (
  'gabriella.santos@aegea.com.br',
  'engenharia.ser@ambientalserra.com.br'
);

-- Garante que nenhum outro usuário tenha a permissão exclusiva.
update public.usuarios
set gerencia_acessos=false,
    updated_at=now()
where lower(email) not in (
  'gabriella.santos@aegea.com.br',
  'engenharia.ser@ambientalserra.com.br'
);

create or replace function public.engsafe_gerencia_acessos()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists(
    select 1
    from public.usuarios u
    where u.id=auth.uid()
      and u.perfil='admin'
      and u.gerencia_acessos=true
      and u.status_acesso='aprovado'
      and u.ativo=true
  )
$$;

create or replace function public.admin_atualizar_usuario(
  p_usuario_id uuid,
  p_status text,
  p_perfil text,
  p_ativo boolean
) returns public.usuarios
language plpgsql
security definer
set search_path = ''
as $$
declare v_result public.usuarios;
begin
  if not public.engsafe_gerencia_acessos() then
    raise exception 'Gestão de acessos não autorizada para esta conta';
  end if;
  if p_status not in ('pendente','aprovado','rejeitado') then raise exception 'Status inválido'; end if;
  if p_perfil not in ('leitor','colaborador','admin') then raise exception 'Perfil inválido'; end if;
  if p_usuario_id=auth.uid() and (p_perfil<>'admin' or p_status<>'aprovado' or p_ativo=false) then
    raise exception 'A administradora não pode remover o próprio acesso';
  end if;
  update public.usuarios
  set status_acesso=p_status,
      perfil=p_perfil,
      ativo=p_ativo,
      aprovado_por=case when p_status='aprovado' then auth.uid() else aprovado_por end,
      aprovado_em=case when p_status='aprovado' then now() else aprovado_em end,
      updated_at=now()
  where id=p_usuario_id
  returning * into v_result;
  if v_result.id is null then raise exception 'Usuário não encontrado'; end if;
  insert into public.logs_auditoria(usuario_id,usuario_email,tabela,operacao,registro_id,dados_novos)
  select auth.uid(),u.email,'usuarios','ACESSO',p_usuario_id::text,to_jsonb(v_result)
  from public.usuarios u where u.id=auth.uid();
  return v_result;
end $$;

revoke all on function public.admin_atualizar_usuario(uuid,text,text,boolean) from public,anon;
grant execute on function public.admin_atualizar_usuario(uuid,text,text,boolean) to authenticated;
grant execute on function public.engsafe_gerencia_acessos() to authenticated;

drop policy if exists engsafe_admin_usuarios on public.usuarios;
create policy engsafe_admin_usuarios on public.usuarios
for select to authenticated using(public.engsafe_gerencia_acessos());

drop policy if exists engsafe_admin_dominios on public.dominios_corporativos;
create policy engsafe_admin_dominios on public.dominios_corporativos
for all to authenticated
using(public.engsafe_gerencia_acessos())
with check(public.engsafe_gerencia_acessos());

drop policy if exists engsafe_admin_auditoria on public.logs_auditoria;
create policy engsafe_admin_auditoria on public.logs_auditoria
for select to authenticated using(public.engsafe_gerencia_acessos());

commit;

-- Conferência:
select nome,email,perfil,status_acesso,ativo,gerencia_acessos
from public.usuarios
order by created_at;

