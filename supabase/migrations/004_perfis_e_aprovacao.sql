-- EngSafe | Perfis Administrador, Colaborador e Leitor
-- Execute após as migrations 001, 002 e 003 em projetos já existentes.

begin;

-- Remove a validação antiga antes de converter Gestor para Colaborador.
do $$
declare v_constraint text;
begin
  for v_constraint in
    select c.conname
    from pg_constraint c
    where c.conrelid='public.usuarios'::regclass
      and c.contype='c'
      and pg_get_constraintdef(c.oid) ilike '%perfil%'
  loop
    execute format('alter table public.usuarios drop constraint %I',v_constraint);
  end loop;
end $$;

update public.usuarios
set perfil='colaborador',updated_at=now()
where perfil='gestor';

alter table public.usuarios
  add constraint usuarios_perfil_check
  check (perfil in ('leitor','colaborador','admin'));

create or replace function public.engsafe_pode_gerir()
returns boolean
language sql
stable
security definer
set search_path=''
as $$
  select coalesce(public.engsafe_perfil() in ('colaborador','admin'),false)
$$;

create or replace function public.admin_atualizar_usuario(
  p_usuario_id uuid,
  p_status text,
  p_perfil text,
  p_ativo boolean
) returns public.usuarios
language plpgsql
security definer
set search_path=''
as $$
declare v_result public.usuarios;
begin
  if not public.engsafe_gerencia_acessos() then
    raise exception 'Gestão de acessos não autorizada para esta conta';
  end if;
  if p_status not in ('pendente','aprovado','rejeitado') then
    raise exception 'Status inválido';
  end if;
  if p_perfil not in ('leitor','colaborador','admin') then
    raise exception 'Perfil inválido';
  end if;
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

-- Mantém a Gestão de Acessos somente nas duas contas da responsável.
update public.usuarios
set perfil='admin',status_acesso='aprovado',ativo=true,gerencia_acessos=true,
    aprovado_em=coalesce(aprovado_em,now()),updated_at=now()
where lower(email) in (
  'gabriella.santos@aegea.com.br',
  'engenharia.ser@ambientalserra.com.br'
);

update public.usuarios
set gerencia_acessos=false,updated_at=now()
where lower(email) not in (
  'gabriella.santos@aegea.com.br',
  'engenharia.ser@ambientalserra.com.br'
);

-- Atualiza a permissão do histórico para o novo perfil Colaborador.
drop policy if exists engsafe_historico_insert_manual on public.historico_emails;
create policy engsafe_historico_insert_manual on public.historico_emails
for insert to authenticated
with check (
  public.engsafe_perfil() in ('colaborador','admin')
  and status_envio='enviado_manual'
  and enviado_por=lower(coalesce(auth.jwt()->>'email',''))
);

revoke all on function public.admin_atualizar_usuario(uuid,text,text,boolean) from public,anon;
grant execute on function public.admin_atualizar_usuario(uuid,text,text,boolean) to authenticated;
grant execute on function public.engsafe_pode_gerir() to authenticated;
grant select,insert on public.historico_emails to authenticated;

commit;

-- Conferência:
select nome,email,perfil,status_acesso,ativo,gerencia_acessos
from public.usuarios
order by created_at desc;
