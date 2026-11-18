-- EngSafe | Estrutura idempotente para notificações e histórico de e-mails
begin;

create table if not exists public.historico_emails (
  id uuid primary key default gen_random_uuid(),
  ocorrencia_id uuid references public.ocorrencias(id) on delete set null,
  contratada_id uuid references public.contratadas(id) on delete set null,
  tipo_envio text not null default 'notificacao' check (tipo_envio in ('notificacao','relatorio')),
  destinatarios text[] not null,
  assunto text not null,
  status_envio text not null check (status_envio in ('enviado','falha','enviado_manual')),
  provedor_id text,
  erro text,
  enviado_por text not null,
  filtros text[],
  data_envio timestamptz not null default now()
);

alter table public.historico_emails alter column ocorrencia_id drop not null;
alter table public.historico_emails alter column contratada_id drop not null;
alter table public.historico_emails add column if not exists tipo_envio text not null default 'notificacao';
alter table public.historico_emails add column if not exists provedor_id text;
alter table public.historico_emails add column if not exists erro text;
alter table public.historico_emails add column if not exists filtros text[];

-- Compatibilidade para bancos que já criaram a restrição anterior.
alter table public.historico_emails drop constraint if exists historico_emails_status_envio_check;
alter table public.historico_emails
  add constraint historico_emails_status_envio_check
  check (status_envio in ('enviado','falha','enviado_manual'));

alter table public.ocorrencias
  add column if not exists contratada_id uuid references public.contratadas(id) on delete set null;

create index if not exists historico_emails_ocorrencia_idx on public.historico_emails(ocorrencia_id,data_envio desc);
create index if not exists historico_emails_contratada_idx on public.historico_emails(contratada_id,data_envio desc);
create index if not exists ocorrencias_contratada_idx on public.ocorrencias(contratada_id);
create index if not exists autoridades_contratada_email_idx on public.autoridades_contratada(contratada_id,ativo,email);

alter table public.historico_emails enable row level security;
drop policy if exists engsafe_historico_select on public.historico_emails;
create policy engsafe_historico_select on public.historico_emails
for select to authenticated using (public.engsafe_tem_acesso());

drop policy if exists engsafe_historico_insert_manual on public.historico_emails;
create policy engsafe_historico_insert_manual on public.historico_emails
for insert to authenticated
with check (
  public.engsafe_perfil() in ('colaborador','admin')
  and status_envio = 'enviado_manual'
  and enviado_por = lower(coalesce(auth.jwt() ->> 'email',''))
);

revoke all on public.historico_emails from anon,authenticated;
grant select,insert on public.historico_emails to authenticated;

commit;
