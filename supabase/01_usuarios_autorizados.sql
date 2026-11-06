-- Vívora — PASSO 1 de 2: rode ANTES de publicar a versão com login.
-- Não muda nada para o site atual: só cria a lista de autorizados e prepara as regras.
-- Antes de rodar: crie os usuários em Authentication → Users e troque os e-mails abaixo.

create table if not exists public.vivo_usuarios (
  user_id uuid primary key references auth.users (id) on delete cascade,
  criado_em timestamptz not null default now()
);

alter table public.vivo_usuarios enable row level security;
revoke all on public.vivo_usuarios from anon, authenticated;
grant select on public.vivo_usuarios to authenticated;

-- Cada usuário só enxerga a própria linha. Não há regra de escrita: a lista só muda por aqui.
drop policy if exists "vivo_usuarios: ler a propria linha" on public.vivo_usuarios;
create policy "vivo_usuarios: ler a propria linha" on public.vivo_usuarios
  for select to authenticated
  using (user_id = (select auth.uid()));

-- Regras de vivo_storage: valem de verdade só depois do passo 2, que liga o RLS na tabela.
drop policy if exists "vivo_storage: autorizados leem" on public.vivo_storage;
create policy "vivo_storage: autorizados leem" on public.vivo_storage
  for select to authenticated
  using (exists (select 1 from public.vivo_usuarios u where u.user_id = (select auth.uid())));

drop policy if exists "vivo_storage: autorizados inserem" on public.vivo_storage;
create policy "vivo_storage: autorizados inserem" on public.vivo_storage
  for insert to authenticated
  with check (exists (select 1 from public.vivo_usuarios u where u.user_id = (select auth.uid())));

drop policy if exists "vivo_storage: autorizados alteram" on public.vivo_storage;
create policy "vivo_storage: autorizados alteram" on public.vivo_storage
  for update to authenticated
  using (exists (select 1 from public.vivo_usuarios u where u.user_id = (select auth.uid())))
  with check (exists (select 1 from public.vivo_usuarios u where u.user_id = (select auth.uid())));

grant select, insert, update on public.vivo_storage to authenticated;

-- TROQUE pelos e-mails de quem pode usar o Vívora.
insert into public.vivo_usuarios (user_id)
select id from auth.users
where lower(email) in ('email1@exemplo.com', 'email2@exemplo.com')
on conflict (user_id) do nothing;

-- Confira: deve listar todos os e-mails autorizados.
select u.email from public.vivo_usuarios v join auth.users u on u.id = v.user_id order by 1;
