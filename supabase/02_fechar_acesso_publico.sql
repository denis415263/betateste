-- Vívora — PASSO 2 de 2: rode só DEPOIS de testar o login no site publicado.
-- A partir daqui a chave pública não lê nem grava mais. Faça um backup antes
-- (Relatórios → Limpeza de itens → Exportar completo).

-- Remove qualquer regra antiga em vivo_storage, menos as criadas no passo 1.
do $$
declare p record;
begin
  for p in
    select policyname from pg_policies
    where schemaname = 'public' and tablename = 'vivo_storage'
      and policyname not like 'vivo_storage: autorizados %'
  loop
    execute format('drop policy %I on public.vivo_storage', p.policyname);
  end loop;
end $$;

alter table public.vivo_storage enable row level security;
revoke all on public.vivo_storage from anon;

-- Confira: vivo_storage deve estar com rowsecurity = true.
-- Se aparecer outra tabela com rowsecurity = false, ela continua exposta: me avise.
select tablename, rowsecurity from pg_tables where schemaname = 'public' order by rowsecurity, tablename;
