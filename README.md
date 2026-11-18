# EngSafe — Gestão de Segurança do Trabalho

Projeto estático preparado para VS Code, GitHub e Cloudflare Pages, com autenticação e dados no Supabase.

O envio de notificações é manual: o EngSafe abre uma nova mensagem no Outlook Web com destinatários, assunto e texto preenchidos. A pessoa responsável confirma a conta `engenharia.ser@ambientalserra.com.br`, revisa e clica em **Enviar**. Não há Resend, API Key de e-mail, Worker de e-mail ou alteração de DNS.

## Estrutura

```text
frontend/index.html       estrutura das telas
frontend/config.js        marcadores da URL/chave publishable do Supabase (preenchidos no build)
frontend/assets/styles.css estilos, cores e responsividade
frontend/assets/app.js    regras do sistema, Supabase, PDF e Outlook
frontend/assets/pdf-template.js imagem institucional usada no PDF
frontend/assets/images/   logos e elementos gráficos
supabase/migrations/      SQL do banco, RLS e histórico
scripts/build-pages.mjs   build do Cloudflare Pages
.github/workflows/        validação e deploy automático
```

Para atualizar o sistema, edite somente o arquivo correspondente à mudança. Preserve os caminhos relativos e publique a pasta completa; o `index.html` isolado não contém mais estilos, lógica ou imagens.

## Supabase

Execute no SQL Editor, nesta ordem:

1. `supabase/migrations/001_auth_rls_auditoria.sql`
2. `supabase/migrations/002_gestao_acessos_exclusiva.sql`
3. `supabase/migrations/003_email_integracao.sql`
4. `supabase/migrations/004_perfis_e_aprovacao.sql`

O terceiro script permite registrar no histórico uma cobrança enviada manualmente pelo Outlook. O quarto ajusta instalações existentes para os perfis Administrador, Colaborador e Leitor.

## Perfis e aprovação

- **Administrador:** acesso operacional total. A Gestão de Acessos permanece exclusiva para `gabriella.santos@aegea.com.br` e `engenharia.ser@ambientalserra.com.br`.
- **Colaborador:** registra e exclui ocorrências, conclui, fecha, reabre, gerencia evidências e prepara notificações pelo Outlook.
- **Leitor:** somente visualiza informações e baixa relatórios em PDF.

O usuário solicita acesso pela tela de login usando o e-mail corporativo. O cadastro entra como `pendente` e só acessa o sistema depois que uma das administradoras autorizadas escolhe o perfil e altera o status para `aprovado` na página **Gestão de Acessos**.

## Configuração do Supabase (front-end)

A URL e a chave **publishable** não ficam no código. Copie `.env.example` para `.env` (ignorado pelo git) e preencha `SUPABASE_URL` e `SUPABASE_PUBLISHABLE_KEY`. O build grava os valores em `dist/config.js` e na CSP de `dist/_headers`, e falha se faltar algum valor ou se for usada uma chave secreta.

## Testar no VS Code

Rode `npm run build` e abra `dist/index.html` com um servidor estático, como a extensão Live Server. Abrir `frontend/index.html` direto mostra o aviso de configuração ausente.

## Build

```bash
npm run check
npm run build
```

O conteúdo pronto para publicar será criado na pasta `dist`.

## GitHub e Cloudflare Pages

Crie um repositório privado e envie o projeto:

```bash
git init
git add .
git commit -m "Configura EngSafe com Outlook manual"
git branch -M main
git remote add origin URL_DO_REPOSITORIO
git push -u origin main
```

No GitHub, cadastre os secrets `CLOUDFLARE_API_TOKEN` e `CLOUDFLARE_ACCOUNT_ID`, e as variables `CLOUDFLARE_PAGES_PROJECT` (nome do projeto Pages), `SUPABASE_URL` e `SUPABASE_PUBLISHABLE_KEY`. O passo de build do workflow precisa recebê-las como `env`. Cada push na branch `main` executará a validação e publicará a pasta `dist`.

No Supabase, configure **Authentication > URL Configuration** com a URL definitiva do Pages em **Site URL** e **Redirect URLs**.

## Fluxo de notificação

1. Colaborador/Admin abre uma ocorrência elegível.
2. Seleciona os e-mails cadastrados e, se desejar, informa outros destinatários.
3. O EngSafe abre o Outlook Web com a mensagem preenchida.
4. A pessoa confirma o remetente `engenharia.ser@ambientalserra.com.br`, revisa e envia.
5. De volta ao EngSafe, clica em **Já enviei — registrar cobrança**.

O navegador não consegue confirmar se o Outlook enviou a mensagem nem anexar automaticamente o PDF por segurança. Para relatórios, o PDF é baixado e deve ser anexado manualmente na janela do Outlook.
