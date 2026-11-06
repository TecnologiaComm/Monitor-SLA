# Monitor SLA Telefônica/Vivo

Painel web estático (GitHub Pages) que lê a caixa **infra@commcenter.com.br** e mostra as conversas com
`iranilde.santos@telefonica.com`, `gestao_vgr.br@telefonica.com` e `servicosdigitais@vivo.com.br` que estão sem resposta.

- **Sem custo de tokens de IA**: é só HTML + JavaScript. O painel faz login com a sua conta Microsoft e consulta o Microsoft Graph direto do navegador.
- **Sem servidor e sem dados no repositório**: os e-mails nunca vão para o GitHub. Eles são lidos e processados apenas no navegador de quem fez login.
- **Somente leitura**: o painel não responde, não move, não apaga e não marca nenhum e-mail.

## Regras aplicadas

| Item | Regra |
|---|---|
| Período | últimos 7 dias (`diasRetroativos`) |
| Agrupamento | assunto normalizado (sem RE:/RES:/ENC:/FW:/TR:) |
| Status | última mensagem de @telefonica.com / @vivo.com.br → **Aguardando resposta da Comm**; de @commcenter.com.br → **Aguardando resposta da Telefônica/Vivo** |
| Níveis | 4–8h **ALTA** · 8–24h **CRÍTICA** · >24h **ESCALAR** |
| Urgência ⚡ | assunto/resumo com urgente, prazo, hoje, escalonamento, crítico, bloqueio, multa, parado, sem acesso, down, inoperante → sobe um nível |
| Tipo | Incidente/Link indisponível, Validação Elétrica, Faturamento/VGR, Serviços Digitais, Operacional/Outros |
| Ignorados | avisos automáticos (`vivob2b@service-now.com`, `postmaster@...`) não contam como última resposta |
| Prováveis encerradas | última mensagem com obrigado, validado, normalizado, restabelecido, resolvido, encerrado → sai dos indicadores e vai para uma lista de conferência |
| Dispensar | esconde uma conversa (só no seu navegador) até chegar mensagem nova nela |
| Chamado | coluna própria com o número do incidente (INC…) aberto pela Telefônica/Vivo; sem INC, usa o protocolo CS…. Se o número não estiver no resumo do e-mail, o painel lê o corpo das mensagens da Telefônica/Vivo (ignorando o histórico citado) e guarda só o número em cache. Também entra na busca, no CSV e no resumo |
| Atualização | automática a cada 5 min com a aba aberta; alertas do navegador para novas conversas CRÍTICA/ESCALAR |
| Copiar resumo | gera um texto pronto para colar no Teams/WhatsApp com os 10 itens mais urgentes |
| Painel de TV | tela cheia só com os indicadores e prioridades, para deixar num monitor da equipe |

Tudo isso é ajustável em `config.js`.

---

## Instalação (≈ 15 minutos)

### 1. Publicar no GitHub Pages

1. Crie um repositório (ex.: `monitor-sla`) e envie todos os arquivos desta pasta (inclusive `vendor/` e `.nojekyll`).
2. No repositório: **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main` / `(root)` → Save**.
3. Anote a URL gerada, ex.: `https://SEU-USUARIO.github.io/monitor-sla/`.

> Um repositório privado também funciona com GitHub Pages nos planos pagos. Mesmo em repositório público não há risco de vazamento de e-mails: o site não contém dados, só código.

### 2. Registrar o aplicativo no Microsoft Entra (precisa de um admin do Microsoft 365)

1. Acesse **entra.microsoft.com → Identidade → Aplicativos → Registros de aplicativo → Novo registro**.
2. Nome: `Monitor SLA Telefônica/Vivo`. Tipos de conta: **somente este diretório organizacional**.
3. **URI de redirecionamento**: plataforma **Aplicativo de página única (SPA)**, valor = a URL do passo 1 **com a barra final** (ex.: `https://SEU-USUARIO.github.io/monitor-sla/`).
   - Para testar localmente, adicione também `http://localhost:8000/`.
4. Após criar, copie o **ID do aplicativo (cliente)** e o **ID do diretório (locatário)**.
5. **Permissões de API → Adicionar permissão → Microsoft Graph → Permissões delegadas**: marque `User.Read` e `Mail.Read.Shared`.
6. Clique em **Conceder consentimento do administrador para Comm**.

Não é necessário criar segredo (client secret) — o app usa login interativo (PKCE).

### 3. Garantir acesso à caixa infra@

Quem for usar o painel precisa ter permissão de leitura na caixa compartilhada.
No **Exchange Admin Center → Destinatários → Caixas de correio → infra@commcenter.com.br → Delegação → Ler e gerenciar (Acesso Total)**, adicione os usuários.

### 4. Configurar `config.js`

```js
clientId: "ID-DO-APLICATIVO-DO-PASSO-2",
tenantId: "0760b291-f27e-4201-ab06-3e7e68daa846",
mailbox:  "infra@commcenter.com.br",
```

Faça o commit. Em 1–2 minutos o GitHub Pages publica a nova versão.

### 5. Usar

Abra a URL, clique em **Entrar com Microsoft**, e pronto. Deixe a aba aberta para atualizar sozinha e clique em **Ativar alertas** para receber notificações do navegador.
Sem configurar nada, o botão **Ver demonstração** mostra o painel com dados fictícios.

### Testar no computador

```bash
python -m http.server 8000
# abra http://localhost:8000/
```

---

## Limitações

- O painel só enxerga o que chega na caixa **infra@**. Se alguém da Comm responder pela própria caixa sem copiar infra@, essa resposta não aparece e a conversa continuará como "Aguardando resposta da Comm". Combine com a equipe de sempre manter infra@ em cópia.
- Conversas em que o assunto foi alterado no meio da troca aparecem como linhas separadas (mesmo comportamento da rotina anterior).
- Alertas e atualização automática só funcionam com a aba aberta. Para alertas com o navegador fechado seria necessário um serviço de backend (ex.: Power Automate), que não faz parte deste projeto.

## Arquivos

| Arquivo | Função |
|---|---|
| `index.html` | página |
| `styles.css` | visual (claro/escuro automático) |
| `app.js` | login, leitura via Graph, regras de SLA, tabela, CSV, alertas |
| `config.js` | **o único arquivo que você precisa editar** |
| `vendor/msal-browser.min.js` | biblioteca oficial de login da Microsoft (v3.27.0, licença MIT) |
| `.nojekyll` | evita processamento Jekyll no GitHub Pages |
