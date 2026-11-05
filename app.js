/* Monitor SLA Telefônica/Vivo — roda 100% no navegador.
 * Login Microsoft (MSAL) -> Microsoft Graph (somente leitura) -> análise local.
 * Nenhum e-mail é alterado, respondido, movido ou enviado a terceiros. */
(function () {
  "use strict";
  const CFG = window.MONITOR_CONFIG;
  const SCOPES = ["User.Read", "Mail.Read.Shared"];
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const LEVELS = ["OK", "ALTA", "CRÍTICA", "ESCALAR"];
  const TIPOS = ["Incidente/Link indisponível", "Validação Elétrica", "Faturamento/VGR", "Serviços Digitais", "Operacional/Outros"];

  const $ = (s) => document.querySelector(s);
  const state = { rows: [], demo: false, timer: null, nextAt: 0, seen: null, account: null, pca: null };

  // ---------- utilidades ----------
  const lc = (s) => (s || "").toLowerCase();
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const domainOf = (addr) => lc(addr).split("@")[1] || "";
  const fmtDate = (d) => new Intl.DateTimeFormat("pt-BR", { timeZone: CFG.fuso, day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }).format(d);
  const fmtFull = (d) => new Intl.DateTimeFormat("pt-BR", { timeZone: CFG.fuso, dateStyle: "short", timeStyle: "short" }).format(d);

  function normSubject(s) {
    let t = (s || "").trim();
    const re = /^(re|res|enc|fw|fwd|tr|rv|resp)\s*:\s*/i;
    while (re.test(t)) t = t.replace(re, "");
    return t.replace(/\s+/g, " ").trim().toLowerCase();
  }
  function addrList(m) {
    const all = [m.from, ...(m.toRecipients || []), ...(m.ccRecipients || [])];
    return all.map((r) => lc(r && r.emailAddress && r.emailAddress.address)).filter(Boolean);
  }
  function sideOf(m) {
    const a = lc(m.from && m.from.emailAddress && m.from.emailAddress.address);
    if (!a || a.startsWith("/o=")) return "comm"; // remetente interno em formato Exchange
    const d = domainOf(a);
    if (CFG.dominiosInternos.some((x) => d === x || d.endsWith("." + x))) return "comm";
    if (CFG.dominiosParceiro.some((x) => d === x || d.endsWith("." + x))) return "parceiro";
    return "externo";
  }
  function tipoOf(subject, preview, partnerAddrs) {
    const t = lc(subject + " " + preview);
    if (/valida[cç][aã]o\s+el[eé]trica/.test(t)) return TIPOS[1];
    if (/fatura|faturamento|cobran[cç]a|boleto|nota fiscal|contesta[cç][aã]o|2[aª] via/.test(t)) return TIPOS[2];
    if (/incidente|link|down|instabilidade|sem internet|sem conex|inoperante|intermit|queda|vpn|indispon/.test(t)) return TIPOS[0];
    if (partnerAddrs.includes("servicosdigitais@vivo.com.br")) return TIPOS[3];
    return TIPOS[4];
  }
  function levelOf(hours) {
    if (hours > CFG.critica[1]) return 3;
    if (hours > CFG.alta[1]) return 2;
    if (hours >= CFG.alta[0]) return 1;
    return 0;
  }

  // ---------- análise (pura; usada pelo modo real e pela demonstração) ----------
  function analyze(messages, now) {
    const parceiros = CFG.parceiros.map(lc);
    const ignorar = CFG.ignorarRemetentes.map(lc);
    const groups = new Map();
    for (const m of messages) {
      if (m.isDraft) continue;
      const key = normSubject(m.subject);
      if (!key) continue;
      if (!groups.has(key)) groups.set(key, { msgs: [], partners: new Set() });
      const g = groups.get(key);
      g.msgs.push(m);
      for (const a of addrList(m)) if (parceiros.includes(a)) g.partners.add(a);
    }
    const rows = [];
    const urg = CFG.palavrasUrgencia.map(lc);
    for (const [key, g] of groups) {
      if (!g.partners.size) continue; // conversa sem os remetentes monitorados
      const valid = g.msgs.filter((m) => !ignorar.includes(lc(m.from && m.from.emailAddress && m.from.emailAddress.address)));
      if (!valid.length) continue;
      valid.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
      const last = valid[0];
      const when = new Date(last.receivedDateTime);
      const hours = (now - when) / 36e5;
      const side = sideOf(last);
      const fromAddr = lc(last.from && last.from.emailAddress && last.from.emailAddress.address);
      const fromName = (last.from && last.from.emailAddress && last.from.emailAddress.name) || fromAddr;
      const text = lc(last.subject + " " + (last.bodyPreview || ""));
      const urgent = urg.some((w) => text.includes(w));
      let lvl = levelOf(hours);
      if (urgent && lvl > 0) lvl = Math.min(3, lvl + 1);
      rows.push({
        key, id: last.id, subject: last.subject || "(sem assunto)", preview: (last.bodyPreview || "").replace(/\s+/g, " ").slice(0, 220),
        from: fromName, fromAddr, side, status: side === "comm" ? "parceiro" : "comm",
        tipo: tipoOf(last.subject || "", last.bodyPreview || "", [...g.partners]),
        hours, when, level: LEVELS[lvl], lvl, urgent, webLink: last.webLink, count: g.msgs.length
      });
    }
    rows.sort((a, b) => b.lvl - a.lvl || b.hours - a.hours);
    return rows;
  }

  // ---------- Microsoft Graph ----------
  async function getToken() {
    const req = { scopes: SCOPES, account: state.account };
    try {
      return (await state.pca.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|consent_required|login_required/.test(e.errorCode || "")) {
        await state.pca.acquireTokenRedirect(req);
        return null;
      }
      throw e;
    }
  }
  async function graphGet(url, token) {
    for (let attempt = 0; attempt < 4; attempt++) {
      const r = await fetch(url, { headers: { Authorization: "Bearer " + token, Prefer: 'outlook.timezone="UTC"' } });
      if (r.status === 429 || r.status === 503) {
        const wait = (parseInt(r.headers.get("Retry-After") || "5", 10) || 5) * 1000;
        await new Promise((res) => setTimeout(res, wait));
        continue;
      }
      if (!r.ok) {
        let detail = "";
        try { const j = await r.json(); detail = (j.error && (j.error.code + ": " + j.error.message)) || ""; } catch (_) {}
        const err = new Error("Graph " + r.status + " " + detail);
        err.status = r.status;
        throw err;
      }
      return r.json();
    }
    throw new Error("Graph: muitas tentativas (limite de requisições).");
  }
  async function fetchMessages(token) {
    const since = new Date(Date.now() - CFG.diasRetroativos * 864e5).toISOString();
    const select = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,webLink,isDraft,conversationId";
    let url = `${GRAPH}/users/${encodeURIComponent(CFG.mailbox)}/messages?$filter=receivedDateTime ge ${since}` +
      `&$orderby=receivedDateTime desc&$top=100&$select=${select}`;
    const out = [];
    for (let page = 0; url && page < 50; page++) {
      const j = await graphGet(url, token);
      out.push(...(j.value || []));
      url = j["@odata.nextLink"];
    }
    return out;
  }

  // ---------- interface ----------
  function show(id, on) { const el = $(id); if (el) el.hidden = !on; }
  function setError(msg) { const el = $("#error"); el.textContent = msg || ""; el.hidden = !msg; }

  function explainError(e) {
    const s = String(e && e.message || e);
    if (e && e.status === 403 || /ErrorAccessDenied|Access is denied/i.test(s))
      return `Sem acesso à caixa ${CFG.mailbox}.\n• Confirme que seu usuário tem permissão de leitura (Acesso Total) nessa caixa compartilhada.\n• Confirme que o aplicativo tem a permissão delegada Mail.Read.Shared com consentimento.\n\nDetalhe: ${s}`;
    if (e && e.status === 404 || /ErrorItemNotFound|MailboxNotEnabled|ResourceNotFound/i.test(s))
      return `Caixa ${CFG.mailbox} não encontrada. Verifique o endereço em config.js.\n\nDetalhe: ${s}`;
    if (/AADSTS65001|consent/i.test(s))
      return `O aplicativo ainda não tem consentimento para ler e-mails. Peça a um administrador para conceder consentimento (README, passo 2).\n\nDetalhe: ${s}`;
    return "Erro: " + s;
  }

  function renderKpis(rows) {
    const over = rows.filter((r) => r.hours >= CFG.limiteHoras);
    const c = (f) => over.filter(f).length;
    const k = [
      ["Sem resposta há +" + CFG.limiteHoras + "h", over.length, ""],
      ["ESCALAR", c((r) => r.level === "ESCALAR"), "esc"],
      ["CRÍTICA", c((r) => r.level === "CRÍTICA"), "cri"],
      ["ALTA", c((r) => r.level === "ALTA"), "alt"],
      ["Aguardando a Comm", c((r) => r.status === "comm"), ""],
      ["Aguardando Telefônica/Vivo", c((r) => r.status === "parceiro"), ""],
      ["Com ⚡ urgência", c((r) => r.urgent), ""]
    ];
    $("#kpis").innerHTML = k.map(([l, v, cls]) => `<div class="kpi ${cls}"><div class="v">${v}</div><div class="l">${esc(l)}</div></div>`).join("");
  }

  function filtered() {
    const q = lc($("#fText").value), st = $("#fStatus").value, tp = $("#fTipo").value, all = $("#fAll").checked;
    return state.rows.filter((r) =>
      (all || r.hours >= CFG.limiteHoras) && (!st || r.status === st) && (!tp || r.tipo === tp) &&
      (!q || lc(r.subject + " " + r.from + " " + r.fromAddr + " " + r.preview).includes(q)));
  }

  function renderTable() {
    const rows = filtered();
    $("#rows").innerHTML = rows.map((r) => `
      <tr>
        <td><span class="badge b-${esc(r.level)}">${esc(r.level)}</span></td>
        <td><div class="subj">${r.urgent ? "⚡ " : ""}${esc(r.subject)}</div><div class="prev" title="${esc(r.preview)}">${esc(r.preview)}</div></td>
        <td>${esc(r.from)}<div class="prev">${esc(r.fromAddr)}</div></td>
        <td>${esc(r.tipo)}</td>
        <td class="st-${r.status}">${r.status === "comm" ? "Aguardando resposta da Comm" : "Aguardando resposta da Telefônica/Vivo"}</td>
        <td class="num">${r.hours.toFixed(1)}h</td>
        <td>${esc(fmtDate(r.when))}</td>
        <td>${r.webLink ? `<a href="${esc(r.webLink)}" target="_blank" rel="noopener">abrir</a>` : ""}</td>
      </tr>`).join("");
    const emp = $("#empty");
    emp.hidden = rows.length > 0;
    emp.textContent = state.rows.length ? "Nenhuma conversa com Telefônica/Vivo sem resposta há mais de " + CFG.limiteHoras + "h." : "Nenhuma conversa com os remetentes monitorados no período.";
  }

  function render() {
    renderKpis(state.rows);
    renderTable();
    show("#dash", true);
  }

  function notifyNew(rows) {
    const keyOf = (r) => r.key + "|" + r.id;
    const important = rows.filter((r) => r.lvl >= 2);
    if (state.seen && "Notification" in window && Notification.permission === "granted") {
      const fresh = important.filter((r) => !state.seen.has(keyOf(r)));
      if (fresh.length) {
        const top = fresh[0];
        new Notification(`${fresh.length} conversa(s) ${top.level} Telefônica/Vivo`, {
          body: `${top.urgent ? "⚡ " : ""}${top.subject} — ${top.hours.toFixed(1)}h sem resposta`, tag: "sla-monitor"
        });
      }
    }
    state.seen = new Set(important.map(keyOf));
  }

  async function refresh() {
    if (state.demo) { state.rows = analyze(demoMessages(), new Date()); render(); stampRun(); return; }
    setError("");
    $("#btnRefresh").disabled = true;
    $("#lastRun").textContent = "Atualizando...";
    try {
      const token = await getToken();
      if (!token) return;
      const msgs = await fetchMessages(token);
      state.rows = analyze(msgs, new Date());
      notifyNew(state.rows);
      render();
      $("#foot").textContent = `${msgs.length} mensagens lidas dos últimos ${CFG.diasRetroativos} dias · ${state.rows.length} conversas com os remetentes monitorados.`;
      stampRun();
    } catch (e) {
      setError(explainError(e));
      $("#lastRun").textContent = "";
    } finally {
      $("#btnRefresh").disabled = false;
    }
  }

  function stampRun() {
    state.nextAt = CFG.atualizarACadaMin > 0 ? Date.now() + CFG.atualizarACadaMin * 60e3 : 0;
    tick();
  }
  function tick() {
    const base = "Atualizado " + fmtFull(new Date(state.nextAt ? state.nextAt - CFG.atualizarACadaMin * 60e3 : Date.now()));
    if (!state.nextAt) { $("#lastRun").textContent = base; return; }
    const s = Math.max(0, Math.round((state.nextAt - Date.now()) / 1000));
    $("#lastRun").textContent = `${base} · próxima em ${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
  }
  function startTimer() {
    clearInterval(state.timer);
    state.timer = setInterval(() => {
      if (state.nextAt && Date.now() >= state.nextAt) { state.nextAt = 0; refresh(); }
      else if (state.nextAt) tick();
    }, 1000);
  }

  function exportCsv() {
    const head = ["Nivel", "Urgencia", "Assunto", "Ultima msg de", "Email", "Tipo", "Status", "Horas", "Desde", "Link"];
    const lines = filtered().map((r) => [r.level, r.urgent ? "sim" : "", r.subject, r.from, r.fromAddr, r.tipo,
      r.status === "comm" ? "Aguardando resposta da Comm" : "Aguardando resposta da Telefonica/Vivo",
      r.hours.toFixed(1).replace(".", ","), fmtFull(r.when), r.webLink || ""]);
    const csv = [head, ...lines].map((l) => l.map((v) => '"' + String(v).replace(/"/g, '""') + '"').join(";")).join("\r\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }));
    a.download = "sla-telefonica-vivo-" + new Date().toISOString().slice(0, 10) + ".csv";
    a.click();
  }

  // ---------- demonstração (dados fictícios) ----------
  function demoMessages() {
    const now = Date.now(), h = (n) => new Date(now - n * 36e5).toISOString();
    const P = (addr, name) => ({ emailAddress: { address: addr, name } });
    const vgr = P("gestao_vgr.br@telefonica.com", "Gestão VGR"), sd = P("servicosdigitais@vivo.com.br", "servicosdigitais"),
      ira = P("iranilde.santos@telefonica.com", "Iranilde (exemplo)"), comm = P("analista@commcenter.com.br", "Analista Comm (exemplo)"),
      infra = P("infra@commcenter.com.br", "Infra");
    let n = 0;
    const m = (subject, from, to, hoursAgo, preview) => ({ id: "demo" + (++n), subject, from, toRecipients: [to], ccRecipients: [],
      receivedDateTime: h(hoursAgo), bodyPreview: preview, webLink: "", isDraft: false });
    return [
      m("Validação Elétrica | 0000001 | Loja Exemplo Centro", vgr, infra, 30, "Identificamos indisponibilidade na unidade. Favor validar energia elétrica."),
      m("Link Down Loja Exemplo Norte", comm, sd, 52, "Loja sem internet desde a manhã."),
      m("Re: Link Down Loja Exemplo Norte", sd, comm, 50, "Chamado aberto, em tratativa pela equipe técnica."),
      m("INCIDENTE - LOJA EXEMPLO SUL", comm, ira, 14, "Poderia nos apoiar com o status do chamado?"),
      m("RES: INCIDENTE - LOJA EXEMPLO SUL", ira, comm, 16, "Chamado escalado para a operadora."),
      m("Contestação de fatura - outubro", vgr, infra, 6, "Segue fatura para conferência."),
      m("Instabilidade de internet - Loja Exemplo Leste", sd, comm, 2, "Chamado aberto proativamente."),
      m("RE: Configuração VPN - Loja Exemplo Oeste", comm, sd, 9.5, "Urgente: loja sem acesso ao sistema de pagamento."),
      m("Atualização cadastral", ira, infra, 26, "Favor confirmar os contatos técnicos."),
      m("ATUALIZAÇÃO: VIVO VITA | Exemplo | INC0000000", P("vivob2b@service-now.com", "Vivo B2B"), vgr, 1, "Notificação automática.")
    ];
  }
  function startDemo() {
    state.demo = true;
    show("#setup", false); show("#welcome", false); show("#demoBanner", true);
    $("#btnRefresh").hidden = false;
    state.rows = analyze(demoMessages(), new Date());
    render();
    $("#foot").textContent = "Dados fictícios gerados no navegador.";
    stampRun();
  }

  // ---------- inicialização ----------
  async function init() {
    $("#mbx").textContent = CFG.mailbox; $("#mbx2").textContent = CFG.mailbox; $("#limH").textContent = CFG.limiteHoras;
    $("#fTipo").innerHTML += TIPOS.map((t) => `<option>${esc(t)}</option>`).join("");
    ["#fText", "#fStatus", "#fTipo", "#fAll"].forEach((s) => $(s).addEventListener("input", renderTable));
    $("#btnCsv").onclick = exportCsv;
    $("#btnRefresh").onclick = () => { state.nextAt = 0; refresh(); };
    $("#btnDemo").onclick = startDemo; $("#btnDemoSetup").onclick = startDemo;
    startTimer();

    if (!CFG.clientId || /COLE-AQUI/.test(CFG.clientId)) {
      show("#setup", true); $("#btnLogin").hidden = true; return;
    }
    if (typeof msal === "undefined") { setError("Biblioteca de login não carregou (vendor/msal-browser.min.js)."); return; }

    state.pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: location.origin + location.pathname.replace(/index\.html$/, "") },
      cache: { cacheLocation: "localStorage" }
    });
    await state.pca.initialize();
    let resp = null;
    try { resp = await state.pca.handleRedirectPromise(); } catch (e) { setError(explainError(e)); }
    state.account = (resp && resp.account) || state.pca.getAllAccounts()[0] || null;

    const login = () => state.pca.loginRedirect({ scopes: SCOPES });
    $("#btnLogin").onclick = login; $("#btnLogin2").onclick = login;
    $("#btnLogout").onclick = () => state.pca.logoutRedirect({ account: state.account });

    if (!state.account) { show("#welcome", true); return; }

    state.pca.setActiveAccount(state.account);
    $("#btnLogin").hidden = true; $("#btnLogout").hidden = false; $("#btnRefresh").hidden = false;
    $("#sub").innerHTML = `Caixa monitorada: <strong>${esc(CFG.mailbox)}</strong> · ${esc(state.account.username)}`;
    if ("Notification" in window && Notification.permission === "default") {
      const b = $("#btnNotif"); b.hidden = false;
      b.onclick = () => Notification.requestPermission().then(() => { b.hidden = true; });
    }
    refresh();
  }

  window.__monitor = { analyze, normSubject, levelOf }; // útil para testes
  document.addEventListener("DOMContentLoaded", init);
})();
