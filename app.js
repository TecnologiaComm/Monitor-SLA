/* Monitor SLA Telefônica/Vivo — roda 100% no navegador.
 * Login Microsoft (MSAL) -> Microsoft Graph (somente leitura) -> análise local.
 * Nenhum e-mail é alterado, respondido, movido ou enviado a terceiros. */
(function () {
  "use strict";
  const CFG = window.MONITOR_CONFIG;
  const SCOPES = ["User.Read", "Mail.Read.Shared"];
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const LEVELS = ["OK", "ALTA", "CRÍTICA", "ESCALAR"];
  const TIPOS = [
    { n: "Incidente/Link indisponível", ic: "wifi", t: "t-esc" },
    { n: "Validação Elétrica", ic: "plug", t: "t-alt" },
    { n: "Faturamento/VGR", ic: "receipt", t: "t-ok" },
    { n: "Serviços Digitais", ic: "globe", t: "t-blue" },
    { n: "Operacional/Outros", ic: "wrench", t: "t-vio" }
  ];
  const ICONS = {
    home: '<path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
    chat: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
    antenna: '<path d="M5 12a7 7 0 0 1 14 0M8.5 12a3.5 3.5 0 0 1 7 0"/><path d="M12 12v9M9 21h6"/>',
    check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
    eyeoff: '<path d="M3 3l18 18M10.6 10.6a2 2 0 0 0 2.8 2.8M9.9 5.1A9.8 9.8 0 0 1 12 5c6 0 10 7 10 7a17 17 0 0 1-3.2 3.9M6.6 6.6C3.9 8.3 2 12 2 12s4 7 10 7a9.7 9.7 0 0 0 5.4-1.6"/>',
    tv: '<rect x="2" y="4" width="20" height="13" rx="2"/><path d="M8 21h8M12 17v4"/>',
    download: '<path d="M12 3v12m0 0 4-4m-4 4-4-4M4 21h16"/>',
    logout: '<path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4M10 17l5-5-5-5M15 12H3"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
    refresh: '<path d="M21 12a9 9 0 1 1-3-6.7L21 8M21 3v5h-5"/>',
    copy: '<rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>',
    bell: '<path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9M13.7 21a2 2 0 0 1-3.4 0"/>',
    chev: '<path d="m9 6 6 6-6 6"/>',
    alert: '<path d="M12 3 2 20h20zM12 10v4M12 17.5v.5"/>',
    flame: '<path d="M12 22a7 7 0 0 0 7-7c0-4-3-6-4-10-2 2-3 4-3 6-1-1-2-2-2-4-2 2-5 5-5 8a7 7 0 0 0 7 7z"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    bolt: '<path d="M13 2 4 14h7l-1 8 9-12h-7z"/>',
    wifi: '<path d="M2 8.8a15 15 0 0 1 20 0M5 12.6a10 10 0 0 1 14 0M8.5 16.4a5 5 0 0 1 7 0"/><path d="M3 3l18 18"/>',
    plug: '<path d="M9 2v6M15 2v6M6 8h12v4a6 6 0 0 1-12 0zM12 18v4"/>',
    receipt: '<path d="M5 2h14v20l-3-2-2 2-2-2-2 2-2-2-3 2zM9 7h6M9 11h6M9 15h4"/>',
    globe: '<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"/>',
    wrench: '<path d="M14.7 6.3a4 4 0 0 0 5 5L22 14l-8 8-2.3-2.3a4 4 0 0 0-5-5L3 11l8-8z"/>'
  };
  const svg = (k) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${ICONS[k] || ""}</svg>`;
  const icon = (k) => `<i data-i="${k}">${svg(k)}</i>`;

  const $ = (s) => document.querySelector(s);
  const state = { rows: [], demo: false, timer: null, nextAt: 0, lastRun: 0, seen: null, account: null, pca: null,
    filter: null, nMsgs: 0, dismissed: loadDismissed() };

  // ---------- utilidades ----------
  const lc = (s) => (s || "").toLowerCase();
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const domainOf = (addr) => lc(addr).split("@")[1] || "";
  const fmt = (d, o) => new Intl.DateTimeFormat("pt-BR", Object.assign({ timeZone: CFG.fuso }, o)).format(d);
  const fmtDate = (d) => fmt(d, { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
  const fmtFull = (d) => fmt(d, { dateStyle: "short", timeStyle: "short" });
  const hfmt = (h) => h >= 48 ? `${Math.floor(h / 24)}d ${Math.round(h % 24)}h` : `${h.toFixed(1)}h`;

  function loadDismissed() { try { return JSON.parse(localStorage.getItem("sla-dismissed") || "{}"); } catch (_) { return {}; } }
  function saveDismissed() { try { localStorage.setItem("sla-dismissed", JSON.stringify(state.dismissed)); } catch (_) {} }

  function normSubject(s) {
    let t = (s || "").trim();
    const re = /^(re|res|enc|fw|fwd|tr|rv|resp)\s*:\s*/i;
    while (re.test(t)) t = t.replace(re, "");
    return t.replace(/\s+/g, " ").trim().toLowerCase();
  }
  function addrList(m) {
    return [m.from, ...(m.toRecipients || []), ...(m.ccRecipients || [])]
      .map((r) => lc(r && r.emailAddress && r.emailAddress.address)).filter(Boolean);
  }
  function sideOf(m) {
    const a = lc(m.from && m.from.emailAddress && m.from.emailAddress.address);
    if (!a || a.startsWith("/o=")) return "comm";
    const d = domainOf(a);
    if (CFG.dominiosInternos.some((x) => d === x || d.endsWith("." + x))) return "comm";
    if (CFG.dominiosParceiro.some((x) => d === x || d.endsWith("." + x))) return "parceiro";
    return "externo";
  }
  function tipoOf(subject, preview, partners) {
    const t = lc(subject + " " + preview);
    if (/valida[cç][aã]o\s+el[eé]trica/.test(t)) return TIPOS[1].n;
    if (/fatura|faturamento|cobran[cç]a|boleto|nota fiscal|contesta[cç][aã]o|2[aª] via/.test(t)) return TIPOS[2].n;
    if (/incidente|link|down|instabilidade|sem internet|sem conex|inoperante|intermit|queda|vpn|indispon/.test(t)) return TIPOS[0].n;
    if (partners.includes("servicosdigitais@vivo.com.br")) return TIPOS[3].n;
    return TIPOS[4].n;
  }
  const TK_RE = /\b(INC|CS)\s*[:#nº°-]?\s*(\d{6,})\b/gi;
  function ticketsIn(text) {
    const out = []; let m; TK_RE.lastIndex = 0;
    while ((m = TK_RE.exec(text || ""))) out.push(m[1].toUpperCase() + m[2]);
    return out;
  }
  function levelOf(h) { return h > CFG.critica[1] ? 3 : h > CFG.alta[1] ? 2 : h >= CFG.alta[0] ? 1 : 0; }

  // ---------- análise (pura) ----------
  function analyze(messages, now) {
    const parceiros = CFG.parceiros.map(lc), ignorar = CFG.ignorarRemetentes.map(lc);
    const urg = CFG.palavrasUrgencia.map(lc), fim = (CFG.palavrasEncerramento || []).map(lc);
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
    for (const [key, g] of groups) {
      if (!g.partners.size) continue;
      const valid = g.msgs.filter((m) => !ignorar.includes(lc(m.from && m.from.emailAddress && m.from.emailAddress.address)));
      if (!valid.length) continue;
      valid.sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
      const last = valid[0], when = new Date(last.receivedDateTime), hours = (now - when) / 36e5;
      const side = sideOf(last);
      const fromAddr = lc(last.from && last.from.emailAddress && last.from.emailAddress.address);
      const fromName = (last.from && last.from.emailAddress && last.from.emailAddress.name) || fromAddr;
      const preview = (last.bodyPreview || "").replace(/\s+/g, " ");
      const text = lc(last.subject + " " + preview);
      const head = lc(preview.slice(0, 160));
      const urgent = urg.some((w) => text.includes(w));
      const closed = !!CFG.separarEncerradas && fim.some((w) => head.includes(w));
      let lvl = levelOf(hours);
      if (urgent && lvl > 0) lvl = Math.min(3, lvl + 1);
      // chamados: varre da mensagem mais nova para a mais antiga; prioriza o INC aberto pela Telefônica/Vivo
      const byNew = [...g.msgs].sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime));
      const tickets = new Set(); let incP = "", inc = "", cs = "";
      for (const m of byNew) {
        const found = ticketsIn((m.subject || "") + " " + (m.bodyPreview || "") + " " + (m._tk || ""));
        for (const t of found) {
          tickets.add(t);
          if (t.startsWith("INC")) { if (!incP && sideOf(m) === "parceiro") incP = t; if (!inc) inc = t; }
          else if (!cs) cs = t;
        }
      }
      const chamado = incP || inc || cs;
      rows.push({
        key, id: last.id, subject: last.subject || "(sem assunto)", preview: preview.slice(0, 220),
        from: fromName, fromAddr, status: side === "comm" ? "parceiro" : "comm",
        tipo: tipoOf(last.subject || "", preview, [...g.partners]),
        hours, when, level: LEVELS[lvl], lvl, urgent, closed, chamado, tickets: [...tickets].filter((t) => t !== chamado).slice(0, 3),
        webLink: last.webLink, count: g.msgs.length
      });
    }
    rows.sort((a, b) => b.lvl - a.lvl || b.hours - a.hours);
    return rows;
  }

  // ---------- Microsoft Graph ----------
  async function getToken() {
    const req = { scopes: SCOPES, account: state.account };
    try { return (await state.pca.acquireTokenSilent(req)).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|consent_required|login_required/.test(e.errorCode || "")) {
        await state.pca.acquireTokenRedirect(req); return null;
      }
      throw e;
    }
  }
  async function graphGet(url, token, extra) {
    for (let i = 0; i < 4; i++) {
      const r = await fetch(url, { headers: Object.assign({ Authorization: "Bearer " + token }, extra || {}) });
      if (r.status === 429 || r.status === 503) {
        await new Promise((res) => setTimeout(res, (parseInt(r.headers.get("Retry-After") || "5", 10) || 5) * 1000));
        continue;
      }
      if (!r.ok) {
        let d = ""; try { const j = await r.json(); d = j.error ? j.error.code + ": " + j.error.message : ""; } catch (_) {}
        const err = new Error("Graph " + r.status + " " + d); err.status = r.status; throw err;
      }
      return r.json();
    }
    throw new Error("Graph: limite de requisições atingido. Tente em alguns minutos.");
  }
  async function fetchMessages(token) {
    const since = new Date(Date.now() - CFG.diasRetroativos * 864e5).toISOString();
    const select = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,webLink,isDraft";
    let url = `${GRAPH}/users/${encodeURIComponent(CFG.mailbox)}/messages?$filter=receivedDateTime ge ${since}&$orderby=receivedDateTime desc&$top=100&$select=${select}`;
    const out = [];
    for (let p = 0; url && p < 50; p++) { const j = await graphGet(url, token); out.push(...(j.value || [])); url = j["@odata.nextLink"]; }
    return out;
  }

  // Busca o número do incidente no corpo completo quando ele não aparece no resumo (bodyPreview).
  // Só guarda os números encontrados (nunca o texto do e-mail), em cache no navegador.
  function loadTkCache() { try { return JSON.parse(localStorage.getItem("sla-tk") || "{}"); } catch (_) { return {}; } }
  function saveTkCache(c) { try { localStorage.setItem("sla-tk", JSON.stringify(c)); } catch (_) {} }
  async function enrichTickets(msgs, token) {
    const cache = loadTkCache(), keep = {};
    for (const m of msgs) if (cache[m.id] !== undefined) { m._tk = cache[m.id]; keep[m.id] = cache[m.id]; }
    const rows = analyze(msgs, new Date());
    const missing = new Set(rows.filter((r) => !r.chamado).map((r) => r.key));
    const todo = msgs.filter((m) => missing.has(normSubject(m.subject)) && m._tk === undefined && sideOf(m) === "parceiro" && !m.isDraft)
      .sort((a, b) => new Date(b.receivedDateTime) - new Date(a.receivedDateTime)).slice(0, CFG.maxCorposPorAtualizacao || 40);
    for (const m of todo) {
      try {
        const j = await graphGet(`${GRAPH}/users/${encodeURIComponent(CFG.mailbox)}/messages/${encodeURIComponent(m.id)}?$select=body`, token,
          { Prefer: 'outlook.body-content-type="text"' });
        // ignora o histórico citado abaixo de "De:/From:" para não herdar chamado de outra conversa
        const body = ((j.body && j.body.content) || "").split(/\n\s*(?:De|From)\s*:/)[0];
        m._tk = [...new Set(ticketsIn(body))].join(" ");
      } catch (_) { m._tk = ""; }
      keep[m.id] = m._tk;
    }
    saveTkCache(keep);
  }

  // ---------- visões ----------
  const isDismissed = (r) => state.dismissed[r.key] === r.id;
  const active = () => state.rows.filter((r) => !isDismissed(r) && !r.closed);
  const overLimit = () => active().filter((r) => r.hours >= CFG.limiteHoras);

  const FILTERS = {
    ESCALAR: { label: "ESCALAR", f: (r) => r.level === "ESCALAR" },
    CRÍTICA: { label: "CRÍTICA", f: (r) => r.level === "CRÍTICA" },
    ALTA: { label: "ALTA", f: (r) => r.level === "ALTA" },
    urg: { label: "⚡ Urgentes", f: (r) => r.urgent },
    comm: { label: "Aguardando a Comm", f: (r) => r.status === "comm" },
    parceiro: { label: "Aguardando Telefônica/Vivo", f: (r) => r.status === "parceiro" },
    encerradas: { label: "Prováveis encerradas", src: "closed" },
    dispensadas: { label: "Dispensadas", src: "dismissed" }
  };
  TIPOS.forEach((t) => (FILTERS["tipo:" + t.n] = { label: t.n, f: (r) => r.tipo === t.n }));

  function listRows() {
    const F = state.filter && FILTERS[state.filter];
    let base;
    if (F && F.src === "closed") base = state.rows.filter((r) => r.closed && !isDismissed(r));
    else if (F && F.src === "dismissed") base = state.rows.filter(isDismissed);
    else base = active().filter((r) => $("#fAll").checked || r.hours >= CFG.limiteHoras);
    if (F && F.f) base = base.filter(F.f);
    const q = lc($("#fText").value), st = $("#fStatus").value, tp = $("#fTipo").value;
    return base.filter((r) => (!st || r.status === st) && (!tp || r.tipo === tp) &&
      (!q || lc([r.subject, r.from, r.fromAddr, r.preview, r.chamado, r.tickets.join(" ")].join(" ")).includes(q)));
  }

  function card(fkey, ic, tile, title, num, label, extra, foot) {
    return `<button class="card${state.filter === fkey ? " sel" : ""}" data-f="${esc(fkey)}">
      <span class="tile ${tile}">${icon(ic)}</span><h5>${esc(title)}</h5>
      <div class="big"><b>${num}</b><small>${esc(label)}</small>${extra || ""}</div>
      <div class="cf"><span>${esc(foot)}</span>${icon("chev")}</div></button>`;
  }

  function renderCards() {
    const o = overLimit(), c = (f) => o.filter(f).length;
    const lvlCard = (lv, ic, tile, desc) => {
      const set = o.filter((r) => r.level === lv);
      return card(lv, ic, tile, lv, set.length, "conversas", "", `${set.filter((r) => r.status === "comm").length} com a Comm · ${desc}`);
    };
    $("#gSla").innerHTML =
      lvlCard("ESCALAR", "flame", "t-esc", "> 24h") +
      lvlCard("CRÍTICA", "alert", "t-cri", "8–24h") +
      lvlCard("ALTA", "clock", "t-alt", "4–8h") +
      card("urg", "bolt", "t-acc", "⚡ Urgentes", c((r) => r.urgent), "com palavra-chave", "", "subiram um nível");

    const comm = o.filter((r) => r.status === "comm"), parc = o.filter((r) => r.status === "parceiro");
    const oldest = (s) => s.length ? "mais antiga: " + hfmt(Math.max(...s.map((r) => r.hours))) : "nenhuma pendente";
    const closed = state.rows.filter((r) => r.closed && !isDismissed(r)).length;
    $("#gResp").innerHTML =
      card("comm", "user", "t-esc", "Aguardando a Comm", comm.length, "pendentes",
        comm.filter((r) => r.lvl === 3).length ? `<span class="sep">${comm.filter((r) => r.lvl === 3).length}</span><small>escalar</small>` : "", oldest(comm)) +
      card("parceiro", "antenna", "t-blue", "Aguardando Telefônica/Vivo", parc.length, "pendentes",
        parc.filter((r) => r.lvl === 3).length ? `<span class="sep">${parc.filter((r) => r.lvl === 3).length}</span><small>escalar</small>` : "", oldest(parc)) +
      card("encerradas", "check", "t-ok", "Prováveis encerradas", closed, "para conferir", "", "fora dos indicadores");

    $("#gTipo").innerHTML = TIPOS.map((t) => {
      const s = o.filter((r) => r.tipo === t.n);
      return card("tipo:" + t.n, t.ic, t.t, t.n, s.length, "pendentes", "", s.length ? `${s.filter((r) => r.status === "comm").length} com a Comm` : "tudo em dia");
    }).join("");
  }

  function renderSide() {
    const o = overLimit();
    const top = o.slice(0, 6);
    $("#plist").innerHTML = top.length ? top.map((r) => `<li><span class="badge b-${r.level}">${r.level}</span>
      <div class="grow"><a class="t" href="${esc(r.webLink || "#")}" target="_blank" rel="noopener">${r.urgent ? "⚡ " : ""}${esc(r.subject)}</a>
      <div class="m">${r.chamado ? esc(r.chamado) + " · " : ""}${hfmt(r.hours)} · ${r.status === "comm" ? "falta a Comm responder" : "aguardando Vivo"}</div></div></li>`).join("")
      : `<li><span class="m">Nenhuma conversa com Telefônica/Vivo sem resposta há mais de ${CFG.limiteHoras}h.</span></li>`;
    const all = active().length || 1;
    $("#barFill").style.width = Math.round((o.filter((r) => r.lvl >= 2).length / all) * 100) + "%";

    const cl = state.rows.filter((r) => r.closed && !isDismissed(r));
    $("#cClosed").textContent = cl.length;
    $("#clist").innerHTML = cl.slice(0, 6).map((r) => `<li><div class="grow"><div class="t">${esc(r.subject)}</div>
      <div class="m">${esc(r.from)}: “${esc(r.preview.slice(0, 70))}…”</div></div><button data-dismiss="${esc(r.key)}">Dispensar</button></li>`).join("")
      || `<li><span class="m">Nada por aqui.</span></li>`;
  }

  function renderTable() {
    const rows = listRows();
    const F = state.filter && FILTERS[state.filter];
    const chip = $("#chip");
    chip.hidden = !F;
    if (F) chip.innerHTML = `Filtro: ${esc(F.label)} <button title="Limpar filtro" data-clear>✕</button>`;
    const dismissedView = F && F.src === "dismissed";
    $("#rows").innerHTML = rows.map((r) => `<tr>
      <td><span class="badge b-${r.level}">${r.level}</span></td>
      <td class="chamado">${r.chamado ? `<b>${esc(r.chamado)}</b>` : `<span class="since">—</span>`}${r.tickets.map((t) => `<div class="since">${esc(t)}</div>`).join("")}</td>
      <td><div class="subj">${r.urgent ? "⚡ " : ""}${esc(r.subject)}</div><div class="prev" title="${esc(r.preview)}">${esc(r.preview)}</div>
        <span class="tk tipo">${esc(r.tipo)}</span>${r.count > 1 ? `<span class="tk">${r.count} msgs</span>` : ""}</td>
      <td><span class="pill ${r.status}" title="${r.status === "comm" ? "Aguardando resposta da Comm" : "Aguardando resposta da Telefônica/Vivo"}">${r.status === "comm" ? "Falta a Comm" : "Aguardando Vivo"}</span>
        <div class="since who" title="${esc(r.fromAddr)}">última: ${esc(r.from)}</div></td>
      <td class="num"><b>${hfmt(r.hours)}</b><div class="since">desde ${esc(fmtDate(r.when))}</div><div class="acts">${r.webLink ? `<a href="${esc(r.webLink)}" target="_blank" rel="noopener">abrir</a>` : ""}
        ${dismissedView ? `<button data-restore="${esc(r.key)}">restaurar</button>` : `<button data-dismiss="${esc(r.key)}" title="Some até chegar mensagem nova">dispensar</button>`}</div></td>
    </tr>`).join("");
    const emp = $("#empty");
    emp.hidden = rows.length > 0;
    emp.textContent = state.rows.length ? "Nenhuma conversa neste filtro." : "Nenhuma conversa com os remetentes monitorados no período.";
  }

  function render() {
    renderCards(); renderSide(); renderTable();
    const o = overLimit(), esc3 = o.filter((r) => r.lvl === 3).length;
    $("#heroSub").innerHTML = o.length
      ? `${fmt(new Date(), { weekday: "long", day: "numeric", month: "long" })} · <strong>${o.length} conversas</strong> sem resposta há +${CFG.limiteHoras}h${esc3 ? `, ${esc3} para escalar` : ""}`
      : `${fmt(new Date(), { weekday: "long", day: "numeric", month: "long" })} · nada pendente com Telefônica/Vivo`;
    $("#dash").hidden = false;
  }

  function greet(name) {
    const h = +fmt(new Date(), { hour: "numeric", hour12: false });
    const g = h < 12 ? "Bom dia" : h < 18 ? "Boa tarde" : "Boa noite";
    $("#greet").textContent = name ? `${g}, ${name}` : g;
  }

  // ---------- ações ----------
  function toast(msg) { const t = $("#toast"); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), 2600); }
  function setError(msg) { const el = $("#error"); el.textContent = msg || ""; el.hidden = !msg; }
  function explainError(e) {
    const s = String((e && e.message) || e);
    if ((e && e.status === 403) || /ErrorAccessDenied|Access is denied/i.test(s))
      return `Sem acesso à caixa ${CFG.mailbox}.\n• Seu usuário precisa de Acesso Total (Full Access) nessa caixa.\n• O app precisa da permissão Mail.Read.Shared com consentimento do administrador.\n\nDetalhe: ${s}`;
    if ((e && e.status === 404) || /ErrorItemNotFound|MailboxNotEnabled|ResourceNotFound/i.test(s))
      return `Caixa ${CFG.mailbox} não encontrada. Verifique o endereço em config.js.\n\nDetalhe: ${s}`;
    if (/AADSTS65001|AADSTS90094|consent/i.test(s))
      return `Falta o consentimento do administrador para o app (README, passo 2.6).\n\nDetalhe: ${s}`;
    return "Erro: " + s;
  }

  function setFilter(k) {
    state.filter = state.filter === k ? null : k;
    document.querySelectorAll("#nav a").forEach((a) => a.classList.toggle("active",
      a.dataset.nav === (state.filter && ["comm", "parceiro", "encerradas", "dispensadas"].includes(state.filter) ? state.filter : "inicio")));
    renderCards(); renderTable();
    if (state.filter) $("#secConv").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function dismiss(key, restore) {
    const r = state.rows.find((x) => x.key === key);
    if (!r) return;
    if (restore) delete state.dismissed[key]; else state.dismissed[key] = r.id;
    saveDismissed(); render();
    toast(restore ? "Conversa restaurada." : "Dispensada. Volta se chegar mensagem nova.");
  }

  function summaryText() {
    const o = overLimit(), c = (lv) => o.filter((r) => r.level === lv).length;
    const lines = [`*SLA Telefônica/Vivo — ${fmtFull(new Date())}*`,
      `${o.length} conversas sem resposta há +${CFG.limiteHoras}h (ESCALAR ${c("ESCALAR")} · CRÍTICA ${c("CRÍTICA")} · ALTA ${c("ALTA")})`,
      `Aguardando a Comm: ${o.filter((r) => r.status === "comm").length} · Aguardando Vivo: ${o.filter((r) => r.status === "parceiro").length}`, ""];
    o.slice(0, 10).forEach((r, i) => lines.push(`${i + 1}. [${r.level}] ${r.urgent ? "⚡ " : ""}${r.chamado ? r.chamado + " · " : ""}${r.subject} — ${hfmt(r.hours)} — ${r.status === "comm" ? "falta a Comm" : "aguardando Vivo"}`));
    if (o.length > 10) lines.push(`… e mais ${o.length - 10}.`);
    return lines.join("\n");
  }
  async function copySummary() {
    try { await navigator.clipboard.writeText(summaryText()); toast("Resumo copiado. Cole no Teams ou WhatsApp."); }
    catch (_) { toast("Não foi possível copiar."); }
  }

  function exportCsv() {
    const head = ["Nivel", "Urgencia", "Chamado", "Outros chamados", "Assunto", "Ultima msg de", "Email", "Tipo", "Status", "Horas", "Desde", "Provavel encerrada", "Link"];
    const lines = listRows().map((r) => [r.level, r.urgent ? "sim" : "", r.chamado, r.tickets.join(" "), r.subject, r.from, r.fromAddr, r.tipo,
      r.status === "comm" ? "Aguardando resposta da Comm" : "Aguardando resposta da Telefonica/Vivo",
      r.hours.toFixed(1).replace(".", ","), fmtFull(r.when), r.closed ? "sim" : "", r.webLink || ""]);
    const csv = [head, ...lines].map((l) => l.map((v) => '"' + String(v).replace(/"/g, '""') + '"').join(";")).join("\r\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" }));
    a.download = "sla-telefonica-vivo-" + new Date().toISOString().slice(0, 10) + ".csv";
    a.click();
  }

  function toggleTv(on) {
    const tv = on ?? !document.body.classList.contains("tv");
    document.body.classList.toggle("tv", tv);
    if (tv && document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
    if (!tv && document.fullscreenElement) document.exitFullscreen().catch(() => {});
    if (tv) toast("Painel de TV — Esc para sair.");
  }

  function notifyNew(rows) {
    const keyOf = (r) => r.key + "|" + r.id;
    const imp = rows.filter((r) => r.lvl >= 2 && !r.closed && !isDismissed(r));
    if (state.seen && "Notification" in window && Notification.permission === "granted") {
      const fresh = imp.filter((r) => !state.seen.has(keyOf(r)));
      if (fresh.length) new Notification(`${fresh.length} conversa(s) ${fresh[0].level} — Telefônica/Vivo`,
        { body: `${fresh[0].urgent ? "⚡ " : ""}${fresh[0].subject} — ${hfmt(fresh[0].hours)} sem resposta`, tag: "sla-monitor" });
    }
    state.seen = new Set(imp.map(keyOf));
  }

  async function refresh() {
    if (state.demo) { state.rows = analyze(demoMessages(), new Date()); state.nMsgs = demoMessages().length; render(); stamp(); return; }
    if (!state.account) return;
    setError(""); $("#btnRefresh").disabled = true;
    try {
      const token = await getToken();
      if (!token) return;
      const msgs = await fetchMessages(token);
      state.nMsgs = msgs.length;
      await enrichTickets(msgs, token);
      state.rows = analyze(msgs, new Date());
      // limpa dispensas antigas (conversa recebeu mensagem nova ou saiu da janela)
      for (const k of Object.keys(state.dismissed)) { const r = state.rows.find((x) => x.key === k); if (!r || r.id !== state.dismissed[k]) delete state.dismissed[k]; }
      saveDismissed();
      notifyNew(state.rows); render(); stamp();
    } catch (e) { setError(explainError(e)); }
    finally { $("#btnRefresh").disabled = false; }
  }
  function stamp() {
    state.lastRun = Date.now();
    state.nextAt = CFG.atualizarACadaMin > 0 ? Date.now() + CFG.atualizarACadaMin * 60e3 : 0;
    tick();
  }
  function tick() {
    if (!state.lastRun) return;
    let s = `Atualizado às ${fmt(new Date(state.lastRun), { hour: "2-digit", minute: "2-digit" })} · ${state.nMsgs} mensagens lidas (${CFG.diasRetroativos} dias) · ${state.rows.length} conversas monitoradas`;
    if (state.nextAt) { const r = Math.max(0, Math.round((state.nextAt - Date.now()) / 1000)); s += ` · próxima em ${Math.floor(r / 60)}:${String(r % 60).padStart(2, "0")}`; }
    $("#runInfo").textContent = s;
  }

  // ---------- demonstração ----------
  function demoMessages() {
    const now = Date.now(), h = (n) => new Date(now - n * 36e5).toISOString();
    const P = (a, n) => ({ emailAddress: { address: a, name: n } });
    const vgr = P("gestao_vgr.br@telefonica.com", "Gestão VGR"), sd = P("servicosdigitais@vivo.com.br", "servicosdigitais"),
      ira = P("iranilde.santos@telefonica.com", "Iranilde (exemplo)"), comm = P("analista@commcenter.com.br", "Analista Comm (exemplo)"),
      infra = P("infra@commcenter.com.br", "Infra");
    let n = 0;
    const m = (subject, from, to, ago, prev) => ({ id: "demo" + ++n, subject, from, toRecipients: [to], ccRecipients: [], receivedDateTime: h(ago), bodyPreview: prev, webLink: "", isDraft: false });
    return [
      m("Validação Elétrica | 0000001 | Loja Exemplo Centro", vgr, infra, 30, "Identificamos indisponibilidade na unidade. Incidente INC10000001. Favor validar energia."),
      m("Link Down Loja Exemplo Norte", comm, sd, 52, "Loja sem internet desde a manhã."),
      m("Re: Link Down Loja Exemplo Norte", sd, comm, 50, "Chamado CS11000001 aberto, em tratativa pela equipe técnica."),
      m("INCIDENTE - LOJA EXEMPLO SUL", comm, ira, 14, "Poderia nos apoiar com o status do chamado CS11000002?"),
      m("RES: INCIDENTE - LOJA EXEMPLO SUL", ira, comm, 16, "Chamado escalado para a operadora."),
      m("Contestação de fatura - outubro", vgr, infra, 6, "Segue fatura para conferência."),
      m("Instabilidade de internet - Loja Exemplo Leste", sd, comm, 2, "Chamado aberto proativamente."),
      m("RE: Configuração VPN - Loja Exemplo Oeste", comm, sd, 9.5, "Urgente: loja sem acesso ao sistema de pagamento."),
      m("Atualização cadastral", ira, infra, 26, "Favor confirmar os contatos técnicos."),
      m("RE: Incidente - Loja Exemplo Shopping", comm, vgr, 40, "Obrigado a todos pelo apoio, internet restabelecida."),
      m("Re: Instabilidade - Loja Exemplo Praia", sd, comm, 70, "Validamos que a localidade está normalizada após retorno da energia."),
      m("ATUALIZAÇÃO: VIVO VITA | Exemplo | INC0000000", P("vivob2b@service-now.com", "Vivo B2B"), vgr, 1, "Notificação automática.")
    ];
  }
  function startDemo() {
    state.demo = true;
    $("#setup").hidden = true; $("#welcome").hidden = true; $("#demoBanner").hidden = false;
    greet("equipe Infra");
    refresh();
  }

  // ---------- inicialização ----------
  async function init() {
    document.querySelectorAll("i[data-i]").forEach((el) => (el.innerHTML = svg(el.dataset.i)));
    $("#mbx").textContent = CFG.mailbox; $("#mbx2").textContent = CFG.mailbox;
    document.querySelectorAll(".limH").forEach((e) => (e.textContent = CFG.limiteHoras));
    $("#fTipo").innerHTML += TIPOS.map((t) => `<option>${esc(t.n)}</option>`).join("");
    greet("");

    ["#fText", "#fStatus", "#fTipo", "#fAll"].forEach((s) => $(s).addEventListener("input", renderTable));
    document.addEventListener("click", (e) => {
      const c = e.target.closest("[data-f]"); if (c) return setFilter(c.dataset.f);
      const d = e.target.closest("[data-dismiss]"); if (d) return dismiss(d.dataset.dismiss);
      const r = e.target.closest("[data-restore]"); if (r) return dismiss(r.dataset.restore, true);
      if (e.target.closest("[data-clear]")) return setFilter(state.filter);
      const nv = e.target.closest("[data-nav]");
      if (nv) {
        e.preventDefault();
        const k = nv.dataset.nav;
        if (k === "tv") return toggleTv(true);
        if (k === "csv") return exportCsv();
        if (k === "inicio") { if (state.filter) setFilter(state.filter); window.scrollTo({ top: 0, behavior: "smooth" }); return; }
        if (k === "conversas") { if (state.filter) setFilter(state.filter); $("#secConv").scrollIntoView({ behavior: "smooth" }); return; }
        if (state.filter !== k) setFilter(k);
      }
    });
    document.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") { e.preventDefault(); $("#fText").focus(); }
      if (e.key === "Escape" && document.body.classList.contains("tv")) toggleTv(false);
    });
    document.addEventListener("fullscreenchange", () => { if (!document.fullscreenElement) document.body.classList.remove("tv"); });
    $("#btnRefresh").onclick = refresh; $("#btnCopy").onclick = copySummary; $("#btnCsv").onclick = exportCsv; $("#btnTv").onclick = () => toggleTv();
    $("#btnNotif").onclick = () => {
      if (!("Notification" in window)) return toast("Este navegador não suporta alertas.");
      Notification.requestPermission().then((p) => toast(p === "granted" ? "Alertas ativados para CRÍTICA/ESCALAR." : "Alertas não autorizados no navegador."));
    };
    $("#btnDemo").onclick = startDemo; $("#btnDemoSetup").onclick = startDemo;
    setInterval(() => { if (state.nextAt && Date.now() >= state.nextAt) { state.nextAt = 0; refresh(); } else tick(); }, 1000);

    if (!CFG.clientId || /COLE-AQUI/.test(CFG.clientId)) { $("#setup").hidden = false; return; }
    if (typeof msal === "undefined") { setError("Biblioteca de login não carregou (vendor/msal-browser.min.js)."); return; }

    state.pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId, redirectUri: location.origin + location.pathname.replace(/index\.html$/, "") },
      cache: { cacheLocation: "localStorage" }
    });
    await state.pca.initialize();
    let resp = null;
    try { resp = await state.pca.handleRedirectPromise(); } catch (e) { setError(explainError(e)); }
    state.account = (resp && resp.account) || state.pca.getAllAccounts()[0] || null;
    $("#btnLogin").onclick = () => state.pca.loginRedirect({ scopes: SCOPES });
    $("#btnLogout").onclick = (e) => { e.preventDefault(); state.pca.logoutRedirect({ account: state.account }); };
    if (!state.account) { $("#welcome").hidden = false; return; }

    state.pca.setActiveAccount(state.account);
    const name = state.account.name || state.account.username;
    greet(name.split(/\s+/)[0]);
    $("#me").hidden = false; $("#meName").textContent = name; $("#meMail").textContent = state.account.username;
    $("#avatar").textContent = name.split(/\s+/).filter(Boolean).slice(0, 2).map((p) => p[0]).join("").toUpperCase();
    refresh();
  }

  window.__monitor = { analyze, normSubject, levelOf, enrichTickets };
  document.addEventListener("DOMContentLoaded", init);
})();
