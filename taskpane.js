/* global Office, Word */

const CM = 28.35; // centimeters to points

// ── Supabase ──────────────────────────────────────────────────────────────────

const sb = window.supabase
  ? window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY)
  : null;

// ── Boot ──────────────────────────────────────────────────────────────────────

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    sub("Apenas para Microsoft Word");
    return;
  }
  sub("Verificando acesso...");
  await checkAuth();
});

// ── Auth ──────────────────────────────────────────────────────────────────────

async function checkAuth() {
  if (!sb) { showAuthScreen("Erro de configuração. Contate o suporte."); return; }

  try {
    const { data: { session } } = await sb.auth.getSession();
    if (!session) { showAuthScreen(); return; }

    const { data: hasSub } = await sb.rpc("current_user_has_subscription");
    if (hasSub === false) {
      showExpiredScreen();
      return;
    }

    showScreen("home");
    loadDocInfo();
  } catch {
    showAuthScreen("Sem conexão. Verifique sua internet.");
  }
}

function showAuthScreen(msg) {
  document.querySelectorAll(".screen").forEach((s) => s.classList.add("hidden"));
  document.getElementById("screen-auth").classList.remove("hidden");
  document.getElementById("back-btn").classList.add("hidden");
  document.getElementById("header-title").textContent = "LaudoBot";
  sub("");
  const el = document.getElementById("auth-msg");
  if (msg) { el.textContent = msg; el.className = "status error"; el.hidden = false; }
  else { el.hidden = true; }
}

function switchTab(tab) {
  const isLogin = tab === "login";
  document.getElementById("form-login").classList.toggle("hidden", !isLogin);
  document.getElementById("form-register").classList.toggle("hidden", isLogin);
  document.getElementById("tab-login").classList.toggle("active", isLogin);
  document.getElementById("tab-register").classList.toggle("active", !isLogin);
  document.getElementById("auth-msg").hidden = true;
}

async function doLogin() {
  const email = document.getElementById("login-email").value.trim();
  const pass  = document.getElementById("login-pass").value;
  const btn   = document.getElementById("btn-login");
  btn.disabled = true;
  setAuthMsg("Entrando...", "info");

  const { error } = await sb.auth.signInWithPassword({ email, password: pass });
  if (error) { setAuthMsg("E-mail ou senha incorretos.", "error"); btn.disabled = false; return; }
  await checkAuth();
  btn.disabled = false;
}

async function doRegister() {
  const email = document.getElementById("reg-email").value.trim();
  const pass  = document.getElementById("reg-pass").value;
  const btn   = document.getElementById("btn-register");

  if (!email || pass.length < 6) { setAuthMsg("Preencha e-mail e senha (mín. 6 caracteres).", "warn"); return; }
  btn.disabled = true;
  setAuthMsg("Criando conta...", "info");

  const { error } = await sb.auth.signUp({ email, password: pass });
  if (error) { setAuthMsg("Erro: " + error.message, "error"); btn.disabled = false; return; }
  setAuthMsg("Conta criada! Verifique seu e-mail para confirmar.", "success");
  btn.disabled = false;
}

function setAuthMsg(msg, type) {
  const el = document.getElementById("auth-msg");
  el.textContent = msg;
  el.className = "status " + type;
  el.hidden = false;
}

function showExpiredScreen() {
  document.querySelectorAll(".screen").forEach((s) => s.classList.add("hidden"));
  document.getElementById("screen-expired").classList.remove("hidden");
  document.getElementById("back-btn").classList.add("hidden");
  document.getElementById("header-title").textContent = "LaudoBot";
  sub("");
}

async function doLogout() {
  await sb.auth.signOut();
  showAuthScreen();
}

// ── Navigation ────────────────────────────────────────────────────────────────

const TITLES = {
  home:          "LaudoBot",
  legendas:      "Legendas",
  redimensionar: "Redimensionar",
  paragrafo:     "Parágrafo",
  recuo:         "Recuo",
  chat:          "Chat IA",
};

function showScreen(name) {
  document.querySelectorAll(".screen").forEach((s) => s.classList.add("hidden"));
  document.getElementById("screen-" + name).classList.remove("hidden");
  document.getElementById("header-title").textContent = TITLES[name] || name;
  document.getElementById("back-btn").classList.toggle("hidden", name === "home");
}

function showHome() {
  showScreen("home");
  loadDocInfo();
}

function toggleRecuoVal(show) {
  document.getElementById("rec-val-field").style.display = show ? "" : "none";
}

// ── Document info ─────────────────────────────────────────────────────────────

async function loadDocInfo() {
  sub("Carregando...");
  const url = Office.context.document.url;
  document.getElementById("doc-name").textContent = url
    ? decodeURIComponent(url.split(/[/\\]/).pop())
    : "Documento";

  try {
    await Word.run(async (context) => {
      const pics = context.document.body.inlinePictures;
      pics.load("items");
      await context.sync();

      const n = pics.items.length;
      const badge = document.getElementById("img-badge");
      badge.textContent = n + (n === 1 ? " imagem" : " imagens");
      badge.hidden = false;
    });
    sub("Pronto");
  } catch (e) {
    sub("Erro ao carregar");
  }
}

// ── Tool: Legendas ────────────────────────────────────────────────────────────

async function runLegendas() {
  const prefix = radio("leg-prefix");
  const align  = radio("leg-align");
  const scope  = radio("leg-scope");
  const texto  = document.getElementById("leg-texto").value.trim();

  await runTool("legendas", async (context) => {
    const pics = getPics(context, scope);
    pics.load("items");
    await context.sync();

    const n = pics.items.length;
    if (n === 0) throw new Error("Nenhuma imagem encontrada.");

    // Load the paragraph immediately after each picture to detect existing captions
    const nextParas = pics.items.map(pic => pic.paragraph.getNextOrNullObject());
    nextParas.forEach(p => p.load("text"));
    await context.sync();

    let added = 0;
    for (let i = n - 1; i >= 0; i--) {
      const next = nextParas[i];
      if (!next.isNullObject && /^(Foto|Figura)\s+\d+/.test(next.text.trim())) {
        continue; // already has a caption
      }
      const label = texto ? `${prefix} ${i + 1} - ${texto}` : `${prefix} ${i + 1}`;
      const para = pics.items[i].getRange().insertParagraph(label, "After");
      if (align) para.alignment = align;
      added++;
    }
    await context.sync();

    if (added === 0) return "Todas as imagens já têm legenda.";
    const skipped = n - added;
    return skipped > 0
      ? `${added} legenda(s) adicionada(s) (${skipped} já tinham).`
      : `${added} legenda(s) adicionada(s).`;
  });
}

// ── Tool: Redimensionar ───────────────────────────────────────────────────────

async function runRedimensionar() {
  const wCm  = num("resize-w");
  const hCm  = num("resize-h");
  const scope = radio("resize-scope");

  if (wCm === null && hCm === null) {
    setStatus("redimensionar", "Informe a largura ou a altura.", "warn");
    return;
  }

  const wPts = wCm !== null ? wCm * CM : 0;
  const hPts = hCm !== null ? hCm * CM : 0;

  await runTool("redimensionar", async (context) => {
    const pics = getPics(context, scope);
    pics.load("items");
    await context.sync();

    const n = pics.items.length;
    if (n === 0) throw new Error("Nenhuma imagem encontrada.");

    const ambos = wPts > 0 && hPts > 0;

    for (const pic of pics.items) {
      if (ambos) {
        // Ambos preenchidos → desbloqueia proporção e aplica exato
        pic.lockAspectRatio = false;
        pic.width  = wPts;
        pic.height = hPts;
      } else {
        // Apenas um lado → Word ajusta o outro automaticamente
        pic.lockAspectRatio = true;
        if (wPts > 0) pic.width  = wPts;
        else          pic.height = hPts;
      }
    }

    await context.sync();
    return `${n} imagem(ns) redimensionada(s).`;
  });
}

// ── Tool: Parágrafo ───────────────────────────────────────────────────────────

async function runParagrafo() {
  const align = radio("par-align");
  const indL  = num("par-ind-l");
  const indR  = num("par-ind-r");
  const spBef = num("par-sp-bef");
  const spAft = num("par-sp-aft");
  const ls    = radio("par-ls");
  const scope = radio("par-scope");

  await runTool("paragrafo", async (context) => {
    const pics = getPics(context, scope);
    pics.load("items");
    await context.sync();

    const n = pics.items.length;
    if (n === 0) throw new Error("Nenhuma imagem encontrada.");

    for (const pic of pics.items) {
      const para = pic.paragraph;
      if (align)      para.alignment   = align;
      if (indL !== null)  para.leftIndent  = indL * CM;
      if (indR !== null)  para.rightIndent = indR * CM;
      if (spBef !== null) para.spaceBefore = spBef * CM;
      if (spAft !== null) para.spaceAfter  = spAft * CM;
      para.lineSpacing = ls === "1.5" ? 18 : 12;
    }

    await context.sync();
    return `${n} parágrafo(s) configurado(s).`;
  });
}

// ── Tool: Recuo ───────────────────────────────────────────────────────────────

async function runRecuo() {
  const type  = radio("rec-type");
  const value = type === "first" ? (num("rec-val") ?? 1.25) : 0;
  const scope = radio("rec-scope");

  await runTool("recuo", async (context) => {
    const pics = getPics(context, scope);
    pics.load("items");
    await context.sync();

    const n = pics.items.length;
    if (n === 0) throw new Error("Nenhuma imagem encontrada.");

    for (const pic of pics.items) {
      const para = pic.paragraph;
      para.leftIndent      = 0;
      para.rightIndent     = 0;
      para.firstLineIndent = value * CM;
    }

    await context.sync();
    return `${n} imagem(ns) com recuo aplicado.`;
  });
}

// ── Chat IA ───────────────────────────────────────────────────────────────────

const CHAT_SYSTEM = "Você é um assistente especializado em documentos Word e escrita acadêmica/profissional em português brasileiro. Seja direto e útil.";

let chatMessages = [];
let chatLastReply = "";

function openChat() {
  chatMessages = [];
  chatLastReply = "";
  document.getElementById("chat-history").innerHTML = "";
  document.getElementById("btn-insert-reply").classList.add("hidden");
  document.getElementById("chat-input").value = "";
  showScreen("chat");
}

function chatAppend(role, text) {
  const history = document.getElementById("chat-history");
  const div = document.createElement("div");
  div.className = "chat-msg chat-msg-" + role;
  div.textContent = text;
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;
}

async function chatSend() {
  const input = document.getElementById("chat-input");
  const text  = input.value.trim();
  if (!text) return;

  const btn = document.getElementById("btn-send");
  input.value = "";
  btn.disabled = true;

  chatAppend("user", text);
  chatMessages.push({ role: "user", content: text });

  const thinking = document.createElement("div");
  thinking.className = "chat-msg chat-msg-thinking";
  thinking.textContent = "...";
  document.getElementById("chat-history").appendChild(thinking);

  try {
    const { data: { session } } = await sb.auth.getSession();
    const res = await fetch(CHAT_FUNCTION_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + session.access_token,
      },
      body: JSON.stringify({
        messages: [{ role: "system", content: CHAT_SYSTEM }, ...chatMessages],
      }),
    });

    const data = await res.json();
    thinking.remove();

    if (!res.ok) throw new Error(data.error || "Erro na API");

    const reply = data.choices[0].message.content;
    chatMessages.push({ role: "assistant", content: reply });
    chatLastReply = reply;
    chatAppend("assistant", reply);
    document.getElementById("btn-insert-reply").classList.remove("hidden");
  } catch (e) {
    thinking.remove();
    chatAppend("assistant", "Erro: " + e.message);
  } finally {
    btn.disabled = false;
    input.focus();
  }
}

async function chatUseSelection() {
  try {
    let selected = "";
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      selected = sel.text.trim();
    });

    if (!selected) {
      chatAppend("assistant", "Nenhum texto selecionado no Word. Selecione um trecho e tente novamente.");
      return;
    }

    const input = document.getElementById("chat-input");
    const prefix = "Texto selecionado no documento:\n\"\"\"\n" + selected + "\n\"\"\"\n\n";
    input.value = prefix;
    input.focus();
    input.setSelectionRange(prefix.length, prefix.length);
  } catch {
    chatAppend("assistant", "Não foi possível ler a seleção do Word.");
  }
}

async function chatInsertReply() {
  if (!chatLastReply) return;
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertText(chatLastReply, "End");
      await context.sync();
    });
  } catch {
    chatAppend("assistant", "Não foi possível inserir o texto no Word.");
  }
}

document.addEventListener("keydown", (e) => {
  if (e.target.id === "chat-input" && e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    chatSend();
  }
});

// ── Shared helpers ────────────────────────────────────────────────────────────

function getPics(context, scope) {
  return scope === "selected"
    ? context.document.getSelection().inlinePictures
    : context.document.body.inlinePictures;
}

async function runTool(name, fn) {
  const btn = document.querySelector(`#screen-${name} .btn-primary`);
  btn.disabled = true;
  setStatus(name, "Processando...", "info");

  try {
    let msg;
    await Word.run(async (context) => { msg = await fn(context); });
    setStatus(name, "✓  " + msg, "success");
    loadDocInfo();
  } catch (e) {
    setStatus(name, "Erro: " + e.message, "error");
  } finally {
    btn.disabled = false;
  }
}

function setStatus(tool, msg, type) {
  const el = document.getElementById("status-" + tool);
  el.textContent = msg;
  el.className = "status " + type;
  el.hidden = false;
}

// Get selected radio value by name
function radio(name) {
  const el = document.querySelector(`input[name="${name}"]:checked`);
  return el ? el.value : "";
}

// Get numeric input value, or null if empty
function num(id) {
  const v = document.getElementById(id).value.trim();
  return v === "" ? null : parseFloat(v.replace(",", "."));
}

function sub(text) {
  document.getElementById("header-sub").textContent = text;
}
