/* global Office, Word */

const CM = 28.35; // centimeters to points
const LAB_VERSION = "24/05 · 18:40";

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
  lab:           "Lab",
  "legendas-lab": "Legendas (Lab)",
};

function showScreen(name) {
  document.querySelectorAll(".screen").forEach((s) => s.classList.add("hidden"));
  document.getElementById("screen-" + name).classList.remove("hidden");
  document.getElementById("header-title").textContent = TITLES[name] || name;
  document.getElementById("back-btn").classList.toggle("hidden", name === "home");
  if (name === "lab") {
    document.getElementById("lab-version").textContent = "Atualizado em " + LAB_VERSION;
  }
}

function showHome() {
  showScreen("home");
  loadDocInfo();
}

function goBack() {
  const current = document.querySelector(".screen:not(.hidden)")?.id.replace("screen-", "") || "home";
  if (current === "lab" || current === "home") { showHome(); }
  else if (current.endsWith("-lab"))           { showScreen("lab"); }
  else                                          { showHome(); }
}

// ── Lab secret menu (5 taps on title) ────────────────────────────────────────

let _labTaps = 0, _labTimer = null;
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("header-title").addEventListener("click", () => {
    _labTaps++;
    clearTimeout(_labTimer);
    if (_labTaps >= 5) { _labTaps = 0; showScreen("lab"); }
    else { _labTimer = setTimeout(() => { _labTaps = 0; }, 1500); }
  });
});


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
  const scope  = radio("leg-scope");
  const texto  = document.getElementById("leg-texto").value.trim();

  const statusEl = document.getElementById("status-legendas");
  statusEl.textContent = "Processando...";
  statusEl.className = "status info";
  statusEl.hidden = false;

  try {
    await Word.run(async (context) => {
      const pics = getPics(context, scope);
      pics.load("items");
      await context.sync();

      const n = pics.items.length;
      if (n === 0) throw new Error("Nenhuma imagem encontrada.");

      const nextParas = pics.items.map(pic => pic.paragraph.getNextOrNullObject());
      nextParas.forEach(p => p.load("text"));
      await context.sync();

      const placeholders = [];
      for (let i = n - 1; i >= 0; i--) {
        const next = nextParas[i];
        if (!next.isNullObject && /^(Foto|Figura)\s+\d+/.test(next.text.trim())) continue;
        const ph = pics.items[i].getRange().insertParagraph("__lb__", "After");
        placeholders.push(ph);
      }
      if (placeholders.length === 0) {
        statusEl.textContent = "Todas as imagens já têm legenda.";
        statusEl.className = "status warn";
        return;
      }
      await context.sync();

      let styleName = "Caption";
      try {
        for (const ph of placeholders) { ph.style = "Caption"; }
        await context.sync();
      } catch {
        styleName = "Legenda";
        for (const ph of placeholders) { ph.style = "Legenda"; }
        await context.sync();
      }

      const ooxml = buildPkgOoxml(prefix, "left", texto, styleName);
      for (const ph of placeholders) {
        ph.getRange("Whole").insertOoxml(ooxml, "Replace");
      }
      await context.sync();

      const added = placeholders.length;
      const skipped = n - added;
      const msg = skipped > 0
        ? `✓ ${added} legenda(s) adicionada(s) (${skipped} já tinham).`
        : `✓ ${added} legenda(s) adicionada(s).`;
      statusEl.textContent = msg + " Use o Passo 2 para ajustar o alinhamento.";
      statusEl.className = "status success";
    });

    await Word.run(async (context) => {
      await aplicarKeepNextBody(context);
    });

  } catch (e) {
    statusEl.textContent = "Erro: " + e.message;
    statusEl.className = "status error";
  }
}

async function runLegendasAlign() {
  const align = radio("leg-align");
  const scope = radio("leg-align-scope");

  const statusEl = document.getElementById("status-legendas-align");
  statusEl.textContent = "Processando...";
  statusEl.className = "status info";
  statusEl.hidden = false;

  const pattern = scope === "ambos"
    ? /^(Foto|Figura)\s/i
    : new RegExp(`^${scope.charAt(0).toUpperCase() + scope.slice(1)}\\s`, "i");

  try {
    await Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load("items");
      await context.sync();

      paras.items.forEach(p => p.load("text"));
      await context.sync();

      const targets = paras.items.filter(p => pattern.test(p.text.trim()));
      if (targets.length === 0) {
        statusEl.textContent = "Nenhuma legenda encontrada com esse prefixo.";
        statusEl.className = "status warn";
        return;
      }

      targets.forEach(p => { p.alignment = align; });
      await context.sync();

      await aplicarKeepNextBody(context);

      statusEl.textContent = `✓ Alinhamento aplicado em ${targets.length} legenda(s) + "manter com o próximo" nas fotos.`;
      statusEl.className = "status success";
    });
  } catch (e) {
    statusEl.textContent = "Erro: " + e.message;
    statusEl.className = "status error";
  }
}

// ── Auxiliar: injeta w:keepNext nos parágrafos com foto via OOXML do body ────

async function aplicarKeepNextBody(context) {
  const body = context.document.body;
  const ooxmlResult = body.getOoxml();
  await context.sync();

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const xmlDoc = new DOMParser().parseFromString(ooxmlResult.value, "application/xml");
  const paras = xmlDoc.getElementsByTagNameNS(W, "p");
  let total = 0;

  for (let i = 0; i < paras.length; i++) {
    const p = paras[i];
    const temImagem =
      p.getElementsByTagNameNS(W, "drawing").length > 0 ||
      p.getElementsByTagNameNS(W, "pict").length > 0;
    if (!temImagem) continue;

    let pPr = null;
    for (let j = 0; j < p.childNodes.length; j++) {
      const c = p.childNodes[j];
      if (c.nodeType === 1 && c.namespaceURI === W && c.localName === "pPr") {
        pPr = c; break;
      }
    }
    if (!pPr) {
      pPr = xmlDoc.createElementNS(W, "w:pPr");
      p.insertBefore(pPr, p.firstChild);
    }

    if (pPr.getElementsByTagNameNS(W, "keepNext").length === 0) {
      pPr.appendChild(xmlDoc.createElementNS(W, "w:keepNext"));
    }
    total++;
  }

  if (total > 0) {
    const novoOoxml = new XMLSerializer().serializeToString(xmlDoc);
    body.insertOoxml(novoOoxml, Word.InsertLocation.replace);
    await context.sync();
  }

  return total;
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

function xmlEsc(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// pkg:package OOXML com estilo Caption + campo SEQ (versões e4ae5c5 / 69def05)
function buildPkgOoxml(prefix, jc, texto, styleId = "Caption") {
  const extraRun = texto ? `<w:r><w:t xml:space="preserve"> - ${xmlEsc(texto)}</w:t></w:r>` : "";
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:pPr><w:pStyle w:val="${styleId}"/><w:jc w:val="${jc}"/></w:pPr>
            <w:r><w:t xml:space="preserve">${xmlEsc(prefix)} </w:t></w:r>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> SEQ ${xmlEsc(prefix)} \\* ARABIC </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>1</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
            ${extraRun}
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

// ── Lab: Legendas (fix fotos anteriores) ─────────────────────────────────────

async function runLegendasLab() {
  const prefix = radio("xleg-prefix");
  const scope  = radio("xleg-scope");
  const texto  = document.getElementById("xleg-texto").value.trim();

  const statusEl = document.getElementById("status-legendas-lab");
  statusEl.textContent = "Processando...";
  statusEl.className = "status info";
  statusEl.hidden = false;

  try {
    await Word.run(async (context) => {
      const pics = getPics(context, scope);
      pics.load("items");
      await context.sync();

      const n = pics.items.length;
      if (n === 0) throw new Error("Nenhuma imagem encontrada.");

      // Fase 1: inserir placeholder após TODAS as fotos sem pré-verificação.
      // Usar paragraph.getRange("Whole") garante inserção após o parágrafo inteiro.
      const allPhs = [];
      for (let i = n - 1; i >= 0; i--) {
        const ph = pics.items[i].paragraph.getRange("Whole").insertParagraph("__lb__", "After");
        allPhs.push(ph);
      }
      await context.sync();

      // Fase 1.5: verificar o que vem DEPOIS de cada placeholder.
      // Se a foto já tinha legenda, o placeholder ficou entre a foto e a legenda existente
      // → o próximo parágrafo do placeholder será a legenda → deletar esse placeholder.
      const phNexts = allPhs.map(ph => ph.getNextOrNullObject());
      phNexts.forEach(p => p.load("text"));
      await context.sync();

      const placeholders = [];
      const toDelete = [];
      for (let k = 0; k < allPhs.length; k++) {
        const nextText = phNexts[k].isNullObject ? "" : phNexts[k].text.trim();
        if (/^(Foto|Figura)\s+\d+/.test(nextText)) {
          toDelete.push(allPhs[k]);
        } else {
          placeholders.push(allPhs[k]);
        }
      }

      if (toDelete.length > 0) {
        toDelete.forEach(ph => ph.delete());
        await context.sync();
      }

      if (placeholders.length === 0) {
        statusEl.textContent = "Todas as imagens já têm legenda.";
        statusEl.className = "status warn";
        return;
      }

      // Fase 2: aplicar estilo Caption/Legenda e OOXML com campo SEQ
      let styleName = "Caption";
      try {
        for (const ph of placeholders) { ph.style = "Caption"; }
        await context.sync();
      } catch {
        styleName = "Legenda";
        for (const ph of placeholders) { ph.style = "Legenda"; }
        await context.sync();
      }

      const ooxml = buildPkgOoxml(prefix, "left", texto, styleName);
      for (const ph of placeholders) {
        ph.getRange("Whole").insertOoxml(ooxml, "Replace");
      }
      await context.sync();

      const added = placeholders.length;
      const skipped = n - added;
      const msg = skipped > 0
        ? `✓ ${added} legenda(s) adicionada(s) (${skipped} já tinham).`
        : `✓ ${added} legenda(s) adicionada(s).`;
      statusEl.textContent = msg + " Use o Passo 2 para ajustar o alinhamento.";
      statusEl.className = "status success";
    });

    await Word.run(async (context) => {
      await aplicarKeepNextBody(context);
    });

  } catch (e) {
    statusEl.textContent = "Erro: " + e.message;
    statusEl.className = "status error";
  }
}

async function runLegendasLabAlign() {
  const align = radio("xleg-align");
  const scope = radio("xleg-align-scope");

  const statusEl = document.getElementById("status-legendas-lab-align");
  statusEl.textContent = "Processando...";
  statusEl.className = "status info";
  statusEl.hidden = false;

  const pattern = scope === "ambos"
    ? /^(Foto|Figura)\s/i
    : new RegExp(`^${scope.charAt(0).toUpperCase() + scope.slice(1)}\\s`, "i");

  try {
    await Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load("items");
      await context.sync();

      paras.items.forEach(p => p.load("text"));
      await context.sync();

      const targets = paras.items.filter(p => pattern.test(p.text.trim()));
      if (targets.length === 0) {
        statusEl.textContent = "Nenhuma legenda encontrada com esse prefixo.";
        statusEl.className = "status warn";
        return;
      }

      targets.forEach(p => { p.alignment = align; });
      await context.sync();

      await aplicarKeepNextBody(context);

      statusEl.textContent = `✓ Alinhamento aplicado em ${targets.length} legenda(s) + "manter com o próximo" nas fotos.`;
      statusEl.className = "status success";
    });
  } catch (e) {
    statusEl.textContent = "Erro: " + e.message;
    statusEl.className = "status error";
  }
}

