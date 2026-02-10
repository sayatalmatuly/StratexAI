const API_URL = "https://stratexai-production.up.railway.app/api"; 


let currentUser = null;  
let isGuest = true;
let isGenerating = false;
let selectedFiles = [];   


const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => [...document.querySelectorAll(sel)];

let lastAiPlain = "";

function stripHtml(html) {
  return String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .trim();
}



const LS_CHATS = "stratex_chats_v2";
const LS_ACTIVE = "stratex_active_chat_v2";

function loadChats() {
  try {
    const data = JSON.parse(localStorage.getItem(LS_CHATS)) || [];
    return Array.isArray(data) ? data : [];
  } catch {
    return [];
  }
}

function saveChats(chats) {
  localStorage.setItem(LS_CHATS, JSON.stringify(chats));
}

function getActiveChatId() {
  return localStorage.getItem(LS_ACTIVE) || "";
}

function setActiveChatId(id) {
  localStorage.setItem(LS_ACTIVE, id);
}

function makeId() {
  return crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random();
}

function clampTitle(t) {
  return (t || "").replace(/\s+/g, " ").trim().slice(0, 48);
}

function previewFromText(t) {
  const s = (t || "New chat").replace(/\s+/g, " ").trim();
  return s.length ? s.slice(0, 28) : "New chat";
}

function ensureActiveChat() {
  const chats = loadChats();
  const activeId = getActiveChatId();
  if (activeId && chats.some((c) => c.id === activeId)) return;

  const id = makeId();
  const now = Date.now();
  chats.push({ id, title: "New chat", createdAt: now, updatedAt: now, messages: [] });
  saveChats(chats);
  setActiveChatId(id);
}

function getActiveChat() {
  const id = getActiveChatId();
  return loadChats().find((c) => c.id === id);
}

function updateChatById(chatId, updater) {
  const chats = loadChats();
  const idx = chats.findIndex((c) => c.id === chatId);
  if (idx === -1) return;
  chats[idx] = updater({ ...chats[idx] });
  saveChats(chats);
}

function deleteChatById(chatId) {
  let chats = loadChats();
  const wasActive = getActiveChatId() === chatId;

  chats = chats.filter((c) => c.id !== chatId);
  saveChats(chats);

  if (wasActive) {
    const nextActive = chats.slice().sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0))[0]?.id;
    if (nextActive) {
      setActiveChatId(nextActive);
    } else {
      setActiveChatId("");
      ensureActiveChat();
    }
    loadActiveChatMessagesIntoUI();
  }

  renderChatList();
}

function formatWhen(ts) {
  return new Date(ts || Date.now()).toLocaleString();
}

function scrollToBottom() {
  const chatContainer = document.getElementById("chat-container");
  if (chatContainer) chatContainer.scrollTop = chatContainer.scrollHeight;
}

function escapeHTML(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function tokenizeHTML(input) {
  const s = String(input ?? "");
  const tokens = [];
  let i = 0;
  while (i < s.length) {
    if (s[i] === "<") {
      const end = s.indexOf(">", i);
      if (end === -1) {
        tokens.push({ type: "text", value: s.slice(i) });
        break;
      }
      tokens.push({ type: "tag", value: s.slice(i, end + 1) });
      i = end + 1;
    } else {
      let j = i;
      while (j < s.length && s[j] !== "<") j++;
      tokens.push({ type: "text", value: s.slice(i, j) });
      i = j;
    }
  }
  return tokens;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function typeWriterHTML(html, element) {
  const tokens = tokenizeHTML(html);
  let out = "";
  element.innerHTML = "";

  for (const token of tokens) {
    if (token.type === "tag") {
      out += token.value;
      element.innerHTML = out;
      scrollToBottom();
      await sleep(10 + Math.random() * 20);
    } else {
      const text = token.value;
      for (let i = 0; i < text.length; i++) {
        out += escapeHTML(text[i]);
        element.innerHTML = out;
        scrollToBottom();
        await sleep(12 + Math.random() * 22);
      }
    }
  }
}

function showTypingIndicator() {
  const chatContainer = document.getElementById("chat-container");
  if (!chatContainer) return;

  removeTypingIndicator();

  const indicator = document.createElement("div");
  indicator.id = "typing-indicator";
  indicator.className = "chat-bubble ai";
  indicator.innerHTML = `
    <strong>Stratex AI</strong><br>
    <span style="opacity:0.6;">typing</span>
    <span class="typing-dots"><span>.</span><span>.</span><span>.</span></span>
  `;

  const styleId = "typing-style";
  if (!document.getElementById(styleId)) {
    const style = document.createElement("style");
    style.id = styleId;
    style.innerHTML = `
      .typing-dots span { animation: typing 1.4s infinite; }
      .typing-dots span:nth-child(2) { animation-delay: 0.2s; }
      .typing-dots span:nth-child(3) { animation-delay: 0.4s; }
      @keyframes typing { 0%, 60%, 100% { opacity: 0; } 30% { opacity: 1; } }
      .ai-content strong { color: #fff; }
    `;
    document.head.appendChild(style);
  }

  chatContainer.appendChild(indicator);
  scrollToBottom();
}

function removeTypingIndicator() {
  document.getElementById("typing-indicator")?.remove();
}

function appendMessage(text, sender, meta = {}, animate = false) {
  const chatContainer = document.getElementById("chat-container");
  if (!chatContainer) return;

  const div = document.createElement("div");
  div.classList.add("chat-bubble", sender);

  if (sender === "ai") {
    div.innerHTML = `<strong>Stratex AI</strong><br>`;
    const content = document.createElement("div");
    content.className = "ai-content";
    div.appendChild(content);

    chatContainer.appendChild(div);
    scrollToBottom();

    (async () => {
     if (animate) {
  await typeWriterHTML(String(text ?? ""), content);
} else {
  content.innerHTML = String(text ?? "");
}
      lastAiPlain = stripHtml(String(text ?? ""));
      const sources = meta?.sources;
      if (Array.isArray(sources) && sources.length) {
        const sWrap = document.createElement("div");
        sWrap.style.cssText = "font-size:12px;opacity:.75;border-top:1px solid rgba(255,255,255,.1);padding-top:10px;margin-top:10px;";
        sWrap.innerHTML = `<strong>Sources:</strong><br>`;
        sources.forEach((s) => {
          const a = document.createElement("a");
          a.href = s.url || "#";
          a.target = "_blank";
          a.rel = "noopener noreferrer";
          a.style.cssText = "color:#6f47eb;";
          a.textContent = s.title || s.url || "source";
          sWrap.appendChild(document.createTextNode("📎 "));
          sWrap.appendChild(a);
          sWrap.appendChild(document.createElement("br"));
        });
        content.appendChild(document.createElement("br"));
        content.appendChild(sWrap);
        scrollToBottom();
      }
    })();
  } else {
    div.textContent = String(text ?? "");
    chatContainer.appendChild(div);
    scrollToBottom();
  }
}

function renderChatList() {
  const chatListEl = document.getElementById("chat-list");
  if (!chatListEl) return;

  const activeId = getActiveChatId();
  const chats = loadChats()
    .slice()
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));

  chatListEl.innerHTML = "";

  chats.forEach((chat) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "glass-panel";
    btn.style.background = "transparent";
    btn.style.border = "1px solid rgba(255,255,255,0.05)";
    btn.style.padding = "15px";
    btn.style.textAlign = "left";
    btn.style.cursor = "pointer";
    btn.style.transition = "0.2s";
    btn.style.width = "100%";

    if (chat.id === activeId) {
      btn.style.border = "1px solid rgba(112,71,235,0.6)";
      btn.style.boxShadow = "0 0 0 2px rgba(112,71,235,0.12) inset";
    }

    const wrap = document.createElement("div");
    wrap.style.display = "flex";
    wrap.style.alignItems = "center";
    wrap.style.gap = "10px";
    wrap.style.width = "100%";

    const main = document.createElement("div");
    main.style.flex = "1";
    main.style.minWidth = "0";

    const title = document.createElement("div");
    title.style.fontWeight = "600";
    title.style.color = "white";
    title.style.fontSize = "14px";
    title.style.whiteSpace = "nowrap";
    title.style.overflow = "hidden";
    title.style.textOverflow = "ellipsis";
    title.textContent = `💬 ${chat.title || "Chat"}`;

    const meta = document.createElement("div");
    meta.style.fontSize = "12px";
    meta.style.color = "var(--text-muted)";
    meta.style.marginTop = "4px";
    meta.style.whiteSpace = "nowrap";
    meta.style.overflow = "hidden";
    meta.style.textOverflow = "ellipsis";
    meta.textContent = formatWhen(chat.updatedAt || chat.createdAt);

    main.appendChild(title);
    main.appendChild(meta);

    const actions = document.createElement("div");
    actions.style.display = "flex";
    actions.style.gap = "8px";
    actions.style.opacity = ".85";

    const mkActionBtn = (label, text) => {
      const b = document.createElement("button");
      b.type = "button";
      b.setAttribute("aria-label", label);
      b.textContent = text;
      b.style.width = "28px";
      b.style.height = "28px";
      b.style.borderRadius = "10px";
      b.style.border = "1px solid rgba(255,255,255,.08)";
      b.style.background = "rgba(255,255,255,.03)";
      b.style.color = "white";
      b.style.display = "grid";
      b.style.placeItems = "center";
      b.style.cursor = "pointer";
      b.style.transition = ".2s";
      b.onmouseenter = () => { b.style.background = "rgba(255,255,255,.08)"; b.style.borderColor = "rgba(255,255,255,.18)"; };
      b.onmouseleave = () => { b.style.background = "rgba(255,255,255,.03)"; b.style.borderColor = "rgba(255,255,255,.08)"; };
      return b;
    };

    const rename = mkActionBtn("Rename", "✎");
    const del = mkActionBtn("Delete", "🗑");

    rename.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const current = chat.title || "New chat";
      const nextTitle = clampTitle(prompt("Rename chat:", current));
      if (!nextTitle) return;
      updateChatById(chat.id, (c) => ({ ...c, title: nextTitle, updatedAt: Date.now() }));
      renderChatList();
    });

    del.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const ok = confirm(`Delete "${chat.title || "Chat"}"? This can’t be undone.`);
      if (!ok) return;
      deleteChatById(chat.id);
    });

    actions.appendChild(rename);
    actions.appendChild(del);

    wrap.appendChild(main);
    wrap.appendChild(actions);
    btn.appendChild(wrap);

    btn.addEventListener("click", () => {
      setActiveChatId(chat.id);
      loadActiveChatMessagesIntoUI();
      renderChatList();
    });

    chatListEl.appendChild(btn);
  });
}

function loadActiveChatMessagesIntoUI() {
  const chatContainer = document.getElementById("chat-container");
  if (!chatContainer) return;

  chatContainer.innerHTML = "";
  const chat = getActiveChat();

  if (!chat || !Array.isArray(chat.messages) || chat.messages.length === 0) {
    appendMessage(
      `Hello! I'm ready to help with your business documents.<br><br>
       Try: <strong>"Create a rollout plan"</strong> or attach a <strong>PDF/DOCX/TXT</strong> for analysis.`,
      "ai"
    );
    return;
  }

  chat.messages.forEach((m) => appendMessage(m.text, m.sender, { sources: m.sources || null }));
}

function updateAuthUI() {
  const statusBadge = document.querySelector(".status-badge");
  const loginBtn = document.getElementById("login-btn");
  const logoutBtn = document.getElementById("logout-btn");

  if (currentUser) {
    if (statusBadge) {
      statusBadge.innerHTML = `<span class="status-dot"></span> ${currentUser.username}`;
    }
    if (loginBtn) loginBtn.style.display = "none";
    if (logoutBtn) logoutBtn.style.display = "inline-flex";
  } else {
    if (statusBadge) {
      statusBadge.innerHTML = `<span class="status-dot"></span> Guest`;
    }
    if (loginBtn) loginBtn.style.display = "inline-flex";
    if (logoutBtn) logoutBtn.style.display = "none";
  }
}

function showAuthModal() {
  const modal = document.createElement("div");
  modal.style.cssText = `
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.8);
    display: flex; align-items: center; justify-content: center;
    z-index: 10000; backdrop-filter: blur(10px);
  `;

  modal.innerHTML = `
    <div class="glass-panel" style="padding: 40px; max-width: 420px; width: 92%; border-radius: 20px;">
      <h2 style="margin-bottom: 10px; background: linear-gradient(135deg, #6f47eb, #a855f7);
        -webkit-background-clip:text; -webkit-text-fill-color: transparent;">🔐 Login</h2>
      <p style="color: var(--text-muted); margin-bottom: 26px; font-size: 14px;">
        Sign in to unlock documents + full AI features
      </p>

      <div style="margin-bottom: 16px;">
        <label style="display:block;margin-bottom:8px;color:white;font-weight:500;">Username</label>
        <input id="auth-username" type="text" placeholder="demo"
          style="width:100%;padding:12px;background:rgba(0,0,0,0.3);border:1px solid var(--glass-border);
          border-radius:8px;color:white;font-family:inherit;">
      </div>

      <div style="margin-bottom: 22px;">
        <label style="display:block;margin-bottom:8px;color:white;font-weight:500;">Password</label>
        <input id="auth-password" type="password" placeholder="••••••"
          style="width:100%;padding:12px;background:rgba(0,0,0,0.3);border:1px solid var(--glass-border);
          border-radius:8px;color:white;font-family:inherit;">
      </div>

      <div style="display:flex; gap:10px;">
        <button id="auth-submit" class="btn btn--primary" style="flex:1;">Sign in</button>
        <button id="auth-cancel" class="btn btn--glass">Cancel</button>
      </div>

      <div style="margin-top: 18px; padding: 14px; background: rgba(111,71,235,0.10); border-radius: 10px;
        border: 1px solid rgba(111,71,235,0.25);">
        <p style="font-size: 12px; color: var(--text-muted); margin: 0;">
          <strong>Test accounts:</strong><br>
          demo / demo123<br>
          business / business2024
        </p>
      </div>
    </div>
  `;

  document.body.appendChild(modal);

  const close = () => modal.remove();
  modal.querySelector("#auth-cancel").onclick = close;

  const doAuth = async () => {
    const username = modal.querySelector("#auth-username").value.trim();
    const password = modal.querySelector("#auth-password").value.trim();
    if (!username || !password) {
      alert("Fill in all fields");
      return;
    }

    try {
      const r = await fetch(`${API_URL}/auth`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });

      if (!r.ok) {
        alert("Wrong login/password");
        return;
      }

      const data = await r.json();
      if (data?.success) {
        if (data.user?.username) {
          currentUser = data.user;
          isGuest = false;
          persistAuth();
          updateAuthUI();
        }

        const msg =
          data.message ||
          (data.user?.username
            ? `Logged in as <strong>${data.user.username}</strong>.`
            : "Logged in.");
        appendMessage(msg, "ai", { sources: data.sources || null }, true);
        close();
      } else {
        alert("Wrong login/password");
      }
    } catch (e) {
      alert("Cannot reach server. Is backend running?");
      console.error(e);
    }
  };

  modal.querySelector("#auth-submit").onclick = doAuth;
  modal.querySelector("#auth-password").addEventListener("keypress", (e) => {
    if (e.key === "Enter") doAuth();
  });
}

function persistAuth() {
  try {
    localStorage.setItem("stratex_auth_v1", JSON.stringify({ currentUser, isGuest }));
  } catch {}
}

function restoreAuth() {
  try {
    const raw = localStorage.getItem("stratex_auth_v1");
    if (!raw) return;
    const data = JSON.parse(raw);
    if (data?.currentUser?.username) {
      currentUser = data.currentUser;
      isGuest = false;
    }
  } catch {}
}

function getServerUserIdForActiveChat() {
  const chatId = getActiveChatId() || "default";
  const u = currentUser?.username ? currentUser.username : "guest";
  return `${u}::${chatId}`; 
}

async function apiChat(message) {
  const payload = {
    message,
    user_id: getServerUserIdForActiveChat(),
    is_guest: isGuest,
  };

  const r = await fetch(`${API_URL}/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const data = await r.json();
  return data;
}

async function apiAnalyzeDocument(file) {
  const form = new FormData();
  form.append("file", file);
  form.append("user_id", getServerUserIdForActiveChat());
  form.append("is_guest", String(isGuest));

  const r = await fetch(`${API_URL}/analyze-document`, { method: "POST", body: form });
  const data = await r.json();
  return { ok: r.ok, data };
}

function localDemoResponse(text) {
  const t = String(text || "").toLowerCase();
  if (t.includes("plan")) {
    return "Here is a <strong>Strategic Plan</strong>:<br><br>1. <strong>Analyze:</strong> Review metrics.<br>2. <strong>Execute:</strong> Launch pilot in 2 weeks.<br>3. <strong>Review:</strong> Measure ROI by month-end.";
  }
  if (t.includes("email")) {
    return "Subject: Update on Project Alpha<br><br>Hi Team,<br><br>We are on track for the Friday release. QA testing is 90% complete. No blockers found.<br><br>Best,<br>[Your Name]";
  }
  return "I’ve prepared a concise executive-style response based on your request.";
}

async function handleSend() {
  const chatInput = document.querySelector(".input-box");
  const text = (chatInput?.value || "").trim();

  if (!text && (!selectedFiles || selectedFiles.length === 0)) return;
  if (isGenerating) return;

  const activeId = getActiveChatId();
  const now = Date.now();

  if (text) {
    appendMessage(text, "user");
    updateChatById(activeId, (chat) => {
      const next = { ...chat };
      next.updatedAt = now;
      next.messages = Array.isArray(next.messages) ? next.messages.slice() : [];
      next.messages.push({ sender: "user", text, at: now });
      if (!next.title || next.title === "New chat") next.title = previewFromText(text);
      return next;
    });
  }

  const filesToSend = selectedFiles.slice();
  if (filesToSend.length) {
    appendMessage(`📎 Attached: ${filesToSend.map((f) => f.name).join(", ")}`, "user");
    updateChatById(activeId, (chat) => {
      const next = { ...chat };
      next.updatedAt = Date.now();
      next.messages = Array.isArray(next.messages) ? next.messages.slice() : [];
      next.messages.push({
        sender: "user",
        text: `📎 Attached: ${filesToSend.map((f) => f.name).join(", ")}`,
        at: Date.now(),
      });
      return next;
    });
  }

  chatInput.value = "";
  selectedFiles = [];
  updateAttachButtonLabel();

  renderChatList();

  isGenerating = true;
  showTypingIndicator();

  try {
    let analyzedSummaries = [];
    if (filesToSend.length) {
      for (const f of filesToSend) {
        const res = await apiAnalyzeDocument(f);

        if (res.data?.require_auth) {
          removeTypingIndicator();
          appendMessage(res.data.message || "Document analysis requires login.", "ai");
          if (confirm("Document analysis requires login. Login now?")) showAuthModal();
          isGenerating = false;
          renderChatList();
          return;
        }

        if (res.ok && res.data?.success) {
          analyzedSummaries.push(`Document "${res.data.filename}":\n${res.data.analysis}`);
        } else {
          analyzedSummaries.push(`Document "${f.name}": ❌ ${res.data?.error || "failed"}`);
        }
      }
    }

    let finalMessage = text || "";
    if (analyzedSummaries.length) {
      finalMessage =
        (finalMessage ? finalMessage + "\n\n" : "") +
        "Please use these document analyses in your answer:\n\n" +
        analyzedSummaries.join("\n\n---\n\n");
    }

    const data = await apiChat(finalMessage);

    removeTypingIndicator();

    if (data?.success) {
      if (data.require_auth) {
        appendMessage(data.message || "Login required.", "ai");
        setTimeout(() => {
          if (confirm("Want to login for full access?")) showAuthModal();
        }, 600);
      } else {
        appendMessage(
  data.message,
  "ai",
  { sources: data.sources || null },
  true
);

        updateChatById(activeId, (chat) => {
          const now2 = Date.now();
          const next = { ...chat };
          next.updatedAt = now2;
          next.messages = Array.isArray(next.messages) ? next.messages.slice() : [];
          next.messages.push({ sender: "ai", text: data.message, at: now2, sources: data.sources || null });
          return next;
        });
      }
    } else {
      const fallback = localDemoResponse(text || "analyze documents");
      appendMessage(`⚠️ Backend error. Using local demo.\n\n${fallback}`.replaceAll("\n", "<br>"), "ai");
    }
  } catch (e) {
    removeTypingIndicator();

    const fallback = localDemoResponse(text || "analyze documents");
    appendMessage(
      `⚠️ Cannot reach backend (<code>${API_URL}</code>). Using local demo:<br><br>${fallback}`,
      "ai"
    );

    console.error(e);
  } finally {
    isGenerating = false;
    renderChatList();
  }
}


function updateAttachButtonLabel() {
  const btn = document.getElementById("attach-btn");
  if (!btn) return;
  if (!selectedFiles.length) btn.textContent = "📎";
  else btn.textContent = `📎 ${selectedFiles.length}`;
}

function getFilenameFromDisposition(disposition) {
  if (!disposition) return "";
  const match = /filename\*?=(?:UTF-8''|")?([^\";]+)/i.exec(disposition);
  if (!match) return "";
  return decodeURIComponent(match[1].trim().replace(/(^\"|\"$)/g, ""));
}

async function downloadFileFromLink(href) {
  try {
    const r = await fetch(href, { mode: "cors", credentials: "include" });
    if (!r.ok) throw new Error(`Download failed: ${r.status}`);

    const ct = (r.headers.get("content-type") || "").toLowerCase();
    if (ct.includes("text/html") || ct.includes("application/json")) {
      throw new Error(`Not a file response: ${ct}`);
    }

    const blob = await r.blob();
    const url = URL.createObjectURL(blob);

    const disp = r.headers.get("content-disposition") || "";
    const headerName = getFilenameFromDisposition(disp);
    const fallbackName = decodeURIComponent(href.split("/").pop() || "download");
    const filename = headerName || fallbackName;

    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (err) {
    console.error(err);
    window.open(href, "_blank", "noopener");
  }
}

function setupFileAttach() {
  const attachBtn = document.getElementById("attach-btn");
  const fileInput = document.getElementById("file-input");
  if (!attachBtn || !fileInput) return;

  attachBtn.addEventListener("click", (e) => {
    e.preventDefault();
    fileInput.click();
  });

  fileInput.addEventListener("change", () => {
    const files = [...(fileInput.files || [])];
    if (!files.length) return;
    selectedFiles = files;
    updateAttachButtonLabel();
  });
}

function setupDragDrop() {
  const chatContainer = document.getElementById("chat-container");
  if (!chatContainer) return;

  chatContainer.addEventListener("dragover", (e) => {
    e.preventDefault();
    chatContainer.style.background = "rgba(111,71,235,0.10)";
  });

  chatContainer.addEventListener("dragleave", () => {
    chatContainer.style.background = "";
  });

  chatContainer.addEventListener("drop", (e) => {
    e.preventDefault();
    chatContainer.style.background = "";

    const files = [...(e.dataTransfer?.files || [])];
    if (!files.length) return;

    selectedFiles = files.filter((f) => /\.(pdf|docx|txt)$/i.test(f.name));
    updateAttachButtonLabel();

    handleSend();
  });
}

function exportActiveChat() {
  const chat = getActiveChat();
  if (!chat) return;

  const lines = [];
  lines.push(`Stratex AI Export`);
  lines.push(`Title: ${chat.title || "Chat"}`);
  lines.push(`When: ${new Date(chat.updatedAt || chat.createdAt || Date.now()).toLocaleString()}`);
  lines.push(`User: ${currentUser?.username || "guest"}`);
  lines.push("");
  lines.push("----");
  lines.push("");

  (chat.messages || []).forEach((m) => {
    const who = m.sender === "ai" ? "AI" : "You";
    const time = new Date(m.at || Date.now()).toLocaleString();
    const txt = String(m.text || "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/?[^>]+>/g, "");
    lines.push(`[${time}] ${who}:`);
    lines.push(txt);
    lines.push("");
  });

  const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${(chat.title || "stratex-chat").replace(/[^\w\-]+/g, "_")}.txt`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
function logout() {
  if (!confirm("Log out from Stratex AI?")) return;

  currentUser = null;
  isGuest = true;

  localStorage.removeItem("stratex_auth_v1");

  updateAuthUI();


  appendMessage(
    "You have logged out. You are now using Stratex AI as a guest.",
    "ai"
  );
}


document.addEventListener("DOMContentLoaded", () => {

  restoreAuth();
  updateAuthUI();

  ensureActiveChat();
  renderChatList();
  loadActiveChatMessagesIntoUI();

  const sendBtn = document.querySelector(".send-btn") || document.querySelector(".composer .btn--primary");
  const chatInput = document.querySelector(".input-box");

  sendBtn?.addEventListener("click", (e) => { e.preventDefault(); handleSend(); });

  chatInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });

  chatInput?.addEventListener("input", () => {
    chatInput.style.height = "auto";
    chatInput.style.height = chatInput.scrollHeight + "px";
  });

  document.getElementById("new-chat")?.addEventListener("click", () => {
    const chats = loadChats();
    const id = makeId();
    const now = Date.now();
    chats.push({ id, title: "New chat", createdAt: now, updatedAt: now, messages: [] });
    saveChats(chats);
    setActiveChatId(id);
    renderChatList();
    loadActiveChatMessagesIntoUI();
  });

  document.getElementById("login-btn")?.addEventListener("click", (e) => {
    e.preventDefault();
    showAuthModal();
  });

  document.getElementById("logout-btn")?.addEventListener("click", (e) => {
    e.preventDefault();
    logout();
  });

  document.getElementById("export-btn")?.addEventListener("click", (e) => {
    e.preventDefault();
    exportActiveChat();
  });

  const genBtn = document.getElementById("gen-file-btn");
  const genType = document.getElementById("gen-file-type");

  genBtn?.addEventListener("click", async (e) => {
    e.preventDefault();
    const content = (lastAiPlain || "").trim();
    if (!content) {
      alert("No AI answer to export yet.");
      return;
    }

    try {
      const r = await fetch(`${API_URL}/generate-file`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          content,
          file_type: genType?.value || "pdf",
        }),
      });

      const data = await r.json();
      if (!r.ok || !data?.success) {
        alert(data?.error || "Failed to generate file");
        return;
      }

      if (data?.download_url) {
        const url = `${API_URL.replace('/api','')}${data.download_url}`;
        appendMessage(
          `File ready: <a href="${url}" target="_blank" rel="noopener" style="color:#6f47eb;font-weight:600;">Download</a>`,
          "ai"
        );
      }
    } catch (err) {
      console.error(err);
      alert("Server error while generating file");
    }
  });

  $$(".prompt-card").forEach((card) => {
    card.addEventListener("click", () => {
      const prompt = card.getAttribute("data-prompt") || "";
      if (!prompt) return;
      if (chatInput) chatInput.value = prompt;
      handleSend();
    });
  });

  setupFileAttach();
  setupDragDrop();

  const chatContainer = document.getElementById("chat-container");
  chatContainer?.addEventListener("click", (e) => {
    const a = e.target?.closest?.("a");
    if (!a) return;
    const href = a.getAttribute("href");
    if (!href) return;
    if (!href.includes("/api/download/")) return;
    e.preventDefault();
    downloadFileFromLink(href);
  });

  fetch(`${API_URL}/status`)
    .then((r) => r.json())
    .then((data) => {
      console.log("✅ Backend status:", data);
    })
    .catch(() => {
      appendMessage(
        `⚠️ Backend is not reachable. If you want real AI + documents, run the Python server on <code>http://localhost:5000</code>.`,
        "ai"
      );
    });
});
const burger = document.getElementById("burger");
const navLinks = document.querySelector(".nav__links");

if (burger && navLinks) {
  const closeMenu = () => {
    navLinks.classList.remove("is-open");
    burger.classList.remove("is-open");
    document.body.classList.remove("menu-open");
  };

  burger.addEventListener("click", () => {
    const isOpen = navLinks.classList.toggle("is-open");
    burger.classList.toggle("is-open", isOpen);
    document.body.classList.toggle("menu-open", isOpen);
  });

  navLinks.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => {
      if (window.innerWidth <= 900) closeMenu();
    });
  });

  window.addEventListener("resize", () => {
    if (window.innerWidth > 900) closeMenu();
  });
}




const mobileMenuBtn = document.getElementById("mobile-menu-toggle");
const sidebar = document.querySelector(".sidebar");
const sidebarBackdrop = document.getElementById("sidebar-backdrop");

if (mobileMenuBtn && sidebar && sidebarBackdrop) {
  const closeSidebar = () => {
    sidebar.classList.remove("is-open");
    sidebarBackdrop.classList.remove("show");
    document.body.classList.remove("menu-open");
  };

  mobileMenuBtn.addEventListener("click", () => {
    const isOpen = sidebar.classList.toggle("is-open");
    sidebarBackdrop.classList.toggle("show", isOpen);
    document.body.classList.toggle("menu-open", isOpen);
  });

  sidebarBackdrop.addEventListener("click", closeSidebar);

  window.addEventListener("resize", () => {
    if (window.innerWidth > 800) closeSidebar();
  });
}
