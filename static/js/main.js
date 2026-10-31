/* =========================================================
   RosterLine — main.js
   ========================================================= */

/* ---------- Timeline "now" marker ---------- */
function positionNowMarker() {
  const marker = document.getElementById("timeline-now");
  if (!marker) return;
  const now = new Date();
  const percent = ((now.getHours() * 60 + now.getMinutes()) / 1440) * 100;
  marker.style.left = percent + "%";
}
positionNowMarker();
setInterval(positionNowMarker, 60000);


/* ---------- Settings page — capacity stepper ---------- */
function stepCapacity(inputId, delta) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.value = Math.max(1, (parseInt(input.value, 10) || 1) + delta);
  updateTotalCapacity();
}

function updateTotalCapacity() {
  const totalEl = document.getElementById("total-capacity");
  if (!totalEl) return;
  let sum = 0;
  document.querySelectorAll(".capacity-input").forEach(inp => {
    sum += parseInt(inp.value, 10) || 0;
  });
  totalEl.textContent = sum;
}

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll(".capacity-input").forEach(inp => {
    inp.addEventListener("input", updateTotalCapacity);
  });
});


/* ---------- Toast ---------- */
let _toastTimer;
function showToast(msg, type = "ok") {
  const el = document.getElementById("toast");
  if (!el) return;
  el.textContent = msg;
  el.className = "toast toast-" + type + " toast-show";
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { el.className = "toast"; }, 2400);
}


/* ---------- Shift dropdown — live save ---------- */
async function handleShiftChange(select) {
  const workerId  = select.dataset.workerId;
  const day       = select.dataset.day;
  const shiftCode = select.value;

  // Update colour class immediately
  select.className = "shift-select shift-select-" + shiftCode.toLowerCase();

  try {
    const res = await fetch("/assign", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ worker_id: workerId, day, shift_code: shiftCode })
    });
    const data = await res.json();
    if (data.ok) {
      showToast("Saved", "ok");
    } else {
      showToast(data.error || "Error saving", "err");
    }
  } catch (e) {
    showToast("Network error", "err");
  }
}


/* ---------- Export helpers ---------- */

/** Returns a promise that resolves when the CDN script is loaded. */
function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error("Failed to load " + src));
    document.head.appendChild(s);
  });
}

/** Download the weekly grid as a PNG image. */
async function downloadImage() {
  const grid = document.getElementById("weekly-grid");
  if (!grid) return;
  showToast("Generating image…", "ok");
  try {
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
    const canvas = await html2canvas(grid, {
      backgroundColor: "#10131f",
      scale: 2,
      useCORS: true,
      logging: false
    });
    const link = document.createElement("a");
    link.download = "weekly_schedule.png";
    link.href = canvas.toDataURL("image/png");
    link.click();
    showToast("Image downloaded", "ok");
  } catch (e) {
    showToast("Image export failed", "err");
    console.error(e);
  }
}

/** Download the weekly grid as a PDF. */
async function downloadPDF() {
  const grid = document.getElementById("weekly-grid");
  if (!grid) return;
  showToast("Generating PDF…", "ok");
  try {
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js");

    const canvas = await html2canvas(grid, {
      backgroundColor: "#10131f",
      scale: 2,
      useCORS: true,
      logging: false
    });

    const { jsPDF } = window.jspdf;
    const imgW  = canvas.width;
    const imgH  = canvas.height;
    const ratio = imgH / imgW;

    // A4 landscape
    const pdf   = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    const pdfW  = pdf.internal.pageSize.getWidth();
    const pdfH  = pdfW * ratio;

    pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 0, pdfW, pdfH);
    pdf.save("weekly_schedule.pdf");
    showToast("PDF downloaded", "ok");
  } catch (e) {
    showToast("PDF export failed", "err");
    console.error(e);
  }
}


/* ================================================================
   AI SHIFT ASSISTANT (OPENROUTER CHATBOT)
   ================================================================ */

let aiChatHistory = [];
let aiIsThinking = false;

/** Toggle the slide-over AI Chat drawer. */
function toggleAIChat() {
  const drawer = document.getElementById("ai-chat-drawer");
  const overlay = document.getElementById("ai-drawer-overlay");
  if (!drawer) return;

  const isOpen = drawer.classList.contains("open");
  if (isOpen) {
    drawer.classList.remove("open");
    if (overlay) overlay.classList.remove("active");
  } else {
    drawer.classList.add("open");
    if (overlay) overlay.classList.add("active");
    const input = document.getElementById("ai-user-input");
    if (input) setTimeout(() => input.focus(), 200);
  }
}

/** Toggle the in-drawer settings panel for OpenRouter. */
function toggleAISettings() {
  const panel = document.getElementById("ai-settings-panel");
  if (!panel) return;
  const isHidden = panel.style.display === "none";
  panel.style.display = isHidden ? "block" : "none";
}

/** Toggle visibility of the API key input. */
function toggleKeyVisibility() {
  const input = document.getElementById("ai-key-input");
  if (!input) return;
  input.type = input.type === "password" ? "text" : "password";
}

/** Save OpenRouter settings from drawer. */
async function saveAISettings() {
  const keyInput = document.getElementById("ai-key-input");
  const modelSelect = document.getElementById("ai-model-select");
  const key = keyInput ? keyInput.value.trim() : "";
  const model = modelSelect ? modelSelect.value.trim() : "";

  try {
    const res = await fetch("/api/openrouter-config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ api_key: key, model })
    });
    const data = await res.json();
    if (data.ok) {
      showToast("AI settings saved", "ok");
      const badge = document.getElementById("ai-active-model-badge");
      if (badge && model) {
        badge.textContent = model.split("/").pop();
      }
      toggleAISettings();
    } else {
      showToast(data.error || "Failed to save settings", "err");
    }
  } catch (e) {
    showToast("Network error saving settings", "err");
  }
}

/** Clear the conversation history in the drawer. */
function clearAIChat() {
  aiChatHistory = [];
  const container = document.getElementById("ai-chat-messages");
  if (!container) return;
  container.innerHTML = `
    <div class="ai-msg ai-msg-assistant">
      <div class="ai-msg-avatar">✨</div>
      <div class="ai-msg-content">
        <div class="ai-bubble">
          <p><strong>Chat history cleared.</strong></p>
          <p>How can I assist you with the shift roster? You can ask to generate shifts, enforce constraints, or modify the table.</p>
        </div>
        <span class="ai-msg-time">Just now</span>
      </div>
    </div>
  `;
  showToast("Chat cleared", "ok");
}

/** Auto-resize input textarea as user types. */
function autoResizeTextarea(textarea) {
  textarea.style.height = "auto";
  textarea.style.height = Math.min(textarea.scrollHeight, 120) + "px";
}

/** Handle Enter key in chat input (Enter to send, Shift+Enter for newline). */
function handleChatKey(event) {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    sendChatMessage();
  }
}

/** Trigger quick action chip prompt. */
function sendQuickPrompt(promptText) {
  const input = document.getElementById("ai-user-input");
  if (!input) return;
  input.value = promptText;
  autoResizeTextarea(input);
  sendChatMessage();
}

/** Escape HTML characters safely. */
function escapeHTML(str) {
  return str.replace(/[&<>'"]/g, tag => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;'
  }[tag] || tag));
}

/** Lightweight markdown formatter for chat bubbles. */
function formatMarkdown(text) {
  if (!text) return "";
  let html = escapeHTML(text);

  // Bold
  html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  // Italics
  html = html.replace(/\*(.+?)\*/g, "<em>$1</em>");
  // Inline code
  html = html.replace(/`([^`]+)`/g, "<code>$1</code>");
  // Convert list items
  html = html.replace(/^[*-]\s+(.+)$/gm, "<li>$1</li>");
  // Wrap li in ul
  html = html.replace(/(<li>.*<\/li>)/s, "<ul>$1</ul>");
  // Line breaks
  html = html.replace(/\n\n/g, "</p><p>");
  html = html.replace(/\n/g, "<br>");
  return `<p>${html}</p>`;
}

/** Send user chat message to /api/chat and handle AI response. */
async function sendChatMessage() {
  if (aiIsThinking) return;

  const input = document.getElementById("ai-user-input");
  if (!input) return;
  const message = input.value.trim();
  if (!message) return;

  const messagesContainer = document.getElementById("ai-chat-messages");
  const sendBtn = document.getElementById("ai-send-btn");

  // 1. Add User message bubble
  const nowStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  const userMsgEl = document.createElement("div");
  userMsgEl.className = "ai-msg ai-msg-user";
  userMsgEl.innerHTML = `
    <div class="ai-msg-content">
      <div class="ai-bubble">${escapeHTML(message).replace(/\n/g, "<br>")}</div>
      <span class="ai-msg-time">${nowStr}</span>
    </div>
  `;
  messagesContainer.appendChild(userMsgEl);
  aiChatHistory.push({ role: "user", content: message });

  // Clear input
  input.value = "";
  autoResizeTextarea(input);
  messagesContainer.scrollTop = messagesContainer.scrollHeight;

  // 2. Add typing indicator
  aiIsThinking = true;
  if (sendBtn) sendBtn.disabled = true;

  const typingEl = document.createElement("div");
  typingEl.className = "ai-msg ai-msg-assistant ai-typing-msg";
  typingEl.id = "ai-typing-indicator";
  typingEl.innerHTML = `
    <div class="ai-msg-avatar">✨</div>
    <div class="ai-msg-content">
      <div class="ai-bubble ai-typing">
        <span></span><span></span><span></span>
      </div>
    </div>
  `;
  messagesContainer.appendChild(typingEl);
  messagesContainer.scrollTop = messagesContainer.scrollHeight;

  // 3. Make API request to /api/chat
  try {
    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        message,
        history: aiChatHistory
      })
    });
    const data = await res.json();

    // Remove typing indicator
    const typingIndicator = document.getElementById("ai-typing-indicator");
    if (typingIndicator) typingIndicator.remove();

    if (data.ok) {
      const assistantMsgEl = document.createElement("div");
      assistantMsgEl.className = "ai-msg ai-msg-assistant";

      let bubbleContent = formatMarkdown(data.reply);

      // If shifts were modified by the AI, add badge
      if (data.actions_count && data.actions_count > 0) {
        bubbleContent += `
          <div class="ai-action-badge">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            Applied ${data.actions_count} shift update(s) to schedule
          </div>
        `;

        // Update live table cells in DOM!
        if (data.actions && Array.isArray(data.actions)) {
          data.actions.forEach(act => {
            const sel = document.querySelector(`.shift-select[data-worker-id="${act.worker_id}"][data-day="${act.day}"]`);
            if (sel) {
              const code = act.shift_code || "—";
              sel.value = code;
              sel.className = `shift-select shift-select-${code.toLowerCase()} shift-updated-flash`;
              setTimeout(() => sel.classList.remove("shift-updated-flash"), 2500);
            }
          });
          showToast(`Updated ${data.actions.length} shift(s)`, "ok");
        } else if (data.updated_table && Array.isArray(data.updated_table)) {
          // Fallback to table data
          data.updated_table.forEach(row => {
            Object.entries(row.days).forEach(([d, code]) => {
              const sel = document.querySelector(`.shift-select[data-worker-id="${row.worker_id}"][data-day="${d}"]`);
              if (sel) {
                sel.value = code;
                sel.className = `shift-select shift-select-${code.toLowerCase()} shift-updated-flash`;
                setTimeout(() => sel.classList.remove("shift-updated-flash"), 2500);
              }
            });
          });
          showToast("Schedule updated", "ok");
        }
      }

      const resTimeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      assistantMsgEl.innerHTML = `
        <div class="ai-msg-avatar">✨</div>
        <div class="ai-msg-content">
          <div class="ai-bubble">${bubbleContent}</div>
          <span class="ai-msg-time">${resTimeStr}</span>
        </div>
      `;
      messagesContainer.appendChild(assistantMsgEl);
      aiChatHistory.push({ role: "assistant", content: data.reply });

      // If user needs API key, open settings automatically
      if (data.needs_key) {
        const panel = document.getElementById("ai-settings-panel");
        if (panel && panel.style.display === "none") {
          panel.style.display = "block";
        }
      }
    } else {
      const errEl = document.createElement("div");
      errEl.className = "ai-msg ai-msg-assistant";
      errEl.innerHTML = `
        <div class="ai-msg-avatar" style="background:#ef4444;">⚠️</div>
        <div class="ai-msg-content">
          <div class="ai-bubble" style="border-color:#fca5a5; background:#fef2f2; color:#991b1b;">
            <strong>Error:</strong> ${escapeHTML(data.error || "An error occurred with OpenRouter.")}
          </div>
          <span class="ai-msg-time">Just now</span>
        </div>
      `;
      messagesContainer.appendChild(errEl);
      showToast(data.error || "AI Error", "err");
    }
  } catch (err) {
    const typingIndicator = document.getElementById("ai-typing-indicator");
    if (typingIndicator) typingIndicator.remove();

    const netErrEl = document.createElement("div");
    netErrEl.className = "ai-msg ai-msg-assistant";
    netErrEl.innerHTML = `
      <div class="ai-msg-avatar" style="background:#ef4444;">⚠️</div>
      <div class="ai-msg-content">
        <div class="ai-bubble" style="border-color:#fca5a5; background:#fef2f2; color:#991b1b;">
          <strong>Network Error:</strong> Could not reach the server. Please check your connection.
        </div>
        <span class="ai-msg-time">Just now</span>
      </div>
    `;
    messagesContainer.appendChild(netErrEl);
    showToast("Network error", "err");
    console.error(err);
  } finally {
    aiIsThinking = false;
    if (sendBtn) sendBtn.disabled = false;
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
  }
}

// On page load: fetch current OpenRouter configuration status
document.addEventListener("DOMContentLoaded", async () => {
  try {
    const res = await fetch("/api/openrouter-config");
    const data = await res.json();
    if (data) {
      const keyInput = document.getElementById("ai-key-input");
      if (keyInput && data.key_preview) {
        keyInput.placeholder = data.key_preview;
      }
      const modelBadge = document.getElementById("ai-active-model-badge");
      if (modelBadge && data.model) {
        modelBadge.textContent = data.model.split("/").pop();
      }
      const modelSelect = document.getElementById("ai-model-select");
      if (modelSelect && data.model) {
        modelSelect.value = data.model;
      }
    }
  } catch (e) {
    // ignore
  }
});
