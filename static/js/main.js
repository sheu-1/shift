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
