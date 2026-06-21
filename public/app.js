const form = document.getElementById("checkForm");
const planInput = document.getElementById("planFile");
const actualInput = document.getElementById("actualFile");
const previewBtn = document.getElementById("previewBtn");
const downloadBtn = document.getElementById("downloadBtn");
const downloadPptxBtn = document.getElementById("downloadPptxBtn");
const statusFilter = document.getElementById("statusFilter");
const errorMsg = document.getElementById("errorMsg");
const resultsSection = document.getElementById("resultsSection");
const resultsBody = document.getElementById("resultsBody");
const summaryRow = document.getElementById("summaryRow");
const loadingOverlay = document.getElementById("loadingOverlay");
const loadingText = document.getElementById("loadingText");

// Pastikan overlay & error selalu tertutup saat halaman baru dimuat
// (mencegah sisa state dari bfcache / form auto-restore browser).
loadingOverlay.hidden = true;
errorMsg.hidden = true;
errorMsg.textContent = "";

window.addEventListener("pageshow", () => {
  loadingOverlay.hidden = true;
  previewBtn.disabled = false;
  downloadBtn.disabled = false;
});

function bindDropfield(input) {
  const field = input.closest(".dropfield");
  const textEl = field.querySelector(".dz-text");
  input.addEventListener("change", () => {
    if (input.files && input.files[0]) {
      field.classList.add("has-file");
      textEl.textContent = input.files[0].name;
    } else {
      field.classList.remove("has-file");
      textEl.textContent = textEl.dataset.default;
    }
  });
}
bindDropfield(planInput);
bindDropfield(actualInput);

function showError(message) {
  errorMsg.textContent = message;
  errorMsg.hidden = false;
}

function hideError() {
  errorMsg.hidden = true;
  errorMsg.textContent = "";
}

function setLoading(isLoading, text) {
  loadingOverlay.hidden = !isLoading;
  if (text) loadingText.textContent = text;
  previewBtn.disabled = isLoading;
  downloadBtn.disabled = isLoading;
  downloadPptxBtn.disabled = isLoading;
}

function validateFiles() {
  if (!planInput.files[0] || !actualInput.files[0]) {
    showError("Mohon pilih kedua file (Plan dan Actual) sebelum lanjut.");
    return false;
  }
  return true;
}

function buildFormData() {
  const fd = new FormData();
  fd.append("plan_file", planInput.files[0]);
  fd.append("actual_file", actualInput.files[0]);
  fd.append("status_filter", statusFilter.value);
  return fd;
}

function statusPillHtml(status) {
  const map = {
    Selesai: ["status-selesai", "Ditemukan"],
    Belum: ["status-belum", "Belum"],
    "Tidak Ditemukan": ["status-tidak", "Tidak Ditemukan"],
  };
  const [cls, label] = map[status] || ["status-belum", status];
  return `<span class="status-pill ${cls}">${label}</span>`;
}

function renderResults(data) {
  resultsBody.innerHTML = "";

  if (data.rows.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="4" style="text-align:center; color:var(--text-dim); padding:24px;">Tidak ada data untuk status yang dipilih.</td>`;
    resultsBody.appendChild(tr);
  }

  data.rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row["tanggal"]}</td>
      <td>${row["start jam"]}</td>
      <td>${row["brand"]}</td>
      <td>${statusPillHtml(row["status"])}</td>
    `;
    resultsBody.appendChild(tr);
  });

  summaryRow.innerHTML = "";
  const totalCard = document.createElement("div");
  totalCard.className = "summary-card";
  totalCard.innerHTML = `<div class="sc-label">Total Baris</div><div class="sc-value">${data.total_rows}</div>`;
  summaryRow.appendChild(totalCard);

  data.summary.forEach((row) => {
    const card = document.createElement("div");
    card.className = "summary-card";
    card.innerHTML = `
      <div class="sc-label">${row.brand}</div>
      <div class="sc-value" style="font-size:16px; display:flex; gap:10px;">
        <span style="color:green">${row["Selesai"] ?? 0} Ditemukan</span>
        <span style="color:amber">${row["Belum"] ?? 0} Belum</span>
        <span style="color:red">${row["Tidak Ditemukan"] ?? 0} Tidak Ditemukan</span>
      </div>
    `;
    summaryRow.appendChild(card);
  });

  resultsSection.hidden = false;
}

async function handlePreview(event) {
  event.preventDefault();
  hideError();
  if (!validateFiles()) return;

  setLoading(true, "Mencocokkan data…");
  try {
    const res = await fetch("/api/check/preview", {
      method: "POST",
      body: buildFormData(),
    });
    const data = await res.json();
    if (!res.ok) {
      showError(data.detail || "Terjadi kesalahan saat memproses file.");
      return;
    }
    renderResults(data);
  } catch (err) {
    showError("Tidak bisa terhubung ke server. Coba lagi.");
  } finally {
    setLoading(false);
  }
}

async function handleDownload() {
  hideError();
  if (!validateFiles()) return;

  setLoading(true, "Menyiapkan file Excel…");
  try {
    const res = await fetch("/api/check", {
      method: "POST",
      body: buildFormData(),
    });
    if (!res.ok) {
      const data = await res.json();
      showError(data.detail || "Terjadi kesalahan saat memproses file.");
      return;
    }
    const blob = await res.blob();
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : "hasil_cek_plan_actual.xlsx";

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    showError("Tidak bisa terhubung ke server. Coba lagi.");
  } finally {
    setLoading(false);
  }
}

async function handleDownloadPptx() {
  hideError();
  if (!validateFiles()) return;

  setLoading(true, "Menyiapkan file PowerPoint…");
  try {
    const res = await fetch("/api/check/pptx", {
      method: "POST",
      body: buildFormData(),
    });
    if (!res.ok) {
      const data = await res.json();
      showError(data.detail || "Terjadi kesalahan saat memproses file.");
      return;
    }
    const blob = await res.blob();
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : "hasil_cek_plan_actual.pptx";

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    showError("Tidak bisa terhubung ke server. Coba lagi.");
  } finally {
    setLoading(false);
  }
}

form.addEventListener("submit", handlePreview);
downloadBtn.addEventListener("click", handleDownload);
downloadPptxBtn.addEventListener("click", handleDownloadPptx);
