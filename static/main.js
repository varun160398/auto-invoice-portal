function showStatus(card, text, isError=false) {
  const el = card.querySelector(".status");
  el.textContent = text;
  el.className = "status " + (isError ? "error" : "ok");
}

// --- Signature upload ---
document.addEventListener("click", async (e) => {
  if (!e.target.classList.contains("uploadBtn")) return;

  const card = e.target.closest(".card");
  const name = card.dataset.name;
  const fileInput = card.querySelector(".fileInput");

  if (!fileInput.files.length) {
    showStatus(card, "Please choose an image first.", true);
    return;
  }

  const form = new FormData();
  form.append("name", name);
  form.append("file", fileInput.files[0]);

  showStatus(card, "Uploading…", false);

  try {
    const res = await fetch("/api/upload-signature", { method: "POST", body: form });
    const data = await res.json();

    if (!res.ok) {
      showStatus(card, data.error || "Upload failed", true);
      return;
    }

    const badge = card.querySelector(".badge");
    badge.textContent = "Signature Uploaded";
    badge.classList.remove("warn");
    badge.classList.add("ok");

    showStatus(card, "Image uploaded ✅ (download PDF now)", false);
  } catch (err) {
    showStatus(card, "Upload error: " + err, true);
  }
});

// --- Period dropdown (Month/Year) ---
async function savePeriod() {
  const month = document.getElementById("monthSel")?.value;
  const year = document.getElementById("yearSel")?.value;
  if (!month || !year) return;

  const form = new FormData();
  form.append("month", month);
  form.append("year", year);

  try {
    const res = await fetch("/api/set-period", { method: "POST", body: form });
    const data = await res.json();
    if (!res.ok) {
      alert(data.error || "Failed to set period");
      return;
    }
    // optional: small toast feel
    console.log("Period saved:", month, year);
  } catch (e) {
    alert("Error saving period: " + e);
  }
}

document.getElementById("monthSel")?.addEventListener("change", savePeriod);
document.getElementById("yearSel")?.addEventListener("change", savePeriod);
