// Focus Tasks: if the app can't start, say why instead of staying on "Loading Focus...".
// The usual reason: some files on GitHub are from a different version than index.html, so the
// browser's tamper check refuses to run them. The fix is uploading all the files together.
(() => {
  const bad = new Set();
  const show = extra => {
    const boot = document.getElementById("boot");
    if (!boot || boot.classList.contains("hidden")) return;
    boot.innerHTML = "";
    const h = document.createElement("p"); h.innerHTML = "<b>Focus couldn't start</b>";
    const p = document.createElement("p");
    p.style.cssText = "color:var(--muted);font-size:14px";
    p.textContent = (bad.size
      ? `These files on GitHub don't match this version of index.html: ${[...bad].join(", ")}. `
      : "") + "Upload all the files from the newest zip to GitHub together (keep your config.js), wait a minute, then close the app completely and open it again." + (extra || "");
    boot.append(h, p);
  };
  window.addEventListener("error", e => {
    const el = e.target;
    if (el && el.tagName === "SCRIPT") { bad.add((el.getAttribute("src") || "").split("?")[0]); show(); }
  }, true);
  setTimeout(() => show(" (It's been 12 seconds.)"), 12000);
})();
