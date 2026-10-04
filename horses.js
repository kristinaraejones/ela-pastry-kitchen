// The ELA Pastry Kitchen — horse theme (beta only).
// Loaded AFTER beta.js. Adds horse options to "Make It Your Kitchen" without
// touching the existing ones: two new sous-chef buddies (Horse, Pony), a
// Horseshoe cookie signature pastry, a Horseshoes page pattern, a Horseshoes
// celebration burst, and two stable-inspired awning colors.

(function () {
  const O = BK_KITCHEN_OPTIONS;
  const add = (arr, v) => { if (arr.indexOf(v) < 0) arr.splice(arr.indexOf("none") >= 0 ? arr.indexOf("none") : arr.length, 0, v); };
  add(O.buddies, "horse"); add(O.buddies, "pony");
  O.pastries.push("horseshoe");
  O.patterns.push("horseshoes");
  O.celebrations.push("horseshoes");
  Object.keys(O.awnings).forEach(kid => {
    O.awnings[kid].push(["pasture", "Pasture", ["#B7D9A8", "#E8D7B0", "#9ED6C6"]]);
    O.awnings[kid].push(["saddle", "Saddle Up", ["#D9A47A", "#F3E2C7", "#B98A5E"]]);
  });

  const SHOE = (col, w) => `<path d="M6 21C2.5 10 6 3.5 12 3.5s9.5 6.5 6 17.5" fill="none" stroke="${col}" stroke-width="${w || 3.4}" stroke-linecap="round"/>`;

  // ---------- Buddies ----------
  const origBuddy = window.bkBuddySVG;
  window.bkBuddySVG = function (kind, size) {
    const s = size || 52;
    if (kind === "horse") return `<svg width="${s}" height="${s}" viewBox="0 0 64 64" aria-hidden="true">
      <path d="M20 60C20 44 22 32 28 22L27 7L36 14C41 12 46 14 49 19C53 28 56 40 57 47C58 51 55 54 51 54C47 54 45 51 44 48C42 46 38 44 36 42C35 48 36 54 38 60Z" fill="#C98B55" stroke="#8A5A33" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M44 20C48 28 52 37 54 43" fill="none" stroke="#FBEFD9" stroke-width="5" stroke-linecap="round"/>
      <path d="M28 19C18 26 15 44 18 60L26 60C24 46 26 34 33 23Z" fill="#5E3B22" stroke="#4A2D19" stroke-width="2.2" stroke-linejoin="round"/>
      <path d="M36 14C34 18 36 22 40 21" fill="none" stroke="#5E3B22" stroke-width="3" stroke-linecap="round"/>
      <circle cx="41" cy="27" r="2.6" fill="#2B1B10"/><circle cx="41.9" cy="26.1" r=".9" fill="#fff"/>
      <ellipse cx="53" cy="49" rx="1.4" ry="1.8" fill="#6B4224"/>
      <path d="M46 52Q50 54 54 52" fill="none" stroke="#6B4224" stroke-width="1.6" stroke-linecap="round"/></svg>`;
    if (kind === "pony") return `<svg width="${s}" height="${s}" viewBox="0 0 64 64" aria-hidden="true">
      <path d="M21 22L18 5L30 12Z" fill="#F3D9A8" stroke="#C99A55" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M43 22L46 5L34 12Z" fill="#F3D9A8" stroke="#C99A55" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M20 24C20 14 26 9 32 9S44 14 44 24L43 44C42 53 38 59 32 59S22 53 21 44Z" fill="#F7E3BC" stroke="#C99A55" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M32 8C26 8 21 13 20 22C23 20 27 21 29 26C30 21 34 19 38 21C41 22 43 23 44 24C44 14 39 8 32 8Z" fill="#F29AC0" stroke="#CF6A98" stroke-width="2.2" stroke-linejoin="round"/>
      <ellipse cx="32" cy="50" rx="9" ry="7" fill="#FBEBD2" stroke="#C99A55" stroke-width="2"/>
      <circle cx="25" cy="34" r="2.6" fill="#3A2A1E"/><circle cx="39" cy="34" r="2.6" fill="#3A2A1E"/>
      <circle cx="25.8" cy="33.2" r=".9" fill="#fff"/><circle cx="39.8" cy="33.2" r=".9" fill="#fff"/>
      <ellipse cx="28.5" cy="49.5" rx="1.2" ry="1.7" fill="#B07A55"/><ellipse cx="35.5" cy="49.5" rx="1.2" ry="1.7" fill="#B07A55"/>
      <path d="M29 54Q32 56 35 54" fill="none" stroke="#B07A55" stroke-width="1.6" stroke-linecap="round"/></svg>`;
    return origBuddy(kind, size);
  };

  // ---------- Horseshoe cookie (signature pastry) ----------
  const origPastry = window.bkPastry;
  window.bkPastry = function (kind, child, size, extraAttr) {
    if (kind !== "horseshoe") return origPastry(kind, child, size, extraAttr);
    const sw = size <= 40 ? 3 : 2.5;
    const dots = [[17.5, 46], [15.5, 35.5], [19.5, 24], [32, 17], [44.5, 24], [48.5, 35.5], [46.5, 46]]
      .map(([x, y]) => `<circle cx="${x}" cy="${y}" r="${size > 40 ? 2.1 : 2.4}" fill="#8A5A33"/>`).join("");
    return `<svg width="${size}" height="${size}" viewBox="0 0 64 64" aria-hidden="true" ${extraAttr || ""}>
      <path d="M15 52C8 34 14 12 32 12s24 22 17 40" fill="none" stroke="#C0904F" stroke-width="${14 + sw}" stroke-linecap="round"/>
      <path d="M15 52C8 34 14 12 32 12s24 22 17 40" fill="none" stroke="#EDC48E" stroke-width="14" stroke-linecap="round"/>
      ${dots}
      ${size > 40 ? `<path d="M22 19C26 15 31 14 35 15" fill="none" stroke="#FCE3BC" stroke-width="2.4" stroke-linecap="round"/>` : ""}</svg>`;
  };

  // ---------- Page pattern ----------
  const SHOE_BG = `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Cg transform='rotate(-18 12 12)'%3E%3Cpath d='M6 21C2.5 10 6 3.5 12 3.5s9.5 6.5 6 17.5' fill='none' stroke='white' stroke-opacity='.75' stroke-width='3' stroke-linecap='round'/%3E%3C/g%3E%3C/svg%3E")`;
  const style = document.createElement("style");
  style.textContent = `
.bk-pat-horseshoes{background-image:${SHOE_BG};background-size:26px 26px;}
.bk-preview-body.bk-pat-horseshoes{background-size:30px 30px;}
body[data-view!="parent"][data-bk-pattern="horseshoes"] .bk-hero{background-image:${SHOE_BG};background-size:32px 32px;}
.bk-choice-grid:has(> :nth-child(6)){grid-template-columns:repeat(3,minmax(0,1fr));}
.bk-choice-grid:has(> :nth-child(7)){grid-template-columns:repeat(4,minmax(0,1fr));}
.bk-choice-grid:has(> :nth-child(5)):not(:has(> :nth-child(6))){grid-template-columns:repeat(5,minmax(0,1fr));}
@media (max-width:640px){.bk-choice-grid:has(> :nth-child(5)){grid-template-columns:repeat(2,minmax(0,1fr));}.bk-choice-grid.five:has(> :nth-child(6)){grid-template-columns:repeat(2,minmax(0,1fr));}}`;
  document.head.appendChild(style);

  const origPatSwatch = window.bkPatternSwatchHTML;
  window.bkPatternSwatchHTML = function (id) {
    return id === "horseshoes" ? `<span class="bk-pat-swatch bk-pat-horseshoes"></span>` : origPatSwatch(id);
  };

  // ---------- Celebration ----------
  const origCelSwatch = window.bkCelebrationSwatchHTML;
  window.bkCelebrationSwatchHTML = function (id) {
    if (id !== "horseshoes") return origCelSwatch(id);
    const c = ["var(--bk-aw1)", "var(--bk-aw2)", "var(--bk-aw3)"];
    return `<span class="bk-cel-swatch">${c.map((col, i) => `<svg width="18" height="18" viewBox="0 0 24 24" style="position:absolute;left:${16 + i * 24}px;top:${i % 2 ? 6 : 2}px;transform:rotate(${(i - 1) * 25}deg)">${SHOE(col)}</svg>`).join("")}</span>`;
  };
  const origCelebrate = window.bkCelebrate;
  window.bkCelebrate = function (key, unlockedWeek) {
    const before = document.querySelector(".bk-celebrate");
    origCelebrate(key, unlockedWeek);
    const el = document.querySelector(".bk-celebrate");
    if (!el || el === before || bkKitchenFor().celebration !== "horseshoes") return;
    const colors = ["var(--bk-aw1)", "var(--bk-aw2)", "var(--bk-aw3)", "#F6D77C", "var(--bk-accent)"];
    el.querySelectorAll(".bk-cel-spr").forEach((sp, i) => {
      sp.className = "bk-cel-spr shape";
      sp.style.background = "none";
      sp.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24">${SHOE(colors[i % colors.length])}</svg>`;
    });
  };
})();
