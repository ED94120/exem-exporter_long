(() => {
  const SCRIPT_VERSION = "EXPO_CAPTEUR_V1_2026_02_22";
  const SEUIL_EXPO_MAX = 10.0; // contrainte utilisée pour vérifier si le décodage des pixels donne des niveaux d'Exposition acceptables.
  const FENETRE_DELTA_MINUTES = 10; // // fenêtre de tolérance (± minutes) autour de H + delta

  // --------------------------
  // Utils dates / nombres
  // --------------------------
  function parseFRDate(s) {
    const m = String(s || "").trim().match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
    if (!m) return null;
    return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], 0, 0);
  }

  function fmtFRDate(d) {
    const p = n => String(n).padStart(2, "0");
    return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  function fmtFRNumber(n) {
    return Number(n).toFixed(2).replace(".", ",");
  }

  function fmtCompactLocal(d) {
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
  }

  function sanitizeFileName(s) {
    return String(s || "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9._-]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 120);
  }

  // --------------------------
  // Saisie : Cancel = "je ne sais pas encore"
  // --------------------------
  function askText(label, current) {
    const v = prompt(label, current == null ? "" : String(current));
    if (v === null) return null; // cancel = inconnu
    const s = String(v);
    return s; // vide autorisé (bloqué à la synthèse si obligatoire)
  }

  function askDate(label, currentStr) {
    const v = prompt(label, currentStr || "");
    if (v === null) return { str: null, date: null }; // cancel = inconnu
    const s = String(v).trim();
    if (s === "") return { str: "", date: null }; // vide = inconnu
    const d = parseFRDate(s);
    if (!d) {
      alert("Date invalide. Format attendu : dd/mm/yyyy hh:mm");
      return askDate(label, currentStr);
    }
    return { str: s, date: d };
  }

  function askNumberFR(label, currentNum) {
    const cur = (currentNum == null || !Number.isFinite(currentNum)) ? "" : String(currentNum).replace(".", ",");
    const v = prompt(label, cur);
    if (v === null) return NaN; // cancel = inconnu
    const s = String(v).trim();
    if (s === "") return NaN; // vide = inconnu
    const n = parseFloat(s.replace(",", "."));
    if (!Number.isFinite(n)) {
      alert("Nombre invalide.");
      return askNumberFR(label, currentNum);
    }
    return n;
  }

  function isMissingText(v) {
    return (v == null) || (String(v).trim() === "");
  }

  function isMissingDate(d) {
    return !(d instanceof Date) || isNaN(d.getTime());
  }

  function isMissingNumber(n) {
    return !Number.isFinite(n);
  }

  function getMissingFields(P) {
    const miss = [];
    if (isMissingText(P.reference)) miss.push("Référence");
    if (isMissingText(P.adresse)) miss.push("Adresse");
    if (isMissingDate(P.DateDeb)) miss.push("Date début");
    if (isMissingDate(P.DateFin)) miss.push("Date fin");
    if (isMissingNumber(P.ExpoDeb)) miss.push("Expo début");
    if (isMissingNumber(P.ExpoFin)) miss.push("Expo fin");
    if (isMissingNumber(P.ExpoMax)) miss.push("Expo max");
    return miss;
  }

  function buildRecap(P) {
    const show = v => (isMissingText(v) ? "NON RENSEIGNÉ" : String(v));
    const showDate = (s, d) => (isMissingDate(d) ? "NON RENSEIGNÉ" : s);
    const showNum = n => (isMissingNumber(n) ? "NON RENSEIGNÉ" : (fmtFRNumber(n) + " V/m"));
    const yesno = b => (b ? "OUI" : "NON");
    const missing = getMissingFields(P);

    return [
      "Récapitulatif",
      "--------------------------------",
      `Référence : ${show(P.reference)}`,
      `Adresse : ${show(P.adresse)}`,
      `Date début: ${showDate(P.sDateDeb, P.DateDeb)}`,
      `Date fin : ${showDate(P.sDateFin, P.DateFin)}`,
      `Expo début: ${showNum(P.ExpoDeb)}`,
      `Expo fin : ${showNum(P.ExpoFin)}`,
      `Expo max : ${showNum(P.ExpoMax)}`,
      `Pixels bruts : ${yesno(P.archiverPixels)}`,
      "",
      missing.length ? ("Champs manquants : " + missing.join(", ")) : "Tous les champs sont renseignés.",
      "",
      "OK = Lancer (si complet) / Annuler = Modifier"
    ].join("\n");
  }

  function editOneField(P) {
    const choice = prompt(
      "Modifier quel champ ?\n" +
      "1 Référence\n" +
      "2 Adresse\n" +
      "3 Date début\n" +
      "4 Date fin\n" +
      "5 Expo début\n" +
      "6 Expo fin\n" +
      "7 Expo max\n" +
      "8 Pixels bruts\n" +
      "0 Retour\n" +
      "00 EXIT (abandonner)",
      "0"
    );
    if (choice === null) return true; // Cancel = retour synthèse
    const c = String(choice).trim();
    if (c === "00") return false; // EXIT complet
    if (c === "0") return true; // retour synthèse

    if (c === "1") { P.reference = askText("Référence Capteur :", P.reference); return true; }
    if (c === "2") { P.adresse = askText("Adresse Capteur :", P.adresse); return true; }
    if (c === "3") { const r = askDate("Date début (dd/mm/yyyy hh:mm) :", P.sDateDeb); P.sDateDeb = r.str; P.DateDeb = r.date; return true; }
    if (c === "4") { const r = askDate("Date fin (dd/mm/yyyy hh:mm) :", P.sDateFin); P.sDateFin = r.str; P.DateFin = r.date; return true; }
    if (c === "5") { P.ExpoDeb = askNumberFR("Expo début (V/m) :", P.ExpoDeb); return true; }
    if (c === "6") { P.ExpoFin = askNumberFR("Expo fin (V/m) :", P.ExpoFin); return true; }
    if (c === "7") { P.ExpoMax = askNumberFR("Expo max (V/m) :", P.ExpoMax); return true; }
    if (c === "8") { P.archiverPixels = confirm("Archiver pixels bruts ?"); return true; }

    alert("Choix invalide.");
    return true;
  }

  function collectAndConfirmUserInputs() {
    const P = {
      reference: null,
      adresse: null,
      sDateDeb: null,
      sDateFin: null,
      DateDeb: null,
      DateFin: null,
      ExpoDeb: NaN,
      ExpoFin: NaN,
      ExpoMax: NaN,
      archiverPixels: false
    };

    // Saisie initiale (Cancel = inconnu, on continue)
    P.reference = askText("Référence Capteur (ex: Site #Nantes_46) :", "");
    P.adresse = askText("Adresse Capteur :", "");

    {
      const rDeb = askDate("Date début (dd/mm/yyyy hh:mm) :", "");
      P.sDateDeb = rDeb.str;
      P.DateDeb = rDeb.date;
    }
    {
      const rFin = askDate("Date fin (dd/mm/yyyy hh:mm) :", "");
      P.sDateFin = rFin.str;
      P.DateFin = rFin.date;
    }

    P.ExpoDeb = askNumberFR("Expo début (V/m) :", NaN);
    P.ExpoFin = askNumberFR("Expo fin (V/m) :", NaN);
    P.ExpoMax = askNumberFR("Expo max (V/m) :", NaN);
    P.archiverPixels = confirm("Archiver pixels bruts ?");

    // Boucle synthèse / correction
    while (true) {
      // Vérification DateFin > DateDeb si présentes
      if (!isMissingDate(P.DateDeb) && !isMissingDate(P.DateFin)) {
        if (P.DateFin.getTime() <= P.DateDeb.getTime()) {
          alert("Erreur : Date fin doit être strictement après Date début.");
          const rFin = askDate("Date fin (dd/mm/yyyy hh:mm) :", P.sDateFin);
          P.sDateFin = rFin.str;
          P.DateFin = rFin.date;
        }
      }

      const missing = getMissingFields(P);
      const ok = confirm(buildRecap(P));

      if (ok) {
        if (missing.length) {
          alert("Impossible de lancer : complète les champs manquants.\n" + missing.join(", "));
          const keep = editOneField(P);
          if (!keep) return null; // EXIT demandé
          continue;
        }
        return P; // tout est OK
      }

      const keep = editOneField(P);
      if (!keep) return null; // EXIT demandé
    }
  }

  // ============================================================
  // UI EXPORT
  // ============================================================
  function getOrCreateExportBox() {
    let box = document.getElementById("EXEM_EXPORT_BOX");
    if (!box) {
      box = document.createElement("div");
      box.id = "EXEM_EXPORT_BOX";
      box.style.cssText =
        "position:fixed;right:12px;bottom:12px;z-index:999999;" +
        "background:#fff;border:1px solid #999;border-radius:6px;" +
        "padding:10px;max-width:55vw;max-height:45vh;overflow:auto;" +
        "font:12px/1.3 Arial;box-shadow:0 2px 10px rgba(0,0,0,0.2)";

      const closeBtn = document.createElement("button");
      closeBtn.textContent = "✕";
      closeBtn.style.cssText =
        "position:absolute;top:6px;right:8px;border:none;background:none;" +
        "cursor:pointer;font-weight:bold;font-size:14px";
      closeBtn.onclick = () => box.remove();
      box.appendChild(closeBtn);

      const title = document.createElement("div");
      title.textContent = "Exports";
      title.style.cssText = "font-weight:bold;margin-bottom:8px";
      box.appendChild(title);

      const btns = document.createElement("div");
      btns.id = "EXEM_EXPORT_BTNS";
      btns.style.cssText = "display:flex;flex-direction:column;gap:6px";
      box.appendChild(btns);

      const hint = document.createElement("div");
      hint.textContent = "Clique sur un bouton pour enregistrer le fichier.";
      hint.style.cssText = "margin-top:8px;color:#444";
      box.appendChild(hint);

      document.body.appendChild(box);
    }

    let btnWrap = box.querySelector("#EXEM_EXPORT_BTNS");
    if (!btnWrap) {
      btnWrap = document.createElement("div");
      btnWrap.id = "EXEM_EXPORT_BTNS";
      btnWrap.style.cssText = "display:flex;flex-direction:column;gap:6px";
      box.appendChild(btnWrap);
    }
    return box;
  }

  async function downloadFileUserClick(name, content, label = "Télécharger") {
    const box = getOrCreateExportBox();
    const btnWrap = box.querySelector("#EXEM_EXPORT_BTNS");

    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = `${label} : ${name}`;
    btn.style.cssText =
      "cursor:pointer;padding:6px 10px;text-align:left;border:1px solid #777;" +
      "border-radius:4px;background:#f7f7f7";

    btn.onclick = async () => {
      if (window.showSaveFilePicker) {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: name,
            types: [{ description: "CSV", accept: { "text/csv": [".csv"] } }]
          });
          const writable = await handle.createWritable();
          await writable.write(content);
          await writable.close();
          btn.disabled = true;
          btn.textContent = `Enregistré : ${name}`;
          btn.style.opacity = "0.7";
          return;
        } catch (err) {
          if (err && err.name === "AbortError") return;
          alert("Erreur sauvegarde : " + (err && err.message ? err.message : err));
          console.error(err);
          return;
        }
      }

      // Fallback universel : Blob download
      try {
        const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);

        btn.disabled = true;
        btn.textContent = `Téléchargé : ${name}`;
        btn.style.opacity = "0.7";
      } catch (e) {
        alert("Erreur téléchargement : " + (e && e.message ? e.message : e));
        console.error(e);
      }
    };

    btnWrap.appendChild(btn);
  }

  function auditDeltas(dataRows) {
    if (!dataRows || dataRows.length < 2) {
      console.log("[AUDIT] Pas assez de points pour calculer les deltas.");
      return null;
    }
  
    // Sécurité : tri chronologique
    dataRows.sort((a, b) => a.dt - b.dt);
  
    const hist = Object.create(null);
    let dMin = Infinity, dMax = -Infinity;
  
    let nbEq120 = 0;
    let nbLt120 = 0;
    let nbGt120 = 0;
  
    const gaps = [];   // trous
    const tight = [];  // trop serré
  
    for (let i = 1; i < dataRows.length; i++) {
      const d0 = dataRows[i - 1].dt.getTime();
      const d1 = dataRows[i].dt.getTime();
      const dMinu = Math.round((d1 - d0) / 60000); // en minutes, arrondi
  
      if (!isFinite(dMinu)) continue;
  
      hist[dMinu] = (hist[dMinu] || 0) + 1;
  
      if (dMinu < dMin) dMin = dMinu;
      if (dMinu > dMax) dMax = dMinu;
  
      if (dMinu === 120) nbEq120++;
      else if (dMinu < 120) { nbLt120++; tight.push({ i, dMinu, t0: dataRows[i-1].dt, t1: dataRows[i].dt }); }
      else { nbGt120++; gaps.push({ i, dMinu, t0: dataRows[i-1].dt, t1: dataRows[i].dt }); }
    }
  
    // Delta dominant
    let topDelta = null, topCount = -1;
    for (const k in hist) {
      const c = hist[k];
      if (c > topCount) { topCount = c; topDelta = +k; }
    }
  
    const nPairs = dataRows.length - 1;
    const pctTop = nPairs > 0 ? (100 * topCount / nPairs) : 0;
    const pct120 = nPairs > 0 ? (100 * nbEq120 / nPairs) : 0;
  
    console.log("[AUDIT] N points =", dataRows.length, " | N deltas =", nPairs);
    console.log("[AUDIT] DeltaMin =", dMin, "min | DeltaMax =", dMax, "min");
    console.log("[AUDIT] Delta dominant =", topDelta, "min |", topCount, "(", pctTop.toFixed(2), "% )");
    console.log("[AUDIT] Delta = 120 min :", nbEq120, "(", pct120.toFixed(2), "% )");
    console.log("[AUDIT] Delta < 120 min :", nbLt120);
    console.log("[AUDIT] Delta > 120 min :", nbGt120);
  
    // Histogramme lisible (trié)
    const keys = Object.keys(hist).map(Number).sort((a,b)=>a-b);
    console.log("[AUDIT] Histogramme deltas (min -> count) :");
    for (const k of keys) console.log("  ", k, "->", hist[k]);
  
    // Optionnel : afficher quelques anomalies
    if (gaps.length) {
      console.log("[AUDIT] Exemples de trous (delta > 120) :");
      for (let j = 0; j < Math.min(8, gaps.length); j++) {
        const g = gaps[j];
        console.log("  delta", g.dMinu, "min :", g.t0, "->", g.t1);
      }
    }
    if (tight.length) {
      console.log("[AUDIT] Exemples trop serrés (delta < 120) :");
      for (let j = 0; j < Math.min(8, tight.length); j++) {
        const t = tight[j];
        console.log("  delta", t.dMinu, "min :", t.t0, "->", t.t1);
      }
    }
  
    return { hist, dMin, dMax, topDelta, topCount, pctTop, nbEq120, nbLt120, nbGt120, gaps, tight };
  }

  // ============================================================
  // AUDIT DELTAS (pas temporel entre mesures)
  // ============================================================
  function auditDeltasCSV(decodedHourly, audit) {
  
    if (!decodedHourly || decodedHourly.length < 2) {
      audit.push("AUDIT;DELTA_STEP;PasAssezDePoints");
      return;
    }
  
    let dMin = Infinity;
    let dMax = -Infinity;
  
    const hist = {};
    let nbEq120 = 0;
    let nbLt120 = 0;
    let nbGt120 = 0;
  
    for (let i = 1; i < decodedHourly.length; i++) {
  
      const t0 = decodedHourly[i - 1][0];
      const t1 = decodedHourly[i][0];
  
      const deltaMin = Math.round((t1 - t0) / 60000);
  
      hist[deltaMin] = (hist[deltaMin] || 0) + 1;
  
      if (deltaMin < dMin) dMin = deltaMin;
      if (deltaMin > dMax) dMax = deltaMin;
  
      if (deltaMin === 120) nbEq120++;
      else if (deltaMin < 120) nbLt120++;
      else nbGt120++;
    }
  
    // delta dominant
    let topDelta = null;
    let topCount = -1;
  
    for (const k in hist) {
      if (hist[k] > topCount) {
        topCount = hist[k];
        topDelta = Number(k);
      }
    }
  
    const total = decodedHourly.length - 1;
    const pct120 = total > 0 ? (100 * nbEq120 / total).toFixed(2) : "0";
  
    audit.push(`AUDIT;DELTA_STEP;TotalDeltas=${total}`);
    audit.push(`AUDIT;DELTA_STEP;DeltaMin=${dMin}`);
    audit.push(`AUDIT;DELTA_STEP;DeltaMax=${dMax}`);
    audit.push(`AUDIT;DELTA_STEP;DeltaDominant=${topDelta};Count=${topCount}`);
    audit.push(`AUDIT;DELTA_STEP;Delta120_Count=${nbEq120};Pct=${pct120}%`);
    audit.push(`AUDIT;DELTA_STEP;DeltaLt120=${nbLt120}`);
    audit.push(`AUDIT;DELTA_STEP;DeltaGt120=${nbGt120}`);
  
    // histogramme compact
    const keys = Object.keys(hist).map(Number).sort((a,b)=>a-b);
    for (const k of keys) {
      audit.push(`AUDIT;DELTA_HISTO;${k};${hist[k]}`);
    }
  }
  
  // --------------------------
  // Extraction pixels
  // --------------------------
  const graph = document.querySelector("path.highcharts-graph");
  if (!graph) {
    alert("Graph not found");
    return;
  }

  const d = graph.getAttribute("d") || "";
  const tokens = d.match(/[A-Za-z]|-?\d*\.?\d+(?:e[+-]?\d+)?/g) || [];
  const pts = [];
  let i = 0;

  while (i < tokens.length) {
    const t = tokens[i++];
    if (t === "M" || t === "L") {
      const x = parseFloat(tokens[i++]);
      const y = parseFloat(tokens[i++]);
      if (Number.isFinite(x) && Number.isFinite(y)) pts.push([x, y]);
    } else if (t === "C") {
      i += 4;
      const x = parseFloat(tokens[i++]);
      const y = parseFloat(tokens[i++]);
      if (Number.isFinite(x) && Number.isFinite(y)) pts.push([x, y]);
    }
  }

  if (pts.length < 2) {
    alert("Pas assez de points.");
    return;
  }

  pts.sort((a, b) => a[0] - b[0]);

  // --------------------------
  // Saisie + synthèse/corrections
  // --------------------------
  const P = collectAndConfirmUserInputs();
  if (!P) {
    alert("Abandon.");
    return;
  }

  const reference = P.reference;
  const adresse = P.adresse;
  const sDateDeb = P.sDateDeb;
  const sDateFin = P.sDateFin;
  const DateDeb = P.DateDeb;
  const DateFin = P.DateFin;
  const ExpoDeb = P.ExpoDeb;
  const ExpoFin = P.ExpoFin;
  const ExpoMaxUser = P.ExpoMax;
  const archiverPixels = P.archiverPixels;

  // --------------------------
  // Calibration
  // --------------------------
  const xDeb = pts[0][0];
  const yDeb = pts[0][1];
  const xFin = pts[pts.length - 1][0];
  const yFin = pts[pts.length - 1][1];

  const tDeb = DateDeb.getTime();
  const tFin = DateFin.getTime();

  const penteTemps = (tFin - tDeb) / (xFin - xDeb);
  const penteExpo = (ExpoFin - ExpoDeb) / (yFin - yDeb);

  // --------------------------
  // Décodage
  // --------------------------
  const decoded = [];
  const audit = [];
  let inversions = 0;
  let prevTcode = -Infinity;

  for (let k = 0; k < pts.length; k++) {
    const x = pts[k][0];
    const y = pts[k][1];

    if (x <= prevTcode) {
      inversions++;
      audit.push(`AUDIT;INVERSION_TEMPS_CODE;${k};x=${x}`);
    }
    prevTcode = x;

    const t_ms = tDeb + (x - xDeb) * penteTemps;
    const E = ExpoDeb + (y - yDeb) * penteExpo;
    decoded.push([t_ms, E]);
  }

  // --------------------------
  // Vérifications finales
  // --------------------------

  for (let k = 0; k < decoded.length; k++) {
  if (decoded[k][1] !== null && decoded[k][1] >= SEUIL_EXPO_MAX) {
    audit.push(`AUDIT;EXPO_SUP_10;${k};E=${decoded[k][1]}`);
    decoded[k][1] = null;
   }
  }

  // ----------------------------------------------------
  // Calcul de delta (minute dominante) par histogramme
  // ----------------------------------------------------
  const histMin = new Array(60).fill(0);
  let nOk = 0;
  
  for (let k = 0; k < decoded.length; k++) {
    const E = decoded[k][1];
    if (E === null) continue;
    const t = decoded[k][0];
    const m = new Date(t).getMinutes(); // heure locale affichée (EXEM)
    histMin[m]++;
    nOk++;
  }
  
  // Trouver le meilleur pic (top1) et le second (top2)
  let top1 = 0, top2 = 0;
  for (let m = 1; m < 60; m++) {
    if (histMin[m] > histMin[top1]) top1 = m;
  }
  top2 = (top1 === 0) ? 1 : 0;
  for (let m = 0; m < 60; m++) {
    if (m === top1) continue;
    if (histMin[m] > histMin[top2]) top2 = m;
  }
  
  // Masse autour du pic : ±2 minutes (avec modulo 60)
  function sumAround(hist, center, radius) {
    let s = 0;
    for (let d = -radius; d <= radius; d++) {
      const mm = (center + d + 60) % 60;
      s += hist[mm];
    }
    return s;
  }
  
  const mass = sumAround(histMin, top1, 2);
  const countTop1 = histMin[top1];
  const countTop2 = histMin[top2];
  
  // Règles de validation (robustes)
  const okMass = (nOk > 0) && (mass >= 0.50 * nOk);
  const okRatio = (countTop2 === 0) ? (countTop1 >= 5) : (countTop1 >= 2 * countTop2);
  
  const DELTA_MINUTES = top1;         // delta proposé (minute dominante)
  const DELTA_OK = okMass && okRatio; // delta jugé fiable
  
  audit.push(`AUDIT;DELTA_ESTIME;DeltaMin=${DELTA_MINUTES};DeltaOK=${DELTA_OK ? "OUI" : "NON"};N=${nOk};Top1=${countTop1};Top2=${countTop2};Mass+/-2=${mass}`);
  
 // on impose : ≤ 1 point par heure, si δ est fiable : on privilégie H:δ avec tolérance ±FENETRE_DELTA_MINUTES, sinon : on retombe sur H:00
 const decodedHourly = [];
  const byHour = new Map(); // clé = heure (entier), valeur = objet {t,E,dist,inside}
  
  for (let k = 0; k < decoded.length; k++) {
    const t = decoded[k][0];
    const E = decoded[k][1];
    if (E === null) continue;
  
    const h = Math.floor(t / 3600000);
    const t0 = h * 3600000;
  
    // Cible : H:delta si delta fiable, sinon H:00
    const target = DELTA_OK ? (t0 + DELTA_MINUTES * 60000) : t0;
  
    const dist = Math.abs(t - target);
  
    // "inside" = dans la fenêtre ± FENETRE_DELTA_MINUTES autour de la cible
    // (utile seulement si DELTA_OK, sinon on s'en fout)
    const inside = DELTA_OK ? (dist <= FENETRE_DELTA_MINUTES * 60000) : true;
  
    const prev = byHour.get(h);
    if (!prev) {
      byHour.set(h, { t, E, dist, inside });
    } else {
      // Priorité à un point "inside" si l'autre ne l'est pas
      if (inside && !prev.inside) {
        byHour.set(h, { t, E, dist, inside });
      } else if (inside === prev.inside) {
        // Si même statut (inside/inside ou dehors/dehors), prendre la plus petite distance
        if (dist < prev.dist) byHour.set(h, { t, E, dist, inside });
      }
      // Sinon, on garde prev
    }
  }
  
  for (const v of byHour.values()) decodedHourly.push([v.t, v.E]);
  decodedHourly.sort((a, b) => a[0] - b[0]);
  
  auditDeltas(decodedHourly);
  audit.push(`AUDIT;NB_HOURLY;NbHeuresAvecMesure=${decodedHourly.length}`);
  auditDeltasCSV(decodedHourly, audit); // écrit dans CSV
  
  const nbMesures = decoded.length;
  const nbMesuresValides = decoded.reduce((acc, d) => acc + (d[1] === null ? 0 : 1), 0);

  // --------------------------
  // Stats a posteriori : Min / Moy (moyenne simple EXEM) / Max
  // --------------------------
  const vals = decodedHourly.map(d => d[1]).filter(v => v !== null && Number.isFinite(v));
  let Emin = NaN, Emoy = NaN, Emax = NaN;

  if (vals.length) {
    let s = 0;
    Emin = Infinity;
    Emax = -Infinity;
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      if (v < Emin) Emin = v;
      if (v > Emax) Emax = v;
      s += v;
    }
    Emoy = s / vals.length;
  }

  // Contrôle visuel : comparaison max saisi vs max décodé (informatif)
  if (Number.isFinite(Emax) && Number.isFinite(ExpoMaxUser)) {
    const diff = Math.abs(Emax - ExpoMaxUser);
    if (diff > 0.15 && diff > 0.05 * ExpoMaxUser) {
      audit.push(
        `AUDIT;MAX_INCOHERENT;EmaxDecoded=${fmtFRNumber(Emax)};EmaxSaisi=${fmtFRNumber(ExpoMaxUser)};Diff=${fmtFRNumber(diff)}`
      );
    }
  }

  audit.push(
    `AUDIT;STATS;Emin=${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""};Emoy=${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""};Emax=${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`
  );

  alert(
    "Stats calculées (à comparer avec EXEM) :\n" +
    "Min : " + (Number.isFinite(Emin) ? fmtFRNumber(Emin) : "NA") + " V/m\n" +
    "Moy : " + (Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : "NA") + " V/m\n" +
    "Max : " + (Number.isFinite(Emax) ? fmtFRNumber(Emax) : "NA") + " V/m"
  );

  // --------------------------
  // Construction CSV
  // --------------------------
  const now = new Date();
  const refSafe = sanitizeFileName(reference);
  const baseName = `${refSafe || "Capteur"}__${fmtCompactLocal(DateDeb)}__${fmtCompactLocal(DateFin)}`;

  const lines = [];
  lines.push("META;Format;EXPO_CAPTEUR_V1");
  lines.push(`META;ScriptVersion;${SCRIPT_VERSION}`);
  lines.push(`META;DateCreationExport;${fmtFRDate(now)}`);
  lines.push(`META;Reference_Capteur;${reference}`);
  lines.push(`META;Adresse_Capteur;${adresse}`);
  lines.push(`META;DateDebut;${sDateDeb}`);
  lines.push(`META;DateFin;${sDateFin}`);
  lines.push(`META;ExpoDebut_Vm;${fmtFRNumber(ExpoDeb)}`);
  lines.push(`META;ExpoFin_Vm;${fmtFRNumber(ExpoFin)}`);
  lines.push(`META;ExpoMax_Saisie_Vm;${Number.isFinite(ExpoMaxUser) ? fmtFRNumber(ExpoMaxUser) : ""}`);
  lines.push(`META;ExpoMin_Decodee_Vm;${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""}`);
  lines.push(`META;ExpoMoy_Decodee_Vm;${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""}`);
  lines.push(`META;ExpoMax_Decodee_Vm;${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`);
  lines.push(`META;InversionDetectee;${inversions > 0 ? "OUI" : "NON"}`);
  lines.push(`META;NbCouplesInverses;${inversions}`);
  lines.push(`META;Pixels_Archive;${archiverPixels ? "OUI" : "NON"}`);
  lines.push(`META;NbMesures;${nbMesures}`);
  lines.push(`META;NbMesuresValides;${nbMesuresValides}`);
  lines.push(`META;SeuilExpoMax_Vm;${fmtFRNumber(SEUIL_EXPO_MAX)}`);
  lines.push(`META;FenetreDeltaMinutes;${FENETRE_DELTA_MINUTES}`);
  lines.push(`META;RegleFiltrage;1_point_max_par_heure;Cible=H+DeltaSiValide(+/-${FENETRE_DELTA_MINUTES}min);Expo>=${fmtFRNumber(SEUIL_EXPO_MAX)}Vm_exclu`);
  lines.push(`META;DeltaEstimeMinutes;${DELTA_MINUTES}`);
  lines.push(`META;DeltaEstimeOK;${DELTA_OK ? "OUI" : "NON"}`);

  lines.push("DATA;DateHeure;Exposition_Vm");
  decodedHourly.forEach(d => {
    lines.push(`DATA;${fmtFRDate(new Date(d[0]))};${d[1] === null ? "" : fmtFRNumber(d[1])}`);
  });

  audit.forEach(a => lines.push(a));

  // --------------------------
  // Export
  // --------------------------
  try {
    downloadFileUserClick(baseName + ".csv", lines.join("\n"), "Télécharger CSV");

    if (archiverPixels) {
      const pix = ["PIXELS;xp;yp"];
      pts.forEach(p => pix.push(`PIXELS;${p[0]};${p[1]}`));
      downloadFileUserClick(baseName + "_pixels.csv", pix.join("\n"), "Télécharger PIXELS");
    }

    alert("Export prêt. Les boutons sont en bas à droite : " + baseName);
  } catch (e) {
    alert("Erreur export UI : " + (e && e.message ? e.message : e));
    console.error(e);
  }
})();
