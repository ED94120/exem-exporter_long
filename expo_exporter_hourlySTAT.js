(() => {

  const SCRIPT_VERSION = "EXPO_CAPTEUR_SERIES_LONGUES_AVEC STAT_V1_2026_02_24";
  const PixelActif = "NO"; // "YES" => l'option pixels apparaît, "NO" => invisible (prod)

  const SEUIL_EXPO_MAX = 10.0; // valeur maximale des Expositions décodées prise en compte
  const FENETRE_DELTA_MINUTES = 10; // plage en minutes autour du delta moyen pris en compte pour sélectionner les dates de mesure 

  // --------------------------
  // PARAMÈTRES STATISTIQUES
  // --------------------------
  const E_MIN_RATIO = 0.1; // valeur minimale exposition à 9 h prise en compte pour les ratios horaires
  const E9_WARN_LOW = 0.2; // valeur de warning exposition à 9 h prise en compte pour les ratios horaires
  const R_MIN = 50; // valeur minimale des ratios horaires ExpoH sur Expo9 en %
  const R_MAX = 600; // valeur maximale des ratios horaires ExpoH sur Expo9 en %

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
      ...(PixelActif === "YES" ? [`Pixels bruts : ${yesno(P.archiverPixels)}`] : []),
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
      (PixelActif === "YES" ? "8 Pixels bruts\n" : "") +
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
    if (c === "8") {
    if (PixelActif !== "YES") return true; // invisible en prod : ignore
    P.archiverPixels = confirm("Archiver pixels bruts ?");
    return true;
    }

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
    P.archiverPixels = (PixelActif === "YES") ? confirm("Archiver pixels bruts ?") : false;

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
    dataRows.sort((a, b) => a[0] - b[0]);
  
    const hist = Object.create(null);
    let dMin = Infinity, dMax = -Infinity;
  
    let nbEq120 = 0;
    let nbLt120 = 0;
    let nbGt120 = 0;
  
    const gaps = [];   // trous
    const tight = [];  // trop serré
  
    for (let i = 1; i < dataRows.length; i++) {
      const d0 = dataRows[i - 1][0];
      const d1 = dataRows[i][0];
      const dMinu = Math.round((d1 - d0) / 60000); // en minutes, arrondi
  
      if (!isFinite(dMinu)) continue;
  
      hist[dMinu] = (hist[dMinu] || 0) + 1;
  
      if (dMinu < dMin) dMin = dMinu;
      if (dMinu > dMax) dMax = dMinu;
  
      if (dMinu === 120) {
        nbEq120++;
      } else if (dMinu < 120) {
        nbLt120++;
        tight.push({
          i,
          dMinu,
          t0: new Date(dataRows[i - 1][0]),
          t1: new Date(dataRows[i][0])
        });
      } else {
        nbGt120++;
        gaps.push({
          i,
          dMinu,
          t0: new Date(dataRows[i - 1][0]),
          t1: new Date(dataRows[i][0])
        });
      }
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
      audit.push(`AUDIT_HISTO;DELTA_HISTO;${k};${hist[k]}`);
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
    const m = (Math.round(t / 60000) % 60 + 60) % 60;
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
  
  audit.push(`AUDIT_MAIN;DELTA_ESTIME;DeltaMin=${DELTA_MINUTES};DeltaOK=${DELTA_OK ? "OUI" : "NON"};N=${nOk};Top1=${countTop1};Top2=${countTop2};Mass+/-2=${mass}`);
  
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
  auditDeltasCSV(decodedHourly, audit); // écrit dans CSV
  audit.push(`AUDIT_MAIN;NB_HOURLY;NbHeuresAvecMesure=${decodedHourly.length}`);
  
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
    `AUDIT_MAIN;STATS;Emin=${Number.isFinite(Emin) ? fmtFRNumber(Emin) : ""};Emoy=${Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : ""};Emax=${Number.isFinite(Emax) ? fmtFRNumber(Emax) : ""}`
  );

  alert(
    "Stats calculées (à comparer avec EXEM) :\n" +
    "Min : " + (Number.isFinite(Emin) ? fmtFRNumber(Emin) : "NA") + " V/m\n" +
    "Moy : " + (Number.isFinite(Emoy) ? fmtFRNumber(Emoy) : "NA") + " V/m\n" +
    "Max : " + (Number.isFinite(Emax) ? fmtFRNumber(Emax) : "NA") + " V/m"
  );

  // --------------------------
// Analyse statistique ?
// --------------------------
const activerStats = confirm(
  "Ajouter l’analyse statistique complète au CSV ?\n\n" +
  "OUI : STAT_HORAIRE, STAT_ANNUELLE (Année / Été / HorsÉté), TENDANCE, STAT_GLOBAL\n" +
  "NON : Export simple"
);

// ============================================================
// ====================== STAT_GLOBAL ==========================
// ============================================================
let N_global = 0, sum_global = 0, sumsq_global = 0, N_lt_0p1_global = 0;
let Moy_Global = NaN, Sigma_Global = NaN, CV_Global = NaN, Pct_lt_0p1_Global = NaN;

if (activerStats) {
  for (let i = 0; i < decodedHourly.length; i++) {
    const E = decodedHourly[i][1];
    if (!Number.isFinite(E)) continue;

    N_global++;
    sum_global += E;
    sumsq_global += E * E;
    if (E < 0.1) N_lt_0p1_global++;
  }

  if (N_global > 0) {
    Moy_Global = sum_global / N_global;
    const variance = (sumsq_global / N_global) - (Moy_Global * Moy_Global);
    Sigma_Global = variance > 0 ? Math.sqrt(variance) : 0;
    CV_Global = Moy_Global > 1e-9 ? Sigma_Global / Moy_Global : NaN;
    Pct_lt_0p1_Global = 100 * N_lt_0p1_global / N_global;
  }
}

// ============================================================
// ====================== STAT_ANNUELLE ========================
// ============================================================
const statsAnnees = Object.create(null);

function initAgg() {
  return { N: 0, sum: 0, sumsq: 0, N_lt_0p1: 0 };
}

function updateAgg(a, E) {
  a.N++;
  a.sum += E;
  a.sumsq += E * E;
  if (E < 0.1) a.N_lt_0p1++;
}

function finalizeAgg(a) {
  if (!a || a.N === 0) return null;
  const mean = a.sum / a.N;
  const variance = (a.sumsq / a.N) - (mean * mean);
  const sigma = variance > 0 ? Math.sqrt(variance) : 0;
  const cv = mean > 1e-9 ? sigma / mean : NaN;
  const pct = 100 * a.N_lt_0p1 / a.N;
  return { mean, sigma, cv, pct, N: a.N };
}

if (activerStats) {
  for (let i = 0; i < decodedHourly.length; i++) {
    const t = decodedHourly[i][0];
    const E = decodedHourly[i][1];
    if (!Number.isFinite(E)) continue;

    const d = new Date(t);
    const year = d.getFullYear();
    const month = d.getMonth() + 1;

    if (!statsAnnees[year]) {
      statsAnnees[year] = {
        Annee: initAgg(),
        Ete: initAgg(),
        HorsEte: initAgg()
      };
    }

    updateAgg(statsAnnees[year].Annee, E);

    if (month >= 6 && month <= 8)
      updateAgg(statsAnnees[year].Ete, E);
    else
      updateAgg(statsAnnees[year].HorsEte, E);
  }
}

// ============================================================
// ======================== TENDANCE ===========================
// ============================================================
let Trend_Slope = NaN, Trend_R2 = NaN, Trend_K = 0;

if (activerStats) {
  const years = Object.keys(statsAnnees).map(Number).sort((a,b)=>a-b);
  const xs = [], ys = [];

  for (let i = 0; i < years.length; i++) {
    const y = years[i];
    const f = finalizeAgg(statsAnnees[y].Annee);
    if (f) {
      xs.push(y);
      ys.push(f.mean);
    }
  }

  Trend_K = xs.length;

  if (Trend_K >= 3) {
    let sx=0, sy=0, sxx=0, sxy=0;

    for (let i=0;i<Trend_K;i++){
      sx+=xs[i];
      sy+=ys[i];
      sxx+=xs[i]*xs[i];
      sxy+=xs[i]*ys[i];
    }

    const denom = Trend_K*sxx - sx*sx;

    if (denom !== 0){
      Trend_Slope = (Trend_K*sxy - sx*sy)/denom;
      const intercept = (sy - Trend_Slope*sx)/Trend_K;

      let ssRes=0, ssTot=0;
      const ybar = sy/Trend_K;

      for (let i=0;i<Trend_K;i++){
        const yhat = Trend_Slope*xs[i] + intercept;
        const r = ys[i]-yhat;
        ssRes += r*r;
        const t = ys[i]-ybar;
        ssTot += t*t;
      }

      Trend_R2 = ssTot>0 ? 1-ssRes/ssTot : 1;
    }
  }
}

// ============================================================
// ======================== STAT_HORAIRE =======================
// ============================================================

// Périodes (bornes incluses)
// Définition : Periode = [H:00 ; (H+1):00 + delta + marge]
// CONTRAINTE : delta+marge ne doit JAMAIS empiéter sur l'heure suivante (au-delà de (H+1)+59 min),
// sinon : pas de ratio + AUDIT;RATIO_WINDOW_INVALID;...
const PERIODS = [
  { key: "Matin",       startHour: 9  },
  { key: "Midi",        startHour: 12 },
  { key: "Sortie",      startHour: 16 },
  { key: "DebutSoiree", startHour: 19 },
  { key: "FinSoiree",   startHour: 22 }
];

const GROUPS_RATIO = ["Ouvre", "Samedi", "Dimanche", "WE"];

function clamp(x, a, b) { return x < a ? a : (x > b ? b : x); }

function initRatioAgg() {
  return { N: 0, sum: 0, Ncap: 0, RminRaw: Infinity, RmaxRaw: -Infinity, N_EmatinLow: 0 };
}

function updateRatioAgg(agg, Ematin, Eper) {
  if (!Number.isFinite(Ematin) || !Number.isFinite(Eper)) return;
  if (Ematin < E_MIN_RATIO || Eper < E_MIN_RATIO) return;

  const rraw = 100 * (Eper / Ematin);

  agg.N++;
  agg.sum += clamp(rraw, R_MIN, R_MAX);

  if (rraw < R_MIN || rraw > R_MAX) agg.Ncap++;
  if (rraw < agg.RminRaw) agg.RminRaw = rraw;
  if (rraw > agg.RmaxRaw) agg.RmaxRaw = rraw;

  if (Ematin >= E_MIN_RATIO && Ematin < E9_WARN_LOW) agg.N_EmatinLow++;
}

function dayKeyLocal(d) {
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function groupFromDow(dow) {
  if (dow === 0) return "Dimanche";
  if (dow === 6) return "Samedi";
  return "Ouvre";
}

// delta utilisé pour les fenêtres : si delta non fiable, on retombe sur 0
const deltaUsed = (typeof DELTA_OK !== "undefined" && DELTA_OK) ? DELTA_MINUTES : 0;
const margeUsed = FENETRE_DELTA_MINUTES;

// Validation des fenêtres : delta+marge ne doit pas empiéter sur l'heure suivante
// Ici, on impose deltaUsed + margeUsed <= 59 (sinon la borne haute dépasse (H+1):59)
const deltaPlusMarge = deltaUsed + margeUsed;
const ratioWindowsOk = (deltaPlusMarge <= 59);

if (activerStats && !ratioWindowsOk) {
  for (let pi = 0; pi < PERIODS.length; pi++) {
    audit.push(
      `AUDIT_MAIN;RATIO_WINDOW_INVALID;Periode=${PERIODS[pi].key};delta+marge=${deltaPlusMarge};raison=DEPASSE_HEURE_SUIVANTE`
    );
  }
}

// Agrégats ratios : ratio de chaque période (sauf Matin) sur Matin, par groupe
const ratioAgg = Object.create(null);
for (let gi = 0; gi < GROUPS_RATIO.length; gi++) {
  const g = GROUPS_RATIO[gi];
  ratioAgg[g] = Object.create(null);
  for (let pi = 0; pi < PERIODS.length; pi++) {
    const pk = PERIODS[pi].key;
    if (pk === "Matin") continue;
    ratioAgg[g][pk] = initRatioAgg();
  }
}

// Map des jours -> points (minute du jour)
const dayMap = Object.create(null);

function minuteOfDay(d) {
  return d.getHours() * 60 + d.getMinutes();
}

function meanInWindow(points, startMin, endMinIncl) {
  // bornes incluses
  let s = 0, n = 0;
  for (let i = 0; i < points.length; i++) {
    const m = points[i].m;
    if (m >= startMin && m <= endMinIncl) {
      const E = points[i].E;
      if (Number.isFinite(E)) { s += E; n++; }
    }
  }
  return n > 0 ? (s / n) : NaN;
}

if (activerStats && ratioWindowsOk) {

  // 1) Construire dayMap : pour chaque jour, liste des points {m,E}
  for (let i = 0; i < decodedHourly.length; i++) {
    const t = decodedHourly[i][0];
    const E = decodedHourly[i][1];
    if (!Number.isFinite(E)) continue;

    const d = new Date(t);
    const key = dayKeyLocal(d);
    if (!dayMap[key]) dayMap[key] = { points: [] };
    dayMap[key].points.push({ m: minuteOfDay(d), E });
  }

  const keys = Object.keys(dayMap);

  // 2) Pour chaque jour, calcul des moyennes de période puis ratios / Matin
  for (let ik = 0; ik < keys.length; ik++) {
    const key = keys[ik];
    const ptsDay = dayMap[key].points;

    // groupe jour
    const parts = key.split("-");
    const dref = new Date(+parts[0], +parts[1] - 1, +parts[2], 12, 0, 0, 0);
    const g = groupFromDow(dref.getDay());

    // Moyennes par période
    const meanByPeriod = Object.create(null);

    for (let pi = 0; pi < PERIODS.length; pi++) {
      const Pp = PERIODS[pi];
      const H0 = Pp.startHour;

      const startMin = H0 * 60; // H:00
      const endMinIncl = (H0 + 1) * 60 + deltaPlusMarge; // (H+1):delta+marge inclus

      meanByPeriod[Pp.key] = meanInWindow(ptsDay, startMin, endMinIncl);
    }

    const Ematin = meanByPeriod["Matin"];
    if (!Number.isFinite(Ematin) || Ematin < E_MIN_RATIO) {
      // pas de ratio ce jour-là (dénominateur invalide)
      continue;
    }

    for (let pi = 0; pi < PERIODS.length; pi++) {
      const pk = PERIODS[pi].key;
      if (pk === "Matin") continue;

      const Eper = meanByPeriod[pk];
      if (!Number.isFinite(Eper) || Eper < E_MIN_RATIO) continue;

      updateRatioAgg(ratioAgg[g][pk], Ematin, Eper);

      if (g === "Samedi" || g === "Dimanche") {
        updateRatioAgg(ratioAgg["WE"][pk], Ematin, Eper);
      }
    }
  }

  audit.push(`AUDIT_MAIN;NB_JOURS_ANALYSES;${Object.keys(dayMap).length}`);
}

  // ============================================================
  // ========= HISTOGRAMME HEURE DES "HAUTS NIVEAUX" (P95) =======
  // ========= ALL + JO (lun-ven) + WE (sam-dim) =================
  // ============================================================
  
  const histP95_ALL = new Array(24).fill(0);
  const histP95_JO  = new Array(24).fill(0);
  const histP95_WE  = new Array(24).fill(0);
  
  let nbJoursP95_ALL = 0, nbJoursP95_JO = 0, nbJoursP95_WE = 0;
  let nbPointsP95_ALL = 0, nbPointsP95_JO = 0, nbPointsP95_WE = 0;
  let nbJoursP95_Exces_ALL = 0, nbJoursP95_Exces_JO = 0, nbJoursP95_Exces_WE = 0;
  
  const dayP95Map = Object.create(null);
  
  // dayKeyLocal(d) existe déjà dans ton code (utilisé pour ratios)
  for (let i = 0; i < decodedHourly.length; i++) {
    const t = decodedHourly[i][0];
    const E = decodedHourly[i][1];
    if (!Number.isFinite(E)) continue;
  
    const d = new Date(t);
    const key = dayKeyLocal(d);
    const h = d.getHours();
    const dow = d.getDay(); // 0=Dim ... 6=Sam
  
    if (!dayP95Map[key]) dayP95Map[key] = { pts: [], dow };
    dayP95Map[key].pts.push({ h, E });
  }
  
  // Percentile linéaire (p en [0..1])
  function percentileLinear(sortedAsc, p) {
    const n = sortedAsc.length;
    if (n === 0) return NaN;
    if (n === 1) return sortedAsc[0];
    const x = (n - 1) * p;
    const i0 = Math.floor(x);
    const i1 = Math.min(n - 1, i0 + 1);
    const w = x - i0;
    return sortedAsc[i0] * (1 - w) + sortedAsc[i1] * w;
  }
  
  function top2FromHist(hist) {
    let h1 = -1, c1 = -1, h2 = -1, c2 = -1;
    for (let h = 0; h < 24; h++) {
      const c = hist[h];
      if (c > c1) { h2 = h1; c2 = c1; h1 = h; c1 = c; }
      else if (c > c2) { h2 = h; c2 = c; }
    }
    return { h1, c1, h2, c2 };
  }
  
  const keysP95 = Object.keys(dayP95Map);
  
  for (let k = 0; k < keysP95.length; k++) {
    const rec = dayP95Map[keysP95[k]];
    if (!rec || !rec.pts || rec.pts.length < 2) continue;
  
    const pts = rec.pts;
    const dow = rec.dow;
    const isWE = (dow === 0 || dow === 6);
  
    // P95 du jour
    const Es = [];
    for (let j = 0; j < pts.length; j++) Es.push(pts[j].E);
    Es.sort((a, b) => a - b);
  
    const p95 = percentileLinear(Es, 0.95);
    if (!Number.isFinite(p95)) continue;
  
    // jours comptés
    nbJoursP95_ALL++;
    if (isWE) nbJoursP95_WE++; else nbJoursP95_JO++;
  
    // comptage des heures >= P95 du jour
    let nbAbove = 0;
    for (let j = 0; j < pts.length; j++) {
      if (pts[j].E >= p95) {
        const h = pts[j].h;
  
        histP95_ALL[h]++; nbPointsP95_ALL++;
        if (isWE) { histP95_WE[h]++; nbPointsP95_WE++; }
        else      { histP95_JO[h]++; nbPointsP95_JO++; }
  
        nbAbove++;
      }
    }
  
    if (nbAbove > 1) {
      nbJoursP95_Exces_ALL++;
      if (isWE) nbJoursP95_Exces_WE++; else nbJoursP95_Exces_JO++;
    }
  }
  
  const bestP95_ALL = top2FromHist(histP95_ALL);
  const bestP95_JO  = top2FromHist(histP95_JO);
  const bestP95_WE  = top2FromHist(histP95_WE);
  
  // AUDIT (main)
  audit.push(`AUDIT_MAIN;HOURP95_ALL;Jours=${nbJoursP95_ALL};Points>=P95=${nbPointsP95_ALL};Jours_avec_>=2pts=${nbJoursP95_Exces_ALL};Top1H=${bestP95_ALL.h1};Top1Count=${bestP95_ALL.c1};Top2H=${bestP95_ALL.h2};Top2Count=${bestP95_ALL.c2}`);
  audit.push(`AUDIT_MAIN;HOURP95_JO;Jours=${nbJoursP95_JO};Points>=P95=${nbPointsP95_JO};Jours_avec_>=2pts=${nbJoursP95_Exces_JO};Top1H=${bestP95_JO.h1};Top1Count=${bestP95_JO.c1};Top2H=${bestP95_JO.h2};Top2Count=${bestP95_JO.c2}`);
  audit.push(`AUDIT_MAIN;HOURP95_WE;Jours=${nbJoursP95_WE};Points>=P95=${nbPointsP95_WE};Jours_avec_>=2pts=${nbJoursP95_Exces_WE};Top1H=${bestP95_WE.h1};Top1Count=${bestP95_WE.c1};Top2H=${bestP95_WE.h2};Top2Count=${bestP95_WE.c2}`);
    
// ============================================================
// ======================== CONSTRUCTION CSV ===================
// ============================================================
const now = new Date();
const refSafe = sanitizeFileName(reference);
const baseName = `${refSafe || "Capteur"}__${fmtCompactLocal(DateDeb)}__${fmtCompactLocal(DateFin)}`;

const lines=[];

// ----- INFOS -----
lines.push("SECTION;INFOS_GENERALES");
lines.push(`META;ScriptVersion;${SCRIPT_VERSION}`);
lines.push(`META;DateCreationExport;${fmtFRDate(now)}`);
lines.push(`META;AnalyseStatistique;${activerStats?"OUI":"NON"}`);
lines.push(`META;E9_MinPourRatio_Vm;${fmtFRNumber(E_MIN_RATIO)}`);
lines.push(`META;E9_WARN_LOW_Vm;${fmtFRNumber(E9_WARN_LOW)}`);
lines.push(`META;R_MIN_pct;${fmtFRNumber(R_MIN)}`);
lines.push(`META;R_MAX_pct;${fmtFRNumber(R_MAX)}`);

// ----- STAT -----
if (activerStats){

  // STAT_HORAIRE (périodes)
lines.push("SECTION;STAT_HORAIRE");

// Si fenêtres invalides (delta+marge>59), les agrégats sont vides (N=0) + AUDIT déjà écrit
for (let gi = 0; gi < GROUPS_RATIO.length; gi++) {
  const g = GROUPS_RATIO[gi];

  for (let pi = 0; pi < PERIODS.length; pi++) {
    const pk = PERIODS[pi].key;
    if (pk === "Matin") continue;

    const agg = (ratioAgg[g] && ratioAgg[g][pk]) ? ratioAgg[g][pk] : initRatioAgg();

    const mean = agg.N > 0 ? (agg.sum / agg.N) : NaN;
    const rmin = (agg.N > 0 && isFinite(agg.RminRaw)) ? agg.RminRaw : NaN;
    const rmax = (agg.N > 0 && isFinite(agg.RmaxRaw)) ? agg.RmaxRaw : NaN;

    const gTag =
    (g === "Ouvre")   ? "JO" :
    (g === "Samedi")  ? "Sa" :
    (g === "Dimanche")? "Di" :
    (g === "WE")      ? "WE" : "XX";
  
  // Ligne principale (celle qu’on lit en premier) : tag par groupe
  lines.push(`STAT_RATIO_MAIN_${gTag};Ratio_${pk}_sur_Matin_${g}_N${agg.N};${Number.isFinite(mean) ? fmtFRNumber(mean) : ""}`);
  
  // Lignes informatives (détails) : tu peux laisser en STAT_RATIO (ou aussi les tagger si tu veux)
  lines.push(`STAT_RATIO;Ratio_${pk}_sur_Matin_${g}_Ncap;${agg.Ncap}`);
  lines.push(`STAT_RATIO;Ratio_${pk}_sur_Matin_${g}_RminRaw;${Number.isFinite(rmin) ? fmtFRNumber(rmin) : ""}`);
  lines.push(`STAT_RATIO;Ratio_${pk}_sur_Matin_${g}_RmaxRaw;${Number.isFinite(rmax) ? fmtFRNumber(rmax) : ""}`);
  lines.push(`STAT_RATIO;Ratio_${pk}_sur_Matin_${g}_N_EmatinLow;${agg.N_EmatinLow}`);
    }
  }

  // STAT_ANNUELLE (tag colonne A filtrable)
lines.push("SECTION;STAT_ANNUELLE");
const yearsStat = Object.keys(statsAnnees).map(Number).sort((a,b)=>a-b);

for (let i = 0; i < yearsStat.length; i++) {
  const y = yearsStat[i];

  const A = finalizeAgg(statsAnnees[y].Annee);
  const E = finalizeAgg(statsAnnees[y].Ete);
  const H = finalizeAgg(statsAnnees[y].HorsEte);

  if (A) {
    lines.push(`STAT_ANNEE;Moy_Annee_${y}_N${A.N};${fmtFRNumber(A.mean)}`);
    lines.push(`STAT_ANNEE;Sigma_Annee_${y};${fmtFRNumber(A.sigma)}`);
    lines.push(`STAT_ANNEE;CV_Annee_${y};${fmtFRNumber(A.cv)}`);
    lines.push(`STAT_ANNEE;Pct_lt_0p1_Annee_${y};${fmtFRNumber(A.pct)}`);
  }
  if (E) {
    lines.push(`STAT_ETE;Moy_Ete_${y}_N${E.N};${fmtFRNumber(E.mean)}`);
    lines.push(`STAT_ETE;Sigma_Ete_${y};${fmtFRNumber(E.sigma)}`);
    lines.push(`STAT_ETE;CV_Ete_${y};${fmtFRNumber(E.cv)}`);
    lines.push(`STAT_ETE;Pct_lt_0p1_Ete_${y};${fmtFRNumber(E.pct)}`);
  }
  if (H) {
    lines.push(`STAT_HORSETE;Moy_HorsEte_${y}_N${H.N};${fmtFRNumber(H.mean)}`);
    lines.push(`STAT_HORSETE;Sigma_HorsEte_${y};${fmtFRNumber(H.sigma)}`);
    lines.push(`STAT_HORSETE;CV_HorsEte_${y};${fmtFRNumber(H.cv)}`);
    lines.push(`STAT_HORSETE;Pct_lt_0p1_HorsEte_${y};${fmtFRNumber(H.pct)}`);
  }
}

  // TENDANCE
  lines.push("SECTION;TENDANCE");
  if(Trend_K>=3 && Number.isFinite(Trend_Slope) && Number.isFinite(Trend_R2)){
    lines.push(`STAT_TREND;Trend_Annee_Slope_VmParAn_N${Trend_K};${fmtFRNumber(Trend_Slope)}`);
    lines.push(`STAT_TREND;Trend_Annee_R2_N${Trend_K};${fmtFRNumber(Trend_R2)}`);
  }

  // STAT_GLOBAL
  lines.push("SECTION;STAT_GLOBAL");
  if(N_global>0){
    lines.push(`STAT_GLOBAL;Moy_Global_N${N_global};${fmtFRNumber(Moy_Global)}`);
    lines.push(`STAT_GLOBAL;Sigma_Global;${fmtFRNumber(Sigma_Global)}`);
    lines.push(`STAT_GLOBAL;CV_Global;${fmtFRNumber(CV_Global)}`);
    lines.push(`STAT_GLOBAL;Pct_lt_0p1_Global;${fmtFRNumber(Pct_lt_0p1_Global)}`);
  }

  // ======================= STAT_HOURP95_ALL =======================
  lines.push("SECTION;STAT_HOURP95_ALL");
  lines.push(`STAT_HOURP95_ALL_MAIN;Top1_Heure_ge_P95;${bestP95_ALL.h1}`);
  lines.push(`STAT_HOURP95_ALL_MAIN;Top1_Count;${bestP95_ALL.c1}`);
  lines.push(`STAT_HOURP95_ALL_MAIN;Top2_Heure_ge_P95;${bestP95_ALL.h2}`);
  lines.push(`STAT_HOURP95_ALL_MAIN;Top2_Count;${bestP95_ALL.c2}`);
  lines.push(`STAT_HOURP95_ALL_MAIN;Top1_PctJours;${nbJoursP95_ALL > 0 ? fmtFRNumber(100 * bestP95_ALL.c1 / nbJoursP95_ALL) : ""}`);
  lines.push(`STAT_HOURP95_ALL_MAIN;Top2_PctJours;${nbJoursP95_ALL > 0 ? fmtFRNumber(100 * bestP95_ALL.c2 / nbJoursP95_ALL) : ""}`);
  lines.push(`STAT_HOURP95_ALL;NbJours_P95;${nbJoursP95_ALL}`);
  lines.push(`STAT_HOURP95_ALL;NbPoints_ge_P95;${nbPointsP95_ALL}`);
  lines.push(`STAT_HOURP95_ALL;NbJours_avec_ge_2_points;${nbJoursP95_Exces_ALL}`);
  for (let h = 0; h < 24; h++) lines.push(`STAT_HOURP95_ALL;H${String(h).padStart(2,"0")};${histP95_ALL[h]}`);
  
  // ======================= STAT_HOURP95_JO ========================
  lines.push("SECTION;STAT_HOURP95_JO");
  lines.push(`STAT_HOURP95_JO_MAIN;Top1_Heure_ge_P95;${bestP95_JO.h1}`);
  lines.push(`STAT_HOURP95_JO_MAIN;Top1_Count;${bestP95_JO.c1}`);
  lines.push(`STAT_HOURP95_JO_MAIN;Top2_Heure_ge_P95;${bestP95_JO.h2}`);
  lines.push(`STAT_HOURP95_JO_MAIN;Top2_Count;${bestP95_JO.c2}`);
  lines.push(`STAT_HOURP95_JO_MAIN;Top1_PctJours;${nbJoursP95_JO > 0 ? fmtFRNumber(100 * bestP95_JO.c1 / nbJoursP95_JO) : ""}`);
  lines.push(`STAT_HOURP95_JO_MAIN;Top2_PctJours;${nbJoursP95_JO > 0 ? fmtFRNumber(100 * bestP95_JO.c2 / nbJoursP95_JO) : ""}`);
  lines.push(`STAT_HOURP95_JO;NbJours_P95;${nbJoursP95_JO}`);
  lines.push(`STAT_HOURP95_JO;NbPoints_ge_P95;${nbPointsP95_JO}`);
  lines.push(`STAT_HOURP95_JO;NbJours_avec_ge_2_points;${nbJoursP95_Exces_JO}`);
  for (let h = 0; h < 24; h++) lines.push(`STAT_HOURP95_JO;H${String(h).padStart(2,"0")};${histP95_JO[h]}`);
  
  // ======================= STAT_HOURP95_WE ========================
  lines.push("SECTION;STAT_HOURP95_WE");
  lines.push(`STAT_HOURP95_WE_MAIN;Top1_Heure_ge_P95;${bestP95_WE.h1}`);
  lines.push(`STAT_HOURP95_WE_MAIN;Top1_Count;${bestP95_WE.c1}`);
  lines.push(`STAT_HOURP95_WE_MAIN;Top2_Heure_ge_P95;${bestP95_WE.h2}`);
  lines.push(`STAT_HOURP95_WE_MAIN;Top2_Count;${bestP95_WE.c2}`);
  lines.push(`STAT_HOURP95_WE_MAIN;Top1_PctJours;${nbJoursP95_WE > 0 ? fmtFRNumber(100 * bestP95_WE.c1 / nbJoursP95_WE) : ""}`);
  lines.push(`STAT_HOURP95_WE_MAIN;Top2_PctJours;${nbJoursP95_WE > 0 ? fmtFRNumber(100 * bestP95_WE.c2 / nbJoursP95_WE) : ""}`);
  lines.push(`STAT_HOURP95_WE;NbJours_P95;${nbJoursP95_WE}`);
  lines.push(`STAT_HOURP95_WE;NbPoints_ge_P95;${nbPointsP95_WE}`);
  lines.push(`STAT_HOURP95_WE;NbJours_avec_ge_2_points;${nbJoursP95_Exces_WE}`);
  for (let h = 0; h < 24; h++) lines.push(`STAT_HOURP95_WE;H${String(h).padStart(2,"0")};${histP95_WE[h]}`);
  
}

// ----- DONNEES -----
lines.push("SECTION;DONNEES");
lines.push("DATA;DateHeure;Exposition_Vm");

decodedHourly.forEach(d=>{
  lines.push(`DATA;${fmtFRDate(new Date(d[0]))};${fmtFRNumber(d[1])}`);
});

// ----- AUDIT -----
lines.push("SECTION;AUDIT");
audit.forEach(a=>lines.push(a));

// --------------------------
// Export
// --------------------------
downloadFileUserClick(baseName + ".csv", lines.join("\n"), "Télécharger CSV");

})();
