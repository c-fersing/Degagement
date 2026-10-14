let currentData = [];

// ======================================================
// CONFIGURATION METIER
// ======================================================

const CONFIG = {

  // Colonnes du fichier source importé
  COL: {
    CODE: 0,                       // Code article
    DESIGNATION: 1,                // Désignation
    GAMME: 2,                      // Gamme
    DLC: 3,                        // Date Limite de Consommation
    PCBV: 4,                       // PCBV
    UNITE: 5,                      // Unité
    COLIS: 6,                      // Colis
    PIECES: 7,                     // Pièces
    BASE_CO: 8,                    // Base Co
    PRIX: 9,                       // Base 8 / Prix utilisé pour le calcul
    STOCK: 10,                     // Stock
    UNITE_FACTURATION: 11,         // Unité de facturation vente

    // Colonnes calculées ajoutées par calculerPrix()
    PRIX_DEGAGEMENT: 12,           // Prix après application du barème
    PRIX_TRONQUE: 13,              // Prix de vente final
    POURCENTAGE: 14,               // % appliqué
    SAMEDI: 15,                    // Nombre de samedis
    DIMANCHE: 16,                  // Nombre de dimanches
    FERIES: 17,                    // Nombre de jours fériés
    AJUSTEMENT: 18                 // Réajustement fériés/week-end
  },

  // Barème de réduction
  BAREME: [
    { jour: 3, reduc: 0.5 },
    { jour: 4, reduc: 0.6 },
    { jour: 5, reduc: 0.7 },
    { jour: 6, reduc: 0.8 },
    { jour: 10, reduc: 0.9 }
  ],

  // Colonnes exportées dans le fichier mail magasin
  EXPORT_MAIL_COLS: [
    "CODE",
    "DESIGNATION",
    "GAMME",
    "DLC",
    "PCBV",
    "UNITE",
    "COLIS",
    "PIECES",
    "PRIX_TRONQUE",
    "UNITE_FACTURATION"
  ],

  // Largeurs des colonnes XLSX pour l'export mail
  EXPORT_MAIL_WIDTHS: [
    15, // Code
    60, // Désignation
    10, // Gamme
    20, // DLC
    10, // PCBV
    13, // Unité
    10, // Colis
    10, // Pièces
    15, // Prix vente
    20  // Unité facturation vente
  ]
};

// Liste dynamique des jours fériés
const joursFeriesCache = new Map();

document
  .getElementById("fileInput")
  .addEventListener("change", function (event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function (e) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(e.target.result, "text/html");
      traiterHTML(doc);
    };
    reader.readAsText(file);
  });

// Gestion du drag & drop
const dropZone = document.getElementById("dropZone");

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");

  const file = e.dataTransfer.files[0];
  if (file && file.name.endsWith(".html")) {
    const reader = new FileReader();
    reader.onload = function(ev) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(ev.target.result, "text/html");
      traiterHTML(doc); // réutilise ta fonction existante
    };
    reader.readAsText(file);
  } else {
    alert("Veuillez déposer un fichier HTML valide.");
  }
});

function traiterHTML(doc) {
  const rows = [...doc.querySelectorAll("table tr")];
  currentData = rows.map((row) =>
    [...row.querySelectorAll("td, th")].map((cell) => cell.textContent.trim()),
  );

  // 🔹 Suppression de la première ligne (titre)
  if (currentData.length > 0) {
    currentData.shift();
  }

  currentData = calculerPrix(currentData);
  afficherResultat(currentData);
  document.getElementById("exportBtn").style.display = "inline-block";
  document.getElementById("exportBtnXls").style.display = "inline-block";
  document.getElementById("exportBtnXlsMail").style.display = "inline-block";
  document.getElementById("toggleBtn").style.display = "inline-block";

}


function calculerPrix(data) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // normalisation à minuit

  // Ajout des en-têtes supplémentaires (après retrait du titre, l’en-tête est à l’index 0)
  if (data[0]) {
    data[0].push(
      "Prix Dégagement",
      "Prix tronqué",
      "% appliqué",
      "Samedi",
      "Dimanche",
      "Fériés",
      "Ajustement",
    );
  }

  // Parcourt les lignes de données à partir de l'index 2 (ligne 3 Excel)
  for (let i = 1; i < data.length; i++) {
    const dlcStr = data[i][CONFIG.COL.DLC]; 
    const prixStr = data[i][CONFIG.COL.PRIX]; 

    const dlc = parseDateFR(dlcStr);
    if (!dlc) {
      data[i].push("", "", "", 0, 0, 0, 0);
      continue;
    }
    dlc.setHours(0, 0, 0, 0);

    // Écart en jours (UTC)
    let ecart = daysBetween(today, dlc);

    // Si écart négatif → périmé
    if (ecart < 0) {
      data[i].push("Périmé", "---", 0, 0, 0, 0);
      continue;
    }

    // Comptage inclusif des jours entre today et dlc
    let nbSamedi = 0,
      nbDimanche = 0,
      nbFeries = 0,
      nbAjustement = 0;
    const dateDebut =
      today.getTime() <= dlc.getTime() ? new Date(today) : new Date(dlc);
    const dateFin =
      today.getTime() <= dlc.getTime() ? new Date(dlc) : new Date(today);

    for (
      let d = new Date(dateDebut);
      d.getTime() <= dateFin.getTime();
      d.setDate(d.getDate() + 1)
    ) {
      d.setHours(0, 0, 0, 0);
      const jour = d.getDay(); // 0=dimanche, 6=samedi
      const estFerie = isHoliday(d);

      // Un jour férié tombant un week-end a déjà été décompté
      // via le week-end. On le réajoute donc pour éviter un double retrait
      if (jour === 6) {
        nbSamedi++;
        if (estFerie) nbAjustement++;
      } else if (jour === 0) {
        nbDimanche++;
        if (estFerie) nbAjustement++;
      }
      if (estFerie) nbFeries++;
    }

    // Ajustement de l'écart comme en VBA
    ecart = ecart - nbSamedi - nbDimanche + nbAjustement;

    // Application du barème
    const reduction = appliquerBareme(ecart);

    // Prix
    const prix = parseFloat(String(prixStr).replace(",", "."));
    let prixDegagement = "";
    let pourcentage = "";
    let prixTronque = "";
    
    if (!isNaN(prix)) {
      const prixDegagementNum = Math.floor(prix * reduction * 100) / 100;
    
      prixDegagement = prixDegagementNum.toFixed(2);
      pourcentage = ((1 - reduction) * 100).toFixed(0) + "%";
      prixTronque = (Math.floor(prixDegagementNum * 10) / 10).toFixed(2);
    
    } else {
      prixDegagement = "Prix manquant";
      pourcentage = "";
      prixTronque = "";
    }

    // Ajout des colonnes calculées
    data[i].push(
      prixDegagement,
      prixTronque,
      pourcentage,
      nbSamedi,
      nbDimanche,
      nbFeries,
      nbAjustement,
    );
  }

  return data;
}

// Helpers
function parseDateFR(str) {
  if (!str) return null;
  const s = String(str).trim();

  // Format dd/mm/yyyy ou dd-mm-yyyy
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    const yyyy = parseInt(m[3], 10);
    return new Date(yyyy, mm - 1, dd); // toujours interprété correctement
  }

  // Si déjà au format ISO (2025-12-01), on laisse passer
  const iso = new Date(s);
  if (!isNaN(iso.getTime())) return iso;

  return null;
}

function daysBetween(a, b) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const utcA = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const utcB = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((utcB - utcA) / msPerDay);
}

function calculerPaques(annee) {
  const a = annee % 19;
  const b = Math.floor(annee / 100);
  const c = annee % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);

  const mois = Math.floor((h + l - 7 * m + 114) / 31);
  const jour = ((h + l - 7 * m + 114) % 31) + 1;

  return new Date(annee, mois - 1, jour);
}

function getJoursFeries(annee) {
  const feries = [];

  // Fêtes fixes
  feries.push(
    new Date(annee, 0, 1).setHours(0, 0, 0, 0),
    new Date(annee, 4, 1).setHours(0, 0, 0, 0),
    new Date(annee, 4, 8).setHours(0, 0, 0, 0),
    new Date(annee, 6, 14).setHours(0, 0, 0, 0),
    new Date(annee, 7, 15).setHours(0, 0, 0, 0),
    new Date(annee, 10, 1).setHours(0, 0, 0, 0),
    new Date(annee, 10, 11).setHours(0, 0, 0, 0),
    new Date(annee, 11, 25).setHours(0, 0, 0, 0)
  );

  const paques = calculerPaques(annee);

  const lundiPaques = new Date(paques);
  lundiPaques.setDate(lundiPaques.getDate() + 1);

  const ascension = new Date(paques);
  ascension.setDate(ascension.getDate() + 39);

  const pentecote = new Date(paques);
  pentecote.setDate(pentecote.getDate() + 50);

  feries.push(
    lundiPaques.setHours(0, 0, 0, 0),
    ascension.setHours(0, 0, 0, 0),
    pentecote.setHours(0, 0, 0, 0)
  );

  return new Set(feries);
}

function isHoliday(dateObj) {
  const annee = dateObj.getFullYear();

  if (!joursFeriesCache.has(annee)) {
    joursFeriesCache.set(
      annee,
      getJoursFeries(annee)
    );
  }

  const t = new Date(dateObj);
  t.setHours(0, 0, 0, 0);

  return joursFeriesCache
    .get(annee)
    .has(t.getTime());
}

function appliquerBareme(ecart) {
  if (ecart < 0) return null;

  for (const ligne of CONFIG.BAREME) {
    if (ecart <= ligne.jour) {
      return ligne.reduc;
    }
  }

  return 1;
}

function afficherResultat(data) {
  const container = document.getElementById("result");
  let html = '<table id="myTable">';

  data.forEach((row, index) => {

    // Détection des lignes d'en-tête (première cellule = "Code")
    const isHeader = row[CONFIG.COL.CODE] === "Code";

    // Ligne grise toutes les 5 lignes
    const grey = (index % 5 === 0) ? ' class="row-grey"' : '';

    // Ajout de la classe header-row si c'est un header
    const headerClass = isHeader ? ' class="header-row"' : '';

    html += `<tr${grey}${headerClass}>` +
      row.map((cell, colIndex) => {
        const isblue = (index > 0 && colIndex === CONFIG.COL.PRIX_TRONQUE) ? ' class="blueprice"' : '';
        return `<td${isblue}>${cell}</td>`;
      }).join("") +
      "</tr>";
  });

  html += "</table>";
  container.innerHTML = html;
}



document.getElementById("exportBtn").addEventListener("click", function () {
  exporterCSV(currentData);
});

function exporterCSV(data) {

  // Fonction pour sécuriser les champs CSV
  const safeRow = row => {
    const subset = [
      row[CONFIG.COL.CODE] ?? "",
      row[CONFIG.COL.DESIGNATION] ?? "",
      row[CONFIG.COL.DLC] ?? "",
      row[CONFIG.COL.PRIX] ?? "",
      // conversion du prix : nombre → string avec virgule
      (row[CONFIG.COL.PRIX_DEGAGEMENT] != null
      ? String(row[CONFIG.COL.PRIX_DEGAGEMENT]).replace(".", ",")
      : ""),
      (row[CONFIG.COL.PRIX_TRONQUE] != null
      ? String(row[CONFIG.COL.PRIX_TRONQUE]).replace(".", ",")
      : ""),
  ];
    return subset.map(val => {
      const s = String(val);
      if ([",", "\"", "\n", "\r"].some(ch => s.includes(ch))) {
        return "\"" + s.replace(/"/g, "\"\"") + "\"";
      }
      return s;
    }).join(",");
  };

  // Conserver uniquement l’en-tête (index 1) et les données (index ≥ 2)
  const filteredData = data.slice(0);

  // Générer le CSV avec BOM UTF‑8 pour gérer les accents
  const csvContent = "\uFEFF" + filteredData.map(safeRow).join("\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = "resultat.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
}

function toggleColumns(indexes) {
  const table = document.getElementById("myTable");
  const rows = table.querySelectorAll("tr");

  rows.forEach(row => {
    const cells = row.querySelectorAll("th, td");
    indexes.forEach(i => {
      if (cells[i]) {
        cells[i].classList.toggle("hide-col");
      }
    });
  });
}

document.getElementById("exportBtnXls").addEventListener("click", function () {
  exporterXLSX(currentData);
});

function exporterXLSX(data) {
  const finalData = [];

  // --- Style de l'en-tête ---
  const headerStyle = {
    fill: { fgColor: { rgb: "D9D9D9" } }, // gris clair
    font: { bold: true },
    alignment: { horizontal: "center", vertical: "center" },
    border: {
      top:    { style: "medium", color: { rgb: "000000" } },
      bottom: { style: "medium", color: { rgb: "000000" } },
      left:   { style: "medium", color: { rgb: "000000" } },
      right:  { style: "medium", color: { rgb: "000000" } }
    }
  };

  // --- Style normal pour les données ---
  const normalStyle = {
    alignment: { horizontal: "center", vertical: "center" },
    font: { bold: true},
    border: {
      top:    { style: "medium", color: { rgb: "000000" } },
      bottom: { style: "medium", color: { rgb: "000000" } },
      left:   { style: "medium", color: { rgb: "000000" } },
      right:  { style: "medium", color: { rgb: "000000" } }
    }
  };

    // --- Style rouge pour le prix tronqué ---
  const redStyle = {
    alignment: { horizontal: "center", vertical: "center" },
    font: { bold: true, color: { rgb: "FF0000"} },
    border: {
      top:    { style: "medium", color: { rgb: "FF0000" } },
      bottom: { style: "medium", color: { rgb: "FF0000" } },
      left:   { style: "medium", color: { rgb: "FF0000" } },
      right:  { style: "medium", color: { rgb: "FF0000" } }
    }
  };

  // --- Ligne d'en-tête (data[0]) ---
  finalData.push([
    { v: data[0][CONFIG.COL.CODE], s: headerStyle },
    { v: data[0][CONFIG.COL.DESIGNATION], s: headerStyle },
    { v: data[0][CONFIG.COL.DLC], s: headerStyle },
    { v: data[0][CONFIG.COL.PRIX], s: headerStyle },
    { v: data[0][CONFIG.COL.PRIX_DEGAGEMENT], s: headerStyle },
    { v: data[0][CONFIG.COL.PRIX_TRONQUE], s: headerStyle }
  ]);

  // --- Booléen pour savoir si la répétition est passée ---
  let repeatPassed = false;

  // --- Données ---
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Détection d'une ligne d'entête répétée
    const firstCell = String(row[CONFIG.COL.CODE] ?? "").trim();
    const isHeaderRepeat = /^[A-Za-z]/.test(firstCell);

    let style;

    if (isHeaderRepeat) {
      style = headerStyle;
      repeatPassed = true;
    } else {
      style = repeatPassed ? normalStyle : redStyle;
    }

    finalData.push([
      { v: row[CONFIG.COL.CODE] ?? "", s: style },
      { v: row[CONFIG.COL.DESIGNATION] ?? "" ,s: style},
      { v: row[CONFIG.COL.DLC] ?? "", s: style },
      { v: cleanNumber(row[CONFIG.COL.PRIX]), s: style },
      { v: cleanNumber(row[CONFIG.COL.PRIX_DEGAGEMENT]), s: style },
      { v: cleanNumber(row[CONFIG.COL.PRIX_TRONQUE]), s: style }
    ]);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(finalData);

  // Largeurs de colonnes
  ws['!cols'] = [
    { wch: 8 },
    { wch: 48 },
    { wch: 12 },
    { wch: 10 },
    { wch: 12 },
    { wch: 12 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Résultats");
  XLSX.writeFile(wb, "resultat.xlsx");
}

function cleanNumber(value) {
  if (!value) return "";

  const s = String(value).trim();

  // Si la valeur contient une virgule ou un point, on tente une conversion
  const normalized = s.replace(",", ".");
  const num = parseFloat(normalized);

  // Si c'est un vrai nombre → on le renvoie
  if (!isNaN(num)) return num;

  // Sinon → valeur métier → on la laisse telle quelle
  return s;
}

document.getElementById("exportBtnXlsMail").addEventListener("click", function () {
  exporterXLSX2(currentData);
});

function parseFrenchDate(str) {
  if (!str) return null;
  const parts = str.split("/");
  if (parts.length !== 3) return null;
  const [jour, mois, annee] = parts.map(p => parseInt(p, 10));
  if (!jour || !mois || !annee) return null;
  return new Date(annee, mois - 1, jour);
}

// Conversion Date JS -> numéro de série Excel
function toExcelDate(date) {
  const excelEpoch = new Date(1899, 11, 30); // 30/12/1899
  const diffMs = date - excelEpoch;
  return diffMs / (24 * 60 * 60 * 1000);
}

function exporterXLSX2(data) {
  // Colonnes à exporter
  const cols = CONFIG.EXPORT_MAIL_COLS.map(
    col => CONFIG.COL[col]
  );

  // Largeurs correspondantes
  const colWidths = CONFIG.EXPORT_MAIL_WIDTHS;

  const finalData = [];

  // --- Styles identiques à ta version actuelle ---
  const headerStyle = {
    fill: { fgColor: { rgb: "D9D9D9" } },
    font: { bold: true },
    alignment: { horizontal: "center", vertical: "center" },
    border: {
      top:    { style: "medium", color: { rgb: "000000" } },
      bottom: { style: "medium", color: { rgb: "000000" } },
      left:   { style: "medium", color: { rgb: "000000" } },
      right:  { style: "medium", color: { rgb: "000000" } }
    }
  };

  const normalStyle = {
    alignment: { horizontal: "center", vertical: "center" },
    font: { bold: true },
    border: {
      top:    { style: "medium", color: { rgb: "000000" } },
      bottom: { style: "medium", color: { rgb: "000000" } },
      left:   { style: "medium", color: { rgb: "000000" } },
      right:  { style: "medium", color: { rgb: "000000" } }
    }
  };

  const redStyle = {
    alignment: { horizontal: "center", vertical: "center" },
    font: { bold: true, color: { rgb: "FF0000"} },
    border: {
      top:    { style: "medium", color: { rgb: "FF0000" } },
      bottom: { style: "medium", color: { rgb: "FF0000" } },
      left:   { style: "medium", color: { rgb: "FF0000" } },
      right:  { style: "medium", color: { rgb: "FF0000" } }
    }
  };

  // --- Ligne d'entête ---
  finalData.push(
    cols.map(c => ({
      v: c === CONFIG.COL.PRIX_TRONQUE ? "Prix vente" : data[0][c],
      s: headerStyle
    }))
  );

  // --- Booléen pour savoir si la répétition est passée ---
  let repeatPassed = false;

  // --- Données ---
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const firstCell = String(row[cols[0]] ?? "").trim();
    const isHeaderRepeat = /^[A-Za-z]/.test(firstCell);

    if (isHeaderRepeat) {
      repeatPassed = true;

      finalData.push(
        cols.map(c => ({
          v: c === CONFIG.COL.PRIX_TRONQUE ? "Prix vente" : row[c],
          s: headerStyle
        }))
      );

      continue;
    }

    const style = repeatPassed ? normalStyle : redStyle;

    finalData.push(
      cols.map(c => {
        let value = row[c];
        let format = undefined;

        // Colonne DLC : format date courte
        if (c === CONFIG.COL.DLC) {
          const d = parseFrenchDate(String(value).trim());
          if (d) {
            value = Math.floor(toExcelDate(d));
            return {
              v: value,
              t: "n",
              z: "dd/mm/yyyy",
              s: style
            };
          }
        }

        // Colonne prix tronqué : format monétaire
        if (c === CONFIG.COL.PRIX_TRONQUE) {
          const num = parseFloat(String(value).trim());
          if (!isNaN(num)) {
            return {
              v: num,
              t: "n",
              z: "#,##0.00",
              s: style
            };
          }
        }


        // Autres colonnes
        else {
          value = cleanNumber(value);
        }

        return {
          v: value,
          s: style,   // ton style rouge / normal / header
          z: format   // format Excel
        };
      })
    );
  }

  // --- Création du fichier ---
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(finalData);

  // Largeurs de colonnes
  ws['!cols'] = colWidths.map(w => ({ wch: w }));

  XLSX.utils.book_append_sheet(wb, ws, "Résultats");
  XLSX.writeFile(wb, "resultat.xlsx");
}

