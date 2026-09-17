let currentData = [];

// Barème
const bareme = [
  { jour: 3, reduc: 0.5 },
  { jour: 4, reduc: 0.6 },
  { jour: 5, reduc: 0.7 },
  { jour: 6, reduc: 0.8 },
  { jour: 10, reduc: 0.9 },
];

// Liste des jours fériés
const joursFeries = [
  "2025-01-01",
  "2025-04-21",
  "2025-05-01",
  "2025-05-08",
  "2025-05-29",
  "2025-06-09",
  "2025-07-14",
  "2025-08-15",
  "2025-11-01",
  "2025-11-11",
  "2025-12-25",
  "2026-01-01",
  "2026-04-06",
  "2026-05-01",
  "2026-05-08",
  "2026-05-14",
  "2026-05-25",
  "2026-07-14",
  "2026-08-15",
  "2026-11-01",
  "2026-11-11",
  "2026-12-25",
].map((d) => new Date(d).setHours(0, 0, 0, 0));

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
    const dlcStr = data[i][3]; // DLC en colonne D (index 3)
    // const prixStr = data[i][8]; // Prix en colonne I (index 8)
    const prixStr = data[i][9]; // Prix en colonne J (index 9) Base 8

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
      prixDegagement = (prix * reduction).toFixed(2);
      pourcentage = ((1 - reduction) * 100).toFixed(0) + "%";
      prixTronque = (Math.floor(parseFloat(prixDegagement) * 10) / 10).toFixed(2);

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

function isHoliday(dateObj) {
  const t = new Date(dateObj);
  t.setHours(0, 0, 0, 0);
  return joursFeries.includes(t.getTime());
}

function appliquerBareme(ecart) {
  // Si la DLC est passée (écart négatif), on signale "Périmé"
  if (ecart < 0) return null;

  if (ecart <= 3) return 0.5;
  if (ecart === 4) return 0.6;
  if (ecart === 5) return 0.7;
  if (ecart === 6) return 0.8;
  if (ecart <= 10) return 0.9;
  return 1; // au-delà de 10 jours → prix d'origine
}

function afficherResultat(data) {
  const container = document.getElementById("result");
  let html = '<table id="myTable">';

  data.forEach((row, index) => {

    // Détection des lignes d'en-tête (première cellule = "Code")
    const isHeader = row[0] === "Code";

    // Ligne grise toutes les 5 lignes
    const grey = (index % 5 === 0) ? ' class="row-grey"' : '';

    // Ajout de la classe header-row si c'est un header
    const headerClass = isHeader ? ' class="header-row"' : '';

    html += `<tr${grey}${headerClass}>` +
      row.map((cell, colIndex) => {
        const isblue = (index > 0 && colIndex === 13) ? ' class="blueprice"' : '';
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
  // Définir les index des colonnes à garder
  const colCode = 0;          // Colonne A = Code
  const colDesignation = 1;   // Colonne B = Designation
  const colDLC = 3;           // Colonne D = DLC
  const colBaseCo = 9;        // Colonne I = Base 8
  const colPrixDegagement = 12; // Prix Dégagement ajouté en premier des colonnes calculées
  const colPrixTronque = 13; // Prix Tronque soit la derniere colonne

  // Fonction pour sécuriser les champs CSV
  const safeRow = row => {
    const subset = [
      row[colCode] ?? "",
      row[colDesignation] ?? "",
      row[colDLC] ?? "",
      row[colBaseCo] ?? "",
      // conversion du prix : nombre → string avec virgule
      (row[colPrixDegagement] != null
      ? String(row[colPrixDegagement]).replace(".", ",")
      : ""),
      (row[colPrixTronque] != null
      ? String(row[colPrixTronque]).replace(".", ",")
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
  const colCode = 0;
  const colDesignation = 1;
  const colDLC = 3;
  const colBase8 = 9;
  const colPrixDegagement = 12;
  const colPrixTronque = 13;

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
    { v: data[0][colCode], s: headerStyle },
    { v: data[0][colDesignation], s: headerStyle },
    { v: data[0][colDLC], s: headerStyle },
    { v: data[0][colBase8], s: headerStyle },
    { v: data[0][colPrixDegagement], s: headerStyle },
    { v: data[0][colPrixTronque], s: headerStyle }
  ]);

  // --- Booléen pour savoir si la répétition est passée ---
  let repeatPassed = false;

  // --- Données ---
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Détection d'une ligne d'entête répétée
    const firstCell = String(row[colCode] ?? "").trim();
    const isHeaderRepeat = /^[A-Za-z]/.test(firstCell);

    let style;

    if (isHeaderRepeat) {
      style = headerStyle;
      repeatPassed = true;
    } else {
      style = repeatPassed ? normalStyle : redStyle;
    }

    finalData.push([
      { v: row[colCode] ?? "", s: style },
      { v: row[colDesignation] ?? "" ,s: style},
      { v: row[colDLC] ?? "", s: style },
      { v: cleanNumber(row[colBase8]), s: style },
      { v: cleanNumber(row[colPrixDegagement]), s: style },
      { v: cleanNumber(row[colPrixTronque]), s: style }
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
  const cols = [0, 1, 2, 3, 4, 5, 6, 7, 13, 11];

  // Largeurs correspondantes
  const colWidths = [15, 60, 10, 20, 10, 13, 10, 10, 15, 20];

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
      v: c === 13 ? "Prix vente" : data[0][c],
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
          v: c === 13 ? "Prix vente" : row[c],
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

        // Colonne 3 : date courte
        if (c === 3) {
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

        // Colonne 13 : monétaire
        if (c === 13) {
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


