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

function traiterHTML(doc) {
  const rows = [...doc.querySelectorAll("table tr")];
  currentData = rows.map((row) =>
    [...row.querySelectorAll("td, th")].map((cell) => cell.textContent.trim()),
  );
  currentData = calculerPrix(currentData);
  afficherResultat(currentData);
  document.getElementById("exportBtn").style.display = "inline-block";
}

function calculerPrix(data) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // normalisation à minuit

  // Ajout des en-têtes supplémentaires (ligne 2 en VBA => index 1)
  if (data[1]) {
    data[1].push(
      "Prix Dégagement",
      "% appliqué",
      "Samedi",
      "Dimanche",
      "Fériés",
      "Ajustement",
    );
  }

  // Parcourt les lignes de données à partir de l'index 2 (ligne 3 Excel)
  for (let i = 2; i < data.length; i++) {
    const dlcStr = data[i][3]; // DLC en colonne D (index 3)
    const prixStr = data[i][8]; // Prix en colonne I (index 8)

    const dlc = parseDateFR(dlcStr);
    if (!dlc) {
      data[i].push("", "", 0, 0, 0, 0);
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
    if (!isNaN(prix)) {
      prixDegagement = (prix * reduction).toFixed(2);
      pourcentage = ((1 - reduction) * 100).toFixed(0) + "%";
    } else {
      prixDegagement = "Prix manquant";
      pourcentage = "";
    }

    // Ajout des colonnes calculées
    data[i].push(
      prixDegagement,
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
  let html = "<table>";
  data.forEach((row) => {
    html += "<tr>" + row.map((cell) => `<td>${cell}</td>`).join("") + "</tr>";
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
  const colBaseCo = 8;        // Colonne H = Base Co
  const colPrixDegagement = data[1].length - 6; // Prix Dégagement ajouté en premier des colonnes calculées

  // Fonction pour sécuriser les champs CSV
  const safeRow = row => {
    const subset = [
      row[colCode] ?? "",
      row[colDesignation] ?? "",
      row[colDLC] ?? "",
      row[colBaseCo] ?? "",
      row[colPrixDegagement] ?? ""
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
  const filteredData = data.slice(1);

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


