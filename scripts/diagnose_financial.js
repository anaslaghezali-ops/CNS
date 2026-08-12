/* Diagnose financial reconciliation gaps */
var XLSX = require("../docs/vendor/xlsx.full.min.js");
global.XLSX = XLSX;
var CNS = require("../docs/reconcile.js");
var fs = require("fs");
var path = require("path");

var UP = "/home/ubuntu/.cursor/projects/workspace/uploads";
var posBuf = fs.readFileSync(path.join(UP, "mahaal_sales_history__10__3b98.xlsx"));
var glovoBuf = fs.readFileSync(path.join(UP, "orderDetails__9__2548.xlsx"));
var siteBuf = fs.readFileSync(path.join(UP, "orders_1786548042_d43a.xlsx"));
var napsBuf = fs.readFileSync(path.join(UP, "0007235524_MON_RELEVE_NAPSPRO_DU_2026-08-01_AU_2026-08-12_3a25.xlsx"));

function firstSheet(buf) {
  var wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  return wb.Sheets[wb.SheetNames[0]];
}

var pos = CNS.loadPOS(firstSheet(posBuf));
var glovo = CNS.loadGlovo(firstSheet(glovoBuf));
var site = CNS.loadSite(firstSheet(siteBuf));
var naps = CNS.loadNAPS(firstSheet(napsBuf));

var result = CNS.run(pos, glovo, naps, site);
var fin = result.summary.financial;
var anomalies = result.anomalies;

console.log("=== FINANCIAL LINES ===");
fin.lines.forEach(function (l) {
  var contribs = CNS.getFinancialContributors(l.lineKey, anomalies);
  var sum = CNS.sumFinancialContributions(l.lineKey, contribs);
  console.log(l.source, "ecart=" + l.ecart, "contribSum=" + sum, "nContrib=" + contribs.length);
});

var lineKey = "glovo_online";
var line = fin.lines.filter(function (l) { return l.lineKey === lineKey; })[0];
var contribs = CNS.getFinancialContributors(lineKey, anomalies);
console.log("\n=== GLOVO ONLINE CONTRIBUTORS (" + contribs.length + ") ===");
contribs.forEach(function (a) {
  var imp = CNS.ecartContributionForAnomaly(a, lineKey);
  console.log(imp.toFixed(1), a.type, a.source_ref || "", a.payment_pos || "", a.payment_source || "");
});

// Related anomalies with zero contribution
console.log("\n=== RELATED ANOMALIES ZERO CONTRIB ===");
anomalies.forEach(function (a) {
  if (a.source !== "Glovo") return;
  var imp = CNS.ecartContributionForAnomaly(a, lineKey);
  if (Math.abs(imp) < 0.01) {
    var rel = false;
    if (a.type.indexOf("Glovo") >= 0) rel = true;
    if (a.payment_pos === "Bank Transfer") rel = true;
    if (a.payment_source === "Bank Transfer") rel = true;
    if (a.type === "Commande Glovo absente du POS" && a.payment_source !== "Cash") rel = true;
    if (rel) console.log(a.type, a.detail.slice(0, 100));
  }
});

// List absent online + orphan BT
console.log("\n=== ABSENT GLOVO ONLINE ===");
anomalies.filter(function (a) {
  return a.type === "Commande Glovo absente du POS" && a.payment_source !== "Cash";
}).forEach(function (a) {
  console.log("+", a.amount_source, a.source_ref);
});

console.log("\n=== ORPHAN POS BT ===");
anomalies.filter(function (a) {
  return a.type === "Ticket Glovo au POS sans commande correspondante" && a.payment_pos === "Bank Transfer";
}).forEach(function (a) {
  console.log("-", a.amount_pos, a.ticket_name, a.pos_ticket_no);
});

console.log("\n=== AMOUNT MISMATCH ONLINE ===");
anomalies.filter(function (a) {
  return a.type === "Écart de montant" && a.source === "Glovo" && a.payment_source !== "Cash";
}).forEach(function (a) {
  var imp = CNS.ecartContributionForAnomaly(a, lineKey);
  console.log(imp.toFixed(1), a.amount_source, "vs", a.amount_pos, a.source_ref);
});

console.log("\n=== PAYMENT INCORRECT ===");
anomalies.filter(function (a) {
  return a.type === "Mode de paiement incorrect";
}).forEach(function (a) {
  var imp = CNS.ecartContributionForAnomaly(a, lineKey);
  console.log("online", imp.toFixed(1), a.payment_source, "->", a.payment_pos, a.source_ref);
});

console.log("\n=== DUPLICATES GLOVO ===");
anomalies.filter(function (a) {
  return a.type === "Ticket en double (correction)" && a.source === "Glovo";
}).forEach(function (a) {
  var imp = CNS.ecartContributionForAnomaly(a, lineKey);
  console.log("online", imp.toFixed(1), a.payment_pos, a.amount_pos, a.source_ref);
});
