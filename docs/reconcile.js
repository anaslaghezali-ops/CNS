/*
 * ChickNSter — Moteur de réconciliation (version navigateur, JavaScript pur).
 * Port fidèle du moteur Python. Aucune dépendance hors SheetJS (global `XLSX`).
 *
 * Fonctionne aussi en Node (pour les tests) : voir le garde module.exports en fin.
 */
(function (root) {
  "use strict";

  // ----------------------------------------------------------------------- //
  // Constantes / règles métier
  // ----------------------------------------------------------------------- //
  var CH_GLOVO = "Glovo";
  var CH_SITE = "Site";
  var CH_DINEIN = "Sur place / Emporter";
  var CH_UNASSIGNED = "À rattacher";
  var CH_OTHER = "Autre";

  var GLOVO_PAYMENT_MAP = { Online: "Bank Transfer", Cash: "Cash" };
  var SITE_EXPECTED_PAYMENT = "Bank Transfer";
  var DINEIN_ALLOWED = ["Cash", "Credit card"];
  var PAYS_POS = ["Cash", "Bank Transfer", "Credit card"];

  /** POS « Payment type » peut lister plusieurs modes : « Cash, Cash » ou « Cash, Credit card ». */
  function parsePaymentTypes(raw) {
    if (!raw) return [];
    return String(raw).split(",").map(function (s) { return s.trim(); }).filter(Boolean);
  }
  function isPurePaymentType(raw, expected) {
    var types = parsePaymentTypes(raw);
    if (!types.length) return String(raw).trim() === expected;
    return types.every(function (t) { return t === expected; });
  }
  function isAllDineinPayments(raw) {
    var types = parsePaymentTypes(raw);
    if (!types.length) return DINEIN_ALLOWED.indexOf(String(raw).trim()) >= 0;
    return types.every(function (t) { return DINEIN_ALLOWED.indexOf(t) >= 0; });
  }
  function isSplitCashCreditCard(raw) {
    var types = parsePaymentTypes(raw);
    if (types.length < 2) return false;
    var hasCash = false, hasCC = false;
    types.forEach(function (t) {
      if (t === "Cash") hasCash = true;
      if (t === "Credit card") hasCC = true;
    });
    return hasCash && hasCC && types.every(function (t) {
      return t === "Cash" || t === "Credit card";
    });
  }
  /** Ventile le total POS par mode (fractionnement Cash ou Cash+CB avec NAPS). */
  function allocatePosPaymentAmounts(p) {
    var t = isNaN(p.total) ? 0 : p.total;
    var out = { "Cash": 0, "Bank Transfer": 0, "Credit card": 0, "Autre": 0 };
    if (p._naps_split_cc != null && !isNaN(p._naps_split_cc)) {
      var cc = Math.min(t, Math.max(0, p._naps_split_cc));
      out["Credit card"] = cc;
      out["Cash"] = t - cc;
      return out;
    }
    var types = parsePaymentTypes(p.payment_type);
    if (!types.length) {
      var pay = PAYS_POS.indexOf(p.payment_type) >= 0 ? p.payment_type : "Autre";
      out[pay] = t;
      return out;
    }
    if (types.every(function (tp) { return tp === "Cash"; })) {
      out["Cash"] = t;
      return out;
    }
    if (types.every(function (tp) { return tp === "Credit card"; })) {
      out["Credit card"] = t;
      return out;
    }
    if (types.length === 1 && PAYS_POS.indexOf(types[0]) >= 0) {
      out[types[0]] = t;
      return out;
    }
    if (isSplitCashCreditCard(p.payment_type)) {
      out["Cash"] = t;
      return out;
    }
    out["Cash"] = t;
    return out;
  }
  function sumPosCcOnDay(pos, d) {
    var s = 0;
    pos.forEach(function (p) {
      if (dateKey(p.datetime) !== d) return;
      if (isPurePaymentType(p.payment_type, "Credit card")) {
        s += isNaN(p.total) ? 0 : p.total;
      } else if (p._naps_split_cc != null && !isNaN(p._naps_split_cc)) {
        s += p._naps_split_cc;
      }
    });
    return s;
  }

  var WINDOW_BEFORE_MIN = 1;
  var WINDOW_AFTER_MIN = 10;
  var AMOUNT_TOL = 0.5;

  // ----------------------------------------------------------------------- //
  // Helpers
  // ----------------------------------------------------------------------- //
  function s(v) { return v == null ? "" : String(v).trim(); }
  function num(v) {
    if (v == null || v === "") return NaN;
    if (typeof v === "number") return v;
    var x = parseFloat(String(v).replace(/\s/g, "").replace(",", "."));
    return isNaN(x) ? NaN : x;
  }

  function toDate(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === "number") {
      // Numéro de série Excel -> date (base 1899-12-30)
      var ms = Math.round((v - 25569) * 86400 * 1000);
      var d0 = new Date(ms);
      return new Date(d0.getUTCFullYear(), d0.getUTCMonth(), d0.getUTCDate(),
                      d0.getUTCHours(), d0.getUTCMinutes(), d0.getUTCSeconds());
    }
    var str = String(v).trim();
    var m = str.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
    var m2 = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m2) return new Date(+m2[1], +m2[2] - 1, +m2[3]);
    var d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  }

  function dateKey(d) {
    if (!d) return "";
    var y = d.getFullYear(), m = d.getMonth() + 1, dd = d.getDate();
    return y + "-" + (m < 10 ? "0" + m : m) + "-" + (dd < 10 ? "0" + dd : dd);
  }
  function minutesBetween(a, b) { return (a.getTime() - b.getTime()) / 60000; }
  function hhmm(d) {
    if (!d) return "";
    var h = d.getHours(), m = d.getMinutes();
    return (h < 10 ? "0" + h : h) + ":" + (m < 10 ? "0" + m : m);
  }

  function dtFull(d) {
    return d ? dateKey(d) + " " + hhmm(d) : "";
  }
  function anomalyId(a) {
    return [
      a.source, a.type, a.ticket_name || "", a.pos_ticket_no || "",
      a.source_ref || "",
      a.amount_pos == null ? "" : a.amount_pos,
      a.amount_source == null ? "" : a.amount_source,
      a.payment_pos || "", a.payment_source || "",
      a.file || "", a.row == null ? "" : a.row,
    ].join("|");
  }
  function anomaly(o) {
    var a = {
      source: o.source, severity: o.severity, type: o.type,
      ticket_name: o.ticket_name || "", pos_ticket_no: o.pos_ticket_no || "",
      pos_datetime: o.pos_datetime || null,
      source_ref: o.source_ref || "", detail: o.detail,
      amount_pos: o.amount_pos == null ? null : o.amount_pos,
      amount_source: o.amount_source == null ? null : o.amount_source,
      payment_pos: o.payment_pos || "", payment_source: o.payment_source || "",
      file: o.file || "", row: o.row == null ? "" : o.row,
      when: o.when || (o.pos_datetime ? dtFull(o.pos_datetime) : ""),
    };
    a.id = anomalyId(a);
    return a;
  }
  /** Champs POS pour lier une anomalie à UNE ligne caisse (ticket_no unique). */
  function posFields(p) {
    return {
      ticket_name: p.ticket_name,
      pos_ticket_no: p.ticket_no || "",
      pos_datetime: p.datetime,
      file: p.file || "POS",
      row: p.row,
      when: dtFull(p.datetime),
    };
  }
  /** Anomalie liée à une ligne POS précise (ne se mélange pas avec un autre ticket_name identique). */
  function posAnomaly(p, o) {
    var base = posFields(p);
    for (var k in o) if (o.hasOwnProperty(k)) base[k] = o[k];
    return anomaly(base);
  }

  // ----------------------------------------------------------------------- //
  // Lecture des feuilles -> lignes normalisées
  // ----------------------------------------------------------------------- //
  function sheetAOA(ws) {
    // blankrows:true -> l'index de ligne correspond au n° de ligne Excel (i+1).
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: true });
  }
  function findHeaderRow(aoa, tokens) {
    for (var i = 0; i < aoa.length; i++) {
      var cells = aoa[i].map(function (c) { return s(c); });
      for (var t = 0; t < tokens.length; t++) {
        if (cells.indexOf(tokens[t]) !== -1) return i;
      }
    }
    return -1;
  }
  function rowsWithHeader(ws, tokens) {
    var aoa = sheetAOA(ws);
    var hi = findHeaderRow(aoa, tokens);
    if (hi < 0) throw new Error("En-tête introuvable (attendus : " + tokens.join(", ") + ")");
    // Origine de la plage : aoa[0] correspond à la ligne Excel (r0 + 1).
    var r0 = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).s.r : 0;
    var headers = aoa[hi].map(function (h) { return s(h); });
    var out = [];
    for (var i = hi + 1; i < aoa.length; i++) {
      var r = aoa[i] || [], obj = { __row: r0 + i + 1 };  // n° de ligne Excel (1-based)
      for (var c = 0; c < headers.length; c++) obj[headers[c]] = r[c];
      out.push(obj);
    }
    return out;
  }

  function loadPOS(ws) {
    var raw = rowsWithHeader(ws, ["Ticket No.", "Ticket name"]);
    var out = [];
    raw.forEach(function (r) {
      if (s(r["Ticket No."]) === "") return;
      out.push({
        row: r.__row, file: "POS",
        ticket_no: s(r["Ticket No."]),
        date: s(r["Date"]),
        hour: s(r["Hour"]),
        user: s(r["User"]),
        ticket_name: s(r["Ticket name"]),
        designations: s(r["Designations (Reference)"]),
        total: num(r["Total"]),
        payment_type: s(r["Payment type"]),
        datetime: toDate(s(r["Date"]) + " " + s(r["Hour"])),
      });
    });
    return out;
  }

  function loadGlovo(ws) {
    var raw = rowsWithHeader(ws, ["Order ID", "Payment type"]);
    var out = [];
    raw.forEach(function (r) {
      if (s(r["Order ID"]) === "") return;
      out.push({
        row: r.__row, file: "Glovo",
        order_id: s(r["Order ID"]),
        payment_type: s(r["Payment type"]),
        status: s(r["Order status"]),
        received_at: toDate(r["Order received at"]),
        earnings: num(r["Estimated earnings"]),
        subtotal: num(r["Subtotal"]),
        discount_funded: num(r["Discount Funded by you"]),
        order_items: s(r["Order Items"]),
        // Montant rapprochement Glovo = col W − col AE — voir CURSOR_JOURNAL.md
        amount: (function () {
          var sub = num(r["Subtotal"]);
          if (isNaN(sub)) return NaN;
          var disc = num(r["Discount Funded by you"]);
          return sub - (isNaN(disc) ? 0 : disc);
        })(),
      });
    });
    return out;
  }

  function loadNAPS(ws) {
    var raw = rowsWithHeader(ws, ["Date de transaction", "Montant"]);
    var out = [];
    raw.forEach(function (r) {
      if (s(r["Montant"]) === "") return;
      var d = toDate(r["Date de transaction"]);
      out.push({ row: r.__row, file: "NAPS", date_transaction: d, date: dateKey(d),
                 montant: num(r["Montant"]) });
    });
    return out;
  }

  function loadSite(ws) {
    var raw = rowsWithHeader(ws, ["identifiant", "Order Total"]);
    var out = [];
    raw.forEach(function (r) {
      if (s(r["identifiant"]) === "") return;
      out.push({
        row: r.__row, file: "Site",
        identifiant: s(r["identifiant"]),
        created_at: toDate(r["Created At"]),
        last_status: s(r["Last Status"]),
        order_total: num(r["Order Total"]),
        delivery_method: s(r["Delivery Method"]),
        delivery_status: s(r["Delivery Provider Status"]),
      });
    });
    return out;
  }

  // ----------------------------------------------------------------------- //
  // Classification
  // ----------------------------------------------------------------------- //
  var RE_1_3 = /^\d{1,3}$/;
  var RE_5 = /^\d{5}$/;
  var RE_SPEMP = /^\d*(sp|emp)\d*$/i;

  function classify(name, siteIds) {
    var n = s(name);
    if (n === "" || n.toLowerCase() === "nan") return CH_UNASSIGNED;
    if (siteIds && siteIds.has(n)) return CH_SITE;
    // « Ticket » (placeholder) = numéro oublié par le caissier -> à rattacher.
    if (n.toLowerCase() === "ticket") return CH_UNASSIGNED;
    var compact = n.replace(/\s/g, "");
    if (RE_SPEMP.test(compact)) return CH_DINEIN;
    if (RE_1_3.test(n)) return CH_GLOVO;
    if (RE_5.test(n)) return CH_SITE;
    return CH_OTHER;
  }
  function addChannel(pos, siteIds) {
    pos.forEach(function (p) { p.channel = classify(p.ticket_name, siteIds); });
    return pos;
  }

  // ----------------------------------------------------------------------- //
  // SITE
  // ----------------------------------------------------------------------- //
  // Fenêtre pour détecter une faute de frappe sur le numéro de commande site.
  var SITE_TYPO_WINDOW_MIN = 20;
  var SITE_TYPO_MAX_EDITS = 2;   // 1 chiffre en trop/en moins/modifié (voire 2)

  function siteOrderIsTakeout(o) {
    var m = s(o.delivery_method).toLowerCase().replace(/à/g, "a").replace(/é/g, "e");
    return m.indexOf("emporter") >= 0;
  }
  function siteOrderIsClosed(o) {
    var ls = s(o.last_status).toLowerCase().replace(/é/g, "e");
    return ls === "fermee" || ls === "closed";
  }
  // Livraison → DELIVERED (col L). Emporter → Fermée (pas de statut livreur DELIVERED).
  function siteOrderCountsInReconciliation(o) {
    if ((o.delivery_status || "").toUpperCase() === "DELIVERED") return true;
    if (siteOrderIsTakeout(o) && siteOrderIsClosed(o)) return true;
    return false;
  }
  function sitePaymentWarning(p, o) {
    if (siteOrderIsTakeout(o)) {
      if (parsePaymentTypes(p.payment_type).indexOf("Bank Transfer") >= 0 ||
          p.payment_type === "Bank Transfer") {
        return " ⚠️ Commande à emporter (col H) : « Bank Transfer » interdit au POS " +
               "(Cash ou Credit card uniquement).";
      }
      if (!isAllDineinPayments(p.payment_type)) {
        return " ⚠️ Commande à emporter : paiement '" + p.payment_type +
               "' inattendu (Cash ou Credit card).";
      }
      return "";
    }
    if (p.payment_type !== SITE_EXPECTED_PAYMENT) {
      return " ⚠️ De plus, son paiement est '" + p.payment_type + "' au lieu de '" +
             SITE_EXPECTED_PAYMENT + "'.";
    }
    return "";
  }
  function pushSitePaymentAnomalies(p, o, anomalies) {
    if (siteOrderIsTakeout(o)) {
      if (parsePaymentTypes(p.payment_type).indexOf("Bank Transfer") >= 0 ||
          p.payment_type === "Bank Transfer") {
        anomalies.push(posAnomaly(p, {
          source: "Site", severity: "haute",
          type: "Mode de paiement incorrect (commande à emporter)",
          detail: "Commande site " + o.identifiant + " à emporter (col H « " +
                  s(o.delivery_method) + " ») : « Bank Transfer » au POS est interdit — " +
                  "seuls Cash ou Credit card sont possibles pour un emporter site.",
          source_ref: o.identifiant,
          payment_pos: p.payment_type,
          payment_source: "Cash ou Credit card" }));
      } else if (!isAllDineinPayments(p.payment_type)) {
        anomalies.push(posAnomaly(p, {
          source: "Site", severity: "haute",
          type: "Mode de paiement incorrect (commande à emporter)",
          detail: "Commande site " + o.identifiant + " à emporter : attendu Cash ou " +
                  "Credit card, trouvé '" + p.payment_type + "' au POS.",
          source_ref: o.identifiant,
          payment_pos: p.payment_type,
          payment_source: "Cash ou Credit card" }));
      }
      return;
    }
    if (p.payment_type !== SITE_EXPECTED_PAYMENT) {
      anomalies.push(posAnomaly(p, {
        source: "Site", severity: "haute",
        type: "Mode de paiement incorrect",
        detail: "Commande site " + o.identifiant + " (livraison) : attendu '" +
                SITE_EXPECTED_PAYMENT + "', trouvé '" + p.payment_type + "' au POS.",
        source_ref: o.identifiant,
        payment_pos: p.payment_type,
        payment_source: SITE_EXPECTED_PAYMENT }));
    }
  }

  // Distance de Levenshtein (nb d'insertions/suppressions/substitutions).
  function editDistance(a, b) {
    a = String(a); b = String(b);
    var m = a.length, n = b.length;
    var prev = [], cur = [], i, j;
    for (j = 0; j <= n; j++) prev[j] = j;
    for (i = 1; i <= m; i++) {
      cur[0] = i;
      for (j = 1; j <= n; j++) {
        var cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
      }
      for (j = 0; j <= n; j++) prev[j] = cur[j];
    }
    return prev[n];
  }

  // Rapproche le site par NUMÉRO (exact) et par faute de frappe (orphelin
  // 5 chiffres). Les tickets SANS numéro sont laissés à la passe commune.
  // Renvoie { anomalies, missing } (commandes livrées non encore rattachées).
  function reconcileSite(pos, site) {
    var anomalies = [], missing = [];
    var byName = {};
    pos.forEach(function (p) { byName[p.ticket_name] = p; });
    var siteIds = new Set(site.map(function (x) { return x.identifiant; }));
    var orphans = pos.filter(function (p) {
      return p.channel === CH_SITE && !siteIds.has(p.ticket_name);
    });
    // Pool des fautes de frappe : tout ticket numérique de 4 à 6 chiffres absent
    // du fichier site (couvre un chiffre en trop/en moins → « 5776 » pour 57767,
    // ou « 58158 » pour 58159). Exclut les tickets Glovo (1-3 chiffres).
    var typoPool = pos.filter(function (p) {
      return /^\d{4,6}$/.test(p.ticket_name) && !siteIds.has(p.ticket_name);
    });
    var typoUsed = new Set();  // ticket_no consommés par une faute de frappe

    var unmatchedDelivered = [];
    site.forEach(function (o) {
      var sid = o.identifiant;
      var counts = siteOrderCountsInReconciliation(o);
      var p = byName[sid];
      if (p) pushSitePaymentAnomalies(p, o, anomalies);
      if (counts) {
        if (!p) { unmatchedDelivered.push(o); return; }
        if (!isNaN(p.total) && Math.abs(p.total - o.order_total) > AMOUNT_TOL) {
          anomalies.push(posAnomaly(p, { source: "Site", severity: "haute",
            type: "Écart de montant",
            detail: "Commande site " + sid + " : " + o.order_total + " DH (site) vs " +
                    p.total + " DH (POS).",
            source_ref: sid,
            amount_pos: p.total, amount_source: o.order_total }));
        }
      }
      // Une commande non livrée (refusée/annulée) PEUT être présente au POS.
    });

    // Faute de frappe sur le numéro : appariement GLOBAL des commandes livrées
    // introuvables avec les tickets du pool (n° proche + même montant + heure).
    var typoPairs = [];
    unmatchedDelivered.forEach(function (o, oi) {
      if (!o.created_at) return;
      typoPool.forEach(function (p, pi) {
        if (isNaN(p.total) || Math.abs(p.total - o.order_total) > AMOUNT_TOL) return;
        if (!p.datetime) return;
        var gap = Math.abs(minutesBetween(p.datetime, o.created_at));
        if (gap > SITE_TYPO_WINDOW_MIN) return;
        var ed = editDistance(p.ticket_name, o.identifiant);
        if (ed > SITE_TYPO_MAX_EDITS) return;
        typoPairs.push({ oi: oi, pi: pi, cost: ed * 1000 + gap, ed: ed, gap: gap });
      });
    });
    typoPairs.sort(function (a, b) { return a.cost - b.cost; });
    var matchedO = new Set(), usedPi = new Set();
    typoPairs.forEach(function (pr) {
      if (matchedO.has(pr.oi) || usedPi.has(pr.pi)) return;
      matchedO.add(pr.oi); usedPi.add(pr.pi);
      var o = unmatchedDelivered[pr.oi], m = typoPool[pr.pi];
      m.channel = CH_SITE;
      typoUsed.add(m.ticket_no);
      var payNote = sitePaymentWarning(m, o);
      anomalies.push(posAnomaly(m, { source: "Site", severity: "moyenne",
        type: "Numéro de commande mal saisi (faute de frappe)",
        detail: "Commande site " + o.identifiant + " livrée : introuvable sous ce numéro, " +
                "mais le ticket POS " + m.ticket_name + " correspond (même montant " +
                o.order_total.toFixed(0) + " DH, +" + pr.gap.toFixed(0) + " min, numéro " +
                "quasi identique). Le caissier a probablement tapé " + m.ticket_name +
                " au lieu de " + o.identifiant + "." + payNote,
        source_ref: o.identifiant,
        amount_pos: m.total, amount_source: o.order_total,
        payment_pos: m.payment_type,
        payment_source: siteOrderIsTakeout(o) ? "Cash ou Credit card" : SITE_EXPECTED_PAYMENT }));
    });

    unmatchedDelivered.forEach(function (o, oi) {
      if (!matchedO.has(oi)) missing.push(o);  // -> passe commune puis « absente »
    });

    // Orphelins « site-like » (5 chiffres) non expliqués.
    orphans.forEach(function (p) {
      if (typoUsed.has(p.ticket_no)) return;
      anomalies.push(posAnomaly(p, { source: "Site", severity: "moyenne",
        type: "Ticket Site au POS sans commande correspondante",
        detail: "Ticket POS " + p.ticket_name + " ressemble à une commande site " +
                "mais n'existe pas dans le fichier site.",
        amount_pos: p.total, payment_pos: p.payment_type }));
    });
    return { anomalies: anomalies, missing: missing };
  }

  // ----------------------------------------------------------------------- //
  // NAPS
  // ----------------------------------------------------------------------- //
  function reconcileNaps(pos, naps) {
    var anomalies = [];
    if (!naps.length) return anomalies;
    pos.forEach(function (p) {
      delete p._naps_split_cc;
      delete p._naps_split_row;
    });

    var posCC = pos.filter(function (p) { return isPurePaymentType(p.payment_type, "Credit card"); });

    var napsDates = naps.map(function (n) { return n.date; }).filter(Boolean).sort();
    var nMin = napsDates[0], nMax = napsDates[napsDates.length - 1];

    var posDates = {};
    pos.forEach(function (p) {
      var d = dateKey(p.datetime);
      if (!d) return;
      if (isPurePaymentType(p.payment_type, "Credit card") || isSplitCashCreditCard(p.payment_type)) {
        posDates[d] = true;
      }
    });

    Object.keys(posDates).sort().forEach(function (d) {
      var posDay = posCC.filter(function (p) { return dateKey(p.datetime) === d; });
      var splitDay = pos.filter(function (p) {
        return dateKey(p.datetime) === d && isSplitCashCreditCard(p.payment_type);
      });

      if (!(d >= nMin && d <= nMax)) {
        var total = posDay.reduce(function (a, p) { return a + (isNaN(p.total) ? 0 : p.total); }, 0);
        if (total > 0) {
          anomalies.push(anomaly({ source: "NAPS", severity: "info",
            type: "Journée non couverte par le relevé NAPS",
            detail: posDay.length + " paiement(s) 'Credit card' du " + d + " (" +
                    total.toFixed(0) + " DH) : le relevé NAPS fourni couvre du " + nMin +
                    " au " + nMax + " (décalage de télécollecte probable).",
            source_ref: d }));
        }
        return;
      }

      var napsDay = naps.filter(function (n) { return n.date === d; });
      var napsUsed = new Set();

      posDay.forEach(function (p) {
        var amt = Math.round((isNaN(p.total) ? 0 : p.total) * 100) / 100;
        var found = -1;
        for (var i = 0; i < napsDay.length; i++) {
          if (napsUsed.has(i)) continue;
          if (Math.round(napsDay[i].montant * 100) / 100 === amt) { found = i; break; }
        }
        if (found >= 0) { napsUsed.add(found); return; }
        anomalies.push(posAnomaly(p, { source: "NAPS", severity: "haute",
          type: "Paiement POS absent du TPE",
          detail: "Paiement 'Credit card' de " + amt.toFixed(0) + " DH au POS le " + d +
                  " à " + hhmm(p.datetime) + " (ticket " + p.ticket_name + ") sans équivalent " +
                  "dans le relevé NAPS.",
          source_ref: p.ticket_no,
          amount_pos: amt }));
      });

      // Cash + CB sur place/emporter : la ligne NAPS non appariée = part carte du ticket.
      splitDay.sort(function (a, b) {
        return (isNaN(b.total) ? 0 : b.total) - (isNaN(a.total) ? 0 : a.total);
      });
      splitDay.forEach(function (p) {
        var ticketTotal = Math.round((isNaN(p.total) ? 0 : p.total) * 100) / 100;
        if (ticketTotal <= 0) return;
        var found = -1, bestAmt = 0;
        for (var j = 0; j < napsDay.length; j++) {
          if (napsUsed.has(j)) continue;
          var nAmt = Math.round(napsDay[j].montant * 100) / 100;
          if (nAmt <= 0 || nAmt >= ticketTotal) continue;
          if (found < 0 || nAmt > bestAmt) { found = j; bestAmt = nAmt; }
        }
        if (found >= 0) {
          napsUsed.add(found);
          p._naps_split_cc = bestAmt;
          p._naps_split_row = napsDay[found].row;
        }
      });

      napsDay.forEach(function (n, i) {
        if (napsUsed.has(i)) return;
        anomalies.push(anomaly({ source: "NAPS", severity: "haute",
          type: "Transaction TPE absente du POS",
          detail: "Transaction NAPS de " + n.montant.toFixed(0) + " DH le " + d +
                  " (ligne " + n.row + " du relevé) sans équivalent 'Credit card' au POS.",
          source_ref: d, amount_source: n.montant, file: "NAPS", row: n.row, when: d }));
      });
    });
    return anomalies;
  }

  /** Date d'une anomalie d'appariement NAPS (POS ou relevé). */
  function napsPairingAnomalyDate(a) {
    if (a.pos_datetime) return dateKey(a.pos_datetime);
    var w = a.when || "";
    if (/^\d{4}-\d{2}-\d{2}/.test(w)) return w.slice(0, 10);
    if (a.source_ref && /^\d{4}-\d{2}-\d{2}/.test(a.source_ref)) return a.source_ref.slice(0, 10);
    return "";
  }

  var NAPS_PAIRING_TYPES = ["Paiement POS absent du TPE", "Transaction TPE absente du POS"];

  /**
   * Journées où total POS CB = total NAPS mais l'appariement transaction par transaction échoue
   * (ex. 110 DH au POS vs 95+15 DH sur le TPE) — pas d'écart financier réel.
   */
  function buildNapsBalancedPairing(pos, naps, anomalies) {
    if (!naps || !naps.length) return [];
    var pairing = anomalies.filter(function (a) {
      return a.source === "NAPS" && NAPS_PAIRING_TYPES.indexOf(a.type) >= 0;
    });
    if (!pairing.length) return [];

    var dates = {};
    pairing.forEach(function (a) {
      var d = napsPairingAnomalyDate(a);
      if (d) dates[d] = true;
    });

    var groups = [];
    Object.keys(dates).sort().forEach(function (d) {
      var dayPairing = pairing.filter(function (a) { return napsPairingAnomalyDate(a) === d; });
      if (!dayPairing.length) return;

      var posCC = pos.filter(function (p) {
        return dateKey(p.datetime) === d && (
          isPurePaymentType(p.payment_type, "Credit card") ||
          (p._naps_split_cc != null && !isNaN(p._naps_split_cc)));
      });
      var posTotal = posCC.reduce(function (s, p) {
        if (isPurePaymentType(p.payment_type, "Credit card")) {
          return s + (isNaN(p.total) ? 0 : p.total);
        }
        return s + p._naps_split_cc;
      }, 0);
      var napsDay = naps.filter(function (n) { return n.date === d; });
      var napsTotal = napsDay.reduce(function (s, n) {
        return s + (isNaN(n.montant) ? 0 : n.montant); }, 0);

      if (Math.abs(Math.round(posTotal) - Math.round(napsTotal)) >= 1) return;

      var posItems = [], napsItems = [], posUnmatched = 0, napsUnmatched = 0;
      dayPairing.forEach(function (a) {
        if (a.type === "Paiement POS absent du TPE") {
          var ap = a.amount_pos || 0;
          posUnmatched += ap;
          posItems.push({
            id: a.id, ticket_no: a.pos_ticket_no, amount: ap,
            ticket_name: a.ticket_name, when: a.when,
          });
        } else {
          var as = a.amount_source || 0;
          napsUnmatched += as;
          napsItems.push({ id: a.id, row: a.row, amount: as, when: a.when });
        }
      });

      groups.push({
        date: d,
        pos_cc_total: Math.round(posTotal),
        naps_total: Math.round(napsTotal),
        pos_unmatched_total: Math.round(posUnmatched),
        naps_unmatched_total: Math.round(napsUnmatched),
        pos_items: posItems,
        naps_items: napsItems,
        anomaly_ids: dayPairing.map(function (a) { return a.id; }),
      });
    });
    return groups;
  }

  function markNapsTotalsOk(anomalies, groups) {
    groups.forEach(function (g) {
      g.anomaly_ids.forEach(function (id) {
        anomalies.forEach(function (a) {
          if (a.id === id) a.naps_totals_ok = true;
        });
      });
    });
  }

  // ----------------------------------------------------------------------- //
  // GLOVO
  // ----------------------------------------------------------------------- //
  // Trouve le meilleur ticket du pool pour une commande Glovo.
  //   payment       : n'accepter que ce mode de paiement (null = indifférent)
  //   requireAmount : n'accepter qu'un montant identique
  //   preferAmount  : à défaut d'exiger, privilégier un montant identique
  // Départage par proximité temporelle.
  function glovoNearest(g, pool, used, lo, hi, payment, requireAmount, preferAmount) {
    var best = -1, bestScore = Infinity;
    for (var i = 0; i < pool.length; i++) {
      if (used.has(i)) continue;
      var p = pool[i];
      if (!p.datetime || p.datetime < lo || p.datetime > hi) continue;
      if (payment != null && p.payment_type !== payment) continue;
      var amountMatch = !isNaN(g.amount) && !isNaN(p.total) &&
                        Math.abs(p.total - g.amount) <= AMOUNT_TOL;
      if (requireAmount && !amountMatch) continue;
      var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
      // Priorités : montant identique > ticket Glovo numéroté > proximité temps.
      var score = gap + (preferAmount && !amountMatch ? 100000 : 0)
                      + (p.channel === CH_UNASSIGNED ? 1000 : 0);
      if (score < bestScore) { bestScore = score; best = i; }
    }
    return best;
  }

  // Rapproche Glovo avec les tickets POS NUMÉROTÉS (canal Glovo). Les commandes
  // encore introuvables sont renvoyées (elles passeront par la passe commune des
  // tickets sans numéro avant d'être déclarées absentes). Renvoie {anomalies, missing}.
  function reconcileGlovo(pos, glovo) {
    var anomalies = [], missing = [];
    var delivered = glovo.filter(function (g) {
      return (g.status || "").toLowerCase() === "delivered";
    }).slice().sort(function (a, b) {
      return (a.received_at ? a.received_at.getTime() : 0) -
             (b.received_at ? b.received_at.getTime() : 0);
    });
    var pool = pos.filter(function (p) { return p.channel === CH_GLOVO; })
                  .slice().sort(function (a, b) {
      return (a.datetime ? a.datetime.getTime() : 0) - (b.datetime ? b.datetime.getTime() : 0);
    });

    var used = new Set();       // indices de tickets (pool) consommés
    var matched = new Set();     // indices de commandes (delivered) appariées
    var MIN = 60000;

    function amountMatch(g, p) {
      return !isNaN(g.amount) && !isNaN(p.total) &&
             Math.abs(p.total - g.amount) <= AMOUNT_TOL;
    }
    function inWindow(g, p, beforeMin, afterMin) {
      if (!p.datetime || !g.received_at) return false;
      return p.datetime >= new Date(g.received_at.getTime() - beforeMin * MIN) &&
             p.datetime <= new Date(g.received_at.getTime() + afterMin * MIN);
    }
    // Appariement GLOBAL glouton : paires éligibles classées par score (temps +
    // montant). Phase 1 EXIGE le même montant (W−AE) — sinon un ticket proche en
    // temps mais mauvais montant « vole » la commande (ex. 725/120 DH vs 658/199 DH).
    function assignGlobalList(list, matchSet, payFilter, beforeMin, afterMin,
                              requireAmount, preferAmount) {
      var pairs = [];
      list.forEach(function (g, gi) {
        if (matchSet.has(gi) || !g.received_at) return;
        var exp = GLOVO_PAYMENT_MAP[g.payment_type];
        pool.forEach(function (p, pi) {
          if (used.has(pi) || !payFilter(p.payment_type, exp)) return;
          if (!inWindow(g, p, beforeMin, afterMin)) return;
          var am = amountMatch(g, p);
          if (requireAmount && !am) return;
          var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
          var score = gap + (preferAmount && !am ? 100000 : 0) - (am ? 0.4 : 0);
          pairs.push({ gi: gi, pi: pi, cost: score });
        });
      });
      pairs.sort(function (a, b) { return a.cost - b.cost; });
      var res = [];
      pairs.forEach(function (pr) {
        if (matchSet.has(pr.gi) || used.has(pr.pi)) return;
        matchSet.add(pr.gi); used.add(pr.pi);
        markGlovoPosMatch(list[pr.gi], pool[pr.pi]);
        res.push(pr);
      });
      return res;
    }
    var isExp = function (pt, exp) { return pt === exp; };
    var isWrong = function (pt, exp) { return pt !== exp; };

    // Phase 1 : fenêtre proche, bon paiement, montant identique (obligatoire).
    assignGlobalList(delivered, matched, isExp, WINDOW_BEFORE_MIN, WINDOW_AFTER_MIN, true, false);

    // Phase 2 : fenêtre proche, MAUVAIS paiement (montant identique privilégié).
    assignGlobalList(delivered, matched, isWrong, WINDOW_BEFORE_MIN, WINDOW_AFTER_MIN, false, true).forEach(function (pr) {
      var g = delivered[pr.gi], p = pool[pr.pi], exp = GLOVO_PAYMENT_MAP[g.payment_type];
      anomalies.push(posAnomaly(p, { source: "Glovo", severity: "haute",
        type: "Mode de paiement incorrect",
        detail: "Commande Glovo " + g.order_id + " (" + g.payment_type + ", " +
                g.amount.toFixed(0) + " DH) : attendu '" + exp + "' au POS, trouvé '" +
                p.payment_type + "' (ticket " + (p.ticket_name || p.ticket_no) +
                " à " + hhmm(p.datetime) + ").",
        source_ref: g.order_id,
        payment_pos: p.payment_type, payment_source: exp }));
    });

    // Phase 3 : saisie tardive (±120 min, bon paiement, montant identique).
    assignGlobalList(delivered, matched, isExp, 120, 120, true, false).forEach(function (pr) {
      var g = delivered[pr.gi], p = pool[pr.pi];
      var delay = minutesBetween(p.datetime, g.received_at);
      anomalies.push(posAnomaly(p, { source: "Glovo", severity: "info",
        type: "Saisie tardive (hors fenêtre 10 min)",
        detail: "Commande Glovo " + g.order_id + " (" + g.amount.toFixed(0) +
                " DH) reçue à " + hhmm(g.received_at) + ", tapée au POS à " +
                hhmm(p.datetime) + " (ticket " + (p.ticket_name || p.ticket_no) + ", " +
                (delay >= 0 ? "+" : "") + delay.toFixed(0) + " min) — présente mais tardive.",
        source_ref: g.order_id,
        amount_pos: p.total, amount_source: g.amount }));
    });

    // Phase 4 : commandes ANNULÉES tapées au POS (avant annulation) — pas des orphelins.
    var cancelled = glovo.filter(function (g) {
      return (g.status || "").toLowerCase() === "cancelled";
    }).slice().sort(function (a, b) {
      return (a.received_at ? a.received_at.getTime() : 0) -
             (b.received_at ? b.received_at.getTime() : 0);
    });
    var cancelledMatched = new Set();
    assignGlobalList(cancelled, cancelledMatched, isExp, WINDOW_BEFORE_MIN, WINDOW_AFTER_MIN, true, false);
    assignGlobalList(cancelled, cancelledMatched, isExp, 120, 120, true, false).forEach(function (pr) {
      var g = cancelled[pr.gi], p = pool[pr.pi];
      g.matched_pos = true;
      var delay = minutesBetween(p.datetime, g.received_at);
      anomalies.push(posAnomaly(p, { source: "Glovo", severity: "info",
        type: "Commande Glovo annulée — présente au POS",
        detail: "Commande Glovo " + g.order_id + " annulée (" + g.payment_type + ", " +
                g.amount.toFixed(0) + " DH) reçue à " + hhmm(g.received_at) +
                ", tapée au POS (ticket " + (p.ticket_name || p.ticket_no) + " à " +
                hhmm(p.datetime) + ", " + (delay >= 0 ? "+" : "") + delay.toFixed(0) +
                " min) — commande annulée sur Glovo mais ticket caisse présent.",
        source_ref: g.order_id,
        amount_pos: p.total, amount_source: g.amount,
        payment_pos: p.payment_type,
        payment_source: GLOVO_PAYMENT_MAP[g.payment_type] }));
    });

    // Phase 5 : bon paiement + fenêtre, montant différent (ex. subtotal W tapé au POS
    // sans déduire la remise AE — W−AE attendu au POS).
    function assignAmountMismatch(list, matchSet, beforeMin, afterMin) {
      var pairs = [];
      list.forEach(function (g, gi) {
        if (matchSet.has(gi) || !g.received_at) return;
        var exp = GLOVO_PAYMENT_MAP[g.payment_type];
        pool.forEach(function (p, pi) {
          if (used.has(pi) || p.payment_type !== exp) return;
          if (!inWindow(g, p, beforeMin, afterMin)) return;
          if (amountMatch(g, p)) return;
          if (isNaN(g.amount) || isNaN(p.total)) return;
          var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
          var tapBrut = !isNaN(g.subtotal) && Math.abs(p.total - g.subtotal) <= AMOUNT_TOL;
          pairs.push({ gi: gi, pi: pi, cost: gap + (tapBrut ? 0 : 0.5) });
        });
      });
      pairs.sort(function (a, b) { return a.cost - b.cost; });
      var res = [];
      pairs.forEach(function (pr) {
        if (matchSet.has(pr.gi) || used.has(pr.pi)) return;
        matchSet.add(pr.gi); used.add(pr.pi);
        markGlovoPosMatch(list[pr.gi], pool[pr.pi]);
        res.push(pr);
      });
      return res;
    }
    assignAmountMismatch(delivered, matched, WINDOW_BEFORE_MIN, WINDOW_AFTER_MIN).forEach(function (pr) {
      var g = delivered[pr.gi], p = pool[pr.pi];
      var note = "";
      if (!isNaN(g.subtotal) && Math.abs(p.total - g.subtotal) <= AMOUNT_TOL &&
          !isNaN(g.discount_funded) && g.discount_funded > 0) {
        note = " Le POS a probablement le subtotal brut (W=" + g.subtotal.toFixed(0) +
               " DH) sans déduire la remise AE (−" + g.discount_funded.toFixed(0) + " DH).";
      }
      anomalies.push(posAnomaly(p, { source: "Glovo", severity: "moyenne",
        type: "Écart de montant",
        detail: "Commande Glovo " + g.order_id + " (" + g.payment_type + ") : " +
                g.amount.toFixed(0) + " DH (W−AE) vs " + p.total.toFixed(0) +
                " DH au POS (ticket " + (p.ticket_name || p.ticket_no) + " à " +
                hhmm(p.datetime) + ")." + note,
        source_ref: g.order_id,
        amount_pos: p.total, amount_source: g.amount,
        payment_pos: p.payment_type,
        payment_source: GLOVO_PAYMENT_MAP[g.payment_type] }));
    });

    // Commandes livrées non appariées -> passe commune puis « absente ».
    delivered.forEach(function (g, gi) {
      if (matched.has(gi)) return;
      if (!g.received_at) {
        anomalies.push(anomaly({ source: "Glovo", severity: "moyenne",
          type: "Commande Glovo sans heure de réception",
          detail: "Commande Glovo " + g.order_id + " sans heure exploitable — " +
                  "rapprochement manuel nécessaire.", source_ref: g.order_id,
          file: "Glovo", row: g.row }));
      } else {
        missing.push(g);
      }
    });

    // Tickets Glovo non appariés.
    pool.forEach(function (p, i) {
      if (!used.has(i)) {
        anomalies.push(posAnomaly(p, { source: "Glovo", severity: "moyenne",
          type: "Ticket Glovo au POS sans commande correspondante",
          detail: "Ticket POS " + p.ticket_name + " (" + hhmm(p.datetime) + ", " +
                  p.payment_type + ", " + (isNaN(p.total) ? "?" : p.total.toFixed(0)) +
                  " DH) classé Glovo mais sans commande Glovo dans la fenêtre.",
          amount_pos: p.total, payment_pos: p.payment_type }));
      }
    });
    return { anomalies: anomalies, missing: missing };
  }

  // ----------------------------------------------------------------------- //
  // PASSE COMMUNE : tickets sans numéro (« Ticket »/vide) attribués JOINTEMENT
  // aux commandes Glovo ET Site encore manquantes — à la plus proche en temps,
  // de même montant et de paiement compatible (un ticket ne peut appartenir
  // qu'à une seule commande, quel que soit le canal).
  // ----------------------------------------------------------------------- //
  function reconcileUnassigned(pos, missingSite, missingGlovo) {
    var anomalies = [];
    var tickets = pos.filter(function (p) { return p.channel === CH_UNASSIGNED; });
    var usedT = new Set();
    var MIN = 60000;

    var demands = [];
    missingSite.forEach(function (o) {
      demands.push({ src: "Site", ref: o.created_at, amount: o.order_total,
                     pay: SITE_EXPECTED_PAYMENT, id: o.identifiant, o: o });
    });
    missingGlovo.forEach(function (g) {
      demands.push({ src: "Glovo", ref: g.received_at, amount: g.amount,
                     pay: GLOVO_PAYMENT_MAP[g.payment_type], id: g.order_id, o: g });
    });
    // Appariement GLOBAL glouton (comme Glovo) : on classe toutes les paires
    // (commande manquante, ticket sans numéro) éligibles — même montant +
    // paiement compatible + fenêtre — par proximité temporelle, et on apparie
    // les plus proches d'abord. Un ticket ne va qu'à une seule commande.
    var pairs = [];
    demands.forEach(function (d, di) {
      if (!d.ref) return;
      tickets.forEach(function (t, ti) {
        if (!t.datetime || t.payment_type !== d.pay) return;
        if (isNaN(t.total) || isNaN(d.amount) || Math.abs(t.total - d.amount) > AMOUNT_TOL) return;
        var gap = Math.abs(minutesBetween(t.datetime, d.ref));
        var ok = d.src === "Site"
          ? gap <= SITE_TYPO_WINDOW_MIN
          : (t.datetime >= new Date(d.ref.getTime() - WINDOW_BEFORE_MIN * MIN) &&
             t.datetime <= new Date(d.ref.getTime() + WINDOW_AFTER_MIN * MIN));
        if (ok) pairs.push({ di: di, ti: ti, gap: gap });
      });
    });
    pairs.sort(function (a, b) { return a.gap - b.gap; });
    var matchedD = new Set(), usedTi = new Set();
    pairs.forEach(function (pr) {
      if (matchedD.has(pr.di) || usedTi.has(pr.ti)) return;
      matchedD.add(pr.di); usedTi.add(pr.ti);
      var d = demands[pr.di], best = tickets[pr.ti];
      usedT.add(best.ticket_no);
      best.channel = d.src === "Site" ? CH_SITE : CH_GLOVO;
      if (d.src === "Glovo") markGlovoPosMatch(d.o, best);
      // Commande retrouvée sous un ticket sans numéro : info (pas une anomalie).
      anomalies.push(posAnomaly(best, { source: d.src, severity: "info",
        type: "Commande rattachée (ticket sans numéro)",
        detail: "Commande " + d.src + " " + d.id + " (" +
                (d.src === "Glovo" ? d.o.payment_type + ", " : "") + d.amount.toFixed(0) +
                " DH) retrouvée au POS sous le ticket sans numéro " + best.ticket_no +
                " (+" + pr.gap.toFixed(0) + " min). Présente — simple oubli de numéro.",
        ticket_name: best.ticket_name || "(vide)",
        source_ref: String(d.id), amount_pos: best.total, amount_source: d.amount,
        payment_pos: best.payment_type }));
    });

    demands.forEach(function (d, di) {
      if (matchedD.has(di)) return;
      if (d.src === "Site") {
        anomalies.push(anomaly({ source: "Site", severity: "haute",
          type: "Commande livrée absente du POS",
          detail: "Commande site " + d.id + " livrée mais introuvable dans le POS.",
          source_ref: String(d.id), amount_source: d.amount,
          file: "Site", row: d.o.row, when: d.ref ? dtFull(d.ref) : "" }));
      } else {
        anomalies.push(anomaly({ source: "Glovo", severity: "haute",
          type: "Commande Glovo absente du POS",
          detail: "Commande Glovo " + d.id + " reçue à " +
                  (d.ref ? dateKey(d.ref) + " " + hhmm(d.ref) : "?") + " (" + d.o.payment_type +
                  ", " + d.amount.toFixed(0) + " DH) non retrouvée au POS " +
                  "(aucun ticket au bon mode de paiement).",
          source_ref: String(d.id), amount_source: d.amount,
          payment_source: d.pay || "?",
          file: "Glovo", row: d.o.row, when: d.ref ? dtFull(d.ref) : "" }));
      }
    });

    // Tickets sans numéro restants -> à rattacher (souvent ventes comptoir).
    tickets.forEach(function (p) {
      if (usedT.has(p.ticket_no)) return;
      anomalies.push(posAnomaly(p, { source: "POS", severity: "info",
        type: "Ticket sans numéro (à rattacher)",
        detail: "Ticket " + p.ticket_no + " du " + dateKey(p.datetime) + " à " +
                hhmm(p.datetime) + " (" + p.payment_type + ", " +
                (isNaN(p.total) ? "?" : p.total) + " DH) sans numéro — non rattaché à " +
                "une commande Glovo ni Site.",
        ticket_name: p.ticket_name || "(vide)",
        amount_pos: p.total, payment_pos: p.payment_type }));
    });
    return anomalies;
  }

  function glovoAggregate(pos, glovo) {
    // Écart global par mode de paiement (info), sur canaux FINAUX (après passes).
    var anomalies = [];
    var delivered = glovo.filter(function (g) { return (g.status || "").toLowerCase() === "delivered"; });
    var posGlovo = pos.filter(function (p) { return p.channel === CH_GLOVO; });
    var gOnline = delivered.filter(function (g) { return g.payment_type === "Online"; }).length;
    var gCash = delivered.filter(function (g) { return g.payment_type === "Cash"; }).length;
    var pBT = posGlovo.filter(function (p) { return p.payment_type === "Bank Transfer"; }).length;
    var pCash = posGlovo.filter(function (p) { return p.payment_type === "Cash"; }).length;
    if (gOnline !== pBT) {
      anomalies.push(anomaly({ source: "Glovo", severity: "info",
        type: "Écart global paiement en ligne",
        detail: "Glovo 'Online' : " + gOnline + " vs POS 'Bank Transfer' (tickets Glovo) : " +
                pBT + " → écart de " + (gOnline - pBT) + "." }));
    }
    if (gCash !== pCash) {
      anomalies.push(anomaly({ source: "Glovo", severity: "info",
        type: "Écart global paiement cash",
        detail: "Glovo 'Cash' : " + gCash + " vs POS 'Cash' (tickets Glovo) : " +
                pCash + " → écart de " + (gCash - pCash) + "." }));
    }
    return anomalies;
  }

  // ----------------------------------------------------------------------- //
  // Sur place / emporter + tickets à rattacher
  // ----------------------------------------------------------------------- //
  function reconcileDinein(pos) {
    var anomalies = [];
    pos.filter(function (p) { return p.channel === CH_DINEIN; }).forEach(function (p) {
      if (!isAllDineinPayments(p.payment_type)) {
        anomalies.push(posAnomaly(p, { source: "Sur place", severity: "moyenne",
          type: "Mode de paiement inattendu (sur place/emporter)",
          detail: "Ticket " + p.ticket_name + " sur place/emporter payé '" +
                  p.payment_type + "' (attendu Cash ou Credit card).",
          amount_pos: p.total, payment_pos: p.payment_type }));
      }
    });
    return anomalies;
  }

  // ----------------------------------------------------------------------- //
  // Produits (POS col J / Glovo col AY) — contrôle secondaire, pas clé d'appariement
  // ----------------------------------------------------------------------- //
  var PRODUCT_SKIP = { "dh": 1, "menu": 1, "x": 1, "ajustement": 1, "compose": 1, "ton": 1 };

  function productTokens(text) {
    if (!text) return [];
    var t = String(text).toLowerCase()
      .replace(/\d+/g, " ")
      .replace(/[×x]/gi, " ")
      .replace(/\[/g, " ")
      .replace(/\]/g, " ")
      .replace(/[^a-zàâäéèêëïîôùûüç0-9\s]/gi, " ");
    var out = [];
    t.split(/[\s,]+/).forEach(function (w) {
      w = w.trim();
      if (w.length > 2 && !PRODUCT_SKIP[w]) out.push(w);
    });
    return out;
  }

  function productSimilarity(posText, glovoText) {
    var ta = productTokens(posText), tb = productTokens(glovoText);
    if (!ta.length || !tb.length) return 0;
    var setA = {}, setB = {}, inter = 0;
    ta.forEach(function (w) { setA[w] = 1; });
    tb.forEach(function (w) {
      if (setA[w]) inter++;
      setB[w] = 1;
    });
    var union = 0;
    Object.keys(setA).forEach(function (w) { union++; });
    Object.keys(setB).forEach(function (w) { if (!setA[w]) union++; });
    return union ? inter / union : 0;
  }

  function findGlovoForDuplicate(cluster, glovo) {
    if (!glovo || cluster[0].channel !== CH_GLOVO) return null;
    var mid = cluster[0].datetime;
    if (!mid) return null;
    var amounts = cluster.map(function (p) { return p.total; });
    var best = null, bestGap = 1e9;
    glovo.forEach(function (g) {
      if ((g.status || "").toLowerCase() !== "delivered" || !g.received_at) return;
      var gap = Math.abs(minutesBetween(mid, g.received_at));
      if (gap > 120) return;
      var ok = amounts.some(function (a) {
        return !isNaN(a) && !isNaN(g.amount) && Math.abs(a - g.amount) <= AMOUNT_TOL;
      });
      if (!ok) return;
      if (gap < bestGap) { bestGap = gap; best = g; }
    });
    if (!best) return null;
    return {
      order_id: best.order_id,
      payment_type: best.payment_type,
      expected_pos: GLOVO_PAYMENT_MAP[best.payment_type],
      order_items: best.order_items || "",
    };
  }

  function enrichProductConfirmation(anomalies, pos, glovo) {
    if (!glovo || !pos.length) return;
    var byNo = {};
    pos.forEach(function (p) { if (p.ticket_no) byNo[p.ticket_no] = p; });
    var byOrder = {};
    glovo.forEach(function (g) { if (g.order_id) byOrder[g.order_id] = g; });

    anomalies.forEach(function (a) {
      var g = a.source_ref ? byOrder[a.source_ref] : null;
      var p = a.pos_ticket_no ? byNo[a.pos_ticket_no] : null;
      if (g && p && p.designations && g.order_items) {
        var score = productSimilarity(p.designations, g.order_items);
        if (score >= 0.2) {
          var pct = Math.round(score * 100);
          a.products_confirm = pct;
          var tag = score >= 0.4 ? "✓ Produits compatibles" : "~ Produits partiellement compatibles";
          a.detail += " [" + tag + " " + pct + "% — contrôle contenu, pas clé d'appariement]";
        }
        return;
      }
      // Commande absente : suggestion par contenu + heure (info seulement)
      if (g && !p && (a.type === "Commande Glovo absente du POS") && g.order_items && g.received_at) {
        var bestP = null, bestScore = 0;
        pos.forEach(function (px) {
          if (px.channel !== CH_GLOVO || !px.datetime || !px.designations) return;
          var gap = Math.abs(minutesBetween(px.datetime, g.received_at));
          if (gap > 120) return;
          var sc = productSimilarity(px.designations, g.order_items);
          if (sc > bestScore) { bestScore = sc; bestP = px; }
        });
        if (bestP && bestScore >= 0.35) {
          a.detail += " [Suggestion : ticket " + bestP.ticket_name + " ligne " + bestP.row +
                      " — produits ~" + Math.round(bestScore * 100) + " % compatibles, à vérifier]";
        }
      }
    });
  }

  function markGlovoPosMatch(g, p) {
    if (!g || !p) return;
    g.matched_pos_ticket_no = p.ticket_no;
    p.matched_glovo_order = g.order_id;
    if ((g.status || "").toLowerCase() === "cancelled") g.matched_pos = true;
  }

  function glovoMatchForPosLine(p, glovo, wideWindow) {
    if (!glovo || p.channel !== CH_GLOVO || !p.datetime) return null;
    var afterMin = wideWindow ? 120 : WINDOW_AFTER_MIN;
    var best = null, bestScore = 1e9;
    glovo.forEach(function (g) {
      if ((g.status || "").toLowerCase() !== "delivered" || !g.received_at) return;
      var exp = GLOVO_PAYMENT_MAP[g.payment_type];
      if (p.payment_type !== exp) return;
      if (isNaN(p.total) || isNaN(g.amount) || Math.abs(p.total - g.amount) > AMOUNT_TOL) return;
      if (p.datetime < new Date(g.received_at.getTime() - WINDOW_BEFORE_MIN * 60000) ||
          p.datetime > new Date(g.received_at.getTime() + afterMin * 60000)) return;
      var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
      var score = gap;
      if (p.designations && g.order_items) score -= productSimilarity(p.designations, g.order_items) * 0.5;
      if (score < bestScore) { bestScore = score; best = g; }
    });
    return best;
  }

  /** Vrai doublon = re-tapé pour corriger LA MÊME commande (pas deux ventes différentes au même n°). */
  function isLikelyCorrectionCluster(cl, glovo) {
    if (cl.length < 2) return false;
    var glovoByOrder = {};
    cl.forEach(function (p) {
      var m = glovoMatchForPosLine(p, glovo, true);
      if (m) glovoByOrder[m.order_id] = m;
    });
    var orderIds = Object.keys(glovoByOrder);
    if (orderIds.length > 1) return false;

    // Une seule commande Glovo retrouvée sur le cluster → correction si les autres
    // lignes sont des saisies erronées (mauvais paiement / montant), pas une 2ᵉ vente.
    if (orderIds.length === 1) {
      var g = glovoByOrder[orderIds[0]];
      var expectedPay = GLOVO_PAYMENT_MAP[g.payment_type];
      for (var j = 0; j < cl.length; j++) {
        var pj = cl[j];
        if (glovoMatchForPosLine(pj, glovo, true)) continue;
        // Cash 0 DH = vente enveloppe distincte (ex. ticket 15), pas une correction.
        if (pj.payment_type === "Cash" &&
            (isNaN(pj.total) || pj.total <= AMOUNT_TOL)) return false;
        // Mauvais mode de paiement vs la commande Glovo → saisie erronée (ex. 98 BT).
        if (pj.payment_type !== expectedPay) continue;
        if (!isNaN(pj.total) && pj.total > AMOUNT_TOL &&
            Math.abs(pj.total - g.amount) > AMOUNT_TOL) {
          if (pj.designations && g.order_items &&
              productSimilarity(pj.designations, g.order_items) < 0.25) return false;
        }
      }
      return true;
    }

    // Aucune commande Glovo : correction seulement si montants identiques ou produits quasi identiques.
    var nonZero = cl.filter(function (p) { return !isNaN(p.total) && p.total > AMOUNT_TOL; });
    if (nonZero.length >= 2) {
      var t0 = nonZero[0].total;
      for (var i = 1; i < nonZero.length; i++) {
        if (Math.abs(nonZero[i].total - t0) > AMOUNT_TOL) return false;
      }
    }

    var base = cl[0];
    for (var k = 1; k < cl.length; k++) {
      if (base.designations && cl[k].designations &&
          productSimilarity(base.designations, cl[k].designations) < 0.35) return false;
      if (!isNaN(base.total) && !isNaN(cl[k].total) &&
          base.total > AMOUNT_TOL && cl[k].total > AMOUNT_TOL &&
          Math.abs(base.total - cl[k].total) > AMOUNT_TOL) return false;
    }
    return true;
  }

  function glovoLineMatchTag(p, glovo) {
    var m = glovoMatchForPosLine(p, glovo, true);
    if (m) {
      return " ✓ commande Glovo " + m.order_id + " (" + m.amount.toFixed(0) +
             " DH " + m.payment_type + ")";
    }
    return " ✗ aucune commande Glovo correspondante";
  }

  function summarizeExtraByPayment(lines) {
    var buckets = {};
    lines.forEach(function (p) {
      var t = isNaN(p.total) ? 0 : p.total;
      if (t <= AMOUNT_TOL) return;
      buckets[p.payment_type] = (buckets[p.payment_type] || 0) + t;
    });
    return Object.keys(buckets).map(function (pay) {
      return buckets[pay].toFixed(0) + " DH en « " + pay + " »";
    }).join(" + ");
  }

  function extraPaymentNote(dupVers) {
    var parts = summarizeExtraByPayment(dupVers);
    if (!parts) return "";
    var hasCash = false, hasBT = false, hasCC = false;
    dupVers.forEach(function (p) {
      if (isNaN(p.total) || p.total <= AMOUNT_TOL) return;
      if (p.payment_type === "Cash") hasCash = true;
      else if (p.payment_type === "Bank Transfer") hasBT = true;
      else if (p.payment_type === "Credit card") hasCC = true;
    });
    var only = [];
    if (hasCash && !hasBT && !hasCC) only.push("surplus uniquement en Cash");
    if (hasBT && !hasCash && !hasCC) only.push("surplus uniquement en Bank Transfer (pas en Cash)");
    if (hasCC && !hasCash && !hasBT) only.push("surplus uniquement en Credit card");
    return parts + (only.length ? " — " + only[0] : "");
  }

  function flagMisusedGlovoNumber(p, matchedLine, matchedG) {
    var amt = isNaN(p.total) ? 0 : p.total;
    var amtTxt = amt > AMOUNT_TOL ? amt.toFixed(0) + " DH" :
      (amt === 0 ? "0 DH (annulée ou montant nul — vérifier)" : "?");
    var sibling = matchedG && matchedLine
      ? " La commande Glovo " + matchedG.order_id + " (" + matchedG.amount.toFixed(0) +
        " DH " + matchedG.payment_type + ") est sur la ligne " + matchedLine.row + "."
      : "";
    return posAnomaly(p, {
      source: "Glovo", severity: "moyenne",
      type: "Numéro Glovo réutilisé (vente distincte)",
      detail: "Ticket POS " + p.ticket_name + " (" + hhmm(p.datetime) + ", " + amtTxt +
              " Cash, ligne " + p.row + ") — autre encaissement au même n° Glovo, pas une correction." +
              sibling +
              " ⚠️ Cash réel : à contrôler dans l'enveloppe caissier.",
      amount_pos: amt > AMOUNT_TOL ? amt : null,
      payment_pos: "Cash",
    });
  }

  // DOUBLONS : même n° tapé plusieurs fois = correction UNIQUEMENT si c'est la
  // même commande (montant / produits / Glovo). Sinon = ventes distinctes au même n°.
  // ----------------------------------------------------------------------- //
  var DUP_WINDOW_MIN = 30;

  function detectDuplicates(pos, anomalies, glovo) {
    var groups = {};
    pos.forEach(function (p) {
      if ((p.channel === CH_GLOVO || p.channel === CH_SITE) && /^\d+$/.test(p.ticket_name)) {
        (groups[p.ticket_name] = groups[p.ticket_name] || []).push(p);
      }
    });
    var handledNo = {};  // ticket_no consommés par un vrai cluster doublon
    var misusedNo = {};  // ticket_no signalés « n° réutilisé » (pas orphelin Glovo)
    Object.keys(groups).forEach(function (name) {
      var list = groups[name];
      if (list.length < 2) return;
      list.sort(function (a, b) {
        return (a.datetime ? a.datetime.getTime() : 0) - (b.datetime ? b.datetime.getTime() : 0);
      });
      // Regrouper en clusters de saisies rapprochées (≤ 30 min).
      var clusters = [], cur = [list[0]];
      for (var i = 1; i < list.length; i++) {
        var prev = cur[cur.length - 1];
        if (list[i].datetime && prev.datetime &&
            Math.abs(minutesBetween(list[i].datetime, prev.datetime)) <= DUP_WINDOW_MIN) {
          cur.push(list[i]);
        } else { clusters.push(cur); cur = [list[i]]; }
      }
      clusters.push(cur);

      clusters.forEach(function (cl) {
        if (cl.length < 2) return;

        if (!isLikelyCorrectionCluster(cl, glovo)) {
          var matchedLine = null, matchedG = null;
          cl.forEach(function (p) {
            var m = glovoMatchForPosLine(p, glovo, true);
            if (m) { matchedLine = p; matchedG = m; }
          });
          cl.forEach(function (p) {
            if (glovoMatchForPosLine(p, glovo, true)) return;
            if (p.channel !== CH_GLOVO) return;
            if (p.payment_type === "Cash") {
              anomalies.push(flagMisusedGlovoNumber(p, matchedLine, matchedG));
              if (p.ticket_no) misusedNo[p.ticket_no] = true;
            }
          });
          return;
        }

        cl.forEach(function (p) { if (p.ticket_no) handledNo[p.ticket_no] = true; });
        var first = cl[0], last = cl[cl.length - 1];
        var glovoRef = findGlovoForDuplicate(cl, glovo);
        var retainPay = glovoRef ? glovoRef.expected_pos : last.payment_type;
        var retainIdx = -1;
        for (var ri = 0; ri < cl.length; ri++) {
          if (cl[ri].payment_type === retainPay) { retainIdx = ri; break; }
        }
        if (retainIdx < 0) retainIdx = cl.length - 1;
        var retainVer = cl[retainIdx];
        var dupVers = cl.filter(function (_, i) { return i !== retainIdx; });
        var versions = cl.map(function (p, idx) {
          return "v" + (idx + 1) + " " + hhmm(p.datetime) + " " +
                 (isNaN(p.total) ? "?" : p.total.toFixed(0)) + " DH " + p.payment_type +
                 " (ligne " + p.row + ")" + glovoLineMatchTag(p, glovo);
        }).join(" → ");
        var diffs = [];
        if (!isNaN(first.total) && !isNaN(last.total) && Math.abs(first.total - last.total) > AMOUNT_TOL)
          diffs.push("montant " + first.total.toFixed(0) + "→" + last.total.toFixed(0));
        if (first.payment_type !== last.payment_type)
          diffs.push("paiement " + first.payment_type + "→" + last.payment_type);
        var extra = cl.reduce(function (a, p) { return a + (isNaN(p.total) ? 0 : p.total); }, 0) -
                    (isNaN(retainVer.total) ? 0 : retainVer.total);
        var dupPay = dupVers.length === 1 ? dupVers[0].payment_type : retainPay;
        var extraBreakdown = extraPaymentNote(dupVers);
        var retainNote = "";
        if (glovoRef) {
          var retainMatch = glovoMatchForPosLine(retainVer, glovo, true);
          retainNote = " Commande Glovo " + glovoRef.order_id + " (" +
            (retainMatch
              ? retainMatch.amount.toFixed(0) + " DH " + retainMatch.payment_type
              : glovoRef.payment_type) +
            ") → bonne saisie : v" + (retainIdx + 1) + " ligne " + retainVer.row +
            " (" + retainVer.total.toFixed(0) + " DH " + retainVer.payment_type + ")." +
            " À annuler au POS : " +
            dupVers.map(function (p) {
              return "v" + (cl.indexOf(p) + 1) + " ligne " + p.row + " (" +
                     (isNaN(p.total) ? "?" : p.total.toFixed(0)) + " DH " + p.payment_type +
                     ")" + glovoLineMatchTag(p, glovo);
            }).join(" · ") + ".";
        } else if (first.payment_type !== last.payment_type) {
          retainNote = " Sans commande Glovo retrouvée : retenir en principe la dernière saisie v" +
            cl.length + " (" + last.payment_type + ") si c'est la correction — vérifier manuellement.";
        }
        anomalies.push(posAnomaly(retainVer, {
          source: cl[0].channel === CH_SITE ? "Site" : "Glovo", severity: "moyenne",
          type: "Ticket en double (correction)",
          detail: "Ticket " + name + " saisi " + cl.length + " fois (doublon / correction) : " +
                  versions + ". " + (diffs.length ? "Évolution : " + diffs.join(", ") + ". " : "") +
                  (extra > AMOUNT_TOL
                    ? "⚠️ Surplus POS à retirer : " + extra.toFixed(0) + " DH" +
                      (extraBreakdown ? " → " + extraBreakdown : "") +
                      ". Une version doit être annulée au POS."
                    : "⚠️ Vérifier qu'une version est bien annulée au POS.") +
                  retainNote,
          source_ref: glovoRef ? glovoRef.order_id : name,
          amount_pos: extra,
          payment_pos: dupPay,
          payment_source: retainPay }));
      });
    });

    // Retirer les anomalies « orphelin » remplacées par une correction.
    return anomalies.filter(function (a) {
      if (a.type !== "Ticket Glovo au POS sans commande correspondante" &&
          a.type !== "Ticket Site au POS sans commande correspondante") return true;
      if (a.pos_ticket_no && handledNo[a.pos_ticket_no]) return false;
      if (a.pos_ticket_no && misusedNo[a.pos_ticket_no] &&
          a.type === "Ticket Glovo au POS sans commande correspondante") return false;
      return true;
    });
  }

  // Anomalies complémentaires pour expliquer les écarts financiers quand le
  // rapprochement ticket↔commande ne produit pas déjà une anomalie contributive.
  function appendFinancialGapAnomalies(pos, glovo, site, anomalies) {
    if (site) {
      var siteById = {};
      site.forEach(function (o) { siteById[o.identifiant] = o; });
      pos.forEach(function (p) {
        if (p.channel !== CH_SITE) return;
        var o = siteById[p.ticket_name];
        if (!o) return;
        if (siteOrderIsTakeout(o)) return;
        if (siteOrderCountsInReconciliation(o)) return;
        if (anomalies.some(function (a) {
          return a.pos_ticket_no === p.ticket_no && a.source === "Site";
        })) return;
        var st = o.delivery_status || o.last_status || "vide";
        anomalies.push(posAnomaly(p, {
          source: "Site", severity: "moyenne",
          type: "Commande site au POS — statut non livré",
          detail: "Ticket POS " + p.ticket_name + " (" +
                  (isNaN(p.total) ? "?" : p.total.toFixed(0)) + " DH, " + p.payment_type +
                  ") : commande site présente mais pas « DELIVERED » (statut : " + st +
                  ") — comptée au POS, absente des livrées site (col L).",
          source_ref: o.identifiant,
          amount_pos: p.total, amount_source: 0,
          payment_pos: p.payment_type }));
      });
    }

    if (!glovo) return;

    function glovoCountedInFinancial(g) {
      var st = (g.status || "").toLowerCase();
      if (st === "delivered") return true;
      return st === "cancelled" && g.matched_pos;
    }
    function hasGlovoFinAnomaly(orderId, posNo, bucket) {
      return anomalies.some(function (a) {
        if (a.source !== "Glovo") return false;
        if (orderId && a.source_ref === String(orderId)) {
          if (a.type === "Commande Glovo absente du POS" ||
              a.type === "Commande Glovo sans heure de réception" ||
              a.type === "Commande Glovo Online absente du POS (écart financier)" ||
              a.type === "Commande Glovo Cash absente du POS (écart financier)") return true;
        }
        if (posNo && a.pos_ticket_no === posNo) {
          if (a.type === "Ticket Glovo au POS sans commande correspondante" ||
              a.type === "Ticket Glovo Online au POS sans appariement (écart financier)" ||
              a.type === "Ticket Glovo Cash au POS sans appariement (écart financier)") return true;
        }
        return false;
      });
    }

    glovo.forEach(function (g) {
      if (!glovoCountedInFinancial(g)) return;
      if (g.matched_pos_ticket_no) return;
      if (hasGlovoFinAnomaly(g.order_id, null, g.payment_type)) return;
      var payLabel = g.payment_type === "Cash" ? "Cash" : "Online";
      var expPay = GLOVO_PAYMENT_MAP[g.payment_type] || "Bank Transfer";
      anomalies.push(anomaly({
        source: "Glovo", severity: "haute",
        type: "Commande Glovo " + payLabel + " absente du POS (écart financier)",
        detail: "Commande Glovo " + g.order_id + " (" + g.amount.toFixed(0) +
                " DH " + payLabel + ") comptée côté source mais aucun ticket POS « " +
                expPay + " » apparié.",
        source_ref: String(g.order_id),
        amount_source: g.amount,
        payment_source: expPay,
        file: "Glovo", row: g.row,
        when: g.received_at ? dtFull(g.received_at) : "" }));
    });

    pos.forEach(function (p) {
      if (p.channel !== CH_GLOVO) return;
      if (p.matched_glovo_order) return;
      if (isNaN(p.total) || p.total <= AMOUNT_TOL) return;
      if (anomalies.some(function (a) {
        return a.pos_ticket_no === p.ticket_no && a.type === "Ticket en double (correction)";
      })) return;
      if (anomalies.some(function (a) {
        return a.type === "Ticket en double (correction)" && a.ticket_name === p.ticket_name;
      })) return;
      if (hasGlovoFinAnomaly(null, p.ticket_no, p.payment_type)) return;
      var bucket = p.payment_type === "Cash" ? "Cash" : "Online";
      anomalies.push(posAnomaly(p, {
        source: "Glovo", severity: "moyenne",
        type: "Ticket Glovo " + bucket + " au POS sans appariement (écart financier)",
        detail: "Ticket POS " + p.ticket_name + " (" + hhmm(p.datetime) + ", " +
                p.total.toFixed(0) + " DH " + p.payment_type + ") compté au POS " +
                bucket + " mais sans commande Glovo " + bucket + " appariée.",
        amount_pos: p.total, payment_pos: p.payment_type }));
    });
  }

  // ----------------------------------------------------------------------- //
  // Orchestration
  // ----------------------------------------------------------------------- //
  // Le fichier POS définit la PÉRIODE d'analyse : les commandes Glovo/Site
  // hors des dates présentes dans le POS sont ignorées (une commande d'un jour
  // non couvert par le POS ne doit pas être signalée comme « absente »).
  function posDateSet(pos) {
    var set = new Set();
    pos.forEach(function (p) { var k = dateKey(p.datetime); if (k) set.add(k); });
    return set;
  }

  function filterToPosDates(rows, getDate, posDates) {
    var kept = [], excluded = 0;
    rows.forEach(function (r) {
      var k = dateKey(getDate(r));
      if (k && !posDates.has(k)) { excluded++; return; }  // hors période -> ignoré
      kept.push(r);
    });
    return { kept: kept, excluded: excluded };
  }

  function run(pos, glovo, naps, site) {
    var siteIds = site ? new Set(site.map(function (x) { return x.identifiant; })) : null;
    addChannel(pos, siteIds);

    // Restreindre Glovo et Site aux dates présentes dans le POS.
    var posDates = posDateSet(pos);
    var glovoExcluded = 0, siteExcluded = 0;
    if (glovo) {
      var gf = filterToPosDates(glovo, function (g) { return g.received_at; }, posDates);
      glovo = gf.kept; glovoExcluded = gf.excluded;
    }
    if (site) {
      var sf = filterToPosDates(site, function (o) { return o.created_at; }, posDates);
      site = sf.kept; siteExcluded = sf.excluded;
    }

    // 1) Rapprochements par NUMÉRO (canaux disjoints, non ambigus).
    // 2) Passe COMMUNE : tickets sans numéro attribués jointement (Glovo+Site).
    var anomalies = [];
    var missingSite = [], missingGlovo = [];
    if (site) {
      var rs = reconcileSite(pos, site);
      anomalies = anomalies.concat(rs.anomalies); missingSite = rs.missing;
    }
    if (naps) anomalies = anomalies.concat(reconcileNaps(pos, naps));
    if (glovo) {
      var rg = reconcileGlovo(pos, glovo);
      anomalies = anomalies.concat(rg.anomalies); missingGlovo = rg.missing;
    }
    anomalies = anomalies.concat(reconcileDinein(pos));
    anomalies = anomalies.concat(reconcileUnassigned(pos, missingSite, missingGlovo));
    if (glovo) anomalies = anomalies.concat(glovoAggregate(pos, glovo));
    anomalies = detectDuplicates(pos, anomalies, glovo);  // doublons/corrections
    appendFinancialGapAnomalies(pos, glovo, site, anomalies);
    enrichProductConfirmation(anomalies, pos, glovo);

    annotate(pos, anomalies);
    var naps_balanced = naps ? buildNapsBalancedPairing(pos, naps, anomalies) : [];
    markNapsTotalsOk(anomalies, naps_balanced);

    var summary = buildSummary(pos, anomalies, glovo, naps, site);
    var dates = Array.from(posDates).sort();
    summary.pos_date_min = dates.length ? dates[0] : "";
    summary.pos_date_max = dates.length ? dates[dates.length - 1] : "";
    summary.glovo_excluded = glovoExcluded;
    summary.site_excluded = siteExcluded;
    summary.financial = computeFinancial(pos, glovo, naps, site, posDates);
    return { anomalies: anomalies, pos: pos, summary: summary, naps_balanced: naps_balanced };
  }

  function listPosDates(pos) {
    return Array.from(posDateSet(pos)).sort();
  }

  function runDailyBreakdown(pos, glovo, naps, site) {
    var byDay = {};
    listPosDates(pos).forEach(function (dk) {
      var dayPos = pos.filter(function (p) { return dateKey(p.datetime) === dk; });
      if (!dayPos.length) return;
      byDay[dk] = run(dayPos, glovo, naps, site);
    });
    return byDay;
  }

  // Réconciliation FINANCIÈRE : totaux par mode de paiement × canal, et écarts
  // POS vs source (TPE/NAPS, Glovo, Site).
  function computeFinancial(pos, glovo, naps, site, posDates) {
    var PAYS = ["Cash", "Bank Transfer", "Credit card"];
    var byPayment = { "Cash": 0, "Bank Transfer": 0, "Credit card": 0, "Autre": 0 };
    var matrix = {};  // canal -> { paiement -> somme }
    pos.forEach(function (p) {
      var alloc = allocatePosPaymentAmounts(p);
      Object.keys(alloc).forEach(function (pay) {
        var v = alloc[pay];
        if (!v) return;
        byPayment[pay] = (byPayment[pay] || 0) + v;
        matrix[p.channel] = matrix[p.channel] || { "Cash": 0, "Bank Transfer": 0, "Credit card": 0, "Autre": 0 };
        matrix[p.channel][pay] += v;
      });
    });

    function sumPos(channel) {
      return pos.reduce(function (a, p) {
        return a + (p.channel === channel && !isNaN(p.total) ? p.total : 0); }, 0);
    }
    function sumPosGlovoPay(pay) {
      return pos.reduce(function (a, p) {
        if (p.channel !== CH_GLOVO) return a;
        var alloc = allocatePosPaymentAmounts(p);
        return a + (alloc[pay] || 0);
      }, 0);
    }
    function sumPosGlovoTicketTotals() {
      return pos.reduce(function (a, p) {
        return a + (p.channel === CH_GLOVO && !isNaN(p.total) ? p.total : 0);
      }, 0);
    }
    function sumGlovoAmount(pay) {
      if (!glovo) return 0;
      return glovo.reduce(function (a, g) {
        if (g.payment_type !== pay) return a;
        var st = (g.status || "").toLowerCase();
        var amt = isNaN(g.amount) ? 0 : g.amount;
        if (st === "delivered") return a + amt;
        // Annulée mais tapée au POS : compter côté source pour la réconciliation financière.
        if (st === "cancelled" && g.matched_pos) return a + amt;
        return a;
      }, 0);
    }

    var posCC = byPayment["Credit card"];
    var napsTotal = 0;
    if (naps) naps.forEach(function (n) {
      if (posDates.has(n.date) && !isNaN(n.montant)) napsTotal += n.montant; });

    var posGlovoTickets = sumPosGlovoTicketTotals(), glovoW = 0;
    if (glovo) glovo.forEach(function (g) {
      var st = (g.status || "").toLowerCase();
      if (isNaN(g.amount)) return;
      if (st === "delivered" || (st === "cancelled" && g.matched_pos)) glovoW += g.amount;
    });

    var posSite = sumPos(CH_SITE), siteL = 0;
    if (site) site.filter(siteOrderCountsInReconciliation)
      .forEach(function (o) { if (!isNaN(o.order_total)) siteL += o.order_total; });

    var lines = [];
    if (naps) lines.push({
      lineKey: "naps", source: "💳 TPE (NAPS)", pos_label: "POS « Credit card »",
      pos: posCC, src_label: "Relevé NAPS", src: napsTotal, ecart: napsTotal - posCC,
    });
    if (glovo) {
      var posGlovoBT = sumPosGlovoPay("Bank Transfer");
      var posGlovoCash = sumPosGlovoPay("Cash");
      var posGlovoCC = sumPosGlovoPay("Credit card");
      var posGlovoAutre = sumPosGlovoPay("Autre");
      // Total POS réconcilié = Online (BT) + Cash — doit égaler les deux lignes ci-dessus.
      var posGlovoFin = posGlovoBT + posGlovoCash;
      var glovoOnline = sumGlovoAmount("Online");
      var glovoCash = sumGlovoAmount("Cash");
      var totalNote = "Côté POS : Total = Online (Bank Transfer) + Cash.";
      if (posGlovoCC >= 1) {
        totalNote += " CB Glovo au POS : " + Math.round(posGlovoCC) + " DH (hors Online/Cash).";
      }
      if (posGlovoAutre >= 1) {
        totalNote += " Autre paiement Glovo : " + Math.round(posGlovoAutre) + " DH.";
      }
      lines.push({
        lineKey: "glovo_online", source: "🛵 Glovo — Online",
        pos_label: "POS Glovo « Bank Transfer »",
        pos: posGlovoBT, src_label: "Glovo Online (W − AE)", src: glovoOnline,
        ecart: glovoOnline - posGlovoBT, group: "glovo",
      });
      lines.push({
        lineKey: "glovo_cash", source: "🛵 Glovo — Cash",
        pos_label: "POS Glovo « Cash »",
        pos: posGlovoCash, src_label: "Glovo Cash (W − AE)", src: glovoCash,
        ecart: glovoCash - posGlovoCash, group: "glovo",
      });
      lines.push({
        lineKey: "glovo_total", source: "🛵 Glovo — Total",
        pos_label: "POS Glovo Online + Cash",
        pos: posGlovoFin, src_label: "Glovo livrées (W − AE)", src: glovoW,
        ecart: glovoW - posGlovoFin, group: "glovo", isTotal: true,
        note: totalNote,
      });
    }
    if (site) lines.push({
      lineKey: "site", source: "🌐 Site", pos_label: "POS tickets Site",
      pos: posSite, src_label: "Site (livrées col L + emporter Fermée)", src: siteL,
      ecart: siteL - posSite,
    });

    var payment_breakdown = buildPaymentBreakdown(byPayment, matrix);

    var posSiteCash = pos.reduce(function (a, p) {
      if (p.channel !== CH_SITE) return a;
      return a + (allocatePosPaymentAmounts(p)["Cash"] || 0);
    }, 0);
    var siteCashSrc = 0;
    if (site) {
      site.filter(siteOrderCountsInReconciliation).forEach(function (o) {
        if (!siteOrderIsTakeout(o)) return;
        if (!isNaN(o.order_total)) siteCashSrc += o.order_total;
      });
    }

    var cash_to_collect = buildCashToCollect(lines, byPayment, posSiteCash, siteCashSrc);

    return { pays: PAYS, by_payment: byPayment, matrix: matrix, lines: lines,
             payment_breakdown: payment_breakdown,
             site_cash_pos: posSiteCash, site_cash_src: siteCashSrc,
             cash_to_collect: cash_to_collect };
  }

  /**
   * Cash à récupérer des caissiers :
   * - Glovo Cash / Site emporter : source > POS (écarts positifs source − POS)
   * - TPE : POS « Credit card » > relevé NAPS → cash à collecter
   * collect_amount > 0 = à collecter · < 0 = sur-saisie POS (réduit le net)
   */
  function buildCashToCollect(lines, byPayment, siteCashPos, siteCashSrc) {
    var items = [];
    var totalCollect = 0, totalOver = 0;

    function addItem(item) {
      var ca = item.collect_amount || 0;
      items.push(item);
      if (ca > 0.5) totalCollect += ca;
      else if (ca < -0.5) totalOver += -ca;
    }

    var glovoLine = lines.filter(function (l) { return l.lineKey === "glovo_cash"; })[0];
    if (glovoLine) {
      var gE = glovoLine.ecart;
      addItem({
        lineKey: "glovo_cash",
        label: glovoLine.source,
        pos_label: glovoLine.pos_label,
        src_label: glovoLine.src_label,
        pos: glovoLine.pos,
        src: glovoLine.src,
        ecart: gE,
        collect_amount: gE > 0.5 ? gE : (gE < -0.5 ? gE : 0),
        hint: "Glovo Cash (W − AE) > POS Glovo Cash : les livreurs ont encaissé plus que saisi au POS.",
      });
    }

    var siteEcart = siteCashSrc - siteCashPos;
    if (Math.abs(siteEcart) >= 0.5 || siteCashSrc > 0 || siteCashPos > 0) {
      addItem({
        lineKey: "site_cash",
        label: "🌐 Site — Cash (emporter)",
        pos_label: "POS Site « Cash »",
        src_label: "Site emporter Fermée",
        pos: siteCashPos,
        src: siteCashSrc,
        ecart: siteEcart,
        collect_amount: siteEcart > 0.5 ? siteEcart : (siteEcart < -0.5 ? siteEcart : 0),
        hint: "Commandes site à emporter payées en cash au comptoir — écart positif = sous-saisie POS.",
      });
    }

    var napsLine = lines.filter(function (l) { return l.lineKey === "naps"; })[0];
    if (napsLine) {
      var tpeOver = napsLine.pos - napsLine.src;
      if (tpeOver > 0.5) {
        addItem({
          lineKey: "naps_tpe_over",
          label: "💳 TPE — POS CB > relevé NAPS",
          pos_label: napsLine.pos_label,
          src_label: napsLine.src_label,
          pos: napsLine.pos,
          src: napsLine.src,
          ecart: napsLine.ecart,
          collect_amount: tpeOver,
          hint: "Le POS enregistre plus de carte que le relevé NAPS — à récupérer en cash des caissiers.",
        });
      }
    }

    var posCash = byPayment["Cash"] || 0;
    var net = totalCollect - totalOver;
    return {
      pos_cash_recorded: Math.round(posCash),
      total_to_collect: Math.round(totalCollect),
      total_over_recorded: Math.round(totalOver),
      net_to_collect: Math.round(net),
      cash_expected_physical: Math.round(posCash + net),
      items: items,
    };
  }

  /** Totaux POS par mode de paiement avec split canal (Glovo / SP&EMP / Site). */
  function buildPaymentBreakdown(byPayment, matrix) {
    function cell(channel, pay) {
      var ch = matrix[channel];
      if (!ch) return 0;
      return ch[pay] || 0;
    }
    function otherChannelsSum(pay, exclude) {
      var s = 0;
      Object.keys(matrix).forEach(function (ch) {
        if (exclude.indexOf(ch) >= 0) return;
        s += cell(ch, pay);
      });
      return s;
    }

    var cashKnown = [CH_GLOVO, CH_DINEIN, CH_SITE];
    var cashSplits = [
      { key: "glovo", label: "Glovo", amount: cell(CH_GLOVO, "Cash") },
      { key: "spemp", label: "SP&EMP", amount: cell(CH_DINEIN, "Cash") },
      { key: "site", label: "Site", amount: cell(CH_SITE, "Cash") },
    ];
    var cashOther = otherChannelsSum("Cash", cashKnown);
    if (Math.round(cashOther) !== 0) {
      cashSplits.push({ key: "autre", label: "Autre", amount: cashOther });
    }

    var ccKnown = [CH_DINEIN, CH_SITE];
    var ccSplits = [
      { key: "spemp", label: "SP&EMP", amount: cell(CH_DINEIN, "Credit card") },
      { key: "site", label: "Site", amount: cell(CH_SITE, "Credit card") },
    ];
    var ccOther = otherChannelsSum("Credit card", ccKnown);
    if (Math.round(ccOther) !== 0) {
      ccSplits.push({ key: "autre", label: "Autre", amount: ccOther });
    }

    var btKnown = [CH_GLOVO, CH_SITE];
    var btSplits = [
      { key: "glovo", label: "Glovo", amount: cell(CH_GLOVO, "Bank Transfer") },
      { key: "site", label: "Site", amount: cell(CH_SITE, "Bank Transfer") },
    ];
    var btOther = otherChannelsSum("Bank Transfer", btKnown);
    if (Math.round(btOther) !== 0) {
      btSplits.push({ key: "autre", label: "Autre", amount: btOther });
    }

    return {
      payments: [
        { key: "Cash", label: "Cash", icon: "💵", total: byPayment["Cash"] || 0, splits: cashSplits },
        { key: "Credit card", label: "CB (Credit card)", icon: "💳",
          total: byPayment["Credit card"] || 0, splits: ccSplits },
        { key: "Bank Transfer", label: "Bank Transfer", icon: "🏦",
          total: byPayment["Bank Transfer"] || 0, splits: btSplits },
      ],
    };
  }

  // Ajustements financiers pour anomalies validées (hors calcul d'écart).
  // Voir CURSOR_JOURNAL.md — chaque type retire le montant du côté qui crée l'écart.
  function emptyAdj() {
    return {
      glovo_pos_bt: 0, glovo_pos_cash: 0,
      glovo_src_online: 0, glovo_src_cash: 0,
      site_pos: 0, site_src: 0,
      pos_cc: 0, naps_src: 0,
    };
  }
  function financialAdjustment(a) {
    var adj = emptyAdj();
    var t = a.type;
    var ap = a.amount_pos, as = a.amount_source;
    if (ap == null || isNaN(ap)) ap = 0;
    if (as == null || isNaN(as)) as = 0;

    if (t === "Commande livrée absente du POS") {
      adj.site_src -= as;
    } else if (t === "Commande Glovo absente du POS") {
      if (a.payment_source === "Cash") adj.glovo_src_cash -= as;
      else adj.glovo_src_online -= as;
    } else if (t === "Ticket Glovo au POS sans commande correspondante") {
      if (a.payment_pos === "Cash") adj.glovo_pos_cash -= ap;
      else adj.glovo_pos_bt -= ap;
    } else if (t === "Ticket Site au POS sans commande correspondante") {
      adj.site_pos -= ap;
    } else if (t === "Commande site au POS — statut non livré") {
      adj.site_pos -= ap;
    } else if (t === "Commande Glovo Online absente du POS (écart financier)") {
      adj.glovo_src_online -= as;
    } else if (t === "Commande Glovo Cash absente du POS (écart financier)") {
      adj.glovo_src_cash -= as;
    } else if (t === "Ticket Glovo Online au POS sans appariement (écart financier)") {
      adj.glovo_pos_bt -= ap;
    } else if (t === "Ticket Glovo Cash au POS sans appariement (écart financier)") {
      adj.glovo_pos_cash -= ap;
    } else if (t === "Paiement POS absent du TPE") {
      adj.pos_cc -= ap;
    } else if (t === "Transaction TPE absente du POS") {
      adj.naps_src -= as;
    } else if (t === "Écart de montant" && a.source === "Site") {
      adj.site_src -= as;
      adj.site_pos -= ap;
    } else if (t === "Écart de montant" && a.source === "Glovo") {
      if (a.payment_source === "Cash") {
        adj.glovo_src_cash -= as;
        adj.glovo_pos_cash -= ap;
      } else {
        adj.glovo_src_online -= as;
        adj.glovo_pos_bt -= ap;
      }
    } else if (t === "Ticket en double (correction)") {
      if (a.source === "Glovo") {
        // payment_pos = mode de la copie en trop (à retirer du total POS)
        if (a.payment_pos === "Cash") adj.glovo_pos_cash -= ap;
        else adj.glovo_pos_bt -= ap;
      } else if (a.source === "Site") {
        adj.site_pos -= ap;
      }
    } else if (t === "Mode de paiement incorrect" && a.source === "Glovo" && ap) {
      var exp = a.payment_source, got = a.payment_pos;
      if (exp === "Bank Transfer" && got === "Cash") {
        adj.glovo_pos_cash -= ap;
        adj.glovo_pos_bt += ap;
      } else if (exp === "Cash" && got === "Bank Transfer") {
        adj.glovo_pos_bt -= ap;
        adj.glovo_pos_cash += ap;
      }
    }
    return adj;
  }
  function ecartDeltaForAdjustment(adj, lineKey) {
    if (lineKey === "glovo_online") return adj.glovo_src_online - adj.glovo_pos_bt;
    if (lineKey === "glovo_cash") return adj.glovo_src_cash - adj.glovo_pos_cash;
    if (lineKey === "glovo_total") {
      return (adj.glovo_src_online + adj.glovo_src_cash) -
             (adj.glovo_pos_bt + adj.glovo_pos_cash);
    }
    if (lineKey === "site") return adj.site_src - adj.site_pos;
    if (lineKey === "naps") return adj.naps_src - adj.pos_cc;
    return 0;
  }
  function ecartContributionForAnomaly(a, lineKey) {
    return -ecartDeltaForAdjustment(financialAdjustment(a), lineKey);
  }
  function getFinancialContributors(lineKey, anomalies) {
    return anomalies.filter(function (a) {
      return Math.abs(ecartContributionForAnomaly(a, lineKey)) >= 0.01;
    });
  }
  function sumFinancialContributions(lineKey, anomalies) {
    var s = 0;
    anomalies.forEach(function (a) { s += ecartContributionForAnomaly(a, lineKey); });
    return s;
  }
  function sumFinancialAdjustments(anomalies) {
    var tot = emptyAdj();
    anomalies.forEach(function (a) {
      var x = financialAdjustment(a);
      Object.keys(tot).forEach(function (k) { tot[k] += x[k]; });
    });
    return tot;
  }
  function applyFinancialAdjustments(fin, adj) {
    if (!fin || !adj) return fin;
    var lines = fin.lines.map(function (l) {
      var pos = l.pos, src = l.src;
      if (l.lineKey === "glovo_online") {
        pos += adj.glovo_pos_bt;
        src += adj.glovo_src_online;
      } else if (l.lineKey === "glovo_cash") {
        pos += adj.glovo_pos_cash;
        src += adj.glovo_src_cash;
      } else if (l.lineKey === "glovo_total") {
        pos += adj.glovo_pos_bt + adj.glovo_pos_cash;
        src += adj.glovo_src_online + adj.glovo_src_cash;
      } else if (l.lineKey === "site") {
        pos += adj.site_pos;
        src += adj.site_src;
      } else if (l.lineKey === "naps") {
        pos += adj.pos_cc;
        src += adj.naps_src;
      }
      return {
        lineKey: l.lineKey, source: l.source, pos_label: l.pos_label, src_label: l.src_label,
        pos: pos, src: src, ecart: src - pos,
        group: l.group, isTotal: l.isTotal, note: l.note,
      };
    });
    return { pays: fin.pays, by_payment: fin.by_payment, matrix: fin.matrix,
             lines: lines, adjustments_applied: true,
             payment_breakdown: fin.payment_breakdown,
             site_cash_pos: fin.site_cash_pos, site_cash_src: fin.site_cash_src,
             cash_to_collect: buildCashToCollect(lines, fin.by_payment,
               fin.site_cash_pos || 0, fin.site_cash_src || 0) };
  }

  function annotate(pos, anomalies) {
    var byTicketNo = {}, sevTicketNo = {};
    anomalies.forEach(function (a) {
      var no = a.pos_ticket_no;
      if (!no) return;
      (byTicketNo[no] = byTicketNo[no] || []).push("[" + a.type + "] " + a.detail);
      (sevTicketNo[no] = sevTicketNo[no] || {})[a.severity] = true;
    });
    pos.forEach(function (p) {
      var s = sevTicketNo[p.ticket_no] || {};
      p.statut = (s.haute || s.moyenne) ? "⚠️ Anomalie" : (s.info ? "ℹ️ Info" : "✅ OK");
      p.anomalies = (byTicketNo[p.ticket_no] || []).join(" | ");
    });
  }

  function buildSummary(pos, anomalies, glovo, naps, site) {
    var sev = { haute: 0, moyenne: 0, info: 0 };
    anomalies.forEach(function (a) { sev[a.severity] = (sev[a.severity] || 0) + 1; });
    var channels = {};
    pos.forEach(function (p) { channels[p.channel] = (channels[p.channel] || 0) + 1; });
    var posTotal = pos.reduce(function (a, p) { return a + (isNaN(p.total) ? 0 : p.total); }, 0);
    return {
      pos_transactions: pos.length,
      pos_total: posTotal,
      channels: channels,
      n_anomalies: sev.haute + sev.moyenne,  // les « info » ne sont pas des anomalies
      n_infos: sev.info,
      severity: sev,
      glovo_orders: glovo ? glovo.filter(function (g) {
        return (g.status || "").toLowerCase() === "delivered"; }).length : 0,
      naps_transactions: naps ? naps.length : 0,
      site_orders: site ? site.length : 0,
    };
  }

  // ----------------------------------------------------------------------- //
  // Export public
  // ----------------------------------------------------------------------- //
  var CNS = {
    loadPOS: loadPOS, loadGlovo: loadGlovo, loadNAPS: loadNAPS, loadSite: loadSite,
    classify: classify, run: run, listPosDates: listPosDates, runDailyBreakdown: runDailyBreakdown,
    sumFinancialAdjustments: sumFinancialAdjustments,
    applyFinancialAdjustments: applyFinancialAdjustments,
    getFinancialContributors: getFinancialContributors,
    ecartContributionForAnomaly: ecartContributionForAnomaly,
    sumFinancialContributions: sumFinancialContributions,
    CH: { GLOVO: CH_GLOVO, SITE: CH_SITE, DINEIN: CH_DINEIN,
          UNASSIGNED: CH_UNASSIGNED, OTHER: CH_OTHER },
  };
  root.CNS = CNS;
  if (typeof module !== "undefined" && module.exports) module.exports = CNS;
})(typeof window !== "undefined" ? window : globalThis);
