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
  function anomaly(o) {
    return {
      source: o.source, severity: o.severity, type: o.type,
      ticket_name: o.ticket_name || "", pos_datetime: o.pos_datetime || null,
      source_ref: o.source_ref || "", detail: o.detail,
      amount_pos: o.amount_pos == null ? null : o.amount_pos,
      amount_source: o.amount_source == null ? null : o.amount_source,
      payment_pos: o.payment_pos || "", payment_source: o.payment_source || "",
      file: o.file || "", row: o.row == null ? "" : o.row,
      when: o.when || (o.pos_datetime ? dtFull(o.pos_datetime) : ""),
    };
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
  var RE_SPEMP = /^(sp|emp)\s*\d*$/i;

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
      var delivered = (o.delivery_status || "").toUpperCase() === "DELIVERED";
      var p = byName[sid];
      if (delivered) {
        if (!p) { unmatchedDelivered.push(o); return; }
        if (!isNaN(p.total) && Math.abs(p.total - o.order_total) > AMOUNT_TOL) {
          anomalies.push(anomaly({ source: "Site", severity: "haute",
            type: "Écart de montant",
            detail: "Commande site " + sid + " : " + o.order_total + " DH (site) vs " +
                    p.total + " DH (POS).",
            ticket_name: sid, pos_datetime: p.datetime, source_ref: sid,
            amount_pos: p.total, amount_source: o.order_total,
            file: "POS", row: p.row, when: dtFull(p.datetime) }));
        }
        if (p.payment_type !== SITE_EXPECTED_PAYMENT) {
          anomalies.push(anomaly({ source: "Site", severity: "haute",
            type: "Mode de paiement incorrect",
            detail: "Commande site " + sid + " : attendu '" + SITE_EXPECTED_PAYMENT +
                    "', trouvé '" + p.payment_type + "' au POS.",
            ticket_name: sid, pos_datetime: p.datetime, source_ref: sid,
            payment_pos: p.payment_type, payment_source: SITE_EXPECTED_PAYMENT,
            file: "POS", row: p.row, when: dtFull(p.datetime) }));
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
      var payNote = m.payment_type !== SITE_EXPECTED_PAYMENT ?
        " ⚠️ De plus, son paiement est '" + m.payment_type + "' au lieu de '" +
        SITE_EXPECTED_PAYMENT + "'." : "";
      anomalies.push(anomaly({ source: "Site", severity: "moyenne",
        type: "Numéro de commande mal saisi (faute de frappe)",
        detail: "Commande site " + o.identifiant + " livrée : introuvable sous ce numéro, " +
                "mais le ticket POS " + m.ticket_name + " correspond (même montant " +
                o.order_total.toFixed(0) + " DH, +" + pr.gap.toFixed(0) + " min, numéro " +
                "quasi identique). Le caissier a probablement tapé " + m.ticket_name +
                " au lieu de " + o.identifiant + "." + payNote,
        ticket_name: m.ticket_name, pos_datetime: m.datetime, source_ref: o.identifiant,
        amount_pos: m.total, amount_source: o.order_total,
        payment_pos: m.payment_type, payment_source: SITE_EXPECTED_PAYMENT,
        file: "POS", row: m.row, when: dtFull(m.datetime) }));
    });

    unmatchedDelivered.forEach(function (o, oi) {
      if (!matchedO.has(oi)) missing.push(o);  // -> passe commune puis « absente »
    });

    // Orphelins « site-like » (5 chiffres) non expliqués.
    orphans.forEach(function (p) {
      if (typoUsed.has(p.ticket_no)) return;
      anomalies.push(anomaly({ source: "Site", severity: "moyenne",
        type: "Ticket Site au POS sans commande correspondante",
        detail: "Ticket POS " + p.ticket_name + " ressemble à une commande site " +
                "mais n'existe pas dans le fichier site.",
        ticket_name: p.ticket_name, pos_datetime: p.datetime,
        amount_pos: p.total, payment_pos: p.payment_type,
        file: "POS", row: p.row, when: dtFull(p.datetime) }));
    });
    return { anomalies: anomalies, missing: missing };
  }

  // ----------------------------------------------------------------------- //
  // NAPS
  // ----------------------------------------------------------------------- //
  function reconcileNaps(pos, naps) {
    var anomalies = [];
    if (!naps.length) return anomalies;
    // Tickets POS « Credit card » (avec date/heure, ligne, ticket).
    var posCC = pos.filter(function (p) { return p.payment_type === "Credit card"; });

    var napsDates = naps.map(function (n) { return n.date; }).filter(Boolean).sort();
    var nMin = napsDates[0], nMax = napsDates[napsDates.length - 1];

    var posDates = {};
    posCC.forEach(function (p) { var d = dateKey(p.datetime); if (d) posDates[d] = true; });

    Object.keys(posDates).sort().forEach(function (d) {
      var posDay = posCC.filter(function (p) { return dateKey(p.datetime) === d; });

      // Journée non couverte par le relevé NAPS -> info (décalage télécollecte).
      if (!(d >= nMin && d <= nMax)) {
        var total = posDay.reduce(function (a, p) { return a + (isNaN(p.total) ? 0 : p.total); }, 0);
        anomalies.push(anomaly({ source: "NAPS", severity: "info",
          type: "Journée non couverte par le relevé NAPS",
          detail: posDay.length + " paiement(s) 'Credit card' du " + d + " (" +
                  total.toFixed(0) + " DH) : le relevé NAPS fourni couvre du " + nMin +
                  " au " + nMax + " (décalage de télécollecte probable).",
          source_ref: d }));
        return;
      }

      // Rapprochement transaction par transaction, par montant, sur la journée.
      var napsDay = naps.filter(function (n) { return n.date === d; });
      var napsUsed = new Set();
      posDay.forEach(function (p) {
        var amt = Math.round((isNaN(p.total) ? 0 : p.total) * 100) / 100;
        var found = -1;
        for (var i = 0; i < napsDay.length; i++) {
          if (napsUsed.has(i)) continue;
          if (Math.round(napsDay[i].montant * 100) / 100 === amt) { found = i; break; }
        }
        if (found >= 0) { napsUsed.add(found); return; }  // apparié -> OK
        anomalies.push(anomaly({ source: "NAPS", severity: "haute",
          type: "Paiement POS absent du TPE",
          detail: "Paiement 'Credit card' de " + amt.toFixed(0) + " DH au POS le " + d +
                  " à " + hhmm(p.datetime) + " (ticket " + p.ticket_name + ") sans équivalent " +
                  "dans le relevé NAPS.",
          ticket_name: p.ticket_name, pos_datetime: p.datetime, source_ref: p.ticket_no,
          amount_pos: amt, file: "POS", row: p.row, when: dtFull(p.datetime) }));
      });
      // Transactions NAPS non appariées.
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
      var amountMatch = !isNaN(g.subtotal) && !isNaN(p.total) &&
                        Math.abs(p.total - g.subtotal) <= AMOUNT_TOL;
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
      return !isNaN(g.subtotal) && !isNaN(p.total) &&
             Math.abs(p.total - g.subtotal) <= AMOUNT_TOL;
    }
    function inWindow(g, p, beforeMin, afterMin) {
      if (!p.datetime || !g.received_at) return false;
      return p.datetime >= new Date(g.received_at.getTime() - beforeMin * MIN) &&
             p.datetime <= new Date(g.received_at.getTime() + afterMin * MIN);
    }
    // Appariement GLOBAL glouton : on classe toutes les paires (commande, ticket)
    // éligibles par proximité temporelle croissante et on apparie les plus proches
    // d'abord — une commande ne peut plus « voler » le ticket d'une autre plus
    // proche. Le montant sert seulement de léger départage (il n'est pas fiable :
    // le POS colle tantôt au montant brut W, tantôt au net AP).
    function assignGlobal(payFilter, beforeMin, afterMin) {
      var pairs = [];
      delivered.forEach(function (g, gi) {
        if (matched.has(gi) || !g.received_at) return;
        var exp = GLOVO_PAYMENT_MAP[g.payment_type];
        pool.forEach(function (p, pi) {
          if (used.has(pi) || !payFilter(p.payment_type, exp)) return;
          if (!inWindow(g, p, beforeMin, afterMin)) return;
          var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
          pairs.push({ gi: gi, pi: pi, cost: gap - (amountMatch(g, p) ? 0.4 : 0) });
        });
      });
      pairs.sort(function (a, b) { return a.cost - b.cost; });
      var res = [];
      pairs.forEach(function (pr) {
        if (matched.has(pr.gi) || used.has(pr.pi)) return;
        matched.add(pr.gi); used.add(pr.pi); res.push(pr);
      });
      return res;
    }
    var isExp = function (pt, exp) { return pt === exp; };
    var isWrong = function (pt, exp) { return pt !== exp; };

    // Phase 1 : fenêtre proche, BON mode de paiement.
    assignGlobal(isExp, WINDOW_BEFORE_MIN, WINDOW_AFTER_MIN);

    // Phase 2 : fenêtre proche, MAUVAIS mode de paiement -> erreur de paiement.
    assignGlobal(isWrong, WINDOW_BEFORE_MIN, WINDOW_AFTER_MIN).forEach(function (pr) {
      var g = delivered[pr.gi], p = pool[pr.pi], exp = GLOVO_PAYMENT_MAP[g.payment_type];
      anomalies.push(anomaly({ source: "Glovo", severity: "haute",
        type: "Mode de paiement incorrect",
        detail: "Commande Glovo " + g.order_id + " (" + g.payment_type + ", " +
                g.subtotal.toFixed(0) + " DH) : attendu '" + exp + "' au POS, trouvé '" +
                p.payment_type + "' (ticket " + (p.ticket_name || p.ticket_no) +
                " à " + hhmm(p.datetime) + ").",
        ticket_name: p.ticket_name, pos_datetime: p.datetime, source_ref: g.order_id,
        payment_pos: p.payment_type, payment_source: exp,
        file: "POS", row: p.row, when: dtFull(p.datetime) }));
    });

    // Phase 3 : saisie tardive (fenêtre élargie ±120 min, bon mode de paiement).
    assignGlobal(isExp, 120, 120).forEach(function (pr) {
      var g = delivered[pr.gi], p = pool[pr.pi];
      var delay = minutesBetween(p.datetime, g.received_at);
      anomalies.push(anomaly({ source: "Glovo", severity: "info",
        type: "Saisie tardive (hors fenêtre 10 min)",
        detail: "Commande Glovo " + g.order_id + " (" + g.subtotal.toFixed(0) +
                " DH) reçue à " + hhmm(g.received_at) + ", tapée au POS à " +
                hhmm(p.datetime) + " (ticket " + (p.ticket_name || p.ticket_no) + ", " +
                (delay >= 0 ? "+" : "") + delay.toFixed(0) + " min) — présente mais tardive.",
        ticket_name: p.ticket_name, pos_datetime: p.datetime, source_ref: g.order_id,
        amount_pos: p.total, amount_source: g.subtotal,
        file: "POS", row: p.row, when: dtFull(p.datetime) }));
    });

    // Commandes non appariées -> passe commune (ticket sans numéro) puis « absente ».
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
        anomalies.push(anomaly({ source: "Glovo", severity: "moyenne",
          type: "Ticket Glovo au POS sans commande correspondante",
          detail: "Ticket POS " + p.ticket_name + " (" + hhmm(p.datetime) + ", " +
                  p.payment_type + ", " + (isNaN(p.total) ? "?" : p.total.toFixed(0)) +
                  " DH) classé Glovo mais sans commande Glovo dans la fenêtre.",
          ticket_name: p.ticket_name, pos_datetime: p.datetime,
          amount_pos: p.total, payment_pos: p.payment_type,
          file: "POS", row: p.row, when: dtFull(p.datetime) }));
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
      demands.push({ src: "Glovo", ref: g.received_at, amount: g.subtotal,
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
      // Commande retrouvée sous un ticket sans numéro : info (pas une anomalie).
      anomalies.push(anomaly({ source: d.src, severity: "info",
        type: "Commande rattachée (ticket sans numéro)",
        detail: "Commande " + d.src + " " + d.id + " (" +
                (d.src === "Glovo" ? d.o.payment_type + ", " : "") + d.amount.toFixed(0) +
                " DH) retrouvée au POS sous le ticket sans numéro " + best.ticket_no +
                " (+" + pr.gap.toFixed(0) + " min). Présente — simple oubli de numéro.",
        ticket_name: best.ticket_name || "(vide)", pos_datetime: best.datetime,
        source_ref: String(d.id), amount_pos: best.total, amount_source: d.amount,
        payment_pos: best.payment_type,
        file: "POS", row: best.row, when: dtFull(best.datetime) }));
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
      anomalies.push(anomaly({ source: "POS", severity: "info",
        type: "Ticket sans numéro (à rattacher)",
        detail: "Ticket " + p.ticket_no + " du " + dateKey(p.datetime) + " à " +
                hhmm(p.datetime) + " (" + p.payment_type + ", " +
                (isNaN(p.total) ? "?" : p.total) + " DH) sans numéro — non rattaché à " +
                "une commande Glovo ni Site.",
        ticket_name: p.ticket_name || "(vide)", pos_datetime: p.datetime,
        amount_pos: p.total, payment_pos: p.payment_type,
        file: "POS", row: p.row, when: dtFull(p.datetime) }));
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
      if (DINEIN_ALLOWED.indexOf(p.payment_type) === -1) {
        anomalies.push(anomaly({ source: "Sur place", severity: "moyenne",
          type: "Mode de paiement inattendu (sur place/emporter)",
          detail: "Ticket " + p.ticket_name + " sur place/emporter payé '" +
                  p.payment_type + "' (attendu Cash ou Credit card).",
          ticket_name: p.ticket_name, pos_datetime: p.datetime,
          amount_pos: p.total, payment_pos: p.payment_type,
          file: "POS", row: p.row, when: dtFull(p.datetime) }));
      }
    });
    return anomalies;
  }

  // ----------------------------------------------------------------------- //
  // DOUBLONS / CORRECTIONS : un même NUMÉRO (Glovo 1-3 chiffres, Site) saisi
  // plusieurs fois à quelques minutes d'intervalle = commande re-tapée (souvent
  // une correction). On ne traite QUE les numéros (jamais sp/emp, qui peuvent
  // légitimement se répéter). Remplace l'anomalie « orphelin » par une
  // « correction » détaillant les changements (montant, paiement).
  // ----------------------------------------------------------------------- //
  var DUP_WINDOW_MIN = 30;

  function detectDuplicates(pos, anomalies) {
    var groups = {};
    pos.forEach(function (p) {
      if ((p.channel === CH_GLOVO || p.channel === CH_SITE) && /^\d+$/.test(p.ticket_name)) {
        (groups[p.ticket_name] = groups[p.ticket_name] || []).push(p);
      }
    });
    var handled = {};  // ticket_name -> true : retirer l'anomalie orpheline
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
        handled[name] = true;
        var first = cl[0], last = cl[cl.length - 1];
        var versions = cl.map(function (p, idx) {
          return "v" + (idx + 1) + " " + hhmm(p.datetime) + " " +
                 (isNaN(p.total) ? "?" : p.total.toFixed(0)) + " DH " + p.payment_type +
                 " (ligne " + p.row + ")";
        }).join(" → ");
        var diffs = [];
        if (!isNaN(first.total) && !isNaN(last.total) && Math.abs(first.total - last.total) > AMOUNT_TOL)
          diffs.push("montant " + first.total.toFixed(0) + "→" + last.total.toFixed(0));
        if (first.payment_type !== last.payment_type)
          diffs.push("paiement " + first.payment_type + "→" + last.payment_type);
        var extra = cl.reduce(function (a, p) { return a + (isNaN(p.total) ? 0 : p.total); }, 0) -
                    (isNaN(last.total) ? 0 : last.total);
        anomalies.push(anomaly({
          source: cl[0].channel === CH_SITE ? "Site" : "Glovo", severity: "moyenne",
          type: "Ticket en double (correction)",
          detail: "Ticket " + name + " saisi " + cl.length + " fois (doublon / correction) : " +
                  versions + ". " + (diffs.length ? "Correction : " + diffs.join(", ") + ". " : "") +
                  "⚠️ La/les copie(s) en trop gonflent le total POS de " + extra.toFixed(0) +
                  " DH — vérifier qu'une version est bien annulée.",
          ticket_name: name, pos_datetime: last.datetime, source_ref: name,
          amount_pos: extra, file: "POS", row: last.row, when: dtFull(last.datetime) }));
      });
    });

    // Retirer les anomalies « orphelin » remplacées par une correction.
    return anomalies.filter(function (a) {
      if (handled[a.ticket_name] &&
          (a.type === "Ticket Glovo au POS sans commande correspondante" ||
           a.type === "Ticket Site au POS sans commande correspondante")) return false;
      return true;
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
    anomalies = detectDuplicates(pos, anomalies);  // doublons/corrections

    annotate(pos, anomalies);
    var summary = buildSummary(pos, anomalies, glovo, naps, site);
    var dates = Array.from(posDates).sort();
    summary.pos_date_min = dates.length ? dates[0] : "";
    summary.pos_date_max = dates.length ? dates[dates.length - 1] : "";
    summary.glovo_excluded = glovoExcluded;
    summary.site_excluded = siteExcluded;
    summary.financial = computeFinancial(pos, glovo, naps, site, posDates);
    return { anomalies: anomalies, pos: pos, summary: summary };
  }

  // Réconciliation FINANCIÈRE : totaux par mode de paiement × canal, et écarts
  // POS vs source (TPE/NAPS, Glovo, Site).
  function computeFinancial(pos, glovo, naps, site, posDates) {
    var PAYS = ["Cash", "Bank Transfer", "Credit card"];
    var byPayment = { "Cash": 0, "Bank Transfer": 0, "Credit card": 0, "Autre": 0 };
    var matrix = {};  // canal -> { paiement -> somme }
    pos.forEach(function (p) {
      var t = isNaN(p.total) ? 0 : p.total;
      var pay = PAYS.indexOf(p.payment_type) >= 0 ? p.payment_type : "Autre";
      byPayment[pay] += t;
      matrix[p.channel] = matrix[p.channel] || { "Cash": 0, "Bank Transfer": 0, "Credit card": 0, "Autre": 0 };
      matrix[p.channel][pay] += t;
    });

    function sumPos(channel) {
      return pos.reduce(function (a, p) {
        return a + (p.channel === channel && !isNaN(p.total) ? p.total : 0); }, 0);
    }
    var posCC = byPayment["Credit card"];
    var napsTotal = 0;
    if (naps) naps.forEach(function (n) {
      if (posDates.has(n.date) && !isNaN(n.montant)) napsTotal += n.montant; });

    var posGlovo = sumPos(CH_GLOVO), glovoW = 0;
    if (glovo) glovo.filter(function (g) { return (g.status || "").toLowerCase() === "delivered"; })
      .forEach(function (g) { if (!isNaN(g.subtotal)) glovoW += g.subtotal; });

    var posSite = sumPos(CH_SITE), siteL = 0;
    if (site) site.filter(function (o) { return (o.delivery_status || "").toUpperCase() === "DELIVERED"; })
      .forEach(function (o) { if (!isNaN(o.order_total)) siteL += o.order_total; });

    var lines = [];
    if (naps) lines.push({ source: "💳 TPE (NAPS)", pos_label: "POS « Credit card »",
      pos: posCC, src_label: "Relevé NAPS", src: napsTotal, ecart: napsTotal - posCC });
    if (glovo) lines.push({ source: "🛵 Glovo", pos_label: "POS tickets Glovo",
      pos: posGlovo, src_label: "Glovo (col W)", src: glovoW, ecart: glovoW - posGlovo });
    if (site) lines.push({ source: "🌐 Site", pos_label: "POS tickets Site",
      pos: posSite, src_label: "Site livrées (col L)", src: siteL, ecart: siteL - posSite });

    return { pays: PAYS, by_payment: byPayment, matrix: matrix, lines: lines };
  }

  function annotate(pos, anomalies) {
    var byTicket = {}, sevTicket = {};
    anomalies.forEach(function (a) {
      var tn = a.ticket_name;
      if (tn && tn !== "" && tn !== "(vide)") {
        (byTicket[tn] = byTicket[tn] || []).push("[" + a.type + "] " + a.detail);
        (sevTicket[tn] = sevTicket[tn] || {})[a.severity] = true;
      }
    });
    pos.forEach(function (p) {
      var s = sevTicket[p.ticket_name] || {};
      p.statut = (s.haute || s.moyenne) ? "⚠️ Anomalie" : (s.info ? "ℹ️ Info" : "✅ OK");
      p.anomalies = (byTicket[p.ticket_name] || []).join(" | ");
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
    classify: classify, run: run,
    CH: { GLOVO: CH_GLOVO, SITE: CH_SITE, DINEIN: CH_DINEIN,
          UNASSIGNED: CH_UNASSIGNED, OTHER: CH_OTHER },
  };
  root.CNS = CNS;
  if (typeof module !== "undefined" && module.exports) module.exports = CNS;
})(typeof window !== "undefined" ? window : globalThis);
