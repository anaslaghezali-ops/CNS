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

  function anomaly(o) {
    return {
      source: o.source, severity: o.severity, type: o.type,
      ticket_name: o.ticket_name || "", pos_datetime: o.pos_datetime || null,
      source_ref: o.source_ref || "", detail: o.detail,
      amount_pos: o.amount_pos == null ? null : o.amount_pos,
      amount_source: o.amount_source == null ? null : o.amount_source,
      payment_pos: o.payment_pos || "", payment_source: o.payment_source || "",
    };
  }

  // ----------------------------------------------------------------------- //
  // Lecture des feuilles -> lignes normalisées
  // ----------------------------------------------------------------------- //
  function sheetAOA(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, blankrows: false });
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
    var headers = aoa[hi].map(function (h) { return s(h); });
    var out = [];
    for (var i = hi + 1; i < aoa.length; i++) {
      var r = aoa[i], obj = {};
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
      out.push({ date_transaction: d, date: dateKey(d), montant: num(r["Montant"]) });
    });
    return out;
  }

  function loadSite(ws) {
    var raw = rowsWithHeader(ws, ["identifiant", "Order Total"]);
    var out = [];
    raw.forEach(function (r) {
      if (s(r["identifiant"]) === "") return;
      out.push({
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
    var usedOrphan = new Set();

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
            amount_pos: p.total, amount_source: o.order_total }));
        }
        if (p.payment_type !== SITE_EXPECTED_PAYMENT) {
          anomalies.push(anomaly({ source: "Site", severity: "haute",
            type: "Mode de paiement incorrect",
            detail: "Commande site " + sid + " : attendu '" + SITE_EXPECTED_PAYMENT +
                    "', trouvé '" + p.payment_type + "' au POS.",
            ticket_name: sid, pos_datetime: p.datetime, source_ref: sid,
            payment_pos: p.payment_type, payment_source: SITE_EXPECTED_PAYMENT }));
        }
      }
      // Une commande non livrée (refusée/annulée) PEUT être présente au POS.
    });

    // Faute de frappe : orphelin « site-like » (5 chiffres) de même montant/heure.
    unmatchedDelivered.forEach(function (o) {
      var sid = o.identifiant, best = -1, bestGap = Infinity;
      orphans.forEach(function (p, i) {
        if (usedOrphan.has(i)) return;
        if (isNaN(p.total) || Math.abs(p.total - o.order_total) > AMOUNT_TOL) return;
        if (!p.datetime || !o.created_at) return;
        var gap = Math.abs(minutesBetween(p.datetime, o.created_at));
        if (gap > SITE_TYPO_WINDOW_MIN) return;
        if (gap < bestGap) { bestGap = gap; best = i; }
      });
      if (best >= 0) {
        usedOrphan.add(best);
        var m = orphans[best];
        m.channel = CH_SITE;
        var payNote = m.payment_type !== SITE_EXPECTED_PAYMENT ?
          " ⚠️ De plus, son paiement est '" + m.payment_type + "' au lieu de '" +
          SITE_EXPECTED_PAYMENT + "'." : "";
        anomalies.push(anomaly({ source: "Site", severity: "moyenne",
          type: "Numéro de commande mal saisi (faute de frappe)",
          detail: "Commande site " + sid + " livrée : introuvable sous ce numéro, mais le " +
                  "ticket POS " + m.ticket_name + " correspond (même montant " +
                  o.order_total.toFixed(0) + " DH, +" + bestGap.toFixed(0) + " min, et " +
                  m.ticket_name + " n'existe pas dans le fichier site). Le caissier a " +
                  "probablement tapé " + m.ticket_name + " au lieu de " + sid + "." + payNote,
          ticket_name: m.ticket_name, pos_datetime: m.datetime, source_ref: sid,
          amount_pos: m.total, amount_source: o.order_total,
          payment_pos: m.payment_type, payment_source: SITE_EXPECTED_PAYMENT }));
      } else {
        missing.push(o);  // -> passe commune (ticket sans numéro) puis « absente »
      }
    });

    // Orphelins « site-like » non expliqués.
    orphans.forEach(function (p, i) {
      if (usedOrphan.has(i)) return;
      anomalies.push(anomaly({ source: "Site", severity: "moyenne",
        type: "Ticket Site au POS sans commande correspondante",
        detail: "Ticket POS " + p.ticket_name + " ressemble à une commande site " +
                "mais n'existe pas dans le fichier site.",
        ticket_name: p.ticket_name, pos_datetime: p.datetime,
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
    var posCC = pos.filter(function (p) { return p.payment_type === "Credit card"; })
                   .map(function (p) { return { date: dateKey(p.datetime), total: p.total }; });

    var napsDates = naps.map(function (n) { return n.date; }).filter(Boolean).sort();
    var nMin = napsDates[0], nMax = napsDates[napsDates.length - 1];

    var posDates = {};
    posCC.forEach(function (p) { if (p.date) posDates[p.date] = true; });

    // 1) Journées non couvertes
    Object.keys(posDates).sort().forEach(function (d) {
      if (!(d >= nMin && d <= nMax)) {
        var rows = posCC.filter(function (p) { return p.date === d; });
        var total = rows.reduce(function (a, p) { return a + (p.total || 0); }, 0);
        anomalies.push(anomaly({ source: "NAPS", severity: "moyenne",
          type: "Journée non couverte par le relevé NAPS",
          detail: rows.length + " paiement(s) 'Credit card' du " + d + " (" +
                  total.toFixed(0) + " DH) : le relevé NAPS fourni couvre du " + nMin +
                  " au " + nMax + " (décalage de télécollecte probable).",
          source_ref: d }));
      }
    });

    // 2) Rapprochement (date, montant) — nombre + montant
    Object.keys(posDates).sort().forEach(function (d) {
      if (!(d >= nMin && d <= nMax)) return;
      var posC = {}, napsC = {};
      posCC.filter(function (p) { return p.date === d; }).forEach(function (p) {
        var k = (Math.round(p.total * 100) / 100);
        posC[k] = (posC[k] || 0) + 1;
      });
      naps.filter(function (n) { return n.date === d; }).forEach(function (n) {
        var k = (Math.round(n.montant * 100) / 100);
        napsC[k] = (napsC[k] || 0) + 1;
      });
      var amounts = {};
      Object.keys(posC).forEach(function (k) { amounts[k] = true; });
      Object.keys(napsC).forEach(function (k) { amounts[k] = true; });
      Object.keys(amounts).map(Number).sort(function (a, b) { return a - b; }).forEach(function (amt) {
        var pc = posC[amt] || 0, nc = napsC[amt] || 0;
        if (pc > nc) {
          anomalies.push(anomaly({ source: "NAPS", severity: "haute",
            type: "Paiement POS absent du TPE",
            detail: (pc - nc) + " paiement(s) 'Credit card' de " + amt.toFixed(0) +
                    " DH le " + d + " au POS sans équivalent dans le relevé NAPS.",
            source_ref: d, amount_pos: amt }));
        } else if (nc > pc) {
          anomalies.push(anomaly({ source: "NAPS", severity: "haute",
            type: "Transaction TPE absente du POS",
            detail: (nc - pc) + " transaction(s) NAPS de " + amt.toFixed(0) +
                    " DH le " + d + " sans équivalent 'Credit card' au POS.",
            source_ref: d, amount_source: amt }));
        }
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
        payment_pos: p.payment_type, payment_source: exp }));
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
        amount_pos: p.total, amount_source: g.subtotal }));
    });

    // Commandes non appariées -> passe commune (ticket sans numéro) puis « absente ».
    delivered.forEach(function (g, gi) {
      if (matched.has(gi)) return;
      if (!g.received_at) {
        anomalies.push(anomaly({ source: "Glovo", severity: "moyenne",
          type: "Commande Glovo sans heure de réception",
          detail: "Commande Glovo " + g.order_id + " sans heure exploitable — " +
                  "rapprochement manuel nécessaire.", source_ref: g.order_id }));
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
        payment_pos: best.payment_type }));
    });

    demands.forEach(function (d, di) {
      if (matchedD.has(di)) return;
      if (d.src === "Site") {
        anomalies.push(anomaly({ source: "Site", severity: "haute",
          type: "Commande livrée absente du POS",
          detail: "Commande site " + d.id + " livrée mais introuvable dans le POS.",
          source_ref: String(d.id), amount_source: d.amount }));
      } else {
        anomalies.push(anomaly({ source: "Glovo", severity: "haute",
          type: "Commande Glovo absente du POS",
          detail: "Commande Glovo " + d.id + " reçue à " +
                  (d.ref ? dateKey(d.ref) + " " + hhmm(d.ref) : "?") + " (" + d.o.payment_type +
                  ", " + d.amount.toFixed(0) + " DH) non retrouvée au POS " +
                  "(aucun ticket au bon mode de paiement).",
          source_ref: String(d.id), amount_source: d.amount,
          payment_source: d.pay || "?" }));
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
      if (DINEIN_ALLOWED.indexOf(p.payment_type) === -1) {
        anomalies.push(anomaly({ source: "Sur place", severity: "moyenne",
          type: "Mode de paiement inattendu (sur place/emporter)",
          detail: "Ticket " + p.ticket_name + " sur place/emporter payé '" +
                  p.payment_type + "' (attendu Cash ou Credit card).",
          ticket_name: p.ticket_name, pos_datetime: p.datetime,
          amount_pos: p.total, payment_pos: p.payment_type }));
      }
    });
    return anomalies;
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

    annotate(pos, anomalies);
    var summary = buildSummary(pos, anomalies, glovo, naps, site);
    var dates = Array.from(posDates).sort();
    summary.pos_date_min = dates.length ? dates[0] : "";
    summary.pos_date_max = dates.length ? dates[dates.length - 1] : "";
    summary.glovo_excluded = glovoExcluded;
    summary.site_excluded = siteExcluded;
    return { anomalies: anomalies, pos: pos, summary: summary };
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
