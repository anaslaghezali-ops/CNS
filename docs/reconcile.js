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

  function reconcileSite(pos, site) {
    var anomalies = [];
    var byName = {};
    pos.forEach(function (p) { byName[p.ticket_name] = p; });
    var siteIds = new Set(site.map(function (x) { return x.identifiant; }));

    // Tickets POS ressemblant à une commande site mais absents du fichier site
    // (orphelins) — candidats à une faute de frappe sur le numéro.
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
        if (!p) { unmatchedDelivered.push(o); return; }  // -> détection faute de frappe
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
      } else if (p) {
        anomalies.push(anomaly({ source: "Site", severity: "haute",
          type: "Commande non livrée mais tapée au POS",
          detail: "Commande site " + sid + " refusée/non livrée (statut '" +
                  o.last_status + "') mais présente au POS.",
          ticket_name: sid, pos_datetime: p.datetime, source_ref: sid,
          amount_pos: p.total }));
      }
    });

    // Commandes livrées introuvables : tenter une faute de frappe sur le numéro.
    unmatchedDelivered.forEach(function (o) {
      var sid = o.identifiant;
      var best = -1, bestGap = Infinity;
      for (var i = 0; i < orphans.length; i++) {
        if (usedOrphan.has(i)) continue;
        var p = orphans[i];
        if (isNaN(p.total) || Math.abs(p.total - o.order_total) > AMOUNT_TOL) continue;
        if (!p.datetime || !o.created_at) continue;
        var gap = Math.abs(minutesBetween(p.datetime, o.created_at));
        if (gap > SITE_TYPO_WINDOW_MIN) continue;
        if (gap < bestGap) { bestGap = gap; best = i; }
      }
      if (best >= 0) {
        usedOrphan.add(best);
        var m = orphans[best];
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
        anomalies.push(anomaly({ source: "Site", severity: "haute",
          type: "Commande livrée absente du POS",
          detail: "Commande site " + sid + " livrée mais introuvable dans le POS.",
          source_ref: sid, amount_source: o.order_total }));
      }
    });

    // Orphelins non expliqués par une faute de frappe.
    orphans.forEach(function (p, i) {
      if (usedOrphan.has(i)) return;
      anomalies.push(anomaly({ source: "Site", severity: "moyenne",
        type: "Ticket Site au POS sans commande correspondante",
        detail: "Ticket POS " + p.ticket_name + " ressemble à une commande site " +
                "mais n'existe pas dans le fichier site.",
        ticket_name: p.ticket_name, pos_datetime: p.datetime,
        amount_pos: p.total, payment_pos: p.payment_type }));
    });
    return anomalies;
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
  function glovoNearest(g, posGlovo, used, lo, hi, payment, requireAmount) {
    var best = -1, bestGap = Infinity;
    for (var i = 0; i < posGlovo.length; i++) {
      if (used.has(i)) continue;
      var p = posGlovo[i];
      if (!p.datetime || p.datetime < lo || p.datetime > hi) continue;
      if (payment != null && p.payment_type !== payment) continue;
      if (requireAmount && !isNaN(g.subtotal) && !isNaN(p.total) &&
          Math.abs(p.total - g.subtotal) > AMOUNT_TOL) continue;
      var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
      if (gap < bestGap) { bestGap = gap; best = i; }
    }
    return best;
  }

  function reconcileGlovo(pos, glovo) {
    var anomalies = [];
    var delivered = glovo.filter(function (g) {
      return (g.status || "").toLowerCase() === "delivered";
    }).slice().sort(function (a, b) {
      return (a.received_at ? a.received_at.getTime() : 0) -
             (b.received_at ? b.received_at.getTime() : 0);
    });
    var posGlovo = pos.filter(function (p) { return p.channel === CH_GLOVO; })
                      .slice().sort(function (a, b) {
      return (a.datetime ? a.datetime.getTime() : 0) - (b.datetime ? b.datetime.getTime() : 0);
    });

    var used = new Set();
    var MIN = 60000;

    // Agrégat fiable par mode de paiement
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

    // Passe 1 : même montant + bon paiement
    var remaining = [];
    delivered.forEach(function (g) {
      if (!g.received_at) { remaining.push(g); return; }
      var lo = new Date(g.received_at.getTime() - WINDOW_BEFORE_MIN * MIN);
      var hi = new Date(g.received_at.getTime() + WINDOW_AFTER_MIN * MIN);
      var exp = GLOVO_PAYMENT_MAP[g.payment_type];
      var idx = glovoNearest(g, posGlovo, used, lo, hi, exp, true);
      if (idx >= 0) used.add(idx);
      else remaining.push(g);
    });

    // Passe 2 : même montant, paiement quelconque -> erreur de paiement
    var remaining2 = [];
    remaining.forEach(function (g) {
      if (!g.received_at) {
        anomalies.push(anomaly({ source: "Glovo", severity: "moyenne",
          type: "Commande Glovo sans heure de réception",
          detail: "Commande Glovo " + g.order_id + " sans heure exploitable — " +
                  "rapprochement manuel nécessaire.", source_ref: g.order_id }));
        return;
      }
      var lo = new Date(g.received_at.getTime() - WINDOW_BEFORE_MIN * MIN);
      var hi = new Date(g.received_at.getTime() + WINDOW_AFTER_MIN * MIN);
      var idx = glovoNearest(g, posGlovo, used, lo, hi, null, true);
      if (idx >= 0) {
        used.add(idx);
        var p = posGlovo[idx];
        var exp = GLOVO_PAYMENT_MAP[g.payment_type];
        anomalies.push(anomaly({ source: "Glovo", severity: "haute",
          type: "Mode de paiement incorrect",
          detail: "Commande Glovo " + g.order_id + " (" + g.payment_type + ", " +
                  g.subtotal.toFixed(0) + " DH) : attendu '" + exp + "' au POS, trouvé '" +
                  p.payment_type + "' (ticket " + p.ticket_name + " à " + hhmm(p.datetime) + ").",
          ticket_name: p.ticket_name, pos_datetime: p.datetime, source_ref: g.order_id,
          payment_pos: p.payment_type, payment_source: exp }));
      } else remaining2.push(g);
    });

    // Passe 3 : rattrapage hors fenêtre (saisie tardive)
    remaining2.forEach(function (g) {
      var exp = GLOVO_PAYMENT_MAP[g.payment_type];
      var big = 24 * 60 * MIN;
      var lo = new Date(g.received_at.getTime() - big);
      var hi = new Date(g.received_at.getTime() + big);
      var idx = glovoNearest(g, posGlovo, used, lo, hi, exp, true);
      if (idx >= 0) {
        used.add(idx);
        var p = posGlovo[idx];
        var delay = minutesBetween(p.datetime, g.received_at);
        anomalies.push(anomaly({ source: "Glovo", severity: "info",
          type: "Saisie tardive (hors fenêtre 10 min)",
          detail: "Commande Glovo " + g.order_id + " (" + g.subtotal.toFixed(0) +
                  " DH) reçue à " + hhmm(g.received_at) + ", tapée au POS à " +
                  hhmm(p.datetime) + " (ticket " + p.ticket_name + ", +" +
                  delay.toFixed(0) + " min) — présente mais tardive.",
          ticket_name: p.ticket_name, pos_datetime: p.datetime, source_ref: g.order_id,
          amount_pos: p.total, amount_source: g.subtotal }));
      } else {
        anomalies.push(anomaly({ source: "Glovo", severity: "haute",
          type: "Commande Glovo absente du POS",
          detail: "Commande Glovo " + g.order_id + " reçue à " +
                  dateKey(g.received_at) + " " + hhmm(g.received_at) + " (" +
                  g.payment_type + ", " + g.subtotal.toFixed(0) + " DH) non retrouvée au POS " +
                  "(aucun ticket de même montant et mode de paiement).",
          source_ref: g.order_id, amount_source: g.subtotal,
          payment_source: GLOVO_PAYMENT_MAP[g.payment_type] || "?" }));
      }
    });

    // Tickets POS Glovo non appariés
    posGlovo.forEach(function (p, i) {
      if (!used.has(i)) {
        anomalies.push(anomaly({ source: "Glovo", severity: "moyenne",
          type: "Ticket Glovo au POS sans commande correspondante",
          detail: "Ticket POS " + p.ticket_name + " (" + hhmm(p.datetime) + ", " +
                  p.payment_type + ", " + (isNaN(p.total) ? "?" : p.total.toFixed(0)) +
                  " DH) classé Glovo mais sans commande Glovo de même montant dans la fenêtre.",
          ticket_name: p.ticket_name, pos_datetime: p.datetime,
          amount_pos: p.total, payment_pos: p.payment_type }));
      }
    });
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

  function flagUnassigned(pos, glovo) {
    var anomalies = [];
    var delivered = (glovo || []).filter(function (g) {
      return (g.status || "").toLowerCase() === "delivered" && g.received_at;
    });
    pos.filter(function (p) { return p.channel === CH_UNASSIGNED; }).forEach(function (p) {
      var hint = "";
      if (delivered.length && p.datetime) {
        var best = null, bestGap = Infinity;
        delivered.forEach(function (g) {
          var gap = Math.abs(minutesBetween(p.datetime, g.received_at));
          if (gap < bestGap) { bestGap = gap; best = g; }
        });
        if (best) hint = " Suggestion : commande Glovo " + best.order_id +
                         " reçue à " + hhmm(best.received_at) + ".";
      }
      anomalies.push(anomaly({ source: "POS", severity: "moyenne",
        type: "Ticket sans nom (à rattacher)",
        detail: "Ticket " + p.ticket_no + " à " + hhmm(p.datetime) + " (" +
                p.payment_type + ", " + (isNaN(p.total) ? "?" : p.total) + " DH) sans ticket name." + hint,
        ticket_name: "(vide)", pos_datetime: p.datetime,
        amount_pos: p.total, payment_pos: p.payment_type }));
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

    var anomalies = [];
    if (site) anomalies = anomalies.concat(reconcileSite(pos, site));
    if (naps) anomalies = anomalies.concat(reconcileNaps(pos, naps));
    if (glovo) anomalies = anomalies.concat(reconcileGlovo(pos, glovo));
    anomalies = anomalies.concat(reconcileDinein(pos));
    anomalies = anomalies.concat(flagUnassigned(pos, glovo));

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
    var byTicket = {};
    anomalies.forEach(function (a) {
      var tn = a.ticket_name;
      if (tn && tn !== "" && tn !== "(vide)") {
        (byTicket[tn] = byTicket[tn] || []).push("[" + a.type + "] " + a.detail);
      }
    });
    pos.forEach(function (p) {
      var issues = byTicket[p.ticket_name] || [];
      p.statut = issues.length ? "⚠️ Anomalie" : "✅ OK";
      p.anomalies = issues.join(" | ");
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
      n_anomalies: anomalies.length,
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
