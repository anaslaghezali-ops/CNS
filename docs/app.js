/* ChickNSter — câblage de l'interface + export Excel (navigateur). */
(function () {
  "use strict";

  var files = { pos: null, glovo: null, naps: null, site: null };
  var state = null; // { anomalies, pos, summary, validated, cashUserAssignments }

  var SEV_BADGE = { haute: "🔴 Haute", moyenne: "🟠 Moyenne", info: "🔵 Info" };
  var SEV_ORDER = { haute: 0, moyenne: 1, info: 2 };
  var finDetailLineKey = null;
  var activeDayKey = null;  // null = vue générale · sinon clé YYYY-MM-DD

  var MONTHS_FR = ["janv.", "févr.", "mars", "avr.", "mai", "juin",
                   "juil.", "août", "sept.", "oct.", "nov.", "déc."];

  // ---- Sélection de fichiers -------------------------------------------- //
  document.querySelectorAll(".drop").forEach(function (drop) {
    var key = drop.getAttribute("data-key");
    var input = drop.querySelector("input");
    var stateEl = drop.querySelector(".drop-state");
    input.addEventListener("change", function () {
      var f = input.files[0];
      files[key] = f || null;
      if (f) {
        drop.classList.add("filled");
        stateEl.textContent = "✓ " + f.name;
      } else {
        drop.classList.remove("filled");
        stateEl.textContent = "Aucun fichier";
      }
    });
  });

  function readSheet(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var wb = XLSX.read(new Uint8Array(e.target.result), {
            type: "array", cellDates: true,
          });
          resolve(wb.Sheets[wb.SheetNames[0]]);
        } catch (err) { reject(err); }
      };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsArrayBuffer(file);
    });
  }

  function setMessage(html) { document.getElementById("message").innerHTML = html || ""; }

  document.getElementById("run").addEventListener("click", function () {
    setMessage("");
    if (!files.pos) {
      setMessage('<div class="warn">⚠️ Le fichier <b>POS</b> est obligatoire.</div>');
      return;
    }
    var btn = this;
    btn.disabled = true;
    btn.textContent = "⏳ Réconciliation en cours…";

    var jobs = [readSheet(files.pos)];
    ["glovo", "naps", "site"].forEach(function (k) {
      jobs.push(files[k] ? readSheet(files[k]) : Promise.resolve(null));
    });

    Promise.all(jobs).then(function (sheets) {
      try {
        var pos = CNS.loadPOS(sheets[0]);
        var glovo = sheets[1] ? CNS.loadGlovo(sheets[1]) : null;
        var naps = sheets[2] ? CNS.loadNAPS(sheets[2]) : null;
        var site = sheets[3] ? CNS.loadSite(sheets[3]) : null;
        state = CNS.run(pos, glovo, naps, site);
        state.byDay = CNS.runDailyBreakdown(pos, glovo, naps, site);
        state.validated = {};
        state.cashUserAssignments = {};
        activeDayKey = null;
        finDetailLineKey = null;
        render();
      } catch (err) {
        setMessage('<div class="err">❌ Erreur : ' + escapeHtml(err.message) + "</div>");
        console.error(err);
      }
    }).catch(function (err) {
      setMessage('<div class="err">❌ Erreur de lecture d\'un fichier : ' +
                 escapeHtml(err.message) + "</div>");
      console.error(err);
    }).then(function () {
      btn.disabled = false;
      btn.textContent = "🚀 Lancer la réconciliation";
    });
  });

  // ---- Validation des anomalies ----------------------------------------- //
  function isValidated(id) {
    return state && state.validated && state.validated[id];
  }

  function getViewState() {
    if (!state) return null;
    if (activeDayKey && state.byDay && state.byDay[activeDayKey]) return state.byDay[activeDayKey];
    return state;
  }

  function fmtDayLabel(dk) {
    var p = dk.split("-");
    if (p.length !== 3) return dk;
    return p[2] + " " + MONTHS_FR[+p[1] - 1] + " " + p[0];
  }

  function fmtDayShort(dk) {
    var p = dk.split("-");
    if (p.length !== 3) return dk;
    return p[2] + "/" + p[1];
  }

  function activeAnomalies() {
    var vs = getViewState();
    if (!vs) return [];
    return vs.anomalies.filter(function (a) { return !isValidated(a.id); });
  }

  function validatedAnomalies() {
    var vs = getViewState();
    if (!vs) return [];
    return vs.anomalies.filter(function (a) { return isValidated(a.id); });
  }

  function fmtTicketLabel(a) {
    var parts = [];
    if (a.ticket_name && a.ticket_name !== "(vide)") parts.push(a.ticket_name);
    if (a.pos_ticket_no) parts.push("#" + a.pos_ticket_no);
    if (a.when) {
      var m = String(a.when).match(/\d{2}:\d{2}/);
      if (m) parts.push(m[0]);
    }
    if (!parts.length) return a.ticket_name || "";
    return parts.join(" · ");
  }

  function anomalyMatchesPos(a, p) {
    if (a.pos_ticket_no && p.ticket_no) return a.pos_ticket_no === p.ticket_no;
    return false;
  }

  function validateAnomaly(id) {
    if (!state) return;
    state.validated[id] = true;
    refreshPosStatuts();
    render();
  }

  function unvalidateAnomaly(id) {
    if (!state || !state.validated[id]) return;
    delete state.validated[id];
    refreshPosStatuts();
    render();
  }

  function refreshPosStatutsFor(runResult) {
    if (!runResult) return;
    var active = runResult.anomalies.filter(function (a) { return !isValidated(a.id); });
    runResult.pos.forEach(function (p) {
      var related = active.filter(function (a) { return anomalyMatchesPos(a, p); });
      var sev = {};
      related.forEach(function (a) { sev[a.severity] = true; });
      p.statut = (sev.haute || sev.moyenne) ? "⚠️ Anomalie" :
                 (sev.info ? "ℹ️ Info" : "✅ OK");
      p.anomalies = related.map(function (a) {
        return "[" + a.type + "] " + a.detail;
      }).join(" | ");
    });
  }

  function refreshPosStatuts() {
    refreshPosStatutsFor(state);
    if (state && state.byDay) {
      Object.keys(state.byDay).forEach(function (dk) {
        refreshPosStatutsFor(state.byDay[dk]);
      });
    }
  }

  function isNapsTotalsOk(a) {
    return a && a.naps_totals_ok;
  }

  function isActionableAnomaly(a) {
    return a.severity !== "info" && !isNapsTotalsOk(a);
  }

  function recomputeCounts() {
    var active = activeAnomalies();
    var sev = { haute: 0, moyenne: 0, info: 0 };
    var napsPairing = 0;
    active.forEach(function (a) {
      if (isNapsTotalsOk(a) && a.severity !== "info") napsPairing++;
      else sev[a.severity] = (sev[a.severity] || 0) + 1;
    });
    return {
      n_anomalies: sev.haute + sev.moyenne,
      n_infos: sev.info,
      severity: sev,
      n_validated: validatedAnomalies().length,
      n_naps_pairing_ok: napsPairing,
    };
  }

  function getAdjustedFinancial() {
    var vs = getViewState();
    if (!vs || !vs.summary.financial) return null;
    var fin;
    var validated = validatedAnomalies();
    if (!validated.length) fin = vs.summary.financial;
    else {
      fin = CNS.applyFinancialAdjustments(vs.summary.financial,
        CNS.sumFinancialAdjustments(validated));
    }
    enrichFinCashCollect(fin, vs);
    return fin;
  }

  function enrichFinCashCollect(fin, viewState) {
    if (!fin || !fin.cash_to_collect || !viewState || !viewState.pos) return fin;
    CNS.finalizeCashToCollectUsers(
      fin.cash_to_collect, viewState.pos, viewState.glovo, viewState.site,
      viewState.naps, CNS.listPosDates(viewState.pos), viewState.anomalies || []);
    CNS.applyCashUserAssignments(fin.cash_to_collect, state.cashUserAssignments || {});
    return fin;
  }

  function assignCashContribution(contribId, userName) {
    if (!state || !contribId || !userName) return;
    state.cashUserAssignments[contribId] = userName;
    render();
  }

  function unassignCashContribution(contribId) {
    if (!state || !state.cashUserAssignments || !state.cashUserAssignments[contribId]) return;
    delete state.cashUserAssignments[contribId];
    render();
  }

  function listViewPosUsers() {
    var vs = getViewState();
    if (!vs || !vs.pos) return [];
    return CNS.listPosUsersWithActivity(vs.pos);
  }

  function isUnattributedCashContrib(t) {
    return (t.original_user || t.user) === CNS.CASH_COLLECT_UNATTRIBUTED;
  }

  function renderDayTabs() {
    var el = document.getElementById("day-tabs");
    if (!state || !state.byDay || Object.keys(state.byDay).length <= 1) {
      el.classList.add("hidden");
      el.innerHTML = "";
      return;
    }
    el.classList.remove("hidden");
    var dates = Object.keys(state.byDay).sort();
    var html = '<button type="button" class="day-tab' +
      (activeDayKey === null ? " active" : "") + '" data-day="">📊 Général</button>';
    dates.forEach(function (d) {
      html += '<button type="button" class="day-tab' +
        (activeDayKey === d ? " active" : "") + '" data-day="' + escapeHtml(d) + '">📅 ' +
        escapeHtml(fmtDayLabel(d)) + "</button>";
    });
    el.innerHTML = html;
    el.querySelectorAll(".day-tab").forEach(function (btn) {
      btn.onclick = function () {
        var day = btn.getAttribute("data-day");
        activeDayKey = day || null;
        finDetailLineKey = null;
        render();
      };
    });
  }

  // ---- Rendu ------------------------------------------------------------- //
  function render() {
    document.getElementById("results").classList.remove("hidden");
    var vs = getViewState();
    if (!vs) return;
    var sm = vs.summary;
    renderDayTabs();
    var counts = recomputeCounts();

    var metrics = [
      ["Transactions POS", sm.pos_transactions],
      ["Total POS (DH)", Math.round(sm.pos_total).toLocaleString("fr-FR")],
      ["Anomalies", counts.n_anomalies],
      ["🔴 Haute", counts.severity.haute || 0],
      ["🟠 Moyenne", counts.severity.moyenne || 0],
      ["🔵 Infos", counts.n_infos || 0],
    ];
    if (counts.n_validated) {
      metrics.push(["✅ Validées (hors calcul)", counts.n_validated]);
    }
    if (counts.n_naps_pairing_ok) {
      metrics.push(["✅ TPE totaux OK (appariement)", counts.n_naps_pairing_ok]);
    }
    var adjFin = getAdjustedFinancial();
    if (adjFin && adjFin.cash_to_collect && adjFin.cash_to_collect.net_to_collect > 0) {
      metrics.push(["💰 Cash à collecter", "+" +
        adjFin.cash_to_collect.net_to_collect.toLocaleString("fr-FR") + " DH"]);
    }
    document.getElementById("metrics").innerHTML = metrics.map(function (m) {
      return '<div class="metric"><div class="label">' + m[0] +
             '</div><div class="value">' + m[1] + "</div></div>";
    }).join("");

    var period = document.getElementById("period");
    var txt;
    if (activeDayKey) {
      txt = "📅 Journée analysée : <b>" + escapeHtml(fmtDayLabel(activeDayKey)) + "</b> " +
            "(vue isolée — Glovo / Site / NAPS filtrés sur cette date POS).";
    } else {
      txt = "📅 Période analysée (d'après le POS) : <b>" +
            escapeHtml(sm.pos_date_min) + "</b> → <b>" + escapeHtml(sm.pos_date_max) + "</b>.";
      if (state.byDay && Object.keys(state.byDay).length > 1) {
        txt += " Utilisez les onglets ci-dessous pour analyser <b>jour par jour</b>.";
      }
    }
    var ex = [];
    if (!activeDayKey) {
      if (sm.glovo_excluded) ex.push(sm.glovo_excluded + " commande(s) Glovo");
      if (sm.site_excluded) ex.push(sm.site_excluded + " commande(s) Site");
    }
    if (ex.length) txt += " " + ex.join(" et ") + " hors de cette période ont été ignorée(s).";
    period.innerHTML = txt;

    var ch = sm.channels;
    var keys = Object.keys(ch);
    var max = Math.max.apply(null, keys.map(function (k) { return ch[k]; })) || 1;
    document.getElementById("chart").innerHTML = keys.map(function (k) {
      var h = Math.round((ch[k] / max) * 100);
      return '<div class="bar"><div class="bar-val">' + ch[k] +
             '</div><div class="fill" style="height:' + h + '%"></div>' +
             '<div class="bar-label">' + escapeHtml(k) + "</div></div>";
    }).join("");

    renderPaymentBreakdown(vs.summary.financial);
    renderCashToCollect(getAdjustedFinancial());
    renderFinancial(getAdjustedFinancial());
    renderNapsBalanced();
    renderSpempReview();

    var sources = uniq(activeAnomalies().map(function (a) { return a.source; }));
    var wrap = document.getElementById("fsrc-wrap");
    wrap.querySelectorAll("label").forEach(function (l) { l.remove(); });
    sources.forEach(function (src) {
      var lab = document.createElement("label");
      lab.innerHTML = '<input type="checkbox" class="fsrc" value="' + escapeHtml(src) +
                      '" checked /> ' + escapeHtml(src);
      wrap.appendChild(lab);
    });

    document.querySelectorAll(".fsev, .fsrc").forEach(function (cb) {
      cb.onchange = function () { renderAnomalies(); renderInfos(); };
    });
    document.getElementById("only-anom").onchange = renderPos;

    renderAnomalies();
    renderValidated();
    renderInfos();
    renderPos();
  }

  function fmtDH(v) {
    if (v == null || isNaN(v)) return "";
    return Math.round(v).toLocaleString("fr-FR") + " DH";
  }

  function renderPaymentBreakdown(fin) {
    var el = document.getElementById("payment-breakdown");
    if (!fin || !fin.payment_breakdown) {
      el.innerHTML = "<p class='muted'>Aucune donnée POS.</p>";
      return;
    }
    el.innerHTML = fin.payment_breakdown.payments.map(function (pay) {
      var splits = pay.splits.map(function (sp) {
        var cls = Math.round(sp.amount) === 0 ? " pay-split zero" : " pay-split";
        return "<div class='" + cls + "'><span class='split-label'>" +
          escapeHtml(sp.label) + "</span><span class='split-amt'>" +
          fmtDH(sp.amount) + "</span></div>";
      }).join("");
      return '<div class="pay-card">' +
        '<div class="pay-card-head">' +
        "<span class='pay-name'>" + escapeHtml(pay.icon) + " " + escapeHtml(pay.label) + "</span>" +
        "<span class='pay-total'>" + fmtDH(pay.total) + "</span>" +
        "</div>" + splits + "</div>";
    }).join("");
  }

  function renderCashToCollect(fin) {
    var section = document.getElementById("cash-collect-section");
    var content = document.getElementById("cash-collect-content");
    if (!fin || !fin.cash_to_collect) {
      section.classList.add("hidden");
      content.innerHTML = "";
      return;
    }
    var cc = fin.cash_to_collect;
    section.classList.remove("hidden");

    var netCls = cc.net_to_collect > 0 ? "cc-highlight" :
      (cc.net_to_collect === 0 ? "cc-ok" : "");
    var netSign = cc.net_to_collect > 0 ? "+" : "";

    var collectWhoInline = "";
    if (cc.by_user && cc.by_user.length && cc.net_to_collect > 0) {
      var toCollectUsers = cc.by_user.filter(function (u) { return u.net_to_collect > 0.5; });
      if (toCollectUsers.length) {
        collectWhoInline = '<div class="cc-who-inline">' +
          toCollectUsers.map(function (u) {
            return "<span><b>" + escapeHtml(u.user) + "</b> +" + fmtDH(u.net_to_collect) + "</span>";
          }).join("") + "</div>";
      }
    }

    var html = '<div class="cash-collect-summary">' +
      '<div class="cc-item"><div class="cc-label">Cash saisi au POS (total)</div>' +
      '<div class="cc-value">' + fmtDH(cc.pos_cash_recorded) + "</div></div>" +
      '<div class="cc-item cc-item-collect"><div class="cc-label">À collecter des caissiers</div>' +
      '<div class="cc-value ' + netCls + '">' + netSign + fmtDH(cc.net_to_collect) + "</div>" +
      collectWhoInline + "</div>" +
      '<div class="cc-item"><div class="cc-label">Cash réel attendu en caisse</div>' +
      '<div class="cc-value cc-highlight">' + fmtDH(cc.cash_expected_physical) + "</div></div>" +
      "</div>";

    if (cc.net_to_collect <= 0 && cc.total_to_collect <= 0) {
      html += "<p class='muted' style='margin:8px 0 0'>✅ Rien à collecter en plus du POS — " +
        "Cash / Glovo Cash / Site emporter / TPE alignés.</p>";
    } else if (cc.net_to_collect > 0) {
      html += "<p class='muted' style='margin:8px 0 0'>" +
        "Vous devez récupérer <b>" + fmtDH(cc.net_to_collect) + "</b> en plus de ce qui est " +
        "affiché au POS Cash (<b>" + fmtDH(cc.pos_cash_recorded) + "</b>) → " +
        "cash physique attendu : <b>" + fmtDH(cc.cash_expected_physical) + "</b>.</p>";
    }

    if (cc.by_user && cc.by_user.length && cc.net_to_collect > 0) {
      var toCollectUsers = cc.by_user.filter(function (u) { return u.net_to_collect > 0.5; });
      if (toCollectUsers.length) {
        html += '<div class="cc-collect-who"><b>À récupérer par utilisateur :</b> ' +
          toCollectUsers.map(function (u) {
            return "<span class='cc-who-chip'><b>" + escapeHtml(u.user) + "</b> +" +
              fmtDH(u.net_to_collect) + "</span>";
          }).join("") + "</div>";
      }
    }

    if (cc.by_user && cc.by_user.length) {
      var withNetEarly = cc.by_user.filter(function (u) {
        return Math.abs(u.net_to_collect) >= 0.5 || u.to_collect >= 0.5 || u.over_recorded >= 0.5;
      });
      if (withNetEarly.length && cc.net_to_collect > 0) {
        html += "<h3 class='cc-user-title'>Qui doit rendre quoi (colonne F POS)</h3>";
        html += '<div class="table-wrap"><table class="cc-user-table"><thead><tr>' +
          "<th>Utilisateur</th><th>À collecter</th><th>Sur-saisie</th><th>Net</th>" +
          "<th>Glovo Cash</th><th>Site emporter</th><th>TPE CB</th></tr></thead><tbody>";
        withNetEarly.forEach(function (u) {
          var bl = u.by_line || {};
          var netCls = u.net_to_collect > 0 ? "cc-ecart-pos" :
            (u.net_to_collect < 0 ? "cc-ecart-neg" : "");
          html += "<tr><td><b>" + escapeHtml(u.user) + "</b></td>" +
            "<td>" + fmtDH(u.to_collect) + "</td>" +
            "<td>" + (u.over_recorded > 0 ? fmtDH(u.over_recorded) : "—") + "</td>" +
            "<td class='" + netCls + "'><b>" +
            (u.net_to_collect > 0 ? "+" + fmtDH(u.net_to_collect) : fmtDH(u.net_to_collect)) +
            "</b></td>" +
            "<td>" + fmtDH(bl.glovo_cash || 0) + "</td>" +
            "<td>" + fmtDH(bl.site_cash || 0) + "</td>" +
            "<td>" + fmtDH(bl.naps_tpe_over || 0) + "</td></tr>";
        });
        html += "</tbody></table></div>";
      }
    }

    var detail = cc.items.filter(function (it) {
      return Math.abs(it.collect_amount || 0) >= 0.5;
    });
    if (detail.length) {
      html += '<div class="cash-collect-items">' + detail.map(function (it) {
        var ca = it.collect_amount || 0;
        var ecartCls = ca > 0 ? "cc-ecart-pos" : "cc-ecart-neg";
        var collectLbl = ca > 0 ? "À collecter : +" + fmtDH(ca) :
          (ca < 0 ? "Sur-saisie POS : " + fmtDH(ca) : "");
        return '<div class="cash-collect-item"><div class="cc-row"><b>' +
          escapeHtml(it.label) + "</b><span>" + escapeHtml(it.pos_label) + " : <b>" +
          fmtDH(it.pos) + "</b></span><span>" + escapeHtml(it.src_label) + " : <b>" +
          fmtDH(it.src) + "</b></span>" +
          (collectLbl ? "<span class='" + ecartCls + "'>" + collectLbl + "</span>" : "") +
          "</div>" +
          (it.hint ? "<div class='cc-hint'>" + escapeHtml(it.hint) + "</div>" : "") +
          "</div>";
      }).join("") + "</div>";
    }

    var unattributed = (cc.ticket_contributions || []).filter(function (t) {
      return Math.abs(t.amount) >= 0.5 && isUnattributedCashContrib(t);
    });
  var posUsers = listViewPosUsers();
    if (unattributed.length) {
      var pending = unattributed.filter(function (t) { return !t.manually_assigned; });
      var manual = unattributed.filter(function (t) { return t.manually_assigned; });

      if (pending.length) {
        html += "<h3 class='cc-user-title'>À clarifier — Non attribué</h3>";
        html += "<p class='muted' style='margin:0 0 8px'>Affectez à un utilisateur POS " +
          "une fois la preuve trouvée (colonne F).</p>";
        html += '<div class="table-wrap"><table class="cc-user-table"><thead><tr>' +
          "<th>Date/heure</th><th>Source</th><th>Montant</th><th>Détail</th>" +
          "<th>Affecter à</th></tr></thead><tbody>";
        pending.forEach(function (t) {
          var lineLbl = t.lineKey === "glovo_cash" ? "Glovo Cash" :
            (t.lineKey === "site_cash" ? "Site emporter" :
              (t.lineKey === "naps_tpe_over" ? "TPE CB" : t.lineKey));
          var opts = posUsers.map(function (u) {
            return "<option value='" + escapeAttr(u.user) + "'>" + escapeHtml(u.label) + "</option>";
          }).join("");
          html += "<tr><td>" + escapeHtml(t.when || "") + "</td>" +
            "<td>" + escapeHtml(lineLbl) + "</td>" +
            "<td class='cc-ecart-pos'><b>" + fmtDH(t.amount) + "</b></td>" +
            "<td class='muted' style='font-size:.82rem'>" + escapeHtml(t.detail || "") + "</td>" +
            "<td class='cc-assign-cell'>" +
            (posUsers.length ?
              "<select class='cc-user-select' data-cid='" + escapeAttr(t.id) + "'>" +
              opts + "</select>" +
              "<button type='button' class='btn-assign-cash' data-cid='" +
              escapeAttr(t.id) + "'>Affecter</button>" :
              "<span class='muted'>Aucun User dans le POS</span>") +
            "</td></tr>";
        });
        html += "</tbody></table></div>";
      }

      if (manual.length) {
        html += "<h3 class='cc-user-title'>Réaffectations manuelles</h3>";
        html += '<div class="table-wrap"><table class="cc-user-table"><thead><tr>' +
          "<th>Utilisateur</th><th>Date/heure</th><th>Source</th><th>Montant</th>" +
          "<th>Détail</th><th></th></tr></thead><tbody>";
        manual.forEach(function (t) {
          var lineLbl = t.lineKey === "glovo_cash" ? "Glovo Cash" :
            (t.lineKey === "site_cash" ? "Site emporter" :
              (t.lineKey === "naps_tpe_over" ? "TPE CB" : t.lineKey));
          html += "<tr class='row-manual-cash'><td><b>" + escapeHtml(t.user) + "</b> " +
            "<span class='cc-manual-badge'>Manuel</span></td>" +
            "<td>" + escapeHtml(t.when || "") + "</td>" +
            "<td>" + escapeHtml(lineLbl) + "</td>" +
            "<td class='cc-ecart-pos'><b>" + fmtDH(t.amount) + "</b></td>" +
            "<td class='muted' style='font-size:.82rem'>" + escapeHtml(t.detail || "") + "</td>" +
            "<td><button type='button' class='btn-unassign-cash' data-cid='" +
            escapeAttr(t.id) + "'>Annuler</button></td></tr>";
        });
        html += "</tbody></table></div>";
      }
    }

    if (cc.by_user && cc.by_user.length) {
      var withNet = cc.by_user.filter(function (u) {
        return Math.abs(u.net_to_collect) >= 0.5 || u.to_collect >= 0.5 || u.over_recorded >= 0.5;
      });
      if (withNet.length && cc.net_to_collect <= 0) {
        html += "<h3 class='cc-user-title'>Par utilisateur (colonne F du POS)</h3>";
        html += '<div class="table-wrap"><table class="cc-user-table"><thead><tr>' +
          "<th>Utilisateur</th><th>À collecter</th><th>Sur-saisie</th><th>Net</th>" +
          "<th>Glovo Cash</th><th>Site emporter</th><th>TPE CB</th></tr></thead><tbody>";
        withNet.forEach(function (u) {
          var bl = u.by_line || {};
          var netCls = u.net_to_collect > 0 ? "cc-ecart-pos" :
            (u.net_to_collect < 0 ? "cc-ecart-neg" : "");
          html += "<tr><td><b>" + escapeHtml(u.user) + "</b></td>" +
            "<td>" + fmtDH(u.to_collect) + "</td>" +
            "<td>" + (u.over_recorded > 0 ? fmtDH(u.over_recorded) : "—") + "</td>" +
            "<td class='" + netCls + "'><b>" +
            (u.net_to_collect > 0 ? "+" + fmtDH(u.net_to_collect) : fmtDH(u.net_to_collect)) +
            "</b></td>" +
            "<td>" + fmtDH(bl.glovo_cash || 0) + "</td>" +
            "<td>" + fmtDH(bl.site_cash || 0) + "</td>" +
            "<td>" + fmtDH(bl.naps_tpe_over || 0) + "</td></tr>";
        });
        html += "</tbody></table></div>";
      }
    }

    if (cc.ticket_contributions && cc.ticket_contributions.length) {
      var ticketRows = cc.ticket_contributions.filter(function (t) {
        return Math.abs(t.amount) >= 0.5;
      });
      if (ticketRows.length) {
        html += "<h3 class='cc-user-title'>Détail par ticket</h3>";
        html += '<div class="table-wrap"><table class="cc-user-table"><thead><tr>' +
          "<th>Utilisateur</th><th>Date/heure</th><th>Ticket</th><th>Source</th>" +
          "<th>Montant (DH)</th><th>Détail</th></tr></thead><tbody>";
        ticketRows.forEach(function (t) {
          var amtCls = t.amount > 0 ? "cc-ecart-pos" : "cc-ecart-neg";
          var lineLbl = t.lineKey === "glovo_cash" ? "Glovo Cash" :
            (t.lineKey === "site_cash" ? "Site emporter" :
              (t.lineKey === "naps_tpe_over" ? "TPE CB" : t.lineKey));
          html += "<tr><td>" + escapeHtml(t.user || "Non attribué") +
            (t.manually_assigned ? " <span class='cc-manual-badge'>Manuel</span>" : "") +
            "</td>" +
            "<td>" + escapeHtml(t.when || "") + "</td>" +
            "<td>" + escapeHtml(t.ticket_name || t.ticket_no || "—") + "</td>" +
            "<td>" + escapeHtml(lineLbl) + "</td>" +
            "<td class='" + amtCls + "'><b>" + fmtDH(t.amount) + "</b></td>" +
            "<td class='muted' style='font-size:.82rem'>" + escapeHtml(t.detail || "") +
            "</td></tr>";
        });
        html += "</tbody></table></div>";
      }
    }

    content.innerHTML = html;

    content.querySelectorAll(".btn-assign-cash").forEach(function (btn) {
      btn.onclick = function () {
        var row = btn.closest("tr");
        var sel = row ? row.querySelector("select.cc-user-select") : null;
        if (!sel || !sel.value) return;
        assignCashContribution(btn.getAttribute("data-cid"), sel.value);
      };
    });
    content.querySelectorAll(".btn-unassign-cash").forEach(function (btn) {
      btn.onclick = function () {
        unassignCashContribution(btn.getAttribute("data-cid"));
      };
    });
  }

  function renderFinancial(fin) {
    if (!fin) {
      document.getElementById("fin-lines").innerHTML = "";
      document.getElementById("fin-ecart-detail").classList.add("hidden");
      return;
    }
    var prev = document.getElementById("fin-adj-note");
    if (prev) prev.remove();
    if (fin.adjustments_applied) {
      var el = document.createElement("p");
      el.id = "fin-adj-note";
      el.className = "muted fin-adj-note";
      el.innerHTML = "Montants ajustés : les anomalies <b>validées</b> ne sont plus comptées dans les écarts.";
      document.getElementById("fin-lines").parentNode.insertBefore(el, document.getElementById("fin-lines"));
    }
    var head = "<thead><tr><th>Source</th><th>Côté POS</th><th>Côté source</th>" +
               "<th>Écart (source − POS)</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + fin.lines.map(function (l) {
      var cls = Math.abs(l.ecart) < 0.5 ? "st-ok" : "st-anom";
      var sign = l.ecart > 0 ? "+" : "";
      var note = l.note ? "<div class='muted' style='font-size:.82rem'>" + escapeHtml(l.note) + "</div>" : "";
      if (l.lineKey === "naps" && Math.abs(l.ecart) < 0.5) {
        var vs = getViewState();
        if (vs && vs.naps_balanced && vs.naps_balanced.length) {
          note += "<div class='muted' style='font-size:.82rem;margin-top:4px'>" +
            "✅ <b>Totaux alignés</b> — écarts d'appariement seulement (voir section ci-dessous).</div>";
        }
      }
      var rowCls = l.isTotal ? "fin-total" : (l.group === "glovo" ? "fin-glovo-sub" : "");
      var contribs = l.lineKey ? CNS.getFinancialContributors(l.lineKey, activeAnomalies()) : [];
      var detailCell = "";
      if (l.lineKey) {
        var active = finDetailLineKey === l.lineKey;
        detailCell = "<button type='button' class='btn-fin-detail" +
          (active ? " active" : "") + "' data-line-key='" + escapeHtml(l.lineKey) +
          "' data-line-label='" + escapeHtml(l.source) + "'>🔍 Voir l'écart (" +
          contribs.length + ")</button>";
      }
      return "<tr class='" + rowCls + "'><td><b>" + escapeHtml(l.source) + "</b></td>" +
             "<td>" + escapeHtml(l.pos_label) + " : <b>" + fmtDH(l.pos) + "</b></td>" +
             "<td>" + escapeHtml(l.src_label) + " : <b>" + fmtDH(l.src) + "</b>" + note + "</td>" +
             "<td class='" + cls + "'><b>" + sign + fmtDH(l.ecart) + "</b></td>" +
             "<td>" + detailCell + "</td></tr>";
    }).join("") + "</tbody>";
    var table = document.getElementById("fin-lines");
    table.innerHTML = head + body;
    table.querySelectorAll(".btn-fin-detail").forEach(function (btn) {
      btn.onclick = function () {
        var key = btn.getAttribute("data-line-key");
        finDetailLineKey = finDetailLineKey === key ? null : key;
        renderFinEcartDetail(fin);
        renderFinancial(fin);
      };
    });
    renderFinEcartDetail(fin);

    var pays = fin.pays;
    var chans = Object.keys(fin.matrix);
    var colTot = {}; pays.forEach(function (p) { colTot[p] = 0; }); var grand = 0;
    var mhead = "<thead><tr><th>Canal</th>" + pays.map(function (p) {
      return "<th>" + escapeHtml(p) + "</th>"; }).join("") + "<th>Total</th></tr></thead>";
    var mbody = "<tbody>" + chans.map(function (ch) {
      var row = fin.matrix[ch], rt = 0;
      var cells = pays.map(function (p) {
        var v = row[p] || 0; rt += v; colTot[p] += v;
        return "<td>" + (v ? fmtDH(v) : "—") + "</td>";
      }).join("");
      grand += rt;
      return "<tr><td>" + escapeHtml(ch) + "</td>" + cells +
             "<td><b>" + fmtDH(rt) + "</b></td></tr>";
    }).join("");
    mbody += "<tr><td><b>Total</b></td>" + pays.map(function (p) {
      return "<td><b>" + fmtDH(colTot[p]) + "</b></td>"; }).join("") +
      "<td><b>" + fmtDH(grand) + "</b></td></tr></tbody>";
    document.getElementById("fin-matrix").innerHTML = mhead + mbody;
  }

  function renderSpempReview() {
    var section = document.getElementById("spemp-review-section");
    var content = document.getElementById("spemp-review-content");
    var vs = getViewState();
    var rows = vs && vs.spemp_review ? vs.spemp_review : [];
    if (!rows.length) {
      section.classList.add("hidden");
      content.innerHTML = "";
      return;
    }
    section.classList.remove("hidden");
    var total = rows.reduce(function (s, r) {
      return s + (r.total == null || isNaN(r.total) ? 0 : r.total);
    }, 0);
    var html = "<p class='muted' style='margin:0 0 10px'><b>" + rows.length +
      "</b> ticket(s) · <b>" + Math.round(total).toLocaleString("fr-FR") + " DH</b></p>";
    html += '<div class="table-wrap"><table><thead><tr>' +
      "<th>Date/heure</th><th>Ticket name</th><th>N° POS</th><th>Total</th>" +
      "<th>Paiement</th><th>Type</th></tr></thead><tbody>";
    rows.forEach(function (r) {
      var kind = r.kind === "a_rattacher" ?
        "<span style='font-size:.78rem;padding:2px 8px;border-radius:4px;background:#fff3e6;" +
        "border:1px solid #e8d4b8;color:#8b5a2b'>À rattacher</span>" :
        (r.kind === "bipeur" ?
          "<span style='font-size:.78rem;padding:2px 8px;border-radius:4px;background:#e8f6ef;" +
          "border:1px solid #a8d5c2;color:#1e6b45'>Bipeur SP&EMP</span>" :
          (r.kind === "cash_comptoir" ?
            "<span style='font-size:.78rem;padding:2px 8px;border-radius:4px;background:#e8f0fa;" +
            "border:1px solid #b8cfe8;color:#2a5080'>Cash comptoir SP&EMP</span>" :
            (r.kind === "cb_naps" ?
              "<span style='font-size:.78rem;padding:2px 8px;border-radius:4px;background:#e8f6ef;" +
              "border:1px solid #a8d5c2;color:#1e6b45'>CB comptoir (NAPS OK)</span>" :
              "<span class='muted'>Libellé libre → SP&EMP</span>")));
      html += "<tr><td>" + escapeHtml(r.when || "") + "</td><td>" +
        escapeHtml(r.ticket_name) + "</td><td>" + escapeHtml(r.ticket_no) + "</td><td><b>" +
        (r.total == null || isNaN(r.total) ? "—" : fmtDH(r.total)) + "</b></td><td>" +
        escapeHtml(r.payment_type) + "</td><td>" + kind + "</td></tr>";
    });
    html += "</tbody></table></div>";
    content.innerHTML = html;
  }

  function renderNapsBalanced() {
    var section = document.getElementById("naps-balanced-section");
    var content = document.getElementById("naps-balanced-content");
    var vs = getViewState();
    var groups = vs && vs.naps_balanced ? vs.naps_balanced : [];
    if (!groups.length) {
      section.classList.add("hidden");
      content.innerHTML = "";
      return;
    }
    section.classList.remove("hidden");
    content.innerHTML = groups.map(function (g) {
      var dayLabel = fmtDayLabel(g.date);
      var posList = g.pos_items.map(function (it) {
        var lbl = it.ticket_name ? escapeHtml(it.ticket_name) : "ticket";
        if (it.ticket_no) lbl += " #" + escapeHtml(it.ticket_no);
        if (it.when) lbl += " · " + escapeHtml(String(it.when).match(/\d{2}:\d{2}/) ?
          String(it.when).match(/\d{2}:\d{2}/)[0] : "");
        return "<li><b>" + Math.round(it.amount) + " DH</b> — " + lbl + "</li>";
      }).join("");
      var napsList = g.naps_items.map(function (it) {
        return "<li><b>" + Math.round(it.amount) + " DH</b> — ligne relevé " +
          escapeHtml(it.row === "" || it.row == null ? "?" : it.row) + "</li>";
      }).join("");
      return '<div class="naps-balanced-day">' +
        "<h3>📅 " + escapeHtml(dayLabel) + " — <span class='match'>✅ pas d'écart</span></h3>" +
        '<div class="naps-balanced-totals">' +
        '<span class="tot">POS Credit card : <b>' + g.pos_cc_total + " DH</b></span>" +
        '<span class="tot">Relevé NAPS : <b>' + g.naps_total + " DH</b></span>" +
        '<span class="tot match">Écart financier : <b>0 DH</b></span>' +
        "</div>" +
        "<p class='muted' style='font-size:.88rem;margin:0 0 10px'>" +
        "Montants non appariés transaction par transaction : " +
        "<b>" + g.pos_unmatched_total + " DH</b> au POS ↔ " +
        "<b>" + g.naps_unmatched_total + " DH</b> sur le TPE — " +
        "les totaux se compensent.</p>" +
        '<div class="naps-balanced-cols">' +
        "<div><h4>Côté POS (sans ligne TPE au même montant)</h4><ul>" +
        (posList || "<li>—</li>") + "</ul></div>" +
        "<div><h4>Côté TPE (sans ticket POS au même montant)</h4><ul>" +
        (napsList || "<li>—</li>") + "</ul></div>" +
        "</div></div>";
    }).join("");
  }

  function renderFinEcartDetail(fin) {
    var panel = document.getElementById("fin-ecart-detail");
    if (!finDetailLineKey || !fin) {
      panel.classList.add("hidden");
      panel.innerHTML = "";
      return;
    }
    var line = fin.lines.filter(function (l) { return l.lineKey === finDetailLineKey; })[0];
    if (!line) {
      panel.classList.add("hidden");
      return;
    }
    var contribs = CNS.getFinancialContributors(finDetailLineKey, activeAnomalies())
      .sort(function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; });
    var sumContrib = CNS.sumFinancialContributions(finDetailLineKey, contribs);
    var html = "<h4>🔍 Écart « " + escapeHtml(line.source) + " » : " +
               (line.ecart > 0 ? "+" : "") + fmtDH(line.ecart) + "</h4>";
    html += "<p class='muted'>Chaque ligne indique comment l'anomalie pousse l'écart " +
            "(source − POS). <b>+</b> = la source encaisse plus que le POS · " +
            "<b>−</b> = le POS a plus que la source.</p>";
    if (!contribs.length) {
      html += "<p>Aucune anomalie en attente explique cet écart (écart résidu ou déjà validé).</p>";
    } else {
      var sumSign = sumContrib > 0 ? "+" : "";
      html += "<p><b>Total expliqué par les anomalies listées : " + sumSign +
              fmtDH(sumContrib) + "</b></p>";
      if (Math.abs(sumContrib - line.ecart) >= 1) {
        html += "<p class='muted'>Le reste (" + fmtDH(line.ecart - sumContrib) +
                ") peut venir de commandes sans anomalie individuelle (écarts de regroupement).</p>";
      }
      html += "<div class='table-wrap'><table class='fin-contrib-table'><thead><tr>" +
              "<th>Valider</th><th>Gravité</th><th>Type</th><th>Impact écart</th>" +
              "<th>Ticket</th><th>Détail</th></tr></thead><tbody>";
      contribs.forEach(function (a) {
        var impact = CNS.ecartContributionForAnomaly(a, finDetailLineKey);
        var impSign = impact > 0 ? "+" : "";
        html += "<tr><td><button type='button' class='btn-validate' data-id='" +
                escapeHtml(a.id) + "'>✅</button></td>" +
                "<td><span class='sev-badge sev-" + a.severity + "'>" +
                SEV_BADGE[a.severity] + "</span></td>" +
                "<td>" + escapeHtml(a.type) + "</td>" +
                "<td class='st-anom'><b>" + impSign + fmtDH(impact) + "</b></td>" +
                "<td>" + escapeHtml(fmtTicketLabel(a)) + "</td>" +
                "<td>" + escapeHtml(a.detail) + "</td></tr>";
      });
      html += "</tbody></table></div>";
    }
    panel.innerHTML = html;
    panel.classList.remove("hidden");
    bindValidateButtons(panel);
  }

  function fmtAnomalyMontant(a) {
    var parts = [];
    if (a.amount_pos != null && !isNaN(a.amount_pos)) {
      parts.push("<b>" + fmtDH(a.amount_pos) + "</b> POS");
    }
    if (a.amount_source != null && !isNaN(a.amount_source)) {
      parts.push("<b>" + fmtDH(a.amount_source) + "</b> source");
    }
    if (!parts.length) return "—";
    return parts.join(" · ");
  }

  function _anomalyRowHTML(a, withValidate, napsOkBadge) {
    var rowCls = napsOkBadge && isNapsTotalsOk(a) ? " row-naps-ok" : "";
    var badge = napsOkBadge && isNapsTotalsOk(a) ?
      " <span class='badge-naps-ok'>✅ Totaux OK</span>" : "";
    var btn = withValidate ?
      '<td><button type="button" class="btn-validate" data-id="' + escapeHtml(a.id) +
      '" title="Valider — hors calcul">✅ Valider</button></td>' : "";
    return "<tr class='" + rowCls + "'>" + btn +
      "<td><span class='sev-badge sev-" + a.severity + "'>" +
      SEV_BADGE[a.severity] + "</span></td><td>" + escapeHtml(a.source) +
      "</td><td>" + escapeHtml(a.type) + badge + "</td><td>" + escapeHtml(a.file || "") +
      "</td><td>" + escapeHtml(a.row === "" || a.row == null ? "" : a.row) +
      "</td><td>" + escapeHtml(a.when || "") + "</td><td>" +
      escapeHtml(fmtTicketLabel(a)) + "</td><td>" + fmtAnomalyMontant(a) + "</td><td>" +
      escapeHtml(a.detail) + "</td></tr>";
  }

  function _tableHTML(rows, withValidate, napsOkBadge) {
    var head = "<thead><tr>";
    if (withValidate) head += "<th>Valider</th>";
    head += "<th>Gravité</th><th>Source</th><th>Type</th>" +
            "<th>Fichier</th><th>Ligne</th><th>Date/heure</th>" +
            "<th>Ticket</th><th>Montant</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + rows.map(function (a) {
      return _anomalyRowHTML(a, withValidate, napsOkBadge);
    }).join("") + "</tbody>";
    return head + body;
  }

  function bindValidateButtons(container) {
    container.querySelectorAll(".btn-validate").forEach(function (btn) {
      btn.onclick = function () {
        validateAnomaly(btn.getAttribute("data-id"));
      };
    });
  }

  function renderAnomalies() {
    var sev = checkedValues("fsev"), src = checkedValues("fsrc");
    var all = activeAnomalies().filter(function (a) {
      return a.severity !== "info" &&
             sev.indexOf(a.severity) !== -1 && src.indexOf(a.source) !== -1;
    }).sort(function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; });

    var actionable = all.filter(isActionableAnomaly);
    var pairingOk = all.filter(isNapsTotalsOk);

    var countTxt = actionable.length + " anomalie(s) à traiter";
    if (pairingOk.length) {
      countTxt += " · " + pairingOk.length + " ligne(s) TPE (totaux alignés — pas d'écart)";
    }
    document.getElementById("anom-count").textContent = countTxt;

    var table = document.getElementById("anom-table");
    if (!all.length) {
      table.innerHTML =
        "<tbody><tr><td colspan='10'>✅ Aucune anomalie en attente pour ces filtres.</td></tr></tbody>";
      return;
    }
    var head = "<thead><tr><th>Valider</th><th>Gravité</th><th>Source</th><th>Type</th>" +
      "<th>Fichier</th><th>Ligne</th><th>Date/heure</th><th>Ticket</th><th>Montant</th><th>Détail</th></tr></thead>";
    var body = "<tbody>";
    if (!actionable.length && pairingOk.length) {
      body += "<tr><td colspan='10' class='muted' style='background:#f4fbf7'>" +
        "✅ Aucune anomalie financière — seulement des écarts d'appariement TPE (totaux OK).</td></tr>";
    }
    actionable.forEach(function (a) { body += _anomalyRowHTML(a, true, false); });
    if (pairingOk.length) {
      body += "<tr><td colspan='10' class='muted' style='background:#f4fbf7;font-weight:600'>" +
        "✅ Appariement TPE — totaux POS CB = NAPS (détail dans la section verte ci-dessus)</td></tr>";
      pairingOk.forEach(function (a) { body += _anomalyRowHTML(a, true, true); });
    }
    body += "</tbody>";
    table.innerHTML = head + body;
    bindValidateButtons(table);
  }

  function renderValidated() {
    var section = document.getElementById("validated-section");
    var rows = validatedAnomalies().filter(function (a) { return a.severity !== "info"; })
      .sort(function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; });
    if (!rows.length) {
      section.classList.add("hidden");
      return;
    }
    section.classList.remove("hidden");
    document.getElementById("validated-count").textContent =
      rows.length + " anomalie(s) validée(s) — exclues des calculs";
    var head = "<thead><tr><th>Action</th><th>Gravité</th><th>Source</th><th>Type</th>" +
               "<th>Ticket</th><th>Montant</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + rows.map(function (a) {
      return "<tr class='row-validated'><td><button type='button' class='btn-unvalidate' " +
             "data-id='" + escapeHtml(a.id) + "' title='Annuler la validation'>↩ Annuler</button></td>" +
             "<td><span class='sev-badge sev-" + a.severity + "'>" + SEV_BADGE[a.severity] +
             "</span></td><td>" + escapeHtml(a.source) + "</td><td>" + escapeHtml(a.type) +
             "</td><td>" + escapeHtml(fmtTicketLabel(a)) + "</td><td>" +
             fmtAnomalyMontant(a) + "</td><td>" + escapeHtml(a.detail) + "</td></tr>";
    }).join("") + "</tbody>";
    var table = document.getElementById("validated-table");
    table.innerHTML = head + body;
    table.querySelectorAll(".btn-unvalidate").forEach(function (btn) {
      btn.onclick = function () { unvalidateAnomaly(btn.getAttribute("data-id")); };
    });
  }

  function renderInfos() {
    var src = checkedValues("fsrc");
    var rows = activeAnomalies().filter(function (a) {
      return a.severity === "info" && src.indexOf(a.source) !== -1;
    });
    document.getElementById("infos-count").textContent = rows.length + " info(s) affichée(s)";
    document.getElementById("infos-table").innerHTML = rows.length ? _tableHTML(rows, false) :
      "<tbody><tr><td>Aucune info pour ces filtres.</td></tr></tbody>";
  }

  function renderPos() {
    var vs = getViewState();
    if (!vs) return;
    var onlyAnom = document.getElementById("only-anom").checked;
    var rows = vs.pos.filter(function (p) {
      return !onlyAnom || (p.statut && p.statut.indexOf("Anomalie") !== -1);
    });
    var cols = [
      ["ticket_no", "Ticket No."], ["date", "Date"], ["hour", "Heure"],
      ["user", "Caissier"], ["ticket_name", "Ticket name"], ["channel", "Canal"],
      ["total", "Total"], ["payment_type", "Paiement"], ["statut", "Statut"],
      ["anomalies", "Anomalies"],
    ];
    var head = "<thead><tr>" + cols.map(function (c) {
      return "<th>" + c[1] + "</th>"; }).join("") + "</tr></thead>";
    var body = "<tbody>" + rows.map(function (p) {
      var isAnom = p.statut && p.statut.indexOf("Anomalie") !== -1;
      return "<tr class='" + (isAnom ? "row-anom" : "") + "'>" + cols.map(function (c) {
        var v = p[c[0]];
        if (c[0] === "total") v = isNaN(v) ? "" : v;
        if (c[0] === "statut") {
          return "<td class='" + (isAnom ? "st-anom" : "st-ok") + "'>" + escapeHtml(v) + "</td>";
        }
        return "<td>" + escapeHtml(v == null ? "" : String(v)) + "</td>";
      }).join("") + "</tr>";
    }).join("") + "</tbody>";
    document.getElementById("pos-table").innerHTML = head + body;
  }

  document.getElementById("download").addEventListener("click", function () {
    if (!state) return;
    var sm = state.summary;
    var counts = recomputeCounts();
    var fin = getAdjustedFinancial();
    var wb = XLSX.utils.book_new();

    function buildResumeRows(runSm, runCounts, label) {
      var rows = [
        ["Indicateur", "Valeur"],
        ["Vue", label || "Général"],
        ["Période / journée (POS)", (runSm.pos_date_min || "") + " → " + (runSm.pos_date_max || "")],
        ["Transactions POS", runSm.pos_transactions],
        ["Total POS (DH)", Math.round(runSm.pos_total * 100) / 100],
        ["Commandes Glovo (livrées, période)", runSm.glovo_orders],
        ["Commandes Glovo hors période (ignorées)", runSm.glovo_excluded || 0],
        ["Transactions NAPS", runSm.naps_transactions],
        ["Commandes Site (période)", runSm.site_orders],
        ["Commandes Site hors période (ignorées)", runSm.site_excluded || 0],
        ["", ""],
        ["Anomalies (Haute + Moyenne)", runCounts.n_anomalies],
        ["  dont haute", runCounts.severity.haute || 0],
        ["  dont moyenne", runCounts.severity.moyenne || 0],
        ["Anomalies validées (hors calcul)", runCounts.n_validated || 0],
        ["Infos (rattachements & notes)", runCounts.n_infos || 0],
        ["", ""],
      ];
      Object.keys(runSm.channels).forEach(function (k) {
        rows.push(["POS — " + k, runSm.channels[k]]);
      });
      return rows;
    }

    function countsForRun(runResult) {
      var sev = { haute: 0, moyenne: 0, info: 0 };
      var napsPairing = 0;
      runResult.anomalies.forEach(function (a) {
        if (isValidated(a.id)) return;
        if (isNapsTotalsOk(a) && a.severity !== "info") napsPairing++;
        else sev[a.severity] = (sev[a.severity] || 0) + 1;
      });
      var nVal = runResult.anomalies.filter(function (a) { return isValidated(a.id); }).length;
      return {
        n_anomalies: sev.haute + sev.moyenne,
        n_infos: sev.info,
        severity: sev,
        n_validated: nVal,
        n_naps_pairing_ok: napsPairing,
      };
    }

    function adjustedFinFor(runResult) {
      if (!runResult.summary.financial) return null;
      var validated = runResult.anomalies.filter(function (a) { return isValidated(a.id); });
      var fin;
      if (!validated.length) fin = runResult.summary.financial;
      else {
        fin = CNS.applyFinancialAdjustments(runResult.summary.financial,
          CNS.sumFinancialAdjustments(validated));
      }
      enrichFinCashCollect(fin, runResult);
      return fin;
    }

    function buildFinRows(runFin) {
      if (!runFin) return [];
      var frows = [["Réconciliation financière (Écart = source − POS)"], []];
      if (runFin.adjustments_applied) {
        frows.push(["(Anomalies validées exclues des montants ci-dessous)"]);
        frows.push([]);
      }
      frows.push(["Source", "Côté POS (libellé)", "Montant POS", "Côté source (libellé)", "Montant source", "Écart"]);
      runFin.lines.forEach(function (l) {
        frows.push([l.source, l.pos_label, Math.round(l.pos), l.src_label, Math.round(l.src), Math.round(l.ecart)]);
        if (l.note) frows.push(["", l.note]);
      });
      frows.push([], ["Répartition POS : mode de paiement × canal"], []);
      frows.push(["Canal"].concat(runFin.pays, ["Total"]));
      var colTot = {}; runFin.pays.forEach(function (p) { colTot[p] = 0; });
      Object.keys(runFin.matrix).forEach(function (ch) {
        var row = runFin.matrix[ch], rt = 0;
        var cells = runFin.pays.map(function (p) {
          var v = row[p] || 0; rt += v; colTot[p] += v; return Math.round(v);
        });
        frows.push([ch].concat(cells, [Math.round(rt)]));
      });
      var grand = 0; runFin.pays.forEach(function (p) { grand += colTot[p]; });
      frows.push(["Total"].concat(runFin.pays.map(function (p) {
        return Math.round(colTot[p]);
      }), [Math.round(grand)]));
      return frows;
    }

    function buildPaymentBreakdownRows(runFin) {
      if (!runFin || !runFin.payment_breakdown) return [];
      var rows = [["Mode de paiement", "Montant (DH)"], []];
      runFin.payment_breakdown.payments.forEach(function (pay) {
        rows.push([pay.label, Math.round(pay.total)]);
        pay.splits.forEach(function (sp) {
          rows.push(["  " + sp.label, Math.round(sp.amount)]);
        });
        rows.push([]);
      });
      return rows;
    }

    function buildCashCollectRows(runFin) {
      if (!runFin || !runFin.cash_to_collect) return [];
      var cc = runFin.cash_to_collect;
      var rows = [
        ["Cash à collecter"],
        ["Cash saisi au POS (total)", cc.pos_cash_recorded],
        ["À collecter des caissiers (net)", cc.net_to_collect],
        ["Cash réel attendu en caisse", cc.cash_expected_physical],
        [],
        ["Détail par source", "POS", "Source", "À collecter (DH)"],
      ];
      cc.items.forEach(function (it) {
        var ca = it.collect_amount || 0;
        if (Math.abs(ca) < 0.5) return;
        rows.push([it.label, Math.round(it.pos), Math.round(it.src), Math.round(ca)]);
      });
      return rows;
    }

    function buildCashCollectUserRows(runFin) {
      if (!runFin || !runFin.cash_to_collect || !runFin.cash_to_collect.by_user) return [];
      var cc = runFin.cash_to_collect;
      var rows = [
        ["Cash à collecter — par utilisateur (col. F POS)"],
        ["Utilisateur", "À collecter", "Sur-saisie", "Net", "Glovo Cash", "Site emporter", "TPE CB"],
      ];
      cc.by_user.forEach(function (u) {
        var bl = u.by_line || {};
        rows.push([
          u.user,
          u.to_collect,
          u.over_recorded,
          u.net_to_collect,
          Math.round(bl.glovo_cash || 0),
          Math.round(bl.site_cash || 0),
          Math.round(bl.naps_tpe_over || 0),
        ]);
      });
      return rows;
    }

    function buildCashCollectTicketRows(runFin) {
      if (!runFin || !runFin.cash_to_collect || !runFin.cash_to_collect.ticket_contributions) {
        return [];
      }
      var rows = [
        ["Cash à collecter — détail tickets"],
        ["Utilisateur", "Date/heure", "N° POS", "Ticket name", "Source",
          "Montant (DH)", "Affectation", "Détail"],
      ];
      runFin.cash_to_collect.ticket_contributions.forEach(function (t) {
        if (Math.abs(t.amount) < 0.5) return;
        var lineLbl = t.lineKey === "glovo_cash" ? "Glovo Cash" :
          (t.lineKey === "site_cash" ? "Site emporter" :
            (t.lineKey === "naps_tpe_over" ? "TPE CB > NAPS" : t.lineKey));
        var aff = t.manually_assigned ? "Manuel → " + t.user :
          ((t.original_user || t.user) === CNS.CASH_COLLECT_UNATTRIBUTED ?
            "Non attribué" : "Auto (ticket POS)");
        rows.push([
          t.user || "Non attribué",
          t.when || "",
          t.ticket_no || "",
          t.ticket_name || "",
          lineLbl,
          Math.round(t.amount),
          aff,
          t.detail || "",
        ]);
      });
      return rows;
    }

    function toRow(a, validated) {
      return {
        "Validée": validated ? "Oui" : "Non",
        "Gravité": SEV_BADGE[a.severity], "Source": a.source, "Type": a.type,
        "Fichier": a.file || "", "Ligne": a.row === "" ? "" : a.row,
        "Date/heure": a.when || "",
        "Ticket POS": a.ticket_name, "N° POS": a.pos_ticket_no || "",
        "Réf. source": a.source_ref,
        "Montant POS": a.amount_pos, "Montant source": a.amount_source,
        "Paiement POS": a.payment_pos, "Paiement attendu": a.payment_source,
        "Détail": a.detail,
      };
    }

    function sheetSafe(name) {
      return String(name).replace(/[\\/*?:\[\]]/g, "-").slice(0, 31);
    }

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
      buildResumeRows(sm, counts, "Général — toute la période")), "Résumé");

    if (fin) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildFinRows(fin)), "Réconciliation €");
      if (fin.payment_breakdown) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
          buildPaymentBreakdownRows(fin)), "Totaux paiement");
      }
      if (fin.cash_to_collect) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
          buildCashCollectRows(fin)), "Cash à collecter");
        if (fin.cash_to_collect.by_user && fin.cash_to_collect.by_user.length) {
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
            buildCashCollectUserRows(fin)), "Cash par utilisateur");
        }
        if (fin.cash_to_collect.ticket_contributions &&
            fin.cash_to_collect.ticket_contributions.length) {
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
            buildCashCollectTicketRows(fin)), "Cash tickets détail");
        }
      }
    }

    function appendAnomalySheets(runResult, prefix) {
      var sheetPrefix = prefix ? prefix + " " : "";
      var byOrder = function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; };
      var anom = runResult.anomalies.filter(function (a) { return a.severity !== "info"; })
        .sort(byOrder).map(function (a) { return toRow(a, isValidated(a.id)); });
      if (!anom.length) anom = [{ "Gravité": "✅ Aucune anomalie", "Détail": "Tout est réconcilié." }];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(anom),
        sheetSafe(sheetPrefix + "Anomalies"));

      var infos = runResult.anomalies.filter(function (a) {
        return a.severity === "info" && !isValidated(a.id);
      }).map(function (a) { return toRow(a, false); });
      if (!infos.length) infos = [{ "Gravité": "—", "Détail": "Aucune info." }];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(infos),
        sheetSafe(sheetPrefix + "Infos"));

      var posRows = runResult.pos.map(function (px) {
        return {
          "Ticket No.": px.ticket_no, "Date": px.date, "Heure": px.hour, "Caissier": px.user,
          "Ticket name": px.ticket_name, "Canal détecté": px.channel,
          "Total": isNaN(px.total) ? "" : px.total, "Mode de paiement": px.payment_type,
          "Statut": px.statut, "Anomalies détectées": px.anomalies,
        };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(posRows),
        sheetSafe(sheetPrefix ? sheetPrefix + "POS" : "POS annoté"));
    }

    appendAnomalySheets(state, "");

    if (state.spemp_review && state.spemp_review.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
        state.spemp_review.map(function (r) {
          return {
            "Date/heure": r.when, "Ticket name": r.ticket_name, "N° POS": r.ticket_no,
            "Total": r.total, "Paiement": r.payment_type,
            "Type": r.kind === "a_rattacher" ? "À rattacher" :
              (r.kind === "bipeur" ? "Bipeur SP&EMP (NAPS OK)" :
                (r.kind === "cash_comptoir" ? "Cash comptoir SP&EMP" :
                  (r.kind === "cb_naps" ? "CB comptoir (NAPS OK)" :
                    "Libellé libre (SP&EMP)"))),
          };
        })), "SP&EMP hors Glovo-Site");
    }

    if (state.byDay && Object.keys(state.byDay).length > 1) {
      Object.keys(state.byDay).sort().forEach(function (dk) {
        var dayRun = state.byDay[dk];
        var dayLabel = fmtDayShort(dk);
        var dayCounts = countsForRun(dayRun);
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
          buildResumeRows(dayRun.summary, dayCounts, fmtDayLabel(dk))),
          sheetSafe(dayLabel + " Résumé"));
        var dayFin = adjustedFinFor(dayRun);
        if (dayFin) {
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildFinRows(dayFin)),
            sheetSafe(dayLabel + " Réconcil"));
          if (dayFin.payment_breakdown) {
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
              buildPaymentBreakdownRows(dayFin)), sheetSafe(dayLabel + " Paiements"));
          }
          if (dayFin.cash_to_collect) {
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
              buildCashCollectRows(dayFin)), sheetSafe(dayLabel + " Cash"));
            if (dayFin.cash_to_collect.by_user && dayFin.cash_to_collect.by_user.length) {
              XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
                buildCashCollectUserRows(dayFin)), sheetSafe(dayLabel + " Cash user"));
            }
            if (dayFin.cash_to_collect.ticket_contributions &&
                dayFin.cash_to_collect.ticket_contributions.length) {
              XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(
                buildCashCollectTicketRows(dayFin)), sheetSafe(dayLabel + " Cash tickets"));
            }
          }
        }
        appendAnomalySheets(dayRun, dayLabel);
      });
    }

    var stamp = new Date().toISOString().slice(0, 16).replace(/[-T:]/g, "").slice(0, 13);
    XLSX.writeFile(wb, "reconciliation_chicknster_" + stamp + ".xlsx");
  });

  function checkedValues(cls) {
    return Array.prototype.slice.call(document.querySelectorAll("." + cls + ":checked"))
      .map(function (c) { return c.value; });
  }
  function uniq(arr) { return arr.filter(function (v, i) { return arr.indexOf(v) === i; }); }
  function escapeAttr(str) {
    return String(str).replace(/&/g, "&amp;").replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;").replace(/</g, "&lt;");
  }
  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
})();
