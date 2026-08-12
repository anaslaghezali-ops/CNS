/* ChickNSter — câblage de l'interface + export Excel (navigateur). */
(function () {
  "use strict";

  var files = { pos: null, glovo: null, naps: null, site: null };
  var state = null; // { anomalies, pos, summary, validated: { id: true } }

  var SEV_BADGE = { haute: "🔴 Haute", moyenne: "🟠 Moyenne", info: "🔵 Info" };
  var SEV_ORDER = { haute: 0, moyenne: 1, info: 2 };
  var finDetailLineKey = null;

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
        state.validated = {};
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

  function activeAnomalies() {
    if (!state) return [];
    return state.anomalies.filter(function (a) { return !isValidated(a.id); });
  }

  function validatedAnomalies() {
    if (!state) return [];
    return state.anomalies.filter(function (a) { return isValidated(a.id); });
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

  function refreshPosStatuts() {
    var active = activeAnomalies();
    state.pos.forEach(function (p) {
      var tn = p.ticket_name;
      var related = active.filter(function (a) {
        return tn && tn !== "(vide)" && a.ticket_name === tn;
      });
      var sev = {};
      related.forEach(function (a) { sev[a.severity] = true; });
      p.statut = (sev.haute || sev.moyenne) ? "⚠️ Anomalie" :
                 (sev.info ? "ℹ️ Info" : "✅ OK");
      p.anomalies = related.map(function (a) {
        return "[" + a.type + "] " + a.detail;
      }).join(" | ");
    });
  }

  function recomputeCounts() {
    var active = activeAnomalies();
    var sev = { haute: 0, moyenne: 0, info: 0 };
    active.forEach(function (a) {
      sev[a.severity] = (sev[a.severity] || 0) + 1;
    });
    return {
      n_anomalies: sev.haute + sev.moyenne,
      n_infos: sev.info,
      severity: sev,
      n_validated: validatedAnomalies().length,
    };
  }

  function getAdjustedFinancial() {
    if (!state || !state.summary.financial) return null;
    var validated = validatedAnomalies();
    if (!validated.length) return state.summary.financial;
    var adj = CNS.sumFinancialAdjustments(validated);
    return CNS.applyFinancialAdjustments(state.summary.financial, adj);
  }

  // ---- Rendu ------------------------------------------------------------- //
  function render() {
    document.getElementById("results").classList.remove("hidden");
    var sm = state.summary;
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
    document.getElementById("metrics").innerHTML = metrics.map(function (m) {
      return '<div class="metric"><div class="label">' + m[0] +
             '</div><div class="value">' + m[1] + "</div></div>";
    }).join("");

    var period = document.getElementById("period");
    var txt = "📅 Période analysée (d'après le POS) : <b>" +
              escapeHtml(sm.pos_date_min) + "</b> → <b>" + escapeHtml(sm.pos_date_max) + "</b>.";
    var ex = [];
    if (sm.glovo_excluded) ex.push(sm.glovo_excluded + " commande(s) Glovo");
    if (sm.site_excluded) ex.push(sm.site_excluded + " commande(s) Site");
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

    renderFinancial(getAdjustedFinancial());

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
                "<td>" + escapeHtml(a.ticket_name) + "</td>" +
                "<td>" + escapeHtml(a.detail) + "</td></tr>";
      });
      html += "</tbody></table></div>";
    }
    panel.innerHTML = html;
    panel.classList.remove("hidden");
    bindValidateButtons(panel);
  }

  function _tableHTML(rows, withValidate) {
    var head = "<thead><tr>";
    if (withValidate) head += "<th>Valider</th>";
    head += "<th>Gravité</th><th>Source</th><th>Type</th>" +
            "<th>Fichier</th><th>Ligne</th><th>Date/heure</th>" +
            "<th>Ticket</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + rows.map(function (a) {
      var btn = withValidate ?
        '<td><button type="button" class="btn-validate" data-id="' + escapeHtml(a.id) +
        '" title="Valider — hors calcul">✅ Valider</button></td>' : "";
      return "<tr>" + btn +
             "<td><span class='sev-badge sev-" + a.severity + "'>" +
             SEV_BADGE[a.severity] + "</span></td><td>" + escapeHtml(a.source) +
             "</td><td>" + escapeHtml(a.type) + "</td><td>" + escapeHtml(a.file || "") +
             "</td><td>" + escapeHtml(a.row === "" || a.row == null ? "" : a.row) +
             "</td><td>" + escapeHtml(a.when || "") + "</td><td>" +
             escapeHtml(a.ticket_name) + "</td><td>" + escapeHtml(a.detail) + "</td></tr>";
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
    var rows = activeAnomalies().filter(function (a) {
      return a.severity !== "info" &&
             sev.indexOf(a.severity) !== -1 && src.indexOf(a.source) !== -1;
    }).sort(function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; });

    document.getElementById("anom-count").textContent = rows.length + " anomalie(s) à traiter";
    var table = document.getElementById("anom-table");
    table.innerHTML = rows.length ? _tableHTML(rows, true) :
      "<tbody><tr><td colspan='9'>✅ Aucune anomalie en attente pour ces filtres.</td></tr></tbody>";
    if (rows.length) bindValidateButtons(table);
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
               "<th>Ticket</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + rows.map(function (a) {
      return "<tr class='row-validated'><td><button type='button' class='btn-unvalidate' " +
             "data-id='" + escapeHtml(a.id) + "' title='Annuler la validation'>↩ Annuler</button></td>" +
             "<td><span class='sev-badge sev-" + a.severity + "'>" + SEV_BADGE[a.severity] +
             "</span></td><td>" + escapeHtml(a.source) + "</td><td>" + escapeHtml(a.type) +
             "</td><td>" + escapeHtml(a.ticket_name) + "</td><td>" + escapeHtml(a.detail) + "</td></tr>";
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
    var onlyAnom = document.getElementById("only-anom").checked;
    var rows = state.pos.filter(function (p) {
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

    var resume = [
      ["Indicateur", "Valeur"],
      ["Période analysée (POS)", (sm.pos_date_min || "") + " → " + (sm.pos_date_max || "")],
      ["Transactions POS", sm.pos_transactions],
      ["Total POS (DH)", Math.round(sm.pos_total * 100) / 100],
      ["Commandes Glovo (livrées, période)", sm.glovo_orders],
      ["Commandes Glovo hors période (ignorées)", sm.glovo_excluded || 0],
      ["Transactions NAPS", sm.naps_transactions],
      ["Commandes Site (période)", sm.site_orders],
      ["Commandes Site hors période (ignorées)", sm.site_excluded || 0],
      ["", ""],
      ["Anomalies (Haute + Moyenne)", counts.n_anomalies],
      ["  dont haute", counts.severity.haute || 0],
      ["  dont moyenne", counts.severity.moyenne || 0],
      ["Anomalies validées (hors calcul)", counts.n_validated || 0],
      ["Infos (rattachements & notes)", counts.n_infos || 0],
      ["", ""],
    ];
    Object.keys(sm.channels).forEach(function (k) {
      resume.push(["POS — " + k, sm.channels[k]]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resume), "Résumé");

    if (fin) {
      var frows = [["Réconciliation financière (Écart = source − POS)"], []];
      if (fin.adjustments_applied) {
        frows.push(["(Anomalies validées exclues des montants ci-dessous)"]);
        frows.push([]);
      }
      frows.push(["Source", "Côté POS (libellé)", "Montant POS", "Côté source (libellé)", "Montant source", "Écart"]);
      fin.lines.forEach(function (l) {
        frows.push([l.source, l.pos_label, Math.round(l.pos), l.src_label, Math.round(l.src), Math.round(l.ecart)]);
        if (l.note) frows.push(["", l.note]);
      });
      frows.push([], ["Répartition POS : mode de paiement × canal"], []);
      frows.push(["Canal"].concat(fin.pays, ["Total"]));
      var colTot = {}; fin.pays.forEach(function (p) { colTot[p] = 0; });
      Object.keys(fin.matrix).forEach(function (ch) {
        var row = fin.matrix[ch], rt = 0;
        var cells = fin.pays.map(function (p) { var v = row[p] || 0; rt += v; colTot[p] += v; return Math.round(v); });
        frows.push([ch].concat(cells, [Math.round(rt)]));
      });
      var grand = 0; fin.pays.forEach(function (p) { grand += colTot[p]; });
      frows.push(["Total"].concat(fin.pays.map(function (p) { return Math.round(colTot[p]); }), [Math.round(grand)]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(frows), "Réconciliation €");
    }

    function toRow(a, validated) {
      return {
        "Validée": validated ? "Oui" : "Non",
        "Gravité": SEV_BADGE[a.severity], "Source": a.source, "Type": a.type,
        "Fichier": a.file || "", "Ligne": a.row === "" ? "" : a.row,
        "Date/heure": a.when || "",
        "Ticket POS": a.ticket_name, "Réf. source": a.source_ref,
        "Montant POS": a.amount_pos, "Montant source": a.amount_source,
        "Paiement POS": a.payment_pos, "Paiement attendu": a.payment_source,
        "Détail": a.detail,
      };
    }
    var byOrder = function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; };

    var anom = state.anomalies.filter(function (a) { return a.severity !== "info"; })
                              .sort(byOrder).map(function (a) {
      return toRow(a, isValidated(a.id));
    });
    if (!anom.length) anom = [{ "Gravité": "✅ Aucune anomalie", "Détail": "Tout est réconcilié." }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(anom), "Anomalies");

    var infos = activeAnomalies().filter(function (a) { return a.severity === "info"; }).map(function (a) {
      return toRow(a, false);
    });
    if (!infos.length) infos = [{ "Gravité": "—", "Détail": "Aucune info." }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(infos), "Infos");

    var posRows = state.pos.map(function (p) {
      return {
        "Ticket No.": p.ticket_no, "Date": p.date, "Heure": p.hour, "Caissier": p.user,
        "Ticket name": p.ticket_name, "Canal détecté": p.channel,
        "Total": isNaN(p.total) ? "" : p.total, "Mode de paiement": p.payment_type,
        "Statut": p.statut, "Anomalies détectées": p.anomalies,
      };
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(posRows), "POS annoté");

    var stamp = new Date().toISOString().slice(0, 16).replace(/[-T:]/g, "").slice(0, 13);
    XLSX.writeFile(wb, "reconciliation_chicknster_" + stamp + ".xlsx");
  });

  function checkedValues(cls) {
    return Array.prototype.slice.call(document.querySelectorAll("." + cls + ":checked"))
      .map(function (c) { return c.value; });
  }
  function uniq(arr) { return arr.filter(function (v, i) { return arr.indexOf(v) === i; }); }
  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
})();
