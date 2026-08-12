/* ChickNSter — câblage de l'interface + export Excel (navigateur). */
(function () {
  "use strict";

  var files = { pos: null, glovo: null, naps: null, site: null };
  var state = null; // { anomalies, pos, summary }

  var SEV_BADGE = { haute: "🔴 Haute", moyenne: "🟠 Moyenne", info: "🔵 Info" };
  var SEV_ORDER = { haute: 0, moyenne: 1, info: 2 };

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

  // ---- Lecture d'un fichier -> worksheet -------------------------------- //
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

  // ---- Lancer la réconciliation ----------------------------------------- //
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

  // ---- Rendu ------------------------------------------------------------- //
  function render() {
    document.getElementById("results").classList.remove("hidden");
    var sm = state.summary;

    // Métriques (les « info » ne comptent pas comme anomalies)
    var metrics = [
      ["Transactions POS", sm.pos_transactions],
      ["Total POS (DH)", Math.round(sm.pos_total).toLocaleString("fr-FR")],
      ["Anomalies", sm.n_anomalies],
      ["🔴 Haute", sm.severity.haute || 0],
      ["🟠 Moyenne", sm.severity.moyenne || 0],
      ["🔵 Infos", sm.n_infos || 0],
    ];
    document.getElementById("metrics").innerHTML = metrics.map(function (m) {
      return '<div class="metric"><div class="label">' + m[0] +
             '</div><div class="value">' + m[1] + "</div></div>";
    }).join("");

    // Période analysée (définie par le POS) + éléments hors-période ignorés
    var period = document.getElementById("period");
    var txt = "📅 Période analysée (d'après le POS) : <b>" +
              escapeHtml(sm.pos_date_min) + "</b> → <b>" + escapeHtml(sm.pos_date_max) + "</b>.";
    var ex = [];
    if (sm.glovo_excluded) ex.push(sm.glovo_excluded + " commande(s) Glovo");
    if (sm.site_excluded) ex.push(sm.site_excluded + " commande(s) Site");
    if (ex.length) txt += " " + ex.join(" et ") + " hors de cette période ont été ignorée(s).";
    period.innerHTML = txt;

    // Graphique par canal
    var ch = sm.channels;
    var keys = Object.keys(ch);
    var max = Math.max.apply(null, keys.map(function (k) { return ch[k]; })) || 1;
    document.getElementById("chart").innerHTML = keys.map(function (k) {
      var h = Math.round((ch[k] / max) * 100);
      return '<div class="bar"><div class="bar-val">' + ch[k] +
             '</div><div class="fill" style="height:' + h + '%"></div>' +
             '<div class="bar-label">' + escapeHtml(k) + "</div></div>";
    }).join("");

    renderFinancial(sm.financial);

    // Filtres source
    var sources = uniq(state.anomalies.map(function (a) { return a.source; }));
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
    renderInfos();
    renderPos();
  }

  function fmtDH(v) {
    if (v == null || isNaN(v)) return "";
    return Math.round(v).toLocaleString("fr-FR") + " DH";
  }

  function renderFinancial(fin) {
    if (!fin) { document.getElementById("fin-lines").innerHTML = ""; return; }
    // Tableau des écarts par source
    var head = "<thead><tr><th>Source</th><th>Côté POS</th><th>Côté source</th>" +
               "<th>Écart (source − POS)</th></tr></thead>";
    var body = "<tbody>" + fin.lines.map(function (l) {
      var cls = Math.abs(l.ecart) < 0.5 ? "st-ok" : "st-anom";
      var sign = l.ecart > 0 ? "+" : "";
      var note = l.note ? "<div class='muted' style='font-size:.82rem'>" + escapeHtml(l.note) + "</div>" : "";
      var rowCls = l.isTotal ? "fin-total" : (l.group === "glovo" ? "fin-glovo-sub" : "");
      return "<tr class='" + rowCls + "'><td><b>" + escapeHtml(l.source) + "</b></td>" +
             "<td>" + escapeHtml(l.pos_label) + " : <b>" + fmtDH(l.pos) + "</b></td>" +
             "<td>" + escapeHtml(l.src_label) + " : <b>" + fmtDH(l.src) + "</b>" + note + "</td>" +
             "<td class='" + cls + "'><b>" + sign + fmtDH(l.ecart) + "</b></td></tr>";
    }).join("") + "</tbody>";
    document.getElementById("fin-lines").innerHTML = head + body;

    // Matrice canal × paiement
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

  function _tableHTML(rows) {
    var head = "<thead><tr><th>Gravité</th><th>Source</th><th>Type</th>" +
               "<th>Fichier</th><th>Ligne</th><th>Date/heure</th>" +
               "<th>Ticket</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + rows.map(function (a) {
      return "<tr><td><span class='sev-badge sev-" + a.severity + "'>" +
             SEV_BADGE[a.severity] + "</span></td><td>" + escapeHtml(a.source) +
             "</td><td>" + escapeHtml(a.type) + "</td><td>" + escapeHtml(a.file || "") +
             "</td><td>" + escapeHtml(a.row === "" || a.row == null ? "" : a.row) +
             "</td><td>" + escapeHtml(a.when || "") + "</td><td>" +
             escapeHtml(a.ticket_name) + "</td><td>" + escapeHtml(a.detail) + "</td></tr>";
    }).join("") + "</tbody>";
    return head + body;
  }

  function renderAnomalies() {
    // Anomalies = Haute + Moyenne uniquement.
    var sev = checkedValues("fsev"), src = checkedValues("fsrc");
    var rows = state.anomalies.filter(function (a) {
      return a.severity !== "info" &&
             sev.indexOf(a.severity) !== -1 && src.indexOf(a.source) !== -1;
    }).sort(function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; });

    document.getElementById("anom-count").textContent = rows.length + " anomalie(s) affichée(s)";
    document.getElementById("anom-table").innerHTML = rows.length ? _tableHTML(rows) :
      "<tbody><tr><td>✅ Aucune anomalie pour ces filtres.</td></tr></tbody>";
  }

  function renderInfos() {
    // Section Infos = éléments informatifs (rattachements, saisies tardives,
    // tickets sans numéro, écarts globaux) — NON comptés comme anomalies.
    var src = checkedValues("fsrc");
    var rows = state.anomalies.filter(function (a) {
      return a.severity === "info" && src.indexOf(a.source) !== -1;
    });
    document.getElementById("infos-count").textContent = rows.length + " info(s) affichée(s)";
    document.getElementById("infos-table").innerHTML = rows.length ? _tableHTML(rows) :
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

  // ---- Export Excel ------------------------------------------------------ //
  document.getElementById("download").addEventListener("click", function () {
    if (!state) return;
    var sm = state.summary;
    var wb = XLSX.utils.book_new();

    // Résumé
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
      ["Anomalies (Haute + Moyenne)", sm.n_anomalies],
      ["  dont haute", sm.severity.haute || 0],
      ["  dont moyenne", sm.severity.moyenne || 0],
      ["Infos (rattachements & notes)", sm.n_infos || 0],
      ["", ""],
    ];
    Object.keys(sm.channels).forEach(function (k) {
      resume.push(["POS — " + k, sm.channels[k]]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resume), "Résumé");

    // Réconciliation financière
    var fin = sm.financial;
    if (fin) {
      var frows = [["Réconciliation financière (Écart = source − POS)"], []];
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

    function toRow(a) {
      return {
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

    // Anomalies (Haute + Moyenne)
    var anom = state.anomalies.filter(function (a) { return a.severity !== "info"; })
                              .sort(byOrder).map(toRow);
    if (!anom.length) anom = [{ "Gravité": "✅ Aucune anomalie", "Détail": "Tout est réconcilié." }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(anom), "Anomalies");

    // Infos (rattachements & notes)
    var infos = state.anomalies.filter(function (a) { return a.severity === "info"; }).map(toRow);
    if (!infos.length) infos = [{ "Gravité": "—", "Détail": "Aucune info." }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(infos), "Infos");

    // POS annoté
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

  // ---- Utilitaires ------------------------------------------------------- //
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
