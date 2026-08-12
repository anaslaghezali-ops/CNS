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

    // Métriques
    var metrics = [
      ["Transactions POS", sm.pos_transactions],
      ["Total POS (DH)", Math.round(sm.pos_total).toLocaleString("fr-FR")],
      ["Anomalies", sm.n_anomalies],
      ["🔴 Haute", sm.severity.haute || 0],
      ["🟠 Moyenne", sm.severity.moyenne || 0],
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
      cb.onchange = renderAnomalies;
    });
    document.getElementById("only-anom").onchange = renderPos;

    renderAnomalies();
    renderPos();
  }

  function renderAnomalies() {
    var sev = checkedValues("fsev"), src = checkedValues("fsrc");
    var rows = state.anomalies.filter(function (a) {
      return sev.indexOf(a.severity) !== -1 && src.indexOf(a.source) !== -1;
    }).sort(function (a, b) { return SEV_ORDER[a.severity] - SEV_ORDER[b.severity]; });

    document.getElementById("anom-count").textContent = rows.length + " anomalie(s) affichée(s)";

    if (!rows.length) {
      document.getElementById("anom-table").innerHTML =
        "<tbody><tr><td>✅ Aucune anomalie pour ces filtres.</td></tr></tbody>";
      return;
    }
    var head = "<thead><tr><th>Gravité</th><th>Source</th><th>Type</th>" +
               "<th>Ticket</th><th>Détail</th></tr></thead>";
    var body = "<tbody>" + rows.map(function (a) {
      return "<tr><td><span class='sev-badge sev-" + a.severity + "'>" +
             SEV_BADGE[a.severity] + "</span></td><td>" + escapeHtml(a.source) +
             "</td><td>" + escapeHtml(a.type) + "</td><td>" +
             escapeHtml(a.ticket_name) + "</td><td>" + escapeHtml(a.detail) + "</td></tr>";
    }).join("") + "</tbody>";
    document.getElementById("anom-table").innerHTML = head + body;
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
      ["Anomalies — total", sm.n_anomalies],
      ["  dont haute", sm.severity.haute || 0],
      ["  dont moyenne", sm.severity.moyenne || 0],
      ["  dont info", sm.severity.info || 0],
      ["", ""],
    ];
    Object.keys(sm.channels).forEach(function (k) {
      resume.push(["POS — " + k, sm.channels[k]]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resume), "Résumé");

    // Anomalies
    var anom = state.anomalies.slice().sort(function (a, b) {
      return SEV_ORDER[a.severity] - SEV_ORDER[b.severity];
    }).map(function (a) {
      return {
        "Gravité": SEV_BADGE[a.severity], "Source": a.source, "Type": a.type,
        "Ticket POS": a.ticket_name, "Réf. source": a.source_ref,
        "Montant POS": a.amount_pos, "Montant source": a.amount_source,
        "Paiement POS": a.payment_pos, "Paiement attendu": a.payment_source,
        "Détail": a.detail,
      };
    });
    if (!anom.length) anom = [{ "Gravité": "✅ Aucune anomalie", "Détail": "Tout est réconcilié." }];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(anom), "Anomalies");

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
