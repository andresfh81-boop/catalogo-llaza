/* Catálogo Llaza instalable (PWA).
   El cliente (client.js, el mismo de la app de Apps Script) llama a google.script.run; aquí se hace con fetch
   a la API (proyecto «Catálogo Llaza API» del repositorio catalogo-llaza-app, carpeta api/), con el ID token
   de «Iniciar sesión con Google». La última copia del catálogo se guarda para abrir al momento y sin conexión. */
(function () {
  "use strict";
  var API = "https://script.google.com/macros/s/AKfycbx6O4pyMDoRBnUyWfemQf4kDcG2AqS0yCgYnO4BsExkOi5_Bg_v3t2QshV18qqMT3RH/exec";
  var CLIENT_ID = "262401648003-agcod92m3lne93rpgf3ulc54p0695p1q.apps.googleusercontent.com";
  var FULL_APP = "https://script.google.com/macros/s/AKfycbyqeQNPmjtzhhU6pr4uA97C4o1Xq4tOZKgtxJ52k36JBoo7cncbaQXtrzK9bwJ39gz0IQ/exec";
  var DATA = "cl-datos";

  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js");

  /* ---------- barra de avisos (sesión, sin conexión, datos nuevos) ---------- */
  function bar(html, kind) {
    var b = document.getElementById("pwaBar"); if (!b) return;
    if (!html) { b.hidden = true; b.innerHTML = ""; return; }
    b.className = kind || ""; b.innerHTML = html; b.hidden = false;
  }
  var fmt = function (t) { var d = new Date(t); return d.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit" }) + " " + d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" }); };

  /* ---------- sesión con Google ---------- */
  var tok = null; try { tok = JSON.parse(localStorage.getItem("cl_tok") || "null"); } catch (e) {}
  var valid = function () { return tok && tok.exp * 1000 > Date.now() + 60000; };
  var waiters = [], gis = false, loginShown = false, rendered = false;
  function payload(t) { var s = t.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"); return JSON.parse(decodeURIComponent(escape(atob(s)))); }
  function onCred(r) {
    var p = payload(r.credential);
    tok = { t: r.credential, exp: p.exp, email: p.email };
    try { localStorage.setItem("cl_tok", JSON.stringify(tok)); } catch (e) {}
    rendered = false;
    if (loginShown) { loginShown = false; bar(""); }
    var w = waiters; waiters = []; w.forEach(function (f) { f(tok.t); });
  }
  function initGis() {
    if (gis || !window.google || !google.accounts || !google.accounts.id) return;
    gis = true;
    google.accounts.id.initialize({ client_id: CLIENT_ID, callback: onCred, auto_select: true, use_fedcm_for_prompt: true, itp_support: true, context: "signin" });
    if (waiters.length) showLogin();
  }
  window.__gisReady = initGis;
  function showLogin() {
    if (rendered) return; /* el botón ya está y Google ya lo ha intentado: llamar otra vez a prompt() cancela el primero */
    loginShown = true;
    bar('<span>Inicia sesión con tu cuenta de Google para ' + (hasCache ? "actualizar" : "cargar") + ' el catálogo.</span><span id="gBtn"></span>', "login");
    if (!gis) { if (!navigator.onLine) bar("<span>Sin conexión. Conéctate para iniciar sesión.</span>", "off"); return; }
    google.accounts.id.renderButton(document.getElementById("gBtn"), { theme: "filled_blue", size: "large", text: "signin_with", shape: "pill", locale: "es" });
    rendered = true;
    google.accounts.id.prompt();
  }
  function getToken() {
    if (valid()) return Promise.resolve(tok.t);
    return new Promise(function (res) { waiters.push(res); if (gis) showLogin(); else if (waiters.length === 1) setTimeout(function () { if (!gis && waiters.length) showLogin(); }, 4000); });
  }

  /* ---------- llamadas a la API ---------- */
  function rpc(fn, args, again) {
    return getToken().then(function (t) {
      return fetch(API, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ t: t, fn: fn, args: args }) });
    }).then(function (r) {
      if (!r.ok) throw new Error("La API ha respondido con un error " + r.status + ".");
      return r.json();
    }).then(function (j) {
      if (!j.ok && j.auth && !again) { tok = null; try { localStorage.removeItem("cl_tok"); } catch (e) {} return rpc(fn, args, true); }
      if (!j.ok) throw new Error(j.error || "Algo ha fallado.");
      return j.r;
    });
  }
  function readC(k) { if (!window.caches) return Promise.resolve(null); return caches.open(DATA).then(function (c) { return c.match("/d/" + k); }).then(function (r) { return r ? r.json() : null; }).catch(function () { return null; }); }
  function saveC(k, v) { if (!window.caches) return Promise.resolve(); return caches.open(DATA).then(function (c) { return c.put("/d/" + k, new Response(JSON.stringify(v), { headers: { "Content-Type": "application/json" } })); }).catch(function () {}); }

  /* Arranque: si hay copia guardada se abre con ella al momento y se actualiza detrás; si hay cambios, se avisa. */
  var hasCache = false;
  function boot(args) {
    return readC("boot").then(function (c) {
      if (!c) return rpc("boot", args).then(function (d) { saveC("boot", { at: Date.now(), d: d }); return d; });
      hasCache = true;
      if (!navigator.onLine) bar("<span>Sin conexión · copia del " + fmt(c.at) + "</span>", "off");
      rpc("boot", args).then(function (d) {
        var changed = JSON.stringify(d) !== JSON.stringify(c.d);
        return saveC("boot", { at: Date.now(), d: d }).then(function () {
          if (changed) bar('<span>Hay cambios en el catálogo.</span><button type="button" class="btn-s" onclick="location.reload()">Ver</button>', "new");
          else if (!loginShown) bar("");
        });
      }).catch(function () { bar("<span>Sin conexión · copia del " + fmt(c.at) + "</span>", "off"); });
      return c.d;
    });
  }
  function call(fn, args) {
    if (fn === "boot") return boot(args);
    if (fn === "getFull") {
      var k = "full/" + args[0];
      return readC(k).then(function (c) { return c || rpc(fn, args).then(function (d) { saveC(k, d); return d; }); });
    }
    if (fn === "syncSheetsBg") return Promise.resolve(null);
    return rpc(fn, args);
  }
  window.addEventListener("online", function () { var b = document.getElementById("pwaBar"); if (b && b.className === "off") bar(""); });
  window.addEventListener("offline", function () { if (hasCache) bar("<span>Sin conexión · estás viendo la última copia guardada.</span>", "off"); });

  /* ---------- google.script.run para client.js ---------- */
  function runner(ok, ko) {
    return new Proxy({}, { get: function (_, k) {
      if (k === "withSuccessHandler") return function (f) { return runner(f, ko); };
      if (k === "withFailureHandler") return function (f) { return runner(ok, f); };
      return function () { var a = [].slice.call(arguments); call(k, a).then(function (r) { ok && ok(r); }, function (e) { ko && ko(e); }); };
    } });
  }
  window.google = window.google || {};
  google.script = { run: runner(null, null) };

  /* ---------- descargar todo para usar sin conexión: fotos grandes y PDF de los manuales ----------
     Va en la misma caché que las fotos ya vistas («full/<fileId>»), así que client.js las encuentra solas. Se puede parar:
     al volver a pulsar sigue por donde iba. Los vídeos y las webs no se pueden guardar. */
  var dlActivo = false, dlParar = false;
  function dlBar(html) {
    var b = document.getElementById("pwaDlBar");
    if (!b) { document.body.insertAdjacentHTML("afterbegin", '<div id="pwaDlBar"></div>'); b = document.getElementById("pwaDlBar"); }
    b.hidden = !html; b.innerHTML = html || "";
  }
  function pdfId(m) {
    var x = /drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:[^#]*&)?id=)([\w-]{20,})/.exec(m.url || "");
    return x ? x[1] : "";
  }
  function dlLista() {
    return readC("boot").then(function (c) {
      if (!c || !c.d) return [];
      var vistos = {}, L = [];
      function pon(id, tipo) { if (id && !vistos[id]) { vistos[id] = 1; L.push({ id: id, tipo: tipo }); } }
      (c.d.fotos || []).forEach(function (f) { pon(f.fileId, "foto"); });
      (c.d.manuales || []).forEach(function (m) { pon(m.fileId || pdfId(m), "pdf"); });
      return L;
    }).then(function (L) {
      return caches.open(DATA).then(function (c) {
        return Promise.all(L.map(function (x) { return c.match("/d/full/" + x.id).then(function (r) { return r ? null : x; }); }));
      }).then(function (R) { return { total: L.length, faltan: R.filter(Boolean) }; });
    });
  }
  function dlOcupa() {
    if (!navigator.storage || !navigator.storage.estimate) return Promise.resolve("");
    return navigator.storage.estimate().then(function (e) { return " Ocupa " + Math.round((e.usage || 0) / 1048576) + " MB en este dispositivo."; }).catch(function () { return ""; });
  }
  function dlEmpieza() {
    if (dlActivo) return;
    if (!navigator.onLine) { dlBar('<span>Sin conexión: conéctate para descargar.</span><button type="button" class="btn-s" data-dl="x">Cerrar</button>'); return; }
    dlLista().then(function (r) {
      var F = r.faltan, nf = F.filter(function (x) { return x.tipo === "foto"; }).length;
      if (!r.total) { dlBar('<span>Primero tiene que cargar el catálogo.</span><button type="button" class="btn-s" data-dl="x">Cerrar</button>'); return; }
      if (!F.length) { return dlOcupa().then(function (o) { dlBar("<span>Ya está todo descargado." + o + '</span><button type="button" class="btn-s" data-dl="x">Cerrar</button>'); }); }
      if (!confirm("Faltan " + nf + (nf === 1 ? " foto" : " fotos") + " y " + (F.length - nf) + (F.length - nf === 1 ? " PDF" : " PDF") + " por descargar. Puede ocupar bastante y tardar: mejor con WiFi. ¿Empiezo?")) return;
      dlActivo = true; dlParar = false;
      var i = 0, hechos = 0, fallos = 0, lleno = false;
      function pinta() { dlBar("<span>Descargando " + (hechos + fallos) + " de " + F.length + "…</span>" + '<button type="button" class="btn-s" data-dl="parar">Parar</button>'); }
      function uno() {
        if (dlParar || lleno || i >= F.length) return Promise.resolve();
        var x = F[i++];
        return rpc("getFull", [x.id]).then(function (d) {
          return caches.open(DATA).then(function (c) { return c.put("/d/full/" + x.id, new Response(JSON.stringify(d), { headers: { "Content-Type": "application/json" } })); });
        }).then(function () { hechos++; }, function (e) { fallos++; if (e && e.name === "QuotaExceededError") lleno = true; }).then(function () { pinta(); return uno(); });
      }
      pinta();
      /* tres a la vez: de uno en uno tardaría demasiado y más le cuesta a la API */
      Promise.all([uno(), uno(), uno()]).then(dlOcupa).then(function (o) {
        dlActivo = false;
        var t = lleno ? "No queda espacio en este dispositivo: " + hechos + " descargados." : dlParar ? "Parado: " + hechos + " descargados. Vuelve a pulsar «Descargar todo» para seguir." : fallos ? hechos + " descargados y " + fallos + " sin poder bajar. Vuelve a pulsar «Descargar todo» para reintentarlos." : "Listo: todo descargado para usar sin conexión.";
        dlBar("<span>" + t + o + '</span><button type="button" class="btn-s" data-dl="x">Cerrar</button>');
      });
    });
  }
  /* Sin conexión, los enlaces a un PDF de Drive abren la copia descargada (en el móvil, como descarga: no abre PDF en una pestaña). */
  function abrePdf(id, nombre) {
    return readC("full/" + id).then(function (d) {
      if (!d) { dlBar('<span>Este documento no está descargado. Con conexión, pulsa «Descargar todo».</span><button type="button" class="btn-s" data-dl="x">Cerrar</button>'); return; }
      var bin = atob(d.b64), u8 = new Uint8Array(bin.length);
      for (var k = 0; k < bin.length; k++) u8[k] = bin.charCodeAt(k);
      var url = URL.createObjectURL(new Blob([u8], { type: d.mime || "application/pdf" }));
      if (/android|iphone|ipad|ipod/i.test(navigator.userAgent)) {
        var a = document.createElement("a"); a.href = url; a.download = (nombre || "documento").replace(/[\\/:*?"<>|]+/g, " ") + ".pdf"; document.body.appendChild(a); a.click(); a.remove();
      } else window.open(url, "_blank");
    });
  }
  document.addEventListener("click", function (e) {
    var b = e.target.closest && e.target.closest("[data-dl]");
    if (b) { if (b.dataset.dl === "parar") { dlParar = true; b.disabled = true; b.textContent = "Parando…"; } else dlBar(""); return; }
    if (navigator.onLine) return;
    var a = e.target.closest && e.target.closest("a[href]"), id = a && pdfId({ url: a.href });
    if (id) { e.preventDefault(); abrePdf(id, a.textContent.trim()); }
  });

  /* ---------- botones propios: app completa (para editar) e instalar ---------- */
  document.addEventListener("DOMContentLoaded", function () {
    var h = document.querySelector(".hdr-btns");
    if (h) h.insertAdjacentHTML("afterbegin", '<a class="btn-s" id="pwaFull" href="' + FULL_APP + '" target="_blank" rel="noopener" title="Para editar, añadir, revisar hojas y exportar">✏ App completa</a>');
    if (h) { h.insertAdjacentHTML("afterbegin", '<button type="button" class="btn-s" id="pwaDl" title="Baja las fotos grandes y los PDF de los manuales para usarlos sin internet">⬇ Descargar todo</button>'); document.getElementById("pwaDl").onclick = dlEmpieza; }
    var instalada = matchMedia("(display-mode: standalone)").matches || navigator.standalone, aviso = null;
    var oculta = false; try { oculta = localStorage.getItem("catalogoLlazaNoInstalar") === "1"; } catch (e) {}
    document.body.insertAdjacentHTML("afterbegin", '<div class="pwaInst" id="pwaInst" hidden><img src="icon/icon-192.png" alt=""><span id="pwaInstTxt">Instala el Catálogo como una app</span><button type="button" id="pwaInstBtn">📲 Instalar</button><button type="button" id="pwaInstNo" aria-label="Ahora no">✕</button></div>');
    var box = document.getElementById("pwaInst"), txt = document.getElementById("pwaInstTxt"), btn = document.getElementById("pwaInstBtn");
    var cerrar = function () { box.hidden = true; try { localStorage.setItem("catalogoLlazaNoInstalar", "1"); } catch (e) {} };
    window.addEventListener("beforeinstallprompt", function (e) { e.preventDefault(); aviso = e; });
    window.addEventListener("appinstalled", cerrar);
    if (!instalada && !oculta) {
      box.hidden = false;
      if (/iphone|ipad|ipod/i.test(navigator.userAgent)) { txt.textContent = "Para instalarla: botón Compartir → «Añadir a pantalla de inicio»"; btn.hidden = true; }
    }
    btn.onclick = function () {
      if (!aviso) { txt.textContent = "Toca los tres puntos ⋮ de Chrome → «Instalar aplicación» (o «Añadir a pantalla de inicio»)"; return; }
      aviso.prompt(); aviso.userChoice.then(function (r) { if (r.outcome === "accepted") cerrar(); aviso = null; });
    };
    document.getElementById("pwaInstNo").onclick = cerrar;
  });
})();
