/**
 * Custom Video Request - MyGeotab Map Add-In (Trips History page)
 *
 * Adds a "Request custom video recording" launcher to the Trips History side
 * panel. Clicking opens a modal to pick a camera, an event start time, and a
 * duration, then queues the clip.
 *
 * Runs in the my.geotab.com document context (map add-ins share the geotab
 * global), so it calls Camera-Services directly - no proxy, no credentials.
 *
 * AUTH (matches what the native camera UI sends in this tenant):
 *   - Base: https://media-services.geotab.com/api
 *   - Authorization: Bearer <token>  (read from localStorage 'authToken',
 *     the OIDC token MyGeotab already minted - we reuse it, same origin)
 *   - X-MyGeotab-Database / -Path / -SessionId / -Userid / -Username
 *
 *   List cameras:  GET  /api/DeviceMappings
 *   Request clip:  POST /api/Media
 *
 * partnerId branches automatically: smarterai (GoFocus) | surfsight | sensata.
 */
geotab.addin.request = function (elt, service) {
  'use strict';

  var MEDIA_BASE = 'https://media-services.geotab.com/api';
  var MAX_DURATION_SECONDS = 120;

  var api = service.api;
  var mapSvc = service.map;       // setBounds / setZoom
  var canvasSvc = service.canvas; // marker / circle / clear
  var PIN_SRC = 'data:image/svg+xml;utf8,' + encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" width="30" height="38" viewBox="0 0 30 38">' +
    '<path d="M15 0C7 0 1 6 1 14c0 9.5 14 24 14 24s14-14.5 14-24C29 6 23 0 15 0z" fill="#25477B"/>' +
    '<circle cx="15" cy="14" r="5.5" fill="#fff"/></svg>');
  var session = null;
  var sessionServer = null;
  var userId = null;
  var cameras = [];
  var serialByDeviceId = {};
  var cameraIdxByGoSerial = {};
  var loaded = false;
  var capturedBearer = null;

  // Reuse the exact OIDC Bearer MyGeotab sends to media-services. The token
  // is a short-lived Keycloak token held in memory (not localStorage). This
  // add-in runs in a sandboxed-but-same-origin iframe, so we patch fetch/XHR
  // on the top window AND every same-origin frame to capture the Authorization
  // header MyGeotab uses for media-services, then reuse it. Pure browser.
  function cvrRemember(auth) {
    if (auth && /^Bearer\s+eyJ/i.test(auth)) {
      capturedBearer = auth.replace(/^Bearer\s+/i, '');
      if (!loaded && !window.__cvrRetried) {
        window.__cvrRetried = true;
        setTimeout(function () { try { init(); } catch (e) {} }, 0);
      }
    }
  }
  function cvrIsMedia(u) { return u && String(u).indexOf('media-services.geotab.com') !== -1; }
  function cvrPatch(win) {
    try {
      var of = win.fetch;
      if (of && !of.__cvrWrapped) {
        win.fetch = function (input, init) {
          try {
            var url = (typeof input === 'string') ? input : (input && input.url) || '';
            if (cvrIsMedia(url)) {
              var a = null;
              if (init && init.headers) {
                try { a = new (win.Headers || Headers)(init.headers).get('authorization'); } catch (e) {}
              }
              if (!a && input && input.headers && input.headers.get) { a = input.headers.get('authorization'); }
              cvrRemember(a);
            }
          } catch (e) {}
          return of.apply(this, arguments);
        };
        win.fetch.__cvrWrapped = true;
      }
    } catch (e) {}
    try {
      var XHR = win.XMLHttpRequest && win.XMLHttpRequest.prototype;
      if (XHR && XHR.setRequestHeader && !XHR.setRequestHeader.__cvrWrapped) {
        var oOpen = XHR.open, oSet = XHR.setRequestHeader;
        XHR.open = function (m, url) { this.__cvrUrl = url || ''; return oOpen.apply(this, arguments); };
        XHR.setRequestHeader = function (name, value) {
          try { if (String(name).toLowerCase() === 'authorization' && cvrIsMedia(this.__cvrUrl)) { cvrRemember(value); } } catch (e) {}
          return oSet.apply(this, arguments);
        };
        XHR.setRequestHeader.__cvrWrapped = true;
      }
    } catch (e) {}
  }
  function cvrPatchAll() {
    cvrPatch(window);
    try { if (window.top) cvrPatch(window.top); } catch (e) {}
    try {
      var f = window.top ? window.top.frames : null;
      if (f) { for (var i = 0; i < f.length; i++) { try { cvrPatch(f[i]); } catch (e) {} } }
    } catch (e) {}
  }
  cvrPatchAll();
  // Re-scan: frames (incl. the native camera add-in) may load after us.
  setInterval(cvrPatchAll, 2000);

  function $(id) { return document.getElementById(id); }

  function setPanelStatus(msg) { $('cvr-panel-status').textContent = msg || ''; }

  function setStatus(msg, kind) {
    var n = $('cvr-status');
    n.textContent = msg || '';
    n.className = 'cvr-status' + (kind ? ' is-' + kind : '');
  }

  // ---------- Session + Camera-Services ----------

  // getSession is a callback in older MyGeotab and a Promise in newer builds.
  // The callback's 2nd arg (and the promise's .path) is the real server host.
  function loadSession() {
    return new Promise(function (resolve, reject) {
      function accept(cred, server) {
        var c = (cred && cred.credentials) ? cred.credentials : cred;
        var srv = server || (cred && cred.path) || (cred && cred.server) || (c && c.domain) || (api && api.server) || null;
        try { console.log('CVR-DIAG ->', JSON.stringify({ server: srv || window.location.host, capturedBearer: !!capturedBearer })); } catch (e) {}
        if (!c || !c.sessionId) {
          reject(new Error('No MyGeotab session available.'));
          return;
        }
        session = c;
        if (srv) { sessionServer = srv; }
        resolve(c);
      }
      try {
        var ret = api.getSession(function (cred, server) { accept(cred, server); });
        if (ret && typeof ret.then === 'function') {
          ret.then(function (result) {
            accept(result, result && (result.path || result.server));
          }, reject);
        }
      } catch (e) {
        reject(e);
      }
    });
  }

  // service.api.call returns a Promise in map add-ins (older builds used
  // callbacks) - support both so we never hang.
  function apiCall(method, params) {
    return new Promise(function (resolve, reject) {
      try {
        var ret = api.call(method, params, function (r) { resolve(r); }, function (e) { reject(e); });
        if (ret && typeof ret.then === 'function') { ret.then(resolve, reject); }
      } catch (e) { reject(e); }
    });
  }

  function getUserId() {
    return apiCall('Get', { typeName: 'User', search: { name: session.userName } }).then(
      function (users) { if (users && users[0]) { userId = users[0].id; } return userId; },
      function () { return null; }
    );
  }

  // The camera UI's Bearer (OIDC) token is kept by MyGeotab in localStorage
  // under 'authToken'. We run in the same origin, so we can read + reuse it.
  // Read fresh each call so we pick up MyGeotab's token refreshes.
  function getAuthToken() {
    if (capturedBearer) { return capturedBearer; }
    try { return (window.localStorage && localStorage.getItem('authToken')) || null; }
    catch (e) { return null; }
  }

  // Give MyGeotab a moment to fire a media-services call we can capture.
  function waitForToken(timeoutMs) {
    var start = Date.now();
    return new Promise(function (resolve) {
      (function poll() {
        if (capturedBearer || Date.now() - start > timeoutMs) { resolve(capturedBearer); return; }
        setTimeout(poll, 300);
      })();
    });
  }

  function cameraHeaders() {
    var path = sessionServer || (window.location && window.location.host) || '';
    path = String(path).replace(/^https?:\/\//, '').replace(/\/$/, '');
    var h = {
      'X-MyGeotab-Database': session.database,
      'X-MyGeotab-Path': path,
      'X-MyGeotab-SessionId': session.sessionId,
      'X-MyGeotab-Username': session.userName,
      'Content-Type': 'application/json'
    };
    if (userId) { h['X-MyGeotab-Userid'] = userId; }
    var token = getAuthToken();
    if (token) { h['Authorization'] = 'Bearer ' + token; }
    return h;
  }

  function mediaFetch(method, path, body) {
    function go() {
      return fetch(MEDIA_BASE + path, {
        method: method,
        headers: cameraHeaders(),
        body: body ? JSON.stringify(body) : undefined
      });
    }
    return go().then(function (r) {
      if (r.status === 401) { return loadSession().then(go); }
      return r;
    }).then(function (r) {
      return r.text().then(function (txt) {
        var data = txt ? JSON.parse(txt) : null;
        if (!r.ok) {
          var msg = (data && (data.message || data.error)) || ('media-services ' + r.status);
          if (r.status === 403) { msg += ' (user lacks a camera role, e.g. ViewRecordedVideo).'; }
          throw new Error(msg);
        }
        return data;
      });
    });
  }

  // ---------- Data ----------

  function getDevices() {
    return apiCall('Get', { typeName: 'Device', resultsLimit: 5000 }).then(
      function (devs) { return devs || []; },
      function () { return []; }
    );
  }

  function loadCameras() {
    return Promise.all([mediaFetch('GET', '/DeviceMappings'), getDevices()]).then(function (res) {
      var mappings = res[0] || [];
      var devices = res[1] || [];
      var nameBySerial = {};
      serialByDeviceId = {};
      devices.forEach(function (d) {
        if (d.serialNumber) {
          nameBySerial[d.serialNumber] = d.name || d.serialNumber;
          serialByDeviceId[d.id] = d.serialNumber;
        }
      });

      cameras = mappings.map(function (m) {
        var goSerial = m.associatedDeviceSerialNumber || m.goDeviceSerialNumber || '';
        var camId = m.partnerDeviceId || m.cameraImei || m.recorderId || '';
        var partner = String(m.partnerId || m.partner || '').toLowerCase();
        var vehicle = nameBySerial[goSerial] || goSerial || '(unpaired)';
        return { vehicle: vehicle, partnerId: partner, partnerDeviceId: camId, goSerial: goSerial };
      }).filter(function (c) {
        return c.partnerDeviceId;
      }).sort(function (a, b) {
        return a.vehicle.localeCompare(b.vehicle);
      });

      cameraIdxByGoSerial = {};
      cameras.forEach(function (c, i) {
        if (c.goSerial) { cameraIdxByGoSerial[c.goSerial] = i; }
      });

      return cameras;
    });
  }

  function populateCameraSelect() {
    var sel = $('cvr-camera');
    sel.innerHTML = '';
    if (!cameras.length) {
      sel.innerHTML = '<option value="" disabled selected>No cameras found</option>';
      return;
    }
    var ph = document.createElement('option');
    ph.value = '';
    ph.disabled = true;
    ph.selected = true;
    ph.textContent = 'Select a camera...';
    sel.appendChild(ph);
    cameras.forEach(function (c, i) {
      var o = document.createElement('option');
      o.value = String(i);
      o.textContent = c.vehicle + '  (' + (c.partnerId || 'camera') + ' - ' + c.partnerDeviceId + ')';
      sel.appendChild(o);
    });
  }

  function preselectByDeviceId(deviceId) {
    if (!deviceId || !loaded) { return; }
    var serial = serialByDeviceId[deviceId];
    var idx = serial != null ? cameraIdxByGoSerial[serial] : undefined;
    if (idx == null) {
      $('cvr-panel-context').textContent = 'Selected vehicle has no paired camera.';
      $('cvr-open').disabled = cameras.length === 0;
      return;
    }
    var sel = $('cvr-camera');
    if (sel) { sel.value = String(idx); }
    $('cvr-panel-context').textContent = 'Ready: ' + cameras[idx].vehicle;
    $('cvr-open').disabled = false;
  }

  // ---------- Address search (jump map + drop pin) ----------

  function setFindStatus(msg) { var n = $('cvr-find-status'); if (n) { n.textContent = msg || ''; } }

  function goToLocation(lat, lng) {
    try {
      if (canvasSvc && canvasSvc.clear) { canvasSvc.clear(); }
      if (canvasSvc && canvasSvc.marker) { canvasSvc.marker({ lat: lat, lng: lng }, 30, 38, PIN_SRC, 1000); }
    } catch (e) {}
    try {
      if (mapSvc && mapSvc.setBounds) {
        var d = 0.0025; // ~250m box -> street-level zoom
        mapSvc.setBounds({ sw: { lat: lat - d, lng: lng - d }, ne: { lat: lat + d, lng: lng + d } });
      }
    } catch (e) {}
  }

  function findAddress() {
    var addr = ($('cvr-address').value || '').trim();
    if (!addr) { setFindStatus('Type an address.'); return; }
    setFindStatus('Searching...');
    // Geotab server-side geocoder: returns coordinates (x = lng, y = lat).
    apiCall('GetCoordinates', { addresses: [addr] }).then(function (res) {
      var c = res && res[0];
      var lat = c && (c.y != null ? c.y : c.lat);
      var lng = c && (c.x != null ? c.x : c.lng);
      if (lat == null || lng == null || (lat === 0 && lng === 0)) {
        setFindStatus('Address not found. Try adding city/state.');
        return;
      }
      goToLocation(lat, lng);
      setFindStatus('Showing: ' + addr + '  (pick the truck on the map, then request video)');
    }).catch(function (err) {
      setFindStatus('Search failed: ' + (err && err.message ? err.message : err));
    });
  }

  // ---------- Request window ----------

  function buildWindow() {
    var startLocal = $('cvr-start').value;
    var durSec = parseInt($('cvr-duration').value, 10);
    if (!startLocal) { throw new Error('Pick an event start time.'); }
    if (!durSec || durSec > MAX_DURATION_SECONDS) {
      throw new Error('Duration must be 1-' + MAX_DURATION_SECONDS + ' seconds.');
    }
    var startDate = new Date(startLocal);
    if (isNaN(startDate.getTime())) { throw new Error('Invalid start time.'); }
    var endDate = new Date(startDate.getTime() + durSec * 1000);
    return { startISO: startDate.toISOString(), endISO: endDate.toISOString(), durSec: durSec };
  }

  function updateWindowHint() {
    try {
      var w = buildWindow();
      $('cvr-window-hint').textContent = 'UTC ' + w.startISO + ' to ' + w.endISO;
    } catch (e) {
      $('cvr-window-hint').textContent = '';
    }
  }

  function setDefaultStart() {
    var now = new Date();
    now.setMilliseconds(0);
    var tzOffsetMs = now.getTimezoneOffset() * 60000;
    $('cvr-start').value = new Date(now.getTime() - tzOffsetMs).toISOString().slice(0, 19);
  }

  // ---------- Modal ----------

  function openModal() {
    if (!loaded) { return; }
    setStatus('');
    setDefaultStart();
    updateWindowHint();
    $('cvr-modal').hidden = false;
  }

  function closeModal() {
    $('cvr-modal').hidden = true;
  }

  function submitRequest() {
    var idx = $('cvr-camera').value;
    if (idx === '') { setStatus('Select a camera first.', 'error'); return; }
    var cam = cameras[parseInt(idx, 10)];

    var w;
    try {
      w = buildWindow();
    } catch (e) {
      setStatus(e.message, 'error');
      return;
    }

    var payload = {
      requestStartTime: w.startISO,
      requestEndTime: w.endISO,
      mediaResourceType: 'Video',
      partnerId: cam.partnerId,
      partnerDeviceId: cam.partnerDeviceId,
      goDeviceSerialNumber: cam.goSerial
    };

    $('cvr-submit').disabled = true;
    setStatus('Submitting request...', 'busy');

    mediaFetch('POST', '/Media', payload).then(function (resp) {
      var reqId = (resp && (resp.requestId || resp.id)) || '(queued)';
      setStatus('Requested for ' + cam.vehicle + '. Tracking ID: ' + reqId + '. Footage appears in the Cameras/Video view once uploaded.', 'ok');
    }).catch(function (err) {
      setStatus('Request failed: ' + (err && err.message ? err.message : err), 'error');
    }).then(function () {
      $('cvr-submit').disabled = false;
    });
  }

  // ---------- Wire up ----------

  function wireControls() {
    $('cvr-open').addEventListener('click', openModal);
    $('cvr-x').addEventListener('click', closeModal);
    $('cvr-cancel').addEventListener('click', closeModal);
    $('cvr-submit').addEventListener('click', submitRequest);
    $('cvr-start').addEventListener('input', updateWindowHint);
    $('cvr-duration').addEventListener('change', updateWindowHint);
    $('cvr-modal').addEventListener('click', function (e) {
      if (e.target === $('cvr-modal')) { closeModal(); }
    });
    var findBtn = $('cvr-find');
    if (findBtn) { findBtn.addEventListener('click', findAddress); }
    var addrInput = $('cvr-address');
    if (addrInput) {
      addrInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); findAddress(); }
      });
    }
  }

  function init() {
    if (loaded) { return; }
    setPanelStatus('Loading cameras...');
    loadSession().then(getUserId).then(function () { return waitForToken(8000); }).then(loadCameras).then(function () {
      populateCameraSelect();
      loaded = true;
      setPanelStatus(cameras.length + ' camera(s) available.');
      $('cvr-open').disabled = cameras.length === 0;
    }).catch(function (err) {
      setPanelStatus('Could not load cameras: ' + (err && err.message ? err.message : err) + '  [path=' + (sessionServer || (window.location && window.location.host)) + ']');
    });
  }

  function attachMapEvent(name) {
    try {
      service.events.attach(name, function (data) {
        if (data && data.type === 'device' && data.entity) {
          preselectByDeviceId(data.entity.id);
        }
      });
    } catch (e) {
      // event not available on this page/version - safe to ignore
    }
  }

  attachMapEvent('over');
  attachMapEvent('click');

  try {
    service.page.attach('focus', init);
  } catch (e) {
    // optional
  }

  wireControls();
  init();
};
