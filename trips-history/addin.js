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

  // ---------- Incident pull (chunked clips + ZIP download) ----------

  var incidentClips = [];

  function escHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function fillIncidentCameras() {
    var sel = $('cvr-inc-camera');
    sel.innerHTML = '';
    var ph = document.createElement('option');
    ph.value = ''; ph.disabled = true; ph.selected = true; ph.textContent = 'Select a camera...';
    sel.appendChild(ph);
    cameras.forEach(function (c, i) {
      var o = document.createElement('option');
      o.value = String(i);
      o.textContent = c.vehicle + '  (' + (c.partnerId || 'camera') + ')';
      sel.appendChild(o);
    });
  }

  function incStatus(msg, kind) {
    var n = $('cvr-inc-status');
    n.textContent = msg || '';
    n.className = 'cvr-status' + (kind ? ' is-' + kind : '');
  }

  function updateIncPlan() {
    var total = parseInt($('cvr-inc-total').value, 10);
    var chunk = parseInt($('cvr-inc-chunk').value, 10);
    if (!total || !chunk) { $('cvr-inc-plan').textContent = ''; return; }
    var n = Math.ceil((total * 60) / chunk);
    $('cvr-inc-plan').textContent = 'Will request ' + n + ' clip(s) of ' + (chunk / 60) + ' min to cover ' + total + ' min.';
  }

  function openIncident() {
    if (!loaded) { return; }
    fillIncidentCameras();
    var now = new Date(); now.setMilliseconds(0);
    $('cvr-inc-start').value = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 19);
    $('cvr-inc-total').value = '20';
    $('cvr-inc-chunk').value = '120';
    incidentClips = [];
    $('cvr-inc-list').innerHTML = '';
    $('cvr-inc-listwrap').hidden = true;
    $('cvr-inc-zip').hidden = true;
    $('cvr-inc-start-btn').disabled = false;
    incStatus('');
    updateIncPlan();
    $('cvr-incident-modal').hidden = false;
  }

  function closeIncident() { $('cvr-incident-modal').hidden = true; }

  function renderIncList() {
    var body = $('cvr-inc-list');
    body.innerHTML = '';
    incidentClips.forEach(function (c) {
      var tr = document.createElement('tr');
      var dl = c.url ? '<a href="' + escHtml(c.url) + '" target="_blank" rel="noopener">download</a>' : '';
      tr.innerHTML = '<td>' + c.idx + '</td><td>' + escHtml(c.label) + '</td><td>' + escHtml(c.status) + '</td><td>' + dl + '</td>';
      body.appendChild(tr);
    });
  }

  function startIncidentPull() {
    var idxStr = $('cvr-inc-camera').value;
    if (idxStr === '') { incStatus('Select a camera first.', 'error'); return; }
    var cam = cameras[parseInt(idxStr, 10)];
    var startLocal = $('cvr-inc-start').value;
    var total = parseInt($('cvr-inc-total').value, 10);
    var chunk = parseInt($('cvr-inc-chunk').value, 10);
    if (!startLocal) { incStatus('Set the incident start time.', 'error'); return; }
    if (!total || total < 1 || total > 120) { incStatus('Total minutes must be 1-120.', 'error'); return; }
    var startMs = new Date(startLocal).getTime();
    if (isNaN(startMs)) { incStatus('Invalid start time.', 'error'); return; }
    var nChunks = Math.ceil((total * 60) / chunk);

    incidentClips = [];
    for (var i = 0; i < nChunks; i++) {
      var s = new Date(startMs + i * chunk * 1000);
      var e = new Date(Math.min(startMs + (i + 1) * chunk * 1000, startMs + total * 60 * 1000));
      incidentClips.push({
        idx: i + 1, startISO: s.toISOString(), endISO: e.toISOString(),
        label: s.toLocaleTimeString() + ' - ' + e.toLocaleTimeString(),
        requestId: null, status: 'Requesting...', url: null
      });
    }
    renderIncList();
    $('cvr-inc-listwrap').hidden = false;
    $('cvr-inc-start-btn').disabled = true;
    $('cvr-inc-zip').hidden = true;
    incStatus('Requesting ' + nChunks + ' clip(s)...', 'busy');

    incidentClips.forEach(function (clip) {
      var payload = {
        requestStartTime: clip.startISO, requestEndTime: clip.endISO, mediaResourceType: 'Video',
        partnerId: cam.partnerId, partnerDeviceId: cam.partnerDeviceId, goDeviceSerialNumber: cam.goSerial
      };
      mediaFetch('POST', '/Media', payload).then(function (resp) {
        clip.requestId = (resp && (resp.mediaRequestId || resp.requestId || resp.id)) || null;
        clip.status = clip.requestId ? 'Processing...' : 'Queued';
        renderIncList();
        if (clip.requestId) { pollClip(clip, 0); } else { checkAllDone(); }
      }).catch(function (err) {
        clip.status = 'Failed: ' + (err && err.message ? err.message : err);
        renderIncList(); checkAllDone();
      });
    });
  }

  function pollClip(clip, attempts) {
    if (attempts > 75) { clip.status = 'Timed out'; renderIncList(); checkAllDone(); return; } // ~5 min
    mediaFetch('GET', '/Media/' + encodeURIComponent(clip.requestId)).then(function (r) {
      var st = r && r.status;
      if (st === 'ResponseReady' || st === 'ResponsePartiallyReady') {
        fetchClipResource(clip);
      } else if (st === 'RequestFailed' || st === 'RequestTimedOut' || st === 'DeviceUnavailable' || st === 'QueueOverflow') {
        clip.status = 'Failed (' + st + ')'; renderIncList(); checkAllDone();
      } else {
        clip.status = 'Processing...'; renderIncList();
        setTimeout(function () { pollClip(clip, attempts + 1); }, 4000);
      }
    }).catch(function () {
      setTimeout(function () { pollClip(clip, attempts + 1); }, 4000);
    });
  }

  function fetchClipResource(clip) {
    mediaFetch('GET', '/Media/' + encodeURIComponent(clip.requestId) + '/Resources').then(function (r) {
      var resources = (r && r.mediaResources) || [];
      var res = null;
      for (var i = 0; i < resources.length; i++) { if (resources[i] && resources[i].mediaUrl) { res = resources[i]; break; } }
      clip.url = res ? res.mediaUrl : null;
      clip.status = res ? 'Ready' : 'Ready (no file URL)';
      renderIncList(); checkAllDone();
    }).catch(function () {
      clip.status = 'Ready (link error)'; renderIncList(); checkAllDone();
    });
  }

  function checkAllDone() {
    var pending = incidentClips.some(function (c) { return /Requesting|Processing/.test(c.status); });
    if (pending) { return; }
    var ready = incidentClips.filter(function (c) { return c.url; });
    $('cvr-inc-start-btn').disabled = false;
    if (ready.length) {
      $('cvr-inc-zip').hidden = false;
      incStatus(ready.length + ' of ' + incidentClips.length + ' clip(s) ready. Download all, or use the per-clip links.', 'ok');
    } else {
      incStatus('No clips retrieved. Check the camera was online for that window, then retry.', 'error');
    }
  }

  // ---- dependency-free ZIP (store / no compression - clips are already compressed) ----
  var _cvrCrcTable = null;
  function cvrCrcTable() {
    if (_cvrCrcTable) { return _cvrCrcTable; }
    var t = [];
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) { c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); }
      t[n] = c >>> 0;
    }
    _cvrCrcTable = t; return t;
  }
  function cvrCrc32(u8) {
    var t = cvrCrcTable(), crc = -1;
    for (var i = 0; i < u8.length; i++) { crc = (crc >>> 8) ^ t[(crc ^ u8[i]) & 0xFF]; }
    return (crc ^ -1) >>> 0;
  }
  function _u16(n) { return new Uint8Array([n & 255, (n >> 8) & 255]); }
  function _u32(n) { return new Uint8Array([n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]); }
  function _cat(arr) {
    var len = 0, i; for (i = 0; i < arr.length; i++) { len += arr[i].length; }
    var out = new Uint8Array(len), o = 0;
    for (i = 0; i < arr.length; i++) { out.set(arr[i], o); o += arr[i].length; }
    return out;
  }
  function cvrBuildZip(files) {
    var enc = new TextEncoder();
    var parts = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name);
      var crc = cvrCrc32(f.data);
      var size = f.data.length;
      var lh = _cat([_u32(0x04034b50), _u16(20), _u16(0), _u16(0), _u16(0), _u16(0x21),
        _u32(crc), _u32(size), _u32(size), _u16(name.length), _u16(0)]);
      parts.push(lh, name, f.data);
      var ch = _cat([_u32(0x02014b50), _u16(20), _u16(20), _u16(0), _u16(0), _u16(0), _u16(0x21),
        _u32(crc), _u32(size), _u32(size), _u16(name.length), _u16(0), _u16(0), _u16(0), _u16(0),
        _u32(0), _u32(offset)]);
      central.push(ch, name);
      offset += lh.length + name.length + size;
    });
    var cd = _cat(central);
    var eocd = _cat([_u32(0x06054b50), _u16(0), _u16(0), _u16(files.length), _u16(files.length),
      _u32(cd.length), _u32(offset), _u16(0)]);
    var all = parts.concat([cd, eocd]);
    return new Blob(all, { type: 'application/zip' });
  }

  function _pad(n) { return (n < 10 ? '0' : '') + n; }
  function _safe(s) { return String(s).replace(/[^a-z0-9]+/gi, '_').replace(/^_|_$/g, '').slice(0, 28); }

  function downloadIncidentZip() {
    var ready = incidentClips.filter(function (c) { return c.url; });
    if (!ready.length) { incStatus('Nothing ready to download.', 'error'); return; }
    incStatus('Fetching ' + ready.length + ' clip(s) to bundle...', 'busy');
    $('cvr-inc-zip').disabled = true;
    var files = [], failed = 0;
    var jobs = ready.map(function (c) {
      return fetch(c.url).then(function (r) {
        if (!r.ok) { throw new Error('http ' + r.status); }
        return r.arrayBuffer();
      }).then(function (buf) {
        files.push({ name: 'clip_' + _pad(c.idx) + '_' + _safe(c.label) + '.mp4', data: new Uint8Array(buf) });
      }).catch(function () { failed++; });
    });
    Promise.all(jobs).then(function () {
      $('cvr-inc-zip').disabled = false;
      if (!files.length) {
        incStatus('Your browser is blocked from downloading the clip files directly (CORS). Use the per-clip "download" links in the list instead.', 'error');
        return;
      }
      try {
        var blob = cvrBuildZip(files);
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url; a.download = 'incident_clips.zip';
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
        incStatus('Downloaded ' + files.length + ' clip(s) as ZIP' + (failed ? (' (' + failed + ' could not be fetched - use links)') : '') + '.', failed ? 'error' : 'ok');
      } catch (e) {
        incStatus('Could not build ZIP: ' + (e && e.message ? e.message : e) + '. Use the per-clip links.', 'error');
      }
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
    $('cvr-incident-open').addEventListener('click', openIncident);
    $('cvr-inc-x').addEventListener('click', closeIncident);
    $('cvr-inc-cancel').addEventListener('click', closeIncident);
    $('cvr-inc-start-btn').addEventListener('click', startIncidentPull);
    $('cvr-inc-zip').addEventListener('click', downloadIncidentZip);
    $('cvr-inc-total').addEventListener('input', updateIncPlan);
    $('cvr-inc-chunk').addEventListener('change', updateIncPlan);
    $('cvr-incident-modal').addEventListener('click', function (e) {
      if (e.target === $('cvr-incident-modal')) { closeIncident(); }
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
      $('cvr-incident-open').disabled = cameras.length === 0;
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
