/* global geotab */
/**
 * MyGeotab Add-In: Custom Video Request
 *
 * Recreates the Surfsight portal's "Request custom video recording" experience
 * inside MyGeotab — a clean start-time + duration picker — for both
 * GoFocus (Smarter AI) and Surfsight cameras.
 *
 * No proxy. The add-in runs in the my.geotab.com origin (same origin the native
 * Cameras view uses), so it calls Camera-Services (media-services.geotab.com)
 * directly, authenticating with the four X-MyGeotab-* headers built from the
 * live session (api.getSession()). The credentials are already handled by
 * MyGeotab — we just forward the session.
 *
 *   List cameras:   GET  /DeviceMappings
 *   Request clip:   POST /Media
 *
 * partnerId branching (handled automatically from /DeviceMappings):
 *   GoFocus / GO Focus Plus -> "smarterai", partnerDeviceId = camera serial
 *   Surfsight               -> "surfsight", partnerDeviceId = camera IMEI
 *   Sensata                 -> "sensata",   partnerDeviceId = recorderId
 */
geotab.addin.customVideoRequest = function () {
  'use strict';

  // ---------- Constants ----------
  var MEDIA_BASE = 'https://media-services.geotab.com';
  var MAX_DURATION_SECONDS = 120; // guard against accidental large/expensive pulls

  // ---------- State ----------
  var api;
  var elAddin;
  var session = null;     // { database, userName, sessionId, ... } from getSession
  var cameras = [];       // normalized: { label, vehicle, partnerId, partnerDeviceId, goSerial }
  var recent = [];        // in-memory log for the table this session

  // ---------- Helpers ----------
  function $(id) { return document.getElementById(id); }

  function setStatus(msg, kind) {
    var node = $('cvr-status');
    node.textContent = msg || '';
    node.className = 'cvr-status' + (kind ? ' is-' + kind : '');
  }

  function escapeHtml(s) {
    if (s == null) return '';
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // ---------- Session + Camera-Services access ----------

  // Pull the active MyGeotab session so we can build the X-MyGeotab-* headers.
  function loadSession() {
    return new Promise(function (resolve, reject) {
      api.getSession(function (s) {
        if (!s) { reject(new Error('No MyGeotab session available.')); return; }
        session = s;
        resolve(s);
      });
    });
  }

  // The add-in document IS my.geotab.com, so the host is just window.location.
  function cameraHeaders() {
    var path = (window.location && window.location.host) || (session && session.path) || '';
    path = String(path).replace(/^https?:\/\//, '').replace(/\/$/, '');
    return {
      'X-MyGeotab-Database': session.database,
      'X-MyGeotab-Path': path,
      'X-MyGeotab-SessionId': session.sessionId,
      'X-MyGeotab-Username': session.userName,
      'Content-Type': 'application/json'
    };
  }

  // Direct call to media-services with one automatic re-fetch of the session on 401.
  function mediaFetch(method, path, body) {
    function doFetch() {
      return fetch(MEDIA_BASE + path, {
        method: method,
        headers: cameraHeaders(),
        body: body ? JSON.stringify(body) : undefined
      });
    }
    return doFetch().then(function (r) {
      if (r.status === 401) {
        // Session aged out — refresh it from MyGeotab and retry once.
        return loadSession().then(doFetch);
      }
      return r;
    }).then(function (r) {
      return r.text().then(function (txt) {
        var data = txt ? JSON.parse(txt) : null;
        if (!r.ok) {
          var msg = (data && (data.message || data.error)) || ('media-services ' + r.status);
          if (r.status === 403) msg += ' (session valid but user lacks the camera role, e.g. ViewRecordedVideo).';
          throw new Error(msg);
        }
        return data;
      });
    });
  }

  // ---------- Data load ----------

  // Camera↔vehicle pairings from /DeviceMappings, joined to Device names.
  function loadCameras() {
    var mappingsP = mediaFetch('GET', '/DeviceMappings');
    var devicesP = new Promise(function (resolve) {
      api.call('Get', { typeName: 'Device', resultsLimit: 5000 },
        function (devs) { resolve(devs || []); },
        function () { resolve([]); }); // friendly names are a nicety; don't fail on them
    });

    return Promise.all([mappingsP, devicesP]).then(function (res) {
      var mappings = res[0] || [];
      var devices = res[1] || [];
      var nameBySerial = {};
      devices.forEach(function (d) {
        if (d.serialNumber) nameBySerial[d.serialNumber] = d.name || d.serialNumber;
      });

      cameras = (mappings || []).map(function (m) {
        var goSerial = m.associatedDeviceSerialNumber || m.goDeviceSerialNumber || '';
        var camId    = m.partnerDeviceId || m.cameraImei || m.recorderId || '';
        var partner  = String(m.partnerId || m.partner || '').toLowerCase();
        var vehicle  = nameBySerial[goSerial] || goSerial || '(unpaired)';
        return {
          vehicle: vehicle,
          partnerId: partner,
          partnerDeviceId: camId,
          goSerial: goSerial,
          label: vehicle + '  —  ' + (partner || 'camera') + ' · ' + camId
        };
      }).filter(function (c) { return c.partnerDeviceId; })
        .sort(function (a, b) { return a.vehicle.localeCompare(b.vehicle); });

      return cameras;
    });
  }

  function populateCameraSelect() {
    var sel = $('cvr-camera');
    sel.innerHTML = '';
    if (!cameras.length) {
      sel.innerHTML = '<option value="" disabled selected>No cameras found</option>';
      $('cvr-camera-hint').textContent = 'No paired cameras returned for this database.';
      return;
    }
    var ph = document.createElement('option');
    ph.value = ''; ph.disabled = true; ph.selected = true;
    ph.textContent = 'Select a camera…';
    sel.appendChild(ph);
    cameras.forEach(function (c, i) {
      var o = document.createElement('option');
      o.value = String(i);
      o.textContent = c.label;
      sel.appendChild(o);
    });
    $('cvr-camera-hint').textContent = cameras.length + ' camera(s) available.';
  }

  // ---------- Request window ----------

  function buildWindow() {
    var startLocal = $('cvr-start').value;          // "YYYY-MM-DDTHH:mm:ss" (local)
    var durSec = parseInt($('cvr-duration').value, 10);
    if (!startLocal) throw new Error('Pick an event start time.');
    if (!durSec || durSec > MAX_DURATION_SECONDS) {
      throw new Error('Duration must be 1–' + MAX_DURATION_SECONDS + ' seconds.');
    }
    var startDate = new Date(startLocal);
    if (isNaN(startDate.getTime())) throw new Error('Invalid start time.');
    var endDate = new Date(startDate.getTime() + durSec * 1000);
    return {
      startISO: startDate.toISOString(),
      endISO: endDate.toISOString(),
      startLocalLabel: startDate.toLocaleString(),
      durSec: durSec
    };
  }

  function updateWindowHint() {
    try {
      var w = buildWindow();
      $('cvr-window-hint').textContent =
        'Window: ' + w.startLocalLabel + ' for ' + w.durSec + 's' +
        '  (UTC ' + w.startISO + ' → ' + w.endISO + ')';
    } catch (e) {
      $('cvr-window-hint').textContent = '';
    }
  }

  function addRecent(vehicle, windowLabel, statusText) {
    recent.unshift({ vehicle: vehicle, window: windowLabel, status: statusText });
    var body = $('cvr-recent-body');
    body.innerHTML = '';
    recent.slice(0, 10).forEach(function (r) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td>' + escapeHtml(r.vehicle) + '</td><td>' +
        escapeHtml(r.window) + '</td><td>' + escapeHtml(r.status) + '</td>';
      body.appendChild(tr);
    });
    $('cvr-recent-section').hidden = false;
  }

  // ---------- Submit ----------

  function submitRequest(ev) {
    ev.preventDefault();
    var idx = $('cvr-camera').value;
    if (idx === '') { setStatus('Select a camera first.', 'error'); return; }
    var cam = cameras[parseInt(idx, 10)];

    var w;
    try { w = buildWindow(); }
    catch (e) { setStatus(e.message, 'error'); return; }

    // media-services POST /Media body.
    var payload = {
      requestStartTime: w.startISO,
      requestEndTime: w.endISO,
      mediaResourceType: 'Video',
      partnerId: cam.partnerId,             // smarterai (GoFocus) | surfsight | sensata
      partnerDeviceId: cam.partnerDeviceId, // serial | IMEI | recorderId per partner
      goDeviceSerialNumber: cam.goSerial
    };

    $('cvr-submit').disabled = true;
    setStatus('Submitting request…', 'busy');

    mediaFetch('POST', '/Media', payload).then(function (resp) {
      var reqId = (resp && (resp.requestId || resp.id)) || '(queued)';
      setStatus('Request submitted. Tracking ID: ' + reqId +
        '. Footage will appear in the Video/Cameras view once the camera uploads it.', 'ok');
      addRecent(cam.vehicle, w.startLocalLabel + ' · ' + w.durSec + 's', 'Queued');
    }).catch(function (err) {
      setStatus('Request failed: ' + (err && err.message ? err.message : err), 'error');
      addRecent(cam.vehicle, w.startLocalLabel + ' · ' + w.durSec + 's', 'Failed');
    }).then(function () {
      $('cvr-submit').disabled = false;
    });
  }

  function resetForm() {
    $('cvr-form').reset();
    $('cvr-duration').value = '30';
    setDefaultStart();
    setStatus('');
    updateWindowHint();
  }

  function setDefaultStart() {
    var now = new Date();
    now.setMilliseconds(0);
    var tzOffsetMs = now.getTimezoneOffset() * 60000;
    $('cvr-start').value = new Date(now.getTime() - tzOffsetMs).toISOString().slice(0, 19);
  }

  // ---------- Geotab add-in lifecycle ----------
  return {
    initialize: function (freshApi, freshState, initializeCallback) {
      api = freshApi;
      elAddin = $('customVideoRequest');

      $('cvr-form').addEventListener('submit', submitRequest);
      $('cvr-cancel').addEventListener('click', resetForm);
      $('cvr-start').addEventListener('input', updateWindowHint);
      $('cvr-duration').addEventListener('change', updateWindowHint);

      setDefaultStart();
      updateWindowHint();
      initializeCallback();
    },

    focus: function (freshApi) {
      api = freshApi;
      elAddin.className = '';
      setStatus('Loading cameras…', 'busy');
      $('cvr-submit').disabled = true;

      loadSession()
        .then(loadCameras)
        .then(function () {
          populateCameraSelect();
          setStatus('');
          $('cvr-submit').disabled = false;
        })
        .catch(function (err) {
          setStatus('Could not load cameras: ' + (err && err.message ? err.message : err), 'error');
        });
    },

    blur: function () {
      elAddin.className = 'hidden';
    }
  };
};
