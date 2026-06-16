/**
 * Custom Video Request - MyGeotab Map Add-In (Trips History page)
 *
 * Adds a "Request custom video recording" launcher to the Trips History side
 * panel. Clicking opens a modal to pick a camera, an event start time, and a
 * duration, then queues the clip.
 *
 * Runs in the my.geotab.com document context (map add-ins share the geotab
 * global), so it calls Camera-Services (media-services.geotab.com) directly
 * with the four X-MyGeotab-* headers from the live session - no proxy.
 *
 *   List cameras:  GET  /DeviceMappings
 *   Request clip:  POST /Media
 *
 * partnerId branches automatically: smarterai (GoFocus) | surfsight | sensata.
 */
geotab.addin.request = function (elt, service) {
  'use strict';

  var MEDIA_BASE = 'https://media-services.geotab.com';
  var MAX_DURATION_SECONDS = 120;

  var api = service.api;
  var session = null;
  var cameras = [];
  var serialByDeviceId = {};
  var cameraIdxByGoSerial = {};
  var loaded = false;

  function $(id) { return document.getElementById(id); }

  function setPanelStatus(msg) { $('cvr-panel-status').textContent = msg || ''; }

  function setStatus(msg, kind) {
    var n = $('cvr-status');
    n.textContent = msg || '';
    n.className = 'cvr-status' + (kind ? ' is-' + kind : '');
  }

  // ---------- Session + Camera-Services ----------

  // getSession is a callback in older MyGeotab and a Promise in newer builds.
  function loadSession() {
    return new Promise(function (resolve, reject) {
      function accept(result) {
        var cred = (result && result.credentials) ? result.credentials : result;
        if (!cred || !cred.sessionId) {
          reject(new Error('No MyGeotab session available.'));
          return;
        }
        if (result && result.path && !cred.path) { cred.path = result.path; }
        if (!session) {
          session = cred;
          resolve(cred);
        }
      }
      try {
        var ret = api.getSession(function (s) { accept(s); });
        if (ret && typeof ret.then === 'function') { ret.then(accept, reject); }
      } catch (e) {
        reject(e);
      }
    });
  }

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
    return new Promise(function (resolve) {
      api.call('Get', { typeName: 'Device', resultsLimit: 5000 }, function (devs) {
        resolve(devs || []);
      }, function () {
        resolve([]);
      });
    });
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
  }

  function init() {
    if (loaded) { return; }
    setPanelStatus('Loading cameras...');
    loadSession().then(loadCameras).then(function () {
      populateCameraSelect();
      loaded = true;
      setPanelStatus(cameras.length + ' camera(s) available.');
      $('cvr-open').disabled = cameras.length === 0;
    }).catch(function (err) {
      setPanelStatus('Could not load cameras: ' + (err && err.message ? err.message : err));
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
