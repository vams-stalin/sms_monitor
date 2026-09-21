(function () {

    var _role     = window.CURRENT_ROLE || 'viewer';
    var _username = (window.CURRENT_USER || 'User').trim();
    var _isAdmin  = _role === 'admin';

    // Role-based nav visibility
    ['trends-nav', 'settings-nav', 'audit-log-nav', 'usage-report-nav', 'master-user-group'].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.style.display = _isAdmin ? '' : 'none';
    });


    // Keep the current page highlighted in the same style as Dashboard.
    var _path = window.location.pathname.replace(/\/$/, '') || '/';
    document.querySelectorAll('.sidebar-nav a.nav-item[href]').forEach(function (item) {
        var href = item.getAttribute('href');
        if (href && href !== '#' && href.replace(/\/$/, '') === _path) {
            item.classList.add('active');
        } else if (href && href !== '#') {
            item.classList.remove('active');
        }
    });

    // Sidebar user info
    var sn = document.getElementById('sidebar-username');
    var sr = document.getElementById('sidebar-userrole');
    var av = document.getElementById('sidebar-avatar');
    if (sn) sn.textContent = _username;
    if (sr) sr.textContent = _isAdmin ? 'Admin' : 'Viewer';
    if (av) av.textContent = _username.charAt(0).toUpperCase();

    // Master User collapsible
    var _dropdownLoaded = false;

    window.toggleMasterUser = function () {
        var panel = document.getElementById('master-panel');
        var btn   = document.getElementById('master-user-toggle');
        if (!panel || !btn) return;
        var isOpen = panel.classList.contains('open');
        panel.classList.toggle('open', !isOpen);
        btn.classList.toggle('open', !isOpen);
        if (!isOpen && !_dropdownLoaded) {
            _populateMasterDropdown();
            _dropdownLoaded = true;
        }
    };

    function _populateMasterDropdown() {
        var sel = document.getElementById('mf-client');
        if (!sel) return;
        sel.innerHTML = '<option value="">Loading clients…</option>';
        sel.disabled  = true;
        fetch('/api/clients')
            .then(function (r) {
                if (r.status === 401 || r.status === 403) { window.location.href = '/'; return null; }
                return r.json();
            })
            .then(function (data) {
                if (!data) return;
                if (data.status === 'ok' && data.clients.length > 0) {
                    sel.innerHTML = '<option value="">— Select client —</option>';
                    data.clients.forEach(function (c) {
                        var opt       = document.createElement('option');
                        opt.value     = c.id;
                        opt.textContent = c.name;
                        sel.appendChild(opt);
                    });
                    sel.disabled = false;
                } else {
                    sel.innerHTML = '<option value="">No clients found</option>';
                }
            })
            .catch(function () {
                sel.innerHTML = '<option value="">Failed to load clients</option>';
            });
    }

    window.assignCredits = async function () {
        var sel      = document.getElementById('mf-client');
        var countIn  = document.getElementById('mf-sms-count');
        var startIn  = document.getElementById('mf-start-date');
        var endIn    = document.getElementById('mf-end-date');
        var feedback = document.getElementById('mf-feedback');
        var btn      = document.getElementById('mf-submit-btn');

        var clientId  = sel.value.trim();
        var smsCount  = parseInt(countIn.value, 10);
        var startDate = startIn ? startIn.value.trim() : '';
        var endDate   = endIn ? endIn.value.trim() : '';

        feedback.className   = 'mf-feedback';
        feedback.textContent = '';

        if (!clientId) {
            feedback.textContent = 'Please select a client.';
            feedback.classList.add('error');
            return;
        }
        if (!smsCount || smsCount < 1) {
            feedback.textContent = 'Enter a valid SMS count.';
            feedback.classList.add('error');
            return;
        }
        if ((startDate || endDate) && !(startDate && endDate)) {
            feedback.textContent = 'Please provide both Start Date and End Date, or leave both blank.';
            feedback.classList.add('error');
            return;
        }
        if (startDate && endDate && startDate >= endDate) {
            feedback.textContent = 'End date must be after start date.';
            feedback.classList.add('error');
            return;
        }

        btn.disabled         = true;
        feedback.textContent = 'Saving…';
        feedback.classList.add('success');

        var payload = { client_id: clientId, allocated_sms: smsCount };
        if (startDate && endDate) {
            payload.start_date = startDate;
            payload.end_date   = endDate;
        }

        try {
            var res  = await fetch('/api/assign-credits', {
                method:  'POST',
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify(payload)
            });
            if (res.status === 401 || res.status === 403) { window.location.href = '/'; return; }
            var data = await res.json();
            if (data.status === 'ok') {
                feedback.textContent = 'Credits assigned successfully!';
                feedback.className   = 'mf-feedback success';
                countIn.value = '';
                if (startIn) startIn.value = '';
                if (endIn) endIn.value = '';
                sel.value     = '';
                if (window.showToast) showToast('SMS credits assigned successfully.', 'success');
            } else {
                feedback.textContent = data.message || 'Failed to assign credits.';
                feedback.className   = 'mf-feedback error';
                if (window.showToast) showToast(data.message || 'Failed to assign credits.', 'error');
            }
        } catch (_) {
            feedback.textContent = 'Request failed. Check your connection.';
            feedback.className   = 'mf-feedback error';
            if (window.showToast) showToast('Network error. Check your connection.', 'error');
        }
        btn.disabled = false;
    };

    window.toggleCreditForm = function () {
        var panel = document.getElementById('credit-form-panel');
        var btn   = document.getElementById('credit-usage-toggle');
        if (!panel || !btn) return;
        var isOpen = panel.style.display !== 'none';
        panel.style.display = isOpen ? 'none' : '';
        btn.classList.toggle('open', !isOpen);
    };

    // Settings modal
    window.openSettings = async function () {
        var overlay   = document.getElementById('st-overlay');
        var warnInput = document.getElementById('st-warning');
        var critInput = document.getElementById('st-critical');
        var fb        = document.getElementById('st-feedback');
        if (!overlay) return;
        fb.textContent = '';
        fb.className   = 'st-feedback';
        try {
            var res  = await fetch('/api/thresholds');
            var data = await res.json();
            warnInput.value = data.warning;
            critInput.value = data.critical;
        } catch (_) {}
        overlay.style.display = 'flex';
        warnInput.focus();
    };

    window.closeSettings = function () {
        var overlay = document.getElementById('st-overlay');
        if (overlay) overlay.style.display = 'none';
    };

    window.saveSettings = async function () {
        var btn      = document.getElementById('st-save-btn');
        var fb       = document.getElementById('st-feedback');
        var warning  = parseInt(document.getElementById('st-warning').value,  10);
        var critical = parseInt(document.getElementById('st-critical').value, 10);

        fb.textContent = '';
        fb.className   = 'st-feedback';

        if (isNaN(warning) || isNaN(critical)) {
            fb.textContent = 'Please enter valid numbers for both fields.';
            fb.className   = 'st-feedback error';
            return;
        }
        if (warning <= 0 || critical <= 0 || warning >= 100 || critical >= 100) {
            fb.textContent = 'Values must be between 1 and 99.';
            fb.className   = 'st-feedback error';
            return;
        }
        if (warning >= critical) {
            fb.textContent = 'Warning must be less than Critical.';
            fb.className   = 'st-feedback error';
            return;
        }

        btn.disabled    = true;
        btn.textContent = 'Saving…';
        try {
            var res = await fetch('/api/thresholds', {
                method:  'POST',
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify({ warning: warning, critical: critical })
            });
            if (res.status === 401 || res.status === 403) { window.location.href = '/'; return; }
            var data = await res.json();
            if (data.status === 'ok') {
                fb.textContent = 'Saved! Thresholds updated successfully.';
                fb.className   = 'st-feedback ok';
                if (window.showToast) showToast('Alert thresholds updated.', 'success');
                setTimeout(window.closeSettings, 1200);
            } else {
                fb.textContent = data.message || 'Failed to save.';
                fb.className   = 'st-feedback error';
            }
        } catch (_) {
            fb.textContent = 'Request failed. Check your connection.';
            fb.className   = 'st-feedback error';
        }
        btn.disabled    = false;
        btn.textContent = 'Save Changes';
    };

})();
