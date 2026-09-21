(function () {

    // ── Thresholds are fetched from /api/thresholds on load — change values only in database.py ──
    const THRESHOLDS = { WARNING: 50, CRITICAL: 90 };

    async function loadThresholds() {
        try {
            const res  = await fetch('/api/thresholds');
            if (!res.ok) return;
            const data = await res.json();
            THRESHOLDS.WARNING  = data.warning;
            THRESHOLDS.CRITICAL = data.critical;
        } catch (_) {}
    }

    window.openSettings = async function () {
        const res = await apiFetch('/api/send-report-now', { method: 'POST' });
        const overlay  = document.getElementById('st-overlay');
        const warnInput = document.getElementById('st-warning');
        const critInput = document.getElementById('st-critical');
        const fb        = document.getElementById('st-feedback');
        fb.textContent  = '';
        fb.className    = 'st-feedback';
        try {
            const res  = await fetch('/api/thresholds');
            const data = await res.json();
            warnInput.value = data.warning;
            critInput.value = data.critical;
        } catch (_) {
            warnInput.value = THRESHOLDS.WARNING;
            critInput.value = THRESHOLDS.CRITICAL;
        }
        overlay.style.display = 'flex';
        warnInput.focus();
    };

    window.closeSettings = function () {
        document.getElementById('st-overlay').style.display = 'none';
    };

    window.saveSettings = async function () {
        const btn       = document.getElementById('st-save-btn');
        const fb        = document.getElementById('st-feedback');
        const warning   = parseInt(document.getElementById('st-warning').value,  10);
        const critical  = parseInt(document.getElementById('st-critical').value, 10);

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

        btn.disabled = true;
        btn.textContent = 'Saving…';
        try {
            const res  = await apiFetch('/api/thresholds', {
                method:  'POST',
                headers: {'Content-Type': 'application/json'},
                body:    JSON.stringify({ warning, critical })
            });
            if (!res) return;
            const data = await res.json();
            if (data.status === 'ok') {
                THRESHOLDS.WARNING  = data.warning;
                THRESHOLDS.CRITICAL = data.critical;
                // Recompute status on every cached client with the new thresholds
                allData.forEach(c => { c.status = getStatus(c.pct); });
                // Update notification tab labels
                const t90 = document.querySelector('[data-filter="threshold_90"]');
                const t50 = document.querySelector('[data-filter="threshold_50"]');
                if (t90) t90.textContent = THRESHOLDS.CRITICAL + '%';
                if (t50) t50.textContent = THRESHOLDS.WARNING  + '%';
                // Re-render table with recomputed statuses
                applyFilters();
                fb.textContent = 'Saved! Thresholds updated successfully.';
                fb.className   = 'st-feedback ok';
                setTimeout(closeSettings, 1200);
            } else {
                fb.textContent = data.message || 'Failed to save.';
                fb.className   = 'st-feedback error';
            }
        } catch (_) {
            fb.textContent = 'Request failed. Check your connection.';
            fb.className   = 'st-feedback error';
        }
        btn.disabled = false;
        btn.textContent = 'Save Changes';
    };

    async function apiFetch(url, options = {}) {
        const res = await fetch(url, options);
        if (res.status === 401 || res.status === 403) {
            window.location.href = '/';
            return null;
        }
        return res;
    }

    let allData      = [];
    let filteredData = [];
    let searchQuery  = '';
    let activeFilter = 'all';
    const PAGE_SIZE  = 15;
    let currentPage  = 1;


    async function populateMasterDropdown() {

        const sel = document.getElementById('mf-client');
        if (!sel) return;

        sel.innerHTML = '<option value="">Loading clients…</option>';
        sel.disabled  = true;

        try {
            const res  = await fetch('/api/clients');
            const data = await res.json();

            if (data.status === 'ok' && data.clients.length > 0) {
                sel.innerHTML = '<option value="">— Select client —</option>';
                data.clients.forEach(c => {
                    const opt       = document.createElement('option');
                    opt.value       = c.id;
                    opt.textContent = c.name;
                    sel.appendChild(opt);
                });
                sel.disabled = false;
            } else {
                sel.innerHTML = '<option value="">No clients found</option>';
            }

        } catch (err) {
            sel.innerHTML = '<option value="">Failed to load clients</option>';
            console.error('[Dashboard] Error fetching clients:', err);
        }
    }


    function getStatus(pct) {
        if (pct >= THRESHOLDS.CRITICAL) return 'critical';
        if (pct >= THRESHOLDS.WARNING)  return 'warning';
        return 'healthy';
    }

    function statusLabel(s) {
        if (s === 'critical') return 'Critical';
        if (s === 'warning')  return 'Warning';
        return 'Healthy';
    }

    function escHtml(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function formatDate(iso) {
        if (!iso) return '';
        const d = new Date(iso + 'T00:00:00');
        return d.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
    }

    function fmt(n) {

        return n.toLocaleString('en-IN');
    }

    function initials(name) {

        return name
            .split(' ')
            .slice(0, 2)
            .map(w => w[0])
            .join('')
            .toUpperCase();
    }


    let dropdownLoaded = false;

    window.toggleMasterUser = function () {

        const panel  = document.getElementById('master-panel');
        const btn    = document.getElementById('master-user-toggle');
        const isOpen = panel.classList.contains('open');

        panel.classList.toggle('open', !isOpen);
        btn.classList.toggle('open', !isOpen);

        // Fetch client list only on first open
        if (!isOpen && !dropdownLoaded) {
            populateMasterDropdown();
            dropdownLoaded = true;
        }
    };


    window.toggleCreditForm = function () {
        const panel = document.getElementById('credit-form-panel');
        const btn   = document.getElementById('credit-usage-toggle');
        if (!panel || !btn) return;
        const isOpen = panel.style.display !== 'none';
        panel.style.display = isOpen ? 'none' : '';
        btn.classList.toggle('open', !isOpen);
    };


    async function loadDashboard() {

        try {

            const res  = await fetch('/api/dashboard-clients');
            const data = await res.json();

            allData = data.clients
                .map(c => ({
                    id:            c.id,
                    name:          c.name,
                    purchased:     c.allocated,
                    consumed:      c.used,
                    remaining:     c.remaining,
                    pct:           Math.round(c.usage_pct),
                    status:        getStatus(Math.round(c.usage_pct)),
                    today:         c.today_usage     || 0,
                    yesterday:     c.yesterday_usage || 0,
                    spike:         c.spike_detected  || false,
                    drop:          c.drop_detected   || false,
                    validityStart: c.validity_start  || null,
                    validityEnd:   c.validity_end    || null,
                    reportEmail:   c.report_email    || null,
                }))
                .sort((a, b) => a.name.localeCompare(b.name));

            applyFilters();

        } catch (err) {

            console.error(
                '[Dashboard] Failed to load dashboard data:',
                err
            );
        }
    }


    /* 
       Master User: assign credits
       Saves directly into DB
     */
    window.assignCredits = async function () {

        const sel        = document.getElementById('mf-client');
        const countIn    = document.getElementById('mf-sms-count');
        const startIn    = document.getElementById('mf-start-date');
        const endIn      = document.getElementById('mf-end-date');
        const feedback   = document.getElementById('mf-feedback');
        const btn        = document.getElementById('mf-submit-btn');

        const clientId   = sel.value.trim();
        const clientName = sel.options[sel.selectedIndex]?.text || '';
        const smsCount   = parseInt(countIn.value, 10);
        const startDate  = startIn ? startIn.value.trim() : '';
        const endDate    = endIn   ? endIn.value.trim()   : '';

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
        feedback.textContent = 'Saving...';
        feedback.classList.add('success');

        const payload = { client_id: clientId, allocated_sms: smsCount };
        if (startDate && endDate) {
            payload.start_date = startDate;
            payload.end_date   = endDate;
        }

        try {
            const res = await fetch('/api/assign-credits', {
                method:  'POST',
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify(payload)
            });

            const data = await res.json();

            if (data.status === 'ok') {
                const c   = data.client;
                const pct = Math.round(c.usage_pct);
                const row = {
                    id:            c.id,
                    name:          c.name,
                    purchased:     c.allocated,
                    consumed:      c.used,
                    remaining:     c.remaining,
                    pct:           pct,
                    status:        getStatus(pct),
                    today:         c.today_usage     || 0,
                    yesterday:     c.yesterday_usage || 0,
                    spike:         c.spike_detected  || false,
                    drop:          c.drop_detected   || false,
                    validityStart: c.validity_start  || null,
                    validityEnd:   c.validity_end    || null,
                };

                const idx = allData.findIndex(x => x.id === c.id);
                if (idx !== -1) allData[idx] = row;
                else allData.push(row);

                applyFilters();

                feedback.textContent =
                    `✓ ${smsCount.toLocaleString('en-IN')} credits assigned to ${clientName}`;

                if (window.showToast) showToast(`Credits assigned to ${escHtml(clientName)}.`, 'success');

            } else {
                feedback.textContent = data.message || 'Error saving credits.';
                feedback.classList.remove('success');
                feedback.classList.add('error');
                if (window.showToast) showToast('Failed to assign credits.', 'error');
            }

        } catch (err) {
            feedback.textContent = 'Network error. Try again.';
            feedback.classList.remove('success');
            feedback.classList.add('error');
            if (window.showToast) showToast('Network error. Check your connection.', 'error');
            console.error(err);
        }

        setTimeout(() => {
            sel.value            = '';
            countIn.value        = '';
            if (startIn) startIn.value = '';
            if (endIn)   endIn.value   = '';
            feedback.textContent = '';
            btn.disabled         = false;
        }, 2500);
    };


    /* Filter + search  */
    function applyFilters() {

        filteredData = allData.filter(c => {

            const matchSearch =
                c.name.toLowerCase().includes(searchQuery.toLowerCase());

            const matchFilter =
                activeFilter === 'all' || c.status === activeFilter;

            return matchSearch && matchFilter;
        });

        currentPage = 1;

        renderTable();
    }


    /* Render table */
    function renderTable() {
    
            const tbody    = document.getElementById('table-body');
            const start    = (currentPage - 1) * PAGE_SIZE;
            const pageData = filteredData.slice(start, start + PAGE_SIZE);
    
            if (allData.length === 0) {
                tbody.innerHTML = `
                    <tr><td colspan="10">
                        <div class="empty-state">
                            <svg width="38" height="38" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" opacity="0.35">
                                <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
                            </svg>
                            <p class="empty-title">No clients assigned yet</p>
                            <p class="empty-sub">Use <strong>Master User</strong> in the sidebar to assign SMS credits to a client.</p>
                        </div>
                    </td></tr>`;
                document.getElementById('table-count').textContent = 'No clients assigned yet';
                document.getElementById('pagination').innerHTML = '';
                return;
            }
    
            if (pageData.length === 0) {
                tbody.innerHTML = `<tr><td colspan="10" class="list-empty">No clients match your filter.</td></tr>`;
                document.getElementById('table-count').textContent = '0 clients';
                renderPagination();
                return;
            }
    
            tbody.innerHTML = pageData.map(c => {
                let validityHtml = '<span class="cr-validity-none">—</span>';
                if (c.validityStart && c.validityEnd) {
                    validityHtml = `<span class="cr-validity-range">${formatDate(c.validityStart)} – ${formatDate(c.validityEnd)}</span>`;
                }
                return `
                <tr>
                    <td class="col-client">
                        <div class="cr-left">
                            <div class="cr-avatar">${initials(c.name)}</div>
                            <div class="cr-info">
                                <div class="cr-name-wrap">
                                    <span class="cr-name">${escHtml(c.name)}</span>
                                    ${c.spike ? '<span class="badge-spike">▲ Spike</span>' : ''}
                                    ${c.drop  ? '<span class="badge-drop">▼ Drop</span>'  : ''}
                                </div>
                            </div>
                        </div>
                    </td>
                    <td class="col-num">${fmt(c.purchased)}</td>
                    <td class="col-num">${fmt(c.consumed)}</td>
                    <td class="col-num">${fmt(c.today)}</td>
                    <td class="col-num">${fmt(c.yesterday)}</td>
                    <td class="col-bar">
                        <div class="cr-bar-wrap">
                            <div class="cr-bar-track">
                                <div class="cr-bar-fill ${c.status}" style="width:${c.pct}%"></div>
                            </div>
                            <span class="cr-pct">${c.pct}%</span>
                        </div>
                    </td>
                    <td class="col-remaining"><span class="cr-remaining ${c.status}">${fmt(c.remaining)}</span></td>
                    <td class="col-validity">${validityHtml}</td>
                    <td class="col-status">
                        <span class="badge ${c.status}">
                            <span class="badge-dot"></span>${statusLabel(c.status)}
                        </span>
                    </td>
                    <td class="col-actions">${_isAdmin ? `<button class="send-client-btn" onclick="sendClientReport('${escHtml(c.id)}', this)">Send Report</button>` : ''}</td>
                </tr>`;
            }).join('');
    
            const total = filteredData.length;
            document.getElementById('table-count').textContent =
                total === 1
                    ? '1 client'
                    : `Showing ${start + 1}–${Math.min(start + PAGE_SIZE, total)} of ${total} clients`;
    
            renderPagination();
        }

    /* Pagination*/
    function renderPagination() {

        const total = Math.ceil(filteredData.length / PAGE_SIZE);
        const pg    = document.getElementById('pagination');

        if (total <= 1) { pg.innerHTML = ''; return; }

        pg.innerHTML = `
            <button class="page-btn" onclick="goPage(${currentPage - 1})"
                    ${currentPage === 1 ? 'disabled' : ''}>‹</button>
            <span class="page-indicator">${currentPage} / ${total}</span>
            <button class="page-btn" onclick="goPage(${currentPage + 1})"
                    ${currentPage === total ? 'disabled' : ''}>›</button>`;
    }

    window.goPage = function (p) {

        const total =
            Math.ceil(filteredData.length / PAGE_SIZE);

        if (p < 1 || p > total) return;

        currentPage = p;

        renderTable();
    };


    /*Live date/time */
    function updateDate() {

        const el =
            document.getElementById('live-date');

        if (!el) return;

        const now = new Date();

        el.textContent =
            now.toLocaleDateString('en-IN', {
                weekday: 'short',
                day:     'numeric',
                month:   'short',
                year:    'numeric'
            }) +
            ' · ' +
            now.toLocaleTimeString('en-IN', {
                hour:   '2-digit',
                minute: '2-digit'
            });
    }

    updateDate();

    setInterval(updateDate, 30000);


    /* Refresh button — re-fetches latest DB values for visible clients  */
    document.getElementById('refresh-btn')
        .addEventListener('click', async function () {

            if (allData.length === 0) return;

            this.classList.add('spinning');
            const btn = this;

            try {
                const res  = await fetch('/api/dashboard-clients?force=true');
                const data = await res.json();

                if (data.status === 'ok' && data.clients.length > 0) {

                    // Build a lookup map from DB by client id
                    const dbMap = {};
                    data.clients.forEach(c => { dbMap[c.id] = c; });

                    // Update only clients currently shown in session
                    allData = allData.map(row => {
                        const fresh = dbMap[row.id];
                        if (!fresh) return row;
                        const pct = Math.round(fresh.usage_pct);
                        return {
                            ...row,
                            purchased:     fresh.allocated,
                            consumed:      fresh.used,
                            remaining:     fresh.remaining,
                            pct:           pct,
                            status:        getStatus(pct),
                            today:         fresh.today_usage     || 0,
                            yesterday:     fresh.yesterday_usage || 0,
                            spike:         fresh.spike_detected  || false,
                            drop:           fresh.drop_detected   || false,
                            validityStart: fresh.validity_start  || null,
                            validityEnd:   fresh.validity_end    || null,
                        };
                    });

                    applyFilters();
                }

            } catch (err) {
                console.error('[Dashboard] Refresh failed:', err);
            }

            btn.classList.remove('spinning');
        });


    /* Search */
    document.getElementById('search-input')
        .addEventListener('input', function () {

            searchQuery = this.value;

            applyFilters();
        });


    /* Filter tabs */
    document.querySelectorAll('.filter-tab')
        .forEach(btn => {

            btn.addEventListener('click', function () {

                document.querySelectorAll('.filter-tab')
                    .forEach(b => b.classList.remove('active'));

                this.classList.add('active');

                activeFilter = this.dataset.filter;

                applyFilters();
            });
        });


    /* 
       NOTIFICATIONS
    */
    let alertsData      = [];
    let activeAlertTab  = 'all';

    async function loadAlerts() {
        try {
            const res  = await fetch('/api/alerts');
            const data = await res.json();
            alertsData = data.alerts || [];
            renderAlerts();
        } catch (err) {
            console.error('[Dashboard] Failed to load alerts:', err);
        }
    }

    function renderAlerts() {

        const badge   = document.getElementById('notif-badge');
        const btn     = document.getElementById('notif-btn');
        const countEl = document.getElementById('notif-count');
        const list    = document.getElementById('notif-list');
        const total   = alertsData.length;

        const visible = activeAlertTab === 'all'
            ? alertsData
            : alertsData.filter(a => a.type === activeAlertTab);

        if (total > 0) {
            badge.textContent   = total > 99 ? '99+' : total;
            badge.style.display = 'flex';
            btn.classList.add('has-alerts');
        } else {
            badge.style.display = 'none';
            btn.classList.remove('has-alerts');
        }

        countEl.textContent = visible.length === 0
            ? 'No alerts'
            : `${visible.length} of ${total} shown`;

        if (visible.length === 0) {
            list.innerHTML = '<div class="notif-empty">No alerts in this category.</div>';
            return;
        }

        list.innerHTML = visible.map((a, idx) => {
            // idx here maps to position in visible — need original index for sendAlertEmail
            const originalIdx = alertsData.indexOf(a);

            const iconClass = a.type === 'threshold_90' ? 'critical'
                            : a.type === 'threshold_50' ? 'warning'
                            : a.type === 'spike'        ? 'spike'
                            : 'drop';

            const iconText  = a.type === 'threshold_90' ? THRESHOLDS.CRITICAL + '%'
                            : a.type === 'threshold_50' ? THRESHOLDS.WARNING  + '%'
                            : a.type === 'spike'        ? '▲'
                            : '▼';

            const desc = a.type === 'threshold_90'
                ? `Usage at ${a.usage_pct}% — critical threshold reached`
                : a.type === 'threshold_50'
                ? `Usage at ${a.usage_pct}% — warning threshold reached`
                : a.type === 'spike'
                ? `Today: ${a.today_usage.toLocaleString('en-IN')} vs 7-day avg ${a.seven_day_avg.toLocaleString('en-IN')}`
                : `Today: ${a.today_usage.toLocaleString('en-IN')} vs 7-day avg ${a.seven_day_avg.toLocaleString('en-IN')}`;

            const btnHtml = a.email_sent
                ? '<button class="notif-send-btn sent" disabled>✓ Sent</button>'
                : `<button class="notif-send-btn" onclick="sendAlertEmail(${originalIdx})">Send Email</button>`;

            return `
                <div class="notif-item" id="notif-item-${originalIdx}">
                    <div class="notif-icon ${iconClass}">${iconText}</div>
                    <div class="notif-body">
                        <div class="notif-name">${escHtml(a.client_name)}</div>
                        <div class="notif-desc">${desc}</div>
                    </div>
                    ${btnHtml}
                </div>`;

        }).join('');
    }

    window.sendClientReport = async function (clientId, btn) {
        btn.disabled    = true;
        btn.textContent = 'Sending…';
        try {
            const res = await apiFetch('/api/send-client-report', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ client_id: clientId }) });
            if (!res) { btn.disabled = false; btn.textContent = 'Send Report'; return; }
            const data = await res.json();
            if (data.status === 'ok') {
                btn.textContent = '✓ Sent';
                if (window.showToast) showToast('Report sent successfully.', 'success');
            } else {
                btn.textContent = 'Failed';
                if (window.showToast) showToast(data.message || 'Failed to send report.', 'error');
            }
        } catch (_) {
            btn.textContent = 'Error';
            if (window.showToast) showToast('Network error. Report not sent.', 'error');
        }
        setTimeout(() => { btn.textContent = 'Send Report'; btn.disabled = false; }, 3000);
    };


    window.sendAlertEmail = async function (idx) {

        const alert = alertsData[idx];
        const btn   = document.querySelector(`#notif-item-${idx} .notif-send-btn`);
        if (!btn || alert.email_sent) return;

        btn.disabled    = true;
        btn.textContent = 'Sending…';

        try {
            const res  = await fetch('/api/send-alert-email', {
                method:  'POST',
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify({ client_id: alert.client_id, alert_type: alert.type })
            });
            const data = await res.json();

            if (data.status === 'ok') {
                alertsData[idx].email_sent = true;
                btn.textContent = '✓ Sent';
                btn.classList.add('sent');
                if (window.showToast) showToast('Alert email sent for ' + escHtml(alert.client_name) + '.', 'success');
            } else {
                btn.disabled    = false;
                btn.textContent = 'Retry';
                if (window.showToast) showToast('Failed to send alert email.', 'error');
            }
        } catch (err) {
            btn.disabled    = false;
            btn.textContent = 'Retry';
            if (window.showToast) showToast('Network error. Email not sent.', 'error');
            console.error(err);
        }
    };

    // Toggle panel open/close
    document.getElementById('notif-btn').addEventListener('click', function (e) {
        e.stopPropagation();
        document.getElementById('notif-panel').classList.toggle('open');
    });

    // Close when clicking anywhere outside
    document.addEventListener('click', function () {
        document.getElementById('notif-panel').classList.remove('open');
    });
    document.getElementById('notif-panel').addEventListener('click', function (e) {
        e.stopPropagation();
    });

    // Alert filter tabs
    document.querySelectorAll('.notif-tab').forEach(tab => {
        tab.addEventListener('click', function () {
            document.querySelectorAll('.notif-tab').forEach(t => t.classList.remove('active'));
            this.classList.add('active');
            activeAlertTab = this.dataset.filter;
            renderAlerts();
        });
    });


    /* Last synced timestamp  */
    let lastSyncedISO = null;

    function timeAgo(isoString) {
        if (!isoString) return null;
        const diff = Math.floor((Date.now() - new Date(isoString)) / 1000);
        if (diff < 60)  return 'Just now';
        if (diff < 3600) return `${Math.floor(diff / 60)} min ago`;
        if (diff < 86400) return `${Math.floor(diff / 3600)} hr ago`;
        return `${Math.floor(diff / 86400)} day(s) ago`;
    }

    function updateSyncLabel() {
        const el = document.getElementById('last-synced');
        if (!el) return;
        if (!lastSyncedISO) {
            el.textContent = 'Not yet synced';
            el.className   = 'last-synced unknown';
            return;
        }
        const diffMin = (Date.now() - new Date(lastSyncedISO)) / 60000;
        el.textContent = `Synced ${timeAgo(lastSyncedISO)}`;
        el.className   = `last-synced${diffMin > 90 ? ' stale' : ''}`;
    }

    async function fetchLastSynced() {
        try {
            const res  = await fetch('/api/last-synced');
            const data = await res.json();
            lastSyncedISO = data.last_synced || null;
            updateSyncLabel();
        } catch (err) {
            console.error('[Dashboard] Failed to fetch last synced:', err);
        }
    }

    fetchLastSynced();
    setInterval(updateSyncLabel, 60000);   // refresh label every minute
    setInterval(fetchLastSynced, 300000);  // re-fetch from server every 5 min


    /*  Init  */
    // Role-based UI visibility
    const _role     = window.CURRENT_ROLE || 'viewer';
    const _username = (window.CURRENT_USER || 'User').trim();
    const _isAdmin  = _role === 'admin';

    const _usersNav    = document.getElementById('users-nav');
    const _auditNav    = document.getElementById('audit-log-nav');
    const _settingsNav = document.getElementById('settings-nav');
    const _trendsNav   = document.getElementById('trends-nav');
    const _usageReportNav = document.getElementById('usage-report-nav');
    const _masterGroup = document.getElementById('master-user-group');
    const _reportBtn   = document.getElementById('send-report-btn');
    const _sidebarName = document.getElementById('sidebar-username');
    const _sidebarRole = document.getElementById('sidebar-userrole');
    const _avatar      = document.getElementById('sidebar-avatar');

    if (_usersNav)    _usersNav.style.display    = _isAdmin ? '' : 'none';
    if (_auditNav)    _auditNav.style.display    = _isAdmin ? '' : 'none';
    if (_settingsNav) _settingsNav.style.display = _isAdmin ? '' : 'none';
    if (_trendsNav)   _trendsNav.style.display   = _isAdmin ? '' : 'none';
    if (_usageReportNav) _usageReportNav.style.display = _isAdmin ? '' : 'none';
    if (_masterGroup) _masterGroup.style.display = _isAdmin ? '' : 'none';
    const _actionsHeader = document.getElementById('col-actions-header');
    if (_actionsHeader) _actionsHeader.textContent = _isAdmin ? 'Report' : '';
    if (_sidebarName) _sidebarName.textContent   = _username;
    if (_sidebarRole) _sidebarRole.textContent   = _isAdmin ? 'Admin' : 'Viewer';
    if (_avatar)      _avatar.textContent        = _username.charAt(0).toUpperCase();

    if (_reportBtn) {
        _reportBtn.style.display = _isAdmin ? '' : 'none';
        _reportBtn.addEventListener('click', async function () {
            const label = document.getElementById('send-report-label');
            _reportBtn.disabled  = true;
            label.textContent    = 'Sending…';
            try {
    // Get the client currently selected in Analytics
    const analyticsClient = document.getElementById('analytics-client');
    const clientId = analyticsClient ? analyticsClient.value.trim() : '';

    // A detailed report needs a specific client
    if (!clientId) {
        label.textContent = 'Select Client';

        if (window.showToast) {
            showToast(
                'Please select a client in Analytics first.',
                'error'
            );
        }

        setTimeout(() => {
            label.textContent = 'Send Report';
            _reportBtn.disabled = false;
        }, 2000);

        return;
    }

    // Send the detailed SMS report request
    const res = await apiFetch('/api/send-client-report', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            client_id: clientId
        })
    });

    if (!res) {
        label.textContent = 'Failed';
        return;
    }

    const data = await res.json();

    if (data.status === 'ok') {
        label.textContent = 'Report Sent!';

        if (window.showToast) {
            showToast(
                'Detailed SMS report sent to inbox.',
                'success'
            );
        }
    } else {
        label.textContent = 'Failed';

        if (window.showToast) {
            showToast(
                data.message || 'Failed to send detailed report.',
                'error'
            );
        }
    }

} catch (_) {
    label.textContent = 'Failed';

    if (window.showToast) {
        showToast(
            'Network error. Report not sent.',
            'error'
        );
    }
}
            setTimeout(() => {
                label.textContent   = 'Send Report';
                _reportBtn.disabled = false;
            }, 3000);
        });
    }

    // Skeleton rows while data loads
    (function () {
        const ws = [65, 80, 55, 72, 60, 75, 50];
        document.getElementById('table-body').innerHTML = ws.map((w, i) => `
            <tr class="sk-row" style="animation-delay:${i * 0.06}s">
                <td class="col-client">
                    <div class="cr-left">
                        <span class="sk-box" style="width:28px;height:28px;flex-shrink:0"></span>
                        <span class="sk-line" style="width:${w}%;max-width:150px"></span>
                    </div>
                </td>
                <td class="col-num"><span class="sk-line" style="width:50%;margin-left:auto"></span></td>
                <td class="col-num"><span class="sk-line" style="width:50%;margin-left:auto"></span></td>
                <td class="col-num"><span class="sk-line" style="width:50%;margin-left:auto"></span></td>
                <td class="col-num"><span class="sk-line" style="width:50%;margin-left:auto"></span></td>
                <td class="col-bar"><span class="sk-line" style="width:88%"></span></td>
                <td class="col-remaining"><span class="sk-line" style="width:55%;margin-left:auto"></span></td>
                <td class="col-validity"><span class="sk-line" style="width:80%"></span></td>
                <td class="col-status"><span class="sk-line" style="width:76%;height:22px;border-radius:100px"></span></td>
                <td class="col-actions"></td>
            </tr>`).join('');
    })();
    loadThresholds().then(() => {
        document.querySelector('[data-filter="threshold_90"]').textContent = THRESHOLDS.CRITICAL + '%';
        document.querySelector('[data-filter="threshold_50"]').textContent = THRESHOLDS.WARNING  + '%';
        loadDashboard();
        loadAlerts();
    });


    /*
       VIEW TABS — Table / Analytics
    */
    function switchView(view) {
        document.querySelectorAll('.view-tab').forEach(t =>
            t.classList.toggle('active', t.dataset.view === view)
        );
        document.getElementById('view-table').style.display     = view === 'table'     ? '' : 'none';
        document.getElementById('view-analytics').style.display = view === 'analytics' ? '' : 'none';

        if (view === 'analytics') populateAnalyticsDropdown();
    }

    document.querySelectorAll('.view-tab').forEach(tab =>
        tab.addEventListener('click', () => switchView(tab.dataset.view))
    );


    /* 
       ANALYTICS — dropdown + charts
     */
    let donutChart = null;
    let barChart   = null;

    function destroyCharts() {
        if (donutChart) { donutChart.destroy(); donutChart = null; }
        if (barChart)   { barChart.destroy();   barChart   = null; }
    }

    function populateAnalyticsDropdown() {
        const sel      = document.getElementById('analytics-client');
        const prevVal  = sel.value;

        sel.innerHTML = '<option value="">— Select a client —</option>';
        allData.forEach(c => {
            const opt       = document.createElement('option');
            opt.value       = c.id;
            opt.textContent = c.name;
            sel.appendChild(opt);
        });

        if (prevVal) {
            sel.value = prevVal;
            if (sel.value) loadClientAnalytics(sel.value);
        }
    }

    async function loadClientAnalytics(clientId) {
        if (!clientId) {
            document.getElementById('analytics-empty').style.display = '';
            document.getElementById('analytics-content').style.display = 'none';
            destroyCharts();
            return;
        }
    
        try {
            const res = await apiFetch(
                `/api/client-analytics/${encodeURIComponent(clientId)}`
            );
    
            if (!res) return;
    
            if (!res.ok) {
                const text = await res.text();
                console.error(
                    '[Analytics] API error:',
                    res.status,
                    text
                );
    
                if (window.showToast) {
                    showToast(
                        `Analytics failed (${res.status})`,
                        'error'
                    );
                }
    
                return;
            }
    
            const data = await res.json();
    
            if (data.status === 'ok') {
                renderAnalytics(data);
            } else {
                console.error('[Analytics] Server error:', data);
    
                if (window.showToast) {
                    showToast(
                        data.message || 'Unable to load analytics.',
                        'error'
                    );
                }
            }
    
        } catch (err) {
            console.error('[Analytics] Failed to load:', err);
    
            if (window.showToast) {
                showToast(
                    'Unable to load analytics. Check the server.',
                    'error'
                );
            }
        }
    }

    function renderAnalytics(data) {
        document.getElementById('analytics-empty').style.display   = 'none';
        document.getElementById('analytics-content').style.display = '';

        const status    = data.usage_pct >= THRESHOLDS.CRITICAL ? 'critical' : data.usage_pct >= THRESHOLDS.WARNING ? 'warning' : 'healthy';
        const usedColor = status === 'critical' ? '#dc2626' : status === 'warning' ? '#d97706' : '#16a34a';
        const todayCount = data.today_count || 0;

        // Stat cards
        document.getElementById('analytics-stats').innerHTML = `
            <div class="an-stat">
                <div class="an-stat-label">Allocated</div>
                <div class="an-stat-value">${data.allocated.toLocaleString('en-IN')}</div>
                <div class="an-stat-sub">Total SMS credits</div>
            </div>
            <div class="an-stat">
                <div class="an-stat-label">Used</div>
                <div class="an-stat-value" style="color:${usedColor}">${data.used.toLocaleString('en-IN')}</div>
                <div class="an-stat-sub">${data.usage_pct}% consumed</div>
            </div>
            <div class="an-stat">
                <div class="an-stat-label">Remaining</div>
                <div class="an-stat-value" style="color:var(--green)">${data.remaining.toLocaleString('en-IN')}</div>
                <div class="an-stat-sub">Credits left</div>
            </div>
            <div class="an-stat">
                <div class="an-stat-label">Today</div>
                <div class="an-stat-value">${todayCount.toLocaleString('en-IN')}</div>
                <div class="an-stat-sub">7-day avg: ${data.seven_day_avg.toLocaleString('en-IN')}</div>
            </div>`;

        // Alert banners
        const alerts = [];
        if (data.usage_pct >= THRESHOLDS.CRITICAL)
            alerts.push({ cls: 'an-alert-critical', icon: '⚠', text: `Critical: ${data.usage_pct}% of credits consumed — ${THRESHOLDS.CRITICAL}% threshold reached` });
        else if (data.usage_pct >= THRESHOLDS.WARNING)
            alerts.push({ cls: 'an-alert-warning',  icon: '⚠', text: `Warning: ${data.usage_pct}% of credits consumed — ${THRESHOLDS.WARNING}% threshold reached` });

        if (data.spike_detected)
            alerts.push({ cls: 'an-alert-spike', icon: '▲', text: `Spike detected: today's usage (${todayCount.toLocaleString('en-IN')}) is ${(todayCount / Math.max(data.seven_day_avg, 1)).toFixed(1)}× the 7-day average` });

        if (data.drop_detected)
            alerts.push({ cls: 'an-alert-drop',  icon: '▼', text: `Drop detected: today's usage (${todayCount.toLocaleString('en-IN')}) is well below the 7-day average (${data.seven_day_avg.toLocaleString('en-IN')})` });

        const alertsEl = document.getElementById('analytics-alerts');
        if (alerts.length > 0) {
            alertsEl.innerHTML = alerts.map(a =>
                `<div class="an-alert ${a.cls}"><span class="an-alert-icon">${a.icon}</span>${a.text}</div>`
            ).join('');
            alertsEl.style.display = '';
        } else {
            alertsEl.innerHTML = '';
            alertsEl.style.display = 'none';
        }

        // Donut center text
        document.getElementById('donut-center-pct').textContent   = data.usage_pct + '%';
        document.getElementById('donut-center-label').textContent =
            status === 'critical' ? 'Critical' : status === 'warning' ? 'Warning' : 'Healthy';

        destroyCharts();

        // Donut chart — Used vs Remaining
        donutChart = new Chart(document.getElementById('donut-chart'), {
            type: 'doughnut',
            data: {
                labels: ['Used', 'Remaining'],
                datasets: [{
                    data: [data.used, Math.max(0, data.remaining)],
                    backgroundColor: [usedColor, '#e2e8f0'],
                    borderWidth: 0,
                    hoverOffset: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '72%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { font: { family: 'Figtree', size: 11 }, padding: 16, usePointStyle: true }
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => ` ${ctx.label}: ${ctx.raw.toLocaleString('en-IN')} SMS`
                        }
                    }
                }
            }
        });

        // Bar chart — 7-day daily trend
        const labels = data.daily.length > 0 ? data.daily.map(d => d.date) : ['No data'];
        const counts = data.daily.length > 0 ? data.daily.map(d => d.count) : [0];

        // Color today's bar by alert status; past days in lighter blue
        const todayIdx   = counts.length - 1;
        const barColors  = counts.map((_, i) => {
            if (i !== todayIdx) return '#93c5fd';
            if (data.spike_detected) return '#dc2626';
            if (data.drop_detected)  return '#2563a8';
            return '#1d4e8f';
        });

        barChart = new Chart(document.getElementById('bar-chart'), {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: 'SMS Sent',
                    data: counts,
                    backgroundColor: barColors,
                    borderRadius: 5,
                    borderSkipped: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: ctx => ` ${ctx.raw.toLocaleString('en-IN')} SMS`
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: '#f1f5f9' },
                        ticks: { font: { family: 'Figtree', size: 11 }, color: '#64748b' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { font: { family: 'Figtree', size: 11 }, color: '#64748b' }
                    }
                }
            }
        });
    }

    const analyticsClient = document.getElementById('analytics-client');
    if (analyticsClient) {
        analyticsClient.addEventListener('change', function () {
            loadClientAnalytics(this.value);
        });
    }

})();