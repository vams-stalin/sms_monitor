(function () {

    /* ── Redirect on 401/403 ── */
    async function apiFetch(url, options) {
        const res = await fetch(url, options || {});
        if (res.status === 401 || res.status === 403) {
            window.location.href = '/';
            return null;
        }
        return res;
    }

    /* ── XSS-safe escaper ── */
    function escHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /* ── Live date ── */
    function updateDate() {
        const el = document.getElementById('live-date');
        if (!el) return;
        const now = new Date();
        el.textContent =
            now.toLocaleDateString('en-IN', { weekday:'short', day:'numeric', month:'short', year:'numeric' }) +
            ' · ' +
            now.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit' });
    }
    updateDate();
    setInterval(updateDate, 30000);

    /* ── Audit log state ── */
    const PAGE_SIZE       = 50;
    let offset            = 0;
    let total             = 0;
    let userFilter        = '';
    let actionFilter      = '';
    let debounceTimer     = null;

    const ACTION_BADGE = {
        LOGIN:                 'badge-action-ok',
        LOGOUT:                'badge-action-neutral',
        LOGIN_FAILED:          'badge-action-danger',
        LOGIN_BLOCKED:         'badge-action-danger',
        VIEW_DASHBOARD:        'badge-action-view',
        VIEW_CLIENT_ANALYTICS: 'badge-action-view',
        ASSIGN_CREDITS:        'badge-action-write',
        MANUAL_SYNC:           'badge-action-write',
        SEED_TEST_DATA:        'badge-action-write',
        SEND_ALERT_EMAIL:      'badge-action-write',
        SEND_REPORT_NOW:       'badge-action-write',
        SEND_CLIENT_REPORT:    'badge-action-write',
        CREATE_USER:           'badge-action-admin',
        ACTIVATE_USER:         'badge-action-admin',
        DEACTIVATE_USER:       'badge-action-danger',
        RESET_PASSWORD:        'badge-action-admin',
    };

    async function load(reset) {
        if (reset) offset = 0;

        const params = new URLSearchParams({
            limit: PAGE_SIZE, offset, user: userFilter, action: actionFilter
        });

        try {
            const res  = await apiFetch('/api/audit-log?' + params);
            if (!res) return;
            const data = await res.json();
            if (data.status === 'ok') { total = data.total; render(data.rows); }
        } catch (err) {
            console.error('[Audit] load failed:', err);
        }
    }

    function render(rows) {
        const tbody = document.getElementById('audit-body');

        if (!rows || rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="6" class="list-empty">No audit entries found.</td></tr>`;
            document.getElementById('audit-count').textContent        = 'No entries found';
            document.getElementById('audit-footer-count').textContent = '0 entries';
            document.getElementById('audit-pagination').innerHTML     = '';
            return;
        }

        tbody.innerHTML = rows.map(r => {
            const cls    = ACTION_BADGE[r.action] || 'badge-action-neutral';
            const target = r.target_id
                ? `<code class="audit-target">${escHtml(r.target_id)}</code>`
                : '<span class="audit-empty-cell">—</span>';
            const detail = r.detail
                ? `<span class="audit-detail">${escHtml(r.detail)}</span>`
                : '<span class="audit-empty-cell">—</span>';
            const ip = r.ip
                ? `<span class="audit-ip">${escHtml(r.ip)}</span>`
                : '<span class="audit-empty-cell">—</span>';

            return `<tr>
                <td class="audit-ts">${escHtml(r.created)}</td>
                <td><strong class="audit-user">${escHtml(r.username)}</strong></td>
                <td><span class="badge-action ${cls}">${escHtml(r.action)}</span></td>
                <td>${target}</td>
                <td>${detail}</td>
                <td>${ip}</td>
            </tr>`;
        }).join('');

        const start = offset + 1;
        const end   = Math.min(offset + rows.length, total);
        const summary = `${start.toLocaleString()}–${end.toLocaleString()} of ${total.toLocaleString()} entries`;
        document.getElementById('audit-count').textContent        = `Activity Records — ${summary}`;
        document.getElementById('audit-footer-count').textContent = `Showing ${summary}`;

        renderPagination();
    }

    function renderPagination() {
        const pg    = document.getElementById('audit-pagination');
        const pages = Math.ceil(total / PAGE_SIZE);
        const cur   = Math.floor(offset / PAGE_SIZE) + 1;

        if (pages <= 1) { pg.innerHTML = ''; return; }

        pg.innerHTML = `
            <button class="page-btn" onclick="auditPage(${cur - 1})" ${cur === 1     ? 'disabled' : ''}>‹</button>
            <span class="page-indicator">${cur} / ${pages}</span>
            <button class="page-btn" onclick="auditPage(${cur + 1})" ${cur === pages ? 'disabled' : ''}>›</button>`;
    }

    window.auditPage = function (p) {
        const pages = Math.ceil(total / PAGE_SIZE);
        if (p < 1 || p > pages) return;
        offset = (p - 1) * PAGE_SIZE;
        load(false);
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };

    /* ── Filter listeners ── */
    document.getElementById('audit-user-input').addEventListener('input', function () {
        clearTimeout(debounceTimer);
        userFilter   = this.value.trim();
        debounceTimer = setTimeout(() => load(true), 350);
    });

    document.getElementById('audit-action-select').addEventListener('change', function () {
        actionFilter = this.value;
        load(true);
    });

    /* ── Initial load ── */
    load(true);

})();
