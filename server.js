const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const sqlite3 = require('sqlite3');
const { open } = require('sqlite');

const app = express();
let db;

const LOG_FILE = '/var/log/policy-portal.log';
function log(level, msg, data) {
    const ts = new Date().toISOString().replace('T', ' ').slice(0, 19);
    const line = '[' + ts + '] [' + level + '] ' + msg + (data ? ' | ' + JSON.stringify(data) : '');
    console.log(line);
    fs.appendFile(LOG_FILE, line + '\n', () => {});
}

const ADMIN_USER = 'admin';
const ADMIN_PASS = process.env.ADMIN_PASS || 'Sanghviadmin2026';
const SESSION_SECRET = crypto.randomBytes(16).toString('hex');

app.set('trust proxy', true);
app.use(express.json());
app.use(express.urlencoded({ extended: true }));


const sessions = new Set();

function parseCookies(req) {
    const cookie = req.headers.cookie || '';
    const result = {};
    cookie.split(';').forEach(c => {
        const [k, v] = c.trim().split('=');
        if (k) result[k.trim()] = (v || '').trim();
    });
    return result;
}

function requireAdmin(req, res, next) {
    const cookies = parseCookies(req);
    if (sessions.has(cookies.admin_session)) {
        next();
    } else {
        res.redirect('/admin/login');
    }
}

async function initDatabase() {
    db = await open({
        filename: 'compliance.db',
        driver: sqlite3.Database
    });

    await db.exec(`
      CREATE TABLE IF NOT EXISTS employees (
        emp_id TEXT PRIMARY KEY,
        name TEXT NOT NULL UNIQUE,
        email TEXT DEFAULT '',
        department TEXT DEFAULT ''
      );
      CREATE TABLE IF NOT EXISTS submissions (
        emp_id TEXT PRIMARY KEY,
        submitted_at TEXT NOT NULL,
        email TEXT DEFAULT '',
        ip_address TEXT DEFAULT '',
        user_agent TEXT DEFAULT '',
        read_policy TEXT DEFAULT 'No',
        q1 TEXT DEFAULT 'No',
        q2 TEXT DEFAULT 'No',
        q3 TEXT DEFAULT 'No',
        q4 TEXT DEFAULT 'No',
        q5 TEXT DEFAULT 'No',
        q6 TEXT DEFAULT 'No',
        q7 TEXT DEFAULT 'No',
        FOREIGN KEY(emp_id) REFERENCES employees(emp_id)
      );
    `);

    // Migrate old DB — add columns that may not exist
    for (const col of ['email', 'department']) {
        try { await db.run(`ALTER TABLE employees ADD COLUMN ${col} TEXT DEFAULT ''`); } catch (_) {}
    }
    for (const col of ['email', 'ip_address', 'user_agent']) {
        try { await db.run(`ALTER TABLE submissions ADD COLUMN ${col} TEXT DEFAULT ''`); } catch (_) {}
    }
    for (const col of ['q7']) {
        try { await db.run(`ALTER TABLE submissions ADD COLUMN ${col} TEXT DEFAULT 'No'`); } catch (_) {}
    }

    const wb = xlsx.readFile('users-sanghvi.xlsx');
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const seen = new Set();
    const result = await db.get('SELECT MAX(CAST(SUBSTR(emp_id,4) AS INTEGER)) AS max_id FROM employees');
    let nextIdx = (result && result.max_id) ? result.max_id + 1 : 1;
    for (const row of rows) {
        if (row[0] === 'Username' || !row[0]) continue;
        const name = row[0].toString().trim();
        if (!name || seen.has(name.toLowerCase())) continue;
        seen.add(name.toLowerCase());
        const email = row[1] ? row[1].toString().trim() : '';
        const department = row[2] ? row[2].toString().trim() : '';
        const existing = await db.get('SELECT emp_id FROM employees WHERE LOWER(TRIM(name)) = LOWER(?)', name);
        if (existing) {
            if (email || department) {
                await db.run(`UPDATE employees SET email = ?, department = ? WHERE emp_id = ?`, email, department, existing.emp_id);
            }
            continue;
        }
        const emp_id = 'SML' + String(nextIdx++).padStart(3, '0');
        await db.run(`INSERT INTO employees (emp_id, name, email, department) VALUES (?, ?, ?, ?)`, emp_id, name, email, department);
    }
    console.log(`Loaded ${seen.size} employees from users-sanghvi.xlsx`);
    log('INFO', 'Server started', { employees: seen.size, db: 'compliance.db' });
}

app.get('/policy.pdf', (req, res) => {
    res.sendFile(path.join(__dirname, 'AI Policy.pdf'));
});

app.get('/logo.jpeg', (req, res) => {
    res.sendFile(path.join(__dirname, 'logo.jpeg'));
});

app.get('/', (req, res) => {
    res.send(getFrontendHTML());
});

app.post('/api/submit', async (req, res) => {
    const { empName, empEmail, client_ip, read_policy, q1, q2, q3, q4, q5, q6, q7 } = req.body;
    const name = (empName || '').trim();
    if (!name) {
        return res.status(400).send(errorPage('Please enter your name.'));
    }
    const email = (empEmail || '').trim();
    if (!email) {
        return res.status(400).send(errorPage('Please enter your email.'));
    }
    try {
        const emp = await db.get('SELECT emp_id FROM employees WHERE LOWER(TRIM(name)) = LOWER(?)', name);
        if (!emp) {
            return res.status(400).send(errorPage(`<strong>${escapeHtml(name)}</strong> was not found in the employee roster. Please check the spelling or contact HR.`));
        }
        const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
        const ip = (client_ip || '').trim() || (req.headers['x-forwarded-for'] ? req.headers['x-forwarded-for'].split(',')[0].trim() : req.ip);
        const ua = req.headers['user-agent'] || '';
        await db.run(
            `INSERT INTO submissions (emp_id, submitted_at, email, ip_address, user_agent, read_policy, q1, q2, q3, q4, q5, q6, q7) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            emp.emp_id, timestamp, email, ip, ua,
            read_policy === 'Yes' ? 'Yes' : 'No',
            q1 === 'Yes' ? 'Yes' : 'No',
            q2 === 'Yes' ? 'Yes' : 'No',
            q3 === 'Yes' ? 'Yes' : 'No',
            q4 === 'Yes' ? 'Yes' : 'No',
            q5 === 'Yes' ? 'Yes' : 'No',
            q6 === 'Yes' ? 'Yes' : 'No',
            q7 === 'Yes' ? 'Yes' : 'No'
        );
        log('SUBMIT', name + ' acknowledged policy', { emp_id: emp.emp_id, email, ip, ua: ua.substring(0,100), read_policy, q1, q2, q3, q4, q5, q6, q7 });
        res.send(successPage(name));
    } catch (err) {
        if (err.message && err.message.includes('UNIQUE constraint failed')) {
            log('WARN', 'Duplicate submission attempt', { name, email });
            res.status(400).send(alreadyPage(name));
        } else {
            log('ERROR', 'Submission failed', { name, email, error: err.message });
            console.error(err);
            res.status(500).send('Internal Server Error.');
        }
    }
});

app.post('/api/verify-email', async (req, res) => {
    const { email } = req.body;
    if (!email || !email.trim()) {
        return res.json({ found: false, error: 'Please enter your email.' });
    }
    try {
        const emp = await db.get('SELECT name, emp_id FROM employees WHERE LOWER(TRIM(email)) = LOWER(?)', email.trim());
        if (emp) {
            log('INFO', 'Email verified', { email, name: emp.name });
            res.json({ found: true, name: emp.name, emp_id: emp.emp_id });
        } else {
            log('INFO', 'Email not found', { email });
            res.json({ found: false, error: 'Email not found in employee roster.' });
        }
    } catch (err) {
        console.error(err);
        res.json({ found: false, error: 'Server error. Try again.' });
    }
});

app.get('/admin/login', (req, res) => {
    res.send(getLoginHTML());
});

app.post('/admin/login', (req, res) => {
    const { username, password } = req.body;
    if (username === ADMIN_USER && password === ADMIN_PASS) {
        const token = crypto.randomBytes(24).toString('hex');
        sessions.add(token);
        res.setHeader('Set-Cookie', `admin_session=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=86400`);
        log('INFO', 'Admin login', { username });
        res.redirect('/admin');
    } else {
        log('WARN', 'Failed admin login attempt', { username });
        res.send(getLoginHTML('Invalid username or password.'));
    }
});

app.get('/admin/logout', (req, res) => {
    const cookies = parseCookies(req);
    sessions.delete(cookies.admin_session);
    res.setHeader('Set-Cookie', 'admin_session=; Path=/; Max-Age=0');
    log('INFO', 'Admin logout', {});
    res.redirect('/admin/login');
});

app.get('/admin', requireAdmin, async (req, res) => {
    const records = await db.all(`
        SELECT e.emp_id, e.name, e.email AS emp_email, e.department, s.submitted_at, s.email AS submitted_email, s.ip_address, s.user_agent, s.read_policy, s.q1, s.q2, s.q3, s.q4, s.q5, s.q6, s.q7
        FROM employees e LEFT JOIN submissions s ON e.emp_id = s.emp_id
    `);
    const completed = records.filter(r => r.submitted_at !== null);
    const pending = records.filter(r => r.submitted_at === null);
    log('INFO', 'Admin dashboard viewed', { total: records.length, completed: completed.length, pending: pending.length });
    res.send(getAdminDashboardHTML(completed, pending, req.query.msg || ''));
});

app.get('/admin/export', requireAdmin, async (req, res) => {
    const records = await db.all(`
        SELECT e.emp_id AS [Employee ID], e.name AS [Employee Name],
               COALESCE(e.department, '') AS [Department],
               COALESCE(e.email, '') AS [Email ID],
               CASE WHEN s.submitted_at IS NOT NULL THEN 'COMPLIANT' ELSE 'PENDING' END AS [Status],
               COALESCE(s.submitted_at, 'N/A') AS [Date of Acknowledgement],
               COALESCE(s.ip_address, '') AS [IP Address],
               COALESCE(s.user_agent, '') AS [Device Info],
               COALESCE(s.read_policy, 'No') AS [Read Policy Document],
                COALESCE(s.q1, 'No') AS [1. Dual Scoring Framework],
                COALESCE(s.q2, 'No') AS [2. Personal Accountability],
                COALESCE(s.q3, 'No') AS [3. Human Review Required],
                COALESCE(s.q4, 'No') AS [4. No Sensitive Info in Prompts],
                COALESCE(s.q5, 'No') AS [5. License Reallocation],
                COALESCE(s.q6, 'No') AS [6. Token Usage],
                COALESCE(s.q7, 'No') AS [7. Disciplinary Action]
        FROM employees e LEFT JOIN submissions s ON e.emp_id = s.emp_id
    `);
    const worksheet = xlsx.utils.json_to_sheet(records);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'AI Compliance Track');
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename="AI_Policy_Compliance_Report.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    log('INFO', 'Admin exported Excel', { records: records.length });
    res.send(buffer);
});

app.post('/admin/sync-roster', requireAdmin, async (req, res) => {
    try {
        const wb = xlsx.readFile('users-sanghvi.xlsx');
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        const result = await db.get('SELECT MAX(CAST(SUBSTR(emp_id,4) AS INTEGER)) AS max_id FROM employees');
        let nextIdx = (result && result.max_id) ? result.max_id + 1 : 1;
        const changes = [];
        let totalChecked = 0;
        for (const row of rows) {
            if (row[0] === 'Username' || !row[0]) continue;
            const name = row[0].toString().trim();
            if (!name) continue;
            totalChecked++;
            const email = row[1] ? row[1].toString().trim() : '';
            const department = row[2] ? row[2].toString().trim() : '';
            const existing = await db.get('SELECT emp_id, email, department FROM employees WHERE LOWER(TRIM(name)) = LOWER(?)', name);
            if (existing) {
                let detail = '';
                const oldEmail = existing.email || '';
                const oldDept = existing.department || '';
                const newEmail = email;
                const newDept = department;
                if (newEmail !== oldEmail) {
                    if (detail) detail += '; ';
                    detail += 'email: "' + oldEmail + '" → "' + newEmail + '"';
                }
                if (newDept !== oldDept) {
                    if (detail) detail += '; ';
                    detail += 'dept: "' + oldDept + '" → "' + newDept + '"';
                }
                if (detail) {
                    await db.run(`UPDATE employees SET email = ?, department = ? WHERE emp_id = ?`, newEmail, newDept, existing.emp_id);
                    changes.push(name + ' — ' + detail);
                }
                continue;
            }
            const emp_id = 'SML' + String(nextIdx++).padStart(3, '0');
            await db.run(`INSERT INTO employees (emp_id, name, email, department) VALUES (?, ?, ?, ?)`, emp_id, name, email, department);
            changes.push('➕ ' + name + ' (new)');
        }
        let msg = '<strong>Sync complete</strong> — ' + totalChecked + ' employee(s) checked from roster.';
        if (changes.length > 0) msg += '<br><br>' + changes.join('<br>');
        else msg += '<br>No changes found — all records up to date.';
        log('INFO', 'Sync roster completed', { checked: totalChecked, updated: changes.filter(c => c.includes('—')).length, added: changes.filter(c => c.includes('(new)')).length });
        if (changes.length > 0) changes.forEach(c => log('INFO', 'Sync detail', { change: c.replace(/<br>/g, '') }));
        res.redirect('/admin?msg=' + encodeURIComponent(msg.trim()));
    } catch (err) {
        log('ERROR', 'Sync roster failed', { error: err.message });
        console.error('Sync error:', err);
        res.redirect('/admin?msg=' + encodeURIComponent('Error reading roster file.'));
    }
});

app.post('/admin/reset-submission', requireAdmin, async (req, res) => {
    const { emp_id, name } = req.body;
    if (!emp_id) return res.redirect('/admin?msg=' + encodeURIComponent('Missing employee ID.'));
    try {
        await db.run('DELETE FROM submissions WHERE emp_id = ?', emp_id);
        log('INFO', 'Submission reset', { emp_id, name });
        res.redirect('/admin?msg=' + encodeURIComponent('Reset: ' + (name || emp_id) + ' can now resubmit.'));
    } catch (err) {
        console.error(err);
        res.redirect('/admin?msg=' + encodeURIComponent('Error resetting submission.'));
    }
});

app.post('/admin/remove-employee', requireAdmin, async (req, res) => {
    const { emp_id, name, password } = req.body;
    if (!emp_id) return res.redirect('/admin?msg=' + encodeURIComponent('Missing employee ID.'));
    if (password !== ADMIN_PASS) {
        return res.redirect('/admin?msg=' + encodeURIComponent('Incorrect admin password — removal cancelled.'));
    }
    try {
        await db.run('DELETE FROM submissions WHERE emp_id = ?', emp_id);
        await db.run('DELETE FROM employees WHERE emp_id = ?', emp_id);
        log('INFO', 'Employee removed', { emp_id, name });
        res.redirect('/admin?msg=' + encodeURIComponent('Removed: ' + (name || emp_id) + ' deleted from DB. Sync roster to re-add from Excel.'));
    } catch (err) {
        console.error(err);
        res.redirect('/admin?msg=' + encodeURIComponent('Error removing employee.'));
    }
});

app.post('/admin/add-employee', requireAdmin, async (req, res) => {
    const { name, email, department } = req.body;
    const empName = (name || '').trim();
    if (!empName) {
        return res.redirect('/admin?msg=' + encodeURIComponent('Error: Name is required.'));
    }
    try {
        const existing = await db.get('SELECT emp_id FROM employees WHERE LOWER(TRIM(name)) = LOWER(?)', empName);
        if (existing) {
            return res.redirect('/admin?msg=' + encodeURIComponent('Error: "' + escapeHtml(empName) + '" already exists (ID: ' + existing.emp_id + ').'));
        }
        const result = await db.get('SELECT MAX(CAST(SUBSTR(emp_id,4) AS INTEGER)) AS max_id FROM employees');
        let nextIdx = (result && result.max_id) ? result.max_id + 1 : 1;
        const emp_id = 'SML' + String(nextIdx).padStart(3, '0');
        await db.run('INSERT INTO employees (emp_id, name, email, department) VALUES (?, ?, ?, ?)', emp_id, empName, (email || '').trim(), (department || '').trim());
        log('INFO', 'Employee added manually', { emp_id, name: empName, email: (email || '').trim(), department: (department || '').trim() });
        res.redirect('/admin?msg=' + encodeURIComponent('Added: ' + escapeHtml(empName) + ' (' + emp_id + ')'));
    } catch (err) {
        log('ERROR', 'Add employee failed', { error: err.message });
        console.error(err);
        res.redirect('/admin?msg=' + encodeURIComponent('Error adding employee.'));
    }
});

app.get('/admin/export-users', requireAdmin, async (req, res) => {
    try {
        const records = await db.all(`
            SELECT e.emp_id AS [Employee ID], e.name AS [Employee Name],
                   COALESCE(e.email, '') AS [Email ID],
                   COALESCE(e.department, '') AS [Department],
                   CASE WHEN s.submitted_at IS NOT NULL THEN 'COMPLIANT' ELSE 'PENDING' END AS [Status]
            FROM employees e LEFT JOIN submissions s ON e.emp_id = s.emp_id
            ORDER BY e.emp_id
        `);
        const worksheet = xlsx.utils.json_to_sheet(records);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Employee Roster');
        const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        res.setHeader('Content-Disposition', 'attachment; filename="Employee_Roster.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        log('INFO', 'Admin exported employee roster', { records: records.length });
        res.send(buffer);
    } catch (err) {
        log('ERROR', 'Export users failed', { error: err.message });
        res.redirect('/admin?msg=' + encodeURIComponent('Error exporting roster.'));
    }
});

initDatabase().then(() => {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log('Server: http://localhost:' + PORT);
        console.log('Admin: http://localhost:' + PORT + '/admin');
    });
}).catch(err => console.error('Startup failed:', err));

function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function resultPage(icon, iconBg, title, msg, link) {
    return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="icon" href="/logo.jpeg"><title>SML - ${title}</title><style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:linear-gradient(160deg,#eef2f7,#dce3ed);display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}
.card{background:#fff;padding:48px 44px 40px;border-radius:20px;box-shadow:0 10px 40px rgba(0,0,0,0.06);text-align:center;max-width:420px;width:100%;animation:fadeIn .4s ease-out}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
h2{color:#8b1a1a;margin-bottom:8px;font-size:1.2rem;font-weight:700}
p{color:#64748b;margin-bottom:22px;line-height:1.6;font-size:0.92rem}
.btn{display:inline-block;background:#8b1a1a;color:#fff;padding:10px 28px;border-radius:10px;text-decoration:none;font-weight:600;font-size:0.88rem;transition:all .25s}
.btn:hover{background:#b91c1c;transform:translateY(-1px);box-shadow:0 4px 12px rgba(139,26,26,0.2)}
.logo{width:40px;margin-bottom:16px;opacity:0.8}
</style></head><body><div class="card"><img src="/logo.jpeg" alt="SML" class="logo">
<h2>${title}</h2><p>${msg}</p>${link ? `<a href="${link.href}" class="btn">${link.text}</a>` : ''}</div></body></html>`;
}
function errorPage(msg) { return resultPage('', '', 'Submission Denied', msg, {href:'/',text:'Go Back'}); }
function successPage(name) { return resultPage('', '', 'Acknowledgement Recorded', 'Thank you, <strong style="color:#8b1a1a">' + escapeHtml(name) + '</strong>! Your formal sign-off on the Enterprise AI Governance policy has been logged securely.', null); }
function alreadyPage(name) { return resultPage('', '', 'Already Submitted', 'An acknowledgement has already been logged for <strong>' + escapeHtml(name) + '</strong>.', {href:'/',text:'Go Back'}); }

function getFrontendHTML() {
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SML - Enterprise AI Policy Portal</title>
<link rel="icon" href="/logo.jpeg">
<style>
@keyframes fadeIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideDown{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:translateY(0)}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
@keyframes spin{to{transform:rotate(360deg)}}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:linear-gradient(160deg,#eef2f7 0,#dce3ed 100%);color:#1e293b;min-height:100vh;padding:24px}
.container{max-width:920px;margin:0 auto;animation:fadeIn .5s ease-out}

.header{background:linear-gradient(135deg,#8b1a1a 0,#b91c1c 50%,#dc2626 100%);border-radius:18px;padding:28px 36px;margin-bottom:28px;display:flex;align-items:center;gap:24px;box-shadow:0 8px 32px rgba(139,26,26,0.28);position:relative;overflow:hidden}
.header::before{content:'';position:absolute;top:-50%;right:-20%;width:300px;height:300px;background:radial-gradient(circle,rgba(255,255,255,0.04),transparent 70%);border-radius:50%}
.header-logo{width:68px;height:68px;border-radius:14px;background:rgba(255,255,255,0.1);display:flex;align-items:center;justify-content:center;flex-shrink:0;overflow:hidden;backdrop-filter:blur(4px);border:1px solid rgba(255,255,255,0.08);transition:transform .3s ease}
.header-logo:hover{transform:scale(1.04)}
.header-logo img{width:100%;height:100%;object-fit:contain;padding:6px}
.header-text{flex:1;position:relative;z-index:1}
.header-text h1{color:#fff;font-size:1.3rem;font-weight:700;letter-spacing:0.3px;margin-bottom:3px}
.header-text .sub{color:rgba(255,255,255,0.65);font-size:0.82rem;font-weight:400}
.header-badge{background:rgba(255,255,255,0.12);color:#fff;padding:5px 14px;border-radius:20px;font-size:0.7rem;font-weight:700;letter-spacing:0.8px;border:1px solid rgba(255,255,255,0.2);white-space:nowrap;position:relative;z-index:1}

.card{background:#fff;border-radius:18px;box-shadow:0 2px 20px rgba(0,0,0,0.05);overflow:hidden;margin-bottom:24px;animation:fadeIn .5s ease-out;animation-fill-mode:both}
.card:nth-child(2){animation-delay:.1s}
.card:nth-child(3){animation-delay:.2s}
.card-header{padding:18px 26px 14px;border-bottom:1px solid #edf2f7;display:flex;align-items:center;gap:10px}
.card-header h2{font-size:1.05rem;font-weight:700;color:#8b1a1a}
.card-header .count{background:#eef2f7;color:#475569;padding:2px 10px;border-radius:10px;font-size:0.75rem;font-weight:600}
.card-body{padding:22px 26px 26px}

.meta-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:22px}
.meta-item{background:#f8fafc;border-radius:10px;padding:10px 14px;border:1px solid #edf2f7;transition:border .2s}
.meta-item:hover{border-color:#cbd5e1}
.meta-item .label{font-size:0.65rem;text-transform:uppercase;letter-spacing:0.6px;color:#64748b;font-weight:700}
.meta-item .value{font-size:0.92rem;color:#0b2450;font-weight:600;margin-top:2px}

.warning-banner{background:linear-gradient(135deg,#fffbeb,#fef3c7);border:1px solid #fcd34d;border-radius:12px;padding:12px 18px;margin-bottom:22px;display:flex;align-items:flex-start;gap:10px;font-size:0.88rem;color:#92400e;animation:slideDown .3s ease-out}
.warning-banner .icon{font-size:1.1rem;flex-shrink:0;margin-top:1px}

.pdf-wrapper{border:2px solid #e2e8f0;border-radius:12px;overflow:hidden;background:#f8fafc;margin-bottom:0}
.pdf-container{width:100%;max-height:600px;overflow-y:auto;background:#fafafa;position:relative;scroll-behavior:smooth}
.pdf-container .pdf-loading{padding:60px;text-align:center}
.pdf-container .pdf-loading .spinner{width:32px;height:32px;border:3px solid #e2e8f0;border-top-color:#0b2450;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 12px}
.pdf-container .pdf-loading span{display:block;color:#94a3b8;font-size:0.85rem}
.pdf-container canvas{display:block;margin:0 auto;width:100%;height:auto;max-width:680px;box-shadow:0 1px 3px rgba(0,0,0,0.05);background:#fff}
.pdf-container .pdf-page-wrap{position:relative;margin:0;border-bottom:1px solid #edf2f7;animation:fadeIn .3s ease-out}
.pdf-container .pdf-page-num{position:absolute;bottom:5px;right:12px;font-size:0.65rem;color:#94a3b8;background:rgba(255,255,255,0.9);padding:2px 8px;border-radius:4px;font-weight:500;backdrop-filter:blur(2px)}

.unlock-bar{background:#f1f5f9;padding:12px 20px;border-top:1px solid #e2e8f0;display:flex;align-items:center;gap:10px;font-size:0.88rem;transition:background .3s}
.unlock-bar:has(input:checked){background:#f0fdf4}
.unlock-bar .check-wrap{position:relative;width:20px;height:20px;flex-shrink:0}
.unlock-bar .check-wrap input{position:absolute;opacity:0;width:100%;height:100%;z-index:2}
.unlock-bar .check-wrap input:disabled{cursor:default}
.unlock-bar .check-wrap .box{width:20px;height:20px;border:2px solid #94a3b8;border-radius:5px;display:flex;align-items:center;justify-content:center;transition:all .25s}
.unlock-bar .check-wrap input:checked+.box{background:#166534;border-color:#166534}
.unlock-bar .check-wrap input:checked:disabled+.box{background:#166534;border-color:#166534;opacity:1}
.unlock-bar .check-wrap input:checked+.box::after{content:'✓';color:#fff;font-size:13px;font-weight:700}
.unlock-bar label{user-select:none}
.unlock-bar label{cursor:pointer;font-weight:600;color:#1e293b;user-select:none}
.unlock-bar label small{font-weight:400;color:#64748b;font-size:0.8rem}

.form-section h2{font-size:1.05rem;font-weight:700;color:#0b2450;margin-bottom:4px}
.form-section>p{font-size:0.88rem;color:#64748b;margin-bottom:18px;line-height:1.5}
.form-group{margin-bottom:16px}
.form-group label{display:block;font-weight:600;font-size:0.88rem;margin-bottom:5px;color:#1e293b}
.form-group .input-wrap{position:relative}
.form-group .input-wrap input{width:100%;padding:11px 14px;border:1.5px solid #d1d9e6;border-radius:10px;font-size:0.92rem;transition:all .25s;outline:none;background:#fff}
.form-group .input-wrap input:focus{border-color:#8b1a1a;box-shadow:0 0 0 3px rgba(139,26,26,0.1)}
.form-group .input-wrap input:user-invalid{border-color:#dc2626;box-shadow:0 0 0 3px rgba(220,38,38,0.08)}
.form-group .input-wrap .focus-ring{position:absolute;inset:0;border-radius:10px;pointer-events:none;transition:box-shadow .25s}
.form-group .input-wrap input:focus~.focus-ring{box-shadow:0 0 0 3px rgba(139,26,26,0.1)}

.section-label{margin-top:18px;margin-bottom:10px;font-weight:700;font-size:0.88rem;color:#8b1a1a;display:flex;align-items:center;gap:8px}
.section-label::after{content:'';flex:1;height:1px;background:#edf2f7}

.checkbox-group{display:flex;align-items:flex-start;gap:12px;padding:12px 16px;border:1px solid #edf2f7;border-radius:10px;margin-bottom:8px;transition:all .2s}
.checkbox-group:hover{background:#fafcff;border-color:#cbd5e1}
.checkbox-group .cb-wrap{position:relative;width:18px;height:18px;flex-shrink:0;margin-top:2px}
.checkbox-group .cb-wrap input{position:absolute;opacity:0;cursor:pointer;width:100%;height:100%;z-index:2}
.checkbox-group .cb-wrap .cb-box{width:18px;height:18px;border:2px solid #cbd5e1;border-radius:4px;display:flex;align-items:center;justify-content:center;transition:all .2s}
.checkbox-group .cb-wrap input:checked+.cb-box{background:#8b1a1a;border-color:#8b1a1a}
.checkbox-group .cb-wrap input:checked+.cb-box::after{content:'✓';color:#fff;font-size:12px;font-weight:700}
.checkbox-group .cb-wrap input:focus-visible+.cb-box{box-shadow:0 0 0 3px rgba(139,26,26,0.15)}
.checkbox-group label{cursor:pointer;font-size:0.88rem;line-height:1.5;color:#334155}
.checkbox-group label strong{color:#0b2450}
.highlight-box{background:#fef2f2;border-color:#fecaca!important}
.highlight-box:hover{background:#fee2e2!important}
.highlight-box .cb-wrap input:checked+.cb-box{background:#dc2626;border-color:#dc2626}

.submit-btn{background:linear-gradient(135deg,#8b1a1a 0,#b91c1c 100%);color:#fff;padding:14px;border:none;border-radius:11px;cursor:pointer;font-size:0.95rem;font-weight:700;width:100%;transition:all .25s;margin-top:4px;display:flex;align-items:center;justify-content:center;gap:8px;position:relative;overflow:hidden}
.submit-btn::after{content:'';position:absolute;inset:0;background:linear-gradient(135deg,transparent 40%,rgba(255,255,255,0.08) 100%);pointer-events:none}
.submit-btn:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 6px 20px rgba(139,26,26,0.25)}
.submit-btn:active:not(:disabled){transform:translateY(0)}
.submit-btn:disabled{background:#cbd5e1;color:#94a3b8;cursor:not-allowed;transform:none;box-shadow:none}
.submit-btn.loading{pointer-events:none}
.submit-btn .spinner{width:18px;height:18px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:spin .6s linear infinite;display:none}
.submit-btn.loading .spinner{display:block}
.submit-btn.loading .btn-text{opacity:0.7}

.locked-overlay{opacity:0.25;pointer-events:none;user-select:none;transition:opacity .5s ease}
.locked-overlay + .unlock-bar{transition:all .3s}

.footer{text-align:center;padding:16px;color:#94a3b8;font-size:0.78rem}
.verify-row{display:flex;gap:8px}
.verify-input{flex:1;padding:11px 14px;border:1.5px solid #d1d9e6;border-radius:10px;font-size:0.92rem;transition:all .25s;outline:none;background:#fff}
.verify-input:focus{border-color:#8b1a1a;box-shadow:0 0 0 3px rgba(139,26,26,0.1)}
.verify-btn{padding:11px 20px;background:#8b1a1a;color:#fff;border:none;border-radius:10px;font-weight:700;font-size:0.85rem;cursor:pointer;transition:background .2s;white-space:nowrap}
.verify-btn:hover{background:#b91c1c}
.verify-btn:disabled{background:#cbd5e1;cursor:default}
.verify-btn.loading{opacity:0.7}
.verify-status{margin-top:6px;font-size:0.82rem;font-weight:500;min-height:20px}
.verify-status.ok{color:#166534}
.verify-status.err{color:#dc2626}

@media(max-width:640px){
body{padding:14px}
.header{flex-wrap:wrap;padding:20px;gap:14px}
.header-text h1{font-size:1.05rem}
.card-body{padding:16px}
.meta-grid{grid-template-columns:1fr 1fr}
.pdf-container{max-height:450px}
.pdf-container canvas{max-width:100%}
}
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <div class="header-logo">
            <img src="/logo.jpeg" alt="SML Logo">
        </div>
        <div class="header-text">
            <h1>Sanghvi Movers Limited</h1>
            <div class="sub">Enterprise AI Governance &amp; License Management Policy</div>
        </div>
        <div class="header-badge">CONFIDENTIAL</div>
    </div>

    <div class="card">
        <div class="card-header">
            <h2>Policy Document</h2>
            <span class="count">v1.0</span>
        </div>
        <div class="card-body">
            <div class="meta-grid">
                <div class="meta-item"><div class="label">Version</div><div class="value">1.0</div></div>
                <div class="meta-item"><div class="label">Effective Date</div><div class="value">08 June 2026</div></div>
                <div class="meta-item"><div class="label">Applicable To</div><div class="value">SML, SFRPL, SLPL, SMME</div></div>
                <div class="meta-item"><div class="label">Classification</div><div class="value" style="color:#b91c1c">CONFIDENTIAL</div></div>
            </div>

            <div class="warning-banner" id="warningBanner">
                <span>Please read the policy below. Check <strong>"I have read the policy"</strong> once done to unlock the acknowledgement form.</span>
            </div>

            <div class="pdf-wrapper" id="pdfWrapper">
                <div id="pdfContainer" class="pdf-container">
                    <div class="pdf-loading">
                        <div class="spinner"></div>
                        <span>Loading policy document&hellip;</span>
                    </div>
                </div>
                <div class="unlock-bar">
                    <div class="check-wrap">
                        <input type="checkbox" id="readVerification" disabled>
                        <div class="box"></div>
                    </div>
                    <label>I have read and understood the policy document <small>(scroll to bottom)</small></label>
                </div>
            </div>
        </div>
    </div>

    <div class="card" id="formCard">
        <div class="card-header">
            <h2>User Acknowledgement</h2>
        </div>
        <div class="card-body">
            <form id="ackForm" action="/api/submit" method="POST" class="form-section locked-overlay">
                <p>Enter your official email to verify your identity. Your name will be auto-filled if found.</p>
                <div class="form-group">
                    <label for="empEmail">Official Email ID *</label>
                    <div class="verify-row">
                        <input type="email" id="empEmail" required placeholder="your.email@sanghviglobal.com" autocomplete="email" class="verify-input">
                        <button type="button" id="verifyBtn" class="verify-btn">Verify</button>
                    </div>
                    <div id="verifyStatus" class="verify-status"></div>
                </div>
                <div class="form-group" id="nameGroup" style="display:none">
                    <label for="empName">Employee Name</label>
                    <div class="input-wrap">
                        <input type="text" id="empName" name="empName" readonly placeholder="Verified employee name" style="background:#f1f5f9;cursor:default;color:#1e293b">
                        <div class="focus-ring"></div>
                    </div>
                </div>
                <input type="hidden" name="empEmail" id="hiddenEmail" value="">

                <div class="section-label">Policy Acknowledgement Items</div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q1" name="q1" value="Yes" required><div class="cb-box"></div></div><label for="q1"><strong>1.</strong> My access to enterprise AI platforms is conditional upon satisfying the requirements of the Dual Scoring Framework and maintaining active, productive utilisation throughout the period of my license.</label></div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q2" name="q2" value="Yes" required><div class="cb-box"></div></div><label for="q2"><strong>2.</strong> I am solely and personally responsible for every AI-generated output that I generate, review, and distribute. The involvement of an AI platform does not transfer or reduce my accountability for the content.</label></div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q3" name="q3" value="Yes" required><div class="cb-box"></div></div><label for="q3"><strong>3.</strong> I will not share, distribute, or submit any AI-generated content internally or externally without first completing the mandatory human review prescribed by this policy.</label></div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q4" name="q4" value="Yes" required><div class="cb-box"></div></div><label for="q4"><strong>4.</strong> I will not enter, upload, or reference sensitive information in any AI platform prompt, as defined in Section 7.2 of this policy.</label></div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q5" name="q5" value="Yes" required><div class="cb-box"></div></div><label for="q5"><strong>5.</strong> My license may be subject to reallocation in the event that I fail to maintain the utilisation thresholds prescribed in Section 6, without regard to my seniority or organisational position.</label></div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q6" name="q6" value="Yes" required><div class="cb-box"></div></div><label for="q6"><strong>6.</strong> I will use the allocated AI platform tokens judiciously and responsibly, and will manage my usage in line with the prescribed token limits.</label></div>
                <div class="checkbox-group"><div class="cb-wrap"><input type="checkbox" id="q7" name="q7" value="Yes" required><div class="cb-box"></div></div><label for="q7"><strong>7.</strong> I understand that breach of this policy may result in disciplinary action in accordance with applicable company policies and legal requirements.</label></div>
                <input type="hidden" name="read_policy" id="readPolicyInput" value="No">
                <input type="hidden" name="client_ip" id="clientIpInput" value="">

                <button type="submit" class="submit-btn" id="submitBtn" disabled>
                    <span class="spinner"></span>
                    <span class="btn-text">Submit Acknowledgement</span>
                </button>
            </form>
        </div>
    </div>

    <div class="footer">&copy; ${new Date().getFullYear()} Sanghvi Movers Limited &mdash; Confidential</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
(function() {
const pdfContainer = document.getElementById('pdfContainer');
const readVerification = document.getElementById('readVerification');
const ackForm = document.getElementById('ackForm');
const submitBtn = document.getElementById('submitBtn');
const warningBanner = document.getElementById('warningBanner');
const nameInput = document.getElementById('empName');
const nameGroup = document.getElementById('nameGroup');
const readPolicyInput = document.getElementById('readPolicyInput');
const empEmailInput = document.getElementById('empEmail');
const verifyBtn = document.getElementById('verifyBtn');
const verifyStatus = document.getElementById('verifyStatus');
const hiddenEmail = document.getElementById('hiddenEmail');

let verifiedName = '';

verifyBtn.addEventListener('click', function() {
    const email = empEmailInput.value.trim();
    if (!email) {
        verifyStatus.className = 'verify-status err';
        verifyStatus.textContent = 'Please enter your email.';
        return;
    }
    verifyBtn.disabled = true;
    verifyBtn.textContent = 'Checking...';
    verifyStatus.className = 'verify-status';
    verifyStatus.textContent = '';
    fetch('/api/verify-email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: email })
    }).then(function(r){return r.json()}).then(function(data){
        if (data.found) {
            verifiedName = data.name;
            nameInput.value = data.name;
            nameGroup.style.display = '';
            hiddenEmail.value = email;
            verifyStatus.className = 'verify-status ok';
            verifyStatus.textContent = 'Verified: ' + data.name;
            empEmailInput.readOnly = true;
            verifyBtn.textContent = 'Verified';
            verifyBtn.disabled = true;
        } else {
            verifyStatus.className = 'verify-status err';
            verifyStatus.textContent = data.error || 'Email not found.';
            verifyBtn.disabled = false;
            verifyBtn.textContent = 'Verify';
        }
    }).catch(function(){
        verifyStatus.className = 'verify-status err';
        verifyStatus.textContent = 'Connection error. Try again.';
        verifyBtn.disabled = false;
        verifyBtn.textContent = 'Verify';
    });
});

empEmailInput.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        verifyBtn.click();
    }
});

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

pdfjsLib.getDocument('/policy.pdf').promise.then(function(pdf) {
    pdfContainer.innerHTML = '';
    let pageNum = 1;
    function renderPage() {
        if (pageNum > pdf.numPages) return;
        pdf.getPage(pageNum).then(function(page) {
            const viewport = page.getViewport({ scale: 1.5 });
            const wrap = document.createElement('div');
            wrap.className = 'pdf-page-wrap';
            const canvas = document.createElement('canvas');
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            wrap.appendChild(canvas);
            const label = document.createElement('div');
            label.className = 'pdf-page-num';
            label.textContent = pageNum + ' / ' + pdf.numPages;
            wrap.appendChild(label);
            pdfContainer.appendChild(wrap);
            page.render({ canvasContext: canvas.getContext('2d'), viewport: viewport }).promise.then(function() {
                pageNum++;
                renderPage();
                if (pageNum === pdf.numPages + 1) {
                    pdfContainer.scrollTop = 0;
                    pdfContainer.addEventListener('scroll', onScrollToBottom);
                    if (pdfContainer.scrollHeight <= pdfContainer.clientHeight + 10) {
                        autoCheckRead();
                    }
                }
            });
        });
    }
    renderPage();

    let autoChecked = false;
    function onScrollToBottom() {
        if (autoChecked) return;
        if (pdfContainer.scrollTop + pdfContainer.clientHeight >= pdfContainer.scrollHeight - 40) {
            autoCheckRead();
        }
    }
    function autoCheckRead() {
        if (autoChecked) return;
        autoChecked = true;
        readVerification.checked = true;
        readVerification.dispatchEvent(new Event('change'));
        pdfContainer.removeEventListener('scroll', onScrollToBottom);
    }
}).catch(function() {
    pdfContainer.innerHTML = '<div class="pdf-loading" style="padding:40px">Failed to load PDF. <a href="/policy.pdf" style="color:#0b2450;font-weight:600">Open directly</a></div>';
});

readVerification.addEventListener('change', function() {
    readPolicyInput.value = this.checked ? 'Yes' : 'No';
    if(this.checked) {
        ackForm.classList.remove('locked-overlay');
        submitBtn.removeAttribute('disabled');
        warningBanner.style.display = 'none';
        setTimeout(function() {
            document.getElementById('formCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 300);
    } else {
        ackForm.classList.add('locked-overlay');
        submitBtn.setAttribute('disabled', 'true');
        warningBanner.style.display = '';
    }
});

ackForm.addEventListener('submit', function(e) {
    if (!verifiedName) {
        verifyStatus.className = 'verify-status err';
        verifyStatus.textContent = 'Please verify your email first.';
        e.preventDefault();
        return;
    }
    submitBtn.classList.add('loading');
    submitBtn.disabled = true;
});

const clientIpInput = document.getElementById('clientIpInput');
fetch('https://api.ipify.org?format=json').then(function(r){return r.json()}).then(function(d){clientIpInput.value=d.ip}).catch(function(){});
})();
</script>
</body>
</html>`;
}

function getLoginHTML(error) {
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Login - SML Policy Portal</title>
<link rel="icon" href="/logo.jpeg">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:linear-gradient(135deg,#8b1a1a 0,#b91c1c 100%);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.login-card{background:#fff;border-radius:20px;padding:45px 40px 40px;width:100%;max-width:400px;box-shadow:0 20px 60px rgba(0,0,0,0.3)}
.logo-area{text-align:center;margin-bottom:30px}
.logo-area img{width:56px;height:56px;border-radius:10px;margin-bottom:12px;background:#f1f5f9;padding:4px}
.logo-area h1{font-size:1.15rem;color:#8b1a1a;font-weight:700}
.logo-area p{font-size:0.8rem;color:#64748b;margin-top:2px}
h2{font-size:1rem;color:#334155;margin-bottom:20px;text-align:center;font-weight:600}
.error{background:#fef2f2;color:#b91c1c;padding:10px 14px;border-radius:10px;font-size:0.85rem;margin-bottom:16px;text-align:center;border:1px solid #fecaca}
.form-group{margin-bottom:16px}
.form-group label{display:block;font-size:0.85rem;font-weight:600;color:#475569;margin-bottom:5px}
.form-group input{width:100%;padding:11px 14px;border:1.5px solid #d1d9e6;border-radius:10px;font-size:0.95rem;outline:none;transition:border .2s,box-shadow .2s}
.form-group input:focus{border-color:#8b1a1a;box-shadow:0 0 0 3px rgba(139,26,26,0.1)}
.login-btn{background:linear-gradient(135deg,#8b1a1a 0,#b91c1c 100%);color:#fff;padding:12px;border:none;border-radius:10px;cursor:pointer;font-size:0.95rem;font-weight:700;width:100%;transition:opacity .2s}
.login-btn:hover{opacity:0.9}
.hint{text-align:center;margin-top:16px;font-size:0.75rem;color:#94a3b8}
</style>
</head>
<body>
<div class="login-card">
    <div class="logo-area">
        <img src="/logo.jpeg" alt="SML">
        <h1>Sanghvi Movers Limited</h1>
        <p>Admin Control Panel</p>
    </div>
    <h2>Sign in to continue</h2>
    ${error ? `<div class="error">&#10060; ${error}</div>` : ''}
    <form method="POST" action="/admin/login">
        <div class="form-group">
            <label for="username">Username</label>
            <input type="text" id="username" name="username" required placeholder="Enter admin username" autocomplete="username">
        </div>
        <div class="form-group">
            <label for="password">Password</label>
            <input type="password" id="password" name="password" required placeholder="Enter admin password" autocomplete="current-password">
        </div>
        <button type="submit" class="login-btn">Sign In</button>
    </form>
    <div class="hint">Authorised personnel only</div>
</div>
</body>
</html>`;
}

function getAdminDashboardHTML(completed, pending, msg) {
    const total = completed.length + pending.length;
    const pct = total ? Math.round((completed.length / total) * 100) : 0;

    const renderRows = (list, isPending) => list.map(r => {
        const detail = JSON.stringify({
            id: r.emp_id,
            name: r.name,
            dept: r.department || '',
            email: r.emp_email || r.submitted_email || '',
            empEmail: r.emp_email || '',
            subEmail: r.submitted_email || '',
            ip: r.ip_address || '',
            ua: r.user_agent || '',
            time: r.submitted_at || 'N/A',
            read_policy: r.read_policy || 'No',
            q1: r.q1 || 'No', q2: r.q2 || 'No', q3: r.q3 || 'No', q4: r.q4 || 'No',
            q5: r.q5 || 'No', q6: r.q6 || 'No', q7: r.q7 || 'No'
        }).replace(/'/g, '&#39;');
        return `<tr class="${isPending ? 'row-pending' : 'row-done'}" data-name="${escapeHtml(r.name).toLowerCase()}" onclick="showDetail('${detail.replace(/"/g, '&quot;')}')" style="cursor:pointer">
        <td class="td-id">${r.emp_id}</td>
        <td class="td-name">${escapeHtml(r.name)}</td>
        <td class="td-email">${escapeHtml(r.emp_email || r.submitted_email || '')}</td>
        <td class="td-dept">${escapeHtml(r.department || '')}</td>
        <td class="td-status">${
            isPending
                ? '<span class="badge badge-pending">PENDING</span>'
                : `<span class="badge badge-done">&#10003; ${r.submitted_at}</span>`
        }</td>
        <td class="td-actions">${
            isPending
            ? `<a href="#" onclick="event.stopPropagation();confirmRemove('${r.emp_id}','${escapeHtml(r.name).replace(/'/g, "\\'")}')" class="remove-link">Remove</a>`
            : `<a href="#" onclick="event.stopPropagation();confirmReset('${r.emp_id}','${escapeHtml(r.name).replace(/'/g, "\\'")}')" class="reset-link">Reset</a>`
            + ` &nbsp; <a href="#" onclick="event.stopPropagation();confirmRemove('${r.emp_id}','${escapeHtml(r.name).replace(/'/g, "\\'")}')" class="remove-link">Remove</a>`
        }</td>
    </tr>`;
    }).join('');

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Dashboard - SML Policy Portal</title>
<link rel="icon" href="/logo.jpeg">
<style>
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideDown{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:translateY(0)}}
@keyframes modalIn{from{opacity:0;transform:scale(0.95)}to{opacity:1;transform:scale(1)}}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#eef2f7;color:#1e293b}
.topbar{background:linear-gradient(135deg,#8b1a1a 0,#b91c1c 100%);padding:0 32px;height:64px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 16px rgba(139,26,26,0.2);position:sticky;top:0;z-index:100}
.topbar-left{display:flex;align-items:center;gap:14px}
.topbar-left img{width:34px;height:34px;border-radius:8px;background:rgba(255,255,255,0.1);padding:3px}
.topbar-left h1{color:#fff;font-size:0.95rem;font-weight:600;letter-spacing:0.2px}
.topbar-right{display:flex;align-items:center;gap:10px}
.topbar-btn{background:rgba(255,255,255,0.1);color:rgba(255,255,255,0.85);padding:7px 14px;border-radius:8px;text-decoration:none;font-size:0.78rem;font-weight:600;transition:all .2s;display:inline-flex;align-items:center;gap:5px;border:none;cursor:pointer}
.topbar-btn:hover{background:rgba(255,255,255,0.2);color:#fff}
.dashboard{max-width:1120px;margin:0 auto;padding:28px 20px;animation:fadeIn .4s ease-out}
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:24px}
.stat-card{background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 2px 10px rgba(0,0,0,0.03);border:1px solid #edf2f7;transition:transform .2s,box-shadow .2s}
.stat-card:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,0.06)}
.stat-card .num{font-size:1.7rem;font-weight:800;color:#8b1a1a;line-height:1.2}
.stat-card .label{font-size:0.78rem;color:#64748b;font-weight:500;margin-top:2px}
.stat-card.green .num{color:#166534}
.stat-card.amber .num{color:#92400e}
.stat-card .bar-wrap{background:#edf2f7;border-radius:20px;height:5px;margin-top:8px;overflow:hidden}
.stat-card .bar-fill{height:100%;border-radius:20px;background:linear-gradient(90deg,#22c55e,#16a34a);transition:width .8s ease}

.panel{background:#fff;border-radius:14px;box-shadow:0 2px 10px rgba(0,0,0,0.03);border:1px solid #edf2f7;overflow:hidden;margin-bottom:18px}
.panel-header{padding:14px 20px;border-bottom:1px solid #edf2f7;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.panel-header-left{display:flex;align-items:center;gap:8px}
.panel-header h2{font-size:0.95rem;font-weight:700;color:#8b1a1a}
.panel-header .badge{background:#edf2f7;color:#475569;padding:2px 8px;border-radius:8px;font-size:0.7rem;font-weight:700}
.search-box{display:flex;align-items:center;gap:6px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:5px 10px;transition:border .2s}
.search-box:focus-within{border-color:#8b1a1a;box-shadow:0 0 0 2px rgba(139,26,26,0.08)}
.search-box input{border:none;background:transparent;outline:none;font-size:0.82rem;padding:3px 0;width:160px;color:#1e293b}
.search-box input::placeholder{color:#94a3b8}
.search-box .icon{color:#94a3b8;font-size:0.8rem}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:0.82rem}
thead th{padding:7px 12px;font-size:0.72rem;color:#333;font-weight:700;background:#e8ecf0;border:1px solid #c4c9ce;text-align:left;white-space:nowrap;position:sticky;top:0;z-index:1}
tbody td{padding:6px 12px;border:1px solid #d4d9de;color:#1e293b;background:#fff}
tbody tr:nth-child(even) td{background:#f5f6f8}
tbody tr:hover td{background:#dce6f2!important}
.row-done td{color:#555}
.td-id{font-size:0.78rem;color:#555;font-family:ui-monospace,monospace;white-space:nowrap}
.td-name{font-weight:500}
.td-email{font-size:0.78rem;color:#475569}
.td-dept{font-size:0.78rem;color:#666}
.td-status{}
.td-actions{text-align:center}
.reset-link{color:#b91c1c;font-size:0.72rem;font-weight:700;text-decoration:none;padding:3px 8px;border-radius:5px;transition:background .2s}
.reset-link:hover{background:#fef2f2;text-decoration:underline}
.remove-link{color:#64748b;font-size:0.72rem;font-weight:600;text-decoration:none;padding:3px 8px;border-radius:5px;transition:all .2s}
.remove-link:hover{background:#fef2f2;color:#b91c1c;text-decoration:underline}
.badge{padding:3px 10px;border-radius:20px;font-size:0.7rem;font-weight:700;display:inline-block;white-space:nowrap}
.badge-pending{background:#fef3c7;color:#92400e}
.badge-done{background:#dcfce7;color:#166534}
.placeholder{text-align:center;padding:36px 20px;color:#94a3b8;font-size:0.88rem}
.placeholder strong{color:#166534;font-size:1rem}
.msg-banner{background:#f0fdf4;color:#166534;padding:10px 18px;border-radius:10px;font-size:0.85rem;font-weight:600;margin-bottom:18px;border:1px solid #bbf7d0;animation:slideDown .3s ease-out}

.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.4);backdrop-filter:blur(4px);z-index:1000;align-items:center;justify-content:center;padding:20px}
.modal-overlay.active{display:flex}
.modal{background:#fff;border-radius:18px;max-width:560px;width:100%;max-height:90vh;overflow-y:auto;animation:modalIn .25s ease-out;box-shadow:0 20px 60px rgba(0,0,0,0.2)}
.modal-head{padding:20px 24px 14px;border-bottom:1px solid #edf2f7;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:#fff;z-index:1}
.modal-head h2{font-size:1.05rem;font-weight:700;color:#8b1a1a}
.modal-close{width:32px;height:32px;border:none;background:#f1f5f9;border-radius:8px;cursor:pointer;font-size:1.1rem;display:flex;align-items:center;justify-content:center;transition:background .2s}
.modal-close:hover{background:#e2e8f0}
.modal-body{padding:20px 24px 24px}
.detail-grid{display:grid;gap:12px}
.detail-row{display:flex;align-items:flex-start;gap:10px;padding:10px 14px;border-radius:10px;background:#f8fafc;border:1px solid #edf2f7}
.detail-row .label{font-size:0.78rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.3px;min-width:130px;flex-shrink:0;padding-top:1px}
.detail-row .value{font-size:0.88rem;color:#1e293b}
.detail-row .value.yes{color:#166534;font-weight:600}
.detail-row .value.no{color:#94a3b8}
.detail-row.sub-head{background:#8b1a1a;color:#fff;border:none;font-weight:700;font-size:0.85rem;text-transform:uppercase;letter-spacing:0.5px}
.detail-row.sub-head .label{color:rgba(255,255,255,0.6);text-transform:none}
.footer{text-align:center;padding:16px;color:#94a3b8;font-size:0.75rem}
@media(max-width:640px){
.topbar{padding:0 14px}
.topbar-left h1{font-size:0.82rem}
.dashboard{padding:16px 10px}
.stats-grid{grid-template-columns:1fr 1fr}
.search-box input{width:100px}
.modal{margin:10px}
.detail-row{flex-wrap:wrap}
.detail-row .label{min-width:100px}
}
</style>
</head>
<body>
<div class="topbar">
    <div class="topbar-left">
        <img src="/logo.jpeg" alt="SML">
        <h1>AI Governance &middot; Admin Panel</h1>
    </div>
    <div class="topbar-right">
        <button class="topbar-btn add" onclick="showAddUser()">&#10010; Add User</button>
        <form method="POST" action="/admin/sync-roster" style="display:inline">
            <button type="submit" class="topbar-btn sync">&#128260; Sync Roster</button>
        </form>
        <a href="/admin/export" class="topbar-btn export">&#128229; Export Compliance</a>
        <a href="/admin/export-users" class="topbar-btn export-users">&#128230; Export Users</a>
        <a href="/admin/logout" class="topbar-btn">&#10140; Logout</a>
    </div>
</div>

<div class="dashboard">
    <div class="stats-grid">
        <div class="stat-card"><div class="num">${total}</div><div class="label">Total Employees</div></div>
        <div class="stat-card green"><div class="num">${completed.length}</div><div class="label">Compliant</div><div class="bar-wrap"><div class="bar-fill" style="width:${pct}%"></div></div></div>
        <div class="stat-card amber"><div class="num">${pending.length}</div><div class="label">Pending</div><div class="bar-wrap"><div class="bar-fill" style="width:${100-pct}%;background:linear-gradient(90deg,#f59e0b,#d97706)"></div></div></div>
        <div class="stat-card"><div class="num">${pct}%</div><div class="label">Completion Rate</div><div class="bar-wrap"><div class="bar-fill" style="width:${pct}%"></div></div></div>
    </div>

    ${msg ? '<div class="msg-banner">&#10003; ' + msg + '</div>' : ''}

    <div class="panel">
        <div class="panel-header">
            <div class="panel-header-left">
                <h2>All Employees</h2>
                <span class="badge">${total}</span>
            </div>
            <div class="search-box">
                <span class="icon">&#128269;</span>
                <input type="text" id="searchInput" placeholder="Search by name or ID..." oninput="filterTable(this.value)">
            </div>
        </div>
        <div class="table-wrap">
            <p style="padding:6px 14px;font-size:0.72rem;color:#94a3b8;border-bottom:1px solid #edf2f7">Click a row to view full submission details</p>
            <table>
                <thead><tr><th>Emp ID</th><th>Name</th><th>Email</th><th>Department</th><th>Status</th><th>Actions</th></tr></thead>
                <tbody id="tableBody">
                    ${pending.length ? renderRows(pending, true) : ''}
                    ${completed.length ? renderRows(completed, false) : ''}
                    ${!pending.length && !completed.length ? '<tr><td colspan="6" class="placeholder">No employees found.</td></tr>' : ''}
                </tbody>
            </table>
        </div>
    </div>
</div>

<div class="modal-overlay" id="modalOverlay" onclick="if(event.target===this)closeDetail()">
    <div class="modal" id="detailModal">
        <div class="modal-head">
            <h2 id="modalTitle">Employee Details</h2>
            <button class="modal-close" onclick="closeDetail()">&#10005;</button>
        </div>
        <div class="modal-body" id="modalBody"></div>
    </div>
</div>

<div class="modal-overlay" id="addUserOverlay" onclick="if(event.target===this)hideAddUser()">
    <div class="modal" style="max-width:420px">
        <div class="modal-head">
            <h2>Add New Employee</h2>
            <button class="modal-close" onclick="hideAddUser()">&#10005;</button>
        </div>
        <div class="modal-body">
            <form method="POST" action="/admin/add-employee" id="addUserForm">
                <div style="margin-bottom:14px">
                    <label style="display:block;font-size:0.82rem;font-weight:600;color:#475569;margin-bottom:4px">Full Name *</label>
                    <input type="text" name="name" required placeholder="e.g. John Doe" style="width:100%;padding:10px 12px;border:1.5px solid #d1d9e6;border-radius:10px;font-size:0.9rem;outline:none;transition:border .2s" onfocus="this.style.borderColor='#8b1a1a'" onblur="this.style.borderColor='#d1d9e6'">
                </div>
                <div style="margin-bottom:14px">
                    <label style="display:block;font-size:0.82rem;font-weight:600;color:#475569;margin-bottom:4px">Email ID</label>
                    <input type="email" name="email" placeholder="e.g. john.doe@sanghviglobal.com" style="width:100%;padding:10px 12px;border:1.5px solid #d1d9e6;border-radius:10px;font-size:0.9rem;outline:none;transition:border .2s" onfocus="this.style.borderColor='#8b1a1a'" onblur="this.style.borderColor='#d1d9e6'">
                </div>
                <div style="margin-bottom:18px">
                    <label style="display:block;font-size:0.82rem;font-weight:600;color:#475569;margin-bottom:4px">Department</label>
                    <input type="text" name="department" placeholder="e.g. Operations" style="width:100%;padding:10px 12px;border:1.5px solid #d1d9e6;border-radius:10px;font-size:0.9rem;outline:none;transition:border .2s" onfocus="this.style.borderColor='#8b1a1a'" onblur="this.style.borderColor='#d1d9e6'">
                </div>
                <button type="submit" style="background:linear-gradient(135deg,#8b1a1a 0,#b91c1c 100%);color:#fff;padding:12px;border:none;border-radius:10px;cursor:pointer;font-size:0.9rem;font-weight:700;width:100%;transition:opacity .2s" onmouseover="this.style.opacity='0.9'" onmouseout="this.style.opacity='1'">Add Employee</button>
            </form>
        </div>
    </div>
</div>

<div class="footer">&copy; ${new Date().getFullYear()} Sanghvi Movers Limited &mdash; Confidential</div>

<script>
function confirmReset(empId, empName) {
    if (confirm('Reset "' + empName + '" so they can resubmit? This cannot be undone.')) {
        var form = document.createElement('form');
        form.method = 'POST';
        form.action = '/admin/reset-submission';
        form.innerHTML = '<input name="emp_id" value="' + empId + '"><input name="name" value="' + empName + '">';
        document.body.appendChild(form);
        form.submit();
    }
}

function confirmRemove(empId, empName) {
    var pw = prompt('Enter admin password to remove "' + empName + '":');
    if (pw) {
        var form = document.createElement('form');
        form.method = 'POST';
        form.action = '/admin/remove-employee';
        form.innerHTML = '<input name="emp_id" value="' + empId + '"><input name="name" value="' + empName + '"><input name="password" value="' + pw + '">';
        document.body.appendChild(form);
        form.submit();
    }
}

function filterTable(val) {
    const q = val.toLowerCase().trim();
    document.querySelectorAll('#tableBody tr').forEach(function(tr) {
        const name = tr.getAttribute('data-name') || '';
        const id = (tr.querySelector('.td-id')?.textContent || '').toLowerCase();
        const dept = (tr.querySelector('.td-dept')?.textContent || '').toLowerCase();
        const email = (tr.querySelector('.td-email')?.textContent || '').toLowerCase();
        tr.style.display = (!q || name.includes(q) || id.includes(q) || dept.includes(q) || email.includes(q)) ? '' : 'none';
    });
}

function showDetail(jsonStr) {
    const data = JSON.parse(jsonStr);
    const overlay = document.getElementById('modalOverlay');
    document.getElementById('modalTitle').textContent = data.name + ' (' + data.id + ')';
    const items = [
        { label: 'Employee', value: data.name + ' (' + data.id + ')' },
        { label: 'Department', value: data.dept || '—' },
        { label: 'Roster Email', value: data.empEmail || '—' },
        { label: 'Submitted Email', value: data.subEmail || '—' },
        { label: 'Status', value: data.time === 'N/A' ? 'PENDING' : 'COMPLIANT', cls: data.time === 'N/A' ? '' : 'yes' },
        { label: 'Submitted At', value: data.time },
        { label: 'IP Address', value: data.ip || '—' },
        { label: 'Device', value: data.ua ? (data.ua.substring(0, 100) + (data.ua.length > 100 ? '...' : '')) : '—' },
        { label: 'Read Policy Doc', value: data.read_policy, cls: data.read_policy === 'Yes' ? 'yes' : 'no' },
        { label: '', subhead: 'Acknowledgement Items', value: '' },
        { label: '1. Dual Scoring Framework', value: data.q1, cls: data.q1 === 'Yes' ? 'yes' : 'no' },
        { label: '2. Personal Accountability', value: data.q2, cls: data.q2 === 'Yes' ? 'yes' : 'no' },
        { label: '3. Human Review Required', value: data.q3, cls: data.q3 === 'Yes' ? 'yes' : 'no' },
        { label: '4. No Sensitive Info in Prompts', value: data.q4, cls: data.q4 === 'Yes' ? 'yes' : 'no' },
        { label: '5. License Reallocation', value: data.q5, cls: data.q5 === 'Yes' ? 'yes' : 'no' },
        { label: '6. Token Usage', value: data.q6, cls: data.q6 === 'Yes' ? 'yes' : 'no' },
        { label: '7. Disciplinary Action', value: data.q7, cls: data.q7 === 'Yes' ? 'yes' : 'no' }
    ];
    document.getElementById('modalBody').innerHTML = '<div class="detail-grid">' + items.map(function(i) {
        if (i.subhead) {
            return '<div class="detail-row sub-head" style="padding:8px 14px"><span>' + i.subhead + '</span></div>';
        }
        var valClass = i.cls ? 'value ' + i.cls : 'value';
        return '<div class="detail-row"><span class="label">' + i.label + '</span><span class="' + valClass + '">' + i.value + '</span></div>';
    }).join('') + '</div>';
    overlay.classList.add('active');
}

function closeDetail() {
    document.getElementById('modalOverlay').classList.remove('active');
}
function showAddUser() {
    document.getElementById('addUserOverlay').classList.add('active');
}
function hideAddUser() {
    document.getElementById('addUserOverlay').classList.remove('active');
}
document.addEventListener('keydown', function(e) { if(e.key === 'Escape') { closeDetail(); hideAddUser(); } });
</script>
</body>
</html>`;
}
