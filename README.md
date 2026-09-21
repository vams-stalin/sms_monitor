# SMS Monitor — Internal Dashboard

Monitors SMS credit usage across clients in real time. Detects spikes, drops, and threshold breaches, and sends email alerts automatically.

---

## Running the app

### 1. Install dependencies
```
pip install -r requirements.txt
```

### 2. Install the SQL Server ODBC driver (if not already installed)
Download **ODBC Driver 18 for SQL Server** from Microsoft and install it.
The app will not connect to the database without it.

### 3. Start the server
```
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 1
```
> **Important:** Always use `--workers 1`. Multiple workers cause duplicate scheduled jobs (hourly sync, daily email report, midnight flag reset).

Open your browser at: **http://localhost:8000**

---

## Login credentials

| Role    | Username   | Password  | Access                          |
|---------|------------|-----------|---------------------------------|
| Admin   | admin      | ADMIN123  | Full access — can assign credits, manage users, trigger sync |
| Viewer  | testviewer | ViewerTest1 | Read-only — dashboard and analytics only |

---

## Configuration files

| File        | Purpose                                      |
|-------------|----------------------------------------------|
| `.env`      | Database connection, JWT secret, app settings |
| `email.env` | SMTP credentials and alert recipient emails   |

### Changing who receives alert emails
Open `email.env` and update:
```
ALERT_TO=recipient1@company.com, recipient2@company.com
ALERT_CC=manager@company.com
```

### Pointing to a different database
Open `.env` and update:
```
DB_SERVER=your-server-ip
DB_NAME=your-database-name
DB_USERNAME=your-username
DB_PASSWORD=your-password
```
Restart the server after any `.env` change.

---

## How it works

- **Hourly sync** — pulls real SMS counts from `Smslog_Vtb` into the dashboard automatically
- **Spike / Drop detection** — flags clients whose today's usage is abnormally high or low vs their 7-day average
- **Threshold alerts** — emails go out when a client crosses 50% or 90% of their allocated SMS
- **Daily report** — summary email sent every day at 6:00 PM with an Excel attachment

---

## Important notes

- The database it connects to is **VAuthenticateWebLive** (live production data)
- For production deployment, set `FORCE_HTTPS=true` in `.env` and put the app behind an nginx reverse proxy
- `.env` and `email.env` contain credentials — do not commit them to git or share them publicly
