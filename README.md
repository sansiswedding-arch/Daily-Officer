# Sansis Daily Officer

Internal employee and administrator workspace for **Sansis Wedding**, built with **Google Apps Script** and **Google Sheets**.

## Current Version

**V8.6 — Administrator Finance**

## Main Modules

### Employee App
- Dashboard
- Attendance
- Job Report
- Task Manager
- KPI
- Finance
- Settings
- Login / Session

### Administrator App
- Dashboard
- Employee Management
- Attendance Monitoring
- Finance
- Login / Session

## Technology
- Google Apps Script
- Google Sheets
- HTML / CSS / JavaScript
- Google Drive for uploaded evidence and documentation

## Repository Structure

```text
sansis-daily-officer/
├── Code.gs
├── Index.html
├── README.md
├── CHANGELOG.md
├── SECURITY.md
├── .gitignore
└── docs/
    ├── DEPLOYMENT.md
    └── VERSION_V8.6.md
```

## Important

GitHub stores the **source code and version history**. The production application is still deployed through **Google Apps Script Web App**, while **Google Sheets** is used as the main data source.

Updating files in GitHub does **not automatically update** the deployed Google Apps Script application unless a deployment workflow is configured separately.

## Deployment

See [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md).

## Security

Keep this repository **Private**.

Never commit:
- Employee passwords
- Administrator passwords
- Session tokens
- Personal employee data exports
- Google credentials or API secrets

See [SECURITY.md](SECURITY.md).

## Versioning

Recommended commit format:

```text
V8.6 - Add Administrator Finance
V8.7 - Add Job Report Monitoring
V8.8 - Add Task Management
```

Use GitHub commits as the main development history instead of keeping duplicate source files for every version inside the repository.
