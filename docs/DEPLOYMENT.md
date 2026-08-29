# Deployment Guide

## Google Apps Script Deployment

### 1. Open the Apps Script project
Open the Google Apps Script project connected to the Sansis Daily Officer Google Sheet.

### 2. Update source files
Replace:
- `Code.gs`
- `Index.html`

with the latest source from this repository.

### 3. Run setup when required
After a schema-changing release, run:

```javascript
setupSansisDailyOfficer()
```

For the first Administrator account, if it has not been created yet, run:

```javascript
createInitialAdministratorLogin()
```

Check **Execution Log** for the generated temporary Administrator credential.

### 4. Deploy a new Web App version
1. Click **Deploy**
2. Select **Manage deployments**
3. Open the current Web App deployment
4. Click **Edit**
5. Select **New version**
6. Click **Deploy**

### 5. Test
Test:
- Employee login/logout
- Administrator login/logout
- Attendance
- Job Report
- Task Manager
- KPI
- Employee Finance
- Employee Management
- Attendance Monitoring
- Administrator Finance

## GitHub vs Production
GitHub is the source/version repository. Google Apps Script is the runtime/deployment environment. A GitHub commit alone does not change the live application.
