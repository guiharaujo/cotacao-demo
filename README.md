# Excel VBA Quotation System — Azure SQL + CRM Integration

A production-grade sales quotation system built entirely in **Excel VBA**, integrating with **Azure SQL** for data persistence and a **REST CRM API** for opportunity management. Designed for B2B sales teams that operate in Excel but need centralized, auditable data.

---

## Features

- **Role-based authentication** — login form validates credentials against Azure SQL; separate approval flow for management-level discounts
- **Live product catalog** — bulk-loads products and pricing from Azure SQL on login; one-click refresh during the session
- **Automated pricing engine** — minimum sale price calculated dynamically based on cost structure, taxes (ICMS, PIS/COFINS, IPI, DIFAL), freight, commission, and target margin
- **Margin authority levels** — seller, supervisor, and management tiers with automatic escalation prompt
- **Freight management** — CIF/FOB toggle; supports air (standard/express) and road modals with weight/volume auto-calculation
- **Word proposal generation** — COM automation generates a fully formatted `.docx` proposal (no external dependencies); sequential numbering via Azure SQL counter
- **Audit log** — every generated proposal is recorded in `COT_PROPOSTAS` with user, timestamp, margin, and CRM opportunity ID
- **CRM integration** — pulls client and opportunity data directly from the CRM API by opportunity ID
- **10-year warranty toggle** — adds 2% to the tax stack when enabled, reflected in real-time pricing

---

## Architecture

```
┌─────────────────────────────────────┐
│         Excel Workbook (.xlsm)      │
│                                     │
│  ┌──────────┐   ┌────────────────┐  │
│  │  LOGIN   │   │   QUOTATION    │  │
│  │  sheet   │   │   sheet        │  │
│  └────┬─────┘   └───────┬────────┘  │
│       │                 │           │
│  ┌────▼─────────────────▼────────┐  │
│  │         VBA Modules           │  │
│  │  Config · Auth · DataSync     │  │
│  │  Utils · Proposta · CRM       │  │
│  └────┬──────────────────┬───────┘  │
└───────┼──────────────────┼──────────┘
        │                  │
        ▼                  ▼
  Azure SQL DB        CRM REST API
  (products,          (opportunities,
   users, log)         contacts)
```

### VBA Modules

| Module | Responsibility |
|---|---|
| `Config` | Connection string (env vars), table constants, CRM base URL |
| `Auth` | Login / logout, session state, management approval |
| `DataSync` | Bulk-load products & payment terms from Azure SQL |
| `Utils` | Clear quotation, format helpers, UI button handlers |
| `WorksheetChange` | Auto-fill minimum price on code entry; recalculate on condition change |
| `Proposta` | Word COM automation — proposal generation and audit logging |
| `CRM` | REST API calls to fetch opportunity and contact data |
| `frmLogin` | Login UserForm |
| `frmAprovacao` | Management approval UserForm |

### Database Tables

| Table | Purpose |
|---|---|
| `COT_USUARIOS` | Users, roles, and credentials |
| `COT_PRODUTOS` | Product catalog with full cost/tax structure |
| `COT_PAGAMENTO` | Payment terms and financial cost rates |
| `COT_PROPOSTAS` | Proposal audit log |
| `COT_CONTADOR` | Sequential proposal number generator |

---

## Pricing Logic

Minimum sale price is calculated as:

```
margin = (price × (1 - taxes) - cost - freight) / price

where:
  taxes  = ICMS + PIS/COFINS + IPI + commission + logistics +
            operational + sales + admin + quality + DIFAL
  cost   = unit cost + extra cost (if any)
  freight = CIF value gross-up (passed through at cost)
```

The seller can price above the minimum. Pricing below triggers an authority check — management must authenticate to approve.

---

## Setup

### 1. Azure SQL

```sql
-- Run the full schema
-- File: sql/schema.sql
```

### 2. Environment Variables (Windows)

Set these as **System Environment Variables** before opening the workbook:

```
COT_DB_SERVER    = your-server.database.windows.net
COT_DB_NAME      = your_database
COT_DB_USER      = your_db_user
COT_DB_PASSWORD  = your_db_password
COT_CRM_TOKEN    = your_crm_api_token
```

See `.env.example` for reference.

### 3. Excel

- Open `SX_COTACAO_DEMO.xlsm`
- Enable macros when prompted
- Log in with a user from `COT_USUARIOS`

---

## Project Structure

```
sx-cotacao-demo/
├── src/
│   ├── modules/
│   │   ├── Config.bas          # Connection + constants
│   │   ├── Auth.bas            # Authentication + session
│   │   ├── DataSync.bas        # Azure SQL data load
│   │   ├── Utils.bas           # Helpers + UI buttons
│   │   └── WorksheetChange.bas # Live pricing recalculation
│   └── forms/
│       ├── frmLogin.frm        # Login form
│       └── frmAprovacao.frm    # Management approval form
├── sql/
│   └── schema.sql              # Full Azure SQL schema + sample data
├── .env.example                # Environment variable reference
├── .gitignore
└── README.md
```

---

## Tech Stack

| Layer | Technology |
|---|---|
| Application | Excel VBA (Office 2016+) |
| Database | Azure SQL (via ADODB / OLEDB) |
| Document generation | Microsoft Word COM automation |
| CRM integration | REST API (MSXML2.XMLHTTP) |
| Auth | Azure SQL table + session variables |

---

## Key Design Decisions

- **No external dependencies** — runs on any Windows machine with Excel and an internet connection to Azure SQL
- **Credentials via env vars** — nothing sensitive hardcoded; safe to distribute the `.xlsm`
- **Bulk load with CopyFromRecordset** — ~10× faster than row-by-row VBA loops for the product catalog
- **Very-hidden sheets** — `DATA` sheet is `xlSheetVeryHidden`; inaccessible without VBA even if macros are disabled
- **Stateless per session** — workbook always opens clean; no stale quotation data between sessions

---

## Author

**Guilherme Araujo** — Supply Chain & Data Consultant  
Specialist in business process automation, FP&A systems, and BI development.  
SAP B1 · Power BI · Python · Azure SQL · Excel VBA

[LinkedIn](https://www.linkedin.com/in/guilhermearaujopcp) · [GitHub](https://github.com/guiharaujo)
