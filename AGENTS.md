# AGENTS.md — Voucher Webapp JB

## Project overview

This is a Flask webapp for John Barber voucher/gift card management.

The app:
- searches voucher data from an Excel file;
- displays voucher details such as order number, card number, status, value, residual amount, service, notes and usage history;
- allows management/editing of selected voucher rows;
- can assign new gift cards;
- is deployed on Render;
- uses an Excel file as database;
- in production, the Excel file is connected through OneDrive / Microsoft Graph API.

## Important rule

Do not rewrite the whole project unless explicitly requested.

Always make the smallest safe change needed for the task.

Before modifying code, inspect the existing files and explain:
1. which files need to be changed;
2. why they need to be changed;
3. what behavior will change;
4. what should be tested after the change.

## Production safety

The live app is deployed on Render.

Do not change deployment configuration unless the task is specifically about deployment.

Do not remove or modify:
- Procfile
- requirements.txt
- Render-related configuration
- Microsoft Graph / OneDrive connection logic

unless strictly necessary and clearly explained.

## Excel / database safety

The Excel file acts as the database.

Never commit real customer data, real voucher data, tokens, secrets or production Excel files.

Do not change Excel column logic unless explicitly requested.

The app previously used these important voucher fields:
- order number
- physical card number
- customer email
- date
- status
- purchased option / box / card
- total value
- residual value
- usage columns 1 to 5
- service
- notes

When writing to Excel:
- preserve formulas;
- preserve formatting;
- preserve existing workbook structure;
- modify only the necessary cells;
- prefer openpyxl for write operations;
- avoid rewriting the entire file with pandas.

## UI / design rules

The visual style must remain minimal, dark and elegant, consistent with John Barber.

Do not redesign approved screens unless explicitly requested.

The initial search screen is considered approved and should not be changed unless the user specifically asks.

Voucher result screens are considered approved and should not be changed unless the user specifically asks.

When modifying UI:
- keep mobile responsiveness;
- keep spacing clean;
- avoid unnecessary visual complexity;
- preserve the existing layout logic.

## Local development

Local development uses .env.dev.

Typical local values:
- DEV_MODE=1
- USE_CLOUD=0
- LOCAL_EXCEL_PATH=./voucher-clienti.xlsx
- PORT=5000
- HOST=127.0.0.1

For local testing, use a local test Excel file only.

## Secrets

Never place secrets in the repository.

Do not commit:
- real Microsoft client secrets;
- access tokens;
- refresh tokens;
- production .env files;
- real customer Excel files.

Use environment variables for production secrets.

## Git workflow

Work on the current branch unless instructed otherwise.

Before changes:
- check git status;
- confirm the working tree is clean.

After changes:
- summarize modified files;
- explain the reason for each change;
- suggest a clear commit message.

## Testing expectations

After each change, test only what is relevant.

Common checks:
- app starts locally;
- home/search page loads;
- voucher search still works;
- voucher result page layout is unchanged unless intentionally modified;
- management page still updates the correct voucher row;
- Excel formulas and formatting are preserved;
- no secrets or real customer data were committed.

## Communication style

Be concise and practical.

Do not over-engineer.

If a requested change risks breaking existing working behavior, say it clearly before modifying files.

If something is unclear, ask before changing code.
