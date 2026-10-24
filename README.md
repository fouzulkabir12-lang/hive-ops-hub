# Hive Operations Hub

Team portal for Powered by H.I.V.E.: daily tasks, end-of-day reports, and the Comb and Waggle agent findings.

## How it is built

| Part | Where it lives | What it does |
|---|---|---|
| `index.html` | This repository, published with GitHub Pages | The page your team opens. Sign-in, tasks, end-of-day reports, team settings. |
| `apps-script/Code.gs` | Google Apps Script project "Hive Operations Hub", attached to the Google Sheet "Hive Operations Hub - Database" | The server. Checks passwords, decides who sees what, reads and writes the sheet. This file is a backup copy; the live code runs in Google. |
| Google Sheet | Fouz's Google Drive | The database: Team, Tasks, EOD, Findings, Notes, Runs tabs. Comb and Waggle read and write it every weekday morning. |

## Access levels

- **Expert** (Chris, Merab, Kamran): only their own tasks and their own end-of-day form.
- **Admin** (Tiffany): all tasks and all end-of-day reports; can assign tasks.
- **Organizer** (Fouz): everything, plus findings, notes for the agents, the Team tab (passwords and levels) and "View as".

Passwords are never stored in this repository. They are stored scrambled in the Team tab of the Google Sheet and checked by the Google script.

## Assign by voice (Gemini)

On the Assignments tab, Admin and Organizer can record a voice brief (or upload a voice note, or type one). Gemini splits it into one task per person, cleans up each task, fills in client, role and due date where they were said, and groups them by person. Nothing is saved until you check the list and press **Assign all**. Voice tasks show "(voice)" in the Source column.

Setup, in the Apps Script editor under Project Settings, Script properties:

- `GEMINI_API_KEY` (required): a key from aistudio.google.com.
- `HIVE_AGENT_INSTRUCTIONS` (optional): paste your Hive Gem's instructions here; they are added to every request.
- `GEMINI_MODEL` (optional): defaults to `gemini-2.5-flash`.

The key stays on Google's server. The page never sees it.

## Changing the server code

1. Edit the code in the Apps Script editor (open the sheet, Extensions, Apps Script).
2. Deploy, Manage deployments, edit the active deployment, Version: New version, Deploy. This keeps the same address, so the page here keeps working.
3. Copy the new code into `apps-script/Code.gs` here so the backup stays current.
