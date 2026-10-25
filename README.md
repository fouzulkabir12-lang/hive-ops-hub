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

Passwords need 8+ characters with a letter, a number and a symbol. To give everyone new passwords, run `resetAllPasswords` in the Apps Script editor: it signs everyone out and writes the new list to a "New passwords" tab in the sheet; share them, then delete that tab. Sign-in locks a name for 10 minutes after 5 wrong tries, and pauses all sign-ins after 25 wrong tries in 10 minutes. Sessions last 7 days.

Passwords are never stored in this repository. They are stored scrambled in the Team tab of the Google Sheet and checked by the Google script.

## Assign by voice (Gemini)

On the Assignments tab, Admin and Organizer can record a voice brief (or upload a voice note, or type one). Gemini splits it into one task per person, cleans up each task, fills in client, role and due date where they were said, and groups them by person. Nothing is saved until you check the list and press **Assign all**. Voice tasks show "(voice)" in the Source column.

## Hive AI (chat, everyone)

Everyone has it. Experts get the same knowledge but only their own tasks and reports, and no "Make tasks" button. A chat tab with two modes. **HIVE Master Ops** answers like the HIVE Master Ops Gem: its instructions are in the **Assistant** tab of the database sheet (row "Instructions"; any other row there is added as background), and it reads the Client Master Sheet live on every question. **Gemini** is plain Gemini with Google Search. Press 🎙 Speak to dictate (Chrome speech-to-text), or tick "Read replies aloud". When you change the Gem in Gemini, paste its new instructions into the Assistant tab.

## Gemini setup

In the Apps Script editor under Project Settings, Script properties:

- `GEMINI_API_KEY` (required): a key from aistudio.google.com.
- `GEMINI_MODEL` (optional): defaults to `gemini-3.5-flash-lite` (about 4 seconds per answer), falling back to `gemini-flash-lite-latest`; set `GEMINI_THINKING` to `high` for deeper, slower answers.
- `CLIENT_MASTER_SHEET_ID` (optional): only if the Client Master Sheet moves.

The key stays on Google's server. The page never sees it.

## Changing the server code

1. Edit the code in the Apps Script editor (open the sheet, Extensions, Apps Script).
2. Deploy, Manage deployments, edit the active deployment, Version: New version, Deploy. This keeps the same address, so the page here keeps working.
3. Copy the new code into `apps-script/Code.gs` here so the backup stays current.
