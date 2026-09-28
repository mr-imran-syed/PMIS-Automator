PMIS Automation Tool - how to use
=================================

What it does
------------
Logs into PMIS and updates data (currently Geographical Details - Block / GP /
Village) from the bundled Excel file, one row at a time. It skips anything
already done, so it is safe to run again if it stops.

First-time setup (once per computer)
------------------------------------
1. Install Python 3.10 or newer from https://www.python.org/downloads/
   IMPORTANT: on the first screen tick "Add python.exe to PATH".
2. Make sure Google Chrome, Firefox, or Edge is installed (Edge already is).

Running it
----------
1. Unzip this folder somewhere (e.g. Desktop).
2. Double-click "run.bat".
3. The app opens and runs a System Readiness check at the top.
   - If Selenium / pandas / openpyxl show "Missing", click their [Install]
     button. The app installs them into its own folder (needs internet the
     first time). When all are green, the checklist collapses on its own.
4. Choose the Environment (Training or Live). Live is highlighted yellow
   because it is the real system.
5. Type your own PMIS username and password (hold the "C" button to peek at
   the password).
6. Choose "Select the Data to be Updated" (e.g. Geographical Details).
7. Editing your data: click "Open Data File" to open the Excel, make your
   changes on the correct sheet, then SAVE and CLOSE Excel.
8. Click "Start". If Excel is still open, the app will ask you to close it.
   For Live, you'll get a confirmation prompt first.
9. Watch the animated status and the live log. Click "Stop" any time - it
   finishes the current row and saves progress. Start again later to resume.
10. "Copy Log" copies the whole current log to the clipboard.

Notes
-----
- Your password is never saved; you type it each time.
- The Excel data file lives in this folder and stores progress, so keep the
  folder together and don't run two copies at once.
- If a row can't be added, the app notes it in the log and keeps going.

Made by Imran Syed - v0.1
