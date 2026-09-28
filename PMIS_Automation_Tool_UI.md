# Claude Code Prompt — PMIS Automation Tool UI Redesign

Redesign the existing Python-based Windows portable application UI into a clean, modern Windows utility while **preserving the existing automation/backend logic**. The current implementation already contains the technical details of what is logged, what updates are performed, and how Selenium automation works. Do not unnecessarily rewrite or change those parts.

The application name is:

**PMIS Automation Tool**

Footer:

**Made by Imran Syed · v0.1**

## 1. Overall UI direction

Replace the existing old-style UI with a **modern, clean, vertically aligned single-screen interface**.

Use:
- Clean light background
- Modern typography
- Rounded cards/containers where appropriate
- Comfortable spacing
- Subtle borders
- Modern buttons
- Minimal visual clutter
- Restrained use of color

Use colors primarily for meaning:
- Green = ready/running/success
- Yellow = warning/update/live environment
- Orange = stopped by user
- Red = error/missing dependency

The UI should remain professional, but small playful elements are acceptable through the status animations.

The application should open as a **normal medium-sized Windows window**, not full-screen. It should be resizable, with a sensible minimum size. The log area should benefit from additional space when the window is enlarged.

Everything should remain on **one vertically arranged screen**. Do not introduce tabs or separate workflow pages.

## 2. Application startup flow

When the application launches, the first thing users should see is a prominent **System Readiness** section.

This section performs an automatic checklist of everything required for the application to run.

The actual dependency list should be **dynamic/extensible**, because additional Python libraries/dependencies may be added later.

The checklist may include:
- Python
- Required Python packages
- Selenium
- Installed browser
- Appropriate browser driver/plugin such as GeckoDriver or Chromium/ChromeDriver
- Any other dependency required by the application

Do not hard-code the UI in a way that makes future dependencies difficult to add.

### Dependency states

**Ready — Green**
`Selenium                         ✅ Ready`

**Update available — Yellow**
`Selenium                         ⚠ Update Available   [Update]`

The application is still usable when a dependency is outdated but otherwise functional. An outdated dependency **must not block Start**.

**Missing/broken — Red**
`ChromeDriver                     ❌ Missing            [Install]`

A missing/broken mandatory dependency **must prevent Start**.

## 3. Dependency installation/update behavior

Do **not** open Command Prompt or Windows Terminal for normal dependency installation.

Use an **in-app installation/update experience**, for example:

> Installing Selenium...  
> ████████████░░░ 78%

Show clear progress and state changes.

Dependencies should be installed **inside the application folder/application-managed environment**, so the portable application remains associated with its own dependencies rather than relying on unrelated global Python packages.

When an installation/update performed by the application succeeds, refresh the dependency status immediately:

`❌ Missing → Installing... → ✅ Ready`

No restart should be required for successful installations performed within the app.

If the user manually changes dependencies outside the application, they can restart the application to trigger a fresh startup check.

## 4. Failed dependency installation

If an installation/update fails, keep the dependency in a red state.

Example:

`❌ Selenium — Installation Failed`

Provide:
**[Try Again] [View Error]**

`View Error` should expose the technical installation error without cluttering the normal UI.

`Try Again` should repeat only that dependency's installation/update.

## 5. Collapsing System Readiness

Once all mandatory dependencies are ready, automatically collapse the full dependency checklist.

Replace it with a compact status bar such as:

**✅ System Ready — All dependencies available**  
`View Details ▾`

If updates are available:

**⚠ System Ready — 2 updates available**  
`View Details ▾` **[Update All]**

### View Details

Clicking **View Details** expands the complete dependency checklist.

Individual outdated dependencies should have their own **Update** button.

When several dependencies are outdated, provide **Update All**.

### Update All

Update dependencies sequentially.

Example:

> Updating dependencies  
> ✅ Selenium — Updated  
> ⏳ Package 2 — Updating...  
> ○ Package 3 — Waiting  
> ████████████░░░ 67%

If one dependency fails, **continue updating the remaining dependencies**.

At the end, clearly identify which dependency/dependencies failed.

## 6. No Recheck button

Do **not** add a separate Recheck button in v0.1.

The startup dependency check runs automatically when the application launches.

If the user makes changes outside the application, restarting the app is sufficient for now.

## 7. Main workflow

After System Readiness collapses, the primary interface should be a simple vertically arranged workflow.

### Environment

Radio buttons:
- ○ Training
- ○ Live

The user must explicitly select the environment **every time the application launches**.

Do not remember the previous environment. There should be no default selection.

### Live environment styling

When **Live** is selected, highlight the Live selection in **yellow** because it is a real/production environment.

Training should have normal styling and no warning color.

## 8. Live environment confirmation

When Live is selected and the user clicks Start, **do not immediately execute the update**.

Show a confirmation dialog:

> ⚠️ **Live Environment**  
> You are about to run this update in the **LIVE environment**.
>
> **[Cancel] [Continue]**

Only Continue begins the actual automation.

When Training is selected, Start can proceed directly without this confirmation.

## 9. Credentials

Below Environment, provide:

### Username
Normal text input.

### Password
Masked password input.

Credentials must be **entered again on every application launch**.

Do not persist/save username or password between launches.

## 10. Password visibility control

Add the agreed **C control** beside the password field.

Behavior:
- Normally password is masked.
- While the C button is being held, show the password.
- As soon as the button is released, hide/mask the password again.
- This is **not a toggle**.

Add a small tooltip explaining the behavior.

## 11. Select data to be updated

Use the exact label:

### Select the Data to be Updated

Provide a dropdown.

For now, options include:
- Geographical Details
- Participant Master Data
- Other
- etc.

Display **names only**. Do not add explanatory descriptions to dropdown items.

Only **one update type can be selected and executed at a time**.

## 12. Start button and preconditions

Provide the main **Start** button.

Start should remain disabled until all required conditions are satisfied:
1. All mandatory dependencies are ready.
2. Training or Live has been selected.
3. Username has been entered.
4. Password has been entered.
5. A data-update option has been selected.

An outdated-but-functional dependency should **not** prevent Start.

## 13. Running state

After Start is clicked and execution begins:
- Change **Start** to **Stop**.
- Begin the running animation.
- Disable environment selection.
- Disable username/password editing.
- Disable data-update selection.
- Prevent accidental changes during execution.
- Begin logging in the log window.

The actual technical log content already exists in the current Claude implementation. **Do not redesign or invent the log messages.**

The log should begin only when the user actually starts an update.

## 14. Running status

### 🐱 Running

Use a green status indicator.

Use a sequence/loop of different **cat emojis/graphics** to create a small animated running effect.

The animation should remain minimal and not distract from the application.

## 15. Stop behavior

When the user clicks Stop:

### ✋ / crossed-hands — Stopped by User

Use **orange**.

The hand/crossed-hands visual should play in a small loop/animation.

Stop the active process safely.

Controls should become available again after the process has stopped.

## 16. Error state

If something fails during execution:

### ❌🚨 Error

Use **red**.

When an error occurs:
- Stop the active process.
- Change status immediately to Error.
- Keep the error details in the log.
- Unlock the controls.
- Change Stop back to Start.
- Allow the user to correct the issue and retry without restarting the application.

## 17. Successful completion

When the automation successfully finishes:

### 🎉 Completed

Use a party-popper/celebration animation.

After completion:
- Unlock environment selection.
- Unlock credentials.
- Unlock data selection.
- Change button back to Start.
- Allow another update to be run immediately.
- The user does not need to restart the application.

## 18. Idle state

The default state before a process begins is:

### IDLE

Idle should have **no strong status color** and should be visually neutral.

## 19. Status state summary

| State | Visual |
|---|---|
| Idle | Neutral |
| Running | 🐱 Green |
| Completed | 🎉 Success/celebration |
| Stopped by User | ✋ / crossed hands, Orange |
| Error | ❌🚨 Red |

## 20. Log window

Place a large log panel below the main controls.

Use a **monospaced font**.

Important behavior:
- Do not show startup dependency-check logs here.
- Logging begins only when Start is clicked.
- When a new run starts, **clear the previous log completely**.
- Each run displays only the current run's log.
- Automatically scroll to the latest log entry as new entries arrive.
- The user must still be able to scroll upward manually to inspect previous entries in the current run.

The technical details to be logged are already defined in the existing implementation.

## 21. Copy Log

The log panel must have:

**[Copy Log]**

Clicking it should copy the **entire current log** to the Windows clipboard.

No Save Log feature is required.

Add a tooltip explaining that it copies the complete current log to the clipboard.

## 22. Closing the application while running

If the user attempts to close the Windows application using the X while a process is running, do not immediately terminate it.

Show:

> ⚠️ **Process is currently running**  
> Closing the application will stop the current process.
>
> **[Keep Running] [Stop & Exit]**

### Keep Running
Cancel the close action and return to the running application.

### Stop & Exit
Safely stop the current process and then close the application.

## 23. Reset behavior between runs

A successful or stopped/failed run should return the controls to an appropriate usable state.

For a new run:
- Select the required environment.
- Enter/edit credentials as needed.
- Select the update type.
- Click Start.

The log is cleared at the beginning of the new run.

Do not require an application restart between runs.

## 24. Tooltips

Add small, unobtrusive hover tooltips only where they improve usability.

At minimum consider tooltips for:
- Password visibility C control
- Live environment
- Dependency status
- Dependency Install/Update buttons
- Copy Log
- Other controls whose purpose may not be immediately obvious

Keep tooltip text short.

## 25. Footer

At the bottom of the application, display:

**Made by Imran Syed · v0.1**

Keep it small, subtle, and centered.

## 26. Important implementation constraint

This request is primarily a **UI/UX redesign and workflow-state implementation**.

Preserve the existing:
- Python automation logic
- Selenium automation
- Existing data-update functionality
- Existing technical logging implementation
- Existing backend behavior
- Existing update-specific logic

Do not replace working backend functionality merely to achieve the UI redesign.

Where existing functionality already corresponds to a requirement above, reuse it.

The UI should act as a polished layer around the existing functionality.

## 27. UX goal

The final application should feel like a **small, modern Windows automation utility**, not an old-fashioned Python desktop form.

The user's experience should be:

**Open app → System readiness check → dependencies automatically verified → readiness collapses → choose Training/Live → enter credentials → select data → Start → watch clear animated status → monitor log → Completed / Stopped / Error → run another update.**

Keep the design simple, professional, and practical for a v0.1 release.
