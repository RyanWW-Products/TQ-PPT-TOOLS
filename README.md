# Trial Quest PowerPoint Add-in

A VBA PowerPoint add-in (the **Trial Quest** ribbon) for building trial / clinical
presentation exhibits: timelines, calendars, vitals scatterplots, a quick-graphics
library, bulk tools and paragraph styles.

This repo makes the add-in **install once, then update from a button press**. It is
a **private** repository — the add-in and installer authenticate to GitHub with a
read-only token to pull updates.

---

## Repo layout

| Path | Purpose |
|------|---------|
| `Modules/*.bas`, `Forms/*` | VBA source (git-tracked exports of the live project) |
| `Modules/Updater.bas` | Update button + version check + authed download + swapper |
| `ribbon/customUI.xml`, `ribbon/images/` | Ribbon definition + button icons (the build injects these) |
| `assets/Trial Ex Addin/` | The ~36 MB runtime templates clients download |
| `build/build.ps1` | Injects the ribbon into a macro-only `.ppam` → `dist/TrialQuest.ppam` |
| `build/make-manifest.ps1` | Regenerates `manifest.json` from `assets/` |
| `build/download-assets.ps1` | Asset downloader used by the installer (and standalone) |
| `build/installer.iss` | Inno Setup script for the slim installer |
| `dist/TrialQuest.ppam` | The built, ribboned add-in clients download |
| `version.json`, `manifest.json` | Update metadata (repo root) |

The macro **source of truth is the `.pptm`**. The `.bas`/`.frm` files are git-tracked
exports (via `A_PROJECT_EXPORTER.ExportAllFormsAndModules`); you re-import them into
the VBA editor for them to take effect. `build.ps1` injects only the ribbon — it does
not compile VBA — so always Save-As your `.ppam` from the current `.pptm` first.

---

## One-time setup

1. **Create the private GitHub repo** and push this folder (see "Initial push").
2. **Mint a fine-grained PAT** (GitHub → Settings → Developer settings → Fine-grained
   tokens):
   * Resource owner: your org/user · Repository access: **only this repo**.
   * Permissions: **Contents → Read-only**. Nothing else.
   * Copy the token (`github_pat_…`). This is the token the installer asks for.
3. **Set the three constants in two places to match the repo:**
   * `Modules/Updater.bas` → `GH_OWNER`, `GH_REPO`, `GH_BRANCH`
   * `build/installer.iss` → `GhOwner`, `GhRepo`, `GhBranch`
   Re-import `Updater.bas` into the `.pptm` after editing.

> **Security:** the token is read-only and scoped to this one repo, and is stored per
> machine at `HKCU\Software\TrialQuest\Addin\GitHubToken` — never in git or in the
> `.ppam`. Rotate it by issuing a new PAT and re-running the installer (or updating
> that registry value). A leaked token can only *read* this one repo.

---


## Troubleshooting

| Symptom | Likely cause / fix |
|--------|--------------------|
| Ribbon tab missing after install | Add-in not registered for your Office version. Check `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\TrialQuest` (Path + AutoLoad). |
| "No update token configured" | Re-run installer, or use **About → set token**. Stored at `HKCU\Software\TrialQuest\Addin\GitHubToken`. |
| Update check fails with HTTP 404 | `GH_OWNER/GH_REPO/GH_BRANCH` mismatch, or token lacks Contents:Read on this repo. |
| Update downloaded but not applied | The swapper waits for **all** PowerPoint windows to close. Close them; it installs and relaunches. |
| Quick Graphics / Vitals can't find templates | Assets not downloaded. Run **Check for Updates** (it re-syncs missing assets). |
| `build.ps1` fails "not a valid .ppam" | You pointed it at a non-add-in file. Save As **PowerPoint Add-in (*.ppam)** from the `.pptm`. |
