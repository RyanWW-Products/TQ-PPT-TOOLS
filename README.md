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

## Release workflow (each new version)

The version is **not** stored in the VBA. It lives in `version.json` (the latest
available) and in each machine's registry `InstalledVersion` (written by the
installer and by the updater). So a release is just:

1. Edit macros in the `.pptm`. (Only re-import `Updater.bas` if you changed *it*.)
2. **Save As** a macro-enabled add-in `.ppam` (e.g. into the repo root).
3. Build the distributable (injects the ribbon, stamps `version.json`):
   ```powershell
   .\build\build.ps1 -InputPpam ".\TrialQuest Addin Master v5-4-7.ppam" -Version 5.4.7
   ```
4. If you changed any **template assets**, bump `assetsVersion` in `version.json`,
   then regenerate the manifest:
   ```powershell
   .\build\make-manifest.ps1
   ```
5. Commit and push `dist/TrialQuest.ppam`, `version.json`, `manifest.json`, and any
   changed `assets/…`. Clients now see the update via **Check for Updates**.
6. (Only when the installer itself changes) bump `#define MyAppVersion` in
   `installer.iss` and recompile it:
   ```powershell
   & "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" .\build\installer.iss
   ```
   Output lands in `build\Output\`.

> No VBA constant to edit and no re-import for a version bump — `build.ps1 -Version`
> stamps `version.json`, the installer writes the registry version on first install,
> and the updater writes it after each self-update.

---

## Beta / stable channels

Two release channels live in the **same repo, on two branches**:

| Channel | Branch | Who | Registry `Channel` |
|---------|--------|-----|--------------------|
| stable  | `main` | live users      | `stable` (or unset) |
| beta    | `beta` | you / testers   | `beta` |

The add-in reads `HKCU\Software\TrialQuest\Addin\Channel` and pulls `version.json`
+ `dist/TrialQuest.ppam` from the matching branch. One build serves both channels —
the channel is just a registry flag, so live users never see beta because their
machine reads `main`.

### Develop a beta release

```powershell
git checkout beta
# edit macros, build.ps1, bump version.json on beta, commit
git push origin beta
```

Anyone on the beta channel gets it via **Check for Updates**.

### Promote beta → stable (ship to everyone)

```powershell
git checkout main
git merge beta        # brings beta's code + version.json + dist into main
git push origin main
```

### Put a machine on a channel

* *Beta installer:* `ISCC.exe /DChannel=beta build\installer.iss` → `…Setup vX.Y.Z beta.exe` (sets `Channel=beta`).
* *Flip an existing install:* `.\build\set-channel.ps1 beta` (or `stable`), then Check for Updates.

> Channel-switch caveat: if beta is *ahead* of stable, moving a machine from beta
> back to stable won't "downgrade" it (the stable `version.json` is lower than what's
> installed), so it simply stops receiving new updates until stable catches up.

---

## Installing on a client

1. Run `TEI Addin Setup vX.Y.Z.exe`.
2. Paste the read-only token when prompted.
3. The installer copies `TrialQuest.ppam` to
   `%APPDATA%\Microsoft\AddIns\Trial Ex Addin\`, registers it for auto-load, and
   downloads the template assets. Start PowerPoint — the **Trial Quest** tab appears.

If asset download fails (network/token), the add-in still installs; the user can
fetch assets later via **Trial Quest → Check for Updates**.

## Updating on a client

**Trial Quest → Check for Updates.** It compares the installed version to
`version.json`, downloads what changed, applies template assets immediately, and —
if the add-in itself changed — asks the user to close PowerPoint so the new `.ppam`
can be swapped in (PowerPoint reopens automatically). **About** shows the installed
version, assets version and token status.

---

## Initial push

```powershell
cd "c:\Users\tqrwhitehead\Documents\VisualStudioCode\PROJECTS\TQ PPT MASTER"
git init
git add .
git commit -m "Initial GitHub-distributable add-in"
git branch -M main
git remote add origin https://github.com/<owner>/<repo>.git   # PRIVATE repo
git push -u origin main
```

`.gitignore` excludes the local `_AppdataBackup/`, installer binaries, large design
sources, and the transient root `.ppam` build inputs.

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
