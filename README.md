# Trial Quest PowerPoint Add-in

A VBA PowerPoint add-in (the **Trial Quest** ribbon) for building trial / clinical
presentation exhibits: timelines, calendars, vitals scatterplots, a quick-graphics
library, bulk tools and paragraph styles.

This repo makes the add-in **install once, then update from a button press**. It is
a **private** repository — the add-in and installer authenticate to GitHub with a
read-only token to pull updates.

---

## EZ Highlights

The **EZ Highlights** button sits between Auto Callout and Shape Tools.
With nothing selected, click it and drag a rectangle on the slide. With one or
more shapes selected, click it to convert them. Press Escape to cancel drawing.
To highlight shapes inside a group, select the group and then click the desired
child shape (Shift-click to select additional children), then click EZ Highlights.
The containing group, nested layout and unselected neighbors are retained.
For a grouped shape containing text, matching ink/text animations run together.

Highlights use embedded Windows Ink with `#FFFF00`, zero transparency and the
native `MaskPen` highlighter operation. Black text remains black beneath the
highlight. Each result is a normal PowerPoint group containing native ink and
an editable copy of the original shape with its fill and outline hidden. Shape
geometry, text, tags, stacking position and animation effects are retained;
animation click triggers are reassigned to the result.

PowerPoint's Fade effect can flash black on native ink, and its animation gallery
hides Replay/Rewind for ink inside groups. EZ Highlights replaces ink entrance
Fades with native Replay and exit Fades with Rewind. Timing, order and shape-click
triggers are retained; editable text keeps its ordinary Fade. This happens when
converting a shape that already has Fade. If you add Fade afterward, select the
existing highlight and click **EZ Highlights** again to repair it. Repeating this
does not add duplicate effects. Unanimated highlights stay unanimated.

The macro-free Replay/Rewind template in `build/templates/InkReplayTemplate.pptx`
is embedded in VBA by `build/embed-ink-replay-template.ps1`; it is not a runtime
download. PowerPoint's VBA enum cannot create these native effects directly.

Conversion supports AutoShapes, freeforms and text boxes, including selected
children inside groups. Whole groups, pictures, charts, lines and shapes with 3-D effects are left unchanged with an
explanation. Changing the editable child's geometry afterward does not redraw
the ink; resize or rotate the complete group to keep them aligned.

The saved `.pptx` displays its highlights in Windows desktop PowerPoint without
this add-in or any linked assets. Native Ink has PowerPoint limitations: standard
PDF/image export and printing can look faded, other animated transitions can show
artifacts, and Mac/web rendering is not equivalent. See
[Microsoft's ink pen documentation](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/mpc/pens-and-pen-options)
and [DoneBy5's native-ink compatibility FAQ](https://doneby5.com/HP.html).

Run `build/test-ez-highlights.ps1` for isolated conversion and save/reopen checks.
Add `-MouseTest` to also draw with the mouse and test Escape cancellation in a
disposable PowerPoint window. This optional test temporarily controls the mouse.
Reports and preview files remain under the ignored `build/Output` directory.
Use `-SourcePptm '<prepared master>.pptm'` to test the assembled release project
in a disposable copy, rather than importing the loose feature modules.

The updater preserves the saved PAT and other settings when recording the new
version. `build/test-updater-settings.ps1` reproduces the old registry-key reset
and verifies preservation using disposable keys and a dummy token. Re-running
the installer also keeps the saved token when its token field is left blank.

---

## Repo layout

| Path | Purpose |
|------|---------|
| `Modules/*.bas`, `Modules/*.cls`, `Forms/*` | VBA source (git-tracked exports of the live project) |
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
