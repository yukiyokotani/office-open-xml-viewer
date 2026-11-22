# Office reference export

Office output is the reference for rendering decisions, under the base
policy in the [core beliefs](../design-docs/core-beliefs.md): exact match is not the goal.
This page says where reference material lives, which export to use, and how to drive Office
without leaving it in a changed state.

## Where files go

- Control and verification documents, with their Office PDFs, go in the
  private corpus folder of their format:
  - `packages/<fmt>/public/private/<fmt>/` for DOCX, XLSX and PPTX;
  - `packages/legacy-office/public/private/{doc,xls,ppt}/` for legacy binary
    files and their paired OOXML controls;
  - `packages/opendocument/public/private/{odt,ods,odp}/` for the
    LibreOffice-converted ODF corpus.
- Name them `NAME.<ext>` + `NAME.pdf`, as
  `packages/*/public/private/CORPUS-RULES.local.md` describes. Group a control
  matrix in a subfolder with a common prefix (for example `ilvl/ilvl-*`), with
  scratch notes beside it as `*.local.md`.
- Never leave controls in `/tmp`, home folders or application containers;
  stage there only to open the file.
- Run folders used for an export job (inputs, logs, launchers) are temporary.
  Copy the outputs into the repository or the private folders, then delete the
  run folder. Never cite a run folder as the record of a decision: settled
  rules go in source comments, findings in the issue or PR.

## Which export is the reference

- **Office for Windows** is the layout ground truth. Its export is saved as
  `<name>.win.pdf` (or as the `NAME.pdf` of a Windows-exported private
  control).
- **Office for Mac** exports use "Best for printing", the Mac's own Quartz
  engine, saved as `<name>.mac.pdf`. Its line layout can differ from Windows
  (for example typo line gap and `fontAlgn` under `lnSpc`); when they
  disagree, follow Windows.
- Do not use PowerPoint for Mac's "Best for electronic distribution and
  accessibility" for new references. It sends the file to a Microsoft online
  service and produces Windows-style output, which the Windows VM now covers.
  Never send private, corpus-derived or third-party files to it. Older rules
  measured from its tagged output (for example #1610) reflect Windows layout.
- Public corpus specimens ship a `<stem>.win.pdf` and a `<stem>.mac.pdf` per
  DOCX/PPTX specimen; for XLSX only chart specimens get PDFs. Re-export both
  whenever the specimen changes, with the corpus fonts installed in both
  environments.
- Private references re-exported on Windows keep the Mac originals in a
  `reference-mac/` subfolder with a `reference-windows-manifest.json`.
- Read the exported PDF yourself (for example `pdftoppm`, `pdffonts`); never
  rely on someone's written description of colors or positions.

## Office for Mac (GUI)

- Open files read-only with macros disabled; do not update links.
- Decline any "repair this file?" prompt (No/Cancel), close only that
  document, record which file triggered it (a repair prompt is a bug in a
  self-authored control), and continue with the rest. When an open or save
  seems to hang, look for such a modal first.
- Stage the file under the app's container tmp
  (`~/Library/Containers/com.microsoft.<App>/Data/tmp/`) and open it with
  `open -a "Microsoft <App>" <path>`; Office cannot open arbitrary paths
  without permission prompts.
- Never change Office's global settings (automation security, display alerts)
  through AppleScript; a killed script leaves them changed. Drive the GUI
  through computer use instead.
- Keep an unsaved blank document open in each app during a GUI phase, so
  closing the last control does not leave the app windowless.
- Computer use cannot attach while the Mac screen is locked
  (`ioreg -n Root -d1 | grep CGSSessionScreenIsLocked`); ask the user to
  unlock it.
- Close only your own documents and leave no dialog open. Serialize GUI
  phases across sessions with `mkdir /private/tmp/office-gui.lock`, released
  by `rmdir` in an EXIT trap.

## Office for Windows in the VMware VM

- Keep VMware Fusion **windowed**. Switch to full screen only for a step that
  cannot be done in the window, and switch back right after it; if it is full
  screen for no such reason, exit full screen first. Put this in every brief
  that drives the VM.
- Serialize VM use across sessions with `/private/tmp/ooxml-vm.lock`.
- Input: the `type` action and the Mac clipboard do not reach the guest. Press
  keys one at a time; `.`, `/`, `+` and `:` may not arrive, so type a launcher
  that rebuilds them, for example
  `iex(irm ('httpjzzHOSTj8765zgoqtxt' -replace 'q',[char]46 -replace 'z',[char]47 -replace 'j',[char]58))`
  with placeholder letters that occur nowhere else in the URL. Afterwards, Up
  and Return rerun it; zoom-check the recalled line first, because history may
  hold other sessions' commands.
- Transfer files through the host share, or, when the guest's saved share
  credential fails, through a small HTTP server bound only to the VM network's
  host address (GET for inputs, PUT for results). Serve launcher text with a
  `.txt` name. Never replace a share file while another process reads it.
- Check that the guest has the document's fonts; a missing font changes the
  layout and invalidates the reference.
- For host-font rendering of this viewer on Windows, open a Vite-served batch
  page in the guest's Edge that uploads its PNGs back.

## Driving Office through COM

- Drive Office COM from VBScript (`cscript //nologo //U`), one script and one
  Office instance per file, never from PowerShell: in the guest used so far,
  Word `ExportAsFixedFormat`/`SaveAs2` and Excel embedded-chart export called
  from PowerShell hang indefinitely with no dialog, while the same calls from
  VBScript finish in seconds.
- Wrap the per-file VBScript in a driver that verifies macros are disabled
  before opening a file, applies a timeout per item, ends only the Office
  process it started, treats a non-zero exit as a failure, and removes its
  staging copies when the run ends.
- Word: `$wa.Quit()` (not `Quit(0)`); iterate merged tables through
  `Range.Cells`; screenshots need `SetProcessDPIAware` and an explicit
  `Zoom.Percentage` (PageFit is ignored). Layout probes can read
  `Information(6)` and shape positions without a PDF export.
- PowerShell: variables are case-insensitive, and aliases beat functions (a
  helper named `H` is shadowed by `h`, Get-History), so avoid one- or
  two-letter function names.
