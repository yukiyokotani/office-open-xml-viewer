# Office PDF export

Tools for exporting Office reference PDFs.
[`docs/workflow/office-reference-export.md`](../../docs/workflow/office-reference-export.md)
says which export is the reference and how to drive Office.

- `windows-export.ps1`: Office for Windows driver that runs in the Windows
  guest. It runs one VBScript and one Office instance per file, with a
  timeout for each item. The header comment explains how to call it and the
  `list.json` format.

  Office COM must be driven from VBScript, not from PowerShell. In the guest
  used so far, `ExportAsFixedFormat` and `SaveAs` hung indefinitely when
  called from PowerShell.

Run folders on the host share are temporary. Copy the PDFs into the
repository (or the ignored private folders), then delete the run folder.
