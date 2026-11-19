# Office for Windows PDF export driver.
#
# Runs inside the Windows guest. Exports each listed Office file to PDF with
# one cscript + one Office instance per file: in this guest, Office COM calls
# hosted in PowerShell (ExportAsFixedFormat, SaveAs) hang with no dialog, while
# the same calls from VBScript finish in seconds. An item that exceeds its
# timeout is killed, together with only the Office processes that appeared
# during it.
#
# Usage (from a launcher in the guest):
#   $WRRoot = '\\<host>\<share>\<run-folder>'   # holds list.json and in\
#   $WRRun  = '<run name>'                           # output subfolder, created
#   $WROnly = @(3, 7)                                # optional: item numbers
#   . .\windows-export.ps1
#
# list.json: [{ "i": 1, "fmt": "docx", "app": "Word", "mode": "doc",
#               "sha256": "<of in\001.docx>", "timeout": 120 }, ...]
#   app Word | Excel | PowerPoint; Excel mode "chart" exports the single
#   chart (chart sheet or embedded), any other mode the whole workbook.
# Inputs are in\NNN.<fmt>; PDFs land in <run>\pdf\NNN.pdf, one JSON line
# per item in <run>\results.jsonl and a log in <run>\drive.log.
$ErrorActionPreference='Stop'
if(-not $WRRoot){ throw 'set $WRRoot to the share folder that holds list.json and in\' }
$WR=$WRRoot
$RunDir=Join-Path $WR $WRRun
New-Item -ItemType Directory -Path $RunDir | Out-Null
$PdfDir=Join-Path $RunDir 'pdf'; New-Item -ItemType Directory -Path $PdfDir | Out-Null
$Stage=Join-Path ([IO.Path]::GetTempPath()) ('winref-'+[Guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $Stage | Out-Null
$Results=Join-Path $RunDir 'results.jsonl'; $Log=Join-Path $RunDir 'drive.log'
function L([string]$s){ $line=[DateTime]::UtcNow.ToString('o')+' '+$s; Write-Host $line; [IO.File]::AppendAllText($Log,$line+"`r`n") }
function WrHash([string]$p){ (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() }
$pmap=@{Word='WINWORD';Excel='EXCEL';PowerPoint='POWERPNT'}
$list=[IO.File]::ReadAllText((Join-Path $WR 'list.json'),[Text.Encoding]::UTF8) | ConvertFrom-Json
if($WROnly){ $list=@($list | Where-Object { $WROnly -contains $_.i }) }
L ('START '+$WRRun+' items '+@($list).Count)
foreach($it in $list){
 $local=Join-Path $Stage ('{0:D3}.{1}' -f $it.i,$it.fmt); $pdf=Join-Path $Stage ('{0:D3}.pdf' -f $it.i); $vbs=Join-Path $Stage ('{0:D3}.vbs' -f $it.i); $outf=Join-Path $Stage ('{0:D3}.out' -f $it.i)
 $r=[ordered]@{i=$it.i; fmt=$it.fmt; app=$it.app; mode=$it.mode; driver='cscript-vbs'; startedUTC=[DateTime]::UtcNow.ToString('o'); ok=$false}
 try {
  Copy-Item -LiteralPath (Join-Path $WR ('in\{0:D3}.{1}' -f $it.i,$it.fmt)) -Destination $local -Force
  if((WrHash $local) -ne $it.sha256){ throw 'staged SHA mismatch' }
  if($it.app -eq 'Word'){
   $code=@"
Set a = CreateObject("Word.Application")
a.Visible = False
a.DisplayAlerts = 0
On Error Resume Next
a.AutomationSecurity = 3
On Error GoTo 0
Set d = a.Documents.Open("$local", False, True, False)
pg = d.ComputeStatistics(2)
d.ExportAsFixedFormat "$pdf", 17, False
d.Close 0
WScript.Echo "APPVER " & a.Version & " " & a.Build
WScript.Echo "PAGES " & pg
a.Quit 0
"@
   $call='Document.ExportAsFixedFormat(path,wdExportFormatPDF,false) via cscript'
  } elseif($it.app -eq 'Excel'){
   $code=@"
Set a = CreateObject("Excel.Application")
a.Visible = True
a.DisplayAlerts = False
a.AskToUpdateLinks = False
a.EnableEvents = False
On Error Resume Next
a.AutomationSecurity = 3
On Error GoTo 0
Set w = a.Workbooks.Open("$local", 0, True)
WScript.Echo "SHEETS " & w.Sheets.Count
mode = "$($it.mode)"
If mode = "chart" Then
 n = w.Charts.Count
 For Each ws In w.Worksheets
  n = n + ws.ChartObjects.Count
 Next
 WScript.Echo "CHARTS " & n
 If n <> 1 Then
  w.Close False
  a.Quit
  WScript.Echo "ERR chart count " & n
  WScript.Quit 2
 End If
 If w.Charts.Count = 1 Then
  w.Charts(1).ExportAsFixedFormat 0, "$pdf", 0, True, False
  WScript.Echo "LOC chartsheet"
 Else
  For Each ws In w.Worksheets
   If ws.ChartObjects.Count = 1 Then
    ws.Activate
    ws.ChartObjects(1).Activate
    a.ActiveChart.ExportAsFixedFormat 0, "$pdf", 0, True, False
    WScript.Echo "LOC embedded " & ws.Name
   End If
  Next
 End If
Else
 w.ExportAsFixedFormat 0, "$pdf", 0, True, False
End If
w.Close False
WScript.Echo "APPVER " & a.Version & " " & a.Build
a.Quit
"@
   $call=$(if($it.mode -eq 'chart'){'Chart.ExportAsFixedFormat(xlTypePDF,...) single chart via cscript'}else{'Workbook.ExportAsFixedFormat(xlTypePDF,...) whole workbook via cscript'})
  } else {
   $code=@"
Set a = CreateObject("PowerPoint.Application")
Set p = a.Presentations.Open("$local", -1, 0, 0)
WScript.Echo "SLIDES " & p.Slides.Count
p.SaveAs "$pdf", 32
p.Close
WScript.Echo "APPVER " & a.Version & " " & a.Build
If a.Presentations.Count = 0 Then a.Quit
"@
   $call='Presentation.SaveAs(path,ppSaveAsPDF) via cscript'
  }
  Set-Content -LiteralPath $vbs -Value $code -Encoding Unicode
  $before=@(Get-Process $pmap[$it.app] -ErrorAction SilentlyContinue | ForEach-Object Id)
  $t=[DateTime]::Now
  $p=Start-Process cscript.exe -ArgumentList '//nologo','//U',$vbs -PassThru -NoNewWindow -RedirectStandardOutput $outf
  $done=$p.WaitForExit([int]$it.timeout*1000)
  $r.seconds=[math]::Round(([DateTime]::Now-$t).TotalSeconds,1)
  if(-not $done){
   try { $p.Kill() } catch {}
   $new=@(Get-Process $pmap[$it.app] -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
   foreach($x in $new){ try { Stop-Process -Id $x.Id -Force; L ('STOPPED new '+$x.ProcessName+' '+$x.Id) } catch {} }
   throw 'timeout'
  }
  $out=@(); if(Test-Path $outf){ $out=@(Get-Content -LiteralPath $outf -Encoding Unicode) }
  $r.vbsOut=$out; $r.exitCode=$p.ExitCode; $r.call=$call
  if(-not (Test-Path -LiteralPath $pdf)){ throw ('no PDF produced; exit '+$p.ExitCode) }
  $dest=Join-Path $PdfDir ('{0:D3}.pdf' -f $it.i)
  Copy-Item -LiteralPath $pdf -Destination $dest -Force
  $r.pdfSHA256=WrHash $dest
  if($r.pdfSHA256 -ne (WrHash $pdf)){ throw 'PDF copy verification failed' }
  $r.ok=$true
 } catch { $r.error=$_.Exception.Message }
 $r.finishedUTC=[DateTime]::UtcNow.ToString('o')
 [IO.File]::AppendAllText($Results,($r | ConvertTo-Json -Compress -Depth 4)+"`r`n")
 L ('ITEM '+$it.i+' '+$it.app+' ok='+$r.ok+' '+$r.seconds+'s '+$r.error)
}
L 'ALLDONE'
