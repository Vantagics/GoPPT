# export_office.ps1 - PowerPoint's own rendering of a deck, as the reference
# that zz_cmp compares a Go render against.
#
#   powershell -File export_office.ps1 -Deck <file.pptx> -OutDir <dir> [-Width 1600]
#
# It drives PowerPoint over COM, so it only runs where PowerPoint is installed.
# The slide size is read from the presentation rather than assumed, and every
# page is exported at the same pixel width the Go renderer was given - two
# renders at different sizes cannot be compared at all.
#
# Notes that cost time to learn:
#   * PageSetup.SlideWidth/SlideHeight are in points; width * (h/w) keeps the
#     aspect ratio the deck actually declares (this one is 720x540, i.e. 4:3).
#   * Open()'s fourth argument ($false) keeps the window hidden. With no window
#     the application still appears in the process list and is killed below.
#   * Export writes Slide1.PNG, Slide2.PNG: no zero padding, different case and
#     no "slide" prefix, so the names are normalised to slide01.png to pair up
#     with the renderer's output.
#   * DisplayAlerts = 1 is ppAlertsNone; without it a repair prompt can block a
#     non-interactive run forever.

param(
    [Parameter(Mandatory = $true)][string]$Deck,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$Width = 1600,
    [string]$Log = ""
)

$ErrorActionPreference = "Stop"
if ($Log -eq "") { $Log = Join-Path $OutDir "office_export.log" }
$lines = New-Object System.Collections.ArrayList
function L($m) { [void]$lines.Add([string]$m) }

$app = $null
try {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

    $app = New-Object -ComObject PowerPoint.Application
    $app.DisplayAlerts = 1

    # ReadOnly = true, Untitled = false, WithWindow = false.
    $pres = $app.Presentations.Open((Resolve-Path $Deck).Path, $true, $false, $false)

    $sw = [double]$pres.PageSetup.SlideWidth
    $sh = [double]$pres.PageSetup.SlideHeight
    $h = [int][Math]::Round($Width * $sh / $sw)

    L ("deck      = " + (Resolve-Path $Deck).Path)
    L ("slides    = " + $pres.Slides.Count)
    L ("pagesetup = " + $sw + " x " + $sh + " pt")
    L ("export    = " + $Width + " x " + $h + " px")

    $pres.Export($OutDir, "PNG", $Width, $h)
    $pres.Close()
    L "export-ok"
}
catch {
    L ("ERROR: " + $_.Exception.Message)
    L ("TRACE: " + $_.ScriptStackTrace)
}
finally {
    if ($app -ne $null) { try { $app.Quit() } catch { L ("quit-err: " + $_.Exception.Message) } }
    Start-Sleep -Seconds 1
    Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
        L ("killing left-over pid=" + $_.Id)
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
}

# Normalise PowerPoint's names to the renderer's, so a name pairs the two runs.
Get-ChildItem -Path $OutDir -Filter *.PNG | ForEach-Object {
    if ($_.Name -match '(\d+)') {
        $n = "slide" + ([int]$Matches[1]).ToString("00") + ".png"
        Move-Item -Path $_.FullName -Destination (Join-Path $OutDir $n) -Force
    }
}
$after = @(Get-ChildItem -Path $OutDir -Filter *.png | Sort-Object Name)
L ("pages on disk = " + $after.Count)

$enc = New-Object System.Text.UTF8Encoding -ArgumentList $false
[System.IO.File]::WriteAllLines($Log, $lines, $enc)
