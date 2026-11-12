# Playwright-Werkzeug – Hauptprogramm (Windows PowerShell 5.1)
# Start über START.bat oder TEAMS-MITSCHNITT.bat, nicht direkt.
param([switch]$Teams)

$Basis = $PSScriptRoot
Set-Location $Basis
$Host.UI.RawUI.WindowTitle = 'Playwright-Werkzeug'
$DE = [Globalization.CultureInfo]::GetCultureInfo('de-DE')

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
if (-not ('PwWerkzeug' -as [type])) {
    Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
public static class PwWerkzeug {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    public static Bitmap Aufnehmen(int x, int y, int w, int h) {
        var b = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(b)) { g.CopyFromScreen(x, y, 0, 0, new Size(w, h)); }
        return b;
    }
    // Verkleinertes Graustufenbild (80 x 45) zum Vergleichen
    public static byte[] Fingerabdruck(Bitmap src) {
        const int W = 80, H = 45;
        using (var k = new Bitmap(W, H, PixelFormat.Format32bppArgb)) {
            using (var g = Graphics.FromImage(k)) {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, W, H);
            }
            var bd = k.LockBits(new Rectangle(0, 0, W, H), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var roh = new byte[bd.Stride * H];
            Marshal.Copy(bd.Scan0, roh, 0, roh.Length);
            int stride = bd.Stride;
            k.UnlockBits(bd);
            var f = new byte[W * H];
            for (int y = 0; y < H; y++)
                for (int x = 0; x < W; x++) {
                    int i = y * stride + x * 4;
                    f[y * W + x] = (byte)((roh[i] * 29 + roh[i + 1] * 150 + roh[i + 2] * 77) >> 8);
                }
            return f;
        }
    }
    // Anteil der Bildpunkte (0..1), die sich deutlich unterscheiden
    public static double Unterschied(byte[] a, byte[] b, int schwelle) {
        if (a == null || b == null || a.Length != b.Length) return 1.0;
        int n = 0;
        for (int i = 0; i < a.Length; i++) if (Math.Abs(a[i] - b[i]) > schwelle) n++;
        return (double)n / a.Length;
    }
    // Fast einfarbig (schwarz, weiß, Ladebildschirm)?
    public static bool Leer(byte[] f) {
        int min = 255, max = 0;
        foreach (var v in f) { if (v < min) min = v; if (v > max) max = v; }
        return (max - min) < 12;
    }
}
'@
}
[void][PwWerkzeug]::SetProcessDPIAware()

# ---------------------------------------------------------------- Einstellungen
$EinstDatei = Join-Path $Basis 'einstellungen.json'
$Standard = [ordered]@{
    Browser                 = 'msedge'
    PlaywrightOrdner        = ''
    TeamsIntervallSekunden  = 2
    TeamsSchwelleProzent    = 3
    TeamsBereich            = ''
    OneNoteAbschnittId      = ''
    OneNoteAbschnittName    = ''
    OneNoteBildbreite       = 720
}

function Lade-Einstellungen {
    $e = [ordered]@{}
    foreach ($k in $Standard.Keys) { $e[$k] = $Standard[$k] }
    if (Test-Path $EinstDatei) {
        try {
            $j = Get-Content $EinstDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $j.PSObject.Properties) { $e[$p.Name] = $p.Value }
        } catch { Write-Host "einstellungen.json ist fehlerhaft, nehme Standardwerte." -ForegroundColor Yellow }
    }
    return $e
}
function Speichere-Einstellungen($e) {
    ($e | ConvertTo-Json -Depth 4) | Set-Content $EinstDatei -Encoding UTF8
}

function Sauberer-Name([string]$t) {
    $t = ($t -replace '[\\/:*?"<>|]', ' ' -replace '\s+', ' ').Trim()
    if ($t.Length -gt 60) { $t = $t.Substring(0, 60).Trim() }
    if (-not $t) { $t = 'ohne Namen' }
    return $t
}
function Pause-Taste { Write-Host ''; [void](Read-Host 'Weiter mit Enter') }

# ================================================================ TEIL 1: Teams → OneNote

function Waehle-Bereich {
    $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $form = New-Object System.Windows.Forms.Form
    $form.FormBorderStyle = 'None'
    $form.StartPosition = 'Manual'
    $form.Bounds = $vs
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.BackColor = [System.Drawing.Color]::Black
    $form.Opacity = 0.4
    $form.Cursor = [System.Windows.Forms.Cursors]::Cross
    $form.KeyPreview = $true
    $form.GetType().GetProperty('DoubleBuffered', [Reflection.BindingFlags]'NonPublic,Instance').SetValue($form, $true, $null)

    $script:bStart = $null
    $script:bRect = [System.Drawing.Rectangle]::Empty
    $script:bFertig = $null

    $machRect = {
        param($a, $b)
        New-Object System.Drawing.Rectangle ([Math]::Min($a.X, $b.X)), ([Math]::Min($a.Y, $b.Y)), ([Math]::Abs($a.X - $b.X)), ([Math]::Abs($a.Y - $b.Y))
    }
    $form.Add_MouseDown({ $script:bStart = $_.Location })
    $form.Add_MouseMove({
        if ($script:bStart) { $script:bRect = & $machRect $script:bStart $_.Location; $form.Invalidate() }
    })
    $form.Add_MouseUp({
        if ($script:bStart) {
            $script:bFertig = & $machRect $script:bStart $_.Location
            $form.Close()
        }
    })
    $form.Add_KeyDown({ if ($_.KeyCode -eq 'Escape') { $script:bFertig = $null; $form.Close() } })
    $form.Add_Paint({
        $g = $_.Graphics
        $schrift = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
        $text = 'Rahmen um die Präsentation ziehen  –  Esc bricht ab'
        $pri = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $g.DrawString($text, $schrift, [System.Drawing.Brushes]::White, ($pri.X - $vs.X + 40), ($pri.Y - $vs.Y + 40))
        if (-not $script:bRect.IsEmpty) {
            $stift = New-Object System.Drawing.Pen([System.Drawing.Color]::Red, 4)
            $g.DrawRectangle($stift, $script:bRect)
            $stift.Dispose()
        }
        $schrift.Dispose()
    })
    [void]$form.ShowDialog()
    $form.Dispose()

    $r = $script:bFertig
    if (-not $r -or $r.Width -lt 80 -or $r.Height -lt 60) { return $null }
    # Formular-Koordinaten → Bildschirm-Koordinaten
    return New-Object System.Drawing.Rectangle ($r.X + $vs.X), ($r.Y + $vs.Y), $r.Width, $r.Height
}

function Teams-Titel {
    $t = Get-Process -Name 'ms-teams', 'Teams', 'msteams' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle } | Select-Object -ExpandProperty MainWindowTitle -First 1
    if ($t) { $t = ($t -replace '\s*\|\s*Microsoft Teams.*$', '').Trim() }
    if (-not $t -or $t -match '^(Microsoft Teams|Teams|Chat|Aktivität|Kalender)$') { $t = 'Teams-Termin' }
    return $t
}

function Position-Steuerfenster($bereich, $groesse) {
    $bildschirme = @([System.Windows.Forms.Screen]::FromRectangle($bereich)) + [System.Windows.Forms.Screen]::AllScreens
    foreach ($s in $bildschirme) {
        $w = $s.WorkingArea
        $ecken = @(
            (New-Object System.Drawing.Point ($w.Right - $groesse.Width - 10), ($w.Bottom - $groesse.Height - 10)),
            (New-Object System.Drawing.Point ($w.Left + 10), ($w.Bottom - $groesse.Height - 10)),
            (New-Object System.Drawing.Point ($w.Right - $groesse.Width - 10), ($w.Top + 10)),
            (New-Object System.Drawing.Point ($w.Left + 10), ($w.Top + 10))
        )
        foreach ($p in $ecken) {
            $kandidat = New-Object System.Drawing.Rectangle $p, $groesse
            if (-not $kandidat.IntersectsWith($bereich)) { return $p }
        }
    }
    $w = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    return New-Object System.Drawing.Point ($w.Right - $groesse.Width - 10), ($w.Bottom - $groesse.Height - 10)
}

function Mitschnitt-Schritt([bool]$erzwingen) {
    $M = $script:M
    if ($M.Pause -and -not $erzwingen) { return }
    $bmp = $null
    try {
        $r = $M.Bereich
        $bmp = [PwWerkzeug]::Aufnehmen($r.X, $r.Y, $r.Width, $r.Height)
        $fp = [PwWerkzeug]::Fingerabdruck($bmp)
        $speichern = $false
        if ($erzwingen) {
            $speichern = $true
        } elseif (-not [PwWerkzeug]::Leer($fp)) {
            $stabil = ($null -ne $M.Vorher) -and ([PwWerkzeug]::Unterschied($fp, $M.Vorher, 20) -lt 0.01)
            $neu = ($null -eq $M.Letzter) -or ([PwWerkzeug]::Unterschied($fp, $M.Letzter, 20) -ge $M.Schwelle)
            $speichern = $stabil -and $neu
        }
        $M.Vorher = $fp
        if ($speichern) {
            $M.Anzahl++
            $zeit = Get-Date
            $datei = Join-Path $M.Ordner ('Folie_{0:000}_{1:HH-mm-ss}.png' -f $M.Anzahl, $zeit)
            $bmp.Save($datei, [System.Drawing.Imaging.ImageFormat]::Png)
            $M.Letzter = $fp
            $M.Status.Text = "{0} Folie(n) gespeichert`nletzte: {1:HH:mm:ss}" -f $M.Anzahl, $zeit
        }
    } catch {
        $M.Status.Text = "Fehler: $($_.Exception.Message)"
    } finally {
        if ($bmp) { $bmp.Dispose() }
    }
}

function Starte-TeamsMitschnitt {
    $e = Lade-Einstellungen
    Clear-Host
    Write-Host '=== Teams-Präsentation mitschneiden ===' -ForegroundColor Cyan
    Write-Host 'Das Werkzeug macht ein Bild, sobald sich die geteilte Präsentation ändert.'
    Write-Host 'Das Teams-Fenster muss sichtbar bleiben (nicht minimieren, nicht verdecken).'
    Write-Host ''
    $vorschlag = Teams-Titel
    $titel = Read-Host "Name des Termins [Enter = $vorschlag]"
    if (-not $titel) { $titel = $vorschlag }

    $bereich = $null
    if ($e.TeamsBereich -match '^\s*(-?\d+),\s*(-?\d+),\s*(\d+),\s*(\d+)\s*$') {
        $alt = New-Object System.Drawing.Rectangle ([int]$Matches[1]), ([int]$Matches[2]), ([int]$Matches[3]), ([int]$Matches[4])
        $a = Read-Host "Bereich vom letzten Mal nehmen ($($alt.Width) x $($alt.Height) Pixel)? [Enter = ja, N = neu aufziehen]"
        if ($a -notmatch '^[nN]') { $bereich = $alt }
    }
    if (-not $bereich) {
        Write-Host ''
        Write-Host 'Gleich wird der Bildschirm abgedunkelt. Ziehe mit der Maus einen Rahmen NUR um die'
        Write-Host 'geteilte Präsentation – ohne Videokacheln, Chat und Teams-Leisten.'
        [void](Read-Host 'Enter drücken, sobald die Präsentation zu sehen ist')
        $bereich = Waehle-Bereich
        if (-not $bereich) { Write-Host 'Abgebrochen (kein Rahmen gezogen).' -ForegroundColor Yellow; Pause-Taste; return }
        $e.TeamsBereich = '{0},{1},{2},{3}' -f $bereich.X, $bereich.Y, $bereich.Width, $bereich.Height
        Speichere-Einstellungen $e
    }

    $beginn = Get-Date
    $ordner = Join-Path $Basis ('Mitschnitte\{0:yyyy-MM-dd_HHmm}_{1}' -f $beginn, (Sauberer-Name $titel))
    New-Item -ItemType Directory -Force -Path $ordner | Out-Null
    @{ Titel = $titel; Beginn = $beginn.ToString('o') } | ConvertTo-Json | Set-Content (Join-Path $ordner 'termin.json') -Encoding UTF8

    # --- kleines Steuerfenster
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Teams-Mitschnitt'
    $form.FormBorderStyle = 'FixedToolWindow'
    $form.TopMost = $true
    $form.StartPosition = 'Manual'
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Padding = New-Object System.Windows.Forms.Padding 8
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.FlowDirection = 'TopDown'
    $panel.AutoSize = $true
    $status = New-Object System.Windows.Forms.Label
    $status.AutoSize = $true
    $status.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $status.Text = "Läuft – warte auf erste Folie …`n "
    $knoepfe = New-Object System.Windows.Forms.FlowLayoutPanel
    $knoepfe.AutoSize = $true
    $bJetzt = New-Object System.Windows.Forms.Button; $bJetzt.Text = 'Bild jetzt'; $bJetzt.AutoSize = $true
    $bPause = New-Object System.Windows.Forms.Button; $bPause.Text = 'Pause'; $bPause.AutoSize = $true
    $bEnde = New-Object System.Windows.Forms.Button; $bEnde.Text = 'Beenden'; $bEnde.AutoSize = $true
    $knoepfe.Controls.AddRange(@($bJetzt, $bPause, $bEnde))
    $panel.Controls.AddRange(@($status, $knoepfe))
    $form.Controls.Add($panel)

    $script:M = @{
        Bereich = $bereich; Ordner = $ordner; Vorher = $null; Letzter = $null; Anzahl = 0; Pause = $false
        Schwelle = [double]$e.TeamsSchwelleProzent / 100; Status = $status
    }
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [int]([double]$e.TeamsIntervallSekunden * 1000)
    $timer.Add_Tick({ Mitschnitt-Schritt $false })
    $bJetzt.Add_Click({ Mitschnitt-Schritt $true })
    $bPause.Add_Click({
        $script:M.Pause = -not $script:M.Pause
        if ($script:M.Pause) { $bPause.Text = 'Weiter'; $status.Text = "PAUSE – $($script:M.Anzahl) Folie(n)`n " } else { $bPause.Text = 'Pause' }
    })
    $bEnde.Add_Click({ $form.Close() })
    $form.Add_Shown({
        $form.Location = Position-Steuerfenster $bereich $form.Size
        $timer.Start()
    })
    Write-Host ''
    Write-Host "Mitschnitt läuft. Steuerfenster: 'Bild jetzt' / 'Pause' / 'Beenden'." -ForegroundColor Green
    [void]$form.ShowDialog()
    $timer.Stop(); $timer.Dispose(); $form.Dispose()

    $anzahl = $script:M.Anzahl
    Write-Host ''
    Write-Host "$anzahl Folie(n) gespeichert in: $ordner"
    if ($anzahl -eq 0) { Pause-Taste; return }
    Schreibe-HtmlNotiz $ordner | Out-Null
    $a = Read-Host 'Jetzt als Notiz in OneNote anlegen? [Enter = ja, N = nein]'
    if ($a -notmatch '^[nN]') { [void](Nach-OneNote $ordner) }
    Pause-Taste
}

function Lies-Mitschnitt([string]$ordner) {
    $info = @{ Titel = (Split-Path $ordner -Leaf); Beginn = (Get-Item $ordner).CreationTime }
    $tj = Join-Path $ordner 'termin.json'
    if (Test-Path $tj) {
        $j = Get-Content $tj -Raw -Encoding UTF8 | ConvertFrom-Json
        $info.Titel = $j.Titel
        if ($j.Beginn -is [datetime]) { $info.Beginn = $j.Beginn.ToLocalTime() }
        else { $info.Beginn = [datetime]::Parse($j.Beginn, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToLocalTime() }
    }
    $info.Bilder = @(Get-ChildItem $ordner -Filter 'Folie_*.png' | Sort-Object Name)
    if ($info.Bilder.Count -gt 0) { $info.Ende = $info.Bilder[-1].LastWriteTime } else { $info.Ende = $info.Beginn }
    return $info
}

function Zeit-Aus-Datei($datei) {
    if ($datei.Name -match '_(\d\d)-(\d\d)-(\d\d)\.png$') { return "$($Matches[1]):$($Matches[2]):$($Matches[3])" }
    return $datei.LastWriteTime.ToString('HH:mm:ss')
}

function Schreibe-HtmlNotiz([string]$ordner) {
    $i = Lies-Mitschnitt $ordner
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!doctype html><html lang="de"><head><meta charset="utf-8"><title>' + (& $enc $i.Titel) + '</title>')
    [void]$sb.Append('<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:1000px;margin:24px auto;padding:0 16px}img{max-width:100%;border:1px solid #ccc}h2{font-size:15px;color:#555;margin-top:28px}</style></head><body>')
    [void]$sb.Append('<h1>' + (& $enc $i.Titel) + '</h1>')
    [void]$sb.Append('<p>' + $i.Beginn.ToString('dddd, dd.MM.yyyy, HH:mm', $DE) + ' – ' + $i.Ende.ToString('HH:mm', $DE) + ' Uhr · ' + $i.Bilder.Count + ' Folie(n)</p><p>Notizen:</p>')
    $n = 0
    foreach ($b in $i.Bilder) {
        $n++
        [void]$sb.Append("<h2>Folie $n – $(Zeit-Aus-Datei $b)</h2><img src=""$($b.Name)"">")
    }
    [void]$sb.Append('</body></html>')
    $ziel = Join-Path $ordner 'Notiz.html'
    [IO.File]::WriteAllText($ziel, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))
    return $ziel
}

function OneNote-Abschnitte($on) {
    $xml = ''
    $on.GetHierarchy('', 3, [ref]$xml)   # 3 = Abschnitte
    [xml]$doc = $xml
    $ns = New-Object System.Xml.XmlNamespaceManager $doc.NameTable
    $ns.AddNamespace('one', $doc.DocumentElement.NamespaceURI)
    $liste = @()
    foreach ($s in $doc.SelectNodes('//one:Section', $ns)) {
        if ($s.isInRecycleBin -eq 'true' -or $s.isDeletedPages -eq 'true' -or $s.locked -eq 'true') { continue }
        $pfad = @()
        $p = $s.ParentNode
        while ($p -and $p.LocalName -ne 'Notebooks') {
            if ($p.isRecycleBin -eq 'true') { $pfad = $null; break }
            if ($p.name) { $pfad = , $p.name + $pfad }
            $p = $p.ParentNode
        }
        if ($null -eq $pfad) { continue }
        $liste += [pscustomobject]@{ Id = $s.ID; Name = (($pfad + $s.name) -join ' > ') }
    }
    return $liste
}

function Nach-OneNote([string]$ordner) {
    $e = Lade-Einstellungen
    try {
        $on = New-Object -ComObject OneNote.Application
    } catch {
        Write-Host 'OneNote (Desktop-App) lässt sich nicht ansteuern.' -ForegroundColor Yellow
        Write-Host 'Ausweg: Notiz.html im Mitschnitt-Ordner öffnen, Strg+A, Strg+C, in OneNote einfügen.'
        Start-Process (Join-Path $ordner 'Notiz.html')
        return $false
    }
    try {
        $abschnitte = @(OneNote-Abschnitte $on)
        if ($abschnitte.Count -eq 0) { throw 'Keine OneNote-Abschnitte gefunden. Bitte OneNote einmal öffnen und ein Notizbuch anlegen.' }
        $ziel = $abschnitte | Where-Object { $_.Id -eq $e.OneNoteAbschnittId } | Select-Object -First 1
        if ($ziel) {
            $a = Read-Host "Abschnitt '$($ziel.Name)' nehmen? [Enter = ja, A = anderen wählen]"
            if ($a -match '^[aA]') { $ziel = $null }
        }
        if (-not $ziel) {
            Write-Host ''
            for ($k = 0; $k -lt $abschnitte.Count; $k++) { Write-Host ('{0,3}  {1}' -f ($k + 1), $abschnitte[$k].Name) }
            $nr = Read-Host 'Nummer des Abschnitts'
            if ($nr -notmatch '^\d+$' -or [int]$nr -lt 1 -or [int]$nr -gt $abschnitte.Count) { throw 'Ungültige Nummer.' }
            $ziel = $abschnitte[[int]$nr - 1]
            $e.OneNoteAbschnittId = $ziel.Id; $e.OneNoteAbschnittName = $ziel.Name
            Speichere-Einstellungen $e
        }

        $i = Lies-Mitschnitt $ordner
        Write-Host "Lege Seite an und füge $($i.Bilder.Count) Bild(er) ein …"
        $seitenId = ''
        $on.CreateNewPage($ziel.Id, [ref]$seitenId, 0)

        $cd = { param($s) '<![CDATA[' + ($s -replace '\]\]>', ']]&gt;') + ']]>' }
        $breite = [double]$e.OneNoteBildbreite
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0"?><one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="' + $seitenId + '">')
        $seitentitel = $i.Beginn.ToString('yyyy-MM-dd', $DE) + ' – ' + $i.Titel
        [void]$sb.Append('<one:Title><one:OE><one:T>' + (& $cd $seitentitel) + '</one:T></one:OE></one:Title>')
        [void]$sb.Append('<one:Outline><one:OEChildren>')
        [void]$sb.Append('<one:OE><one:T>' + (& $cd ('Termin: ' + $i.Titel)) + '</one:T></one:OE>')
        [void]$sb.Append('<one:OE><one:T>' + (& $cd ('Datum: ' + $i.Beginn.ToString('dddd, dd.MM.yyyy', $DE) + ', ' + $i.Beginn.ToString('HH:mm') + ' – ' + $i.Ende.ToString('HH:mm') + ' Uhr')) + '</one:T></one:OE>')
        [void]$sb.Append('<one:OE><one:T>' + (& $cd ('Folien: ' + $i.Bilder.Count + ' (automatisch aufgenommen)')) + '</one:T></one:OE>')
        [void]$sb.Append('<one:OE><one:T>' + (& $cd 'Notizen:') + '</one:T></one:OE><one:OE><one:T><![CDATA[]]></one:T></one:OE>')
        $n = 0
        foreach ($b in $i.Bilder) {
            $n++
            $bild = [System.Drawing.Image]::FromFile($b.FullName)
            $hoehe = [Math]::Round($breite * $bild.Height / $bild.Width, 1)
            $bild.Dispose()
            $daten = [Convert]::ToBase64String([IO.File]::ReadAllBytes($b.FullName))
            [void]$sb.Append('<one:OE><one:T>' + (& $cd ("Folie $n – " + (Zeit-Aus-Datei $b))) + '</one:T></one:OE>')
            [void]$sb.Append('<one:OE><one:Image format="png"><one:Size width="' + $breite.ToString([Globalization.CultureInfo]::InvariantCulture) + '" height="' + $hoehe.ToString([Globalization.CultureInfo]::InvariantCulture) + '" isSetByUser="true" /><one:Data>' + $daten + '</one:Data></one:Image></one:OE>')
        }
        [void]$sb.Append('</one:OEChildren></one:Outline></one:Page>')
        $on.UpdatePageContent($sb.ToString())
        try { $on.NavigateTo($seitenId) } catch {}
        Write-Host "Fertig: Seite '$seitentitel' in '$($ziel.Name)' angelegt." -ForegroundColor Green
        return $true
    } catch {
        Write-Host "OneNote-Fehler: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host 'Ausweg: Notiz.html öffnet sich gleich – Strg+A, Strg+C, in OneNote einfügen.'
        Start-Process (Join-Path $ordner 'Notiz.html')
        return $false
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($on)
    }
}

function Mitschnitt-Nach-OneNote {
    $ordner = @(Get-ChildItem (Join-Path $Basis 'Mitschnitte') -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 15)
    if ($ordner.Count -eq 0) { Write-Host 'Noch keine Mitschnitte vorhanden.'; Pause-Taste; return }
    for ($k = 0; $k -lt $ordner.Count; $k++) { Write-Host ('{0,3}  {1}' -f ($k + 1), $ordner[$k].Name) }
    $nr = Read-Host 'Nummer des Mitschnitts'
    if ($nr -notmatch '^\d+$' -or [int]$nr -lt 1 -or [int]$nr -gt $ordner.Count) { return }
    $o = $ordner[[int]$nr - 1].FullName
    Schreibe-HtmlNotiz $o | Out-Null
    [void](Nach-OneNote $o)
    Pause-Taste
}

# ================================================================ TEIL 2: Playwright

$script:Laufzeit = $null
$script:Python = 'python'
$script:CliPfad = $null

function Finde-Laufzeit {
    if ($script:Laufzeit) { return $script:Laufzeit }
    $e = Lade-Einstellungen
    if (Get-Command node -ErrorAction SilentlyContinue) {
        $pfade = @()
        if ($e.PlaywrightOrdner) { $pfade += (Join-Path $e.PlaywrightOrdner 'node_modules') }
        $pfade += (Join-Path $Basis 'node_modules')
        if (Get-Command npm -ErrorAction SilentlyContinue) {
            $g = (& npm root -g 2>$null | Select-Object -First 1)
            if ($g) { $pfade += $g.Trim() }
        }
        $env:NODE_PATH = ($pfade -join ';')
        $cli = & node -e "const p=require('path');let r;try{r=require.resolve('playwright/package.json')}catch(e){r=require.resolve('@playwright/test/package.json')}process.stdout.write(p.join(p.dirname(r),'cli.js'))" 2>$null
        if ($LASTEXITCODE -eq 0 -and $cli) { $script:CliPfad = $cli; $script:Laufzeit = 'node'; return 'node' }
    }
    foreach ($py in 'python', 'py', 'python3') {
        if (Get-Command $py -ErrorAction SilentlyContinue) {
            & $py -c "import playwright" 2>$null
            if ($LASTEXITCODE -eq 0) { $script:Python = $py; $script:Laufzeit = 'python'; return 'python' }
        }
    }
    return $null
}

function Brauche-Laufzeit {
    $l = Finde-Laufzeit
    if (-not $l) {
        Write-Host 'Playwright wurde weder für Node noch für Python gefunden.' -ForegroundColor Red
        Write-Host 'Siehe ANLEITUNG.html, Abschnitt "Wenn Playwright nicht gefunden wird".'
        Pause-Taste
    }
    return $l
}

function Setze-Umgebung {
    $e = Lade-Einstellungen
    $env:PW_BROWSER = $e.Browser
    $speicher = Join-Path $Basis 'anmeldung.json'
    if (Test-Path $speicher) { $env:PW_SPEICHER = $speicher } else { $env:PW_SPEICHER = '' }
}

function Playwright-Cli([string[]]$argumente) {
    if ($script:Laufzeit -eq 'node') { & node $script:CliPfad @argumente }
    else { & $script:Python -m playwright @argumente }
}

function Helfer([string[]]$argumente) {
    Setze-Umgebung
    if ($script:Laufzeit -eq 'node') { & node (Join-Path $Basis 'helfer.js') @argumente }
    else { & $script:Python (Join-Path $Basis 'helfer.py') @argumente }
}

function Browser-Argumente {
    $e = Lade-Einstellungen
    if ($e.Browser -and $e.Browser -ne 'chromium') { return @("--channel=$($e.Browser)") }
    return @()
}

function Neuer-Ergebnisordner([string]$art) {
    $o = Join-Path $Basis ('Ergebnisse\{0}_{1:yyyy-MM-dd_HHmmss}' -f $art, (Get-Date))
    New-Item -ItemType Directory -Force -Path $o | Out-Null
    return $o
}

function Ablauf-Aufnehmen {
    if (-not (Brauche-Laufzeit)) { return }
    $name = Sauberer-Name (Read-Host 'Name für den Ablauf (z. B. Zeiterfassung)')
    $url = Read-Host 'Startadresse (z. B. https://intranet.firma.de) – Enter = leer'
    $endung = if ($script:Laufzeit -eq 'node') { 'js' } else { 'py' }
    $ziel = if ($script:Laufzeit -eq 'node') { 'javascript' } else { 'python' }
    $datei = Join-Path $Basis "Ablaeufe\$name.$endung"
    $arg = @('codegen', "--target=$ziel", "--output=$datei", '--lang=de-DE') + (Browser-Argumente)
    $speicher = Join-Path $Basis 'anmeldung.json'
    if (Test-Path $speicher) { $arg += "--load-storage=$speicher" }
    if ($url) { $arg += $url }
    Write-Host ''
    Write-Host 'Es öffnen sich zwei Fenster: der Browser und der "Playwright Inspector".'
    Write-Host 'Klicke dich im Browser durch. Jeder Schritt wird mitgeschrieben.'
    Write-Host 'Zum Schluss den Browser schließen – dann wird gespeichert.'
    Playwright-Cli $arg
    if (Test-Path $datei) { Write-Host "Gespeichert: $datei" -ForegroundColor Green } else { Write-Host 'Es wurde nichts gespeichert.' -ForegroundColor Yellow }
    Pause-Taste
}

function Ablauf-Abspielen {
    if (-not (Brauche-Laufzeit)) { return }
    $endung = if ($script:Laufzeit -eq 'node') { '*.js' } else { '*.py' }
    $liste = @(Get-ChildItem (Join-Path $Basis 'Ablaeufe') -Filter $endung -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($liste.Count -eq 0) { Write-Host 'Noch keine Abläufe aufgenommen (Menüpunkt 2).'; Pause-Taste; return }
    for ($k = 0; $k -lt $liste.Count; $k++) { Write-Host ('{0,3}  {1}' -f ($k + 1), $liste[$k].BaseName) }
    $nr = Read-Host 'Nummer des Ablaufs'
    if ($nr -notmatch '^\d+$' -or [int]$nr -lt 1 -or [int]$nr -gt $liste.Count) { return }
    $datei = $liste[[int]$nr - 1].FullName
    if ($script:Laufzeit -eq 'node') {
        # Falls nur @playwright/test installiert ist, require('playwright') umbiegen
        & node -e "require('playwright')" 2>$null
        if ($LASTEXITCODE -ne 0) {
            $tmp = Join-Path $env:TEMP ('pw_ablauf_' + [IO.Path]::GetFileName($datei))
            (Get-Content $datei -Raw -Encoding UTF8) -replace "require\('playwright'\)", "require('@playwright/test')" | Set-Content $tmp -Encoding UTF8
            $datei = $tmp
        }
        & node $datei
    } else {
        & $script:Python $datei
    }
    if ($LASTEXITCODE -eq 0) { Write-Host 'Ablauf fertig.' -ForegroundColor Green } else { Write-Host 'Ablauf mit Fehler beendet (siehe oben).' -ForegroundColor Red }
    Pause-Taste
}

function Frage-Adressen([string]$listenDatei) {
    $u = Read-Host "Adresse eingeben – oder Enter für alle Adressen aus listen\$listenDatei"
    if ($u) { return $u }
    return (Join-Path $Basis "listen\$listenDatei")
}

function Bilder-Machen {
    if (-not (Brauche-Laufzeit)) { return }
    $quelle = Frage-Adressen 'seiten.txt'
    $o = Neuer-Ergebnisordner 'Bilder'
    Helfer @('bilder', $quelle, $o)
    Invoke-Item $o
    Pause-Taste
}

function Pdfs-Machen {
    if (-not (Brauche-Laufzeit)) { return }
    $quelle = Frage-Adressen 'seiten.txt'
    $o = Neuer-Ergebnisordner 'PDF'
    Helfer @('pdf', $quelle, $o)
    Invoke-Item $o
    Pause-Taste
}

function Seiten-Pruefen {
    if (-not (Brauche-Laufzeit)) { return }
    $o = Neuer-Ergebnisordner 'Pruefung'
    Helfer @('pruefen', (Join-Path $Basis 'listen\pruefen.txt'), $o)
    Invoke-Item $o
    Pause-Taste
}

function Tabelle-Auslesen {
    if (-not (Brauche-Laufzeit)) { return }
    $url = Read-Host 'Adresse der Seite mit der Tabelle'
    if (-not $url) { return }
    $sel = Read-Host 'Nur bestimmte Tabelle? CSS-Auswahl, z. B. #umsatz (Enter = alle Tabellen)'
    $o = Neuer-Ergebnisordner 'Tabelle'
    $arg = @('tabelle', $url, $o)
    if ($sel) { $arg += $sel }
    Helfer $arg
    Invoke-Item $o
    Pause-Taste
}

function Anmeldung-Speichern {
    if (-not (Brauche-Laufzeit)) { return }
    $url = Read-Host 'Adresse der Anmeldeseite'
    if (-not $url) { return }
    Helfer @('anmelden', $url, (Join-Path $Basis 'anmeldung.json'))
    Pause-Taste
}

function Einrichtung-Pruefen {
    Clear-Host
    Write-Host '=== Einrichtung prüfen ===' -ForegroundColor Cyan
    Write-Host ("Windows PowerShell : {0}" -f $PSVersionTable.PSVersion)
    $script:Laufzeit = $null
    $l = Finde-Laufzeit
    if ($l -eq 'node') {
        Write-Host ("Playwright        : Node {0}, {1}" -f (& node --version), $script:CliPfad) -ForegroundColor Green
        Write-Host ("Version           : {0}" -f (Playwright-Cli @('--version')))
    } elseif ($l -eq 'python') {
        Write-Host ("Playwright        : Python ({0})" -f $script:Python) -ForegroundColor Green
        Write-Host ("Version           : {0}" -f (Playwright-Cli @('--version')))
    } else {
        Write-Host 'Playwright        : NICHT gefunden (Teams/OneNote geht trotzdem)' -ForegroundColor Yellow
    }
    $e = Lade-Einstellungen
    Write-Host ("Browser           : {0}" -f $e.Browser)
    try { $on = New-Object -ComObject OneNote.Application; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($on); Write-Host 'OneNote           : gefunden' -ForegroundColor Green }
    catch { Write-Host 'OneNote           : Desktop-App nicht ansteuerbar (Ausweg: Notiz.html)' -ForegroundColor Yellow }
    $t = Get-Process -Name 'ms-teams', 'Teams', 'msteams' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($t) { Write-Host "Teams             : läuft ($($t.Name))" -ForegroundColor Green } else { Write-Host 'Teams             : läuft gerade nicht' }
    if (Test-Path (Join-Path $Basis 'anmeldung.json')) { Write-Host 'Anmeldung         : gespeichert (anmeldung.json)' }
    Pause-Taste
}

# ================================================================ Menü

if ($Teams) { Starte-TeamsMitschnitt; exit }

while ($true) {
    Clear-Host
    Write-Host '==============================================' -ForegroundColor Cyan
    Write-Host '   Playwright-Werkzeug' -ForegroundColor Cyan
    Write-Host '==============================================' -ForegroundColor Cyan
    Write-Host '  TEAMS & ONENOTE'
    Write-Host '   1  Teams-Präsentation mitschneiden (→ OneNote)'
    Write-Host '   T  Fertigen Mitschnitt nach OneNote übertragen'
    Write-Host ''
    Write-Host '  BROWSER (Playwright)'
    Write-Host '   2  Ablauf aufnehmen'
    Write-Host '   3  Ablauf abspielen'
    Write-Host '   4  Screenshots von Webseiten'
    Write-Host '   5  PDFs von Webseiten'
    Write-Host '   6  Webseiten prüfen (listen\pruefen.txt)'
    Write-Host '   7  Tabelle von Webseite → Excel (CSV)'
    Write-Host '   8  Anmeldung speichern (für geschützte Seiten)'
    Write-Host ''
    Write-Host '   P  Einrichtung prüfen     E  Einstellungen'
    Write-Host '   A  Anleitung              0  Beenden'
    Write-Host ''
    $wahl = (Read-Host 'Auswahl').Trim().ToUpper()
    switch ($wahl) {
        '1' { Starte-TeamsMitschnitt }
        'T' { Mitschnitt-Nach-OneNote }
        '2' { Ablauf-Aufnehmen }
        '3' { Ablauf-Abspielen }
        '4' { Bilder-Machen }
        '5' { Pdfs-Machen }
        '6' { Seiten-Pruefen }
        '7' { Tabelle-Auslesen }
        '8' { Anmeldung-Speichern }
        'P' { Einrichtung-Pruefen }
        'E' { if (-not (Test-Path $EinstDatei)) { Speichere-Einstellungen (Lade-Einstellungen) }; Start-Process notepad.exe $EinstDatei -Wait; $script:Laufzeit = $null }
        'A' { Start-Process (Join-Path $Basis 'ANLEITUNG.html') }
        '0' { exit }
    }
}
