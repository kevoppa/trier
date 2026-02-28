Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- CONFIGURATION COULEURS ---
$colorBg = [Drawing.Color]::FromArgb(25, 25, 25)
$colorBtnTri = [Drawing.Color]::FromArgb(0, 150, 136) 
$colorBtnAnnuler = [Drawing.Color]::FromArgb(156, 39, 176)
$colorText = [Drawing.Color]::White
$colorAccent = [Drawing.Color]::FromArgb(0, 120, 215)

# --- FENÊTRE PRINCIPALE ---
$form = New-Object Windows.Forms.Form
$form.Text = "TRIEUR INTELLIGENT PRO"
$form.Size = New-Object Drawing.Size(500, 520)
$form.BackColor = $colorBg
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$lblTitre = New-Object Windows.Forms.Label
$lblTitre.Text = "TOOLBOX TRI"
$lblTitre.Font = New-Object Drawing.Font("Segoe UI", 20, [Drawing.FontStyle]::Bold)
$lblTitre.ForeColor = $colorText
$lblTitre.TextAlign = "MiddleCenter"
$lblTitre.Size = New-Object Drawing.Size(480, 60)
$lblTitre.Location = New-Object Drawing.Point(10, 30)
$form.Controls.Add($lblTitre)

$btnTrier = New-Object Windows.Forms.Button
$btnTrier.Text = "LANCER UN TRI"
$btnTrier.Size = New-Object Drawing.Size(380, 80)
$btnTrier.Location = New-Object Drawing.Point(60, 110)
$btnTrier.FlatStyle = "Flat"
$btnTrier.BackColor = $colorBtnTri
$btnTrier.ForeColor = $colorText
$btnTrier.Font = New-Object Drawing.Font("Segoe UI", 14, [Drawing.FontStyle]::Bold)
$btnTrier.FlatAppearance.BorderSize = 0
$btnTrier.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnTrier.Add_Click({ Executer-Tri })
$form.Controls.Add($btnTrier)

$btnAnnuler = New-Object Windows.Forms.Button
$btnAnnuler.Text = "ANNULER UN TRI"
$btnAnnuler.Size = New-Object Drawing.Size(380, 80)
$btnAnnuler.Location = New-Object Drawing.Point(60, 200)
$btnAnnuler.FlatStyle = "Flat"
$btnAnnuler.BackColor = $colorBtnAnnuler
$btnAnnuler.ForeColor = $colorText
$btnAnnuler.Font = New-Object Drawing.Font("Segoe UI", 14, [Drawing.FontStyle]::Bold)
$btnAnnuler.FlatAppearance.BorderSize = 0
$btnAnnuler.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnAnnuler.Add_Click({ Executer-Annulation })
$form.Controls.Add($btnAnnuler)

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Size = New-Object Drawing.Size(380, 15)
$progressBar.Location = New-Object Drawing.Point(60, 310)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

$lblCount = New-Object Windows.Forms.Label
$lblCount.Text = ""
$lblCount.ForeColor = $colorText
$lblCount.TextAlign = "MiddleCenter"
$lblCount.Size = New-Object Drawing.Size(480, 20)
$lblCount.Location = New-Object Drawing.Point(10, 330)
$lblCount.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$form.Controls.Add($lblCount)

$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Text = "Prêt à l'emploi"
$statusLabel.ForeColor = [Drawing.Color]::Gray
$statusLabel.TextAlign = "MiddleCenter"
$statusLabel.Size = New-Object Drawing.Size(480, 40)
$statusLabel.Location = New-Object Drawing.Point(10, 400)
$statusLabel.Font = New-Object Drawing.Font("Segoe UI", 11)
$form.Controls.Add($statusLabel)

# --- FENÊTRE : CONFIRMATION STYLISÉE ---
function Show-StyledConfirmation($count, $totalMB, $totalGB, $limitMo, $limitGo, $folders) {
    $cForm = New-Object Windows.Forms.Form
    $cForm.Text = "CONFIRMATION"
    $cForm.Size = "450,400"
    $cForm.BackColor = $colorBg
    $cForm.StartPosition = "CenterParent"
    $cForm.FormBorderStyle = "FixedDialog"

    $ico = New-Object Windows.Forms.Label
    $ico.Text = "i"
    $ico.Font = New-Object Drawing.Font("Segoe UI", 30, [Drawing.FontStyle]::Bold)
    $ico.ForeColor = $colorAccent
    $ico.Location = "20,20"; $ico.Size = "50,50"; $ico.TextAlign = "MiddleCenter"
    $cForm.Controls.Add($ico)

    $txt = New-Object Windows.Forms.Label
    $txt.Text = "Le logiciel s'apprête à trier :`n`n" +
                "• $count fichiers`n" +
                "• Poids total : $totalMB Mo ($totalGB Go)`n" +
                "• Poids max / dossier : $limitMo Mo ($limitGo Go)`n`n" +
                "Dossiers créés : environ '' $folders ''"
    $txt.ForeColor = $colorText
    $txt.Font = New-Object Drawing.Font("Segoe UI", 11)
    $txt.Location = "80,30"; $txt.Size = "330,200"
    $cForm.Controls.Add($txt)

    $btnOui = New-Object Windows.Forms.Button
    $btnOui.Text = "OUI"; $btnOui.Size = "150,45"; $btnOui.Location = "60,280"
    $btnOui.FlatStyle = "Flat"; $btnOui.BackColor = $colorBtnTri; $btnOui.ForeColor = $colorText
    $btnOui.Add_Click({ $cForm.DialogResult = [Windows.Forms.DialogResult]::Yes; $cForm.Close() })
    $cForm.Controls.Add($btnOui)

    $btnNon = New-Object Windows.Forms.Button
    $btnNon.Text = "NON"; $btnNon.Size = "150,45"; $btnNon.Location = "230,280"
    $btnNon.FlatStyle = "Flat"; $btnNon.BackColor = [Drawing.Color]::DarkRed; $btnNon.ForeColor = $colorText
    $btnNon.Add_Click({ $cForm.DialogResult = [Windows.Forms.DialogResult]::No; $cForm.Close() })
    $cForm.Controls.Add($btnNon)

    return $cForm.ShowDialog()
}

# --- FENÊTRE : RÉSULTAT DÉTAILLÉ (FIN DE TRI) ---
function Show-FinalReport($total, $tooBigDetails) {
    $rForm = New-Object Windows.Forms.Form
    $rForm.Text = "RÉSULTAT DU TRI"
    $rForm.Size = "550,500"; $rForm.BackColor = $colorBg; $rForm.StartPosition = "CenterParent"

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = "TRI TERMINÉ AVEC SUCCÈS !"; $lbl.ForeColor = $colorBtnTri
    $lbl.Font = New-Object Drawing.Font("Segoe UI", 16, [Drawing.FontStyle]::Bold)
    $lbl.Size = "500,40"; $lbl.Location = "15,20"; $lbl.TextAlign = "MiddleCenter"
    $rForm.Controls.Add($lbl)

    $info = New-Object Windows.Forms.Label
    $info.Text = "Total traités : $total fichiers"; $info.ForeColor = $colorText
    $info.Location = "15,70"; $info.Size = "500,25"; $info.Font = New-Object Drawing.Font("Segoe UI", 10)
    $rForm.Controls.Add($info)

    if ($tooBigDetails.Count -gt 0) {
        $warn = New-Object Windows.Forms.Label
        $warn.Text = "Attention : fichiers dépassant la limite (placés seuls) :"; $warn.ForeColor = [Drawing.Color]::Gold
        $warn.Location = "15,100"; $warn.Size = "500,25"; $warn.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
        $rForm.Controls.Add($warn)

        $box = New-Object Windows.Forms.TextBox
        $box.Multiline = $true; $box.ReadOnly = $true; $box.BackColor = [Drawing.Color]::FromArgb(40,40,40); $box.ForeColor = $colorText
        $box.Size = "500,240"; $box.Location = "15,130"; $box.Font = New-Object Drawing.Font("Consolas", 9)
        $box.ScrollBars = "Vertical"
        
        $reportText = ""
        foreach($item in $tooBigDetails) {
            $reportText += "FICHIER : $($item.Name)`r`nEMPLACEMENT : $($item.Folder)`r`n--------------------------`r`n"
        }
        $box.Text = $reportText
        $rForm.Controls.Add($box)
    }

    $btnOk = New-Object Windows.Forms.Button
    $btnOk.Text = "FERMER"; $btnOk.Size = "200,45"; $btnOk.Location = "175,400"
    $btnOk.FlatStyle = "Flat"; $btnOk.BackColor = $colorAccent; $btnOk.ForeColor = $colorText
    $btnOk.Add_Click({ $rForm.Close() })
    $rForm.Controls.Add($btnOk)
    $rForm.ShowDialog()
}

# --- FONCTION : LISTE DÉROULANTE DES ERREURS ---
function Show-DetailedErrorReport($filesList, $limitMo, $limitGo) {
    $reportForm = New-Object Windows.Forms.Form
    $reportForm.Text = "CONFLIT DE TAILLE - DETAIL"
    $reportForm.Size = New-Object Drawing.Size(600, 550)
    $reportForm.BackColor = [Drawing.Color]::FromArgb(30, 30, 30)
    $reportForm.StartPosition = "CenterParent"
    $reportForm.FormBorderStyle = "FixedDialog"

    $header = New-Object Windows.Forms.Label
    $header.Text = "ALERTE : Voici les fichiers qui dépassent la limite de $limitMo Mo ($limitGo Go) :"
    $header.ForeColor = [Drawing.Color]::Gold
    $header.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
    $header.Size = New-Object Drawing.Size(560, 45)
    $header.Location = New-Object Drawing.Point(15, 15)
    $reportForm.Controls.Add($header)

    $container = New-Object Windows.Forms.Panel
    $container.Size = New-Object Drawing.Size(555, 350)
    $container.Location = New-Object Drawing.Point(15, 70)
    $container.AutoScroll = $true
    $reportForm.Controls.Add($container)

    $yPos = 0
    foreach ($f in $filesList) {
        $fMo = [Math]::Round($f.Length / 1MB, 0)
        $fGo = [Math]::Round($f.Length / 1GB, 2)
        
        $p = New-Object Windows.Forms.Panel
        $p.Size = New-Object Drawing.Size(520, 100)
        $p.Location = New-Object Drawing.Point(0, $yPos)
        $p.BackColor = [Drawing.Color]::FromArgb(45, 45, 45)
        $p.BorderStyle = "FixedSingle"
        
        $txtInfo = New-Object Windows.Forms.TextBox
        $txtInfo.Multiline = $true
        $txtInfo.ReadOnly = $true
        $txtInfo.BorderStyle = "None"
        $txtInfo.BackColor = [Drawing.Color]::FromArgb(45, 45, 45)
        $txtInfo.ForeColor = [Drawing.Color]::White
        $txtInfo.Font = New-Object Drawing.Font("Segoe UI", 8, [Drawing.FontStyle]::Bold)
        $txtInfo.Text = "NOM : $($f.Name)`r`nTAILLE : $fMo Mo ($fGo Go)`r`nCHEMIN : $($f.FullName)"
        $txtInfo.Size = New-Object Drawing.Size(480, 80)
        $txtInfo.Location = New-Object Drawing.Point(10, 10)
        $p.Controls.Add($txtInfo)

        $container.Controls.Add($p)
        $yPos += 110
    }

    $btnContinue = New-Object Windows.Forms.Button
    $btnContinue.Text = "CONTINUER LE TRI"
    $btnContinue.Size = New-Object Drawing.Size(200, 45)
    $btnContinue.Location = New-Object Drawing.Point(80, 440)
    $btnContinue.BackColor = $colorAccent
    $btnContinue.ForeColor = [Drawing.Color]::White
    $btnContinue.FlatStyle = "Flat"
    $btnContinue.Add_Click({ $reportForm.DialogResult = [Windows.Forms.DialogResult]::Yes; $reportForm.Close() })
    $reportForm.Controls.Add($btnContinue)

    $btnStop = New-Object Windows.Forms.Button
    $btnStop.Text = "ANNULER"
    $btnStop.Size = New-Object Drawing.Size(200, 45)
    $btnStop.Location = New-Object Drawing.Point(300, 440)
    $btnStop.BackColor = [Drawing.Color]::DarkRed
    $btnStop.ForeColor = [Drawing.Color]::White
    $btnStop.FlatStyle = "Flat"
    $btnStop.Add_Click({ $reportForm.DialogResult = [Windows.Forms.DialogResult]::No; $reportForm.Close() })
    $reportForm.Controls.Add($btnStop)

    return $reportForm.ShowDialog()
}

# --- FENÊTRE DE CHOIX DU NOM ---
function Get-NamingChoice($parent) {
    $nForm = New-Object Windows.Forms.Form
    $nForm.Text = "Option de nommage"; $nForm.Size = "400,320"; $nForm.StartPosition = "CenterParent"; $nForm.BackColor = "White"; $nForm.FormBorderStyle = "FixedToolWindow"
    $l = New-Object Windows.Forms.Label; $l.Text = "Comment renommer ?"; $l.Location = "20,20"; $l.Size = "300,25"; $l.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold); $nForm.Controls.Add($l)
    $opt1 = New-Object Windows.Forms.RadioButton; $opt1.Text = "Nom perso :"; $opt1.Location = "30,60"; $opt1.Size = "110,25"; $opt1.Checked = $true; $nForm.Controls.Add($opt1)
    $txt = New-Object Windows.Forms.TextBox; $txt.Location = "145,62"; $txt.Size = "150,25"; $nForm.Controls.Add($txt)
    $opt2 = New-Object Windows.Forms.RadioButton; $opt2.Text = "Juste chiffres (01, 02...)"; $opt2.Location = "30,100"; $opt2.Size = "250,25"; $nForm.Controls.Add($opt2)
    $opt3 = New-Object Windows.Forms.RadioButton; $opt3.Text = "Nom du dossier original ($parent)"; $opt3.Location = "30,140"; $opt3.Size = "320,25"; $nForm.Controls.Add($opt3)
    $btnOk = New-Object Windows.Forms.Button; $btnOk.Text = "VALIDER"; $btnOk.Location = "100,210"; $btnOk.Size = "180,45"; $btnOk.FlatStyle = "Flat"; $btnOk.BackColor = "Black"; $btnOk.ForeColor = "White"; $btnOk.Add_Click({ $nForm.DialogResult = "OK"; $nForm.Close() }); $nForm.Controls.Add($btnOk)
    if ($nForm.ShowDialog() -eq "OK") { if ($opt1.Checked) { return ($txt.Text + " ").TrimStart() }; if ($opt2.Checked) { return "CHIFRE_ONLY" }; if ($opt3.Checked) { return "$parent " } }
    return $null
}

# --- FENÊTRE DE CHOIX DU POIDS ---
function Get-SizeChoice {
    $sForm = New-Object Windows.Forms.Form
    $sForm.Text = "REPARTITION DE LA MEMOIRE"; $sForm.Size = "450,480"; $sForm.BackColor = "White"; $sForm.StartPosition = "CenterParent"; $sForm.FormBorderStyle = "FixedToolWindow"
    $l1 = New-Object Windows.Forms.Label; $l1.Text = "Poids maximum par dossier (Mo) :"; $l1.Location = "30,30"; $l1.Size = "300,25"; $l1.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold); $sForm.Controls.Add($l1)
    $numSize = New-Object Windows.Forms.NumericUpDown; $numSize.Minimum = 1; $numSize.Maximum = 999999; $numSize.Value = 9500; $numSize.Location = "35,60"; $numSize.Size = "360,30"; $numSize.Font = New-Object Drawing.Font("Segoe UI", 12); $sForm.Controls.Add($numSize)
    $sizes = @(@("500 mo", 500), @("1 Go", 1024), @("2 Go", 2048), @("5 Go", 5120), @("10 Go", 10240), @("20 Go", 20480), @("50 Go", 51200), @("100 Go", 102400), @("200 Go", 204800))
    $posX = 35; $posY = 145; $i = 0
    foreach ($s in $sizes) { $b = New-Object Windows.Forms.Button; $b.Text = $s[0]; $val = $s[1]; $b.Size = "110,40"; $b.Location = New-Object Drawing.Point($posX, $posY); $b.FlatStyle = "Flat"; $b.Add_Click({ $numSize.Value = $val }.GetNewClosure()); $sForm.Controls.Add($b); $posX += 125; $i++; if ($i % 3 -eq 0) { $posX = 35; $posY += 55 } }
    $btnValider = New-Object Windows.Forms.Button; $btnValider.Text = "Confirmer mon choix"; $btnValider.Location = "100,360"; $btnValider.Size = "230,50"; $btnValider.FlatStyle = "Flat"; $btnValider.BackColor = "Black"; $btnValider.ForeColor = "White"; $btnValider.Add_Click({ $sForm.DialogResult = "OK"; $sForm.Close() }); $sForm.Controls.Add($btnValider)
    if ($sForm.ShowDialog() -eq "OK") { return $numSize.Value }; return $null
}

# --- FONCTIONS LOGIQUES ---
function Executer-Tri {
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($browser.ShowDialog() -ne "OK") { return }
    $sourceDir = $browser.SelectedPath
    $baseName = Get-NamingChoice (Split-Path $sourceDir -Leaf)
    if ($null -eq $baseName) { return }
    if ($baseName -eq "CHIFRE_ONLY") { $baseName = "" }
    $sizeMo = Get-SizeChoice
    if ($null -eq $sizeMo) { return }
    $maxBytes = [double]$sizeMo * 1MB
    $files = Get-ChildItem $sourceDir -File -Recurse | Where-Object { $_.DirectoryName -notmatch "\\(\d{2}|.* \d{2})$" -and $_.Name -ne "tri_archive.log" }
    if ($files.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Aucun fichier à trier."); return }

    $tooBigFiles = $files | Where-Object { $_.Length -gt $maxBytes }
    if ($tooBigFiles) {
        $maxSizeGB = [Math]::Round($sizeMo / 1024, 2)
        $check = Show-DetailedErrorReport $tooBigFiles $sizeMo $maxSizeGB
        if ($check -ne "Yes") { return }
    }

    $totalFiles = $files.Count
    $totalSizeBytes = ($files | Measure-Object -Property Length -Sum).Sum
    $totalSizeMB = [Math]::Round($totalSizeBytes / 1MB, 0)
    $totalSizeGB = [Math]::Round($totalSizeBytes / 1GB, 2)
    $maxSizeLimitGB = [Math]::Round($sizeMo / 1024, 2)
    $estFolders = [Math]::Ceiling($totalSizeMB / $sizeMo)

    if ((Show-StyledConfirmation $totalFiles $totalSizeMB $totalSizeGB $sizeMo $maxSizeLimitGB $estFolders) -ne "Yes") { return }

    $log = Join-Path $sourceDir "tri_archive.log"
    $fNum = 1; $curSize = 0; $dest = ""; $progressBar.Maximum = $totalFiles; $current = 0
    $tooBigReport = New-Object System.Collections.Generic.List[Object]

    foreach ($f in $files) {
        $current++; $statusLabel.Text = "Tri en cours..."; $lblCount.Text = "$current / $totalFiles"; $progressBar.Value = $current; $form.Refresh()
        Add-Content $log "$($f.Name)|$($f.DirectoryName)"
        
        if ($dest -eq "" -or ($curSize + $f.Length) -gt $maxBytes) {
            if ($dest -ne "") { $fNum++ }
            $folderName = $baseName + ("{0:D2}" -f $fNum)
            $dest = Join-Path $sourceDir $folderName
            New-Item $dest -ItemType Directory -Force | Out-Null
            $curSize = 0
        }

        if ($f.Length -gt $maxBytes) {
            $tooBigReport.Add([PSCustomObject]@{Name=$f.Name; Folder=(Split-Path $dest -Leaf)})
        }

        Move-Item $f.FullName $dest -Force
        $curSize += $f.Length
    }
    Get-ChildItem $sourceDir -Directory -Recurse | Sort-Object { $_.FullName.Length } -Descending | ForEach-Object { if ($_.Name -notmatch "(\d{2}|.* \d{2})$" -and (Get-ChildItem $_.FullName -Recurse).Count -eq 0) { Remove-Item $_.FullName -Force } }
    
    $statusLabel.Text = "Tri terminé !"; $lblCount.Text = "Succès : $totalFiles fichiers"
    explorer.exe $sourceDir
    Show-FinalReport $totalFiles $tooBigReport
}

function Executer-Annulation {
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($browser.ShowDialog() -ne "OK") { return }
    $sourceDir = $browser.SelectedPath
    $log = Join-Path $sourceDir "tri_archive.log"
    if (-not (Test-Path $log)) { return }
    $lines = Get-Content $log; $total = $lines.Count; $progressBar.Maximum = $total; $current = 0
    foreach ($line in $lines) {
        $current++; $lblCount.Text = "$current / $total"; $progressBar.Value = $current; $form.Refresh()
        $n, $o = $line.Split('|'); $f = Get-ChildItem $sourceDir -Filter $n -Recurse | Where-Object { $_.DirectoryName -match "\\(\d{2}|.* \d{2})$" }
        if ($f) { if (-not (Test-Path $o)) { New-Item $o -ItemType Directory -Force | Out-Null }; Move-Item $f.FullName $o -Force }
    }
    Get-ChildItem $sourceDir -Directory | Where-Object { $_.Name -match "(\d{2}|.* \d{2})$" } | Remove-Item -Recurse -Force
    Remove-Item $log -ErrorAction SilentlyContinue
    explorer.exe $sourceDir
    $statusLabel.Text = "Annulation terminée !"; $lblCount.Text = ""; $progressBar.Value = 0
}

$null = $form.ShowDialog()
[System.Environment]::Exit(0)