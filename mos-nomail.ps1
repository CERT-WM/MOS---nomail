#---------------------------------------------------------------------------------
# HRtV 2022
# Script om standaard voor MX, DMARC en SPF in te stellen voor mailloos domein
#
# 1.1.0
# - invoerveld voor organisatie ingebouwd
# 1.0.1
# - inbouwen basis errorchecks
# - toevoegen van Null-MX record
# 1.0.0
# - eerste versie!
#---------------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$FONT        = New-Object System.Drawing.Font('Verdana',10,[System.Drawing.FontStyle]::Regular)
$ORGANISATIE = "Het Waterschapshuis"

# ------------------------------------------------------------------------------
# Toevoegen tekst aan log, met specifieke kleur
# ------------------------------------------------------------------------------
function Add-Line($color, $addtext) {
  if ( $color -eq $null ) { $color = $tekst.ForeColor }
  $tekst.SelectionStart = $tekst.TextLength;
  $tekst.SelectionLength = 0;
  $tekst.SelectionColor = $color;
  $tekst.AppendText($addtext.toString());
  $tekst.AppendText("`r`n");
  $tekst.SelectionColor = $tekst.ForeColor;
  $tekst.SelectionStart = $tekst.TextLength;
  $tekst.ScrollToCaret()
}


# ------------------------------------------------------------------------------
# Genereren output
# ------------------------------------------------------------------------------
function Report-Status($domain) {
  
  $SOA   = Resolve-DnsName $domain -type SOA
  $MX    = Resolve-DnsName $domain -type MX
  $SPF   = (Resolve-DnsName $domain -type TXT  -erroraction 'SilentlyContinue').Strings  | Select-String "spf"
  $DMARC = [String](Resolve-DnsName ('_dmarc.' + $domain) -type TXT  -erroraction 'SilentlyContinue').Strings
  
  
  Add-Line "black" "------------------------------------------------------------------------------"
  Add-Line "black" "De primaire DNS server voor dit domein is:"
  Add-Line "green" $SOA.PrimaryServer
  Add-Line "darkcyan"  "(hieruit kun je de naam van de DNS provider afleiden)"

  Add-Line "black" ""
  Add-Line "black" "De mailserver(s) voor dit domein zijn:"
  foreach ($record in $MX) {
    Add-Line "green" $record.NameExchange
  }

  Add-Line "black" ""
  Add-Line "black" "Het SPF record voor dit domein is:"
  if ($SPF.length -eq 0)   { 
    Add-Line "red" "niet gevonden"
  } else {
    Add-Line "green" $SPF
  }
  
  Add-Line "black" ""
  Add-Line "black" "Het DMARC record voor dit domein is:"
  if ($DMARC.length -eq 0) { 
    Add-Line "red" "niet gevonden"
  } else {
    Add-Line "green" $DMARC
  }
  Add-Line "black" "------------------------------------------------------------------------------"
  
  Add-Line "black" ""
  Add-Line "black" "Op te stellen mailbericht voor opschonen domein zonder mail:"
  Add-Line "darkcyan" "(selecteer deze tekst en kopieer deze in een op te stellen mailbericht)"
  Add-Line "black" "------------------------------------------------------------------------------"
  Add-Line "black" ""
  Add-Line "black" "LS,"
  Add-Line "black" ""
  Add-Line "black" "$ORGANISATIE gebruikt een groot aantal domeinen, waaronder [$domain], waarvan het domeinbeheer bij jullie ligt."
  Add-Line "black" "Aangezien er niet van of naar dit domein gemaild wordt, wil ik u verzoeken de volgende mailaanpassingen voor het domein door te voeren: "
  Add-Line "black" ""
  Add-Line "black" "- Het verwijderen van de MX records van het domein:"
  foreach ($record in $MX) {
    Add-Line "red" ("     " + $record.NameExchange)
  }
  Add-Line "black" ""
  Add-Line "black" "- Het aanmaken van een Null-MX record voor het domein:"
  Add-Line "green" ("      " + $domain + ".     MX 0 .")
  Add-Line "black" ""
  Add-Line "black" "- Het aanpassen of aanmaken van het SPF record:"
  Add-Line "green" ("     " + $domain + ".     v=spf1 -all")
  Add-Line "black" ""
  Add-Line "black" "- Het aanpassen of aanmaken van het DMARC record:"
  Add-Line "green" ("     _dmarc." + $domain + ".     v=DMARC1; p=quarantine; pct=100;")
  Add-Line "black" ""
  Add-Line "black" "Bij voorbaat dank!"
  Add-Line "black" ""
  Add-Line "black" "------------------------------------------------------------------------------"
}


# ------------------------------------------------------------------------------
# GUI
#---------------------------------------------------------------------------------
$nomail                  = New-Object system.Windows.Forms.Form
$nomail.Text             = 'MOS no-mail'
$nomail.AutoSize         = $False
$nomail.MaximizeBox      = $False
$nomail.MinimizeBox      = $False    
$nomail.FormBorderStyle  = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$nomail.width            = 900
$nomail.Height           = 800

# - label
$lorg              = New-Object System.Windows.Forms.Label
$lorg.Name         = 'ldom'
$lorg.Text         = 'org'
$lorg.Left         = 10
$lorg.Top          = 12
$lorg.Width        = 50

# - tekst voor invoer domein
$torg              = New-Object System.Windows.Forms.TextBox
$torg.Name         = 'tdom'
$torg.Left         = $lorg.Left + $lorg.Width + 10
$torg.Top          = 10
$torg.Width        = 200
$torg.Multiline    = $False
$torg.BorderStyle  = 'FixedSingle'
$torg.Text         = $ORGANISATIE

# - label
$ldom              = New-Object System.Windows.Forms.Label
$ldom.Name         = 'ldom'
$ldom.Text         = 'domein'
$ldom.Left         = 10
$ldom.Top          = 42
$ldom.Width        = 50

# - tekst voor invoer domein
$tdom              = New-Object System.Windows.Forms.TextBox
$tdom.Name         = 'tdom'
$tdom.Left         = $ldom.Left + $ldom.Width + 10
$tdom.Top          = 40
$tdom.Width        = 100
$tdom.Multiline    = $False
$tdom.BorderStyle  = 'FixedSingle'

# - genereer button
$bgen               = New-Object System.Windows.Forms.Button
$bgen.Name          = 'bgen'
$bgen.Top           = 40
$bgen.Left          = $tdom.Left + $tdom.Width + 10
$bgen.Text          = 'genereer'
$bgen.Add_Click({
  $ORGANISATIE = $torg.Text
  $tekst.Clear()
  if ($tdom.Text.length -gt 0) {
    if ((resolve-dnsname $tdom.Text -erroraction 'SilentlyContinue')) {
      Report-Status $tdom.Text
    } else {
      Add-Line "red" "Dan dit domein niet controleren"
    }
  }
})

# - copy button
$bcopy               = New-Object System.Windows.Forms.Button
$bcopy.Name          = 'bcopy'
$bcopy.Top           = 40
$bcopy.Left          = $bgen.Left + $bgen.Width + 10
$bcopy.Text          = 'kopieer'
$bcopy.Add_Click({
  # - Write-Host $tekst.SelectedRtf
  [System.Windows.Forms.Clipboard]::SetText($tekst.SelectedRtf, 2)
})

# - tekstveld voor resultaat
$tekst             = New-Object System.Windows.Forms.RichTextBox
$tekst.Name        = 'tuid'
$tekst.Left        = 10
$tekst.Top         = 70
$tekst.Width       = 860
$tekst.Height      = 710
$tekst.Multiline   = $True
$tekst.Font        = $FONT

# -  main window
$nomail.Controls.Add($lorg)
$nomail.Controls.Add($torg)
$nomail.Controls.Add($ldom)
$nomail.Controls.Add($tdom)
$nomail.Controls.Add($bgen)
$nomail.Controls.Add($bcopy)
$nomail.Controls.Add($tekst)

$dummy = $nomail.ShowDialog()

Write-Host "kloar!"
#---------------------------------------------------------------------------------
