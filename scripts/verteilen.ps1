<#
.SYNOPSIS
  Verteilt das Add-in "Mail-Ablage" als Organisations-Add-in in Exchange Online.

.DESCRIPTION
  1. Meldet bei Exchange Online an, falls noch keine Sitzung besteht.
  2. Entfernt persönliche Installationen der angegebenen Postfächer (sonst erscheint das Add-in doppelt
     bzw. Exchange lehnt die Organisations-Installation für diese Benutzer ab).
  3. Entfernt eine vorhandene Organisations-Installation (Update nach Manifest-Änderung).
  4. Installiert das Add-in für die angegebenen Benutzer oder für alle, standardmäßig aktiviert.

  Mit -WhatIf zeigt das Skript nur an, was es tun würde.

.EXAMPLE
  .\scripts\verteilen.ps1 -Benutzer "hendrik.lindemann@sbhl.de","kollegin@sbhl.de" -WhatIf

.EXAMPLE
  .\scripts\verteilen.ps1 -Benutzer "hendrik.lindemann@sbhl.de","kollegin@sbhl.de"

.EXAMPLE
  .\scripts\verteilen.ps1 -Alle -PersoenlichEntfernen "hendrik.lindemann@sbhl.de"
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "Benutzer")]
param(
    # Postfächer, die das Add-in erhalten
    [Parameter(Mandatory, ParameterSetName = "Benutzer")]
    [string[]] $Benutzer,

    # Für alle Benutzer der Organisation bereitstellen
    [Parameter(Mandatory, ParameterSetName = "Alle")]
    [switch] $Alle,

    # Postfächer mit persönlicher Installation (per New-App -Mailbox). Standard: die Liste aus -Benutzer.
    [string[]] $PersoenlichEntfernen,

    # Benutzer können das Add-in nicht deaktivieren
    [switch] $ImmerAktiv,

    [string] $ManifestUrl = "https://hendriklindemann.github.io/outlook-mail-ablage/manifest.xml"
)

$ErrorActionPreference = "Stop"
$AppId = "5475c721-578c-4dff-965e-1659d7a1b9a2"   # <Id> aus manifest.template.xml

# --- Anmeldung -------------------------------------------------------------
if (-not (Get-Command New-App -ErrorAction SilentlyContinue)) {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowBanner:$false
}

# --- Manifest erreichbar? --------------------------------------------------
try {
    $null = Invoke-WebRequest -Uri $ManifestUrl -UseBasicParsing -Method Head
}
catch {
    throw "Manifest nicht erreichbar: $ManifestUrl ($($_.Exception.Message))"
}

# --- Persönliche Installationen entfernen ----------------------------------
if (-not $PersoenlichEntfernen) { $PersoenlichEntfernen = $Benutzer }
foreach ($mb in $PersoenlichEntfernen) {
    $app = Get-App -Mailbox $mb -Identity $AppId -ErrorAction SilentlyContinue
    if (-not $app) {
        Write-Host "Keine persönliche Installation: $mb"
    }
    elseif ($PSCmdlet.ShouldProcess($mb, "Persönliche Installation von Mail-Ablage entfernen")) {
        Remove-App -Mailbox $mb -Identity $AppId -Confirm:$false
        Write-Host "Persönliche Installation entfernt: $mb"
    }
}

# --- Vorhandene Organisations-Installation ersetzen ------------------------
$org = Get-App -OrganizationApp -Identity $AppId -ErrorAction SilentlyContinue
if ($org -and $PSCmdlet.ShouldProcess("Organisation", "Bisherige Organisations-Installation (Version $($org.AppVersion)) entfernen")) {
    Remove-App -OrganizationApp -Identity $AppId -Confirm:$false
    Write-Host "Bisherige Organisations-Installation entfernt."
}

# --- Installieren ----------------------------------------------------------
$params = @{
    OrganizationApp     = $true
    Url                 = $ManifestUrl
    DefaultStateForUser = $(if ($ImmerAktiv) { "AlwaysEnabled" } else { "Enabled" })
}
if ($Alle) {
    $params.ProvidedTo = "Everyone"
    $ziel = "alle Benutzer"
}
else {
    $params.ProvidedTo = "SpecificUsers"
    $params.UserList = $Benutzer
    $ziel = $Benutzer -join ", "
}

if ($PSCmdlet.ShouldProcess($ziel, "Mail-Ablage als Organisations-Add-in installieren")) {
    New-App @params | Format-Table DisplayName, AppVersion, Enabled, ProvidedTo, DefaultStateForUser -AutoSize
    Write-Host "Installiert für: $ziel"
    Write-Host "Kontrolle: Get-App -OrganizationApp | Format-Table DisplayName,AppVersion,ProvidedTo,UserList"
    Write-Host "Hinweis: Bis das Add-in bei allen Benutzern erscheint, können einige Stunden bis zwei Tage vergehen."
}
