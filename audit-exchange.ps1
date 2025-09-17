<#
.SYNOPSIS
    Auditoria de infraestrutura de e-mail híbrida (Exchange On-Prem, Exchange Online, DNS MX/SPF/DKIM/DMARC)
    Gera relatório HTML detalhado.

.NOTES
    - Execute com conta que tenha privilégios em Exchange on-prem e Exchange Online.
    - Forneça os domínios a auditar via -Domains.
    - Não altera configurações, apenas coleta informações.
#>

param(
    [Parameter(Mandatory=$true)]
    [string[]] $Domains,

    [Parameter(Mandatory=$false)]
    [string] $OnPremExchangeFQDN = "",

    [Parameter(Mandatory=$false)]
    [string] $OutputHtml = ".\MailInfraAuditReport.html",

    [switch] $UseImplicitRemoting
)

# --- Função para garantir módulos ---
function Ensure-Module {
    param($Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Módulo $Name não encontrado. Tentando instalar via PSGallery..." -ForegroundColor Yellow
        try {
            Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop
        } catch {
            $msg = $_.Exception.Message
            Write-Warning "Falha ao instalar módulo $Name: ${msg}. Rode manualmente ou instale o módulo e reexecute."
        }
    } else {
        Write-Host "Módulo $Name disponível."
    }
}

# --- Pré-checagens ---
Write-Host "`n== Pré-checagens ==" -ForegroundColor Cyan
Ensure-Module -Name "ExchangeOnlineManagement"
Ensure-Module -Name "AzureAD"

# --- Conexão Exchange Online ---
$connectedEXO = $false
try {
    Write-Host "Conectando ao Exchange Online..." -ForegroundColor Cyan
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    $connectedEXO = $true
    Write-Host "Conectado ao Exchange Online." -ForegroundColor Green
} catch {
    $msg = $_.Exception.Message
    Write-Warning "Falha ao conectar ao Exchange Online: ${msg}"
}

# --- Conexão Exchange On-Prem (opcional) ---
$onPremConnected = $false
$onPremSession = $null
if ($UseImplicitRemoting -and ($OnPremExchangeFQDN -ne "")) {
    try {
        Write-Host "Conectando ao Exchange On-Prem ($OnPremExchangeFQDN)..." -ForegroundColor Cyan
        $cred = Get-Credential -Message "Credenciais com permissão no Exchange on-prem"
        $uri = "http://$OnPremExchangeFQDN/PowerShell/"
        $onPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $cred -ErrorAction Stop
        Import-PSSession $onPremSession -AllowClobber -DisableNameChecking -ErrorAction Stop | Out-Null
        $onPremConnected = $true
        Write-Host "Conectado ao Exchange On-Prem." -ForegroundColor Green
    } catch {
        $msg = $_.Exception.Message
        Write-Warning "Falha ao conectar ao Exchange On-Prem: ${msg}"
    }
}

# --- Estrutura de resultados ---
$result = [ordered]@{
    GeneratedOn    = (Get-Date).ToString("s")
    Host           = (hostname)
    Domains        = @()
    ExchangeOnline = @{}
    ExchangeOnPrem = @{}
    DNS            = @{}
    Reachability   = @()
    Warnings       = @()
}

# --- Coleta Exchange Online ---
if ($connectedEXO) {
    Write-Host "`n== Coleta Exchange Online ==" -ForegroundColor Cyan
    try {
        $result.ExchangeOnline.Connectivity = "Connected"
        try { $result.ExchangeOnline.InboundConnectors  = Get-InboundConnector  -ErrorAction Stop | Select Name, ConnectorType, SenderIPAddresses, SenderDomains, Enabled } catch { $msg=$_.Exception.Message; $result.ExchangeOnline.InboundConnectors="Erro: ${msg}" }
        try { $result.ExchangeOnline.OutboundConnectors = Get-OutboundConnector -ErrorAction Stop | Select Name, Type, ConnectorSource, Enabled, CloudServicesMailEnabled } catch { $msg=$_.Exception.Message; $result.ExchangeOnline.OutboundConnectors="Erro: ${msg}" }
        try { $result.ExchangeOnline.TransportRules     = Get-TransportRule    -ErrorAction Stop | Select Name, State, Priority } catch { $msg=$_.Exception.Message; $result.ExchangeOnline.TransportRules="Erro: ${msg}" }
        try { $result.ExchangeOnline.AcceptedDomains   = Get-AcceptedDomain  -ErrorAction Stop | Select DomainName, DomainType } catch { $msg=$_.Exception.Message; $result.ExchangeOnline.AcceptedDomains="Erro: ${msg}" }
        try { $result.ExchangeOnline.DKIM               = Get-DkimSigningConfig -ErrorAction Stop | Select Selector, Domain, Enabled } catch { $msg=$_.Exception.Message; $result.ExchangeOnline.DKIM="Erro: ${msg}" }

        $oc = $result.ExchangeOnline.OutboundConnectors | Where-Object { $_.ConnectorSource -match "OnPremises" -or $_.Name -match "Hybrid" }
        if ($oc) {
            $result.ExchangeOnline.PossibleCentralizedMailTransport = $true
            $result.ExchangeOnline.Notes = "Conectores que podem forçar tráfego via on-prem: " + ($oc | Select-Object -ExpandProperty Name -Unique | Out-String)
        } else {
            $result.ExchangeOnline.PossibleCentralizedMailTransport = $false
        }
    } catch {
        $msg = $_.Exception.Message
        $result.ExchangeOnline = "Erro geral na coleta EXO: ${msg}"
    }
} else {
    $result.Warnings += "Não conectado ao Exchange Online; coleta EXO ignorada."
}

# --- Coleta Exchange On-Prem ---
if ($onPremConnected) {
    Write-Host "`n== Coleta Exchange On-Prem ==" -ForegroundColor Cyan
    try {
        $result.ExchangeOnPrem.Connectivity = "Connected"
        try { $result.ExchangeOnPrem.SendConnectors    = Get-SendConnector   | Select Name, AddressSpaces, SmartHosts, Enabled } catch { $msg=$_.Exception.Message; $result.ExchangeOnPrem.SendConnectors="Erro: ${msg}" }
        try { $result.ExchangeOnPrem.ReceiveConnectors = Get-ReceiveConnector| Select Name, Bindings, RemoteIPRanges, AuthMechanism } catch { $msg=$_.Exception.Message; $result.ExchangeOnPrem.ReceiveConnectors="Erro: ${msg}" }
        try { $result.ExchangeOnPrem.TransportRules    = Get-TransportRule   | Select Name, State, Priority } catch { $msg=$_.Exception.Message; $result.ExchangeOnPrem.TransportRules="Erro: ${msg}" }
        try { $result.ExchangeOnPrem.AcceptedDomains   = Get-AcceptedDomain  | Select DomainName, DomainType } catch { $msg=$_.Exception.Message; $result.ExchangeOnPrem.AcceptedDomains="Erro: ${msg}" }
        try { $result.ExchangeOnPrem.HybridConfiguration = Get-HybridConfiguration -ErrorAction Stop } catch { $msg=$_.Exception.Message; $result.ExchangeOnPrem.HybridConfiguration="Erro: ${msg}" }
    } catch {
        $msg = $_.Exception.Message
        $result.ExchangeOnPrem = "Erro geral na coleta on-prem: ${msg}"
    }
} else {
    $result.ExchangeOnPrem.Connectivity = "NotConnected"
    $result.Warnings += "Exchange On-Prem não conectado; coleta incompleta."
}

# --- DNS MX/SPF/DMARC/DKIM ---
foreach ($d in $Domains) {
    $domainObj = [ordered]@{ Domain=$d; MX=@(); MX_A=@(); TXT=@(); SPF=""; DMARC=""; DKIM=@() }
    try {
        $mxRecords = Resolve-DnsName -Name $d -Type MX -ErrorAction Stop
        foreach ($mx in $mxRecords) {
            $domainObj.MX += [ordered]@{ Preference=$mx.Preference; Exchange=$mx.NameExchange }
            try {
                $a = Resolve-DnsName -Name $mx.NameExchange -Type A -ErrorAction Stop
                foreach ($addr in $a) { $domainObj.MX_A += [ordered]@{ Exchange=$mx.NameExchange; IP=$addr.IPAddress } }
            } catch {
                $msg=$_.Exception.Message
                $domainObj.MX_A += [ordered]@{ Exchange=$mx.NameExchange; IP="Erro: ${msg}" }
            }
        }
    } catch {
        $msg=$_.Exception.Message
        $domainObj.MX = "Erro MX: ${msg}"
    }

    # TXT/SPF
    try {
        $txt = Resolve-DnsName -Name $d -Type TXT -ErrorAction Stop
        $txtStrings = $txt | ForEach-Object { ($_.Strings -join "") }
        $domainObj.TXT = $txtStrings
        if ($txtStrings -match "v=spf1") { $domainObj.SPF = ($txtStrings | Where-Object { $_ -match "v=spf1" })[0] } else { $domainObj.SPF="SPF não encontrado" }
    } catch {
        $msg=$_.Exception.Message
        $domainObj.TXT="Erro TXT: ${msg}"
    }

    # DMARC
    try {
        $dmarc = Resolve-DnsName -Name "_dmarc.$d" -Type TXT -ErrorAction Stop
        $domainObj.DMARC = ($dmarc | ForEach-Object { ($_.Strings -join "") }) -join "; "
    } catch { $domainObj.DMARC="DMARC não encontrado" }

    $result.DNS[$d] = $domainObj
}

# --- Teste Reachability TCP25 ---
$mxIps = @()
foreach ($d in $result.DNS.Keys) { foreach ($e in $result.DNS[$d].MX_A) { $mxIps += $e.IP } }
$mxIps = $mxIps | Where-Object { $_ -match "^\d{1,3}(\.\d{1,3}){3}$" } | Sort-Object -Unique
foreach ($ip in $mxIps) {
    try {
        $t = Test-NetConnection -ComputerName $ip -Port 25 -WarningAction SilentlyContinue
        $result.Reachability += [ordered]@{ IP=$ip; Tcp25Reachable=$t.TcpTestSucceeded; RemoteAddress=$t.RemoteAddress }
    } catch { $msg=$_.Exception.Message; $result.Reachability += [ordered]@{ IP=$ip; Tcp25Reachable="Erro: ${msg}" } }
}

# --- Geração HTML ---
function Convert-ToHtmlSection { param($Title, $Object) "<h2>$Title</h2>`n<pre>$($Object | Out-String)</pre>" }

$css = @"
body { font-family: Segoe UI, Arial; font-size: 12px; margin: 20px; }
h1 { font-size: 20px; }
h2 { font-size: 16px; color:#004b87; margin-top:18px; }
pre { background:#f6f6f6; padding:8px; border:1px solid #eee; }
"@

$htmlBody = "<html><head><meta charset='utf-8'><title>Mail Infra Audit Report</title><style>$css</style></head><body>"
$htmlBody += "<h1>Mail Infrastructure Audit Report</h1>"
$htmlBody += "<p>Generated: $($result.GeneratedOn) on host $($result.Host)</p>"

$htmlBody += "<h2>Warnings</h2><ul>"
foreach ($w in $result.Warnings) { $htmlBody += "<li>$w</li>" }
$htmlBody += "</ul>"

$htmlBody += Convert-ToHtmlSection -Title "Exchange Online" -Object $result.ExchangeOnline
$htmlBody += Convert-ToHtmlSection -Title "Exchange On-Prem" -Object $result.ExchangeOnPrem
$htmlBody += Convert-ToHtmlSection -Title "DNS Results" -Object $result.DNS
$htmlBody += Convert-ToHtmlSection -Title "Reachability" -Object $result.Reachability

$htmlBody += "<h2>Raw JSON</h2><pre>$($result | ConvertTo-Json -Depth 6)</pre>"
$htmlBody += "</body></html>"

$htmlBody | Out-File -FilePath $OutputHtml -Encoding UTF8 -Force
Write-Host "Relatório salvo em: $OutputHtml" -ForegroundColor Green

# Cleanup
if ($onPremSession) { Remove-PSSession $onPremSession -ErrorAction SilentlyContinue }
if ($connectedEXO) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
