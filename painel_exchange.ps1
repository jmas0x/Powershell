Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================
# Formulário principal
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Painel Exchange 2016 - Monitoramento"
$form.Size = New-Object System.Drawing.Size(1280,720)
$form.StartPosition = "CenterScreen"
$form.AutoScroll = $true
$form.AutoScrollMinSize = New-Object System.Drawing.Size(1400,900)

# =========================
# Lista de filas
# =========================
$queueList = New-Object System.Windows.Forms.ListBox
$queueList.Location = New-Object System.Drawing.Point(10,40)
$queueList.Size = New-Object System.Drawing.Size(300,500)
$form.Controls.Add($queueList)

# Botão para carregar filas
$btnLoadQueues = New-Object System.Windows.Forms.Button
$btnLoadQueues.Text = "Carregar Filas"
$btnLoadQueues.Location = New-Object System.Drawing.Point(10,10)
$btnLoadQueues.Size = New-Object System.Drawing.Size(300,25)
$form.Controls.Add($btnLoadQueues)

# =========================
# Grid de mensagens da fila
# =========================
$msgGrid = New-Object System.Windows.Forms.DataGridView
$msgGrid.Location = New-Object System.Drawing.Point(320,40)
$msgGrid.Size = New-Object System.Drawing.Size(900,500)
$msgGrid.ReadOnly = $true
$msgGrid.AllowUserToAddRows = $false
$msgGrid.AllowUserToDeleteRows = $false
$msgGrid.AutoSizeColumnsMode = "AllCells"
$msgGrid.ScrollBars = "Both"
$form.Controls.Add($msgGrid)

# =========================
# Grid de serviços Exchange
# =========================
$serviceGrid = New-Object System.Windows.Forms.DataGridView
$serviceGrid.Location = New-Object System.Drawing.Point(10,560)
$serviceGrid.Size = New-Object System.Drawing.Size(1210,120)
$serviceGrid.ReadOnly = $true
$serviceGrid.AllowUserToAddRows = $false
$serviceGrid.AllowUserToDeleteRows = $false
$serviceGrid.AutoSizeColumnsMode = "Fill"
$serviceGrid.ScrollBars = "Both"
$form.Controls.Add($serviceGrid)

# =========================
# Grid de erros do Exchange
# =========================
$errorGrid = New-Object System.Windows.Forms.DataGridView
$errorGrid.Location = New-Object System.Drawing.Point(10,690)
$errorGrid.Size = New-Object System.Drawing.Size(1210,150)
$errorGrid.ReadOnly = $true
$errorGrid.AllowUserToAddRows = $false
$errorGrid.AllowUserToDeleteRows = $false
$errorGrid.AutoSizeColumnsMode = "Fill"
$errorGrid.ScrollBars = "Both"
$form.Controls.Add($errorGrid)

# =========================
# Funções auxiliares
# =========================

function Load-ServiceStatus {
    $exchangeServices = "MSExchangeIS","MSExchangeTransport","MSExchangeFrontEndTransport","MSExchangeMailboxAssistants"
    $services = Get-Service | Where-Object { $_.Name -in $exchangeServices } |
        Select-Object DisplayName,Status,StartType

    $dataTable = New-Object System.Data.DataTable
    "DisplayName","Status","StartType" | ForEach-Object { $null = $dataTable.Columns.Add($_) }

    foreach ($svc in $services) {
        $row = $dataTable.NewRow()
        $row["DisplayName"] = $svc.DisplayName
        $row["Status"]      = $svc.Status
        $row["StartType"]   = $svc.StartType
        $dataTable.Rows.Add($row)
    }
    $serviceGrid.DataSource = $dataTable
}

function Load-ExchangeErrors {
    $events = Get-WinEvent -LogName "Application" -MaxEvents 50 |
        Where-Object { $_.ProviderName -like "MSExchange*" -and $_.LevelDisplayName -in "Error","Warning" } |
        Select-Object TimeCreated,ProviderName,LevelDisplayName,Message

    $dataTable = New-Object System.Data.DataTable
    "TimeCreated","ProviderName","Level","Message" | ForEach-Object { $null = $dataTable.Columns.Add($_) }

    foreach ($evt in $events) {
        $row = $dataTable.NewRow()
        $row["TimeCreated"] = $evt.TimeCreated
        $row["ProviderName"] = $evt.ProviderName
        $row["Level"]        = $evt.LevelDisplayName
        $row["Message"]      = $evt.Message.Substring(0, [Math]::Min(200,$evt.Message.Length))
        $dataTable.Rows.Add($row)
    }

    $errorGrid.DataSource = $dataTable
}

# =========================
# Eventos
# =========================
$btnLoadQueues.Add_Click({
    $queueList.Items.Clear()
    try {
        $queues = Get-Queue | Select-Object Identity,MessageCount
        $queues | ForEach-Object { 
            $queueList.Items.Add("$($_.Identity)  [$($_.MessageCount) mensagens]")
        }
        Load-ServiceStatus
        Load-ExchangeErrors
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro ao obter filas: $_", "Erro")
    }
})

$queueList.Add_SelectedIndexChanged({
    $selected = $queueList.SelectedItem
    if ($selected) {
        $queueName = $selected -replace "\s+\[\d+\s+mensagens\]",""
        try {
            $messages = Get-Message -Queue $queueName | Select-Object FromAddress,Recipients,Subject,Status,Size

            if ($messages -and $messages.Count -gt 0) {
                $dataTable = New-Object System.Data.DataTable
                "FromAddress","Recipients","Subject","Status","Size" | ForEach-Object {
                    $null = $dataTable.Columns.Add($_)
                }

                foreach ($msg in $messages) {
                    $row = $dataTable.NewRow()
                    $row["FromAddress"] = $msg.FromAddress
                    $row["Recipients"]  = ($msg.Recipients -join "; ")
                    $row["Subject"]     = $msg.Subject
                    $row["Status"]      = $msg.Status
                    $row["Size"]        = $msg.Size
                    $dataTable.Rows.Add($row)
                }

                $msgGrid.DataSource = $dataTable
            } else {
                [System.Windows.Forms.MessageBox]::Show("Nenhuma mensagem na fila $queueName", "Fila Vazia")
                $msgGrid.DataSource = $null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao acessar fila: $_", "Erro",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# =========================
# Exibir formulário
# =========================
$form.ShowDialog() | Out-Null
