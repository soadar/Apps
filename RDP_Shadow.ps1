Add-Type -AssemblyName PresentationCore, PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

####### Form #######
$form = New-Object System.Windows.Forms.Form
$form.Text = "Usuarios conectados"
$form.StartPosition = 'CenterScreen'
$form.AutoSize = $true
$form.AutoScaleMode = 'Font'
$form.AutoSizeMode = 'GrowOnly'
$form.FormBorderStyle = 'FixedSingle'
$form.Icon = New-Object system.drawing.icon (".\icon.ico")
####### Label #######
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(280, 20)
$label.Text = 'Seleccione ID del usuario:'
$label.Font = 'Microsoft Sans Serif,10'
$Currentlocation = (Get-Location).path
$icon = $Currentlocation + "icon.ico"
$form.Controls.Add($label)
####### ListView #######
$listView = New-Object System.Windows.Forms.ListView
$listView.View = 'Details'
$listView.Location = '10, 30'
$listView.Width = 300
$listView.Height = 300
$listView.FullRowSelect = $true
$listView.Font = 'Microsoft Sans Serif,10'
$listView.Add_DoubleClick({
		$id = $ListView.SelectedItems[0].Subitems[0].Text
		if ([Microsoft.VisualBasic.Information]::IsNumeric($id))
		{
			$form.Close()
			mstsc /shadow:$id /v:$srv /control
		}
		
	})
$columnA = New-Object System.Windows.Forms.ColumnHeader;
$columnA.Text = "ID"
$columnA.Width = 120
$columnA.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$columnB = New-Object System.Windows.Forms.ColumnHeader;
$columnB.Text = "Usuarios"
$columnB.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$listView.Columns.Add($columnA) | Out-Null;
$listView.Columns.Add($columnB) | Out-Null;

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$srv = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese hostname / IP:", "RDP session", "")
if ($srv)
{
	if (Test-Connection -ComputerName $srv -Count 1 -Quiet)
	{
		reg add "\\$srv\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v AllowRemoteRPC /t REG_DWORD /d 1 /F
		$winver = (Get-WmiObject Win32_OperatingSystem -ComputerName $srv).Caption
		if (-not $winver.Contains("Windows 10"))
		{
			[System.Windows.MessageBox]::Show('Sistema operativo no compatible', 'Alerta', 0, 48)
			$form.close()
		}
		$result = C:\Windows\System32\query.exe user /server:$srv | findstr "Activ"
		if ($result.Count -gt 0)
		{
			$rows = $result -split "`n"
			foreach ($row in $rows)
			{
				$row = -split $row
				$item = New-Object System.Windows.Forms.ListViewItem($row[2])
				$item.SubItems.Add($row[0])
				$listView.Items.AddRange(($item))
			}
			[void]$ListView.AutoResizeColumns(1)
			$form.Controls.Add($listView)
			$form.Topmost = $true
			$form.ShowDialog()
		}
		else
		{
			[System.Windows.MessageBox]::Show('No hay usuarios conectados', 'Alerta', 0, 48)
		}
	}
	else
	{
		[System.Windows.MessageBox]::Show('El servidor se encuentra apagado o no existe', 'Alerta', 0, 48)
	}
}
$form.Dispose()