Clear-Host
Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
Remove-Variable * -ErrorAction SilentlyContinue; $Error.Clear();
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$datetime = Get-Date -Format G
$dt = (Get-Date).ToString("ddMMyyyy_HHmmss")
Try { Get-Module -ListAvailable slpslib, importexcel -ErrorAction Stop | Import-Module -WarningAction SilentlyContinue }
Catch { Write-Warning ($_); Continue }

#region Variable
$Batch = "Batch1"
$Thu = "Thursday"
$Sat = "Saturday"
$Sun = "Sunday"
$Mpl = "C:\Users\Muhammad Suhail\Desktop\MasterPatchingList_v0.9.xlsx"
$folder = @("$ScriptDir\$Thu", "$ScriptDir\$Sat", "$ScriptDir\$Sun")
#endregion

#region Create folder Thursday, Saturday, Sunday
Write-Host "Creating folder $Thu, $Sat and $Sun : " -NoNewline
Try
{
	New-Item -ItemType Directory -Path "$ScriptDir\$Thu" -Force -ErrorAction Stop | Out-Null
	New-Item -ItemType Directory -Path "$ScriptDir\$Sat" -Force -ErrorAction Stop | Out-Null
	New-Item -ItemType Directory -Path "$ScriptDir\$Sun" -Force -ErrorAction Stop | Out-Null
	Write-Host "Done" -ForegroundColor Green
}

Catch
{
	Write-Warning ($_); Continue
}
#endregion

#region Copy MPL to current directory
Write-Host "Coying MPL to current directory : " -NoNewline
Try
{
	Copy-Item -Path $Mpl -Destination "$ScriptDir\" -Force -ErrorAction Stop | Out-Null
	Write-Host "Done" -ForegroundColor Green
}

Catch
{
	Write-Warning ($_); Continue
}
#endregion Copy MPL to current directory

#region Creating Temp_$Batch.xlsx
Write-Host "Creating Temp_$Batch.xlsx : " -NoNewline
Try
{
	New-SLDocument -WorkbookName Temp_$Batch -WorksheetName Overwritten -Path "$ScriptDir" -Force -ErrorAction Stop
	Write-Host "Done" -ForegroundColor Green
}

Catch
{
	Write-Warning ($_); Continue
}
#endregion

#region Convert MPL to CSV
$oleDbConn = New-Object System.Data.OleDb.OleDbConnection
$oleDbCmd = New-Object System.Data.OleDb.OleDbCommand
$oleDbAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
$dataTable = New-Object System.Data.DataTable
$oleDbConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='$ScriptDir\MasterPatchingList_v0.9.xlsx';Extended Properties=Excel 16.0;Persist Security Info=False"
$oleDbConn.Open()
$oleDbCmd.Connection = $OleDbConn
$oleDbCmd.commandtext = "Select * from [Sheet1$]"
$oleDbAdapter.SelectCommand = $OleDbCmd
$ret = $oleDbAdapter.Fill($dataTable)
Write-Host	"Rows returned:$ret" -ForegroundColor green
$dataTable | Export-Csv "$ScriptDir\Temp.csv" -Delimiter ';'
$oleDbConn.Close()
#endregion
