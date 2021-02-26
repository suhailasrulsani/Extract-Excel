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

#region Copy PatchingList sheet from MPL to Temp_$Batch.xlsx
Write-Host "Copying PatchingList sheet from MPL to Temp_$Batch.xlsx : " -NoNewline
Start-Sleep 2
Try
{
	Copy-ExcelWorkSheet -SourceWorkbook "$ScriptDir\MasterPatchingList_v0.9.xlsx" -sourceWorksheet "PatchingList" -DestinationWorkbook "$ScriptDir\Temp_$Batch.xlsx" -DestinationWorksheet "Overwritten" -ErrorAction Stop
	Write-Host "Done" -ForegroundColor Green
}

Catch
{
	Write-Warning ($_)
	Continue
}

Finally
{
	$Error.Clear()
}

#endregion

Pause
