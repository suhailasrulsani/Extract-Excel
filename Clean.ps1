Clear-Host
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

Try
{
	Get-ChildItem -Path "$ScriptDir" -Recurse -Exclude "Extract.ps1", "Clean.ps1" -ErrorAction Stop | Remove-Item -Force -ErrorAction Stop
}

Catch
{
	Write-Warning ($_)
}
