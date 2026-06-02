<# Name : Reboot Multiple Servers

   Descritption : Script to reboot multiple servers remotely
   
   Author : Abhishek Bharath

   NOTE: Change Path of input file as required (VERY IMPORTANT!!)

#>

$Servers=Get-content "C:\temp\reboot\servers.txt" #Input File
#$Servers="server_name"

foreach ($Server in $Servers)
{
    Write-Host $server -ForegroundColor Yellow
    Invoke-Command -ComputerName $server -ScriptBlock {shutdown /r /t 0}
}