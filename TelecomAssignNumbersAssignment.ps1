<#	
    .NOTES
    ===========================================================================
    Created with: 	ISE
    Created on:   	8/02/2018 1:46 PM
    Created by:   	Vikas Sukhija
    Organization: 	
    Filename:     	TelecomAssignNumbers.ps1
    ===========================================================================
    .DESCRIPTION
    This script is used for Telcom Number Migrations (NumberAssignment)
#>
param (
  [string]$CSV = $(Read-Host "Enter the CSV path to process"),
  [string]$domain = $(Read-Host "Enter the onmicrosoft domain")
)
function Write-Log
{
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true,ParameterSetName = 'Create')]
    [array]$Name,
    [Parameter(Mandatory = $true,ParameterSetName = 'Create')]
    [string]$Ext,
    [Parameter(Mandatory = $true,ParameterSetName = 'Create')]
    [string]$folder,
    
    [Parameter(ParameterSetName = 'Create',Position = 0)][switch]$Create,
    
    [Parameter(Mandatory = $true,ParameterSetName = 'Message')]
    [String]$Message,
    [Parameter(Mandatory = $true,ParameterSetName = 'Message')]
    [String]$path,
    [Parameter(Mandatory = $false,ParameterSetName = 'Message')]
    [ValidateSet('Information','Warning','Error')]
    [string]$Severity = 'Information',
    
    [Parameter(ParameterSetName = 'Message',Position = 0)][Switch]$MSG
  )
  switch ($PsCmdlet.ParameterSetName) {
    "Create"
    {
      $log = @()
      $date1 = Get-Date -Format d
      $date1 = $date1.ToString().Replace("/", "-")
      $time = Get-Date -Format t
	
      $time = $time.ToString().Replace(":", "-")
      $time = $time.ToString().Replace(" ", "")
	
      foreach ($n in $Name)
      {$log += (Get-Location).Path + "\" + $folder + "\" + $n + "_" + $date1 + "_" + $time + "_.$Ext"}
      return $log
    }
    "Message"
    {
      $date = Get-Date
      $concatmessage = "|$date" + "|   |" + $Message +"|  |" + "$Severity|"
      switch($Severity){
        "Information"{Write-Host -Object $concatmessage -ForegroundColor Green}
        "Warning"{Write-Host -Object $concatmessage -ForegroundColor Yellow}
        "Error"{Write-Host -Object $concatmessage -ForegroundColor Red}
      }
      
      Add-Content -Path $path -Value $concatmessage
    }
  }
} #Function Write-Log
function Start-ProgressBar
{
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)]
    $Title,
    [Parameter(Mandatory = $true)]
    [int]$Timer
  )
	
  For ($i = 1; $i -le $Timer; $i++)
  {
    Start-Sleep -Seconds 1;
    Write-Progress -Activity $Title -Status "$i" -PercentComplete ($i /100 * 100)
  }
}

function LaunchSOL
{
  param
  (
    [Parameter(Mandatory = $true)]
    $Domain,
    [Parameter(Mandatory = $false)]
    $Credential
  )
  Write-Host -Object "Enter Skype Online Credentials" -ForegroundColor Green
  $dommicrosoft = $domain + ".onmicrosoft.com"
  $CSSession = New-CsOnlineSession -Credential $Credential -OverrideAdminDomain $dommicrosoft 
  Import-Module (Import-PSSession -Session $CSSession -AllowClobber) -Prefix SOL  -Global
} #Function LaunchSOL

Function RemoveSOL
{
  $Session = Get-PSSession | Where-Object -FilterScript { $_.ComputerName -like "*.online.lync.com" }
  Remove-PSSession $Session
} #Function RemoveSOL
#################Check if logs folder is created####
$logpath  = (Get-Location).path + "\logs" 
$testlogpath = Test-Path -Path $logpath
if($testlogpath -eq $false)
{
 Start-ProgressBar -Title "Creating logs folder" -Timer 10
  New-Item -Path (Get-Location).path -Name Logs -Type directory
}

$Reportpath  = (Get-Location).path + "\Report" 
$testlogpath = Test-Path -Path $Reportpath
if($testlogpath -eq $false)
{
  Start-ProgressBar -Title "Creating Report folder" -Timer 10
  New-Item -Path (Get-Location).path -Name Report -Type directory
}
####################Load variables and log##########
$log = Write-Log -Name "TelecomNumberMigration-Log" -folder "logs" -Ext "log"
$Report = Write-Log -Name "TelecomNumberMigration-Report" -folder "Report" -Ext "csv"

#############Start script############################
Write-Log -Message "Start script" -path $log
try 
{
  Write-Log -Message "Load user info" -path $log
  Write-Log -Message "Load Modules" -path $log
  LaunchSOL -Domain $domain
  Write-Log -Message "loaded.... SKOB Online Module" -path $log
}
catch 
{
  $exception = $_.Exception
  Write-Log -Message "Error loading Module" -path $log -Severity Error 
  Write-Log -Message $exception -path $log -Severity error
  exit;
}

#############fetch CSV data############################
Write-Log -Message "Start importing $CSV" -path $log
try
{
  $coll = @()
  $collection = @()
  $data = Import-Csv $CSV
  ForEach($i in $data)
  {
    $mcoll = "" | select UPN, PhNumber
    $upn = $i.upn.trim()
    $PhNumber = $i.PhoneNumber.trim()
    $mcoll.UPN = $upn 
    $mcoll.PhNumber = $PhNumber
    $coll += $mcoll
  }
  $collection = $coll | where{$_.upn -ne ""}
}

catch
{
  $exception = $_.Exception
  Write-Log -Message "Error fetching Information from CSV" -path $log -Severity Error 
  Write-Log -Message $exception -path $log -Severity error
  exit;
}
#####################Process the collection###############
$error.clear()
$coll1 = @()
if($collection)
{
  Write-Log -Message "Processing cleaned collection" -path $log
  foreach($x in $collection)
  {
    $mcoll = "" | select UPN, PhNumber, Status
    $userprincipalname = $x.UPN
    $phone = $x.Phnumber
    $mcoll.UPN = $userprincipalname
    $mcoll.PhNumber = $phone
    try
    {
     Set-SOLCsOnlineVoiceUser -Identity $userprincipalname -TelephoneNumber $phone
      Write-Log -Message "Assign $userprincipalname with phone $phone" -path $log
      if($error)
      {
        $mcoll.Status = "Error"
        Write-Log -Message "Error occured processing $userprincipalname" -path $log -Severity Error 
        Write-Log -Message "$error" -path $log -Severity Error
        $error.clear()
      }
      else{$mcoll.Status = "Success"}
    }
    catch
    {
      $mcoll.Status = "Exception"
      $exception = $_.Exception
      Write-Log -Message "Exception occured processing $userprincipalname" -path $log -Severity Error 
      Write-Log -Message $exception -path $log -Severity error
    }

    $coll1 += $mcoll
  }
}
$coll1 | Export-Csv $Report -NoTypeInformation
########################Recycle reports & logs##############
Write-Log -Message "Script Finished" -path $log