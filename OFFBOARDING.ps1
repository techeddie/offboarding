################################################
# 
# AUTHOR:  Eddie
# EMAIL:   eddie@directbox.de
# BLOG:    https://exchangeblogonline.de
# COMMENT: Migrate Online Mailbox to Exchange OP
#
################################################

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Bitte die Mailbox UPN eingeben")]
    [ValidateNotNullorEmpty()] [string] $userprincipalname,
    
    [Parameter(Mandatory = $true, HelpMessage = "Bitte die externe ExchangeFQDN eintragen")]
    [ValidateNotNullorEmpty()] [string] $ExchangeFQDN,

    [Parameter(Mandatory = $true, HelpMessage = "Bitte die Ziel-Datenbank eintragen")]
    [ValidateNotNullorEmpty()] [string] $TargetDatabase,

    [Parameter(Mandatory = $true, HelpMessage = "Bitte die Ziel-SMTP Domain eintragen")]
    [ValidateNotNullorEmpty()] [string] $TargetDomain 

)

$Host.ui.RawUI.WindowTitle = "EXCHANGE ONLINE OFFBOARDING"
$ErrorActionPreference = "SilentlyContinue"

Clear-Host
function header {
    $datum = Get-Date -Format ("HH:mm  dd/MM/yyyy")
    Write-Host "
 --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
 |e| |x| |c| |h| |a| |n| |g| |e| |b| |l| |o| |g| |o| |n| |l| |i| |n| |e| |.| |d| |e|
 --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
 Powered by Eddie  |  https://exchangeblogonline.de

" -F Green
    Write-Host "$datum                       `n" -b Blue
			
}
header


#remove ps-sessions
Get-PSSession | Remove-PSSession

#session params
$UserCredential = Get-Credential "Office365-Credentials"
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session -AllowClobber -ErrorAction SilentlyContinue

#proxy session
if ((Get-PSSession).computername -ne "outlook.office365.com") {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -SessionOption $proxysettings
    Start-Sleep 1
    Import-PSSession $Session -AllowClobber -ErrorAction SilentlyContinue
}

#function offboarding
function offboarding{ 

    $onpremcred = Get-Credential "DOMAIN\OnPremiseUser"   

    Get-MoveRequest $userprincipalname  -ErrorAction SilentlyContinue | Remove-MoveRequest -Confirm:$false
     
    new-moverequest -identity $userprincipalname -OutBound -RemoteTargetDatabase "$TargetDatabase" -RemoteHostName "$ExchangeFQDN" -RemoteCredential $onpremcred -TargetDeliveryDomain "$TargetDomain" -baditemlimit Unlimited

    Start-Sleep 3

    #status feedback
    do {
	
        $status = (Get-MoveRequest $userprincipalname | Get-MoveRequestStatistics).status
	
        if ($status -match "Failed") {
            Suspend-MoveRequest $userprincipalname -Confirm:$false
            Set-MoveRequest $userprincipalname -BadItemLimit unlimited -LargeItemLimit unlimited -AcceptLargeDataLoss 
            Start-Sleep 5
            Resume-MoveRequest $userprincipalname
        }	
	    
        Write-Host "Status..." -Fore Yellow
        Get-MoveRequestStatistics $userprincipalname | ft DisplayName, StatusDetail, TotalMailboxSize, TotalArchiveSize, PercentComplete
        start-sleep -s 3
        cls           
           
    }until ($status -match "Completed") 

    #remove moverequest if needed
    #Get-MoveRequest $userprincipalname | Remove-MoveRequest -Confirm:$false
    pause
}

#run offboarding
offboarding