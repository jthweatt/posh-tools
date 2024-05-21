#Requires -RunAsAdministrator

<#
.SYNOPSIS
   Delete user profiles on local or remote computer. Utilizes CIM. The session in which you are running the script must be started with elevated user rights (Run as Administrator).
.DESCRIPTION
   This script delete the user profiles on local or remote computer that match the search criteria.
.PARAMETER UserName
   User Name to delete user profile, is possible use the '*' wildchar.
.PARAMETER ExcludeUserName
   User name to exclude, is possible use the '*' wildchar.
.PARAMETER InactiveDays
   Inactive days of the profile, this parameter is optional and specify that the profile will be deleted only if not used for the specifed days.
.PARAMETER ComputerName
   Host name or list of host names on witch delete user profile, this parameter is optional (the default value is local computer).
.PARAMETER Special
   Searches special system service users, this parameter is optional (the default value is False).
.PARAMETER Force
   Force execution without require confirm (the default value is False).
.EXAMPLE
   ./Remove-UserProfileQuick.ps1 -UserName "JasonT"
   Delete the profile of the user with user name equal JasonT.
.EXAMPLE
   ./Remove-UserProfileQuick.ps1 -UserName "Jason*"
   Delete all user profiles of the user with user name begin with "Jason".
.EXAMPLE
   ./Remove-UserProfileQuick.ps1 -UserName "*" -InactiveDays 30
   Delete all user profiles inactive by 30 days.
.EXAMPLE
   ./Remove-UserProfileQuick.ps1 -UserName "*" -ExcludeUserName Admistrator
   Delete all user profiles exclude user name Administrator.
.EXAMPLE
   ./Remove-UserProfileQuick.ps1 -UserName "ServiceUser" -Special
   Delete the profile of the user with user name equal ServiceUser. Searches only profiles with Special flag set to true.
.EXAMPLE
   ./Remove-UserProfileQuick.ps1 -UserName "*" -Force
   Delete all user profiles without require confim.
.NOTES
   Author:  Jason Thweatt
   Blog:    https://github.com/jthweatt
   Date:    05/20/2024
   Version: 1.0.0
.LINK  
#>

<#
Changelog
[1.0.0] - 2024-05-20 - Initial Build.
#>

[cmdletbinding(ConfirmImpact = 'High', SupportsShouldProcess=$True)]
Param(
  [Parameter(Mandatory=$True)]
  [string]$Name,
  [string]$ExcludeName = [string]::Empty,
  [string[]]$ComputerName = $env:computername,
  [uint32]$InactiveDays = [uint32]::MaxValue,
  [switch]$Special = $False,
  [switch]$Force = $False
)

Set-strictmode -version latest

ForEach ($computer in $ComputerName){

  $profilesFound = 0

  # Get profiles from computer
  Try {
    $profiles = Get-CimInstance -Class Win32_UserProfile -Computer $computer -Filter "Special = '$Special'"
  } Catch {            
    Write-Warning "Failed to retreive user profiles on $ComputerName. $_"
    Exit
  }
}

ForEach ($profile in $profiles){
  # Set variables
  $profilePath = $profile.LocalPath
  $profileName = $profile.LocalPath.split('\')[-1]
  $sid = New-Object System.Security.Principal.SecurityIdentifier($profile.SID)
  $loaded = $profile.Loaded
  $lastUseTime = $profile.LastUseTime
  $special = $profile.Special

  # Calculation of the unused days of the profile
  $profileUnusedDays=0
  If (-Not $loaded){
    $profileUnusedDays = (New-TimeSpan -Start $lastUseTime -End (Get-Date)).Days 
  }

  If($profileName.ToLower() -Eq $UserName.ToLower() -Or ($UserName.Contains("*") -And $profileName.ToLower() -Like $UserName.ToLower())) {
    If($ExcludeUserName -ne [string]::Empty -And -Not $ExcludeUserName.Contains("*") -And ($profileName.ToLower() -eq $ExcludeUserName.ToLower())){
      Continue
    }
    If($ExcludeUserName -ne [string]::Empty -And $ExcludeUserName.Contains("*") -And ($profileName.ToLower() -Like $ExcludeUserName.ToLower())){
      Continue
    }
    If($InactiveDays -ne [uint32]::MaxValue -And $profileUnusedDays -le $InactiveDays){
      continue
    }

    $profilesFound ++

    If ($profilesFound -gt 1) {Write-Host "`n"}
    Write-Host "Start deleting profile ""$profileName"" on computer ""$computer"" ..." -ForegroundColor Green
    Write-Host "Account SID: $sid"
    Write-Host "Special system service user: $special"
    Write-Host "Profile Path: $profilePath"
    Write-Host "Loaded : $loaded"
    Write-Host "Last use time: $lastUseTime"
    Write-Host "Profile unused days: $profileUnusedDays"

    If ($loaded) {
      Write-Warning "Cannot delete profile because is in use"
      Continue
    }

    If ($Force -Or $PSCmdlet.ShouldProcess($profileName)) {
      Try {
        $profile | Remove-CimInstance           
        Write-Host "Profile deleted successfully" -ForegroundColor Green        
      } 
      Catch {            
        Write-Host "Error during delete the profile. $_" -ForegroundColor Red
      }
    } 
  }
}
 
If($profilesFound -eq 0){
  Write-Warning "No profiles found on $ComputerName with Name $UserName"
}