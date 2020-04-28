#Requires -Modules MicrosoftTeams, MSOnline

#This script requires a CSV field with the following fields:
#TeamName - The exact case matched name of the team, if this does not exist it will be created
#UserPrincipalName - the Userprincipal name of the user which is their kent login followed by @kent.ac.uk eg ja80@kent.ac.uk
#MemberType - This should be of either of the following values: Member,Owner
<#
.SYNOPSIS
Allows the creation and/or population of teams from CSV data

.DESCRIPTION
Takes input from CSV data to either:
Create a team with an owner(required), populate additional owners(optional) and populate members(optional)
or
Populate a team with additional owners and/or members

.PARAMETER InputPath
The path to the CSV data file
This script requires a CSV field with the following fields:
TeamName - The exact case matched name of the team, if this does not exist it will be created
UserPrincipalName - the Userprincipal name of the user which is their kent login followed by @kent.ac.uk eg ja80@kent.ac.uk
MemberType - This should be of either of the following values: Member,Owner

.PARAMETER Connect
When used the connect switch will tell the script to connect to the appropriate 365 powershell services

.PARAMETER Log
When used the Log switch will enable logging for the script to a time stamped log file in the same directory as the script of the name (CreateAndPopulateTeam-<timestamp>.log)

.EXAMPLE
The following would take the data from Data.csv, connect to 365 services and also enabling logging for this run
CreateAndPopulateTeam.ps1 -InputPath Data.csv -Connect -Log

.NOTES
This script required 2 powershell modules These can be installed with the following commands
Install-Module -Module MicrosoftTeams -Force
Install-Module -Module MSOnline -Force

The scrip will import these modules for the user

#>
[CmdletBinding(PositionalBinding = $false)]
param (
    [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default', Mandatory = $true)]
    [String] $InputPath,
    [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default', Mandatory = $false)]
    [switch] $Connect,
    [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default', Mandatory = $false)]
    [switch] $Log
)

$timestamp = get-date -Format "yyyyMMdd-HHmmss"
$logLocation = "$PSScriptRoot\CreateAndPopulateTeam-$timestamp.log"

function log{
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default', Mandatory = $true)]
        [String] $Message
    )
    if ($log){
        $MessageTimestamp = get-date -Format "yyyyMMdd-HHmmss"
        Add-Content $logLocation ("[$MessageTimestamp]$Message")
    }
}
log -Message "[INFO]Script Started"
Import-Module MSOnline -ErrorAction Stop
Import-Module MicrosoftTeams -ErrorAction Stop
log -Message "[INFO]Modules Loaded"

if ($Connect){
    log -Message "[INFO]Connecting to 365 Powershell Services"
    Connect-MsolService
    Connect-MicrosoftTeams
}

try
{
    Get-MsolDomain -ErrorAction Stop > $null
}
catch
{
    log -Message "[Error]Please run Connect-MsolService before running this script or run with the -Connect switch"
    Write-Error "Please run Connect-MsolService before running this script or run with the -Connect switch"
    exit 1
}

function Verify-User{
    [CmdletBinding(PositionalBinding = $false)]
    param (
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default', Mandatory = $true)]
        [String] $UserPrincipalName
    )

    $user = Get-MsolUser -UserPrincipalName $UserPrincipalName -ErrorAction SilentlyContinue
    if ($null -eq $user){
        log -Message"[WARNING]$user does not exist in 365"
        $false
    }else{
        log -Message "[INFO]$user does exist in 365"
        $true
    }
}

# Test Input File exists
if (-not (Test-Path $InputPath -PathType Leaf)){
    log -Message"[ERROR]Input file at $InputPath does not exist"
    Write-Error "Input file at $InputPath does not exist"
    exit 1
}

#Load File
$inputData = Import-Csv $InputPath

#Test File contents
if ($null -eq $inputData -or $inputData.count -eq 0){
    log -Message "[ERROR]Input file at $InputPath does not contain any csv data"
    Write-Error "Input file at $InputPath does not contain any csv data"
    exit 1
}

#Verify fields for all users
$errorCount = 0
$emailRegex = '^\S*@kent.ac.uk$'
$teamNameRegex = '^(\d|\w|\-| )*$'
foreach ($record in $inputData){
    $errorFound = $false
    if ($null -eq $record.TeamName -or $record.TeamName -eq ""){
        $errorFound = $true
    }elseif($record.TeamName -notmatch $teamNameRegex){
        $errorFound = $true
        log -Message ("[WARNING]"+ $record.TeamName +" does not match the required restriction of letters number or dash characters")
        Write-Warning ($record.TeamName + " does not match the required restriction of letters number or dash characters")
    }
    if ($null -eq $record.UserPrincipalName -or $record.UserPrincipalName -eq ""){
        $errorFound = $true
    }elseif($record.UserPrincipalName -notmatch $emailRegex){
        $errorFound = $true
        log -Message ("[WARNING]"+ $record.UserPrincipalName +" does not match the required format of <login>@kent.ac.uk")
        Write-Warning ($record.UserPrincipalName + " does not match the required format of <login>@kent.ac.uk")
    }
    if ($null -eq $record.MemberType -or $record.MemberType -eq ""){
        $errorFound = $true
    }
    if ($errorFound -gt 0){
        log -Message ("[WARNING]Error found with record missing a required field. TeamName: " + $record.TeamName + " UserPrincipalName: " + $record.UserPrincipalName + " MemberType: " + $record.MemberType)
        Write-Warning ("Error found with record missing a required field. TeamName: " + $record.TeamName + " UserPrincipalName: " + $record.UserPrincipalName + " MemberType: " + $record.MemberType)
        $errorCount ++
    }
}

if ($errorCount -gt 0){
    log -Message "[ERROR]Errors found in Input file data. $errorCount errors found. Script will not proceed while there are errors in the data"
    Write-Error "Errors found in Input file data. $errorCount errors found. Script will not proceed while there are errors in the data"
    exit 1
}

#Build dataobject for creation adn verify users exist in AzureAD

$dataSet = foreach ($teamName in ($inputData.TeamName | Select-Object -Unique)){
    $teamData = $inputData | Where-Object {$_.TeamName -eq $teamName}
    $owners = @()
    ($teamData | Where-Object {$_.MemberType -eq "Owner"}) | foreach-object {
        if (Verify-User -UserPrincipalName $_.UserPrincipalName){
            $owners = $owners + $_.UserPrincipalName
        }else{
            log -Message ("[Warning]"+ $_.UserPrincipalName +" could not be verified in Azure. User will not be added for $teamName")
            Write-Warning ("User "+ $_.UserPrincipalName +" could not be verified in Azure. User will not be added for $teamName")
        }
    }

    $members = @()
    ($teamData | Where-Object {$_.MemberType -eq "Member"}) | foreach-object {
        if (Verify-User -UserPrincipalName $_.UserPrincipalName){
            $members = $members + $_.UserPrincipalName
        }else{
            log -Message ("[WARNING]User "+ $_.UserPrincipalName +" could not be verified in Azure. User will not be added for $teamName")
            Write-Warning ("User "+ $_.UserPrincipalName +" could not be verified in Azure. User will not be added for $teamName")
        }
    }

    if ($owners.count -lt 1-or $null -eq $owners){
        log -Message "No owners found for team: $teamName"
        Write-Warning "No owners found for team: $teamName"
    }
    [PSCustomObject]@{
        TeamName = $teamName.trim()
        Owners = $owners
        Members = $members
    }
}

#Check if teams exist
foreach ($teamData in $dataSet){
    $existingTeam = Get-Team -DisplayName $teamData.TeamName
    if ($null -eq $existingTeam){
        log -Message ("[INF0]User "+ $teamData.TeamName + " does not exist.")
        Write-Verbose -Verbose ($teamData.TeamName + " does not exist.")

        #Create team if 1 or more owners
        if ($teamData.Owners.count -gt 0){
            #Create Team
            $teamMailNickName = $teamData.TeamName -replace " ", "-"
            $newTeam = New-Team -DisplayName $teamData.TeamName -MailNickName $teamMailNickName -Description $teamData.TeamName -Visibility "Private" -Owner $teamData.Owners[0]  -whatif

            if ($null -eq $newTeam){
                log -Message ("[ERROR]Failed to create team " + $teamData.TeamName + ". Please consult with Information Services")
                write-error ("Failed to create team " + $teamData.TeamName + ". Please consult with Information Services")
            }else{
                log -Message ("[INFO]Team created with name " + $teamData.TeamName + " and ID: " + $newTeam.GroupID)
                Write-Verbose -Verbose ("Team created with name " + $teamData.TeamName + " and ID: " + $newTeam.GroupID)

                $teamUsers = Get-TeamUser -GroupId  $newTeam.GroupID
                #add other owners
                $ownersToAdd = $teamData.Owners | where-object {$_ -notin $teamData.Owners[0] -and $_ -notin $teamUsers.User}
                if ($ownersToAdd.count -gt 0){
                    $ownersToAdd| Foreach-object{
                        log -Message ("[INFO]Adding Owner $_ to " + $teamData.TeamName)
                        Write-Host ("Adding Owner $_ to " + $teamData.TeamName)
                        Add-TeamUser -GroupID $newTeam.GroupID -Role "Owner" -User $_
                    }
                }

                #add members
                $membersToAdd = $teamData.Members | where-object {$_ -notin $teamUsers.User}
                if ($membersToAdd.count -gt 0){
                    $membersToAdd| Foreach-object{
                        log -Message ("[INFO]Adding Member $_ to " + $teamData.TeamName)
                        Write-Host ("Adding Member $_ to " + $teamData.TeamName)
                        Add-TeamUser -GroupID $newTeam.GroupID -Role "Member" -User $_
                    }
                }
            }
        }else{
            Write-Error ($teamData.TeamName + " cannot be created. Needs at least 1 Owner. Skiping team.")
        }
    }else{
        #check owners and members then adds the missing
        $teamUsers = Get-TeamUser -GroupId  $existingTeam.GroupID

        #add other owners
        $ownersToAdd = $teamData.Owners | where-object {$_ -notin $teamData.Owners[0] -and $_ -notin $teamUsers.User}
        if ($ownersToAdd.count -gt 0){
            $ownersToAdd| Foreach-object{
                log -Message ("[INFO]Adding Owner $_ to " + $teamData.TeamName)
                Write-Host ("Adding Owner $_ to " + $teamData.TeamName)
                Add-TeamUser -GroupID $existingTeam.GroupID -Role "Owner" -User $_
            }
        }

        #add members
        $membersToAdd = $teamData.Members | where-object {$_ -notin $teamUsers.User}
        if ($membersToAdd.count -gt 0){
            $membersToAdd| Foreach-object{
                log -Message ("[INFO]Adding Member $_ to " + $teamData.TeamName)
                Write-Host ("Adding Member $_ to " + $teamData.TeamName)
                Add-TeamUser -GroupID $existingTeam.GroupID -Role "Member" -User $_
            }
        }
    }
}

log -Message "[INFO]Script Finished"