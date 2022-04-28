class NPSNetworkPolicyAttribute {
    [string] $Name
    [string] $Id
    [string] $Value

    [string] ToString() {
        return "$($this.Name) [$($this.Id)] = $($this.Value)"
    }
}

class NPSNetworkPolicy {
    [string] $Name
    [string] $State
    [int] $ProcessingOrder
    [int] $PolicySource
    [NPSNetworkPolicyAttribute[]] $Conditions
    [NPSNetworkPolicyAttribute[]] $Profiles
}

Function Get-NPSNetworkPolicies {
    [CmdletBinding()]
    param ()

    PROCESS {
        $currentObject = $null
        $currentAttribute = $null

        switch -regex (netsh -c nps show np) {
            "Ok." {
                $currentObject
            }
            "Network Policy configuration:" {
                if ($null -ne $currentObject) {
                    $currentObject
                }

                $currentObject = [NPSNetworkPolicy] @{
                    Conditions = @()
                    Profiles = @()
                }
                continue
            }
            "Name\s+=\s+(.*)" {
                $currentObject.Name = ($matches.1).Trim()
                continue
            }
            "State\s+=\s+(.*)" {
                $currentObject.State = ($matches.1).Trim()
                continue
            }
            "Processing order\s+=\s+(.*)" {
                $currentObject.ProcessingOrder = [int] $matches.1
                continue
            }
            "Policy source\s+=\s+(.*)" {
                $currentObject.PolicySource = [int] $matches.1
                continue
            }
            '^(Condition\d+)\s+(0x[0-9a-f]+)\s+(.*)' {
                $name = ($matches.1).Trim()
                $id = ($matches.2).Trim()
                $value = ($matches.3).Trim()
                $currentObject.Conditions += [NPSNetworkPolicyAttribute] @{
                    Name = $Name
                    Id = $id
                    Value = $value
                }
                continue
            }
            '^([-a-zA-z]+)\s+(0x[0-9a-f]+)\s+(.*)' {
                $name = ($matches.1).Trim()
                $id = ($matches.2).Trim()
                $value = ($matches.3).Trim()
                $currentObject.Profiles += [NPSNetworkPolicyAttribute] @{
                    Name = $Name
                    Id = $id
                    Value = $value
                }
                continue
            }
        }
    }
}

Function Get-NPSNetworkPolicyProfileAttributes {
    [CmdletBinding()]
    param()
    PROCESS {
        switch -Regex (netsh -c nps show npprofileattributes)  {
            '^(.+?) {2,}(.+) {2,}(.+)' { 
                $name = ($Matches.1).Trim() 
                $id = ($Matches.2).Trim()
                $type = ($Matches.3).Trim()
                
                if ( $null -ne $name -and $name -ne "Name" ) {
                    $properties = @{
                        'Name'=$name.Trim();
                        'Id'=$id.Trim();
                        'Type'=$type.Trim()
                    }

                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj
                }
            }
        }
    }
}

Function Get-NPSNetworkPolicyConditionAttributes {
    [CmdletBinding()]
    param()
    PROCESS {
        switch -Regex (netsh -c nps show npconditionattributes)  {
            '^(.+?) {2,}(.+) {2,}(.+)' { 
                $name = $Matches.1 
                $id = $Matches.2
                $type = $Matches.3
                
                if ( $null -ne $name -and $name -ne "Name" ) {
                    $properties = @{
                        'Name'=$name.Trim();
                        'Id'=$id.Trim();
                        'Type'=$type.Trim()
                    }

                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj
                }
            }
        }
    }
}

#region help
<#
.SYNOPSIS
Add Network Policy config in NPS
.DESCRIPTION
Add Network Policy config in NPS

.PARAMETER Name
Name of the Network Policy to be added
#>
#endregion

Function Add-NPSNetworkPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                    HelpMessage="Name of Network Policy")]
        [string]$Name,

        [Parameter(Mandatory=$true,
                    HelpMessage="Processing Order")]
        [int]$ProcessingOrder,

        [Parameter(Mandatory=$false,
                    HelpMessage="Enabled")]
        [switch]$Enabled,

        [Parameter(Mandatory=$true,
                    HelpMessage="Conditions in format: condition-attribute-id=condition-data")]
        [String[]]$Conditions,

        [Parameter(Mandatory=$false,
                    HelpMessage="Profiles in format: profile-attribute-id=profile-data")]
        [String[]]$Profiles
    )

PROCESS {
    $State = "disable"
    if ( $Enabled ) {
        $State = "enable"
    }

    $Cmd = 'netsh -c nps add np name = "' + $Name + '" ' + 'state = "' + $State  + '"' + "processingorder = $ProcessingOrder"

    # Conditions
    for ($i = 0; $i -lt $Conditions.Count; $i++) {
        $Cond = $Conditions[$i].Split("=")
        $CondId = $Cond[0]
        $CondData = $Cond[1]


        if(! $CondId.StartsWith("0x")) {
            $CondName = $CondId
            $CondIdObj = Get-NetworkPolicyConditionAttributes | Where-Object {$_.Name -eq $CondName} | Select-Object Id
            if ( $null -eq $CondIdObj) {
                throw "Unsupported Condition Attribute Name '" + $CondName + "'"
            }
            $CondId = $CondIdObj.Id
        }
        $Cmd = $Cmd + ' conditionid = "' + $CondId + '" conditiondata = "' + $CondData + '"'
    }

    # Profiles
    for ($i = 0; $i -lt $Profiles.Count; $i++) {
        $Prof = $Profiles[$i].Split("=")
        $ProfId = $Prof[0]
        $ProfData = $Prof[1]

        if(! $ProfId.StartsWith("0x")) {
            $ProfName = $ProfId
            $ProfIdObj = Get-NetworkPolicyProfileAttributes | Where-Object {$_.Name -eq $ProfName} | Select-Object Id
            if ( $null -eq $ProfIdObj) {
                throw "Unsupported Profile Attribute Name '" + $ProfName + "'"
            }
            $ProfId = $ProfIdObj.Id
        }
        $Cmd = $Cmd + ' profileid = "' + $ProfId + '" profiledata = "' + $ProfData + '"'
    }

    Write-Debug "Command: $Cmd"
    Invoke-Expression -Command $Cmd | Out-String -OutVariable out | Out-Null
    if (!$out.StartsWith("Ok.")) {
        throw "Error adding the Network Policy"
    }
}
}