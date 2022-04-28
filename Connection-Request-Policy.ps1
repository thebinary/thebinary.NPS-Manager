class NPSConnectionRequestPolicyAttribute {
    [string] $Name
    [string] $Id
    [string] $Value

    [string] ToString() {
        return "$($this.Name) [$($this.Id)] = $($this.Value)"
    }
}

class NPSConnectionRequestPolicy {
    [string] $Name
    [string] $State
    [int] $ProcessingOrder
    [int] $PolicySource
    [NPSConnectionRequestPolicyAttribute[]] $Conditions
    [NPSConnectionRequestPolicyAttribute[]] $Profiles
}

Function Get-NPSConnectionRequestPolicies {
    [CmdletBinding()]
    param ()

    PROCESS {
        $currentObject = $null
        $currentAttribute = $null

        switch -regex (netsh -c nps show crp) {
            "Ok." {
                $currentObject
                continue
            }
            "Connection request policy configuration:" {
                if ($null -ne $currentObject) {
                    $currentObject
                }

                $currentObject = [NPSConnectionRequestPolicy] @{
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
                $currentObject.Conditions += [NPSConnectionRequestPolicyAttribute] @{
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
                $currentObject.Profiles += [NPSConnectionRequestPolicyAttribute] @{
                    Name = $Name
                    Id = $id
                    Value = $value
                }
                continue
            }
        }
    }
}

Function Get-NPSConnectionRequestPolicyProfileAttributes {
    [CmdletBinding()]
    param()
    PROCESS {
        switch -Regex (netsh -c nps show crpprofileattributes)  {
            '^(.+?) {2,}(.+) {2,}(.+)' { 
                $name = $Matches.1 
                $id = $Matches.2
                $type = $Matches.3
                
                if ( $null -ne $name -and $name -ne "Name" ) {
                    $properties = @{
                        'Name'="$name";
                        'Id'=$id;
                        'Type'=$type
                    }

                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj
                }
            }
        }
    }
}

Function Get-NPSConnectionRequestPolicyConditionAttributes {
    [CmdletBinding()]
    param()
    PROCESS {
        switch -Regex (netsh -c nps show crpconditionattributes)  {
            '^(.+?) {2,}(.+) {2,}(.+)' { 
                $name = $Matches.1 
                $id = $Matches.2
                $type = $Matches.3
                
                if ( $null -ne $name -and $name -ne "Name" ) {
                    $properties = @{
                        'Name'="$name";
                        'Id'=$id;
                        'Type'=$type
                    }

                    $obj = New-Object -TypeName PSObject -Property $properties
                    Write-Output $obj
                }
            }
        }
    }
}

Function Add-NPSConnectionRequestPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                    HelpMessage="Name of Connection Request Policy")]
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

    $Cmd = 'netsh -c nps add crp name = "' + $Name + '" ' + 'state = "' + $State  + '"' + "processingorder = $ProcessingOrder"

    # Conditions
    for ($i = 0; $i -lt $Conditions.Count; $i++) {
        $Cond = $Conditions[$i].Split("=")
        $CondId = $Cond[0]
        $CondData = $Cond[1]


        if(! $CondId.StartsWith("0x")) {
            $CondName = $CondId
            $CondIdObj = Get-ConnectionRequestPolicyConditionAttributes | Where-Object {$_.Name -eq $CondName} | Select-Object Id
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
            $ProfIdObj = Get-ConnectionRequestPolicyProfileAttributes | Where-Object {$_.Name -eq $ProfName} | Select-Object Id
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
        throw "Error adding the Connection Request Policy"
    }
}
}