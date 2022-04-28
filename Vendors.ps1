Function Get-NPSVendors {
    [CmdletBinding()]
    param ()

    PROCESS {
        $listStarted = $false
        switch -Regex (netsh -c nps show vendors)  {
            '^-' {
                $listStarted = $true
                continue
            }
            '^([^-]+.*)' { 
                if ($listStarted) {
                    $name = $Matches.1
                    
                    if ( $null -ne $name -and $name -ne "Ok." -and $name -ne "Name" ) {
                        $properties = @{
                            'Name'= $name.Trim()
                        }

                        $obj = New-Object -TypeName PSObject -Property $properties
                        Write-Output $obj
                    }
                }
            }
        }
    }
}