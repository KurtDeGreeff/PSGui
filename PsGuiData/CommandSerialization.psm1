<#
.SYNOPSIS
    Outputs a lot of information for a command
.DESCRIPTION
    Outputs
        - name
        - modulename
        - modules
        - commandtype
        - definition
        - parametersets

    Essentially it's a slightly modified version of the one contained in the native powershell "Show-Command" command.
.EXAMPLE
    PS> Get-SerializedCommandInfo

    Outputs the command infos about every loaded command
.EXAMPLE
    PS> Get-SerializedCommandInfo -Command Get-Help -NoWellKnownParameters

    Outputs the info about the command Get-Help, without the parametersets containing
    well known parameters, like -ErrorAction, -Confirm, etc.
#>
Function Get-SerializedCommandInfo {

    [CmdletBinding()]
    [OutputType([PSCustomObject])]

    Param (
        # The names of the commands
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNull()]
        [String[]] $Commands,

         # If set well known parameters will not be output
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNull()]
        [Switch] $NoWellKnownParameters
    )

    Begin {
        Write-Debug "Starting $($MyInvocation.Mycommand)."

        Function GetParameterType {
            param ([Type] $parameterType)
            
            $returnParameterType = New-Object PSObject
            $returnParameterType | Add-Member -MemberType NoteProperty -Name "FullName" -Value $parameterType.FullName
            $returnParameterType | Add-Member -MemberType NoteProperty -Name "IsEnum" -Value $parameterType.IsEnum
            $returnParameterType | Add-Member -MemberType NoteProperty -Name "IsArray" -Value $parameterType.IsArray
            
            if ($parameterType.IsEnum) {
                $enumValues = [System.Enum]::GetValues($parameterType)
            } else {
                $enumValues = [string[]] @()
            }

            $returnParameterType | Add-Member -MemberType NoteProperty -Name "EnumValues" -Value $enumValues
            if ($parameterType.IsArray) {
                $hasFlagAttribute = ($parameterType.GetCustomAttributes([System.FlagsAttribute], $true).Length -gt 0)
                # Recurse into array elements.
                $elementType = GetParameterType($parameterType.GetElementType())
            } else {
                $hasFlagAttribute = $false
                $elementType = $null
            }
            $returnParameterType | Add-Member -MemberType NoteProperty -Name "HasFlagAttribute" -Value $hasFlagAttribute
            $returnParameterType | Add-Member -MemberType NoteProperty -Name "ElementType" -Value $elementType
            
            if (!($parameterType.IsEnum) -and !($parameterType.IsArray)) {
                $implementsDictionary = [System.Collections.IDictionary].IsAssignableFrom($parameterType)
            } else {
                $implementsDictionary = $false
            }
            $returnParameterType | Add-Member -MemberType NoteProperty -Name "ImplementsDictionary" -Value $implementsDictionary
            return $returnParameterType
        }

        Function GetParameterInfo {
            param ($parameters)

            $parameterInfos = @()
            foreach ($parameter in $parameters) {
                $parameterInfo = new-object PSObject
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $parameter.Name
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "IsMandatory" -Value $parameter.IsMandatory
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "ValueFromPipeline" -Value $parameter.ValueFromPipeline
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "Position" -Value $parameter.Position
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "ParameterType" -Value (GetParameterType($parameter.ParameterType))
                $hasParameterSet = $false
                [string[]] $validValues = @()
                if ($PSVersionTable.PSVersion.Major -gt 2) {
                    $validateSetAttributes = $parameter.Attributes | Where {
                        [ValidateSet].IsAssignableFrom($_.GetType())
                    }
                    if (($validateSetAttributes -ne $null) -and ($validateSetAttributes.Count -gt 0)) {
                        $hasParameterSet = $true
                        $validValues = $validateSetAttributes[0].ValidValues
                    }
                }
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "HasParameterSet" -Value $hasParameterSet
                $parameterInfo | Add-Member -MemberType NoteProperty -Name "ValidParamSetValues" -Value $validValues
                $parameterInfos += $parameterInfo
            }
            return (,$parameterInfos)
        }

        Function GetParameterSets {
            param ([System.Management.Automation.CommandInfo] $cmdInfo, [Switch] $NoWellKnownParameters)
            
            $wellknownparameternames = @("Verbose", "Debug", "ErrorAction", "WarningAction", "ErrorVariable",
                                         "WarningVariable", "OutVariable", "PipelineVariable", "OutBuffer",
                                         "Confirm", "WhatIf" )

            $parameterSets = $null
            try {
                $parameterSets = $cmdInfo.ParameterSets
            }
            catch [System.InvalidOperationException] { }
            catch [System.Management.Automation.PSNotSupportedException] { }
            catch [System.Management.Automation.PSNotImplementedException] { }
            
            if (($parameterSets -eq $null) -or ($parameterSets.Count -eq 0)) {
                return (,@())
            }

            [PSObject[]] $returnParameterSets = @()
            foreach ($parameterSet in $parameterSets) {
                $parameterSetInfo = new-object PSObject
                $parameterSetInfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $parameterSet.Name
                $parameterSetInfo | Add-Member -MemberType NoteProperty -Name "IsDefault" -Value $parameterSet.IsDefault

                if ($NoWellKnownParameters) { # depending on switch
                    $parameterSetInfo | Add-Member -MemberType NoteProperty -Name "Parameters" -Value (GetParameterInfo($parameterSet.Parameters | Where-Object { $_.Name -notin $wellknownparameternames } ))
                } else {
                    $parameterSetInfo | Add-Member -MemberType NoteProperty -Name "Parameters" -Value (GetParameterInfo($parameterSet.Parameters))
                }
                
                $returnParameterSets += $parameterSetInfo
            }
            return (,$returnParameterSets)
        }

        Function GetModuleInfo {
            
            param ([System.Management.Automation.CommandInfo] $cmdInfo)
            if ($cmdInfo.ModuleName -ne $null) {
                $moduleName = $cmdInfo.ModuleName
            } else {
                $moduleName = ""
            }
            $moduleInfo = new-object PSObject
            $moduleInfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $moduleName
            return $moduleInfo
        }

        Function ConvertToShowCommandInfo {
            param ([System.Management.Automation.CommandInfo] $cmdInfo, [Switch] $NoWellKnownParameters)
            $showCommandInfo = new-object PSObject
            $showCommandInfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $cmdInfo.Name
            $showCommandInfo | Add-Member -MemberType NoteProperty -Name "ModuleName" -Value $cmdInfo.ModuleName
            $showCommandInfo | Add-Member -MemberType NoteProperty -Name "Module" -Value (GetModuleInfo($cmdInfo))
            $showCommandInfo | Add-Member -MemberType NoteProperty -Name "CommandType" -Value $cmdInfo.CommandType
            $showCommandInfo | Add-Member -MemberType NoteProperty -Name "Definition" -Value $cmdInfo.Definition
            $showCommandInfo | Add-Member -MemberType NoteProperty -Name "ParameterSets" -Value (GetParameterSets -cmdInfo $cmdInfo -NoWellKnownParameters:$($NoWellKnownParameters.IsPresent))
            return $showCommandInfo
        }
    }

    Process {

        if ($Commands) {
            # If set: Only the given commands

            foreach ($command in @(Get-Command -Name $Commands)) {
                Write-Output (ConvertToShowCommandInfo -cmdInfo $command -NoWellKnownParameters:$($NoWellKnownParameters.IsPresent))
            }

        } else {
            # All registered commands

            $commandList = @("Cmdlet", "Function", "Script", "ExternalScript")
            if ($PSVersionTable.PSVersion.Major -gt 2) {
                $commandList += "Workflow"
            }

            foreach ($command in @(Get-Command -CommandType $commandList)) {
                Write-Output (ConvertToShowCommandInfo -cmdInfo $command -NoWellKnownParameters:$($NoWellKnownParameters.IsPresent))
            }
        }

    }

    End {
        Write-Debug "Finished $($MyInvocation.Mycommand)."
    }
}
