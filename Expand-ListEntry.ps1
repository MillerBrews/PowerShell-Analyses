<#
Data Analysis: Expand-ListEntry.ps1
Author: Eric K. Miller
Last updated: 20 December 2025

This script contains PowerShell code to "unravel" rows that contain
lists as entries, creating a new row for each entry in a PSCustomObject.
It enables use of crosstab, in particular.
#>

function Expand-ListEntry {
    [CmdletBinding(SupportsShouldProcess)]
    [Alias('explode')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject[]]$DataObject,

        [Parameter(Mandatory)]
        [string]$ExplodeField
    )

    process {
        foreach ($item in $DataObject) {
            $list = $item.$ExplodeField

            if ($list -is [string]) {
                $list = $list -split ',' | ForEach-Object {$_.Trim()}
            }
            elseif ($list -eq $null) {
                $output = $item.PSObject.Copy()
                $output.$ExplodeField = $null
                $output
                continue
            }
            foreach ($val in $list) {
                $output = $item.PSObject.Copy()
                $output.$ExplodeField = $val
                $output
            }
        }
    }
}