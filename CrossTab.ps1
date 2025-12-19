<#
Data Analysis: CrossTab.ps1
Author: Eric K. Miller
Last updated: 18 December 2025

This script contains PowerShell code for creating a cross-tabulation
(crosstab), or contingency table, returned as an object. It helps
summarize categorical data to illuminate insights.
#>

function New-CrossTab {
    [CmdletBinding(SupportsShouldProcess)]
    [Alias('crosstab')]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$Data,

        [Parameter(Mandatory)]
        [string]$RowEntries,

        [Parameter(Mandatory)]
        [string]$ColumnEntries
    )

    begin {$allData = @()}

    process {$allData += $Data}

    end {
        if ($allData[0].PSObject.Properties.Name -notcontains $RowEntries) {
            throw "Property '$RowEntries' not found in the data."
        }
        if ($allData[0].PSObject.Properties.Name -notcontains $ColumnEntries) {
            throw "Property '$ColumnEntries' not found in the data."
        }

        $unique_rowValues = $allData.$RowEntries | Select-Object -Unique
        $unique_colValues = $allData.$ColumnEntries | Select-Object -Unique

        $CrossTab = foreach ($row in $unique_rowValues)
        {
            $rowData = $allData | Where-Object $RowEntries -eq $row
            $grouped = $rowData | Group-Object $ColumnEntries

            $properties = [ordered]@{$RowEntries = $row}  # Row labels are first

            # Build the crosstab
            foreach ($col in $unique_colValues) {
                $properties[$col] = ($grouped | Where-Object Name -eq $col).Count
            }
            [PSCustomObject]$properties
        }

        # Add row totals column via calculated property
        $CrossTab = $CrossTab | Select-Object *,
            @{
                N='Row Totals'
                E={
                    ($_.PSObject.Properties | Select-Object -Skip 1 |
                    Measure-Object -Property Value -Sum).Sum
                }
            }

        # Add col totals row via new object
        $colTotal = [ordered]@{$RowEntries = 'Col Totals'}

        # Get col names (skip first, which is row label)
        $columnNames = $CrossTab[0].PSObject.Properties.Name | Select-Object -Skip 1

        foreach ($col in $columnNames) {
            $colTotal[$col] = ($CrossTab | Measure-Object -Property $col -Sum).Sum
        }

        $CrossTab = @($CrossTab) + @([PSCustomObject]$colTotal)

        return $CrossTab
    }
}
