<#
Data Analysis: Get-RandomLetterOrder.ps1
Author: Eric K. Miller
Last updated: 23 January 2026

This script contains PowerShell code to shuffle letters in a word to
scramble or help unscramble the word.
#>

function Get-RandomLetterOrder {
    <#
    .SYNOPSIS
        Shuffles letters in a word and outputs the shuffled result.

    .DESCRIPTION
        This function shuffles letters in a word to scramble or unscramble
    the word. The user has the option of ensuring the output word is
    different from the original.
    
    .PARAMETER Word
        A string of the word to shuffle.

    .PARAMETER EnsureDifferent
        Indicates whether the function should return a new word.
    
    .EXAMPLE
        'iaeontto' | Get-RandomLetterOrder -EnsureDifferent
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [Alias('wordshuffle')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Word,

        [Parameter()]
        [switch]$EnsureDifferent
    )

    process {
        $original = $Word
        $letters = $Word.ToCharArray()
        $lettersCount = $letters.Count
        do {
            $shuffled = $letters | Get-Random -Count $lettersCount
            $result = -join $shuffled
        }
        while ($EnsureDifferent -and $result -eq $original -and $lettersCount -gt 1)
        
        return $result
    }
}