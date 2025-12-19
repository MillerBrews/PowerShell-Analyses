<#
MailingList-Compare.ps1

This script takes lists from two spreadsheets
and compares them to see missing entries in each.

Created: 10 Apr 2025
Updated: 13 Aug 2025
Author: Eric K. Miller
#>

$filepath = '.\Documents\PowerShell'
$csv1 = 'iContact.csv'
$csv2 = 'Salesforce.csv'

# First, write a function to create the data objects,
# so we can reuse the capability for both datasets.
function Make-DataObject {
    param (
        [Parameter(Mandatory)][string]$File,
        [Parameter()][string]$Source = '',
        [Parameter()]$SelectColumns = '*'
    )

    $DataObject = Import-Csv $File -Encoding UTF8
    # add new column Source to object with value input for $Source parameter
    $DataObject = $DataObject | Select-Object *, @{Name='Source'; Expression={$Source}} |
        Select-Object $SelectColumns # "script block"
    return $DataObject
}

# Run the function on the two datasets and their specific attributes.
$params_src1 = @{File = "$filepath\$csv1"
    Source = 'iContact'
    SelectColumns = 'fname','lname','Source'  # we only care about name and source to compare
}
$iC = Make-DataObject @params_src1  # "splatting"

$params_src2 = @{File = "$filepath\$csv2"
    Source = 'Salesforce'
    SelectColumns = 'First Name','Last Name','Source'
}
$Sf = Make-DataObject @params_src2

# Next, write a function to create (somewhat)
# standardized lists of names to compare.
function Make-NameList {
    param (
        [Parameter(Mandatory)]$DataObject,
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName
    )
    
    $NAMES = foreach ($row in $DataObject) {
        $row.$LastName.ToUpper() + ', ' + $row.$FirstName.ToUpper()
    }
    $NAMES = $NAMES | Sort-Object | Get-Unique
    return $NAMES
}

# Run the function on the two data objects created from above.
$params_namelist1 = @{DataObject = $iC
    FirstName = 'fname'
    LastName = 'lname'
}
$NAMES_I = Make-NameList @params_namelist1

$params_namelist2 = @{DataObject = $Sf
    FirstName = 'First Name'
    LastName = 'Last Name'
}
$NAMES_S = Make-NameList @params_namelist2

# Finally, compare the two lists of names. The output will be
# a CSV showing which dataset the names came from.
$COMPARISON_TRANSLATION = @{'<=' = 'Salesforce'
                            '=>' = 'iContact'}

$comparisonResult = Compare-Object -ReferenceObject $NAMES_S -DifferenceObject $NAMES_I
$comparisonResult = $comparisonResult | Select-Object *, @{Name='OnlyIn'; Expression={$COMPARISON_TRANSLATION[$_.SideIndicator]}}

#$comparisonResult.Length  # 2082
#$comparisonResult | Out-GridView  # to view in an interactive pop-up

$comparisonResult | Export-Csv "$filepath\iContact_Salesforce_Names.csv" -NoTypeInformation
