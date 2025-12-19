<#
MailingList-Excel.ps1

This script takes the MailingList data, transfers it
to Excel, and creates a PivotTable and PivotChart.

Created: 17 Aug 2025
Updated: 20 Aug 2025
Author: Eric K. Miller
#>

$filepath = $psworkingfolder  # $profile variable
$csv = 'iContact_Salesforce_Names.csv'  # created from MailingList-Compare.ps1

$Data = Import-Csv "$filepath\$csv"

$headers = $Data[0].psobject.properties.Name[0,2]  # select first and third columns

try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true  # we can see the data populate in real time!

    $Workbook = $Excel.Workbooks.Add()
    $Worksheet = $Workbook.Worksheets.Item(1)
    $Worksheet.Name = "Data_Sheet"

    # Populate the data
    for ($j=0; $j -lt $headers.Length; $j++) {
        for ($i=0; $i -lt $Data.Length; $i++) {
            $Worksheet.Cells.Item(1,$j+1) = $headers[$j]
            $Worksheet.Cells.Item($i+2,$j+1) = $Data[$i].($headers[$j])
        }
    }

    $pivotWorksheet = $Workbook.Worksheets.Add()
    $pivotWorksheet.Name = "PivotTable"

    $numRows = $Data.Length
    $dataRange = $Worksheet.Range("A1:B$numrows")
    
    # Create the PivotTable
    # 1 is [Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase
    $pivotCache = $Workbook.PivotCaches().Create(1, $dataRange)
    $pivotTable = $pivotCache.CreatePivotTable($pivotWorksheet.Range("A2"), "PivotSheet")

    $rowField = $pivotTable.PivotFields($headers[1])
    $datField = $pivotTable.PivotFields($headers[0])

    # 1 is [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
    # 4 is [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
    $rowField.Orientation = 1
    $datField.Orientation = 4

    $valueField = $pivotTable.PivotFields("Count of InputObject")
    $valueField.Caption = "Count of Contacts"

    $pivotTable.RefreshTable() | Out-Null

    # Create the PivotChart
    $pivotChart = $Excel.Charts.Add()
    $pivotChart.SetSourceData($pivotTable.TableRange1)
    # 2 is [Microsoft.Office.Interop.Excel.XlChartLocation]::xlLocationAsObject
    $pivotChart.Location(2, $pivotWorksheet.Name) | Out-Null
    $pivotChart = $pivotWorksheet.ChartObjects(1).Chart
    # 5 is [Microsoft.Office.Interop.Excel.XlChartType]::xlPie  # Defaults to bar chart
    $pivotChart.ChartType = 5
    $pivotChart.HasTitle = $true
    $pivotChart.ChartTitle.Text = "PivotChart of Contacts in Different Databases"
    # -4107 is [Microsoft.Office.Interop.Excel.XlLegendPosition]::xlLegendPositionBottom
    $pivotChart.Legend.position = -4107
    
    # Format the DataLabels
    $series = $pivotChart.SeriesCollection(1)
    # Type, LegendKey, AutoText, HasLeaderLines, ...
    # 3 is xlDataLabelsShowPercent
    # adds LeaderLines, but need to move the dataLabel to see them
    $def = [Type]::Missing
    $series.ApplyDataLabels(3, $def, $def, $true)

    for ($i=0; $i -lt $series.Values.Length; $i++) {
        $dataLabel = $series.DataLabels($i+1)
        $dataLabel.Position = 2  # xlDataLabelPositionOutsideEnd
        $dataLabel.Format.TextFrame2.TextRange.Font.Size = 12
    }

    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "Excel Files (*.xlsx) | *.xlsx"
    $SaveFileDialog.ShowDialog() | Out-Null
    $ExcelFile = $SaveFileDialog.FileName
    
    $Workbook.SaveAs($ExcelFile)  # includes full path
}

catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}

finally {
    if ($Workbook) {
        $Workbook.Close()
    }
    if ($Excel) {
        $Excel.Quit()
    }

}
