<#
MailingList-Chart.ps1

This script takes the MailingList data and creates
a native PowerShell pie chart.

Created: 13 Aug 2025
Updated: 14 Aug 2025
Author: Eric K. Miller
#>

using namespace System.Windows.Forms.DataVisualization.Charting
using namespace System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

$filepath = '.\Documents\PowerShell'
$csv = 'iContact_Salesforce_Names.csv'

$Data = Import-Csv "$filepath\$csv"

# Instantiate the objects we will use to create the chart
$Chart = New-Object Chart
#$Chart = [Chart]::new()  # alternate syntax for the above
$ChartArea = New-Object ChartArea
$Series = New-Object Series
$Series.ChartType = [SeriesChartType]::Pie

# Add ChartArea to Chart, Series to Chart
#(automatically assigned to the ChartArea)
$Chart.ChartAreas.Add($ChartArea)
$Chart.Series.Add($Series)

$Chart.Width = 800
$Chart.Height = 600

# Add data "points" to the Series
($Data.OnlyIn | Get-Unique) | foreach {
    [void]$Series.Points.AddXY($_, ($Data.OnlyIn -like $_).Length)
}

# Add Series properties
$Series["PieLabelStyle"] = "Outside"  # $Series.CustomProperties
$Series["PieLineColor"] = "Black"  # $Series.CustomProperties
$Series.Label = "#AXISLABEL (#PERCENT{P0})"
$Series.LabelToolTip = "#VAL"  # shows values when hovering over the labels

# If we want a legend, uncomment the below
#$Legend = New-Object Legend
#$Chart.Legends.Add($Legend)

# Add ChartTitle properties
$ChartTitle = New-Object Title
$Font = New-Object System.Drawing.Font @('Lucida Console', '15', [System.Drawing.FontStyle]::Bold)
$ChartTitle.Text = "Breakdown of Contacts in Different Databases"
$ChartTitle.Font = $Font
$Chart.Titles.Add($ChartTitle)

# If we want to save the chart as an image
#$Chart.SaveImage("$filepath\MailingList-Chart.png", [ChartImageFormat]::Png)

# Create the Form and add the Chart control
$Form = New-Object Form
$Form.Text = "PowerShell Chart for MailingList Data"
$Form.Width = 1.05*($Chart.Width)
$Form.Height = 1.15*($Chart.Height)
$Form.Controls.Add($Chart)
$Form.ShowDialog() | Out-Null  # Display the Form with Chart visualization
