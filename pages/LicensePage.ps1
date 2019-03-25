$LicensePage = New-UDPage -Name "Licenses" -Content {
    New-UDChart -Type Bar -Endpoint {
        $Cache:SortedLicenses | Out-UDChartData -LabelProperty 'License' -DataSet @(
            New-UDChartDataset -DataProperty 'UsedLicenses' -Label 'Used Licenses' -BackgroundColor '#80962F23' -HoverBackgroundColor '#80962F23'
            New-UDChartDataset -DataProperty 'AvailableLicenses' -Label 'Available Licenses' -BackgroundColor '#8014558C' -HoverBackgroundColor '#8014558C'
        )
    } -Labels @('Licenses') -Options @{
        scales = @{
            xAxes = @(
                @{
                    stacked = $true
                }
            )
            yAxes = @(
                @{
                    stacked = $true
                }
            )
        }
    }
}