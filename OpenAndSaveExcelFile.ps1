while ($true) {
    
    # Configure the interval to check if the file success.flag exists.
    Start-Sleep -Seconds 3
    # Configure the Excel file name.
    $path = "File.xlsx"

    # If the excel file exists and the file "success.flag" does not exist, then open, calculate, save and close.
    if ((Test-Path -Path $path) -and !(Test-Path -Path "success.flag" -PathType Leaf)) {
        $file = Get-Item $path
        $excel = New-Object -COM "Excel.Application"
        $excel.Visible = $false

        $b = $excel.Workbooks.Open((Resolve-Path $path).Path)
        $b.Application.CalculateFull
        $b.Save()
        $b.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($b) | Out-Null
        Remove-Variable -Name b

        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Variable -Name excel

        if (!(Test-Path -Path "success.flag" -PathType Leaf)) {
            New-Item "success.flag" -type file
        }
        Write-Output ((Get-Date).ToString() + " File was opened and closed.")
    }
}
