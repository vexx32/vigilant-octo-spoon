function Convert-ExcelToCsv {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = "Low")]
    param(
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [Alias('InputObject')]
        [ValidateNotNull()]
        [System.IO.FileInfo[]]
        $FilePath,

        [Parameter(Position = 1, Mandatory)]
        [Alias('OutputFolder')]
        [ValidateScript( { Test-Path -PathType Container -Path $_ })]
        [string]
        $DestinationFolder
    )
    begin {
        Write-Verbose "Starting Excel COM handler."
        $ExcelComHandler = New-Object -ComObject 'Excel.Application'

        $ExcelComHandler.Visible = $false
        $ExcelComHandler.DisplayAlerts = $false
    }
    process {
        $FilePath | ForEach-Object {
            Write-Verbose "Processing Excel file '$($_.Name)'"

            if ($PSCmdlet.ShouldProcess($_.Name, "Convert to CSV files")) {
                $Workbook = $ExcelComHandler.Workbooks.Open($_.FullName)
                foreach ($Worksheet in $Workbook.Worksheets) {
                    $CsvFileName = "{0}_{1}.csv" -f $_.BaseName, $Worksheet.Name
                    $SavePath = Join-Path -Path $DestinationFolder -ChildPath $CsvFileName

                    $Worksheet.SaveAs($SavePath, 6)
                }
            }
        }
    }
    end {
        $ExcelComHandler.Quit()
    }
}