function Import-PDFText {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript({ Test-Path $_ })]
        [string]
        $Path
    )
    if (-not ([System.Management.Automation.PSTypeName]'iTextSharp.Text.Pdf.PdfReader').Type) {
        Add-Type -Path "$PSScriptRoot\itextsharp.dll"
    }

    $Reader = New-Object 'iTextSharp.Text.Pdf.PdfReader' -ArgumentList $Path
    $PdfText = New-Object 'System.Text.StringBuilder'

    for ($Page = 1; $Page -le $Reader.NumberOfPages; $Page++) {
        $Strategy = New-Object 'iTextSharp.Text.Pdf.Parser.SimpleTextExtractionStrategy'
        $CurrentText = [iTextSharp.Text.Pdf.Parser.PdfTextExtractor]::GetTextFromPage($Reader, $Page, $Strategy)
        $PdfText.Append([System.Text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert([System.Text.Encoding]::Default, [System.Text.Encoding]::UTF8, [System.Text.Encoding]::Default.GetBytes($CurrentText))))
    }
    $Reader.Close()

    $PdfText.ToString()
}
