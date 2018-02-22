function Import-PDFText {
    <#
    .SYNOPSIS
        Import-PdfText
        Imports the raw text data of a PDF file as readable text.
    .DESCRIPTION
        Takes the path of a PDF file, loads the file in and converts the text into readable
        string data, before outputting it as a complete string.
    .EXAMPLE
        PS C:\> Import-PDFText -Path .\Test.pdf | Set-Content $env:Temp\test.txt

        Returns all of the text in Test.pdf as a string using StringBuilder and stores it to a
        file in the temp folder.
    .INPUTS
        Takes a file path as pipeline input.
    .OUTPUTS
        Outputs the entire text content of the PDF file as string.
    .NOTES
        Requires the iTextSharp library, which can be downloaded from here:
        https://sourceforge.net/projects/itextsharp/

        Place the DLL in the same folder as the script, and execute the function.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateScript({ Test-Path $_ })]
        [string]
        $Path
    )
    begin {
        if (-not ([System.Management.Automation.PSTypeName]'iTextSharp.Text.Pdf.PdfReader').Type) {
            Add-Type -Path "$PSScriptRoot\itextsharp.dll"
        }
    }
    process {
        $Reader = New-Object 'iTextSharp.Text.Pdf.PdfReader' -ArgumentList $Path
        $PdfText = New-Object 'System.Text.StringBuilder'

        for ($Page = 1; $Page -le $Reader.NumberOfPages; $Page++) {
            $Strategy = New-Object 'iTextSharp.Text.Pdf.Parser.SimpleTextExtractionStrategy'
            $CurrentText = [iTextSharp.Text.Pdf.Parser.PdfTextExtractor]::GetTextFromPage($Reader, $Page, $Strategy)
            $PdfText.AppendLine([System.Text.Encoding]::UTF8.GetString([System.Text.ASCIIEncoding]::Convert([System.Text.Encoding]::Default, [System.Text.Encoding]::UTF8, [System.Text.Encoding]::Default.GetBytes($CurrentText))))
        }
        $Reader.Close()

        $PdfText.ToString()
    }
}
