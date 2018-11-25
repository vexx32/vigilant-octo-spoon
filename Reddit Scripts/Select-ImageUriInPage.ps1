<#
.SYNOPSIS
Selects and extracts all image URIs in a given webpage and sends them to the output stream.

.DESCRIPTION
Given an input URI or list of URIs either directly or via the pipeline, automatically examines
the resultant object of an Invoke-WebRequest against the URI for images contained in the page.
Matches URI strings using regex and automatically resolves relative URIs by comparing to the
parent page's own URI.

.PARAMETER Uri
One or more [uri] or [string] objects to be passed to Invoke-WebRequest containing valid URIs.

.EXAMPLE
Select-ImageUriInPage -Uri 'https://www.reddit.com/r/pics'

.EXAMPLE
'https://www.reddit.com/r/earthporn' | Select-ImageUriInPage

.NOTES
See also:
  Invoke-WebRequest
#>
function Select-ImageUriInPage {
    [CmdletBinding()]
    [Alias('suip')]
    param(
        [Parameter(
            Position = 0,
            Mandatory,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [Alias('URL', 'AbsoluteUri')]
        [ValidateNotNull()]
        [Uri[]]
        $Uri
    )
    process {
        foreach ($WebAddress in $Uri) {
            $ImageHtml = @(
                Invoke-WebRequest -Uri $WebAddress |
                    Select-Object -ExpandProperty Images |
                    Select-Object -ExpandProperty OuterHtml
            )
            if ($ImageHtml.Count -lt 1) {
                $PSCmdlet.WriteError("No images found in the page '$WebAddress'")
                continue
            }

            foreach ($Snippet in $ImageHtml) {
                if ($Snippet -match '(?<=src\=")(.*)(?=")') {
                    $ImagePath = $Matches[0]

                    if ($ImagePath -notmatch '^https*://') {
                        $ImagePath = "$($WebAddress.OriginalString)/$ImagePath" -replace "(?<!http(s)*:)//+", "/"
                    }

                    [uri] $ImagePath
                }
            }
        }
    }
}