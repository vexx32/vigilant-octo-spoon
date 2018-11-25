using namespace System.Drawing

function Export-Png {
    [CmdletBinding()]
    [Alias('epng')]
    param(
        [Parameter(
            Position = 0,
            Mandatory,
            ValueFromPipeline
        )]
        [string[]]
        [ValidateNotNullOrEmpty()]
        $String,

        [Parameter(
            Position = 1,
            Mandatory,
            ParameterSetName = "SaveFile"
        )]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path,

        [Parameter(
            Position = 1,
            Mandatory,
            ParameterSetName = "Clipboard"
        )]
        [switch]
        $ToClipboard
    )
    begin {
        [Bitmap]$Image = [Bitmap]::new(1, 1)
        [System.Collections.Generic.List[string]]$StringList = @()

        [int]$Width = 0
        [int]$Height = 0
        # Create the Font object for the image text drawing.
        [Font]$ImageFont = [Font]::new(
            [FontFamily]::GenericMonospace, 
            12, 
            [FontStyle]::Regular, 
            [GraphicsUnit]::Point
        )

        # Create a graphics object to measure the texts width and height.
        [Graphics] $Graphics = [Graphics]::FromImage($Image)
    }
    process {
        foreach ($Line in $String) {
            $StringList.Add($Line)
        }
    }
    end {
        $ImageText = $StringList -join "`n"
        # This is where the bitmap size is determined.
        $Width = [int] $Graphics.MeasureString($ImageText, $ImageFont).Width
        $Height = [int] $Graphics.MeasureString($ImageText, $ImageFont).Height

        # Create the bmpImage again with the correct size for the text and font.
        $Image = [Bitmap]::new($Image, [Size]::new($Width, $Height))

        # Add the colors to the new bitmap.
        $Graphics = [Graphics]::FromImage($Image)

        # Set Background color
        $Graphics.Clear([Color]::Black)
        $Graphics.SmoothingMode = [Drawing2D.SmoothingMode]::Default
        $Graphics.TextRenderingHint = [Text.TextRenderingHint]::SystemDefault
        $Graphics.DrawString($ImageText, $ImageFont, [SolidBrush]::new([Color]::FromArgb(192, 192, 192)), 0, 0)
        $Graphics.Flush()
        switch ($PSCmdlet.ParameterSetName) {
            "SaveFile" {
                try {
                    $Image.Save($Path, [Imaging.ImageFormat]::Png)
                }
                catch {
                    $PSCmdlet.WriteError($_)
                }
            }
            "Clipboard" {
                [System.Windows.Forms.Clipboard]::SetImage($Image)
            }
        }
    }
}