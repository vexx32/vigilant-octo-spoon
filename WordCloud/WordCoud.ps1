using namespace System.Drawing
using namespace System.Collections.Generic
using namespace System.Numerics

# PowerShell Core uses System.Drawing.Common assembly instead of System.Drawing
if ($PSEdition -eq 'Core') {
    Add-Type -AssemblyName 'System.Drawing.Common'
}
else {
    Add-Type -AssemblyName 'System.Drawing'
}

function Convert-ToRadians {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline)]
        [ValidateRange(0, 360)]
        [double]
        $Degrees
    )
    process {
        ([Math]::PI / 180) * $Degrees
    }
}

function New-WordCloud {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [Alias('Text', 'String', 'Words', 'Document', 'Page')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $InputString,

        [Parameter(Mandatory, Position = 1)]
        [Alias('OutFilePath','ExportPath')]
        [ValidateScript(
            { Test-Path -IsValid $_ -PathType Leaf }
        )]
        [string[]]
        $ImagePath,

        [Parameter()]
        [Alias('ColourSet')]
        [KnownColor[]]
        $ColorSet = [Enum]::GetValues([KnownColor]),

        [Parameter()]
        [Alias('MaxColours')]
        [int]
        $MaxColors,

        [Parameter()]
        [ValidateRange(1, 20)]
        $DistanceStep = 5,

        $RadialGranularity,

        $BackgroundColor = [KnownColor]::Black,

        [switch]
        [Alias('Greyscale', 'Grayscale')]
        $Monochrome
    )
}
$Text = Get-Content -Path "$PSScriptRoot\Test.txt"

$ExcludedWords = @(
    'a', 'about', 'above', 'after', 'again', 'against', 'all', 'also', 'am', 'an', 'and', 'any', 'are', 'aren''t', 'as',
    'at', 'be', 'because', 'been', 'before', 'being', 'below', 'between', 'both', 'but', 'by', 'can', 'can''t',
    'cannot', 'com', 'could', 'couldn''t', 'did', 'didn''t', 'do', 'does', 'doesn''t', 'doing', 'don''t', 'down',
    'during', 'each', 'else', 'ever', 'few', 'for', 'from', 'further', 'get', 'had', 'hadn''t', 'has', 'hasn''t',
    'have', 'haven''t', 'having', 'he', 'he''d', 'he''ll', 'he''s', 'hence', 'her', 'here', 'here''s', 'hers',
    'herself', 'him', 'himself', 'his', 'how', 'how''s', 'however', 'http', 'i', 'i''d', 'i''ll', 'i''m', 'i''ve', 'if',
    'in', 'into', 'is', 'isn''t', 'it', 'it''s', 'its', 'itself', 'just', 'k', 'let''s', 'like', 'me', 'more', 'most',
    'mustn''t', 'my', 'myself', 'no', 'nor', 'not', 'of', 'off', 'on', 'once', 'only', 'or', 'other', 'otherwise',
    'ought', 'our', 'ours', 'ourselves', 'out', 'over', 'own', 'r', 'same', 'shall', 'shan''t', 'she', 'she''d',
    'she''ll', 'she''s', 'should', 'shouldn''t', 'since', 'so', 'some', 'such', 'than', 'that', 'that''s', 'the',
    'their', 'theirs', 'them', 'themselves', 'then', 'there', 'there''s', 'therefore', 'these', 'they', 'they''d',
    'they''ll', 'they''re', 'they''ve', 'this', 'those', 'through', 'to', 'too', 'under', 'until', 'up', 'very', 'was',
    'wasn''t', 'we', 'we''d', 'we''ll', 'we''re', 'we''ve', 'were', 'weren''t', 'what', 'what''s', 'when', 'when''s',
    'where', 'where''s', 'which', 'while', 'who', 'who''s', 'whom', 'why', 'why''s', 'with', 'won''t', 'would',
    'wouldn''t', 'www', 'you', 'you''d', 'you''ll', 'you''re', 'you''ve', 'your', 'yours', 'yourself', 'yourselves'
) -join '|'

$SplitChars = [char[]]" `n.,`"?!{}[]:'()`“`”"
$WordList = $Text.Split($SplitChars, [StringSplitOptions]::RemoveEmptyEntries).Where{
    $_ -notmatch "^$ExcludedWords$|^[^a-z]+$"
}

$WordHeight = @{}
foreach ($Word in $WordList) {
    # Count occurrence of each word
    $WordHeight[$Word] = $WordHeight[$Word] + 1
}

[double]$HighestFrequency = $WordHeight.Values | Measure-Object -Maximum | ForEach-Object -MemberName Maximum
[double]$LargestSize = 200 # in Points

foreach ($Word in $WordHeight.Keys.Clone()) {
    # Replace occurrence number with size number
    $WordHeight[$Word] = [Math]::Round( ($WordHeight[$Word] / $HighestFrequency) * $LargestSize)
}

# Dummy image to use for measurements
[Bitmap]$Image = [Bitmap]::new(1, 1)

$FontFamily = "Consolas"

$BackgroundColor = [Color]::Black

# Create a graphics object to measure the text's width and height.
[Graphics] $Graphics = [Graphics]::FromImage($Image)

[SizeF]$TotalSize = [SizeF]::Empty
$WordSizes = @{}
foreach ($Word in $WordHeight.Keys) {
    $Font = [Font]::new(
        $FontFamily,
        $WordHeight[$Word],
        [FontStyle]::Regular,
        [GraphicsUnit]::Point
    )
    $WordSizes[$Word] = $Graphics.MeasureString($Word, $Font)
    $TotalSize += $WordSizes[$Word]
}
$Graphics.Flush()
$Graphics.Dispose()

# Keep image square
$SquareSizeLength = ($TotalSize.Height + $TotalSize.Width) / 2
$TotalSize = [SizeF]::new($SquareSizeLength / 2, $SquareSizeLength / 2)

[Size]$FinalImageSize = [Size]::new($TotalSize.Width, $TotalSize.Height)
Write-Host $FinalImageSize

$FinalImage = [Bitmap]::new($Image, $FinalImageSize)
[Graphics]$DrawingSurface = [Graphics]::FromImage($FinalImage)

$DrawingSurface.Clear($BackgroundColor)
$DrawingSurface.SmoothingMode = [Drawing2D.SmoothingMode]::AntiAlias
$DrawingSurface.TextRenderingHint = [Text.TextRenderingHint]::AntiAlias

[List[KnownColor]]$ColorList = [KnownColor[]](
    [Enum]::GetValues([KnownColor]) |
        Where-Object {
            $_ -notmatch 'black|dark' -and
            [Color]::FromKnownColor([KnownColor]$_).GetSaturation() -gt 0.5
        } |
        Sort-Object {
            $Color = [Color]::FromKnownColor([KnownColor]$_)
            $Value = $Color.GetBrightness()
            $Random = (-$Value..$Value | Get-Random) / (1 - $Color.GetSaturation())
            $Value + $Random
        } -Descending
)

$Centre = $FinalImageSize.Height / 2
$OrderedWords = $WordHeight.Keys |
    Sort-Object { $WordSizes[$_].Width * $WordSizes[$_].Height } -Descending |
    Select-Object -First 100

$RectangleList = [List[RectangleF]]::new()
$Distance = 0
$ColorIndex = 0
:words foreach ($Word in $OrderedWords) {
    $Font = [Font]::new(
        $FontFamily,
        $WordHeight[$Word],
        [FontStyle]::Regular,
        [GraphicsUnit]::Point
    )

    $Rect = $null
    do {
        $IsColliding = $false

        if ($Distance -gt $Centre) {
            continue words
        }

        $RadialGranularity = ($Distance + 1) * 1.5
        for ($Angle = 0; $Angle -le 360; $Angle += 360 / $RadialGranularity) {
            $Radians = $Angle | Convert-ToRadians
            $Complex = [Complex]::FromPolarCoordinates($Distance, $Radians)

            if ($Complex -eq 0) {
                $OffsetX = $WordSizes[$Word].Width / 1.5
                $OffsetY = $WordSizes[$Word].Height / 1.5
                $Location = [PointF]::new($Centre - $OffsetX, $Centre - $OffsetY)
            }
            else {
                $Location = [PointF]::new($Centre + $Complex.Real, $Centre + $Complex.Imaginary)
            }

            $Rect = [RectangleF]::new($Location, $WordSizes[$Word])

            foreach ($Rectangle in $RectangleList) {
                $IsColliding = (
                    $Rect.IntersectsWith($Rectangle) -or
                    $Rect.Top -lt 0 -or
                    $Rect.Bottom -gt $FinalImageSize.Height -or
                    $Rect.Left -lt 0 -or
                    $Rect.Right -gt $FinalImageSize.Width
                )

                if ($IsColliding) {
                    break
                }
            }

            if (!$IsColliding) {
                break
            }
        }

        if ($IsColliding) {
            $Distance += 5
        }
    } while ($IsColliding)
    $RectangleList.Add($Rect)

    $KnownColor = $ColorList[$ColorIndex]
    $Color = [Color]::FromKnownColor($KnownColor)

    $ColorIndex++
    if ($ColorIndex -ge $ColorList.Count) {
        $ColorIndex = 0
    }

    $DrawingSurface.DrawString($Word, $Font, [SolidBrush]::new($Color), $Location)
}

$DrawingSurface.Flush()
$FinalImage.Save("$PSScriptRoot\test.png", [Imaging.ImageFormat]::Png)

# Link to Python word cloud position figuring code:
# https://github.com/amueller/word_cloud/blob/b79b3d69a65643dbd421a027e66760a4398e91b3/wordcloud/wordcloud.py#L471