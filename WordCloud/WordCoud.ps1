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
    [Alias('wordcloud', 'wcloud')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [Alias('Text', 'String', 'Words', 'Document', 'Page')]
        [AllowEmptyString()]
        [string[]]
        $InputString,

        [Parameter(Mandatory, Position = 1)]
        [Alias('OutFile', 'ExportPath', 'ImagePath')]
        [ValidateScript(
            { Test-Path -IsValid $_ -PathType Leaf }
        )]
        [string[]]
        $Path,

        [Parameter()]
        [Alias('ColourSet')]
        [KnownColor[]]
        $ColorSet = [Enum]::GetValues([KnownColor]),

        [Parameter()]
        [Alias('MaxColours')]
        [int]
        $MaxColors,

        [Parameter()]
        [Alias('FontFace')]
        [string]
        $FontFamily = 'Consolas',

        [Parameter()]
        [FontStyle]
        $FontStyle = [FontStyle]::Regular,

        [Parameter()]
        [Alias('MaxWordSize')]
        [double]
        $MaxFontSize = 200,

        [Parameter()]
        [Alias('GraphicsUnit', 'Unit')]
        [GraphicsUnit]
        $SizeUnit = [GraphicsUnit]::Point,

        [Parameter()]
        [ValidateRange(1, 20)]
        $DistanceStep = 5,

        [Parameter()]
        [ValidateRange(1, 50)]
        $RadialGranularity = 15,

        [Parameter()]
        [AllowNull()]
        [Alias('BackgroundColour')]
        [Nullable[KnownColor]]
        $BackgroundColor = [KnownColor]::Black,

        [Parameter()]
        [switch]
        [Alias('Greyscale', 'Grayscale')]
        $Monochrome
    )
    begin {
        if (-not (Test-Path -Path $Path)) {
            $Path = (New-Item -ItemType File -Path $Path).FullName
        }
        else {
            $Path = (Get-Item -Path $Path).FullName
        }

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
        $WordList = [List[string]]::new()

        $WordHeightTable = @{}
        $WordSizeTable = @{}

        # Create a graphics object to measure the text's width and height.
        $DummyImage = [Bitmap]::new(1, 1)

        $TotalSize = [SizeF]::Empty

        $MaxColors = if ($PSBoundParameters.ContainsKey('MaxColors')) { $MaxColors } else { [int]::MaxValue }

        if ($MaxColors) {
            $ColorList = $ColorSet |
                Sort-Object {Get-Random} |
                Select-Object -First $MaxColors |
                ForEach-Object {
                if (-not $Monochrome) {
                    [Color]::FromKnownColor($_)
                }
                else {
                    $Brightness = [Color]::FromKnownColor($_).GetBrightness() * 255
                    [Color]::FromArgb(1, $Brightness, $Brightness, $Brightness)
                }
            }
        }

        $ColorList = $ColorList | Where-Object {
            if ($BackgroundColor) {
                $_.Name -notmatch $BackgroundColor -and
                $_.GetSaturation() -gt 0.5
            }
            else {
                $_.GetSaturation() -gt 0.5
            }
        } | Sort-Object -Descending {
            $Value = $_.GetBrightness()
            $Random = (-$Value..$Value | Get-Random) / (1 - $_.GetSaturation())
            $Value + $Random
        }

        $RectangleList = [List[RectangleF]]::new()

        $RadialDistance = 0
        $ColorIndex = 0
    }
    process {
        $WordList.AddRange(
            $InputString.Split($SplitChars, [StringSplitOptions]::RemoveEmptyEntries).Where{
                $_ -notmatch "^$ExcludedWords$|^[^a-z]+$"
            } -as [string[]]
        )
    }
    end {
        foreach ($Word in $WordList) {
            # Count occurrence of each word
            $WordHeightTable[$Word] ++
        }

        $HighestWordCountt = $WordHeightTable.Values |
            Measure-Object -Maximum |
            ForEach-Object -MemberName Maximum

        try {
            foreach ($Word in $WordHeightTable.Keys.Clone()) {
                $WordHeightTable[$Word] = [Math]::Round( ($WordHeightTable[$Word] / $HighestWordCountt) * $MaxFontSize)

                $Font = [Font]::new(
                    $FontFamily,
                    $WordHeightTable[$Word],
                    [FontStyle]::Regular,
                    [GraphicsUnit]::Point
                )

                $Graphics = [Graphics]::FromImage($DummyImage)
                $WordSizeTable[$Word] = $Graphics.MeasureString($Word, $Font)
                $TotalSize += $WordSizeTable[$Word]
            }
        }
        finally {
            $Graphics.Dispose()
        }

        $SortedWordList = $WordHeightTable.Keys |
            Sort-Object -Descending { $WordSizeTable[$_].Width * $WordSizeTable[$_].Height } |
            Select-Object -First 100

        # Keep image square
        $SideLength = ($TotalSize.Height + $TotalSize.Width) / 2
        $TotalSize = [SizeF]::new($SideLength / 2, $SideLength / 2)

        $FinalImageSize = $TotalSize.ToSize()
        $CentrePoint = [PointF]::new($FinalImageSize.Height / 2, $FinalImageSize.Width / 2)
        Write-Verbose "Final Image size will be $FinalImageSize with centrepoint $CentrePoint"

        $WordCloudImage = [Bitmap]::new($DummyImage, $FinalImageSize)
        $DrawingSurface = [Graphics]::FromImage($WordCloudImage)

        if ($BackgroundColor) {
            $DrawingSurface.Clear([Color]::FromKnownColor($BackgroundColor))
        }
        $DrawingSurface.SmoothingMode = [Drawing2D.SmoothingMode]::AntiAlias
        $DrawingSurface.TextRenderingHint = [Text.TextRenderingHint]::AntiAlias

        :words foreach ($Word in $SortedWordList) {
            $Font = [Font]::new(
                $FontFamily,
                $WordHeightTable[$Word],
                [FontStyle]::Regular,
                [GraphicsUnit]::Point
            )

            $WordRectangle = $null
            do {
                if ($RadialDistance -gt $FinalImageSize.Height) {
                    continue words
                }
                $IsColliding = $false
                $AngleIncrement = 360 / ( ($RadialDistance + 1) * $RadialGranularity / 10 )

                for ($Angle = 0; $Angle -le 360; $Angle += $AngleIncrement) {
                    $Radians = Convert-ToRadians -Degrees $Angle
                    $Complex = [Complex]::FromPolarCoordinates($RadialDistance, $Radians)

                    if ($Complex -eq 0) {
                        # Offset slighty from center
                        $OffsetX = $WordSizeTable[$Word].Width / 1.5
                        $OffsetY = $WordSizeTable[$Word].Height / 1.5
                        $DrawLocation = [PointF]::new($CentrePoint.X - $OffsetX, $CentrePoint.Y - $OffsetY)
                    }
                    else {
                        $DrawLocation = [PointF]::new($CentrePoint.X + $Complex.Real, $CentrePoint.Y + $Complex.Imaginary)
                    }

                    $WordRectangle = [RectangleF]::new($DrawLocation, $WordSizeTable[$Word])

                    foreach ($Rectangle in $RectangleList) {
                        $IsColliding = (
                            $WordRectangle.IntersectsWith($Rectangle) -or
                            $WordRectangle.Top -lt 0 -or
                            $WordRectangle.Bottom -gt $FinalImageSize.Height -or
                            $WordRectangle.Left -lt 0 -or
                            $WordRectangle.Right -gt $FinalImageSize.Width
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
                    $RadialDistance += 5
                }
            } while ($IsColliding)

            $RectangleList.Add($WordRectangle)

            $Color = $ColorList[$ColorIndex]

            $ColorIndex++
            if ($ColorIndex -ge $ColorList.Count) {
                $ColorIndex = 0
            }

            $DrawingSurface.DrawString($Word, $Font, [SolidBrush]::new($Color), $DrawLocation)
        }

        $DrawingSurface.Flush()
        $WordCloudImage.Save($Path, [Imaging.ImageFormat]::Png)
    }
}






# Link to Python word cloud position figuring code:
# https://github.com/amueller/word_cloud/blob/b79b3d69a65643dbd421a027e66760a4398e91b3/wordcloud/wordcloud.py#L471