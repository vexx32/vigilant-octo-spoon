using namespace System.Drawing

# PowerShell Core uses System.Drawing.Common assembly instead of System.Drawing
if ($PSEdition -eq 'Core') {
    Add-Type -AssemblyName 'System.Drawing.Common'
}
else {
    Add-Type -AssemblyName 'System.Drawing'
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

$SplitChars = [char[]]( ' ', "`n" )
$ForbiddenChars = '[^a-z]'
$WordList = $Text.Split($SplitChars, [StringSplitOptions]::RemoveEmptyEntries).Where{ $_ -notmatch "^$ExcludedWords$|^[^a-z]+$" }

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

$TotalSize = [SizeF]::Empty
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

[Size]$FinalImageSize = [Size]::new($TotalSize.Width, $TotalSize.Height)
Write-Host $FinalImageSize

$FinalImage = [Bitmap]::new($Image, $FinalImageSize)
[Graphics]$DrawingSurface = [Graphics]::FromImage($FinalImage)

$DrawingSurface.Clear($BackgroundColor)
$DrawingSurface.SmoothingMode = [Drawing2D.SmoothingMode]::AntiAlias
$DrawingSurface.TextRenderingHint = [Text.TextRenderingHint]::AntiAlias

$x = $FinalImageSize.Width / 2
$y = $FinalImageSize.Height / 2

$RectangleTable = @{}

foreach ($Word in $WordHeight.Keys) {
    $Font = [Font]::new(
        $FontFamily,
        $WordHeight[$Word],
        [FontStyle]::Regular,
        [GraphicsUnit]::Point
    )
    $MaxX = $FinalImageSize.Width - $WordSizes[$Word].Width
    $MaxY = $FinalImageSize.Height - $WordSizes[$Word].Height

    do {
        $IsColliding = $false

        $StartX = 0..$MaxX | Get-Random
        $StartY = 0..$MaxY | Get-Random
        [PointF]$Location = [PointF]::new($StartX, $StartY)
        $Rect = [RectangleF]::new($Location, $WordSizes[$Word])

        foreach ($Rectangle in $RectangleTable.Values) {
            if ($Rect.IntersectsWith($Rectangle)) {
                $IsColliding = $true
            }
        }
    } while ($IsColliding)

    $RectangleTable[$Word] = $Rect

    $RandomColor = [Enum]::GetValues([KnownColor]) | Where-Object {$_ -notmatch 'black|dark'} | Get-Random
    $Color = [Color]::FromKnownColor([KnownColor]$RandomColor)

    $DrawingSurface.DrawString($Word, $Font, [SolidBrush]::new($Color), $Location)
}

$DrawingSurface.Flush()
$FinalImage.Save('C:\Users\Joel\Desktop\TestImage.png', [Imaging.ImageFormat]::Png)

# Link to Python word cloud position figuring code:
# https://github.com/amueller/word_cloud/blob/b79b3d69a65643dbd421a027e66760a4398e91b3/wordcloud/wordcloud.py#L471