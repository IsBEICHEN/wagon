param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $false)]
    [string]$OutputDocx
)

if (-not (Test-Path -LiteralPath $InputDocx)) {
    throw "Input DOCX not found: $InputDocx"
}

if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
    $dir = Split-Path -Parent $InputDocx
    $name = [System.IO.Path]::GetFileNameWithoutExtension($InputDocx)
    $OutputDocx = Join-Path $dir ($name + "_square_wrap_images.docx")
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Open($InputDocx, $false, $false)

    # Word constants:
    # wdWrapSquare = 0
    # wdRelativeHorizontalPositionMargin = 0
    # wdRelativeVerticalPositionParagraph = 2
    # wdShapeCenter = -999995
    # msoTrue = -1
    $usableWidth = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin

    # Convert inline pictures to floating shapes, then apply square wrapping.
    for ($i = $doc.InlineShapes.Count; $i -ge 1; $i--) {
        $inline = $doc.InlineShapes.Item($i)
        try {
            if ($inline.Width -gt $usableWidth) {
                $inline.LockAspectRatio = -1
                $inline.Width = $usableWidth
            }
            $shape = $inline.ConvertToShape()
            $shape.WrapFormat.Type = 0
            $shape.WrapFormat.DistanceTop = 6
            $shape.WrapFormat.DistanceBottom = 6
            $shape.WrapFormat.DistanceLeft = 6
            $shape.WrapFormat.DistanceRight = 6
            $shape.RelativeHorizontalPosition = 0
            $shape.RelativeVerticalPosition = 2
            $shape.Left = -999995
            $shape.LockAnchor = $false
        } catch {
        }
    }

    # Normalize any pictures that were already floating.
    foreach ($shape in $doc.Shapes) {
        try {
            if ($shape.Width -gt $usableWidth) {
                $shape.LockAspectRatio = -1
                $shape.Width = $usableWidth
            }
            $shape.WrapFormat.Type = 0
            $shape.WrapFormat.DistanceTop = 6
            $shape.WrapFormat.DistanceBottom = 6
            $shape.WrapFormat.DistanceLeft = 6
            $shape.WrapFormat.DistanceRight = 6
            $shape.RelativeHorizontalPosition = 0
            $shape.RelativeVerticalPosition = 2
            $shape.Left = -999995
            $shape.LockAnchor = $false
        } catch {
        }
    }

    $doc.SaveAs([ref]$OutputDocx)
    Write-Output $OutputDocx
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
