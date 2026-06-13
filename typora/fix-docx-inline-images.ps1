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
    $OutputDocx = Join-Path $dir ($name + "_inline_images.docx")
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Open($InputDocx, $false, $false)

    # Word WdWrapType: 7 = wdWrapInline
    # Iterate backwards because converting a Shape to InlineShape mutates the collection.
    for ($i = $doc.Shapes.Count; $i -ge 1; $i--) {
        $shape = $doc.Shapes.Item($i)
        try {
            $shape.WrapFormat.Type = 7
            $shape.ConvertToInlineShape() | Out-Null
        } catch {
            # Some shapes cannot be converted; keep going so other images are fixed.
        }
    }

    # Also normalize inline pictures so they do not exceed page width.
    $usableWidth = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
    foreach ($inline in $doc.InlineShapes) {
        try {
            if ($inline.Width -gt $usableWidth) {
                $inline.LockAspectRatio = -1
                $inline.Width = $usableWidth
            }
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
