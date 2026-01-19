param(
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    [Parameter(Mandatory=$true)]
    [string]$OutputDir
)

try {
    # Create Output Dir if not exists
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
    }

    Write-Output "DEBUG: InputPath received: '$InputPath'"
    if (-not (Test-Path -LiteralPath $InputPath)) {
        Write-Error "ERROR: File does not exist at path: $InputPath"
        exit 1
    }
    Write-Output "DEBUG: File exists."

    $ppt = New-Object -ComObject PowerPoint.Application
    # Open ReadOnly, Untitled, WithWindow=msoFalse
    $presentation = $ppt.Presentations.Open($InputPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    $slidesData = @()
    
    foreach ($slide in $presentation.Slides) {
        $imageName = "slide_$($slide.SlideIndex).png"
        $imagePath = Join-Path $OutputDir $imageName
        
        # export path must be full path
        $slide.Export($imagePath, "PNG")
        
        $notes = ""
        try {
            if ($slide.HasNotesPage -eq -1) { # msoTrue is -1 in COM usually
                # Notes body is usually shape 2 on the notes page
                $notes = $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
            }
        } catch {
            Write-Warning "Could not extract notes for slide $($slide.SlideIndex)"
        }
        
        $slidesData += @{
            index = $slide.SlideIndex
            image = $imageName
            notes = $notes
        }
    }
    
    $presentation.Close()
    $ppt.Quit()
    
    # Clean up COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    $manifestPath = Join-Path $OutputDir "manifest.json"
    $slidesData | ConvertTo-Json -Depth 3 | Out-File -FilePath $manifestPath -Encoding UTF8

    Write-Output "SUCCESS: $manifestPath"

} catch {
    Write-Error "Error converting PowerPoint: $_"
    exit 1
}
