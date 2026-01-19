param(
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    [Parameter(Mandatory=$true)]
    [string]$NotesJsonPath
)

try {
    Write-Output "DEBUG: InputPath: '$InputPath'"
    Write-Output "DEBUG: NotesJsonPath: '$NotesJsonPath'"

    if (-not (Test-Path -LiteralPath $InputPath)) {
        Write-Error "ERROR: PPTX file not found."
        exit 1
    }
    if (-not (Test-Path -LiteralPath $NotesJsonPath)) {
        Write-Error "ERROR: Notes JSON file not found."
        exit 1
    }

    # Load JSON data
    $jsonContent = Get-Content -LiteralPath $NotesJsonPath -Raw -Encoding UTF8
    $updates = $jsonContent | ConvertFrom-Json

    $ppt = New-Object -ComObject PowerPoint.Application
    # Open presentation (Visible=True usually safer for write ops in some versions, but False is preferred for background)
    # Using msoFalse for Window argument
    $presentation = $ppt.Presentations.Open($InputPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

    foreach ($update in $updates) {
        $slideIndex = $update.index
        $newNotes = $update.notes
        
        try {
            # indices are 1-based in COM
            $slide = $presentation.Slides.Item($slideIndex)
            
            # Ensure Notes Page exists (usually it does, but just in case)
            # Accessing NotesPage creates one if it doesn't exist? Actually HasNotesPage is read/write in some contexts but read-only in others. 
            # Usually accessing NotesPage property is enough.
            
            # Shape 2 is standard for the body text on the notes master/page
            # We might need to iterate shapes to find the body placeholder if it's not index 2, but 2 is the standard convention.
            $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $newNotes
            
            Write-Output "Updated Slide $slideIndex"
        } catch {
            Write-Warning "Failed to update slide $slideIndex : $_"
        }
    }

    $presentation.Save()
    $presentation.Close()
    $ppt.Quit()

    # Cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Output "SUCCESS"

} catch {
    Write-Error "CRITICAL ERROR: $_"
    exit 1
}
