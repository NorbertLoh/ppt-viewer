param(
    [Parameter(Mandatory=$true)]
    [string]$InputPath,
    [Parameter(Mandatory=$true)]
    [string]$AudioDir,
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

try {
    Write-Output "DEBUG: Starting Video Generation..."
    Write-Output "DEBUG: Input: $InputPath"
    Write-Output "DEBUG: AudioDir: $AudioDir"
    Write-Output "DEBUG: Output: $OutputPath"

    if (-not (Test-Path -LiteralPath $InputPath)) {
        Write-Error "ERROR: Input file does not exist: $InputPath"
        exit 1
    }

    $ppt = New-Object -ComObject PowerPoint.Application
    # Open Presentation (Visible=True usually needed for export operations to work reliably, or standard open)
    $presentation = $ppt.Presentations.Open($InputPath)
    
    foreach ($slide in $presentation.Slides) {
        $index = $slide.SlideIndex
        $audioFile = Join-Path $AudioDir "slide_$index.wav"
        
        if (Test-Path $audioFile) {
            Write-Output "DEBUG: Adding audio to slide $index"
            
            # Add Media Object (Audio)
            # AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
            # Move off-screen instead of hiding to prevent playback issues
            # Add Media Object (Audio)
            # Position visible but will be hidden by settings
            $shape = $slide.Shapes.AddMediaObject2($audioFile, 0, 1, 10, 10, 50, 50)
            
            # Hide when not playing
            $shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = -1 # msoTrue
            
            # Use PlayOnEntry to create the standard Media Play effect
            $shape.AnimationSettings.PlaySettings.PlayOnEntry = -1 # msoTrue

            # Now find the effect in the MainSequence that targets this shape
            $sequence = $slide.TimeLine.MainSequence
            $foundEffect = $null
            
            for ($i = 1; $i -le $sequence.Count; $i++) {
                $eff = $sequence.Item($i)
                if ($eff.Shape.Name -eq $shape.Name) {
                    $foundEffect = $eff
                    break # Assuming one effect per audio shape
                }
            }
            
            if ($foundEffect) {
                # Move to the very top (first animation)
                while ($foundEffect.Index -gt 1) {
                    $foundEffect.MoveTo($foundEffect.Index - 1)
                }
                
                # Set Trigger to AFTER PREVIOUS (3)
                # 1 = WithPrevious, 3 = AfterPrevious
                $foundEffect.Timing.TriggerType = 3 
                $foundEffect.Timing.TriggerDelayTime = 0
            }
            
            # IMPORTANT: Set Slide Transition to Advance on Time
            # We need to know the duration of the audio.
            # The Shape.MediaFormat.Length property gives length in ms.
            $lengthMs = $shape.MediaFormat.Length
            $lengthSec = $lengthMs / 1000
            
            Write-Output "DEBUG: Slide $index Audio Length: $lengthSec seconds"
            
            $slide.SlideShowTransition.AdvanceOnTime = -1 # msoTrue
            $slide.SlideShowTransition.AdvanceTime = $lengthSec
        } else {
            Write-Output "DEBUG: No audio for slide $index. Using default transition."
            # Maybe set a default short duration if no audio?
            if ($slide.SlideShowTransition.AdvanceOnTime -eq 0) {
                 $slide.SlideShowTransition.AdvanceOnTime = -1
                 $slide.SlideShowTransition.AdvanceTime = 3 # Default 3 seconds
            }
        }
    }
    
    Write-Output "DEBUG: Starting CreateVideo..."
    # CreateVideo(FileName, UseTimingsAndNarrations, VertResolution, FramesPerSecond, Quality, CompactIsPresentation)
    $presentation.CreateVideo($OutputPath, -1, 4, 1080, 30, 85)
    
    # Wait for completion
    # Status: 0=None, 1=InProgress, 2=Queued, 3=Done, 4=Failed
    Start-Sleep -Seconds 2 # Give it a moment to queue
    
    while ($presentation.CreateVideoStatus -eq 1 -or $presentation.CreateVideoStatus -eq 2) { 
        Start-Sleep -Seconds 2
        Write-Output "DEBUG: Rendering... Status: $($presentation.CreateVideoStatus)"
    }
    
    if ($presentation.CreateVideoStatus -eq 3) { 
         Write-Output "DEBUG: Video created successfully."
    } elseif ($presentation.CreateVideoStatus -eq 4) {
         throw "CreateVideo failed (Status 4)."
    } else {
         # Should not happen unless cancelled or None(0) if never started
         throw "CreateVideo ended with unexpected status: $($presentation.CreateVideoStatus)"
    }
    
    # Close without saving (we modified the open presentation by adding audio shapes)
    $presentation.Close()
    $ppt.Quit()
    
    # Cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Output "SUCCESS"

} catch {
    Write-Error "Error: $_"
    exit 1
}
