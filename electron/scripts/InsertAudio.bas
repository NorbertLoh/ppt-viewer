Attribute VB_Name = "AudioTools"

' ==============================================================================================
' INSTRUCTIONS FOR USER:
' 1. Open PowerPoint.
' 2. Press Alt+F11 (or Fn+Opt+F11 on Mac) to open the VBA Editor.
' 3. File -> Remove Module (if previous one exists).
' 4. File -> Import File... -> Select this NEW "InsertAudio.bas" file.
' 5. Go to File -> Save As... -> Save as PowerPoint Add-in (.ppam) -> Overwrite previous "AudioTools.ppam".
' 6. Restart PowerPoint to ensure the new Add-in is loaded.
' ==============================================================================================

Sub InsertAudio()
    Dim sld As Slide
    Dim shp As Shape
    Dim pres As Presentation
    Dim paramsPath As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim params() As String
    Dim slideIndex As Integer
    Dim audioPath As String
    
    ' Path to the parameters file in the Office Group Container
    ' Expanding the tilde ~ in VBA is not automatic, so we hardcode the likely path structure or rely on AppleScript passing full path?
    ' Actually, AppleScript will write to the Office sandbox.
    ' VBA on Mac sandboxing usually maps strictly.
    ' Let's use the standard Mac Office sandbox path.
    
    ' Note: On Mac VBA, getting the user home folder correctly can be tricky with sandboxing.
    ' However, generic file reading from the UBF8T346G9.Office folder is usually allowed.
    
    ' We will rely on reading from a fixed location that we know we can write to.
    ' Mac Path: /Users/<User>/Library/Group Containers/UBF8T346G9.Office/audio_params.txt
    
    ' Function to get Home Dir in Mac VBA?
    ' Environ("HOME") often points to the container, not the real user home.
    ' The "UBF8T346G9.Office" folder is effectively the shared root.
    
    ' Let's try to find the file by standard Mac path.
    ' AppleScript will write to: ~/Library/Group Containers/UBF8T346G9.Office/audio_params.txt
    ' VBA should be able to read that if we construct the path correctly.
    
    ' A safe way is to let AppleScript invoke the macro, and we assume the file is there.
    
    On Error Resume Next
    
    ' Attempt to locate existing file
    ' Use MacScript to resolve path? No, deprecated/blocked.
    
    ' Hardcoded check for standard location relative to where we are? No.
    ' Let's try expanding the home dir via a small trick or just assuming standard structure if Environ works.
    ' If Environ("HOME") returns /Users/username/Library/Containers/..., that's the app sandbox.
    ' We need the Group Container.
    
    ' Standard Path construction:
    ' We use the real user path because the Group Container is outside the App Sandbox Data folder
    paramsPath = "/Users/" & Environ("USER") & "/Library/Group Containers/UBF8T346G9.Office/audio_params.txt"
    
    ' Check if file exists
    If Dir(paramsPath) = "" Then
        MsgBox "Error: Could not find audio_params.txt at " & paramsPath
        Exit Sub
    End If
    
    ' Read the file
    fileNum = FreeFile
    Open paramsPath For Input As fileNum
    Line Input #fileNum, fileContent
    Close fileNum
    
    ' Expected content: "SlideIndex|AudioPath|PresentationPath"
    params = Split(fileContent, "|")
    
    If UBound(params) < 2 Then
        Exit Sub
    End If
    
    slideIndex = CInt(params(0))
    audioPath = params(1)
    Dim targetPath As String
    targetPath = params(2)
    
    ' Find the correct presentation
    Dim p As Presentation
    Set pres = Nothing
    
    For Each p In Application.Presentations
        ' Check FullName (absolute path) or Name (filename)
        ' Mac paths can be tricky (HFS vs POSIX), so checking if ends with Name is safer, or fuzzy match.
        ' Let's check if the FullName contains our target filename distinctively.
        ' Or better, if we passed the full path, check for exact match or converted match.
        
        If p.FullName = targetPath Or p.Name = Dir(targetPath) Then
            Set pres = p
            Exit For
        End If
    Next p
    
    ' Fallback: if not found by strict path, maybe it is the active one if the user is looking at it?
    ' But the user said "wrong powerpoint", so ActivePresentation might be the Add-in itself (usually hidden) or another doc.
    If pres Is Nothing Then
        ' Try to find by filename only
        Dim targetName As String
        targetName = Right(targetPath, Len(targetPath) - InStrRev(targetPath, "/"))
        
        For Each p In Application.Presentations
            If p.Name = targetName Then
                Set pres = p
                Exit For
            End If
        Next p
    End If
    
    If pres Is Nothing Then
        MsgBox "Error: Could not find open presentation: " & targetPath
        Exit Sub
    End If
    
    If slideIndex > pres.Slides.Count Or slideIndex < 1 Then Exit Sub
    
    Set sld = pres.Slides(slideIndex)
    
    ' AddMediaObject2(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
    ' Insert at 0,0 first to ensure creation, then move.
    Set shp = sld.Shapes.AddMediaObject2(audioPath, 0, -1, 0, 0, 50, 50)
    
    If Not shp Is Nothing Then
        shp.Name = "GeneratedAudio_" & Format(Now, "hhmmss")
        
        ' Move off-screen explicitly
        shp.Left = -100
        shp.Top = -100
        
        ' Animation Logic: Play After Previous
        Dim eff As Effect
        ' msoAnimEffectMediaPlay = 83
        ' msoAnimTriggerAfterPrevious = 3
        Set eff = sld.TimeLine.MainSequence.AddEffect(shp, 83, , 3)
        
        ' Move to first position so it plays immediately when slide starts
        eff.MoveTo 1
    End If
    
End Sub
