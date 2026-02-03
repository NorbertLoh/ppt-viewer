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
    
    ' Path construction for Mac Office sandbox
    paramsPath = "/Users/" & Environ("USER") & "/Library/Group Containers/UBF8T346G9.Office/audio_params.txt"
    
    If Dir(paramsPath) = "" Then
        MsgBox "Error: Could not find audio_params.txt at " & paramsPath
        Exit Sub
    End If
    
    fileNum = FreeFile
    Open paramsPath For Input As fileNum
    Line Input #fileNum, fileContent
    Close fileNum
    
    params = Split(fileContent, "|")
    If UBound(params) < 2 Then Exit Sub
    
    slideIndex = CInt(params(0))
    audioPath = params(1)
    Dim targetPath As String
    targetPath = params(2)
    
    ' Find the correct presentation
    Dim p As Presentation
    Set pres = Nothing
    For Each p In Application.Presentations
        If p.FullName = targetPath Or p.Name = Dir(targetPath) Then
            Set pres = p
            Exit For
        End If
    Next p
    
    ' Fallback to name search
    If pres Is Nothing Then
        Dim targetName As String
        targetName = Mid(targetPath, InStrRev(targetPath, "/") + 1)
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
    
    ' Insert the audio object
    ' LinkToFile: False (0), SaveWithDocument: True (-1)
    Set shp = sld.Shapes.AddMediaObject2(audioPath, 0, -1, 10, 10)
    
    If Not shp Is Nothing Then
        shp.Name = "GeneratedAudio_" & Format(Now, "hhmmss")
        
        ' 1. Position it off-screen
        shp.Left = -100
        shp.Top = -100
        
        ' 2. Configure Media Settings (No Animation Trigger)
        With shp.MediaFormat
            ' This ensures the icon isn't visible if it were on-screen
            ' and removes the need for interactive triggers
            .Muted = False
            .Volume = 0.5
        End With
        
        ' 3. Ensure no Play effects are added to the timeline
        ' We simply skip the sld.TimeLine.MainSequence.AddEffect block entirely.
    End If
    
End Sub
