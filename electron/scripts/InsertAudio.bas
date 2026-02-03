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

Sub UpdateNotes()
    Dim pres As Presentation
    Dim p As Presentation
    Dim paramsPath As String
    Dim dataPath As String
    Dim targetPath As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim params() As String
    
    ' 1. Read Parameters (Presentation Path | Data File Path)
    paramsPath = "/Users/" & Environ("USER") & "/Library/Group Containers/UBF8T346G9.Office/notes_params.txt"
    
    If Dir(paramsPath) = "" Then
        MsgBox "Error: Could not find notes_params.txt"
        Exit Sub
    End If
    
    fileNum = FreeFile
    Open paramsPath For Input As fileNum
    Line Input #fileNum, fileContent
    Close fileNum
    
    params = Split(fileContent, "|")
    If UBound(params) < 1 Then Exit Sub
    
    targetPath = params(0)
    dataPath = params(1)
    
    ' 2. Find Presentation
    Set pres = Nothing
    For Each p In Application.Presentations
        If p.FullName = targetPath Or p.Name = Dir(targetPath) Then
            Set pres = p
            Exit For
        End If
    Next p
    
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
        MsgBox "Error: Presentation not found: " & targetPath
        Exit Sub
    End If
    
    ' 3. Read Data File
    If Dir(dataPath) = "" Then
        MsgBox "Error: Data file not found: " & dataPath
        Exit Sub
    End If
    
    Dim dataNum As Integer
    Dim lineData As String
    Dim currentSlideIndex As Integer
    Dim currentNotes As String
    Dim isReadingNotes As Boolean
    
    currentSlideIndex = -1
    isReadingNotes = False
    
    dataNum = FreeFile
    Open dataPath For Input As dataNum
    
    Do While Not EOF(dataNum)
        Line Input #dataNum, lineData
        
        If Left(lineData, 17) = "###SLIDE_START###" Then
            ' Format: ###SLIDE_START### <index>
            currentSlideIndex = CInt(Mid(lineData, 19))
            currentNotes = ""
            isReadingNotes = True
        ElseIf Left(lineData, 15) = "###SLIDE_END###" Then
            If currentSlideIndex > 0 And currentSlideIndex <= pres.Slides.Count Then
                ' Apply notes to slide
                On Error Resume Next
                pres.Slides(currentSlideIndex).NotesPage.Shapes(2).TextFrame.TextRange.Text = currentNotes
                On Error GoTo 0
            End If
            isReadingNotes = False
        Else
            If isReadingNotes Then
                If currentNotes = "" Then
                    currentNotes = lineData
                Else
                    currentNotes = currentNotes & vbCrLf & lineData
                End If
            End If
        End If
    Loop
    
    Close dataNum
    
    ' 4. Save
    pres.Save
    pres.Close
    
End Sub
