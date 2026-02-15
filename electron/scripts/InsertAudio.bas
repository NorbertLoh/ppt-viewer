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
    
    ' Batch Process Mode
    Dim lineData As String
    Dim targetPath As String
    Dim hasPres As Boolean
    
    hasPres = False
    
    fileNum = FreeFile
    Open paramsPath For Input As fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, fileContent
        
        If Len(Trim(fileContent)) > 0 Then
            params = Split(fileContent, "|")
            
            If UBound(params) >= 2 Then
                slideIndex = CInt(params(0))
                audioPath = params(1)
                
                ' Only find presentation once (on first valid line)
                If Not hasPres Then
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
                        Close fileNum
                        Exit Sub
                    End If
                    
                    hasPres = True
                End If
                
                ' Process Audio Insertion for this line
                If slideIndex > 0 And slideIndex <= pres.Slides.Count Then
                    Set sld = pres.Slides(slideIndex)
                    
                    ' Tag for unique identification
                    Dim audioTag As String
                    audioTag = "ppt_audio_slide_" & slideIndex
                    
                    ' Find and delete existing audio with this tag (Find & Replace)
                    Dim s As Shape
                    For Each s In sld.Shapes
                        If s.Name = audioTag Then
                            s.Delete
                            Exit For
                        End If
                    Next s
                    
                    ' Insert the audio object
                    Set shp = sld.Shapes.AddMediaObject2(audioPath, 0, -1, 10, 10)
                    
                    If Not shp Is Nothing Then
                        shp.Name = audioTag
                        shp.Left = -100
                        shp.Top = -100
                        
                        ' --- Animation Configuration ---
                        Dim eff As Effect
                        
                        ' 1. Ensure clean slate (remove any auto-added effects for this shape)
                        Dim i As Integer
                        For i = sld.TimeLine.MainSequence.Count To 1 Step -1
                            If Not sld.TimeLine.MainSequence(i).Shape Is Nothing Then
                                If sld.TimeLine.MainSequence(i).Shape.Name = shp.Name Then
                                    sld.TimeLine.MainSequence(i).Delete
                                End If
                            End If
                        Next i
                        
                        ' 2. Add "Play" effect to Main Sequence
                        ' msoAnimEffectMediaPlay = 83 in some versions, but better to rely on Enum or just add it.
                        ' Using the named constant if available, else assuming standard MediaPlay behavior.
                        ' 3 = msoAnimTriggerAfterPrevious
                        Set eff = sld.TimeLine.MainSequence.AddEffect(shp, 83, , 3) 
                        
                        ' 3. Move to Front (Make it the first animation)
                        Do While eff.Index > 1
                            eff.MoveTo 1
                        Loop
                        
                        ' 4. Remove any "Trigger" (Interactive Sequence) created by PPT defaults
                        ' (Often PPT adds an OnClick trigger for media)
                        ' Note: sld.TimeLine.InteractiveSequences contains trigger-based animations
                        ' We iterate via property accessor or just by count if needed, but usually AddEffect logic is enough 
                        ' if we didn't use "AddMediaObject" with "LinkToFile:=False, SaveWithDocument:=True" etc which sometimes auto-triggers.
                        ' But AddMediaObject2 is generally cleaner.
                        ' We will assume the MainSequence addition is sufficient, but let's double check MainSequence order.
                        
                        With shp.MediaFormat
                            .Muted = False
                            .Volume = 0.5
                        End With
                    End If
                End If
            End If
        End If
    Loop
    
    Close fileNum
    
    ' Save ONCE after batch processing
    If hasPres Then
        pres.Save
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
    ' pres.Close
    
End Sub
