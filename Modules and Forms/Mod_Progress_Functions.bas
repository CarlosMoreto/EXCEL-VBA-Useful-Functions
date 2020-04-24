Attribute VB_Name = "Mod_Progress_Functions"
'Option Private Module

Private progPos As Double
Private progDist As Double
Private winWidth As Double
Private imgWidth As Double
Private numUpdates As Long
Private updateCount As Long

'Public Const lblColors = Array(vbBlack, 15921906)

Sub Progress_Initialize()
    'Export_Modules
    Image_Initialize
    
    With Form_Progress
        .Img_Bug.Visible = True
        .Img_Vision.Visible = True
        '.Caption = "VISION - em progresso..."
    End With
    
    Form_Progress.Show (vbModeless)
End Sub

Public Sub Image_Initialize(Optional ByVal Image As Boolean = True)
    numUpdates = 10
    updateCount = 0
    
    If Image = True Then
        winWidth = val(Replace(Form_Progress.Width, ",", "."))
        imgWidth = val(Replace(Form_Progress.Img_Bug.Width, ",", "."))
        progPos = -0.91 * imgWidth
        progDist = 0.91 * imgWidth
    End If
End Sub

Public Sub Progress_Update(Optional ByVal waitLabel As String = "Aguarde...", _
                           Optional ByVal maxProg As Double = 0.91, _
                           Optional ByVal clrLabel As Boolean = True)
    
    updateCount = updateCount + 1
    progDist = maxProg * winWidth
    oldPos = progPos
    progPos = progPos + progDist / numUpdates
    
    'waitLabel = "Aguarde..."
    
    If progPos >= (maxProg * winWidth - imgWidth) Then progPos = maxProg * winWidth - imgWidth
    
    With Form_Progress
'        thePos = oldPos
'
'        For i = 1 To 10
'            thePos = thePos + progPos / 10
'            .Img_Bug.Left = thePos
'
'            Sleep 50
'        Next
        
        .Img_Bug.Left = progPos
        '.L_Wait.Visible = Not .Label1.Visible
        
        If clrLabel Then
            .L_Wait.Caption = waitLabel
        Else
            .L_Wait.Caption = .L_Wait.Caption
        End If
    End With
    
    DoEvents
End Sub

Public Sub Progress_Almost_Terminate()
    numUpdates = 20
    sleepTime = 20
    progLeft = 0.91 * winWidth - progPos
    progOS = progLeft / numUpdates
    
    For i = 1 To numUpdates
        Form_Progress.Img_Bug.Left = Form_Progress.Img_Bug.Left + progOS
        DoEvents
        Sleep sleepTime
    Next
End Sub

Public Sub Progress_Terminate()
    Unload Form_Progress
End Sub

Public Function Get_File_Num_Lines(ByVal thePath As String) As Integer
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set theText = objFSO.OpenTextFile(thePath, ForReading)

    'Skip lines one by one
    Do While theText.AtEndOfStream <> True
        theText.SkipLine ' or strTemp = txsInput.ReadLine
    Loop
    
    numLines = theText.Line - 1
    
    'Cleanup
    Set objFSO = Nothing
    Set theText = Nothing
    
    Get_File_Num_Lines = numLines
End Function

Public Sub Progress_Set_Update_Count(ByVal updateCount As Long, _
                                     Optional ByVal increment As Boolean = True)
    If increment Then
        numUpdates = numUpdates + updateCount
    Else
        numUpdates = updateCount
    End If
End Sub

