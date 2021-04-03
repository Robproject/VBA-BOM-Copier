Attribute VB_Name = "Module1"

Option Explicit

Sub aSetKey()

    Application.OnKey "{@}", "gotoCell"
    
End Sub


Sub gotoCell()

Dim FindString As String
Dim Rng As Range
FindString = InputBox("Val")
If Trim(FindString) <> "" Then
    With Worksheets("Sheet1").Range("A:A")
        Set Rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                Application.Goto Rng.Offset(0, 4), True
                Beep
            Else
                 Dim AckTime As Integer, InfoBox As Object

                Set InfoBox = CreateObject("WScript.Shell")
                AckTime = 1
                Select Case InfoBox.Popup("Click OK or do nothing.", _
                AckTime, "Nothing Found", 0)

 Case 1, -1
  Exit Sub
 End Select
                
            End If
    End With
End If

Dim Feeder As String
Dim Feede As String
Dim Feed As String
Feeder = InputBox("Feed")
Feede = Right(Feeder, Len(Feeder) - 2)
If Len(Feede) = 1 And Feede = "1" Then

Else
Feed = (Left(Feede, 1))
Select Case Feed
    Case "B"
    Feede = Right(Feede, Len(Feede) - 1)
    Feede = Feede - 1
    Feede = "B" & Feede
    Case "D"
    Feede = Right(Feede, Len(Feede) - 1)
    Feede = Feede - 1
    Feede = "D" & Feede
    Case "G"
    Feede = Right(Feede, Len(Feede) - 1)
    Feede = Feede - 1
    Feede = "G" & Feede
    Case "R"
    Feede = Right(Feede, Len(Feede) - 1)
    Application.Goto Rng.Offset(0, 3), True
End Select
Application.ActiveCell = Feede
Beep
End If




End Sub

