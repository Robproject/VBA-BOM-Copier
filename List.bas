Attribute VB_Name = "Module1"
Option Explicit

Sub aSetKey()
    
    Application.OnKey "{@}", "gotoFeeder"
    
End Sub


Sub gotoFeeder()
    
    Dim FindString As String                    'Declare input box value
    Dim Rng As Range                            'Declare rng value
    FindString = InputBox("Val")                'Assign value from box
    If Trim(FindString) <> "" Then              'If value isn't empty
    FindString = Right(FindString, Len(FindString) - 1)  'remove leftmost character
        With Worksheets("Sheet1").Range("A:A")  'Find value read from input box, in column A
            Set Rng = .Find(What:=FindString, _
            After:=.Cells(.Cells.Count), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)
            If Not Rng Is Nothing Then          'If not empty, select cell offset 4 to the right
                Application.Goto Rng.Offset(0, 3), True
                Beep
            Else
                Dim AckTime As Integer, InfoBox As Object   'error box if empty
                Set InfoBox = CreateObject("WScript.Shell")
                AckTime = 1
                Select Case InfoBox.Popup("Click OK or do nothing.", _
                    AckTime, "Nothing Found", 0)
                Case 1, -1
                End Select
            End If
        End With
    End If
    
    Dim Value As String
    Dim LessPreValue As String
    Value = InputBox("Value")       'get part value
    LessPreValue = Right(Value, Len(Value) - 1)     'trim prefix
    If Len(LessPreValue) = 1 And LessPreValue = "1" Then    'if not 1, apply value
        
    Else
        
        Application.ActiveCell = LessPreValue   'write value to cell
        Beep
    End If
    
    
    
    
End Sub







