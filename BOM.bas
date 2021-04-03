Attribute VB_Name = "Module1"
Option Explicit



Sub aSetKey()

    Application.OnKey "{@}", "gotoCell"
    
End Sub


Sub gotoCell()
    
    Dim FindString As String            'Declare var for textbox
    Dim Rng As Range                    'Declare var for finding feeder name
    FindString = InputBox("Val")        'Get part number
    If Trim(FindString) <> "" Then      'Check if part number empty
        With Worksheets("Sheet1").Range("C:C")                  'Using the Part number range
            Set Rng = .Find(What:=FindString, _
            After:=.Cells(.Cells.Count), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)       'set range based on what's found after looking through the rows
            If Not Rng Is Nothing Then          'Check something has been found, move to feeder column if true
                Application.Goto Rng.Offset(0, 5), True
                Beep
            Else
                Dim AckTime As Integer, InfoBox As Object   'If not found, nothing found message
                Set InfoBox = CreateObject("WScript.Shell")
                AckTime = 1
                Select Case InfoBox.Popup("Click OK or do nothing.", _
                    AckTime, "Nothing Found", 0)
                    
                End If
            End With
        End If
        
        Dim Feeder As String
        Dim Feede As String
        Dim Feed As String
        Feeder = InputBox("Feed")   'get feeder from scan
        Feede = Right(Feeder, Len(Feeder) - 2)  'delete @ and ~
        If Len(Feede) = 1 And Feede = "1" Then  'If length is 1 char, and the char is 1, do nothing (exit/cancel)
            
        Else
            Feed = (Left(Feede, 1)) 'else, select feeder first letter and recreate based on which letter
            Feede = Right(Feede, Len(Feede) - 1)
            Feede = Feede - 1
            Select Case Feed
            Case "B"
                Feede = "B" & Feede
            Case "D"
                Feede = "D" & Feede
            Case "G"
                Feede = "G" & Feede
            Case "R"
                Application.Goto Rng.Offset(0, 3), True 'Select Rotation column if @~R selected
            End Select
            Application.ActiveCell = Feede
            Beep
        End If
        
        
        
        
    End Sub




Sub UpdateFeeder_List()
    
    
    Dim BOMPath As String
    Dim FeederPath As String
    Dim DPos As Long
    Dim WBOMFile As String
    Dim FeedFile As String
    
    FeedFile = "Loaded_Feeders.xlsm"
    WBOMFile = ThisWorkbook.Name
    BOMPath = Application.ThisWorkbook.Path
    DPos = InStrRev(BOMPath, "Desktop")
    FeederPath = Left(BOMPath, DPos - 1) & "Desktop/Jobs/" & FeedFile
    
 
    Workbooks.Open (FeederPath)
    
    Dim FeederBOM As Range
    Dim FeederList As Range
    Dim LastRow As Long
    Dim i As Integer
    
    With Workbooks(WBOMFile).Worksheets("Sheet1")
        LastRow = .Cells(.Rows.Count, "H").End(xlUp).Row
        For i = 2 To LastRow Step 1
            Set FeederBOM = Workbooks(WBOMFile).Worksheets("Sheet1").Range("H" & i)    'Loooping through BOM feeder column
            If FeederBOM.Value = "" Then                                                        'If it's nothing, skip
            Else
                With Workbooks(FeedFile).Worksheets("Sheet1").Range("A:A")         'If it's something, look in feeder list sheet
                    Set FeederList = .Find(What:=FeederBOM.Value, _
                    After:=.Cells(.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)                                                           'Find and select the matching cell
                    If Not FeederList Is Nothing Then
                        Application.Goto FeederList.Offset(0, 3), True                          'Go to feeder value
                        Application.ActiveCell = Workbooks(WBOMFile).Worksheets("Sheet1").Range("C" & i).Value 'Make feeder value the same as BOM
                        Application.ActiveCell.Offset(0, 1) = Workbooks(WBOMFile).Worksheets("Sheet1").Range("D" & i).Value 'Make feeder profile same as BOM
                    End If
                End With
            End If
        Next
       
        
        
    End With
    
    
   
End Sub


