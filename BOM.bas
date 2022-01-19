Option Explicit



Sub aSetKey()
    Application.OnKey "{@}", "gotoValue"
    
End Sub


Sub gotoValue()
    
    Dim FindString As String            'Declare var for textbox
    Dim Rng As Range                    'Declare var for finding feeder name
    FindString = InputBox("Value")        'Get part number
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
                    
                End Select
            End If
        End With
    End If
    
    Dim Feeder As String
    Dim Feede As String
    Dim Feed As String
    
    Feeder = InputBox("Feeder")   'get feeder from scan
    Feede = Right(Feeder, Len(Feeder) - 2)  'delete @ and ~
    If Len(Feede) = 1 And Feede = "1" Then  'If length is 1 char, and the char is 1, do nothing (exit/cancel)
    ElseIf Len(Feede) = 1 And Feede = "2" Then    'erase when scanning 2
        Application.ActiveCell.Value = ""
    Else
        Feed = (Left(Feede, 1)) 'else, assign feeder letter to Feed
        Feede = Right(Feede, Len(Feede) - 1)    'isolate feeder number
        Feede = Feede - 1                       'subtract 1; feeder is 27, cell is a28, so qr reads 28....
        Select Case Feed    'select feeder based on letter, recreate
        Case "B"
            Feede = "B" & Feede
        Case "D"
            Feede = "D" & Feede
        Case "G"
            Feede = "G" & Feede
        Case "R"
            Feede = Feede + 1
            Application.Goto Rng.Offset(0, 3), True 'Select Rotation column if @~R selected
        End Select
        Application.ActiveCell = Feede  'write cell with feeder / rotation
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
    DPos = InStrRev(BOMPath, "Workflow")
    FeederPath = Left(BOMPath, DPos - 1) & "Workflow/Shared Documents/General/" & FeedFile
    
    
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
                        Application.ActiveCell.Offset(0, 2) = Date
                    End If
                End With
            End If
        Next
        
        
        
    End With
    
    
    
End Sub




Sub UpdatePartNumbersByFeeder()
    
    Dim CsvBOMPath As String
    Dim WBOMFile As String
    Dim CsvBOMFile As String
    Dim CsvPOSFile As String
    Dim SlashPos As Long
    Dim VPPPath As String
    Dim VPPFile As String
    Dim WDir As String
    Dim POSFileName As String
    Dim textline As String, text As String, LocPOSeq As Integer, LocBOMeq As Integer, LocPOScsv As Integer, LocBOMcsv As Integer
    
    
    
    WBOMFile = ThisWorkbook.Name
    
    
    MsgBox "In the following window, please find and select the VisualPlace file used for this job."
    With Application.FileDialog(msoFileDialogFilePicker)
'Opens dialog with path of this workbook
    .InitialFileName = Application.ThisWorkbook.Path
'Makes sure the user can select only one file
        .AllowMultiSelect = False
'Filter to just the following types of files to narrow down selection options
        .Filters.Add "vpp", "*.vpp", 1
'Show the dialog box
        .Show
'Store in fullpath variable
        VPPPath = .SelectedItems.Item(1)
    End With
    
    'Get paths and name of VPP
    SlashPos = InStrRev(VPPPath, "/") + InStrRev(VPPPath, "\")
    VPPFile = Right(VPPPath, Len(VPPPath) - SlashPos)
    
    'gets and reads text from vpp file
    Open VPPPath For Input As #1
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline
    Loop
    Close #1
    
    'gets locations of pos and bom file name strings from vpp text
    LocPOSeq = InStr(text, "pos-top=")
    LocBOMeq = InStr(text, "bom=")
    'gets csv or each file type as end position
    LocPOScsv = InStr(LocPOSeq, text, ".csv")
    LocBOMcsv = InStr(LocBOMeq, text, ".csv")
   'extracts filenames from vpp text and saves
   CsvBOMFile = Mid(text, LocBOMeq + 4, LocBOMcsv - LocBOMeq)
   CsvPOSFile = Mid(text, LocPOSeq + 8, LocPOScsv - LocPOSeq - 4)
   POSFileName = Left(CsvPOSFile, Len(CsvPOSFile) - 8)
   
    'get working directory by subtracting length of vpp file name inside vpp file path
    WDir = Left(VPPPath, Len(VPPPath) - Len(VPPFile))
    
    

    
    
    Dim ValueWBOM As Variant
    Dim ValueCsvBOM As Range
    Dim ValueCsvPOS As Range
    Dim LastRow As Long
    Dim i As Integer
    Dim FeederLoc As String
    Dim FeederType As String
    Dim BotCell As Range, BotRow As Integer, TopSel As String, BotSel As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    
    
    
    
    Workbooks.Open (WDir & CsvBOMFile)
    Workbooks.Open (WDir & CsvPOSFile)
    With Workbooks(WBOMFile).Worksheets("Sheet1")                       'On sheet1, only sheet in file
        LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row                'get last row containing value in WBOM
        For i = 2 To LastRow Step 1
            Set ValueWBOM = Workbooks(WBOMFile).Worksheets("Sheet1").Range("C" & i)    'Loooping through BOM feeder column
            FeederLoc = ValueWBOM.Offset(0, 5).Value                                   'get FeederLoc
            FeederType = Left(FeederLoc, 1)                                            'Look at prefix of feeder
            If FeederLoc = "" Then                                                        'If it's nothing, skip
            ElseIf FeederType = "D" Or FeederType = "S" Then
                'If it's T or S
                With Workbooks(CsvBOMFile).ActiveSheet.Range("A:I")
                'Search csvBom for value
                    Set ValueCsvBOM = .Find(What:=ValueWBOM.Value, _
                    After:=.Cells(.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)                               'Find and select the matching cell, looking at values of entire cells by rows without case
                    If Not ValueCsvBOM Is Nothing Then              'If found, append S- or T- prefix based on feeder type
                        Select Case FeederType
                        Case "S"
                            ValueCsvBOM.Value = "S-" & ValueCsvBOM.Value
                            'Open csv file and replace all matching part names with S- prefix
                            Workbooks(CsvPOSFile).ActiveSheet.Range("B:B").Replace _
                            What:=ValueWBOM.Value, Replacement:="S-" & ValueWBOM.Value, _
                            SearchOrder:=xlByColumns, MatchCase:=True
                        Case "D"
                            ValueCsvBOM.Value = "T-" & ValueCsvBOM.Value
                            'Open csv file and repalce all matching part names with T- prefix
                            Workbooks(CsvPOSFile).ActiveSheet.Range("B:B").Replace _
                            What:=ValueWBOM.Value, Replacement:="T-" & ValueWBOM.Value, _
                            SearchOrder:=xlByColumns, MatchCase:=True
                        End Select
                    End If
                End With
            End If
        Next
    End With

Workbooks(CsvPOSFile).Save
Workbooks(CsvBOMFile).Close SaveChanges:=True

    With Workbooks(CsvPOSFile).ActiveSheet
        LastRow = .Cells(.Rows.Count, "B:B").End(xlUp).Row
        'Find first Bottomside row, to split the file with
        Set BotCell = .Range("F:F").Find(What:="Bottom")
        If Not BotCell Is Nothing Then
            BotRow = BotCell.Row
        Else
            BotRow = LastRow + 1
        End If
        'Select rows for topside and bottomside
        TopSel = "A6:H" & BotRow - 1
        BotSel = "A" & BotRow & ":H" & LastRow
        'copy topselection
        .Range(TopSel).Select
        Selection.Copy
        'create new sheet temporarily
        Sheets.Add.Name = "Top"
        Sheets("Top").Select
        ActiveSheet.Paste
        'save active sheet
        ActiveWorkbook.SaveAs Filename:=WDir & POSFileName & "T.txt", FileFormat:=xlText, CreateBackup:=False
        'delete active sheet
        ActiveWindow.SelectedSheets.Delete
        'If there is 2 sides:
        If BotRow <> LastRow Then
            'select bottom part of file, repeat procedure above
            .Range(BotSel).Select
            Selection.Copy
            Sheets.Add.Name = "Bot"
            Sheets("Bot").Select
            ActiveSheet.Paste
            ActiveWorkbook.SaveAs Filename:=WDir & POSFileName & "B.txt", FileFormat:=xlText, CreateBackup:=False
            ActiveWindow.SelectedSheets.Delete
        End If
        'turn on screen
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
    End With
    
    ActiveWorkbook.Close SaveChanges:=False
 
    MsgBox "Success - Feeder Prefixes Added, PnP Files Created"
    
End Sub

