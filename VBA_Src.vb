Public streamWB As Workbook
Public streamOverall As Worksheet
Public modelWB As Workbook
Public modelOverall As Worksheet
Function main()

    'MsgBox "Initializing Setup..."

    Debug.Print ActiveWorkbook.name
    Debug.Print ThisWorkbook.name

    'Dim modelsWB As Workbook

    'Streams
    Call Streams
    'Models
    Call Models

    ActiveWindow.WindowState = xlMinimized

    MsgBox "You can now find the worksheets minimized in the lower left corner."

End Function
Sub Models()

    Dim filename As String
    Dim sheetname As String
    filename = Init_Models()

    'Gets Paste Data from txt file provided
    Path = Application.ActiveWorkbook.Path
    Debug.Print Path

    Dim arr() As String
    arr = getText()


    modelWB.Worksheets("Aspen Data Tables").Copy After:=Worksheets(Sheets.count)
    On Error Resume Next
        ActiveSheet.name = "Overall"

    Set modelOverall = modelWB.Worksheets("Overall")
    'modelOverall.Cells(50, 1).Value = "hi"

    'Overall
    'sheetname = "Overall"
    'sheetname = CopySheetAndRename()
    'Worksheets("Aspen Data Tables").Copy After:=Worksheets(Sheets.count)
    'Call CopySheetAndRenamePredefined
    'ActiveSheet.name = sheetname
    'Call A64_A108_Data(filename, sheetName)

    'Worksheets("Overall").Activate

    Call freezePanes("B4")
    'Debug.Print "calling setupBlockDataCells..."
    Call setupBlockDataCells(arr)
    ActiveWindow.WindowState = xlMinimized


End Sub
Sub Streams()

    Dim filename As String
    filename = Init_Streams()
    'Dim sheetname As String

    'Aspen Data Tables Modified
    'sheetname = CopySheetAndRename

    Worksheets("Aspen Data Tables").Copy After:=Worksheets(Sheets.count)
    On Error Resume Next
        ActiveSheet.name = "Aspen Data Tables Modified"
    'Debug.Print "sheetname: ", sheetname

    Call DeleteRowsBelow
    Call DeleteRowsAbove
    Call removeExtraRows
    Call addRowEntropyFlow
    Call addRowExergyFlow
    Call freezePanes("C7")

    'Overall
    Worksheets("Aspen Data Tables Modified").Copy After:=Worksheets(Sheets.count)
    On Error Resume Next
        ActiveSheet.name = "Overall"
    'Call CopySheetAndRename
    Call removeToFromColumns
    Call calc_balances

    Set streamOverall = streamWB.Worksheets("Aspen Data Tables Modified")
    ActiveWindow.WindowState = xlMinimized

End Sub
Function Init_Streams() As String

    MsgBox ("Please select a file to be opened for STREAMS (click ok first)")

    Dim filename As Variant

    filename = Application.GetOpenFilename(FileFilter:="Excel Files, *.xl*;*.xm*")
    '  “xl Files,*.xl*,xm Files,*.xm*” if you want separate file types in dialogue box

    If filename <> False Then
        'Workbooks.Open filename:=filename
        Set streamWB = Workbooks.Open(filename)
        On Error Resume Next
    End If

    filename = Mid(filename, InStrRev(filename, "\") + 1)
    Debug.Print "streams filename: ", filename
    'Set streamOverall = ActiveWorkbook.Worksheets(filename)

End Function
Function Init_Models() As String

    MsgBox ("Please select a file to be opened for MODELS (click ok first)")

    Dim filename As Variant

    filename = Application.GetOpenFilename(FileFilter:="Excel Files, *.xl*;*.xm*")
    '  “xl Files,*.xl*,xm Files,*.xm*” if you want separate file types in dialogue box

    If filename <> False Then
        'Workbooks.Open filename:=filename
        Set modelWB = Workbooks.Open(filename)
    End If

    Debug.Print "FILENAME: ", filename
    filename = Mid(filename, InStrRev(filename, "\") + 1)
    Debug.Print "name_only: ", filename

    Init_Models = filename

End Function
Sub checkEnthalpyUnits(ByVal row As Integer)

    Debug.Print "units: " & Cells(row, 2)
    Dim lastCol, i As Integer
    lastCol = Cells(row, Columns.count).End(xlToLeft).Column
    Debug.Print "lastCol: " & lastCol

    If Cells(row, 2) = "Watt" Then
        For i = 3 To lastCol
            Debug.Print "value: " & Cells(row, i)
            Cells(row, i) = Cells(row, i) / (10 ^ 6)
        Next i
    End If
    Cells(row, 2) = "MW"

End Sub
Public Sub CopySheetAndRenamePredefined()
    ActiveSheet.Copy After:=Worksheets(Sheets.count)
    On Error Resume Next
    ActiveSheet.name = "Overall"
End Sub

Sub setupBlockDataCells(arr)

    'modelWB.Worksheets("modelOverall").Activate

    Dim cell As Range
    Dim blockType As String

    For Each cell In Range("3:3")

        'Debug.Print "before If cell.value = name"

        If cell.Value = "Name" Then

            'Debug.Print "after!"
            'Fills Data Block Values
            Call PrintArray(64, cell.Column, arr)

            'Sets Block Type
            blockType = Cells(2, cell.Column).Value
            Cells(64, cell.Column + 1) = blockType

            'Sets Temperature Row
            If blockType = "Heater" Then
                Cells(111, cell.Column) = "Temperature, K"
            End If

            'Sets Heat IN/OUT rows
            If blockType = "RadFrac" Then
                Cells(111, cell.Column) = "Heat In, MW"
                Cells(112, cell.Column) = "Heat OUT, MW"
            End If

            'Sets each Block Name
            Dim i As Integer
            Dim blockname As Variant
            i = 1

            'Loops through each blockname for each block type
            'Inside each loop is run for EVERY blockname
            Do While (IsEmpty(Cells(3, cell.Column + i).Value)) <> True

                blockname = Cells(3, cell.Column + i).Value
                Cells(65, cell.Column + i).Value = blockname

                'for each block name
                'Debug.Print "before call, blockname: ", blockname

                Call getBlockNameData_To(blockname, cell.Column + i)
                Call getBlockNameData_From(blockname, cell.Column + i)

                If blockType = "Pump" Then
                    If Cells(34, cell.Column) = "Net work required [Watt]" Then
                        Cells(106, cell.Column + i) = Cells(34, cell.Column + i) / 1000000
                    Else
                        Cells(106, cell.Column + i) = Cells(34, cell.Column + i)
                    End If
                End If

                If blockType = "Compr" Then
                    If Cells(21, cell.Column) = "Net work required [Watt]" Then
                        Cells(106, cell.Column + i) = Cells(21, cell.Column + i) / 1000000
                    Else
                        Cells(106, cell.Column + i) = Cells(21, cell.Column + i)
                    End If
                End If

                If blockType = "RadFrac" Then
                    If Cells(27, cell.Column) = "Condenser / top stage heat duty [Watt]" Then
                        Cells(107, cell.Column + i) = (Cells(33, cell.Column + i) + Cells(27, cell.Column + i)) / 1000000
                        Cells(111, cell.Column + i) = Cells(33, cell.Column + i) / 1000000
                        Cells(112, cell.Column + i) = Cells(27, cell.Column + i) / 1000000
                    Else
                        Cells(107, cell.Column + i) = Cells(33, cell.Column + i) + Cells(27, cell.Column + i)
                        Cells(111, cell.Column + i) = Cells(33, cell.Column + i)
                        Cells(112, cell.Column + i) = Cells(27, cell.Column + i)
                    End If
                End If

                If blockType = "Heater" Then
                    If Cells(18, cell.Column) = "Calculated heat duty [Watt]" Then
                        Cells(107, cell.Column + i) = Cells(18, cell.Column + i) / 1000000
                    Else
                        Cells(107, cell.Column + i) = Cells(18, cell.Column + i)
                    End If
                End If

                Call MassBalCalc2(cell.Column + i)
                Call EnergyBalCalc(cell.Column + i)
                Call EntropyBalCalc(cell.Column + i)

                If blockType = "Heater" Then
                    'Cells(107, cell.Column + i) = Cells(18, cell.Column + i)
                    If Cells(16, cell.Column) = "Calculated temperature [K]" Then
                        Cells(111, cell.Column + i) = Cells(16, cell.Column + i)
                    Else
                        Cells(111, cell.Column + i) = Cells(16, cell.Column + i) + 273.15
                    End If
                    Cells(110, cell.Column + i) = Cells(110, cell.Column + i) - (Cells(107, cell.Column + i) / Cells(111, cell.Column + i) * 1000)
                End If

                'finish calc after entropyCalc call
                '-(AX111/(AX32+273.15)+AX112/(AX25+273.15))*1000
                If blockType = "RadFrac" Then
                    Dim temp_sum As Double
                    temp_sum = (Cells(111, cell.Column + i) / (Cells(32, cell.Column + i) + 273.15)) + (Cells(112, cell.Column + i) / (Cells(25, cell.Column + i) + 273.15))
                    temp_sum = temp_sum * 1000
                    Cells(110, cell.Column + i) = Cells(110, cell.Column + i) - temp_sum
                End If

                i = i + 1
            Loop

        End If
    Next cell

End Sub
Sub EntropyBalCalc(ByVal col As Integer)

    Dim sum, in_sum, out_sum As Double
    in_sum = Cells(69, col) + Cells(73, col) + Cells(77, col) + Cells(81, col)
    out_sum = Cells(85, col) + Cells(89, col) + Cells(93, col) + Cells(97, col) + Cells(101, col) + Cells(105, col)
    sum = out_sum - in_sum
    'Debug.Print in_sum, out_sum, sum
    Cells(110, col) = sum

End Sub
Sub EnergyBalCalc(ByVal col As Integer)

    Dim sum, in_sum, out_sum As Double
    in_sum = Cells(68, col) + Cells(72, col) + Cells(76, col) + Cells(80, col) + Cells(106, col) + Cells(107, col)
    out_sum = Cells(84, col) + Cells(88, col) + Cells(92, col) + Cells(96, col) + Cells(100, col) + Cells(104, col)
    sum = out_sum - in_sum
    'Debug.Print in_sum, out_sum, sum
    Cells(109, col) = sum

End Sub
Sub MassBalCalc2(ByVal col As Integer)

    Dim sum, in_sum, out_sum As Double
    in_sum = Cells(67, col) + Cells(71, col) + Cells(75, col) + Cells(79, col)
    out_sum = Cells(83, col) + Cells(87, col) + Cells(91, col) + Cells(95, col) + Cells(99, col) + Cells(103, col)
    sum = out_sum - in_sum
    'Debug.Print in_sum, out_sum, sum
    Cells(108, col) = sum

End Sub
Sub MassBalCalc(ByVal col As Integer)

    Dim sum, in_sum, out_sum As Long
    Dim i As Integer
    For i = 67 To 79 Step 4
        in_sum = in_sum + Cells(i, col)
    Next i
    For i = 83 To 103 Step 4
        out_sum = out_sum + Cells(i, col)
    Next i
    sum = in_sum - out_sum
    Cells(108, col) = sum

End Sub
Function fillBlockName_In(ByVal fill_col As Integer, ByVal get_col As Integer, ByVal count As Integer)

    Dim curRow As Integer
    Dim fill_row As Integer
    fill_row = 66 + (count - 1) * 4

    curRow = FindRowByNameStreams("Stream Name")
    Cells(fill_row, fill_col).Value = streamOverall.Cells(curRow, get_col)

    curRow = FindRowByNameStreams("Mass Flows")
    Cells(fill_row + 1, fill_col).Value = streamOverall.Cells(curRow, get_col)

    curRow = FindRowByNameStreams("Enthalpy Flow")
    Cells(fill_row + 2, fill_col).Value = streamOverall.Cells(curRow, get_col)

    curRow = FindRowByNameStreams("Entropy Flow")
    Cells(fill_row + 3, fill_col).Value = streamOverall.Cells(curRow, get_col)


End Function
Function fillBlockName_Out(ByVal fill_col As Integer, ByVal get_col As Integer, ByVal count As Integer)

    Dim fill_row, curRow As Integer
    fill_row = 82 + (count - 1) * 4

    curRow = FindRowByNameStreams("Stream Name")
    Cells(fill_row, fill_col).Value = streamOverall.Cells(curRow, get_col)

    curRow = FindRowByNameStreams("Mass Flows")
    Cells(fill_row + 1, fill_col).Value = streamOverall.Cells(curRow, get_col)

    curRow = FindRowByNameStreams("Enthalpy Flow")
    Cells(fill_row + 2, fill_col).Value = streamOverall.Cells(curRow, get_col)

    curRow = FindRowByNameStreams("Entropy Flow")
    Cells(fill_row + 3, fill_col).Value = streamOverall.Cells(curRow, get_col)


End Function
Function getBlockNameData_To(ByVal blockname As Variant, ByVal fill_col As Integer)

    'TO row = 6
    Dim count As Integer
    Dim lastCol As Long
    count = 1

    Dim toRow As Integer
    toRow = FindRowByNameStreams("To")
    'Debug.Print "blockname: ", blockname

    lastCol = streamOverall.Cells(10, Columns.count).End(xlToLeft).Column

    'Debug.Print lastCol

    Dim i As Integer
    For i = 3 To lastCol

        'Debug.Print "val: ", streamOverall.Cells(6, i).Value
        If streamOverall.Cells(toRow, i).Value = blockname Then
            'found blockname, now do below for each TO found
            'Debug.Print "Found Match for ", blockname, " in col ", i
            Call fillBlockName_In(fill_col, i, count)
            count = count + 1
        End If
    Next i


End Function
Function getBlockNameData_From(ByVal blockname As Variant, ByVal fill_col As Integer)

    Dim count As Integer
    Dim lastCol As Long
    count = 1

    Dim fromRow As Integer
    fromRow = FindRowByNameStreams("From")
    'Debug.Print "blockname: ", blockname

    lastCol = streamOverall.Cells(10, Columns.count).End(xlToLeft).Column
    'Debug.Print lastCol

    Dim i As Integer
    For i = 3 To lastCol


        If streamOverall.Cells(fromRow, i).Value = blockname Then
            'found blockname, now do below for each TO found
            'Debug.Print "Found Match for ", blockname, " in col ", i
            Call fillBlockName_Out(fill_col, i, count)
            count = count + 1
        End If
    Next i


End Function
'Puts Data from a 1-D Array into cells (row,col) of sheetName
Sub PrintArray(row, col, arr)

    'Debug.Print "printing A64_A108 array..."
    Dim startRow As Integer
    startRow = row

    ' 0 -> 46
    For i = LBound(arr, 1) To UBound(arr, 1)
        modelOverall.Cells(startRow, col).Value = arr(i)
        startRow = startRow + 1
    Next i

End Sub
Sub PrintArray2(Data() As String, Cl As Range)

    Cl.Resize(UBound(Data, 1), UBound(Data, 1)) = Data

End Sub
Sub PrintArray3(Data, sheetname, startRow, StartCol)

    Dim row As Integer
    Dim col As Integer

    row = startRow

    For i = LBound(Data, 1) To UBound(Data, 1)
        col = StartCol
        For j = LBound(Data, 2) To UBound(Data, 2)
            Sheets(sheetname).Cells(row, col).Value = Data(i, j)
            col = col + 1
        Next j
            row = row + 1
    Next i

End Sub
Function getText() As String()

    Dim arr() As String
    Dim i As Integer
    i = 0
    Path = Application.ActiveWorkbook.Path
    Debug.Print Path

    Dim filename As String
    filename = Path & "\data.txt"

    Debug.Print filename

    Open filename For Input As #1
    Do While Not EOF(1) 'Loop until End Of File
        ReDim Preserve arr(i) 'Redim array for new element
        Line Input #1, arr(i) 'read next line and insert into array
        i = i + 1

    Loop
    Close #1 ' Close file.

    getText = arr

End Function
Sub A64_A108_Data(ByVal filename As String, ByVal sheetname As String)

    'Dim f_name As String
    'Set f_name = Path.GetFileName(fullFilename)

    Windows("AspenModels05182020Done.xlsx").Activate
    Range("A64:A110").Select
    Selection.Copy
    Windows(filename).Activate
    Sheets(sheetname).Select
    Range("A64").Select
    ActiveSheet.Paste

End Sub
Public Function CopySheetAndRename(ByVal name As String) As String
    'Dim newName As String

    'On Error Resume Next
    'newName = InputBox("Enter the name for the copied worksheet")

    If name <> "" Then
        ActiveSheet.Copy After:=Worksheets(Sheets.count)
        On Error Resume Next
        ActiveSheet.name = newName
    End If

    CopySheetAndRename = newName
    Debug.Print newName

End Function

Sub DeleteRowsBelow()

    'Worksheets("Overall").Activate
    Dim cell As Range
    Dim savedRow, i As Integer

    For Each cell In ActiveSheet.usedRange

        If cell.Value = "Mass Flows" Then
            savedRow = cell.row
            'MsgBox (savedRow)
            Exit For
        End If
    Next cell

    'MsgBox (savedRow)

    Dim sheet_name As String
    sheet_name = ActiveSheet.name

    With Sheets(sheet_name)
        .Rows(savedRow + 1 & ":" & .Rows.count).Delete
    End With

End Sub

Sub DeleteRowsAbove()

    'Worksheets("Overall").Activate
    Dim cell As Range
    Dim savedRow, i As Integer

    For Each cell In ActiveSheet.usedRange

        If cell.Value = "Stream Name" Then
            savedRow = cell.row
            'MsgBox (savedRow)
            Exit For
        End If
    Next cell

    Range("A1:A" & savedRow - 1).EntireRow.Delete

End Sub

Sub removeExtraRows()

    Call removeRow("Maximum Relative Error")
    Call removeRow("Cost Flow")
    'Call removeRow("Stream Class")
    'Call removeRow("Phase")
    Call removeRow("Molar Liquid Fraction")
    Call removeRow("Molar Solid Fraction")
    Call removeRow("Mass Vapor Fraction")
    Call removeRow("Mass Liquid Fraction")
    Call removeRow("Mass Solid Fraction")
    Call removeRow("Mass Enthalpy")
    Call removeRow("Mass Entropy")
    Call removeRow("Mass Density")
    Call removeRow("Mole Fractions")
    Call removeRow("MIXED Substream")
    Call removeRow("Description")

End Sub
Sub addBeginningRows()

    '{A1}.Value shorthand for Range("A1").Value

    ActiveSheet.Rows("1:3").EntireRow.Insert
    [A1].Value = "Example Streams 5-18"
    'Where did this come from and do I need to prompt user to enter name?
    [A2].Value = "In"
    [A3].Value = "Out"

End Sub
Sub insertRowBelowActiveCell()

    ActiveCell.Offset(1).EntireRow.Insert shift:=xlShiftDown

End Sub

Sub addRowEntropyFlow()

    Dim row As Integer
    row = FindRowByName("Enthalpy Flow")

    Debug.Print "Checking Ethalpy Units"
    Call checkEnthalpyUnits(row)

    ActiveSheet.Range("A1").Activate
    ActiveCell.Offset(row).EntireRow.Insert shift:=xlShiftDown

    'Rows(row).Insert shift:=xlShiftDown

    'With ActiveCell
    '   .Offset(row).EntireRow.Insert Shift:=xlShiftDown
    'End With

    row = row + 1

    Dim rowAsString As String
    rowAsString = "A" & CStr(row)
    Range(rowAsString).Select
    Selection.Value = "Entropy Flow"

    rowAsString = "B" & CStr(row)
    Range(rowAsString).Select
    Selection.Value = "kW/K"

    rowAsString = "C" & CStr(row)

    Call PopulateEntropyFlow(row)

    'ActiveCell.Value = rowAsString

    'Debug.Print ActiveCell.Value

    'ActiveCell.FormulaR1C1 = "=R[-3]C*R[2]C/3600*0.001"

    'Set SourceRange = ActiveSheet.Range(rowAsString)

    'Dim rowAsString2 As String
    'rowAsString = "D" & CStr(row)
    'rowAsString2 = "CO" & CStr(row)

    'Set fillRange = ActiveSheet.Range(rowAsString & ":" & rowAsString2)

    'SourceRange.Copy fillRange

    'Rows(rw + 1).Columns("B:C").Interior.Color = RGB(191, 191, 191)

End Sub
Sub PopulateEntropyFlow(ByVal row As Integer)

    Dim rng As Variant
    rng = "C" & row
    ActiveSheet.Range(rng).Select

    Debug.Print "mole flow units: " & Cells(row + 2, 2)

    If Cells(row + 2, 2) = "kmol/sec" Then
        ActiveCell.FormulaR1C1 = "=R[-3]C*R[2]C*0.001"
    Else
        ActiveCell.FormulaR1C1 = "=R[-3]C*R[2]C/3600*0.001"
    End If

    Set SourceRange = ActiveSheet.Range(rng)

    Dim lastCol As Variant
    lastCol = Cells(row - 1, Columns.count).End(xlToLeft).Column

    lastCol = Number2Letter(lastCol)

    'Debug.Print "last Col: " & lastCol

    Dim fillRng As Variant

    fillRng = "D" & row & ":" & lastCol & row
    Set fillRange = ActiveSheet.Range(fillRng)
    SourceRange.Copy fillRange

End Sub
Function Number2Letter(ByVal col As Integer) As String
    'PURPOSE: Convert a given number into it's corresponding Letter Reference
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim ColumnNumber As Long

    'Input Column Number
      ColumnNumber = col

    'Convert To Column Letter
     Number2Letter = Split(Cells(1, ColumnNumber).Address, "$")(1)

End Function
Sub addRowExergyFlow()

    Dim row As Integer
    row = FindRowByName("Entropy Flow")
    'should return 16 here

    ActiveSheet.Range("A1").Activate
    ActiveCell.Offset(row).EntireRow.Insert shift:=xlShiftDown

    row = row + 1
    'Debug.Print row, ActiveCell

    'Rows(row).EntireRow.Insert

    Dim rowAsString As String
    rowAsString = "A" & CStr(row)
    Range(rowAsString).Select
    Selection.Value = "Exergy Flow"

    rowAsString = "B" & CStr(row)
    Range(rowAsString).Select
    Selection.Value = "MW"

    Call PopulateExergyFlow(row)

End Sub
Sub PopulateExergyFlow(ByVal row As Integer)

    Dim rng As Variant
    rng = "C" & row
    ActiveSheet.Range(rng).Select
    ActiveCell.FormulaR1C1 = "=R[-2]C-0.3*R[-1]C"
    Set SourceRange = ActiveSheet.Range(rng)
    Dim lastCol As Variant
    lastCol = Cells(row - 1, Columns.count).End(xlToLeft).Column
    lastCol = Number2Letter(lastCol)
    'Debug.Print "last Col: " & lastCol
    Dim fillRng As Variant
    fillRng = "D" & row & ":" & lastCol & row
    Set fillRange = ActiveSheet.Range(fillRng)
    'Set SourceRange = ActiveSheet.Range("C17")
    'Set fillRange = ActiveSheet.Range("D17:CO17")
    'SourceRange.AutoFill Destination:=fillRange, Type:=xlFillCopy
    SourceRange.Copy fillRange

End Sub
Sub freezePanes(ByVal cell As Variant)

    ActiveSheet.Range(cell).Select
    ActiveWindow.freezePanes = True

End Sub
Function FindRowByName(ByVal key As String) As Integer

    Dim cell As Range

    For Each cell In ActiveSheet.usedRange

        If cell.Value = key Then
            FindRowByName = cell.row
            Exit For
        End If
    Next cell

End Function
Function FindRowByNameStreams(ByVal key As String) As Integer

    Dim cell As Range

    For Each cell In streamOverall.usedRange

        If cell.Value = key Then
            FindRowByNameStreams = cell.row
            Exit For
        End If
    Next cell

End Function

Public Sub removeRow(ByVal key As String)

    Dim cell As Range
    Dim savedRow, i As Integer

    For Each cell In ActiveSheet.usedRange

        If cell.Value = key Then
            savedRow = cell.row
            'MsgBox (savedRow)
            Exit For
        End If
    Next cell

    'MsgBox (savedRow)

    With ActiveSheet
        .Rows(savedRow).Delete
    End With

End Sub
Public Sub calc_balances()

    Dim curRow As Integer
    Dim sum As Variant

    curRow = FindRowByName("Enthalpy Flow")
    sum = calc_balances2(curRow) 'Enthalpy Flow
    Call InsertSum(sum, curRow)

    curRow = FindRowByName("Entropy Flow")
    sum = calc_balances2(curRow) 'Entropy Flow
    Call InsertSum(sum, curRow)

    curRow = FindRowByName("Exergy Flow")
    sum = calc_balances2(curRow) 'Exergy Flow
    Call InsertSum(sum, curRow)

    curRow = FindRowByName("Mass Flows")
    sum = calc_balances2(curRow) 'Mass Flows
    If sum > 10 Then
        sum = "MB Error"
    End If
    Call InsertSum(sum, curRow)

End Sub

Public Function calc_balances2(ByVal row As Integer) As Variant

    Dim lastCol As Variant
    lastCol = Cells(10, Columns.count).End(xlToLeft).Column
    lastCol = Number2Letter(lastCol)

    Dim my_range As String
    Let my_range = "C" & row & ":" & lastCol & row
    Dim rng As Range: Set rng = Application.Range(my_range)
    'Debug.Print my_range

    Dim in_sum As Variant
    Dim out_sum As Variant
    Dim tot_sum As Variant
    in_sum = 0
    out_sum = 0
    tot_sum = 0

    For i = 3 To rng.Cells.count + 2 Step 1

        If Cells(1, i).Value = "In" Then
            in_sum = in_sum + Cells(row, i).Value
        End If

        If Cells(1, i).Value = "Out" Then
            out_sum = out_sum + Cells(row, i).Value
        End If

        'Debug.Print i, in_sum, out_sum
    Next i

    tot_sum = out_sum - in_sum
    'Debug.Print "tot_sum:", tot_sum
    calc_balances2 = tot_sum

End Function
Public Sub removeToFromColumns()

    Rows(1).Insert shift:=xlShiftDown
    Dim lastCol As Variant
    lastCol = Cells(10, Columns.count).End(xlToLeft).Column
    lastCol = Number2Letter(lastCol)

    Dim fromRow, toRow As Integer
    fromRow = FindRowByName("From")
    toRow = fromRow + 1
    Dim appRng As Variant
    appRng = "C" & fromRow & ":" & lastCol & toRow

    Debug.Print "appRng: " & appRng

    Dim rng As Range: Set rng = Application.Range(appRng)

    Dim A As String
    Dim B As String
    Dim i As Integer

    For i = rng.Cells.count To 3 Step -1
        'A = Cells(fromRow, i)
        'B = Cells(toRow, i)
        'Debug.Print A, i, B
        If IsEmpty(Cells(fromRow, i)) = False And IsEmpty(Cells(toRow, i)) = False Then
            'Debug.Print "A is EMPTY"
            Columns(i).EntireColumn.Delete
            'Debug.Print "Double Column"
            'Debug.Print "B is EMPTY"
        End If
    Next i

    lastCol = Cells(10, Columns.count).End(xlToLeft).Column

    For i = lastCol To 3 Step -1
        If IsEmpty(Cells(fromRow, i)) = True Then
            'means FROM is blank, so add "In"
            Cells(1, i).Value = "In"
        ElseIf IsEmpty(Cells(toRow, i)) = True Then
            Cells(1, i).Value = "Out"
        End If
    Next i

End Sub

Public Sub InsertSum(ByVal sum As Variant, ByVal row As Integer)

    Dim lastCol As Variant
    lastCol = Cells(10, Columns.count).End(xlToLeft).Column + 1
    lastCol = Number2Letter(lastCol)

    Cells(row, lastCol).Value = sum
    Cells(row, lastCol).Interior.Color = RGB(255, 255, 0)

End Sub
