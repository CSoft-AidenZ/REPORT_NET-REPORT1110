Option Strict Off
Option Explicit On
Module basGridFunctionsRpts
	
	Public Const MAX_COLUMN_WIDTH As Short = 100 'Minimum width of a column in the grid on frmSetup
	Public Const MIN_COLUMN_WIDTH As Short = 10 'Minimum width of a column in the grid on frmSetup
	Public Const MAX_FIELDS_PER_COLUMN As Short = 12 'Maximum number of items shown per column on frmEditData
	Public Const INITIAL_LABEL_WIDTH As Short = 720 'Initial width of a label on frmSetup
	Public Const INITIAL_CONTROL_WIDTH As Short = 720 'Initial width of a non-label on frmSetup
	Public Const MAX_CONTROL_WIDTH As Short = 3500 'Maximum width of a non-label on frmSetup
	Public Const CONTROL_WIDTH_STEP As Short = 250 'Controls will be rounded to the nearest control_width_step for width
	
	Public Const FP_ACTION_CLEARTEXT As Short = 12
	Public Const FP_ACTION_CLEAR As Short = 3
	
    Public Sub CopyAndPaste(ByVal aGrid As Object, ByVal KeyCode As Short, ByVal Shift As Short, ByRef iCopy As Integer, ByRef mRow1 As Integer, ByRef mRow2 As Integer, ByRef mCol1 As Integer, ByRef mCol2 As Integer)


        Dim i As Short
        Dim c2, c, r, r2 As Object

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        Dim T() As String
        Dim xStr As String
        Dim iR As Integer
        Dim ic As Integer
        Dim iCell As Integer

        'If Shift = VB6.ShiftConstants.CtrlMask Then
        If Shift = vbCtrlMask Then

            If KeyCode = System.Windows.Forms.Keys.C Then

                If aGrid.SelectionCount > 0 Then

                    c = Nothing
                    r = Nothing
                    c2 = Nothing
                    r2 = Nothing
                    aGrid.GetSelection(0, c, r, c2, r2)

                Else

                    c = aGrid.ActiveCol
                    r = aGrid.ActiveRow
                    c2 = c
                    r2 = r

                End If

                If r > 0 And c > 0 Then

                    aGrid.ClipboardCopy()
                    iCopy = True
                    mCol1 = c
                    mCol2 = c2
                    mRow1 = r
                    mRow2 = r2

                End If

            ElseIf KeyCode = System.Windows.Forms.Keys.V Then

                If aGrid.SelectionCount > 0 Then

                    c = Nothing
                    r = Nothing
                    c2 = Nothing
                    r2 = Nothing
                    aGrid.GetSelection(0, c, r, c2, r2)

                Else

                    c = aGrid.ActiveCol
                    r = aGrid.ActiveRow
                    c2 = c
                    r2 = r

                End If

                If r = r2 And c = c2 Then 'One Cell

                    aGrid.ClipboardPaste()
                ElseIf (r2 - r) = (mRow2 - mRow1) And (c2 - c) = (mCol2 - mCol1) Then  'Same Size Block

                    aGrid.ClipboardPaste()

                ElseIf (mRow2 - mRow1) = 0 Then  'One or more cells in ONE row as source;  Multiple rows to paste

                    xStr = My.Computer.Clipboard.GetText

                    T = Split(xStr, Chr(9))


                    For iR = r To r2
                        iCell = -1
                        For ic = c To c2
                            iCell = iCell + 1
                            If iCell <= UBound(T) Then
                                aGrid.SetText(ic, iR, CObj(T(iCell)))
                            End If
                        Next ic
                    Next iR

                End If

            End If


        End If

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub



    '-----------------------------------------
    '
    '  Add Report Cell
    '  Changed its Name 6/13/11 because it duplicated AddACell in BasicFunctions.bas
    '
    '-----------------------------------------
    Public Sub AddaRptCell(ByVal aGrid As Object, ByRef aColl As Collection, ByVal Col As Integer, ByVal Row As Integer, ByVal Fld As String, ByVal fldType As Integer, ByRef fields As Collection, ByVal aCaption As String)

        Dim aCell As clsLib_DataAccess.clsCell

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aGrid.Row = Row
        aGrid.Col = Col
        aGrid.ForeColor = System.Drawing.Color.Blue

        aCell = New clsLib_DataAccess.clsCell
        aCell.Row = Row
        aCell.Col = Col
        aCell.FieldName = Fld
        aCell.FieldType = fldType
        aCell.RepCaption = aCaption
        aCell.fieldColl = fields
        aColl.Add(aCell, aCell.FieldName)

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Sub writeTextToCell(ByVal aGrid As Object, ByRef T As String, ByRef r As Integer, ByVal c1 As Integer, ByVal c2 As Integer, ByVal isBold As Integer, ByRef objerrors As clsLib_DataAccess.clsErrors)

        Dim c As Integer
        Dim ii As Integer
        Dim j As Integer
        Dim maxCellW As Integer
        Dim aStr As String
        Dim r0 As Integer
        Dim r1 As Integer
        Dim xT() As String
        Dim n As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Len(T) = 0 Then
            Exit Sub
        End If

        xT = Split(T, vbCrLf)

        For j = UBound(xT) To 0 Step -1
            If Len(xT(j)) > 0 Then
                n = j
                Exit For
            End If
        Next j

        ReDim Preserve xT(n)

        r = r - 1

        For j = 0 To UBound(xT)

            r = r + 1
            If r > aGrid.MaxRows Then
                aGrid.MaxRows = r
            End If

            maxCellW = 0
            T = xT(j)
            T = Trim(T)

            r0 = r
            r1 = r

            With aGrid

                For c = c1 To c2
                    maxCellW = maxCellW + .ColWidth(c)
                Next c

                .Row = r
                '.row2 = r
                .Col = c1
                '.Col2 = c1
                .TypeMaxEditLen = 200
                .FontBold = isBold
                .TypeHAlign = TypeHAlignLeft
                .AllowCellOverflow = True
                .SetText(c1, r, CStr(T))
                '.BlockMode = True
                '.BlockMode = False

                ii = .MaxTextCellWidth

                If (ii > maxCellW) Then

                    aStr = GetSnippet(aGrid, T, maxCellW, c1, r, objerrors)
                    .Row = r
                    .Col = c1
                    .FontBold = isBold
                    .SetText(c1, r, CStr(aStr))
                    ii = .MaxTextCellWidth

                    'Write what's left over
                    r = r + 1
                    If r > .MaxRows Then
                        .MaxRows = r
                    End If

                    .Row = r
                    .Col = c1
                    .FontBold = isBold
                    .TypeMaxEditLen = 200
                    .TypeHAlign = TypeHAlignLeft
                    .AllowCellOverflow = True
                    .SetText(c1, r, CStr(T))
                    ii = .MaxTextCellWidth

                    Do Until ii <= maxCellW

                        aStr = GetSnippet(aGrid, T, maxCellW, c1, r, objerrors)
                        .SetText(c1, r, CStr(aStr))

                        'Write what's left over
                        r = r + 1
                        If r > .MaxRows Then
                            .MaxRows = r
                        End If
                        .Row = r
                        .Col = c1
                        .FontBold = isBold
                        .TypeMaxEditLen = 200
                        .TypeHAlign = TypeHAlignLeft
                        .AllowCellOverflow = True
                        .SetText(c1, r, CStr(T))
                        ii = .MaxTextCellWidth

                    Loop

                End If

                r1 = r

            End With

        Next j

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Function GetSnippet(ByVal aGrid As Object, ByRef T As String, ByVal maxCellW As Object, ByVal c As Integer, ByVal r As Integer, ByRef objerrors As clsLib_DataAccess.clsErrors) As String

        Dim xStr As String
        Dim i As Integer
        Dim tmp As String
        Dim n As Integer
        Dim ii As Integer
        Dim jj As Integer
        Dim aChar As String
        Dim lastSpace As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        T = LTrim(T) 'Trim

        'Split into blocks of text
        'tmp = Split(T, " ")
        'tmp = SplitTbl(T, delimiterArray, objerrors, " ")
        tmp = T

        n = Len(tmp)

        xStr = ""
        lastSpace = -1

        For i = 1 To n

            With aGrid

                aChar = Mid(tmp, i, 1)
                'xStr = xStr & tmp(i) & delimiterArray(i)
                If StrComp(aChar, " ", CompareMethod.Text) = 0 Then
                    lastSpace = i
                End If
                xStr = xStr & aChar

                .SetText(c, r, CStr(xStr))
                .Col = c
                .Row = r
                ii = .MaxTextCellWidth

                If (ii > maxCellW) Then
                    GoTo KickOut
                    'Done, Revert back to last space
                End If

            End With

        Next i

        lastSpace = n

KickOut:
        'Handle if the first string itself is too long...
        'Just back up one char
        If lastSpace = -1 Then
            lastSpace = Len(xStr) - 1
            xStr = Left(xStr, lastSpace)
        Else
            xStr = Left(xStr, lastSpace - 1) 'Knock off last character (it is a space)
        End If


        If (jj < n) Then
            T = Right(T, Len(T) - Len(xStr))
        Else
            T = ""
        End If

        T = LTrim(T)

        GetSnippet = xStr

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function



    Public Sub DoGridSetup(ByRef Table As clsLib_DataAccess.clsTable, ByVal aGrid As Object, ByRef aForm As System.Windows.Forms.Form, Optional ByRef Editable As Boolean = False)

        Dim charsW As Integer
        Dim rwTwips As Integer
        Dim lngField As Integer
        Dim lScreen As Integer
        Dim lngFirstCol As Integer
        Dim lngFirstRow As Integer
        Dim lngLastCol As Integer
        Dim lngLastRow As Integer
        Dim colScreen As Collection

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        With aGrid
            .GetFirstValidCell(lngFirstCol, lngFirstRow)
            .GetLastValidCell(lngLastCol, lngLastRow)
            ' Select a block of cells
            .Row = lngFirstRow
            .Col = lngFirstCol
            .Row2 = lngLastRow
            .Col2 = lngLastCol
            .BlockMode = True

            ' Clear the Table from the cells
            .Action = 12 ' SS_ACTION_CLEAR_TEXT

            .RowHeightToTwips(1, aGrid.rowHeight(1), rwTwips)
            .MaxCols = Table.Fields.Count
            .ColWidth(0) = 6
            .rowHeight(0) = 15
            charsW = .ColWidth(0)

            'Record Column Headers
            .SetText(0, 0, "Record")
            'aGrid.ColWidth(1) = 0

            .Row = -1
            .Row2 = -1
            '.BlockMode = False
        End With

        If Table.BuildScreenOrder(gObjErrors) = False Then
            DisplayErrors()
            Exit Sub
        End If

        colScreen = Table.GetScreenOrder(gObjErrors)
        If colScreen Is Nothing Then
            DisplayErrors()
            Exit Sub
        End If

        For lScreen = 1 To colScreen.Count
            lngField = colScreen(lScreen)
            If Table.Fields(lngField).ScreenOrder = lScreen Then
                If (Table.Fields(lngField).DisplayOrder > -1) Then
                    'We found the next non-hidden column

                    aGrid.SetText(lScreen, 0, UCase(Table.Fields(lngField).DisplayName))
                    'Setting the limitations for the column width
                    aGrid.ColWidth(lScreen) = aGrid.MaxTextColWidth(lngField)
                    If aGrid.ColWidth(lScreen) > MAX_COLUMN_WIDTH Then
                        aGrid.ColWidth(lScreen) = MAX_COLUMN_WIDTH
                    ElseIf aGrid.ColWidth(lScreen) < MIN_COLUMN_WIDTH Then
                        aGrid.ColWidth(lScreen) = MIN_COLUMN_WIDTH
                    End If
                    'aGrid.ColWidth(lScreen) = Min(Table.Fields(lngField).Size, 100)
                    charsW = charsW + aGrid.ColWidth(lScreen)
                    aGrid.Col = lScreen
                    aGrid.Col2 = lScreen
                    If Editable Then
                        If Table.Fields(lngField).Required = True Then
                            aGrid.BackColor = System.Drawing.Color.Yellow
                        Else
                            aGrid.BackColor = System.Drawing.Color.White
                        End If
                    End If
                    aGrid.CellBorderStyle = 1 'SS_BORDER_STYLE_SOLID
                    aGrid.CellBorderType = 16 'SS_BORDER_TYPE_OUTLINE
                    aGrid.CellBorderColor = System.Drawing.Color.Black
                    aGrid.Action = ActionSetCellBorder
                Else
                    aGrid.SetText(lScreen, 0, UCase(Table.Fields(lngField).DisplayName))
                    aGrid.Col = lScreen
                    aGrid.Col2 = lScreen
                    aGrid.ColHidden = True
                End If
            End If
        Next lScreen

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Sub UpdateGrid(ByVal aGrid As Object, ByVal Editable As Boolean)

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        With aGrid

            .Col = 1
            .Col2 = .MaxCols
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            If Editable = True Then
                '    .EditModePermanent = True
                .ProcessTab = True
                .Protect = True
            Else
                .Lock = True
                .Protect = True
                .EditModePermanent = False
                .ProcessTab = False
            End If
            .BlockMode = False
        End With

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Sub ClearGrid(ByVal aGrid As Object, ByRef ClearAll As Boolean)

        Dim lngFirstCol As Integer
        Dim lngFirstRow As Integer
        Dim lngLastCol As Integer
        Dim lngLastRow As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        ' Clear the grid
        'aGrid.Action = ActionClearText
        aGrid.GetFirstValidCell(lngFirstCol, lngFirstRow)
        aGrid.GetLastValidCell(lngLastCol, lngLastRow)

        ' Select a block of cells
        aGrid.Row = lngFirstRow
        aGrid.Col = lngFirstCol
        aGrid.Row2 = lngLastRow
        aGrid.Col2 = lngLastCol
        aGrid.BlockMode = True
        If (ClearAll) Then
            ' Clear the data from the cells
            aGrid.Action = FP_ACTION_CLEAR 'SS_ACTION_CLEAR_all
        Else
            ' Clear the data from the cells
            aGrid.Action = 12 'SS_ACTION_CLEAR_TEXT
        End If
        'aGrid.FontBold = False
        'aGrid.FontUnderline = False
        ' Turn block mode off
        aGrid.BlockMode = False

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub


    Public Sub ResortEx(ByVal aCol As Integer, ByVal aGrid As Object, ByRef theSortCol As Integer, ByRef theSortDir As Integer, ByVal maxRow As Integer)

        Static sortCol As Integer
        Static sortAscending As Integer
        'Dim aKey As Integer
        'Dim i As Integer , intCPosition As Integer

        TurnMeOn()

        With aGrid

            If .MaxRows < 3 Then
                Exit Sub
            End If

            If .MaxCols < aCol Then
                Exit Sub
            End If

            If sortCol = aCol Then

                If sortAscending Then
                    sortAscending = False
                    .SortKey(1) = aCol
                    .SortKeyOrder(1) = 1
                Else
                    sortAscending = True
                    .SortKey(1) = aCol
                    .SortKeyOrder(1) = 2
                End If

            Else

                sortCol = aCol
                .SortKey(1) = aCol
                'default sort to ascending
                .SortKeyOrder(1) = 1

            End If

            theSortCol = aCol
            theSortDir = .SortKeyOrder(1)

            .SortBy = 0
            .Col = 0
            .Col2 = .MaxCols
            .Row = 0
            'Use passed in max row
            .Row2 = maxRow
            .Action = 25

        End With

        TurnMeOff()


    End Sub


    Public Sub Resort(ByVal aCol As Integer, ByVal aGrid As Object, ByRef theSortCol As Integer, ByRef theSortDir As Integer)

        Static sortCol As Integer
        Static sortAscending As Integer
        'Dim aKey As Integer
        'Dim i, intCPosition As Integer

        TurnMeOn()

        With aGrid

            If .MaxRows < 3 Then
                Exit Sub
            End If

            If .MaxCols < aCol Then
                Exit Sub
            End If

            If sortCol = aCol Then

                If sortAscending Then
                    sortAscending = False
                    .SortKey(1) = aCol
                    .SortKeyOrder(1) = 1
                Else
                    sortAscending = True
                    .SortKey(1) = aCol
                    .SortKeyOrder(1) = 2
                End If

            Else

                sortCol = aCol
                .SortKey(1) = aCol
                'default sort to ascending
                .SortKeyOrder(1) = 1

            End If

            theSortCol = aCol
            theSortDir = .SortKeyOrder(1)

            .SortBy = 0
            .Col = 0
            .Col2 = .MaxCols
            .Row = 0
            .Row2 = .MaxRows
            .Action = 25

        End With

        TurnMeOff()


    End Sub


    Sub AlignCells(ByVal aGrid As Object, ByVal r1 As Integer, ByVal c1 As Integer, ByVal r2 As Integer, ByVal c2 As Integer, ByRef align As Integer)

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aGrid.Row = r1
        aGrid.Row2 = r2
        aGrid.Col = c1
        aGrid.Col2 = c2
        aGrid.BlockMode = True
        aGrid.TypeHAlign = align
        aGrid.BlockMode = False

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Sub BoldCellsColor(ByVal aGrid As Object, ByVal r1 As Integer, ByVal c1 As Integer, ByVal r2 As Integer, ByVal c2 As Integer, ByRef aColor As Integer)

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aGrid.Row = r1
        aGrid.Row2 = r2
        aGrid.Col = c1
        aGrid.Col2 = c2
        aGrid.BlockMode = True
        aGrid.FontBold = True
        aGrid.ForeColor = System.Drawing.ColorTranslator.FromWin32(aColor)
        aGrid.BlockMode = False

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Sub BorderCellsExt(ByVal aGrid As Object, ByVal r1 As Integer, ByVal c1 As Integer, ByVal r2 As Integer, ByVal c2 As Integer, ByRef BT As Integer)

        basGridFunctionsRpts.BorderCells(aGrid, r1, c1, r2, c2, BT)

    End Sub

    Public Sub BorderCells(ByVal aGrid As Object, ByVal r1 As Integer, ByVal c1 As Integer, ByVal r2 As Integer, ByVal c2 As Integer, ByRef BT As Integer)

        '------------------------------------------------------------------
        '   BT Values
        '   16 = Outline Cells
        '   1 = Left
        '   2 = Right
        '   4 = Top
        '   8 = Bottom
        '   16 = Outline Cells
        '------------------------------------------------------------------

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aGrid.Row = r1
        aGrid.Row2 = r2
        aGrid.Col = c1
        aGrid.Col2 = c2
        aGrid.BlockMode = True
        aGrid.CellBorderType = BT
        aGrid.CellBorderStyle = 1
        aGrid.CellBorderColor = System.Drawing.Color.Black
        aGrid.Action = 16
        aGrid.BlockMode = False

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)


    End Sub

End Module
