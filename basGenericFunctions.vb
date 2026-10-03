Option Strict Off
Option Explicit On
Module basGenericFunctions
	
	Public Function doCase(ByRef aStr As String) As String
		
        If gDatalayer.DatabaseType = DB_POSTGRESQL Then
            doCase = UCase(aStr)
        Else
            doCase = aStr
        End If

    End Function


    Public Function MxDate(ByRef aDate As Date) As Date

        If aDate > #1/1/9999# Then
            MxDate = #1/1/9999#
        Else
            MxDate = aDate
        End If


    End Function


    Public Function is_TSQL_Date(varDate As Object) As Boolean

        Dim T As String
        Dim aDate As Date

        T = CStr(varDate)

        If Len(T) = 0 Then
            is_TSQL_Date = False
        ElseIf IsDate(T) Then
            aDate = CDate(T)
            If aDate < #1/1/1753# Or aDate > #12/31/9999# Then
                is_TSQL_Date = False
            Else
                is_TSQL_Date = True
            End If
        Else
            is_TSQL_Date = False
        End If

    End Function


    Public Sub AddDetField(ByRef aColl As Collection, ByVal colName As String, ByVal colHeader As String, ByVal DataType As Integer, ByVal prefix As String, _
                                            ByVal decimals As Integer, ByVal isTotal As Integer)

        Dim aField As CSoft_DataAccess.clsEditValue

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aField = New CSoft_DataAccess.clsEditValue
        aField.FieldName = colName
        aField.Header = colHeader
        aField.DataType = DataType
        aField.Prefix = prefix
        aField.Decimals = decimals
        aField.FieldVal = isTotal
        aColl.Add(aField)

        Exit Sub

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Sub AddDetField(ByRef aColl As CSoft_DataAccess.vbCollection, ByVal colName As String, ByVal colHeader As String, ByVal DataType As Integer,
                                ByVal prefix As String, ByVal decimals As Integer, ByVal isTotal As Integer)

        Dim aField As CSoft_DataAccess.clsEditValue

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aField = New CSoft_DataAccess.clsEditValue
        aField.FieldName = colName
        aField.Header = colHeader
        aField.DataType = DataType
        aField.Prefix = prefix
        aField.Decimals = decimals
        aField.FieldVal = isTotal
        aColl.Add(aField)

        Exit Sub

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

    Public Function GetDestinationFromScaleIni() As String

        Dim tempStr As String
        Dim iLen As Integer
        Dim strFile As String
        Dim iRet As Integer
        Dim Destination As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetDestinationFromScaleIni = ""

        strFile = My.Application.Info.DirectoryPath & "\Scale.ini"

        Destination = ""
        tempStr = New String(" ", 255)
        iLen = 255
        iRet = GetPrivateProfileString("GENERAL", "Destination", "", tempStr, iLen, strFile)
        If iRet <> 0 Then
            If Len(tempStr) <> 0 Then
                Destination = Mid(tempStr, 1, iRet)
            End If
        End If

        GetDestinationFromScaleIni = Destination

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetDirName(ByRef ScanString As String) As String

        Dim intPos As Short
        Dim intPosSave As Short

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        intPos = 1
        Do
            intPos = InStr(intPos, ScanString, "\")
            If intPos = 0 Then
                Exit Do
            Else
                intPos = intPos + 1
                intPosSave = intPos - 1
            End If
        Loop
        GetDirName = Left(ScanString, intPosSave)

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetFileName(ByRef ScanString As String) As String

        Dim intPos As Short
        Dim intPosSave As Short

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        intPos = 1
        Do
            intPos = InStr(intPos, ScanString, "\")
            If intPos = 0 Then
                Exit Do
            Else
                intPos = intPos + 1
                intPosSave = intPos - 1
            End If
        Loop
        GetFileName = Trim(Mid(ScanString, intPosSave + 1))

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function Sortem(ByRef X() As Double, ByRef Y() As Integer, ByVal n As Integer, ByVal d As Short) As Short
        Sortem = 0
        '----------------------------------------------------------------------------------------------------------
        ' Real Number Sorter
        ' x in an array of doubles -> input
        ' Y is an array of longs
        ' Y(i) is the original rows sorted from low to high
        '----------------------------------------------------------------------------------------------------------

        Dim IR, i, j, L As Integer
        Dim xa As Double
        Dim yb As Integer

        If n < 2 Then
            Exit Function
        End If

        L = n / 2 + 1
        IR = n
        d = 88

Label1:

        If (L > 1) Then

            L = L - 1
            xa = X(L)
            yb = Y(L)

        Else

            xa = X(IR)
            yb = Y(IR)
            X(IR) = X(1)
            Y(IR) = Y(1)
            IR = IR - 1

            If (IR = 1) Then


                X(1) = xa
                Y(1) = yb

                Sortem = d
                Exit Function
            End If

        End If

        i = L
        j = L + L

Label2:

        If (j <= IR) Then

            If (j < IR) Then

                If (X(j) < X(j + 1)) Then
                    j = j + 1
                End If

            End If


            If (xa < X(j)) Then

                X(i) = X(j)
                Y(i) = Y(j)
                i = j
                j = j + j

            Else

                j = IR + 1

            End If

            GoTo Label2

        End If

        X(i) = xa
        Y(i) = yb

        GoTo Label1

        Sortem = n
        Exit Function

    End Function

    Function SortemStr(ByRef X() As String, ByRef Y() As Integer, ByVal n As Integer, ByVal d As Integer) As Integer
        '----------------------------------------------------------
        '                            String Sorter
        '----------------------------------------------------------
        Dim i As Integer
        Dim j As Integer
        Dim IR As Integer
        Dim L As Integer
        Dim xa As String
        Dim yb As Integer

        L = n / 2 + 1
        IR = n
        d = 88

Label1:

        If (L > 1) Then

            L = L - 1
            xa = X(L)
            yb = Y(L)

        Else

            xa = X(IR)
            yb = Y(IR)
            X(IR) = X(1)
            Y(IR) = Y(1)
            IR = IR - 1

            If (IR = 1) Then

                X(1) = xa
                Y(1) = yb

                SortemStr = d
                Exit Function
            End If

        End If

        i = L
        j = L + L

Label2:

        If (j <= IR) Then

            If (j < IR) Then

                If (X(j) < X(j + 1)) Then
                    j = j + 1
                End If

            End If


            If (xa < X(j)) Then

                X(i) = X(j)
                Y(i) = Y(j)
                i = j
                j = j + j

            Else

                j = IR + 1

            End If

            GoTo Label2

        End If

        X(i) = xa
        Y(i) = yb

        GoTo Label1

        SortemStr = n

        Exit Function

    End Function

    Public Function Pow(ByRef X As Double, ByRef Ex As Double) As Double

        Pow = System.Math.Exp(Ex * System.Math.Log(X))

    End Function

    Public Function isKeyinVBCollection(ByRef aColl As CSoft_DataAccess.vbCollection, ByRef aKey As String) As Boolean

        Dim aVar As Object

        On Error GoTo 100

        aVar = aColl(aKey)
        If IsNothing(aVar) Then
            isKeyinVBCollection = False
        Else
            isKeyinVBCollection = True
        End If

200:
        Exit Function

100:
        isKeyinVBCollection = False
        Resume 200

    End Function

    Public Function isKeyinCollection(ByRef aColl As Collection, ByRef aKey As String) As Boolean

        Dim aVar As Object

        On Error GoTo 100

        If aColl.Contains(aKey) Then
            isKeyinCollection = True
        Else
            isKeyinCollection = False
        End If
        'aVar = aColl(aKey)
        'isObjKeyinVBCollection = True

200:
        Exit Function

100:
        isKeyinCollection = False
        Resume 200

    End Function



    Public Function isObjKeyinCollection(ByRef aColl As Collection, ByRef aKey As String) As Boolean

        Dim aVar As Object

        On Error GoTo 100

        If aColl.Contains(aKey) Then
            isObjKeyinCollection = True
        Else
            isObjKeyinCollection = False
        End If

200:
        Exit Function

100:
        isObjKeyinCollection = False
        Resume 200

    End Function

    Public Function isObjKeyinVBCollection(ByRef aColl As Collection, ByRef aKey As String) As Boolean

        Dim aVar As Object

        On Error GoTo 100

        If aColl.Contains(aKey) Then
            isObjKeyinVBCollection = True
        Else
            isObjKeyinVBCollection = False
        End If
        'aVar = aColl(aKey)
        'isObjKeyinVBCollection = True

200:
        Exit Function

100:
        isObjKeyinVBCollection = False
        Resume 200

    End Function

    Public Function GetIndexOfIteminCollection(ByRef aColl As CSoft_DataAccess.vbCollection, ByRef Item As String) As Integer

        Dim i As Integer

        On Error GoTo 100

        GetIndexOfIteminCollection = NA

        For i = 1 To aColl.Count
            If (StrComp(aColl.Item(i), Item, CompareMethod.Text) = 0) Then
                GetIndexOfIteminCollection = i
                Exit Function
            End If
        Next i

        Exit Function

100:
        GetIndexOfIteminCollection = NA

    End Function



    Public Function getAValue(ByRef aFieldName As String, ByRef aTableName As String, ByRef anIDField As String, ByVal anIDValue As Object) As Object

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSQL As String
        Dim objErrors As clsLib_DataAccess.clsErrors

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        objErrors = New clsLib_DataAccess.clsErrors

        If (VarType(anIDValue) = vbString) Then

            'anIDValue = Replace(anIDValue, "'", "''") 
            strSQL = "SELECT " & aFieldName & " FROM " & aTableName & " WHERE " & anIDField & " = |" &
                     anIDValue & "|"

        ElseIf (VarType(anIDValue) = vbDate) Then

            strSQL = "SELECT " & aFieldName & " FROM " & aTableName & " WHERE " & anIDField & " = " & gVars.gDI &
                     gBF.IntlDateFormat(CDate(anIDValue)) & gVars.gDI

        Else

            strSQL = "SELECT " & aFieldName & " FROM " & aTableName & " WHERE " & anIDField & " = " &
                     anIDValue

        End If

        aRS = gDataLayer.LoadRecordset(strSQL, objErrors)

        'aRS.Open strSql, gDB, adOpenDynamic, adLockReadOnly, adCmdText 

        If Not aRS.EOF Then
            getAValue = aRS(aFieldName).Value
        Else
            getAValue = Nothing
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function getAValue0(ByRef aFieldName As String, ByRef aTableName As String, ByRef anIDField As String, ByVal anIDValue As Object, ByRef objerrors As clsLib_DataAccess.clsErrors) As Object

        Dim aRS As clsLib_DataAccess.clsRecordset
        Dim strSQL As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If (VarType(anIDValue) = VariantType.String) Then

            'anIDValue = Replace(anIDValue, "'", "''")
            strSQL = "SELECT " & aFieldName & " FROM " & aTableName & " WHERE " & anIDField & " = |" & anIDValue & "|"

        ElseIf (VarType(anIDValue) = VariantType.Date) Then

            strSQL = "SELECT " & aFieldName & " FROM " & aTableName & " WHERE " & anIDField & " = " & gVars.gDI & gBF.IntlDateFormat(CDate(anIDValue)) & gVars.gDI

        Else

            strSQL = "SELECT " & aFieldName & " FROM " & aTableName & " WHERE " & anIDField & " = " & anIDValue

        End If

        aRS = gDataLayer.LoadRecordset(strSQL, objerrors)
        If (gObjErrors.Count = 0) Then
        Else
            aRS = Nothing
            getAValue0 = Nothing
            Exit Function
        End If

        'aRS.Open strSql, gDB, adOpenDynamic, adLockReadOnly, adCmdText

        If Not aRS.EOF Then
            getAValue0 = aRS(aFieldName).Value
        Else
            getAValue0 = Nothing
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        objerrors.Add(Err.Description, Err.Number, Err.Source & ":basFunctions.getAValue")

    End Function
End Module
