Option Strict Off
Option Explicit On
Imports System.Windows.Forms

Module basCommonFuncs

    Public Function GetRSbyCommand(ByRef cmdNum As Integer, ByVal ParamArray params() As Object) As CSOFT_RECORDSET_EXT.clsRecordsetExt

        Dim rs As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim nP As Integer
        Dim sqlStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        Select Case cmdNum

            Case 1
                sqlStr = "SELECT CODE FROM VENDOR_STATEMENTS WHERE PAY_PERIOD_CODE = |" & params(0) & "| AND OWNER_CODE = |" & params(1) & "|"
                GetRSbyCommand = gDataLayer.LoadRecordset(sqlStr, gObjErrors)

            Case 2
                sqlStr = "SELECT EQUIPMENT_CODE, OWNER_CODE, LOAD_PCT FROM ((LOADSLIP_PAY_ACTIVITIES LS INNER JOIN EQUIPMENT EQ ON LS.EQUIPMENT_CODE = EQ.CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) WHERE LS.LOAD_CODE = |" & params(0) & "| AND LS.PAY_ACTIVITY_CODE = |" & params(1) & "| " & " AND OW.OWNER = 0 AND LS.EQUIPMENT_CODE = |" & params(2) & "| "

                GetRSbyCommand = gDataLayer.LoadRecordset(sqlStr, gObjErrors)


            Case 3
                sqlStr = "SELECT EQUIPMENT_CODE, LOAD_PCT FROM ((LOADSLIP_PAY_ACTIVITIES LS INNER JOIN EQUIPMENT EQ ON LS.EQUIPMENT_CODE = EQ.CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) WHERE LS.LOAD_CODE = |" & params(0) & "| AND LS.PAY_ACTIVITY_CODE = |" & params(1) & "| " & " AND OW.OWNER = 0 AND OW.CODE = |" & params(2) & "| AND LS.EQUIPMENT_CODE = |" & params(3) & "| "

                GetRSbyCommand = gDataLayer.LoadRecordset(sqlStr, gObjErrors)


            Case 4
                sqlStr = "SELECT NEXT_ID FROM INCREMENT WHERE SEQUENCE = |" & params(0) & "|"
                GetRSbyCommand = gDataLayer.LoadRecordset(sqlStr, gObjErrors)

            Case 5
                sqlStr = "UPDATE INCREMENT SET NEXT_ID = NEXT_ID + 1 WHERE SEQUENCE = |" & params(0) & "|"
                gDataLayer.ExecuteSQL(sqlStr, gObjErrors, i)
                GetRSbyCommand = Nothing

            Case Else
                GetRSbyCommand = Nothing

        End Select

        Exit Function

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function isIteminCollection(ByRef aColl As Collection, ByRef Item As String) As Boolean

        Dim i As Integer

        On Error GoTo 100

        isIteminCollection = False

        For i = 1 To aColl.Count
            If (StrComp(aColl.Item(i), Item, CompareMethod.Text) = 0) Then
                isIteminCollection = True
                Exit Function
            End If
        Next i

        Exit Function

100:
        isIteminCollection = False


    End Function

    Public Function isFieldinRecordset(ByRef aRs As CSOFT_RECORDSET_EXT.clsRecordsetExt, ByVal aKey As String) As Boolean

        Dim aVar As Object
        Dim i As Integer

        On Error GoTo 100

        aVar = aRs(aKey).Value

        isFieldinRecordset = True

200:
        Exit Function

100:
        isFieldinRecordset = False
        Resume 200

    End Function

    Public Function isFieldinCollection(aColl As clsLib_DataAccess.clsFields, ByVal aKey As String) As Boolean

        Dim aVar As Object
        Dim i As Integer

        On Error GoTo 100

        For i = 1 To aColl.Count
            If (StrComp(aColl(i).Name, aKey, vbTextCompare) = 0) Then
                isFieldinCollection = True
                Exit Function
            End If
        Next i

        isFieldinCollection = False

200:
        Exit Function

100:
        isFieldinCollection = False
        Resume 200

    End Function

    Public Sub CenterForm(ByVal frm As Form)

        frm.SetBounds((System.Windows.Forms.Screen.GetBounds(frm).Width / 2) - (frm.Width / 2), _
            (System.Windows.Forms.Screen.GetBounds(frm).Height / 2) - (frm.Height / 2), _
            frm.Width, frm.Height, System.Windows.Forms.BoundsSpecified.Location)

    End Sub

    Public Function GetIncrement(Sequence As String) As Integer

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim sql As String
        Dim results As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        Dim lIncrement As Integer
        GetIncrement = -1

        sql = "SELECT NEXT_ID FROM INCREMENT WHERE SEQUENCE = |" & Sequence & "|"
        aRS = gDataLayer.LoadRecordset(sql, gObjErrors)

        If (gObjErrors.Count = 0) Then

            If (Not aRS.EOF) Then

                GetIncrement = aRS(0).Value

                sql = "UPDATE INCREMENT SET NEXT_ID = NEXT_ID + 1 WHERE SEQUENCE= |" & Sequence & "|"
                If (gDataLayer.ExecuteSQL(sql, gObjErrors, results)) Then
                Else
                    DisplayErrors()
                End If

            Else

                MsgBox("Auto Increment Sequence " & Sequence & " Not Found.", vbInformation)

            End If

        Else

            DisplayErrors()

        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function Max(ByVal X As Object, ByVal Y As Object) As Object

        If (X > Y) Then
            Max = X
        Else
            Max = Y
        End If

    End Function

    Public Function Max(ByVal X As Date, ByVal Y As Date) As Date

        If (X > Y) Then
            Max = X
        Else
            Max = Y
        End If

    End Function

    Public Function Min(ByVal X As Object, ByVal Y As Object) As Object

        If (X < Y) Then
            Min = X
        Else
            Min = Y
        End If

    End Function

    Public Function Min(ByVal X As Date, ByVal Y As Date) As Date

        If (X < Y) Then
            Min = X
        Else
            Min = Y
        End If

    End Function

    Public Function IntlDateFormat(ByVal aDate As Date) As String

        gBF = New CSOFT_BASIC_FUNCTIONS.clsBasicFunctions(gDebugMode) 'DP Added 7/13/2022 because things like IntlDateFormat methods within this dll just call the gBF.IntlDateFormat anyway, and that was not working.

        Return gBF.IntlDateFormat(aDate)

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        'IntlDateFormat = Format(aDate, "YYYY-MMM-DD HH:NN:SS")
        'IntlDateFormat = Format(aDate, "yyyy-MMM-dd HH:mm:ss")
        IntlDateFormat = aDate.ToString("yyyy-MMM-dd HH:mm:ss")

        Select Case Month(aDate)

            Case 1
                IntlDateFormat = Replace(IntlDateFormat, "janv", "Jan", 1, -1, vbTextCompare)

            Case 2
                IntlDateFormat = Replace(IntlDateFormat, "f�vr", "Feb", 1, -1, vbTextCompare)

            Case 3
                IntlDateFormat = Replace(IntlDateFormat, "mars", "Mar", 1, -1, vbTextCompare)

            Case 4
                IntlDateFormat = Replace(IntlDateFormat, "avr", "Apr", 1, -1, vbTextCompare)

            Case 5
                IntlDateFormat = Replace(IntlDateFormat, "mai", "May", 1, -1, vbTextCompare)

            Case 6
                IntlDateFormat = Replace(IntlDateFormat, "juin", "Jun", 1, -1, vbTextCompare)

            Case 7
                IntlDateFormat = Replace(IntlDateFormat, "juil", "Jul", 1, -1, vbTextCompare)

            Case 8
                IntlDateFormat = Replace(IntlDateFormat, "ao�t", "Aug", 1, -1, vbTextCompare)

            Case 9
                IntlDateFormat = Replace(IntlDateFormat, "sept", "Sep", 1, -1, vbTextCompare)

            Case 10
                IntlDateFormat = Replace(IntlDateFormat, "oct", "Oct", 1, -1, vbTextCompare)

            Case 11
                IntlDateFormat = Replace(IntlDateFormat, "nov", "Nov", 1, -1, vbTextCompare)

            Case 12
                IntlDateFormat = Replace(IntlDateFormat, "d�c", "Dec", 1, -1, vbTextCompare)

        End Select

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Sub TurnMeOff()

        'TurnMeOff
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Sub TurnMeOn()

        'TurnMeOff
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

    End Sub

    Public Sub DisplayErrors(Optional ByVal Number As Integer = 0, Optional ByVal Description As String = "", Optional ByVal Source As String = "")

        Dim lngErrorNumber As Integer
        Dim T As String
        Dim ii As Integer
        Dim jj As Integer
        Dim tbl As String
        Dim xtraMsg As String

        TurnMeOff()

        If gObjErrors.Count > 0 Then

            For lngErrorNumber = 1 To gObjErrors.Count

                If gObjErrors(lngErrorNumber).Show = True Then

                    If ((InStr(1, gObjErrors(lngErrorNumber).Description, "The record cannot be deleted", vbTextCompare) > 0) And _
                        (InStr(1, gObjErrors(lngErrorNumber).Description, "includes related records", vbTextCompare) > 0)) Then

                        ii = InStr(1, gObjErrors(lngErrorNumber).Description, "'", vbTextCompare)
                        jj = InStr(ii + 1, gObjErrors(lngErrorNumber).Description, "'", vbTextCompare)
                        tbl = Mid(gObjErrors(lngErrorNumber).Description, ii + 1, jj - ii - 1)

                        Select Case tbl

                            Case "PRODUCTION_DETAILS"
                                T = "production entries for employees"
                            Case "PRODUCTION_DETAILS_CONTRACTORS"
                                T = "production entries for contractors"
                            Case "TIME_DETAILS", "BLOCK_COST_DETAIL", "EMPLOYEE_STATEMENT_DETAILS"
                                T = "time entries for employees"
                            Case "TIME_DETAILS_EQUIP"
                                T = "time entries for equipment"
                            Case "TIME_DETAILS_CONTRACTORS"
                                T = "time entries for contractors"
                            Case "LOAD_SLIPS", "LOAD_TRUCK_COST", "LOADSLIP_PAY_ACTIVITIES", "LOADSLIP_REVENUE"
                                T = "loads"
                            Case "MISCELLANEOUS_INCOME"
                                T = "miscellaneous income entries"
                            Case "MISCELLANEOUS_EXPENSE"
                                T = "miscellaneous expense entries"
                            Case "EMPLOYEE_EXPENSES"
                                T = "employee expense entries"
                            Case "BENEFIT_RATES"
                                T = "benefit rates"
                            Case "BUDGET_ACTIVITY_RATES"
                                T = "budget activity rates"
                            Case "EMPLOYEE_BASE_RATES"
                                T = "employee base pay rates"
                            Case "REVENUE_CONTRACT_RATES"
                                T = "revenue contract rates"
                            Case "PAY_CONTRACT_RATES"
                                T = "pay contract rates"
                            Case "VENDOR_STATEMENT_DETAILS", "INVOICE_DETAILS"
                                T = "time, production or load entries"
                            Case "TRUCK_RATES"
                                T = "truck charge out rates"
                            Case "EQUIPMENT_RATES"
                                T = "equipment charge out rates"
                            Case "EMPLOYEE_RATES"
                                T = "employee charge out rates"
                            Case "CUSTOMER_ADVANCES"
                                T = "customer advances"
                            Case Else
                                T = ""

                        End Select

                        If (Len(T) > 0) Then
                            xtraMsg = vbCrLf & "You already have " & T & " entered using this data value.  You cannot make this deletion unless there are NO " & T & " using this value"
                        Else
                            xtraMsg = ""
                        End If

                        gObjErrors(lngErrorNumber).Description = gObjErrors(lngErrorNumber).Description & xtraMsg

                    End If

                    If (InStr(1, gObjErrors(lngErrorNumber).Description, "deadlock", vbTextCompare) > 0) Then

                        gObjErrors(lngErrorNumber).Description = "Your Current Action Conflicted with Another User.  Please Try Again."

                    End If

                    Select Case gObjErrors(lngErrorNumber).Number

                        Case 0

                            MsgBox("Error: " & gObjErrors(lngErrorNumber).Description, vbCritical, gObjErrors(lngErrorNumber).Source)

                        Case -2147217887, -2147217873

                            If (InStr(1, gObjErrors(lngErrorNumber).Description, "PRIMARY", vbTextCompare) > 0) Then
                                If Len(Source) > 0 Then
                                    MsgBox("An existing entry already exists for " & Source & ".  Please change the value to a unique entry and save.", vbCritical, gVars.gApp_Title)
                                Else
                                    MsgBox("An existing entry already exists.  Please change the value to a unique entry and save.", vbCritical, gVars.gApp_Title)
                                End If
                            ElseIf (InStr(1, gObjErrors(lngErrorNumber).Description, "related records", vbTextCompare) > 0) Then
                                T = Replace(gObjErrors(lngErrorNumber).Description, "related records", "entries that depend on the entry you are deleting.  You must first delete the dependent entries before deleting the current entry.")
                                MsgBox(T, vbCritical, gVars.gApp_Title)
                            Else
                                MsgBox("Error: " & gObjErrors(lngErrorNumber).Description, vbCritical, gObjErrors(lngErrorNumber).Source)
                            End If

                            'MsgBox "A matching entry already exists for " & Source & ".  Please change the value and re-save.", vbCritical, gVars.gApp_Title 

                        Case -2147467259

                            If (InStr(1, gObjErrors(lngErrorNumber).Description, "Could Not Find File", vbTextCompare) > 0) Then
                                MsgBox("Check Your Database File Location: " & gObjErrors(lngErrorNumber).Description & vbCrLf & "Use the 'File | Database Operations | Switch Database' menu item to set the location of your database.", vbCritical, gObjErrors(lngErrorNumber).Source)
                            Else
                                MsgBox(gObjErrors(lngErrorNumber).Number & ":" & gObjErrors(lngErrorNumber).Description, vbCritical, gObjErrors(lngErrorNumber).Source)
                            End If

                        Case Else

                            MsgBox(gObjErrors(lngErrorNumber).Number & ":" & gObjErrors(lngErrorNumber).Description, vbCritical, gObjErrors(lngErrorNumber).Source)

                    End Select

                End If

            Next lngErrorNumber

        Else

            If (Number <> 0) Then
                MsgBox("Error #" & Number & ":" & Description, vbCritical, Source)
            Else
                If Len(Description) > 0 Then
                    MsgBox("Error: " & Description, vbCritical, Source)
                End If
            End If
        End If

        gObjErrors = New clsLib_DataAccess.clsErrors

    End Sub

End Module
