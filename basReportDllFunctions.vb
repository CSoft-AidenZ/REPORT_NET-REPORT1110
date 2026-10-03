Option Explicit On

Imports CSOFT_RECORDSET_EXT
Imports clsLib_DataAccess

Module basReportDllFunctions

    Public Function getInSQLforCPartTable(ByVal aDest As String, ByVal partNo As Integer, ByVal aType As Integer) As String

        '------------------------------- 
        '  AType: 1 = Headers 
        '         2 = Detail 
        '         3 = Both 
        '------------------------------- 

        Dim inSQL As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim n As Integer
        Dim WhereClause As String = ""

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        Select Case aType

            Case 1
                WhereClause = " AND IS_HEADER <> 0 "
            Case 2
                WhereClause = " AND IS_DETAIL <> 0 "
            Case 3
                WhereClause = ""

        End Select

        inSQL = "SELECT COUNT(*) FROM DESTINATION_CPARTS T0 INNER JOIN C_PARTS T1 ON T0.LOAD_ITEM = T1.DESCRIPTION WHERE T0.IS_ACTIVE <> 0 AND T1.PART_NUMBER = " & partNo & _
                  " AND DESTINATION_CODE = |" & aDest & "| " & WhereClause
        aRS = gDataLayer.LoadRecordset(inSQL, gObjErrors)
        If gObjErrors.Count > 0 Then
            gBF.DisplayErrors(BasLocals.gobjErrors, BasLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return ""
            Exit Function
        End If

        n = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

        If n > 0 Then
            inSQL = "SELECT ITEM_VALUE FROM DESTINATION_CPARTS T0 INNER JOIN C_PARTS T1 ON T0.LOAD_ITEM = T1.DESCRIPTION WHERE T0.IS_ACTIVE <> 0 AND T1.PART_NUMBER = " & partNo & _
                    " AND DESTINATION_CODE = |" & aDest & "| " & WhereClause
        Else
            inSQL = ""
        End If

        getInSQLforCPartTable = inSQL

        Exit Function

Err_Handler:
        gbf.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetPayPeriod(ByVal ppType As String, ByVal doLoads As Integer, ByVal aDate As Date, ByVal BLOCK_CODE As String, ByVal DESTINATION_CODE As String) As String

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim nwPPCode As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        nwPPCode = ""

        '--------------------------------------------------------------------- 
        'Get the Pay Period it should be paid in 
        '--------------------------------------------------------------------- 
        If (doLoads And Val(gVars.gBasicSetup.Value("ENABLE_PAY_BY_CUSTOMER")) <> 0) Then

            strSQL = " SELECT PAY_PERIOD_TYPE_CODE FROM ((( CUSTOMERS C " &
                        " INNER JOIN REVENUE_CONTRACTS R ON R.CUSTOMER_CODE = C.CODE) " &
                        " INNER JOIN REVENUE_CONTRACT_BLOCKS RB ON R.CODE = RB.REVENUE_CONTRACT_CODE) " &
                        " INNER JOIN REVENUE_CONTRACT_DESTINATIONS DS ON R.CODE = DS.REVENUE_CONTRACT_CODE) " &
                        " WHERE RB.BLOCK_CODE = |" & BLOCK_CODE & "| AND DS.DESTINATION_CODE = |" & DESTINATION_CODE & "|"

            aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                gBF.TurnMeOff()
                Return ""
                Exit Function
            End If

            If (Not aRS.EOF) Then

                strSQL = " SELECT CODE FROM PAY_PERIODS WHERE START_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI & " AND " &
                            " DATEADD(|D|,1,END_DATE) > " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                            " AND PAY_PERIOD_TYPE_CODE = |" & CStr(aRS(0).Value) & "| AND CODE <> |SPOT| "

            Else
                gBF.TurnMeOff()
                MsgBox("No Valid Period Setup.  You must first setup a VENDOR/CUSTOMER pay period that encompasses the pay date: " & aDate & " of the Load.")
                Return ""
                Exit Function
            End If

        Else

            strSQL = " SELECT CODE FROM PAY_PERIODS WHERE START_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI & " AND " &
                        " DATEADD(|D|,1,END_DATE) > " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                        " AND PAY_PERIOD_TYPE_CODE IN (|" & ppType & "|,|ALL|) AND CODE <> |SPOT| "

        End If

        xRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            gBF.TurnMeOff()
            Return ""
            Exit Function
        End If

        If (xRS.EOF) Then
            gBF.TurnMeOff()
            MsgBox("No Valid Period Setup.  You must first setup a VENDOR pay period that encompasses the pay date: " & aDate & " of the Entry.")
            Return ""
            Exit Function
        Else
            nwPPCode = xRS(0).Value
        End If

        GetPayPeriod = nwPPCode

        xRS = Nothing
        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function getLoadUoM(ByVal BlCode As String, ByVal DstCode As String) As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Val(gVars.gBasicSetup.Value("IS_USE_DEST_FOR_UOMS")) <> 0 Then
            getLoadUoM = gComFn.getAValue("UOM_CODE", "DESTINATIONS", "CODE", DstCode)
        Else
            getLoadUoM = gComFn.getAValue("LOAD_UOM_CODE", "BLOCKS", "CODE", BlCode)
        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function BuildActiveCParts() As Collection

        Dim cParts As Collection
        Dim cpRs As clsRecordsetExt
        Dim aPart As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        cpRs = gDataLayer.LoadRecordset("SELECT * FROM C_PARTS WHERE IS_ACTIVE <> 0 AND PART_NUMBER > " & MAX_PRIME_PARTS & " ORDER BY PART_NUMBER ", gObjErrors)

        cParts = New Collection

        aPart = New clsLib_DataAccess.clsCPart

        aPart.PartNo = BLOCK_CODE
        aPart.PartDescription = "BLOCK_CODE"

        cParts.Add(aPart, CStr(BLOCK_CODE))

        aPart = New clsLib_DataAccess.clsCPart

        aPart.PartNo = EQUIPMENT_CODE
        aPart.PartDescription = "EQUIPMENT_CODE"

        cParts.Add(aPart, CStr(EQUIPMENT_CODE))

        aPart = New clsLib_DataAccess.clsCPart

        aPart.PartNo = DESTINATION
        aPart.PartDescription = "DESTINATION_CODE"

        cParts.Add(aPart, CStr(DESTINATION))


        Do While cpRs.EOF = False

            aPart = New clsLib_DataAccess.clsCPart

            aPart.PartNo = cpRs("PART_NUMBER").Value
            aPart.PartDescription = cpRs("C_PART").Value

            cParts.Add(aPart, CStr(cpRs("PART_NUMBER").Value))
            cpRs.MoveNext()

        Loop

        cpRs.Close1()

        If gObjtables("LOADSLIPS").Fields("ROUTE_CODE").DisplayOrder > 0 Then
            aPart = New clsLib_DataAccess.clsCPart

            aPart.PartNo = ROUTE
            aPart.PartDescription = "ROUTE_CODE"

            cParts.Add(aPart, CStr(ROUTE))
        End If

        If gObjtables("LOADSLIPS").Fields("TRUCK_TYPE_CODE").DisplayOrder > 0 Then
            aPart = New clsLib_DataAccess.clsCPart

            aPart.PartNo = TRUCK_TYPE
            aPart.PartDescription = "TRUCK_TYPE_CODE"

            cParts.Add(aPart, CStr(TRUCK_TYPE))
        End If

        BuildActiveCParts = cParts

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function buildCPartColl(ByVal aRS As clsRecordsetExt) As Collection

        Dim i As Integer
        Dim cPartColl As Collection

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        cPartColl = New Collection
        For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
            cPartColl.Add(CStr(aRS("C_PART_" & i & "_CODE").Value), CStr(i))
        Next i
        cPartColl.Add(CStr(aRS("DESTINATION_CODE").Value), CStr(DESTINATION))
        cPartColl.Add(CStr(aRS("TRUCK_TYPE_CODE").Value), CStr(TRUCK_TYPE))
        cPartColl.Add(CStr(aRS("ROUTE_CODE").Value), CStr(ROUTE))

        buildCPartColl = cPartColl

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetWeekCode(ByVal TIME_SLIP_DATE As String, ByRef errorStr As String) As String

        Dim aRS As clsRecordsetExt
        Dim strSQL As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        strSQL = "SELECT TOP 1 CODE FROM WEEKS WHERE THE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(CDate(TIME_SLIP_DATE)) & gVars.gDI & " ORDER BY THE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return ""
            Exit Function
        End If

        If aRS.EOF Then
            errorStr = "Week Code not Found for: " & TIME_SLIP_DATE
            Return ""
        Else
            'Base Pay Rate 
            GetWeekCode = aRS("CODE").Value
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)


    End Function

    Public Function GetBaseRateFactor(ByRef isPercent As Integer, ByVal EMPLOYEE_CODE As String, ByVal EQUIPMENT_CODE As String, ByVal TIME_SLIP_DATE As Date, ByVal PAY_ACTIVITY_CODE As String, ByVal RATE_TYPE_CODE As String, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        '    strSQL = "SELECT TOP 1 PAY_RATE, IS_PERCENT FROM EMPLOYEE_RATES ER INNER JOIN EQUIPMENT EQ ON ER.EQUIPMENT_CODE = EQ.CODE " & _ 
        '             " WHERE RATE_TYPE_CODE = |" & RATE_TYPE_CODE) & "| AND EMPLOYEE_CODE = |" & EMPLOYEE_CODE) & "|" & _ 
        '             " AND PHASE_CODE = |" & PAY_ACTIVITY_CODE) & "| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & _ 
        '             " AND ( EQUIPMENT_CODE = |" & EQUIPMENT_CODE) & "| OR EQUIPMENT_CODE = |NONE| ) " & _ 
        '             " ORDER BY EQ.IS_NONE ASC, EFFECTIVE_DATE DESC " 

        'Bob 3/6/06 Invoked ALL for Employee Code 
        strSQL = "SELECT TOP 1 PAY_RATE, IS_PERCENT FROM (( EMPLOYEE_RATES ER " &
                 " INNER JOIN EQUIPMENT EQ ON ER.EQUIPMENT_CODE = EQ.CODE) " &
                 " INNER JOIN EMPLOYEES EM ON ER.EMPLOYEE_CODE = EM.CODE) " &
                 " WHERE RATE_TYPE_CODE = |" & RATE_TYPE_CODE & "|" &
                 " AND PHASE_CODE = |" & PAY_ACTIVITY_CODE & "| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI &
                 " AND ( EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "| OR EQUIPMENT_CODE = |NONE| ) " &
                 " AND ( EMPLOYEE_CODE = |" & EMPLOYEE_CODE & "| OR EMPLOYEE_CODE = |ALL|) " &
                 " ORDER BY EM.IS_ALL ASC, EQ.IS_NONE ASC, EFFECTIVE_DATE DESC "

        'Bob 1/24/05 removed the is_active flag 
        ' " AND IS_ACTIVE <> 0 " & _ 

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then
            'ErrorStr = "Rate not found for Employee, Activity, Date: " & EMPLOYEE_CODE & "," & PAY_ACTIVITY_CODE & "," & TIME_SLIP_DATE 
            GetBaseRateFactor = 0
            isPercent = 0
        Else
            'Base Pay Rate 
            GetBaseRateFactor = aRS("PAY_RATE").Value
            isPercent = aRS("IS_PERCENT").Value
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function calcAvgCost(ByVal Q As Double, ByVal PartCost_NO_Core As Double, ByVal Freight As Double, ByVal isApplyTax As Integer) As Double

        Dim TCost As Double
        Dim TaxOnFreight As Double
        Dim TaxOnPart As Double
        Dim SalesTaxRate As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        calcAvgCost = 0

        If isApplyTax Then
            SalesTaxRate = gVars.gBasicSetup.Value("SALES_TAX_RATE")
        Else
            SalesTaxRate = 0
        End If

        TCost = PartCost_NO_Core

        If Val("" & gVars.gBasicSetup.Value("INCLUDE_SALES_TAX_IN_COSTS")) <> 0 Then

            TaxOnPart = SalesTaxRate * PartCost_NO_Core

        Else

            TaxOnPart = 0

        End If


        If Val("" & gVars.gBasicSetup.Value("INCLUDE_FREIGHT_TAX_IN_COSTS")) <> 0 Then

            TCost = TCost + Freight

            '------------------------------------------------------------------------------ 
            '  Tax on Freight only included if: 
            '  1.  Freight is included in Costs 
            '  2.  Sales Tax applies to freight 
            '  3.  Sales Tax is included in costs 
            '------------------------------------------------------------------------------ 

            If Val("" & gVars.gBasicSetup.Value("APPLY_TAX_TO_FREIGHT")) <> 0 Then

                If Val("" & gVars.gBasicSetup.Value("INCLUDE_SALES_TAX_IN_COSTS")) <> 0 Then

                    TaxOnFreight = SalesTaxRate * Freight

                Else

                    TaxOnFreight = 0

                End If

            Else

                TaxOnFreight = 0

            End If

        Else

            TaxOnFreight = 0

        End If

        TCost = TCost + TaxOnFreight + TaxOnPart

        If Math.Abs(Q) > EPS Then
            calcAvgCost = TCost / Q
        Else
            'calcAvgCost = 0 
            calcAvgCost = TCost
        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        calcAvgCost = 0
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function getBlockStartEnd(ByRef blStr As String, ByRef blSD As Date, ByRef blED As Date) As Integer

        Dim xRS As clsRecordsetExt
        Dim fstDate As Date
        Dim lstDate As Date
        Dim strSQL As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt = Nothing

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        blSD = #1/1/2050#
        blED = #1/1/1900#

        strSQL = " SELECT MIN(T0.DATE_OUT) AS SD, MAX(T0.DATE_OUT) AS ED FROM LOADSLIPS T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = CDate(aRS(0).Value)
            blED = CDate(aRS(1).Value)

        End If

        strSQL = " SELECT MIN(T0.TIME_SLIP_DATE) AS SD, MAX(T0.TIME_SLIP_DATE) AS ED FROM TIME_DETAILS T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, CDate(aRS(0).Value))
            blED = gBF.Max(blED, CDate(aRS(1).Value))

        End If

        strSQL = " SELECT MIN(T0.TIME_SLIP_DATE) AS SD, MAX(T0.TIME_SLIP_DATE) AS ED FROM TIME_DETAILS_CONTRACTORS T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, aRS(0).Value)
            blED = gBF.Max(blED, aRS(1).Value)

        End If

        strSQL = " SELECT MIN(T0.TIME_SLIP_DATE) AS SD, MAX(T0.TIME_SLIP_DATE) AS ED FROM TIME_DETAILS_EQUIPMENT T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, aRS(0).Value)
            blED = gBF.Max(blED, aRS(1).Value)

        End If

        strSQL = " SELECT MIN(T0.TIME_SLIP_DATE) AS SD, MAX(T0.TIME_SLIP_DATE) AS ED FROM PRODUCTION_DETAILS T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, aRS(0).Value)
            blED = gBF.Max(blED, aRS(1).Value)

        End If

        strSQL = " SELECT MIN(T0.TIME_SLIP_DATE) AS SD, MAX(T0.TIME_SLIP_DATE) AS ED FROM PRODUCTION_DETAILS_CONTRACTORS T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, aRS(0).Value)
            blED = gBF.Max(blED, aRS(1).Value)

        End If

        strSQL = " SELECT MIN(T0.ENTRY_DATE) AS SD, MAX(T0.ENTRY_DATE) AS ED FROM MISCELLANEOUS_EXPENSES T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, aRS(0).Value)
            blED = gBF.Max(blED, aRS(1).Value)

        End If

        strSQL = " SELECT MIN(T0.ENTRY_DATE) AS SD, MAX(T0.ENTRY_DATE) AS ED FROM EMPLOYEE_EXPENSES T0 WHERE T0.BLOCK_CODE IN ( " & blStr & ")"
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then

        Else

            blSD = gBF.Min(blSD, aRS(0).Value)
            blED = gBF.Max(blED, aRS(1).Value)

        End If

        getBlockStartEnd = -1

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetWeeklyRevenue(ByRef aColl1 As Collection, ByVal xsd As Date, ByVal xED As Date, ByVal doAll As Integer, ByVal useInvoiceRevenue As Integer) As Collection

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim fstDate As Date
        Dim lstDate As Date
        Dim blStr As String
        Dim aObs As clsAValue
        Dim i As Integer
        Dim aColl As Collection
        Dim revTable As String
        Dim DoW0 As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetWeeklyRevenue = New Collection

        If (aColl1.Count = 0) Then
            Exit Function
        End If

        revTable = "INVOICE_DETAILS"

        aColl = New Collection

        If doAll Then

            strSQL = " SELECT MIN(RV.ENTRY_DATE) AS SD, MAX(RV.ENTRY_DATE) AS ED FROM " & revTable & " RV  " &
                      " WHERE RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                      " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI
        Else

            blStr = ""
            For i = 1 To aColl1.Count
                blStr = blStr & "|" & CStr(aColl1(i)) & "|, "
            Next i
            blStr = Left(blStr, Len(blStr) - 2)

            strSQL = " SELECT MIN(RV.ENTRY_DATE) AS SD, MAX(RV.ENTRY_DATE) AS ED FROM " & revTable & " RV  " &
                      " WHERE RV.BLOCK_CODE IN (" & blStr & ") " &
                      " AND RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                      " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI
        End If

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        If (IsDBNull(aRS(0).Value)) Then
            Exit Function
        Else

            DoW0 = Weekday(xsd)  'Sunday is 1, Monday is 2 
            'xsd = CDate(Format(aRS(0).Value, "ddddd"))
            xsd = CDate(FormatDateTime(aRS(0).Value, DateFormat.ShortDate))

            Do Until Weekday(xsd) = DoW0
                xsd = New vb6Date(xsd) - 1
            Loop

            xED = New vb6Date(CDate(FormatDateTime(aRS(1).Value, DateFormat.ShortDate))) + 1

        End If


        fstDate = xsd

        lstDate = New vb6Date(fstDate) + 7

        Do While fstDate < xED

            If doAll Then

                strSQL = " SELECT SUM(RV.PAY) AS PAY FROM " & revTable & " RV  " &
                          " WHERE RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(lstDate) & gVars.gDI &
                          " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(fstDate) & gVars.gDI
            Else

                blStr = ""
                For i = 1 To aColl1.Count
                    blStr = blStr & "|" & CStr(aColl1(i)) & "|, "
                Next i
                blStr = Left(blStr, Len(blStr) - 2)

                strSQL = " SELECT SUM(RV.PAY) AS PAY FROM " & revTable & " RV  " &
                          " WHERE RV.BLOCK_CODE IN (" & blStr & ") " &
                          " AND RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(lstDate) & gVars.gDI &
                          " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(fstDate) & gVars.gDI
            End If

            aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                Exit Function
            End If

            aObs = New clsAValue
            aObs.dblValue = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
            aObs.strValue = "WE: " & Format(lstDate, "mm/dd")

            aColl.Add(aObs, CStr(lstDate))

            fstDate = New vb6Date(fstDate) + 7
            lstDate = New vb6Date(fstDate) + 7

        Loop

        GetWeeklyRevenue = aColl

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetLoads_ByBlockSP(ByVal blStr As String, ByVal xsd As Date, ByVal xED As Date) As Collection

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim fstDate As Date
        Dim lstDate As Date
        Dim aObs As clsAValue
        Dim i As Integer
        Dim aColl As Collection

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetLoads_ByBlockSP = New Collection

        aColl = New Collection

        strSQL = " SELECT SUM(LS.NET) AS KGS, SUM(LS.NET/1000) AS TONS, LS.C_PART_4_CODE, LS.C_PART_5_CODE FROM LOADSLIPS LS " &
                  " WHERE LS.BLOCK_CODE IN (" & blStr & ") " &
                  " AND LS.DATE_OUT <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                  " AND LS.DATE_OUT >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI &
                  " GROUP BY LS.C_PART_4_CODE, LS.C_PART_5_CODE "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False

            aObs = New clsAValue
            aObs.dblValue = IIf(IsDBNull(aRS(1).Value), 0, aRS(1).Value) 'Tons 

            aObs.lngValue = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value) 'KGs 

            aObs.strValue = aRS(1).Value

            aColl.Add(aObs, CStr(aRS(2).Value & ":" & aRS(3).Value))

            aRS.MoveNext()

        Loop

        GetLoads_ByBlockSP = aColl

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetRevenueByActivity_ByBlockSP(ByVal blStr As String, ByVal actStr As String, ByVal xsd As Date, ByVal xED As Date) As Collection

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim fstDate As Date
        Dim lstDate As Date
        Dim aObs As clsAValue
        Dim i As Integer
        Dim aColl As Collection
        Dim revTable As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetRevenueByActivity_ByBlockSP = New Collection

        revTable = "INVOICE_DETAILS"

        aColl = New Collection

        strSQL = " SELECT SUM(RV.PAY) AS PAY, LS.C_PART_4_CODE, LS.C_PART_5_CODE FROM " & revTable & " RV  " &
                  " INNER JOIN LOADSLIPS LS ON LS.CODE = RV.LOAD_CODE " &
                  " WHERE RV.PAY_TYPE IN (" & REV_CONTRACT_LOADS & ", " & REV_CONTRACT_CUSTOMER_LOADS & ") " &
                  " AND LS.BLOCK_CODE IN (" & blStr & ") " &
                  " AND LS.DATE_OUT <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                  " AND LS.DATE_OUT >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI &
                  " AND RV.ACTIVITY_CODE = |" & actStr & "|" &
                  " GROUP BY LS.C_PART_4_CODE, LS.C_PART_5_CODE "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False

            aObs = New clsAValue
            aObs.dblValue = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

            aObs.strValue = aRS(1).Value

            aColl.Add(aObs, CStr(aRS(1).Value & ":" & aRS(2).Value))

            aRS.MoveNext()

        Loop

        GetRevenueByActivity_ByBlockSP = aColl

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetRevenueByActivity_ByBlock(ByVal blStr As String, ByVal xsd As Date, ByVal xED As Date) As Collection

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim fstDate As Date
        Dim lstDate As Date
        Dim aObs As clsAValue
        Dim i As Integer
        Dim aColl As Collection
        Dim revTable As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetRevenueByActivity_ByBlock = New Collection

        revTable = "INVOICE_DETAILS"

        aColl = New Collection

        strSQL = " SELECT SUM(RV.PAY) AS PAY, RV.ACTIVITY_CODE FROM " & revTable & " RV  " &
                  " WHERE RV.BLOCK_CODE IN (" & blStr & ") " &
                  " AND RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                  " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI &
                  " GROUP BY RV.ACTIVITY_CODE " &
                  " ORDER BY RV.ACTIVITY_CODE "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False

            aObs = New clsAValue
            aObs.dblValue = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

            aObs.strValue = aRS(1).Value

            aColl.Add(aObs, CStr(aRS(1).Value))

            aRS.MoveNext()

        Loop

        GetRevenueByActivity_ByBlock = aColl

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetRevenueByBlock(ByRef aColl1 As Collection, ByVal blStr As String, ByVal xsd As Date, ByVal xED As Date, ByVal doAll As Integer, ByVal useInvoiceRevenue As Integer) As Collection

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim fstDate As Date
        Dim lstDate As Date
        Dim aObs As clsAValue
        Dim i As Integer
        Dim aColl As Collection
        Dim revTable As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetRevenueByBlock = New Collection

        If (aColl1.Count = 0) Then
            Exit Function
        End If

        revTable = "INVOICE_DETAILS"

        aColl = New Collection

        If doAll Then

            strSQL = " SELECT SUM(RV.PAY) AS PAY, RV.BLOCK_CODE FROM " & revTable & " RV  " &
                      " WHERE RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                      " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI &
                      " GROUP BY RV.BLOCK_CODE " &
                      " ORDER BY SUM(RV.PAY) DESC "

        Else

            strSQL = " SELECT SUM(RV.PAY) AS PAY, RV.BLOCK_CODE FROM " & revTable & " RV  " &
                      " WHERE RV.BLOCK_CODE IN (" & blStr & ") " &
                      " AND RV.ENTRY_DATE <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI &
                      " AND RV.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI &
                      " GROUP BY RV.BLOCK_CODE " &
                      " ORDER BY SUM(RV.PAY) DESC "
        End If

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False

            aObs = New clsAValue
            aObs.dblValue = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

            aObs.strValue = aRS(1).Value

            aColl.Add(aObs, CStr(aRS(1).Value))

            aRS.MoveNext()

        Loop

        GetRevenueByBlock = aColl

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function




    Public Function GetTruckTareWeight(ByVal TRUCK_CODE As String, ByVal ROUTE_CODE As String, ByVal TRUCK_TYPE_CODE As String,
                                       ByVal DESTINATION_CODE As String, ByVal Date_In As Date) As Double

        '-------------------------------------- 
        '  Return NA if NO Minimum Weight 
        '-------------------------------------- 

        Dim strSQL As String
        Dim aRS As clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetTruckTareWeight = 0

        strSQL = "SELECT TOP 1 T1.STANDARD_TARE FROM (((( TRUCK_MAX_WEIGHTS T1 " &
                  " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE ) " &
                  " INNER JOIN DESTINATIONS D ON D.CODE = T1.DESTINATION_CODE ) " &
                  " INNER JOIN ROUTES R ON R.CODE = T1.ROUTE_CODE ) " &
                  " INNER JOIN TRUCK_TYPES TT ON TT.CODE = T1.TRUCK_TYPE_CODE ) " &
                  " WHERE (T1.TRUCK_CODE = |" & TRUCK_CODE & "| OR T1.TRUCK_CODE = |NONE|) " &
                  " AND (T1.ROUTE_CODE = |" & ROUTE_CODE & "| OR T1.ROUTE_CODE = |NONE|) " &
                  " AND (T1.TRUCK_TYPE_CODE = |" & TRUCK_TYPE_CODE & "| OR T1.TRUCK_TYPE_CODE = |NONE|)" &
                  " AND (T1.DESTINATION_CODE = |" & DESTINATION_CODE & "| OR T1.DESTINATION_CODE = |NONE|)" &
                  " AND T1.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(Date_In) & gVars.gDI &
                  " ORDER BY TR.IS_NONE ASC, D.IS_NONE, R.IS_NONE ASC, TT.IS_NONE ASC, T1.EFFECTIVE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        '-------------------------------------- 
        '  Return 0 if NO Minimum Weight Setup 
        '-------------------------------------- 

        If (aRS.EOF) Then
            GetTruckTareWeight = 0
            Exit Function
        End If

        GetTruckTareWeight = aRS("STANDARD_TARE").Value

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function LoadCPartCollection() As Collection

        Dim aRS As clsRecordsetExt
        Dim i As Integer
        Dim strSQL As String
        Dim aPart As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadCPartCollection = New Collection

        aRS = New clsRecordsetExt
        strSQL = "SELECT DESCRIPTION, PART_NUMBER FROM C_PARTS WHERE IS_ACTIVE <> 0 AND PART_NUMBER > " & MAX_PRIME_PARTS & " ORDER BY PART_NUMBER "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False
            aPart = New clsLib_DataAccess.clsCPart
            aPart.PartDescription = CStr(aRS("DESCRIPTION").Value)
            aPart.PartNo = aRS("PART_NUMBER").Value
            LoadCPartCollection.Add(aPart)
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)


    End Function

    Public Function GetTripTime(ByVal aLoad As clsRecordsetExt, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String
        Dim JoinStr As String
        Dim OrderByStr As String
        Dim whereStr As String
        Dim i As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt

        'strSql = "SELECT AVG_TIME FROM BLOCK_DISTANCES WHERE BLOCK_CODE = |" & aLoad("BLOCK_CODE").Value & "| AND DESTINATION_CODE = |" & aLoad("DESTINATION_CODE").Value & "|" 

        JoinStr = " BLOCK_DISTANCES T0 "
        OrderByStr = "ORDER BY T0.BLOCK_CODE, "
        whereStr = " WHERE T0.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aLoad("DATE_OUT").Value) & gVars.gDI & " AND T0.BLOCK_CODE = |" & aLoad("BLOCK_CODE").Value & "|" &
                      " AND T0.DESTINATION_CODE = |" & aLoad("DESTINATION_CODE").Value & "| "

        For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS

            If (gObjtables("BLOCK_DISTANCES").Fields("C_PART_" & i & "_CODE").DisplayOrder > 0) Then
                JoinStr = "(" & JoinStr & " INNER JOIN C_PART_" & i & " AS T" & i & " ON T" & i & ".CODE = T0.C_PART_" & i & "_CODE ) "
                whereStr = whereStr & " AND (C_PART_" & i & "_CODE = |" & aLoad("C_PART_" & i & "_CODE").Value & "| OR C_PART_" & i & "_CODE = |NONE| )"
                OrderByStr = OrderByStr & " T" & i & ".IS_NONE ASC, "
            End If

        Next i

        If (gObjtables("BLOCK_DISTANCES").Fields("TRUCK_TYPE_CODE").DisplayOrder > 0) Then
            JoinStr = "(" & JoinStr & " INNER JOIN TRUCK_TYPES AS TT ON TT.CODE = T0.TRUCK_TYPE_CODE ) "
            whereStr = whereStr & " AND (TRUCK_TYPE_CODE = |" & aLoad("TRUCK_TYPE_CODE").Value & "| OR TRUCK_TYPE_CODE = |NONE| )"
            OrderByStr = OrderByStr & " TT.IS_NONE ASC, "
        End If

        If (gObjtables("BLOCK_DISTANCES").Fields("ROUTE_CODE").DisplayOrder > 0) Then
            JoinStr = "(" & JoinStr & " INNER JOIN ROUTES AS RO ON RO.CODE = T0.ROUTE_CODE ) "
            whereStr = whereStr & " AND (ROUTE_CODE = |" & aLoad("ROUTE_CODE").Value & "| OR ROUTE_CODE = |NONE| )"
            OrderByStr = OrderByStr & " RO.IS_NONE, "
        End If

        'OrderByStr = VBA.Left(OrderByStr, Len(OrderByStr) - 2) 
        OrderByStr = OrderByStr & " T0.EFFECTIVE_DATE DESC "

        strSQL = "SELECT AVG_TIME FROM " & JoinStr & whereStr & OrderByStr

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (aRS.EOF) Then
            GetTripTime = 0
            '~errorStr = "Error #9001: Avg Trip Time not Setup for " & gVars.gBlock_Name & " [" & aLoad("BLOCK_CODE").Value & "] and Destination [" & aLoad("DESTINATION_CODE").Value & "]" 
            errorStr = "Error #9001: Avg Trip Time not Setup From " & aLoad("BLOCK_CODE").Value & " To " & aLoad("DESTINATION_CODE").Value
            Return 0
        Else
            GetTripTime = aRS(0).Value
        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    'Public Function GetTripTime_clsLoad(aLoad As clsLib_Csoft_Classes.clsLoad, ByRef ErrorStr As String)
    '
    '    Dim aRS As clsRecordsetExt
    '    Dim strSql As String
    '    Dim JoinStr As String
    '    Dim OrderByStr As String
    '    Dim WhereStr As String
    '    Dim i as Integer
    '
    '    If (gDebugMode = 0) Then On Error GoTo Err_Handler
    '
    '    Set aRS = New clsRecordsetExt
    '
    '    'strSql = "SELECT AVG_TIME FROM BLOCK_DISTANCES WHERE BLOCK_CODE = |" & aLoad("BLOCK_CODE") & "| AND DESTINATION_CODE = |" & aLoad("DESTINATION_CODE") & "|"
    '
    '    JoinStr = " BLOCK_DISTANCES T0 "
    '    OrderByStr = "ORDER BY T0.BLOCK_CODE, "
    '    WhereStr = " WHERE T0.BLOCK_CODE = |" & aLoad.BLOCK_CODE & "| AND T0.DESTINATION_CODE = |" & aLoad.DESTINATION_CODE & "|"
    '
    '    For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
    '
    '        If (gobjTables("BLOCK_DISTANCES").Fields("C_PART_" & i & "_CODE").DisplayOrder > 0) Then
    '            JoinStr = "(" & JoinStr & " INNER JOIN C_PART_" & i & " AS T" & i & " ON T" & i & ".CODE = T0.C_PART_" & i & "_CODE ) "
    '            WhereStr = WhereStr & " AND (C_PART_" & i & "_CODE = |" & aLoad.aPart(i).PartCode & "| OR C_PART_" & i & "_CODE = |NONE| )"
    '            OrderByStr = OrderByStr & " T" & i & ".IS_NONE ASC, "
    '        End If
    '
    '    Next i
    '
    '    If (gobjTables("BLOCK_DISTANCES").Fields("TRUCK_TYPE_CODE").DisplayOrder > 0) Then
    '        JoinStr = "(" & JoinStr & " INNER JOIN TRUCK_TYPES AS TT ON TT.CODE = T0.TRUCK_TYPE_CODE ) "
    '        WhereStr = WhereStr & " AND (TRUCK_TYPE_CODE = |" & aLoad.TRUCK_TYPE_CODE & "| OR TRUCK_TYPE_CODE = |NONE| )"
    '        OrderByStr = OrderByStr & " TT.IS_NONE ASC, "
    '    End If
    '
    '    If (gobjTables("BLOCK_DISTANCES").Fields("ROUTE_CODE").DisplayOrder > 0) Then
    '        JoinStr = "(" & JoinStr & " INNER JOIN ROUTES AS RO ON RO.CODE = T0.ROUTE_CODE ) "
    '        WhereStr = WhereStr & " AND (ROUTE_CODE = |" & aLoad.ROUTE_CODE & "| OR ROUTE_CODE = |NONE| )"
    '        OrderByStr = OrderByStr & " RO.IS_NONE, "
    '    End If
    '
    '    OrderByStr = VBA.Left(OrderByStr, Len(OrderByStr) - 2)
    '    strSql = "SELECT AVG_TIME FROM " & JoinStr & WhereStr & OrderByStr
    '
    '    Set aRS = gDataLayer.LoadRecordset(strSql, gobjErrors)
    '    If (gobjErrors.Count = 0) Then
    '    Else
    '        DisplayErrors
    '        Exit Function
    '    End If
    '
    '    If (aRS.EOF) Then
    '        GetTripTime_clsLoad = 0
    '        ErrorStr = "Error #9001: Avg Trip Time not Setup for " & gVars.gBlock_Name & " [" & aLoad.BLOCK_CODE & "] and Destination [" & aLoad.DESTINATION_CODE & "]"
    '    Else
    '        GetTripTime_clsLoad = aRS(0)
    '    End If
    '
    '    Exit Function
    'Err_Handler:
    '    DisplayErrors Err.Number, Err.Description, Err.Source & ":basReportDllFunctions.GetLoadValue"
    '
    'End Function

    Public Function GetFuelPrice(ByVal TIME_SLIP_DATE As Date, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        strSQL = "SELECT TOP 1 PRICE FROM FUEL_PRICES WHERE EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then
            errorStr = "No Price"
            Return 0
        Else
            'Base Pay Rate 
            errorStr = ""
            GetFuelPrice = aRS("PRICE").Value
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetEquipmentCostRate_NonHourly(ByVal CostField As String, ByVal EQUIPMENT_CODE As String, ByVal TIME_SLIP_DATE As Date, ByVal RATE_TYPE_CODE As String, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String
        Dim deltaFuel_HR As Double
        Dim deltaFuel As Double
        Dim xFact As Double
        Dim errStr As String
        Dim fuelPrice As Double
        Dim eqCost As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt

        '------------------------------------------------------------------------------- 
        ' Len(RATE_TYPE_CODE) = 0 when recalculating based on the block cost table 
        '->  Assumes Rate Type has not changed 
        '------------------------------------------------------------------------------- 
        If Len(RATE_TYPE_CODE) = 0 Then
            strSQL = "SELECT TOP 1 " & CostField & ", FUEL_PRICE, FUEL_PER_HOUR, FLUIDS_PCT/100 AS FLUIDS, ANNUAL_PRODUCTION, ANNUAL_HOURS, UTILIZATION_PCT " &
                     " FROM EQUIPMENT_COST WHERE CODE = |" & EQUIPMENT_CODE & "|" &
                     " AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "
        Else
            strSQL = "SELECT TOP 1 " & CostField & ", FUEL_PRICE, FUEL_PER_HOUR, FLUIDS_PCT/100 AS FLUIDS, ANNUAL_PRODUCTION, ANNUAL_HOURS, UTILIZATION_PCT " &
                     " FROM EQUIPMENT_COST WHERE CODE = |" & EQUIPMENT_CODE & "|" &
                     " AND RATE_TYPE_CODE = |" & RATE_TYPE_CODE & "| " &
                     " AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "
        End If
        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then
            '~errorStr = "Error #1302: Equipment " & CostField & " Cost not found for Equipment Code: [" & EQUIPMENT_CODE & "] and Time Slip Date: [" & TIME_SLIP_DATE & "] and Rate Basis is [" & RATE_TYPE_CODE & "]" 
            errorStr = "Equipment " & CostField & " [" & RATE_TYPE_CODE & "] Cost Record Not Found for Equipment Code: [" & EQUIPMENT_CODE & "]"
            Return 0
        Else

            'Base Pay Rate 
            eqCost = aRS(CostField).Value

            If Math.Abs(eqCost) < EPS And Val(gVars.gBasicSetup.Value("WARN_IF_EQ_COST_IS_ZERO")) <> 0 Then
                errorStr = "Equipment " & CostField & " [" & RATE_TYPE_CODE & "] Cost is ZERO for Equipment Code: [" & EQUIPMENT_CODE & "]"
                Return 0
            Else

                errStr = ""
                fuelPrice = GetFuelPrice(TIME_SLIP_DATE, errStr)
                If Len(errStr) = 0 Then
                    'Recalc 
                    deltaFuel_HR = (fuelPrice - aRS("FUEL_PRICE").Value) * aRS("FUEL_PER_HOUR").Value * (1.0# + aRS("FLUIDS").Value) 'Cost Difference based on Difference in fuel prices 
                    If Math.Abs(aRS("ANNUAL_PRODUCTION").Value) > EPS Then
                        xFact = (aRS("ANNUAL_HOURS").Value * aRS("UTILIZATION_PCT").Value / 100.0#) / aRS("ANNUAL_PRODUCTION").Value
                    Else
                        xFact = 0.0#
                    End If
                    deltaFuel = xFact * deltaFuel_HR  'Covert to $ per unit of production rather than hours 
                Else
                    deltaFuel_HR = 0
                    'Skip 
                End If

                GetEquipmentCostRate_NonHourly = eqCost + deltaFuel

            End If

        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetEquipmentCostRate(ByVal CostField As String, ByVal EQUIPMENT_CODE As String, ByVal TIME_SLIP_DATE As Date, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String
        Dim fuelPrice As Double
        Dim errStr As String
        Dim deltaFuel As Double
        Dim eqCost As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        strSQL = "SELECT TOP 1 " & CostField & ", FUEL_PRICE, FUEL_PER_HOUR, FLUIDS_PCT/100 AS FLUIDS FROM EQUIPMENT_COST WHERE CODE = |" & EQUIPMENT_CODE & "|" &
                 " AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then
            '~errorStr = "Error #1302: Equipment " & CostField & " Cost not found for Equipment Code: [" & EQUIPMENT_CODE & "] and Time Slip Date: [" & TIME_SLIP_DATE & "]" 
            errorStr = "Equipment " & CostField & " Cost Record Not Found for Equipment Code: [" & EQUIPMENT_CODE & "]"
            Return 0
        Else

            'Base Pay Rate 
            eqCost = aRS(CostField).Value

            If Math.Abs(eqCost) < EPS And Val(gVars.gBasicSetup.Value("WARN_IF_EQ_COST_IS_ZERO")) <> 0 Then
                errorStr = "Equipment " & CostField & " Cost is ZERO for Equipment Code: [" & EQUIPMENT_CODE & "]"
                Return 0
            Else

                errStr = ""
                fuelPrice = GetFuelPrice(TIME_SLIP_DATE, errStr)
                If Len(errStr) = 0 Then
                    'Recalc 
                    deltaFuel = (fuelPrice - aRS("FUEL_PRICE").Value) * aRS("FUEL_PER_HOUR").Value * (1.0# + aRS("FLUIDS").Value)
                Else
                    deltaFuel = 0
                    'Skip 
                End If

            End If

            GetEquipmentCostRate = eqCost + deltaFuel

        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetTruckCostHourlyRate(ByVal CostField As String, ByVal TRUCK_CODE As String, ByVal TIME_SLIP_DATE As Date, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String
        Dim fuelPrice As Double
        Dim errStr As String
        Dim deltaFuel As Double
        Dim eqCost As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        strSQL = "SELECT TOP 1 " & CostField & ", FUEL_PRICE, FUEL_PER_HOUR, FLUIDS_PCT/100 AS FLUIDS FROM TRUCK_COST WHERE CODE = |" & TRUCK_CODE & "|" &
                 " AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then

            '~errorStr = "Error #1301: Truck " & CostField & "  not found for Truck Code: [" & TRUCK_CODE & "] and Time Slip Date: [" & TIME_SLIP_DATE & "]" 
            errorStr = "Truck " & CostField & " Cost Record Not Found for Truck Code: [" & TRUCK_CODE & "]"
            Return 0

        Else

            eqCost = aRS(CostField).Value

            If Math.Abs(eqCost) < EPS And Val(gVars.gBasicSetup.Value("WARN_IF_EQ_COST_IS_ZERO")) <> 0 Then
                errorStr = "Truck " & CostField & " Cost is ZERO for Truck Code: [" & TRUCK_CODE & "]"
                Return 0
            Else

                errStr = ""
                fuelPrice = GetFuelPrice(TIME_SLIP_DATE, errStr)
                If Len(errStr) = 0 Then
                    'Recalc 
                    deltaFuel = (fuelPrice - aRS("FUEL_PRICE").Value) * aRS("FUEL_PER_HOUR").Value * (1.0# + aRS("FLUIDS").Value)
                Else
                    deltaFuel = 0
                    'Skip 
                End If

                GetTruckCostHourlyRate = eqCost + deltaFuel

            End If

        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetConversionRate(ByVal BlockCode As String, ByVal aDate As Date, ByVal FromUOM As String, ByVal ToUOM As String, ByVal cPartColl As Collection,
                                      ByRef netWeight As Double, ByVal scaleVolume As Double, ByVal scaleVolumeUOM As String, ByRef errorStr As String) As Double

        Dim strSQL As String
        Dim strSql0 As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim i As Integer
        Dim strWhere As String
        Dim isNoConvert As Object
        Dim OrderByStr As String
        Dim JoinStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetConversionRate = 0
        If (Len(BlockCode) = 0) Then
            GetConversionRate = 0
            Exit Function
        End If

        isNoConvert = 0
        isNoConvert = gComFn.getAValue("IS_NO_CONVERSION", "BLOCKS", "CODE", BlockCode)

        If (IsNothing(isNoConvert)) Then
            'set to NA if Conversions should not be calculated 
            '-> Check to see if an all block will work 
            GetConversionRate = NA
        Else

            If (isNoConvert = ConversionType.NO_CONVERSIONS) Then
                'set to NA if Conversions should not be calculated 
                '-> Check to see if an all block will work 
                GetConversionRate = NA
                'Exit Function 
            End If

        End If

        If (Len(FromUOM) = 0) Then
            GetConversionRate = 0
            Exit Function
        End If

        If (Len(ToUOM) = 0) Then
            GetConversionRate = 0
            Exit Function
        End If

        If Val(gVars.gBasicSetup.Value("ENABLE_SCALE_VOLUME")) <> 0 And Math.Abs(scaleVolume - NA) > EPS Then

            If StrComp(scaleVolumeUOM, ToUOM, vbTextCompare) = 0 And Math.Abs(scaleVolume) > EPS Then
                'GetConversionRate = scaleVolume / netWeight 
                GetConversionRate = VOL_EQUALS_SCALE_VOLUME
                Exit Function
            End If

        End If

        If (StrComp(FromUOM, ToUOM, vbTextCompare) = 0) Then
            GetConversionRate = 1.0#
            Exit Function
        End If


        If (Len(aDate) = 0) Then
            GetConversionRate = 0
            Exit Function
        ElseIf (Not IsDate(aDate)) Then
            GetConversionRate = 0
            Exit Function
        End If

        '--------------------------------------------------------------------------------------------------------- 
        'Try with exact Block Code 
        '--------------------------------------------------------------------------------------------------------- 

        If (isNoConvert <> ConversionType.NO_CONVERSIONS) Then

            strSQL = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |" & BlockCode & "|"

            xRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                Exit Function
            End If

            If (Not xRS.EOF) Then

                strWhere = ""
                OrderByStr = ""
                JoinStr = ""
                strSQL = "SELECT CONVERSION_FACTOR FROM "
                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        strWhere = strWhere & " AND (C_PART_" & i & "_CODE = |" & cPartColl(CStr(i)) & "| OR C_PART_" & i & "_CODE = |NONE| ) "
                        OrderByStr = OrderByStr & " CP" & i & ".IS_NONE, "
                        JoinStr = JoinStr & " INNER JOIN C_PART_" & i & " AS CP" & i & " ON CF.C_PART_" & i & "_CODE = CP" & i & ".CODE  ) "
                        strSQL = strSQL & "("
                    End If
                Next i
                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    strWhere = strWhere & " AND (DESTINATION_CODE = |" & cPartColl(CStr(DESTINATION)) & "| OR DESTINATION_CODE = |NONE| ) "
                    OrderByStr = OrderByStr & " D.IS_NONE, "
                    JoinStr = JoinStr & " INNER JOIN DESTINATIONS AS D ON CF.DESTINATION_CODE = D.CODE ) "
                    strSQL = strSQL & "("
                End If
                OrderByStr = OrderByStr & " EFFECTIVE_DATE DESC"

                strSql0 = strSQL

                strSQL = strSql0 & " CONVERSIONS AS CF " & JoinStr &
                         " WHERE BLOCK_CODE = |" & BlockCode & "| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                         " AND FROM_UOM_CODE = |" & FromUOM & "| AND TO_UOM_CODE = |" & ToUOM & "|" &
                         strWhere & " ORDER BY " & OrderByStr

                aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
                If gObjErrors.Count > 0 Then
                    gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                    Exit Function
                End If

                If (Not aRS.EOF) Then

                    GetConversionRate = aRS("CONVERSION_FACTOR").Value

                Else

                    '--------------------------------------------------------- 
                    'Try the Inverse 
                    'Flip the To and From Values 
                    '--------------------------------------------------------- 
                    strSQL = strSql0 & " CONVERSIONS AS CF " & JoinStr &
                             " WHERE BLOCK_CODE = |" & BlockCode & "| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                             " AND FROM_UOM_CODE = |" & ToUOM & "| AND TO_UOM_CODE = |" & FromUOM & "|" &
                             strWhere & " ORDER BY " & OrderByStr

                    aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
                    If gObjErrors.Count > 0 Then
                        gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                        Exit Function
                    End If

                    If (Not aRS.EOF) Then

                        GetConversionRate = IIf(IsDBNull(aRS("CONVERSION_FACTOR").Value), 0, aRS("CONVERSION_FACTOR").Value)

                        If Math.Abs(GetConversionRate) > 0.000001 Then
                            GetConversionRate = 1.0# / GetConversionRate
                        Else
                            GetConversionRate = NA
                        End If

                    Else

                        GetConversionRate = NA

                    End If


                End If

            Else

                GetConversionRate = NA

            End If

        End If

        '--------------------------------------------------------------------------------------------------------- 
        'Try with |All| Block Code 
        'Do every thing the same but match on the ALL Block 
        '--------------------------------------------------------------------------------------------------------- 

        If (Math.Abs(GetConversionRate - NA) < EPS) Then

            strSQL = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |ALL|"

            xRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                Exit Function
            End If

            If (Not xRS.EOF) Then

                strWhere = ""
                OrderByStr = ""
                JoinStr = ""
                strSQL = "SELECT CONVERSION_FACTOR FROM "
                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        strWhere = strWhere & " AND (C_PART_" & i & "_CODE = |" & cPartColl(CStr(i)) & "| OR C_PART_" & i & "_CODE = |NONE| ) "
                        OrderByStr = OrderByStr & " CP" & i & ".IS_NONE, "
                        JoinStr = JoinStr & " INNER JOIN C_PART_" & i & " AS CP" & i & " ON CF.C_PART_" & i & "_CODE = CP" & i & ".CODE  ) "
                        strSQL = strSQL & "("
                    End If
                Next i
                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    strWhere = strWhere & " AND (DESTINATION_CODE = |" & cPartColl(CStr(DESTINATION)) & "| OR DESTINATION_CODE = |NONE| ) "
                    OrderByStr = OrderByStr & " D.IS_NONE, "
                    JoinStr = JoinStr & " INNER JOIN DESTINATIONS AS D ON CF.DESTINATION_CODE = D.CODE ) "
                    strSQL = strSQL & "("
                End If
                OrderByStr = OrderByStr & " EFFECTIVE_DATE DESC"

                strSql0 = strSQL

                strSQL = strSql0 & " CONVERSIONS AS CF " & JoinStr &
                         " WHERE BLOCK_CODE = |ALL| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                         " AND FROM_UOM_CODE = |" & FromUOM & "| AND TO_UOM_CODE = |" & ToUOM & "|" &
                         strWhere & " ORDER BY " & OrderByStr

                aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
                If gObjErrors.Count > 0 Then
                    gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                    Exit Function
                End If

                If aRS.EOF Then


                    '--------------------------------------------------------- 
                    'Try the Inverse 
                    'Flip the To and From Values 
                    '--------------------------------------------------------- 
                    strSQL = strSql0 & " CONVERSIONS AS CF " & JoinStr &
                             " WHERE BLOCK_CODE = |ALL| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                             " AND FROM_UOM_CODE = |" & ToUOM & "| AND TO_UOM_CODE = |" & FromUOM & "|" &
                             strWhere & " ORDER BY " & OrderByStr

                    aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
                    If gObjErrors.Count > 0 Then
                        gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                        Exit Function
                    End If

                    If (Not aRS.EOF) Then

                        GetConversionRate = IIf(IsDBNull(aRS("CONVERSION_FACTOR").Value), 0, aRS("CONVERSION_FACTOR").Value)

                        If Math.Abs(GetConversionRate) > 0.000001 Then

                            GetConversionRate = 1.0# / GetConversionRate

                        Else

                            'No Factor 
                            GetConversionRate = 0
                            '~errorStr = "Error #1101: No Conversion Set Up For [" & FromUOM & "] to [" & ToUOM & "] For " & gVars.gBlock_Name & " [" & BlockCode & "]" 
                            errorStr = "Error #1101: No Conversion Set Up For [" & FromUOM & "] to [" & ToUOM & "]"
                            Exit Function

                        End If

                    Else

                        'No Factor 
                        GetConversionRate = 0
                        '~errorStr = "Error #1101: No Conversion Set Up For [" & FromUOM & "] to [" & ToUOM & "] For " & gVars.gBlock_Name & " [" & BlockCode & "]" 
                        errorStr = "Error #1101: No Conversion Set Up For [" & FromUOM & "] to [" & ToUOM & "]"
                        Exit Function

                    End If


                Else

                    GetConversionRate = aRS("CONVERSION_FACTOR").Value

                End If

            Else

                'No Factor 
                GetConversionRate = 0
                errorStr = "Error #1102: No Conversion Parts Setup for the ALL " & gVars.gBlock_Name
                Exit Function

            End If

        Else

            'GetConversionRate = aRS("CONVERSION_FACTOR").Value 
            GetConversionRate = GetConversionRate

        End If

        'Check for conversion in wght to volume 
        If (Val(gVars.gBasicSetup.Value("CONVERSION_WGHT_TO_VOL")) = 1) Then
            If (Val(GetConversionRate) > EPS) Then
                GetConversionRate = 1.0# / GetConversionRate
            End If
        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function doIHaveHourRecords(ByRef iEmp As Integer, ByRef employeeCode As String, ByRef SD As Date, ByRef ED As Date) As Integer

        Dim sql As String
        Dim aRS As clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If iEmp <> 0 Then

            sql = " SELECT COUNT(*) FROM HOURLY_PAY_DETAIL WHERE EMPLOYEE_CODE = |" & employeeCode & "| AND " &
                     " TIME_SLIP_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND TIME_SLIP_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI

        Else

            sql = " SELECT COUNT(*) FROM TIME_DETAILS_CONTRACTORS TD INNER JOIN VENDOR_STATEMENT_DETAILS VSD ON TD.CODE = VSD.LOAD_CODE " &
                 " WHERE TD.EMPLOYEE_CODE = |" & employeeCode & "| AND VSD.PAY_TYPE = " & LOAD_PAY_2 &
                 " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI

        End If

        aRS = gDataLayer.LoadRecordset(sql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If Math.Abs(IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)) > EPS Then
            doIHaveHourRecords = 1
        Else
            doIHaveHourRecords = 0
        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function doIHaveHours(ByRef iEmp As Integer, ByRef employeeCode As String, ByRef SD As Date, ByRef ED As Date) As Integer

        Dim sql As String
        Dim aRS As clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If iEmp <> 0 Then

            sql = " SELECT SUM(HOURS) FROM HOURLY_PAY_DETAIL WHERE EMPLOYEE_CODE = |" & employeeCode & "| AND " &
                     " TIME_SLIP_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND TIME_SLIP_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI

        Else

            sql = " SELECT SUM(VSD.PAY_WEIGHT) FROM TIME_DETAILS_CONTRACTORS TD INNER JOIN VENDOR_STATEMENT_DETAILS VSD ON TD.CODE = VSD.LOAD_CODE " &
                 " WHERE TD.EMPLOYEE_CODE = |" & employeeCode & "| AND VSD.PAY_TYPE = " & LOAD_PAY_2 &
                 " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI

        End If

        aRS = gDataLayer.LoadRecordset(sql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If Math.Abs(IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)) > EPS Then
            doIHaveHours = 1
        Else
            doIHaveHours = 0
        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function doIHaveTravelHours(ByRef iEmp As Integer, ByRef employeeCode As String, ByRef Trvl As String, ByRef SD As Date, ByRef ED As Date) As Integer

        Dim sql As String
        Dim aRS As clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If iEmp <> 0 Then

            sql = " SELECT SUM(HOURS) FROM HOURLY_PAY_DETAIL WHERE EMPLOYEE_CODE = |" & employeeCode & "| AND " &
                     " TIME_SLIP_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND TIME_SLIP_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI &
                     " AND ACTIVITY_CODE = |" & Trvl & "|"

        Else

            sql = " SELECT SUM(VSD.PAY_WEIGHT) FROM TIME_DETAILS_CONTRACTORS TD INNER JOIN VENDOR_STATEMENT_DETAILS VSD ON TD.CODE = VSD.LOAD_CODE " &
                  " WHERE TD.EMPLOYEE_CODE = |" & employeeCode & "| AND VSD.PAY_TYPE = " & LOAD_PAY_2 &
                  " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI &
                  " AND ACTIVITY_CODE = |" & Trvl & "|"

        End If

        aRS = gDataLayer.LoadRecordset(sql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If Math.Abs(IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)) > EPS Then
            doIHaveTravelHours = 1
        Else
            doIHaveTravelHours = 0
        End If

        aRS.Close1()
        aRS = Nothing


        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function doIHaveSubsistence(ByVal iEmp As Integer, ByVal SubSist As String, ByVal employeeCode As String, ByRef SD As Date, ByRef ED As Date) As Integer

        Dim c As Integer
        Dim T As String
        Dim sql As String
        Dim xRS As clsRecordsetExt
        Dim aRS As clsRecordsetExt
        Dim total As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If iEmp <> 0 Then

            sql = " SELECT SUM(AMOUNT) FROM EMPLOYEE_EXPENSES WHERE EMPLOYEE_CODE = |" & employeeCode & "| AND " &
                     " ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND ENTRY_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI &
                     " AND ACCOUNT_CODE = |" & SubSist & "| "
        Else

            sql = " SELECT SUM(AMOUNT) FROM MISCELLANEOUS_EXPENSES WHERE OWNER_CODE = |" & employeeCode & "| AND " &
                     " ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND ENTRY_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI &
                     " AND ACCOUNT_CODE = |" & SubSist & "| "

        End If

        aRS = gDataLayer.LoadRecordset(sql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If Math.Abs(IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)) > EPS Then
            doIHaveSubsistence = 1
        Else
            doIHaveSubsistence = 0
        End If

        aRS.Close1()
        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function is_Guy_an_Employee(ByRef empCode As Object) As Integer

        '0=Not An employee 
        '1=Is Employee 

        Dim strSQL As String
        Dim aRS As clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        is_Guy_an_Employee = 0

        'Check to see if employee is an Employee 
        strSQL = "SELECT OWNER FROM OWNERS OW INNER JOIN EMPLOYEES EM ON EM.OWNER_CODE = OW.CODE WHERE EM.CODE = |" & empCode & "|"
        aRS = New clsRecordsetExt

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If gObjErrors.Count = 0 Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        '0=Not An employee 
        '1=Is Employee 
        If (aRS.EOF) Then
            is_Guy_an_Employee = 0
        Else
            is_Guy_an_Employee = aRS(0).Value
        End If

        Exit Function

        gBF.TurnMeOff()

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function isGuySelfEmployed(ByVal guy As String) As Integer

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim myOwner As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        myOwner = gComFn.getAValue("OWNER_CODE", "EMPLOYEES", "CODE", guy)

        strSQL = "SELECT COUNT(*) FROM EMPLOYEES WHERE OWNER_CODE = |" & myOwner & "|"

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count > 0) Then
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        Else
        End If

        If (aRS(0).Value = 1) Then
            isGuySelfEmployed = 1
        Else
            isGuySelfEmployed = 0
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployeeNameFromCode(ByRef VendorCode As String) As String

        Dim aRS0 As clsRecordsetExt, SQLStmnt0 As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        SQLStmnt0 = "SELECT * FROM EMPLOYEES WHERE CODE = |" & VendorCode & "|"

        aRS0 = gDataLayer.LoadRecordset(SQLStmnt0, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return ""
            Exit Function
        End If

        GetEmployeeNameFromCode = aRS0("FULL_NAME").Value

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function CalcMillTime(ByRef Date_Out As Date, ByRef Date_In As Date) As String

        Dim d As Double, Minutes As Integer, HOURS As Integer, lngSeconds As Integer
        Dim SS As String, HH As String, mm As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        lngSeconds = DateDiff("s", Date_In, Date_Out)
        Minutes = lngSeconds \ 60
        lngSeconds = lngSeconds Mod 60
        HOURS = Minutes \ 60
        Minutes = Minutes Mod 60
        If HOURS > 100 Then HOURS = 99

        HH = CStr(HOURS)
        mm = CStr(Minutes)
        SS = CStr(lngSeconds)
        If (Len(HH) = 1) Then HH = "0" & HH
        If (Len(mm) = 1) Then mm = "0" & mm
        If (Len(SS) = 1) Then SS = "0" & SS

        CalcMillTime = HH & ":" & mm & ":" & SS

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetAssociatedRevenueContractValues(ByRef payWeight As Double, ByVal payMeasureCode As String, ByVal payRate As Double, ByVal revRate As Double, ByVal LoadCode As String) As Integer

        '------------------------------------------------------------------------------------ 
        '   Reports 
        '   607 - rptTrckCntrPay 
        '   607 Roga - rptTrckCntrPayRoga 
        '   613 - rptTrckCntrXport 
        '   617 - rptTrckDriverPayRoga 
        '------------------------------------------------------------------------------------ 

        Dim strSQLx As String
        Dim xRS As clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        'Find associated Load revenue record 

        'Gat any load revenue for trucking 
        strSQLx = "SELECT RATE_TYPE_CODE, PAY_WEIGHT, RATE FROM LOADSLIP_REVENUE WHERE LOAD_CODE = |" & LoadCode & "|" &
                 " AND ACTIVITY_CODE = |" & gVars.gTruckingCode & "|"

        xRS = gDataLayer.LoadRecordset(strSQLx, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If Not xRS.EOF Then

            'Three replace ment columns from load revenue 
            '1.  Pay Weight 
            '2.  Pay Basis 
            '3.  Rate = pay rate * revenue rate 
            payWeight = xRS("PAY_WEIGHT").Value
            payMeasureCode = xRS("RATE_TYPE_CODE").Value
            payRate = xRS("RATE").Value * revRate
            GetAssociatedRevenueContractValues = 1

        Else

            'Return what comes is 
            'payWeight = cRs("PAY_WEIGHT") 
            'payMeasureCode = cRs("PAY_MEASURE_CODE") 
            'payRate = cRs("PAY_RATE") 
            GetAssociatedRevenueContractValues = 0

        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    '    Public Function GetOverload_Grace(ByVal TicketNo As String, ByVal Gross As Double, ByVal Net As Double, ByVal TruckCode As String, ByVal RouteCode As String, ByVal TruckTypeCode As String, ByVal DestinationCode As String, ByVal DateIn As Date, ByVal IsNoOverload As Integer, _
    '                                      ByRef Overload As Double, ByRef mxWeight As Double, ByRef partsColl As Collection, ByRef cPartColl As Collection, ByRef errorStr As String) As Integer

    '        '------------------------------------------------------------------------------------ 
    '        '   Reports 
    '        '   638 - rpt TrckOverLoads for Nogel Trucking 
    '        '   Uses Destination Grace Value to Deterime Overload Amount 
    '        '------------------------------------------------------------------------------------ 

    '        Dim aRS As clsRecordsetExt
    '        Dim strSQL As String
    '        Dim isOver As Integer
    '        Dim AltTruckMaxWeight As Double

    '        If (gDebugMode = 0) Then On Error GoTo Err_Handler

    '        Overload = 0

    '        strSQL = "SELECT OVERLOAD_DELTA, IS_DEDUCT_OVERLOADS, IS_USE_OVERLOAD_RATE, IS_USE_OVERLOAD_SCHEDULE, IS_OVERLOAD_BASED_ON_NET FROM DESTINATIONS WHERE CODE = |" & DestinationCode & "|"

    '        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
    '        If (gobjErrors.Count = 0) Then
    '        Else
    '            gBF.DisplayErrors(BasLocals.gobjErrors, BasLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
    '            Return 0
    '            Exit Function
    '        End If

    '        If ((aRS("IS_DEDUCT_OVERLOADS").Value <> 0 Or aRS("IS_USE_OVERLOAD_RATE").Value <> 0 Or aRS("IS_USE_OVERLOAD_SCHEDULE").Value <> 0) And IsNoOverload = 0) Then

    '            mxWeight = GetTruckMaxWeight(TruckCode, RouteCode, TruckTypeCode, DestinationCode, DateIn, aRS("IS_OVERLOAD_BASED_ON_NET").Value, partsColl, cPartColl, I_PAY, 0, AltTruckMaxWeight)

    '            mxWeight = mxWeight + aRS("OVERLOAD_DELTA").Value

    '            If (mxWeight > EPS) Then

    '                If (aRS("IS_OVERLOAD_BASED_ON_NET").Value) Then
    '                    Overload = Max(Net - mxWeight, 0.0#)
    '                Else
    '                    Overload = Max(Gross - mxWeight, 0.0#)
    '                End If

    '                'Convert in to Pay UOM 
    '                If (Math.Abs(Overload) > EPS) Then

    '                    isOver = 1

    '                End If

    '            Else

    '                isOver = 0
    '                mxWeight = 0
    '                Overload = 0
    '                errorStr = "Warning #1103: No Maximum Weight Set for Truck: [" & TruckCode & "], NO Overload applied. Ticket: [" & TicketNo & "]"

    '            End If

    '        Else

    '            isOver = 0
    '            mxWeight = 0
    '            Overload = 0

    '        End If

    '        aRS = Nothing

    '        GetOverload_Grace = isOver

    '        Exit Function

    'Err_Handler:
    '        DisplayErrors(Err.Number, Err.Description, Err.Source & ":frmDataSheet.LoadMe")

    '    End Function


    '    Public Function GetOverload(ByVal TicketNo As String, ByVal Gross As Double, ByVal Net As Double, ByVal TruckCode As String, ByVal RouteCode As String, ByVal TruckTypeCode As String, ByVal DestinationCode As String, ByVal DateIn As Date, ByVal IsNoOverload As Integer, _
    '                                ByRef Overload As Double, ByRef mxWeight As Double, ByRef partsColl As Collection, ByRef cPartColl As Collection, ByVal isPayOrRev As Integer, ByVal isUsePayWeight As Integer, ByRef errorStr As String) As Integer

    '        '------------------------------------------------------------------------------------ 
    '        '   Reports 
    '        '   604 - rptTrckOverloads 
    '        '   607 - rptTrckCntrPay 
    '        '   607 Roga - rptTrckCntrPayRoga 
    '        '   613 - rptTrckCntrXport 
    '        '   617 - rptTrckDriverPayRoga 
    '        '------------------------------------------------------------------------------------ 

    '        Dim aRS As clsRecordsetExt
    '        Dim strSQL As String
    '        Dim isOver As Integer
    '        Dim AltTruckMaxWeight As Double

    '        If (gDebugMode = 0) Then On Error GoTo Err_Handler

    '        Overload = 0

    '        strSQL = "SELECT IS_DEDUCT_OVERLOADS,IS_USE_OVERLOAD_RATE, IS_USE_OVERLOAD_SCHEDULE,IS_OVERLOAD_BASED_ON_NET FROM DESTINATIONS WHERE CODE = |" & DestinationCode & "|"

    '        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
    '        If (gobjErrors.Count = 0) Then
    '        Else
    '            gBF.DisplayErrors(BasLocals.gobjErrors, BasLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
    '            Return 0
    '            Exit Function
    '        End If

    '        If ((aRS("IS_DEDUCT_OVERLOADS").Value <> 0 Or aRS("IS_USE_OVERLOAD_RATE").Value <> 0 Or aRS("IS_USE_OVERLOAD_SCHEDULE").Value <> 0) And IsNoOverload = 0) Then

    '            mxWeight = GetTruckMaxWeight(TruckCode, RouteCode, TruckTypeCode, DestinationCode, DateIn, aRS("IS_OVERLOAD_BASED_ON_NET").Value, partsColl, cPartColl, isPayOrRev, isUsePayWeight, AltTruckMaxWeight)

    '            If (mxWeight > EPS) Then

    '                If (aRS("IS_OVERLOAD_BASED_ON_NET").Value) Then
    '                    Overload = Max(Net - mxWeight, 0.0#)
    '                Else
    '                    Overload = Max(Gross - mxWeight, 0.0#)
    '                End If

    '                'Convert in to Pay UOM 
    '                If (Math.Abs(Overload) > EPS) Then

    '                    isOver = 1

    '                End If

    '            Else

    '                isOver = 0
    '                mxWeight = 0
    '                Overload = 0
    '                errorStr = "Warning #1103: No Maximum Weight Set for Truck: [" & TruckCode & "], NO Overload applied. Ticket: [" & TicketNo & "]"

    '            End If

    '        Else

    '            isOver = 0
    '            mxWeight = 0
    '            Overload = 0

    '        End If

    '        aRS = Nothing

    '        GetOverload = isOver

    '        Exit Function

    'Err_Handler:
    '        DisplayErrors(Err.Number, Err.Description, Err.Source & ":frmDataSheet.LoadMe")

    '    End Function


    '    Public Function GetTruckMaxWeight(ByVal TRUCK_CODE As String, ByVal ROUTE_CODE As String, ByVal TRUCK_TYPE_CODE As String, ByVal DESTINATION_CODE As String, ByVal Date_In As Date, _
    '                                      ByVal IS_OVERLOAD_BASED_ON_NET As Integer, ByRef partsColl As Collection, ByRef cPartColl As Collection, ByVal isPayOrRev As Integer, ByVal isUsePayWeight As Integer, _
    '                                      ByRef AltTruckMaxWeight As Double) As Double


    '        '------------------------------------------------------------------------------------ 
    '        '   LOGGERS EDGE 
    '        '   Reports 
    '        '   616 - rptTruckSlackerTrcker 
    '        '------------------------------------------------------------------------------------ 

    '        Dim strSQL As String
    '        Dim aRS As clsRecordsetExt
    '        Dim xRS As clsRecordsetExt
    '        Dim i As Integer
    '        Dim TruckMaxWeight As Double

    '        If (gDebugMode = 0) Then On Error GoTo Err_Handler

    '        GetTruckMaxWeight = 0

    '        'strSQL = "SELECT TOP 1 T1.SUMMER_MAX, T1.WINTER_MAX, T1.SUMMER_MAX_NET, T1.WINTER_MAX_NET FROM (((( TRUCK_MAX_WEIGHTS T1 " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE ) " & _ 
    '        '    " INNER JOIN DESTINATIONS D ON D.CODE = T1.DESTINATION_CODE ) " & _ 
    '        '    " INNER JOIN ROUTES R ON R.CODE = T1.ROUTE_CODE ) " & _ 
    '        '    " INNER JOIN TRUCK_TYPES TT ON TT.CODE = T1.TRUCK_TYPE_CODE ) " & _ 
    '        '    " WHERE (T1.TRUCK_CODE = |" & TRUCK_CODE & "| OR T1.TRUCK_CODE = |NONE|) " & _ 
    '        '    " AND (T1.ROUTE_CODE = |" & ROUTE_CODE & "| OR T1.ROUTE_CODE = |NONE|) " & _ 
    '        '    " AND (T1.TRUCK_TYPE_CODE = |" & TRUCK_TYPE_CODE & "| OR T1.TRUCK_TYPE_CODE = |NONE|)" & _ 
    '        '    " AND (T1.DESTINATION_CODE = |" & DESTINATION_CODE & "| OR T1.DESTINATION_CODE = |NONE|)" & _ 
    '        '    " AND T1.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(Date_In) & gVars.gDI & " ORDER BY TR.IS_NONE ASC, D.IS_NONE, R.IS_NONE ASC, TT.IS_NONE ASC, T1.EFFECTIVE_DATE DESC " 


    '        Dim JoinStr As String
    '        Dim whereStr As String
    '        Dim OrderByStr As String

    '        JoinStr = " (((( TRUCK_MAX_WEIGHTS T1 INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE ) " & _
    '           " INNER JOIN DESTINATIONS D ON D.CODE = T1.DESTINATION_CODE ) " & _
    '           " INNER JOIN ROUTES R ON R.CODE = T1.ROUTE_CODE ) " & _
    '           " INNER JOIN TRUCK_TYPES TT ON TT.CODE = T1.TRUCK_TYPE_CODE ) "

    '        whereStr = " WHERE (T1.TRUCK_CODE = |" & TRUCK_CODE & "| OR T1.TRUCK_CODE = |NONE|) " & _
    '            " AND (T1.ROUTE_CODE = |" & ROUTE_CODE & "| OR T1.ROUTE_CODE = |NONE|) " & _
    '            " AND (T1.TRUCK_TYPE_CODE = |" & TRUCK_TYPE_CODE & "| OR T1.TRUCK_TYPE_CODE = |NONE|)" & _
    '            " AND (T1.DESTINATION_CODE = |" & DESTINATION_CODE & "| OR T1.DESTINATION_CODE = |NONE|) "

    '        OrderByStr = " ORDER BY TR.IS_NONE ASC, D.IS_NONE, R.IS_NONE ASC, TT.IS_NONE ASC, "

    '        For i = 1 To partsColl.Count
    '            If (partsColl(i).partNo > MAX_PRIME_PARTS And partsColl(i).partNo < 100) Then
    '                If (gobjTables("TRUCK_MAX_WEIGHTS").Fields("C_PART_" & partsColl(i).partNo & "_CODE").DisplayOrder > 0) Then
    '                    JoinStr = "(" & JoinStr & " INNER JOIN C_PART_" & partsColl(i).partNo & " AS T" & partsColl(i).partNo & " ON T" & partsColl(i).partNo & ".CODE = T1.C_PART_" & partsColl(i).partNo & "_CODE ) "
    '                    whereStr = whereStr & " AND (T1.C_PART_" & partsColl(i).partNo & "_CODE = |" & cPartColl(CStr(partsColl(i).partNo)) & "| OR T1.C_PART_" & partsColl(i).partNo & "_CODE = |NONE| )"
    '                    OrderByStr = OrderByStr & " T" & partsColl(i).partNo & ".IS_NONE ASC, "
    '                End If
    '            End If
    '        Next i

    '        strSQL = "SELECT TOP 1 T1.SUMMER_MAX, T1.WINTER_MAX, T1.SUMMER_MAX_NET, T1.WINTER_MAX_NET, T1.MAX_WEIGHT_FOR_PAY, T1.MAX_WEIGHT_FOR_REV, T1.MAX_WEIGHT_FOR_PAY_SUMMER, T1.MAX_WEIGHT_FOR_REV_SUMMER " & _
    '            " FROM  " & JoinStr & whereStr & _
    '            " AND T1.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(Date_In) & gVars.gDI & OrderByStr & " T1.EFFECTIVE_DATE DESC "

    '        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
    '        If (gobjErrors.Count = 0) Then
    '        Else
    '            gBF.DisplayErrors(BasLocals.gobjErrors, BasLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
    '            Exit Function
    '        End If

    '        If (aRS.EOF) Then
    '            GetTruckMaxWeight = 0
    '            Exit Function
    '        End If

    '        'Determine if Summer or Winter 
    '        strSQL = " SELECT TOP 1 IS_WINTER FROM SEASONS AS T0 INNER JOIN ROUTES AS R ON T0.ROUTE_CODE = R.CODE WHERE " & gVars.gDI & gBF.IntlDateFormat(Date_In) & gVars.gDI & " >= T0.THE_DATE " & " AND (T0.ROUTE_CODE = |NONE| OR T0.ROUTE_CODE = |" & ROUTE_CODE & "|)" & " ORDER BY R.IS_NONE ASC, T0.THE_DATE DESC "

    '        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)

    '        If (xRS.EOF) Then
    '            GetTruckMaxWeight = 0
    '            Exit Function
    '        Else

    '            If (xRS(0).Value <> 0) Then
    '                If (IS_OVERLOAD_BASED_ON_NET <> 0) Then
    '                    TruckMaxWeight = aRS("WINTER_MAX_NET").Value
    '                Else
    '                    TruckMaxWeight = aRS("WINTER_MAX").Value
    '                End If
    '            Else
    '                If (IS_OVERLOAD_BASED_ON_NET <> 0) Then
    '                    TruckMaxWeight = aRS("SUMMER_MAX_NET").Value
    '                Else
    '                    TruckMaxWeight = aRS("SUMMER_MAX").Value
    '                End If
    '            End If

    '            If isUsePayWeight <> 0 Then

    '                If isPayOrRev = I_PAY Then

    '                    If (xRS(0).Value <> 0) Then
    '                        AltTruckMaxWeight = aRS("MAX_WEIGHT_FOR_PAY").Value
    '                    Else
    '                        AltTruckMaxWeight = aRS("MAX_WEIGHT_FOR_PAY_SUMMER").Value
    '                    End If

    '                ElseIf isPayOrRev = I_REVENUE Then

    '                    If (xRS(0).Value <> 0) Then
    '                        AltTruckMaxWeight = aRS("MAX_WEIGHT_FOR_REV").Value
    '                    Else
    '                        AltTruckMaxWeight = aRS("MAX_WEIGHT_FOR_REV_SUMMER").Value
    '                    End If

    '                End If

    '            Else

    '                AltTruckMaxWeight = TruckMaxWeight

    '            End If

    '            GetTruckMaxWeight = TruckMaxWeight

    '        End If

    '        aRS.Close1()
    '        aRS = Nothing

    '        xRS.Close1()
    '        xRS = Nothing

    '        Exit Function

    'Err_Handler:
    '        DisplayErrors(Err.Number, Err.Description, Err.Source & ":basReportDllFunctions.GetTruckMaxWeight")

    '    End Function

    '    Public Function GetTruckMaxWeight_Obsolete(ByVal TRUCK_CODE As String, ByVal ROUTE_CODE As String, ByVal TRUCK_TYPE_CODE As String, _
    '                                               ByVal DESTINATION_CODE As String, ByVal Date_In As Date, ByVal IS_OVERLOAD_BASED_ON_NET As Integer) As Double


    '        '------------------------------------------------------------------------------------ 
    '        '   LOGGERS EDGE 
    '        '   Reports 
    '        '   616 - rptTruckSlackerTrcker 
    '        '------------------------------------------------------------------------------------ 

    '        Dim strSQL As String
    '        Dim aRS As clsRecordsetExt
    '        Dim xRS As clsRecordsetExt

    '        If (gDebugMode = 0) Then On Error GoTo Err_Handler

    '        GetTruckMaxWeight_Obsolete = 0

    '        strSQL = "SELECT TOP 1 T1.SUMMER_MAX, T1.WINTER_MAX, T1.SUMMER_MAX_NET, T1.WINTER_MAX_NET FROM (((( TRUCK_MAX_WEIGHTS T1 " & _
    '                  " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE ) " & _
    '                  " INNER JOIN DESTINATIONS D ON D.CODE = T1.DESTINATION_CODE ) " & _
    '                  " INNER JOIN ROUTES R ON R.CODE = T1.ROUTE_CODE ) " & _
    '                  " INNER JOIN TRUCK_TYPES TT ON TT.CODE = T1.TRUCK_TYPE_CODE ) " & _
    '                  " WHERE (T1.TRUCK_CODE = |" & TRUCK_CODE & "| OR T1.TRUCK_CODE = |NONE|) " & _
    '                  " AND (T1.ROUTE_CODE = |" & ROUTE_CODE & "| OR T1.ROUTE_CODE = |NONE|) " & _
    '                  " AND (T1.TRUCK_TYPE_CODE = |" & TRUCK_TYPE_CODE & "| OR T1.TRUCK_TYPE_CODE = |NONE|)" & _
    '                  " AND (T1.DESTINATION_CODE = |" & DESTINATION_CODE & "| OR T1.DESTINATION_CODE = |NONE|)" & _
    '                  " AND T1.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(Date_In) & gVars.gDI & _
    '                  " ORDER BY TR.IS_NONE ASC, D.IS_NONE, R.IS_NONE ASC, TT.IS_NONE ASC, T1.EFFECTIVE_DATE DESC "

    '        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
    '        If (gobjErrors.Count = 0) Then
    '        Else
    '            gBF.DisplayErrors(BasLocals.gobjErrors, BasLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
    '            Exit Function
    '        End If

    '        If (aRS.EOF) Then
    '            GetTruckMaxWeight_Obsolete = 0
    '            Exit Function
    '        End If

    '        'Determine if Summer or Winter 
    '        strSQL = " SELECT TOP 1 IS_WINTER FROM SEASONS AS T0 INNER JOIN ROUTES AS R ON T0.ROUTE_CODE = R.CODE WHERE " & gVars.gDI & gBF.IntlDateFormat(Date_In) & gVars.gDI & " >= T0.THE_DATE " & _
    '                 " AND (T0.ROUTE_CODE = |NONE| OR T0.ROUTE_CODE = |" & ROUTE_CODE & "|)" & _
    '                 " ORDER BY R.IS_NONE ASC, T0.THE_DATE DESC "

    '        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)

    '        If (xRS.EOF) Then
    '            GetTruckMaxWeight_Obsolete = 0
    '            Exit Function
    '        Else
    '            If (xRS(0).Value <> 0) Then
    '                If (IS_OVERLOAD_BASED_ON_NET <> 0) Then
    '                    GetTruckMaxWeight_Obsolete = aRS("WINTER_MAX_NET").Value
    '                Else
    '                    GetTruckMaxWeight_Obsolete = aRS("WINTER_MAX").Value
    '                End If
    '            Else
    '                If (IS_OVERLOAD_BASED_ON_NET <> 0) Then
    '                    GetTruckMaxWeight_Obsolete = aRS("SUMMER_MAX_NET").Value
    '                Else
    '                    GetTruckMaxWeight_Obsolete = aRS("SUMMER_MAX").Value
    '                End If
    '            End If
    '        End If

    '        aRS = Nothing
    '        xRS = Nothing

    '        Exit Function

    'Err_Handler:
    '        DisplayErrors(Err.Number, Err.Description, Err.Source & ":basReportDllFunctions.GetTruckMaxWeight_Obsolete")

    '    End Function



    Public Function GetTheBenefitRate(ByVal aLoad As clsRecordsetExt, ByVal activityCode As String, ByRef aBasis As Double, ByRef aRate As Double, ByRef aUoM As String, ByRef partsColl As Collection) As Integer

        '--------------------------------------------------------------- 
        'Returns NA if NO benefit rate found 
        '--------------------------------------------------------------- 

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim rtRS As clsRecordsetExt
        Dim LoadVolume As Double
        Dim aCode As String
        Dim cPartColl As Collection
        Dim Conversion As Double
        Dim IS_NO_CONVERSION As Integer
        Dim i As Integer
        Dim errorStr As String = ""
        Dim JoinStr As String
        Dim whereStr As String
        Dim OrderByStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        JoinStr = " ((( BENEFIT_RATES AS BR " &
                 " INNER JOIN BLOCKS AS BL ON BR.BLOCK_CODE = BL.CODE) " &
                 " INNER JOIN DESTINATIONS AS DS ON BR.DESTINATION_CODE = DS.CODE) " &
                 " INNER JOIN ROUTES AS RT ON BR.ROUTE_CODE = RT.CODE) "

        whereStr = " WHERE (BR.BLOCK_CODE = |NONE| OR BR.BLOCK_CODE = |" & aLoad("BLOCK_CODE").Value & "|)" &
                 " AND (BR.DESTINATION_CODE = |NONE| OR BR.DESTINATION_CODE = |" & aLoad("DESTINATION_CODE").Value & "| )" &
                 " AND (BR.ROUTE_CODE = |NONE| OR BR.ROUTE_CODE = |" & aLoad("ROUTE_CODE").Value & "| )" &
                 " AND (BR.ACTIVITY_CODE = |" & activityCode & "| )" &
                 " AND BR.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aLoad("DATE_OUT").Value) & gVars.gDI

        OrderByStr = " ORDER BY BL.IS_NONE ASC, DS.IS_NONE ASC, "

        For i = 1 To MAX_CPARTS

            If (gBF.isObjKeyinVBCollection(partsColl, "C_PART_" & i & "_CODE")) Then
                JoinStr = "(" & JoinStr & " INNER JOIN C_PART_" & i & " AS T" & i & " ON T" & i & ".CODE = BR.C_PART_" & i & "_CODE ) "
                whereStr = whereStr & " AND (BR.C_PART_" & i & "_CODE = |" & aLoad("C_PART_" & i & "_CODE").Value & "| OR BR.C_PART_" & i & "_CODE = |NONE| )"
                OrderByStr = OrderByStr & " T" & i & ".IS_NONE ASC, "
            End If

        Next i

        OrderByStr = OrderByStr & " RT.IS_NONE ASC, BR.EFFECTIVE_DATE DESC "

        strSQL = "SELECT TOP 1 BENEFIT_RATE AS RATE, RATE_TYPE_CODE FROM " &
                 JoinStr & whereStr & OrderByStr

        rtRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (rtRS.EOF) Then
            GetTheBenefitRate = NA
            Exit Function
        Else
            aUoM = rtRS("RATE_TYPE_CODE").Value
            aRate = rtRS("RATE").Value
        End If

        LoadVolume = NA

        If (StrComp(aLoad("SCALE_VOLUME_UOM_CODE").Value, rtRS("RATE_TYPE_CODE").Value) = 0) Then

            LoadVolume = aLoad("SCALE_VOLUME").Value

        ElseIf (StrComp(aUoM, "LOAD", vbTextCompare) = 0) Then

            LoadVolume = aLoad("LOADS").Value

        Else

            aRS = New clsRecordsetExt

            strSQL = "SELECT LOAD_UOM_CODE, VOLUME_UOM_CODE, IS_NO_CONVERSION FROM BLOCKS WHERE CODE = |" & aLoad("BLOCK_CODE").Value & "|"

            aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                Return 0
                Exit Function
            End If

            IS_NO_CONVERSION = aRS("IS_NO_CONVERSION").Value

            'If NO conversion -> Must CONVERT if flag is set to do conversions 
            'And Volume on Load is the Same as the Pay Basis 
            If (IS_NO_CONVERSION = ConversionType.NO_CONVERSIONS And StrComp(rtRS("RATE_TYPE_CODE").Value, aRS("VOLUME_UOM_CODE").Value, vbTextCompare) = 0) Then

                'If block convert is different from pay convert for volume -> must convert below 
                LoadVolume = aLoad("VOLUME").Value

            ElseIf (StrComp(rtRS("RATE_TYPE_CODE").Value, aRS("LOAD_UOM_CODE").Value, vbTextCompare) = 0) Then

                LoadVolume = aLoad("NET").Value

            Else

                cPartColl = New Collection

                If (Not IsDBNull(aLoad("DESTINATION_CODE").Value)) Then
                    aCode = aLoad("DESTINATION_CODE").Value
                    If (Len(aCode) > 0) Then
                        '-> Got it 
                    Else
                        aCode = "NONE"
                    End If
                Else
                    aCode = "NONE"
                End If
                cPartColl.Add(aCode, CStr(DESTINATION))

                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (Not IsDBNull(aLoad("C_PART_" & i & "_CODE"))) Then
                        aCode = aLoad("C_PART_" & i & "_CODE").Value
                        If (Len(aCode) > 0) Then
                            '-> Got it 
                        Else
                            aCode = "NONE"
                        End If
                    Else
                        aCode = "NONE"
                    End If
                    cPartColl.Add(aCode, CStr(i))
                Next i

                'Dim isConvertBasedonNET As Double 
                'isConvertBasedonNET = gComFn.getAValue("IS_CONVERT_USING_NET", "BLOCKS", "CODE", aBlock) 

                'Try to Convert weight into payBasis 
                Conversion = GetConversionRate(aLoad("BLOCK_CODE").Value, aLoad("DATE_OUT").Value, aRS("LOAD_UOM_CODE").Value, rtRS("RATE_TYPE_CODE").Value, cPartColl, aLoad("NET").Value, aLoad("SCALE_VOLUME").Value, aLoad("SCALE_VOLUME_UOM_CODE").Value, errorStr)

                If Math.Abs(Conversion - VOL_EQUALS_SCALE_VOLUME) < EPS Then

                    LoadVolume = aLoad("SCALE_VOLUME").Value

                ElseIf (Len(errorStr) = 0 And Math.Abs(Conversion - NA) > EPS) Then

                    LoadVolume = Conversion * Val(aLoad("NET").Value)

                Else
                    GetTheBenefitRate = False
                    Exit Function
                End If

            End If

        End If

        aBasis = LoadVolume
        GetTheBenefitRate = True

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function

    Public Function GetBaseChargeOutRate(ByVal EMPLOYEE_CODE As String, ByVal TIME_SLIP_DATE As Date, ByVal RATE_TYPE_CODE As String, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        strSQL = "SELECT TOP 1 CHARGE_OUT_RATE FROM EMPLOYEE_BASE_RATES WHERE RATE_TYPE_CODE = |" & RATE_TYPE_CODE &
                 "| AND EMPLOYEE_CODE = |" & EMPLOYEE_CODE & "|" &
                 " AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "

        'Bob 1/24/05 removed the is_active flag 
        ' " AND IS_ACTIVE <> 0 " & _ 

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then
            errorStr = "Error #1402: Base Charge Out Rate not found for Employee: [" & EMPLOYEE_CODE & "]; Time Slip Date: [" & TIME_SLIP_DATE & "]"
            Return 0
        Else
            'Base Pay Rate 
            GetBaseChargeOutRate = aRS("CHARGE_OUT_RATE").Value
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetChargeOutRateFactor(ByRef isPercent As Integer, ByVal EMPLOYEE_CODE As String, ByVal TIME_SLIP_DATE As Date, ByVal PAY_ACTIVITY_CODE As String,
                                           ByVal REVENUE_CONTRACT_CODE As String, ByVal RATE_TYPE_CODE As String, ByRef errorStr As String) As Double

        Dim aRS As clsRecordsetExt
        Dim strSQL As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New clsRecordsetExt
        strSQL = "SELECT TOP 1 CHARGE_RATE, IS_PERCENT FROM (( EMPLOYEE_CHARGE_RATES ER " &
                 " INNER JOIN EMPLOYEES EE ON EE.CODE = ER.EMPLOYEE_CODE) " &
                 " INNER JOIN REVENUE_CONTRACTS AS RV ON RV.CODE = ER.REVENUE_CONTRACT_CODE ) " &
                 " WHERE CHARGE_RATE_TYPE_CODE = |" & RATE_TYPE_CODE & "|" &
                 " AND (ER.EMPLOYEE_CODE = |" & EMPLOYEE_CODE & "| OR ER.EMPLOYEE_CODE = |ALL|) " &
                 " AND ER.PHASE_CODE = |" & PAY_ACTIVITY_CODE & "| " &
                 " AND (ER.REVENUE_CONTRACT_CODE = |NONE| OR ER.REVENUE_CONTRACT_CODE = |" & REVENUE_CONTRACT_CODE & "|) " &
                 " AND ER.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI &
                 " ORDER BY EE.IS_ALL ASC, RV.IS_NONE ASC, ER.EFFECTIVE_DATE DESC "

        ' Bob 4/15/05 Made the match based on REVENUE_CONTRACT_CODE = NONE or exact match 
        'AND ER.REVENUE_CONTRACT_CODE = |" & REVENUE_CONTRACT_CODE & "| AND ER.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & _ 

        aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If aRS.EOF Then

            'Try the ALL Employee 
            aRS = New clsRecordsetExt
            strSQL = "SELECT TOP 1 CHARGE_RATE, IS_PERCENT FROM EMPLOYEE_CHARGE_RATES WHERE CHARGE_RATE_TYPE_CODE = |" & RATE_TYPE_CODE &
                     "| AND EMPLOYEE_CODE = |ALL|" &
                     " AND PHASE_CODE = |" & PAY_ACTIVITY_CODE & "| AND REVENUE_CONTRACT_CODE = |" & REVENUE_CONTRACT_CODE & "| AND EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(TIME_SLIP_DATE) & gVars.gDI & " ORDER BY EFFECTIVE_DATE DESC "

            aRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                Return 0
                Exit Function
            End If

            If aRS.EOF Then

                If Val("" & gVars.gBasicSetup.Value("IS_ONLY_USE_EMPLOYEE_CHARGE_RATES")) <> 0 Then
                    '~errorStr = "Charge Rate in EMPLOYEE_CHARGE_RATES is not found for Employee: " & EMPLOYEE_CODE & ", Activity: " & PAY_ACTIVITY_CODE & ", Date: " & TIME_SLIP_DATE & ", Basis: " & RATE_TYPE_CODE & ", Rev Contract: " & REVENUE_CONTRACT_CODE 
                    errorStr = "Charge Rate Not found for EMPLOYEE_CODE: [" & EMPLOYEE_CODE & "]; Basis: " & RATE_TYPE_CODE
                Else
                    'No error here 
                    'errorStr = "Charge Rate Differential not found for Employee, Activity, Date: " & EMPLOYEE_CODE & "," & PAY_ACTIVITY_CODE & "," & TIME_SLIP_DATE 
                End If

                'If Not found, set to zero 
                GetChargeOutRateFactor = 0
                isPercent = 0

            Else
                'Base Pay Rate 
                GetChargeOutRateFactor = aRS("CHARGE_RATE").Value
                isPercent = aRS("IS_PERCENT").Value
            End If

        Else

            'Base Pay Rate 
            GetChargeOutRateFactor = aRS("CHARGE_RATE").Value
            isPercent = aRS("IS_PERCENT").Value

        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetDepCost(ByVal aRS As clsRecordsetExt, ByVal SD As Date, ByVal ED As Date, ByVal BlockCode As String, ByVal isCpart7Enabled As Object, ByVal isExactMatch As Integer) As Double

        Dim xRS As clsRecordsetExt
        Dim strSQL As String
        Dim ldWhereClause As String
        Dim depCost As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler


        ldWhereClause = " LS.DATE_OUT >= " & gVars.gDI & gBF.IntlDateFormat("1/1/1900") & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI &
                  " AND LS.BLOCK_CODE = |" & BlockCode & "| "

        If isExactMatch Then
            If (isCpart7Enabled <> 0) Then
                ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "| AND LS.C_PART_4_CODE = |" & aRS("C_PART_4_CODE").Value & "| AND LS.C_PART_7_CODE = |" & aRS("C_PART_7_CODE").Value & "|"
            Else
                ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "| AND LS.C_PART_4_CODE = |" & aRS("C_PART_4_CODE").Value & "|"
            End If
        Else
            If (StrComp(aRS("C_PART_4_CODE").Value, "NONE", vbTextCompare) = 0) Then
                If (isCpart7Enabled <> 0) Then
                    If (StrComp(aRS("C_PART_7_CODE").Value, "NONE", vbTextCompare) = 0) Then
                        ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "|"
                    Else
                        ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "| AND LS.C_PART_7_CODE = |" & aRS("C_PART_7_CODE").Value & "|"
                    End If
                Else
                    ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "|"
                End If
            Else
                If (isCpart7Enabled <> 0) Then
                    If (StrComp(aRS("C_PART_7_CODE").Value, "NONE", vbTextCompare) = 0) Then
                        ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "| AND LS.C_PART_4_CODE = |" & aRS("C_PART_4_CODE").Value & "|"
                    Else
                        ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "| AND LS.C_PART_4_CODE = |" & aRS("C_PART_4_CODE").Value & "| AND LS.C_PART_7_CODE = |" & aRS("C_PART_7_CODE").Value & "|"
                    End If
                Else
                    ldWhereClause = ldWhereClause & " AND LS.C_PART_5_CODE = |" & aRS("C_PART_5_CODE").Value & "| AND LS.C_PART_4_CODE = |" & aRS("C_PART_4_CODE").Value & "|"
                End If
            End If
        End If

        strSQL = "SELECT SUM(VS.PAY) AS COST FROM VENDOR_STATEMENT_DETAILS VS INNER JOIN LOADSLIPS LS ON VS.LOAD_CODE = LS.CODE " &
             " WHERE " & ldWhereClause &
             " AND VS.PAY_TYPE IN( " & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_30 & "," & LOAD_PAY_5 & " ) " &
             " AND VS.ACTIVITY_CODE = |DEPLETION|"

        xRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (xRS.EOF) Then
            depCost = 0
        Else
            If (IsDBNull(xRS(0).Value)) Then
                depCost = 0
            Else
                If (xRS(0).Value = 0) Then
                    depCost = 0
                Else
                    depCost = xRS(0).Value
                End If
            End If
        End If

        GetDepCost = depCost

        xRS = Nothing

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetDepCostHD(aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt, ByVal SD As Date, ByVal ED As Date, ByVal BlockCode As String, ByVal acoll As Collection) As Double

        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSQL As String
        Dim ldWhereClause As String
        Dim depCost As Double
        Dim cpWhere As String
        Dim i As Long
        Dim aPart As clsLib_DataAccess.clsCPart

        If (Not gDebugMode) Then On Error GoTo Err_Handler

        ldWhereClause = " LS.DATE_OUT >= " & gVars.gDI & gBF.IntlDateFormat("1/1/1900") & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI &
                  " AND LS.BLOCK_CODE = |" & BlockCode & "| "

        If (acoll.Count > 0) Then
            cpWhere = ""
            For i = 1 To acoll.Count
                aPart = acoll(i)
                If String.Compare(aPart.PartCode, "NONE", True) = 0 Then
                    'Account for the NONE WildCard
                Else
                    cpWhere = cpWhere & " AND C_PART_" & aPart.PartNo & "_CODE = |" & aPart.PartCode & "|"
                End If
            Next i
        Else
            cpWhere = ""
        End If

        strSQL = "SELECT SUM(VS.PAY) AS COST FROM VENDOR_STATEMENT_DETAILS VS INNER JOIN LOADSLIPS LS ON VS.LOAD_CODE = LS.CODE " &
             " WHERE " & ldWhereClause & cpWhere &
             " AND VS.PAY_TYPE IN( " & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_30 & "," & LOAD_PAY_5 & " ) " &
             " AND VS.ACTIVITY_CODE = |DEPLETION|"

        xRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        If (xRS.EOF) Then
            depCost = 0
        Else
            If (IsDBNull(xRS(0).Value)) Then
                depCost = 0
            Else
                If (xRS(0).Value = 0) Then
                    depCost = 0
                Else
                    depCost = xRS(0).Value
                End If
            End If
        End If

        GetDepCostHD = depCost

        xRS = Nothing

        Exit Function
Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function


    Public Function GetTheDepletionRateHD(ByVal aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt, ByVal acoll As Collection, ByVal BLOCK_CODE As String, ByVal DESTINATION_CODE As String, ByVal aDate As Date, ByVal CruiseUoM As String) As Double

        '---------------------------------------------------------------
        'Returns NA if NO benefit rate found
        '---------------------------------------------------------------

        Dim strSQL As String
        'Dim aRS As clsRecordset
        'Dim xRS As clsRecordset
        Dim rtRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim LoadVolume As Double
        Dim aCode As String
        Dim cPartColl As Collection
        Dim Conversion As Double
        Dim IS_NO_CONVERSION As Long
        Dim i As Long
        Dim errorStr As String = ""
        Dim rateUoM As String
        Dim aRate As Double
        Dim j As Long
        Dim aPart As clsCPart
        Dim strInnerJoin As String
        Dim cpWhere As String
        Dim orderBy As String
        Dim leftParens As String

        If (Not gDebugMode) Then On Error GoTo Err_Handler

        leftParens = ""
        If (acoll.Count > 0) Then
            strInnerJoin = ""
            For i = 1 To acoll.Count
                aPart = acoll(i)
                strInnerJoin = strInnerJoin & " INNER JOIN C_PART_" & aPart.PartNo & " AS CP" & aPart.PartNo & " ON R.C_PART_" & aPart.PartNo & "_CODE = CP" & aPart.PartNo & ".CODE) "
                leftParens = leftParens & "("
            Next i
        Else
            strInnerJoin = ""
        End If

        '    MsgBox "strInnerJoin: " & strInnerJoin

        If (acoll.Count > 0) Then
            cpWhere = ""
            For i = 1 To acoll.Count
                aPart = acoll(i)
                cpWhere = cpWhere & " AND (R.C_PART_" & aPart.PartNo & "_CODE = |NONE| OR R.C_PART_" & aPart.PartNo & "_CODE = |" & aRS("C_PART_" & aPart.PartNo & "_CODE").Value & "| )"
            Next i
        Else
            cpWhere = ""
        End If

        '    MsgBox "cpWhere: " & cpWhere

        orderBy = " ORDER BY "
        If (acoll.Count > 0) Then
            For i = 1 To acoll.Count
                aPart = acoll(i)
                orderBy = orderBy & " CP" & aPart.PartNo & ".IS_NONE ASC, "
            Next i
        Else
            '    orderBy = ""
        End If

        orderBy = orderBy & "BL.IS_NONE ASC, R.EFFECTIVE_DATE DESC "

        '    MsgBox "orderBy: " & orderBy

        '    strSQL = "SELECT TOP 1 RATE, RATE_TYPE_CODE " & _
        '             " FROM (((( PAY_CONTRACT_RATES AS R " & _
        '             " INNER JOIN BLOCKS AS BL ON R.BLOCK_CODE = BL.CODE) " & _
        '             " INNER JOIN C_PART_4 AS CP4 ON R.C_PART_4_CODE = CP4.CODE) " & _
        '             " INNER JOIN C_PART_5 AS CP5 ON R.C_PART_5_CODE = CP5.CODE) " & _
        '             " INNER JOIN C_PART_7 AS CP7 ON R.C_PART_7_CODE = CP7.CODE) " & _
        '             " WHERE (R.BLOCK_CODE = |NONE| OR R.BLOCK_CODE = |" & BLOCK_CODE & "|) " & _
        '             " AND (R.C_PART_4_CODE = |NONE| OR R.C_PART_4_CODE = |" & C_PART_4_CODE & "| )" & _
        '             " AND (R.C_PART_5_CODE = |NONE| OR R.C_PART_5_CODE = |" & C_PART_5_CODE & "| )" & _
        '             " AND (R.C_PART_7_CODE = |NONE| OR R.C_PART_7_CODE = |" & C_PART_7_CODE & "| )" & _
        '             " AND (R.PAY_ACTIVITY_CODE = |" & "DEPLETION" & "| )" & _
        '             " AND R.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI & _
        '             " ORDER BY CP4.IS_NONE ASC, CP5.IS_NONE ASC, CP7.IS_NONE ASC, BL.IS_NONE ASC, R.EFFECTIVE_DATE DESC "


        strSQL = "SELECT TOP 1 RATE, RATE_TYPE_CODE " &
                 " FROM " & leftParens & "( PAY_CONTRACT_RATES AS R " &
                 " INNER JOIN BLOCKS AS BL ON R.BLOCK_CODE = BL.CODE) " &
                 strInnerJoin &
                 " WHERE (R.BLOCK_CODE = |NONE| OR R.BLOCK_CODE = |" & BLOCK_CODE & "|) " &
                 cpWhere &
                 " AND (R.PAY_ACTIVITY_CODE = |" & "DEPLETION" & "| )" &
                 " AND R.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                 orderBy

        '    MsgBox "strSQL: " & strSQL

        rtRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        If (rtRS.EOF) Then
            'MsgBox "NO RATE strSQL: " & strSQL
            GetTheDepletionRateHD = 0
            Exit Function
        Else
            rateUoM = rtRS("RATE_TYPE_CODE").Value
            aRate = rtRS("RATE").Value
        End If

        'MsgBox "My DepRate: " & aRate

        If (StrComp(rateUoM, CruiseUoM, vbTextCompare) = 0) Then

            'We're good to go
            GetTheDepletionRateHD = aRate

        Else
            'Convert

            cPartColl = New Collection

            aCode = DESTINATION_CODE
            cPartColl.Add(aCode, CStr(DESTINATION))

            For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                aCode = "NONE"
                For j = 1 To acoll.Count
                    aPart = acoll(j)
                    If aPart.PartNo = i Then
                        aCode = aRS("C_PART_" & i & "_CODE").Value
                        Exit For
                    End If
                Next j
                cPartColl.Add(aCode, CStr(i))
            Next i

            'Try to Convert weight into payBasis
            Conversion = GetConversionRate(BLOCK_CODE, aDate, CruiseUoM, rateUoM, cPartColl, NA, NA, "", errorStr)

            'MsgBox "Conversion: " & Conversion & ";  CruiseUoM = " & CruiseUoM & "; rateUoM = " & rateUoM

            If Math.Abs(Conversion - VOL_EQUALS_SCALE_VOLUME) < EPS Then

                'Should never go here...vol = NA above

            ElseIf (Len(errorStr) = 0 And Math.Abs(Conversion - NA) > EPS) Then

                aRate = Conversion * aRate

            Else

                Conversion = GetConversionRate(BLOCK_CODE, aDate, rateUoM, CruiseUoM, cPartColl, NA, NA, "", errorStr)

                If Math.Abs(Conversion - VOL_EQUALS_SCALE_VOLUME) < EPS Then

                    'Should never go here...vol = NA above

                ElseIf (Len(errorStr) = 0 And Math.Abs(Conversion - NA) > EPS) Then

                    If (Math.Abs(Conversion) > EPS) Then
                        aRate = aRate / Conversion
                    Else
                        aRate = 0
                    End If

                Else

                    aRate = 0

                End If
            End If

            GetTheDepletionRateHD = aRate

        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function




    Public Function GetTheDepletionRate(ByVal C_PART_4_CODE As String, ByVal C_PART_5_CODE As String, ByVal C_PART_7_CODE As String,
                                        ByVal BLOCK_CODE As String, ByVal DESTINATION_CODE As String, ByVal aDate As Date, ByVal CruiseUoM As String) As Double

        '--------------------------------------------------------------- 
        'Returns NA if NO benefit rate found 
        '--------------------------------------------------------------- 

        Dim strSQL As String
        Dim aRS As clsRecordsetExt
        Dim xRS As clsRecordsetExt
        Dim rtRS As clsRecordsetExt
        Dim LoadVolume As Double
        Dim aCode As String
        Dim cPartColl As Collection
        Dim Conversion As Double
        Dim IS_NO_CONVERSION As Integer
        Dim i As Integer
        Dim errorStr As String = ""
        Dim rateUoM As String
        Dim aRate As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        strSQL = "SELECT TOP 1 RATE, RATE_TYPE_CODE " &
                 " FROM (((( PAY_CONTRACT_RATES AS R " &
                 " INNER JOIN BLOCKS AS BL ON R.BLOCK_CODE = BL.CODE) " &
                 " INNER JOIN C_PART_4 AS CP4 ON R.C_PART_4_CODE = CP4.CODE) " &
                 " INNER JOIN C_PART_5 AS CP5 ON R.C_PART_5_CODE = CP5.CODE) " &
                 " INNER JOIN C_PART_7 AS CP7 ON R.C_PART_7_CODE = CP7.CODE) " &
                 " WHERE (R.BLOCK_CODE = |NONE| OR R.BLOCK_CODE = |" & BLOCK_CODE & "|) " &
                 " AND (R.C_PART_4_CODE = |NONE| OR R.C_PART_4_CODE = |" & C_PART_4_CODE & "| )" &
                 " AND (R.C_PART_5_CODE = |NONE| OR R.C_PART_5_CODE = |" & C_PART_5_CODE & "| )" &
                 " AND (R.C_PART_7_CODE = |NONE| OR R.C_PART_7_CODE = |" & C_PART_7_CODE & "| )" &
                 " AND (R.PAY_ACTIVITY_CODE = |" & "DEPLETION" & "| )" &
                 " AND R.EFFECTIVE_DATE <= " & gVars.gDI & gBF.IntlDateFormat(aDate) & gVars.gDI &
                 " ORDER BY CP4.IS_NONE ASC, CP5.IS_NONE ASC, CP7.IS_NONE ASC, BL.IS_NONE ASC, R.EFFECTIVE_DATE DESC "

        rtRS = gDataLayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            gBF.DisplayErrors(basLocals.gObjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Return 0
            Exit Function
        End If

        If (rtRS.EOF) Then
            GetTheDepletionRate = 0
            Exit Function
        Else
            rateUoM = rtRS("RATE_TYPE_CODE").Value
            aRate = rtRS("RATE").Value
        End If

        If (StrComp(rateUoM, CruiseUoM, vbTextCompare) = 0) Then
            'We're good to go 
            GetTheDepletionRate = aRate

        Else
            'Convert 



            cPartColl = New Collection

            aCode = DESTINATION_CODE
            cPartColl.Add(aCode, CStr(DESTINATION))

            For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                Select Case i
                    Case 4
                        aCode = C_PART_4_CODE
                    Case 5
                        aCode = C_PART_5_CODE
                    Case 7
                        aCode = C_PART_7_CODE
                    Case Else
                        aCode = "NONE"
                End Select
                cPartColl.Add(aCode, CStr(i))
            Next i

            'Try to Convert weight into payBasis 
            Conversion = GetConversionRate(BLOCK_CODE, aDate, CruiseUoM, rateUoM, cPartColl, NA, NA, "", errorStr)

            If Math.Abs(Conversion - VOL_EQUALS_SCALE_VOLUME) < EPS Then

                'Should never go here...vol = NA above 

            ElseIf (Len(errorStr) = 0 And Math.Abs(Conversion - NA) > EPS) Then

                aRate = Conversion * aRate

            Else

                Conversion = GetConversionRate(BLOCK_CODE, aDate, rateUoM, CruiseUoM, cPartColl, NA, NA, "", errorStr)

                If Math.Abs(Conversion - VOL_EQUALS_SCALE_VOLUME) < EPS Then

                    'Should never go here...vol = NA above 

                ElseIf (Len(errorStr) = 0 And Math.Abs(Conversion - NA) > EPS) Then

                    If (Math.Abs(Conversion) > EPS) Then
                        aRate = aRate / Conversion
                    Else
                        aRate = 0
                    End If

                Else

                    aRate = 0

                End If
            End If

            GetTheDepletionRate = aRate

        End If

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function




End Module
