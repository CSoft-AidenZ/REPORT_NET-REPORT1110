Option Strict Off
Option Explicit On
'<System.Runtime.InteropServices.ProgId("clsRpt1110_NET.clsRpt1110")> Public Class clsRpt1110
Public Class clsRpt1110
    'Public Function RunReport(ByRef aForm As Object, ByRef frm_Filter As Object, ByRef Vars As clsLib_DataAccess.clsVars, ByRef objTables As clsLib_DataAccess.clsTables, ByRef objerrors As clsLib_DataAccess.clsErrors) As Object
    Public Function RunReport(ByVal aform As Object, ByVal frm_Filter As Object, ByVal myCalcs As clsCaribouCalculators.clsCalcs, ByVal objErrors As clsLib_DataAccess.clsErrors) As Integer
        RunReport = 0

        Dim crit As String
        Dim crit2 As String
        Dim aColl1 As Collection
        Dim aColl2 As Collection
        Dim SD As Date
        Dim ED As Date
        Dim OrderBy As Integer
        Dim ret As Integer
        Dim PrintRawData As Integer
        Dim doLoadedCost As Integer
        Dim doSummary As Integer
        Dim useInvoiceRevenue As Integer
        Dim doAll As Integer
        Dim dateForPay As Integer

        'gVars = Vars
        'gobjTables = objTables
        'gobjErrors = objErrors


        gVars = myCalcs.get_gVars
        gDataLayer = myCalcs.get_gDataLayer
        gobjTables = myCalcs.get_gobjTables
        gobjTables.Datalayer = gDataLayer
        gobjErrors = objErrors
        gConst = New CSOFT_BASIC_FUNCTIONS.clsConstants
        gBF = New CSOFT_BASIC_FUNCTIONS.clsBasicFunctions(gVars.gDebugMode)
        gComFn = New clsBasComnFuncs.clsCommonFuncs(myCalcs.getCalcInterface, gVars.gDebugMode, gVars.gLicMode)
        basReportGlobals = New clsLib_ReportClasses.claBasReportGlobals(myCalcs.getCalcInterface, gVars.gDebugMode, gVars.gLicMode)
        basReportDllFunctions = New clsLib_ReportClasses.clsBasReportDllFunctions(myCalcs.getCalcInterface, gVars.gDebugMode, gVars.gLicMode)
        If (basLocals.gVars.gDebugMode = 0) Then
            On Error GoTo Err_Handler
        End If

        crit = "SELECT CODE, DESCRIPTION FROM CREW_CODES WHERE IS_ACTIVE <> 0 ORDER BY CODE "

        aColl1 = New Collection
        aColl2 = New Collection

        frm_Filter.Frame2.Visible = False
        frm_Filter.check3.Visible = False
        'frm_Filter.option1(0).Caption = gVars.gBasicSetup.Value("GEOG_WORK_AREA_LABEL")
        frm_Filter.combo1.Enabled = True
        frm_Filter.txtstartdate.Enabled = True
        frm_Filter.txtenddate.Enabled = True
        frm_Filter.Frame1.Enabled = True
        frm_Filter.check3.Visible = True
        frm_Filter.check3.Caption = "Use Date for Pay"

        ret = frm_Filter.LoadMeNet(crit, "", aColl1, aColl2, SD, ED, "DESCRIPTION", "Crew", "", OrderBy, dateForPay, "", "Use Invoice Revenue", useInvoiceRevenue, 1, doSummary, gConst.PP_REPORTING)

        If (ret) Then
            ret = DoCruise(aform, aColl1, aColl2, SD, ED, OrderBy, doLoadedCost, PrintRawData, doSummary, useInvoiceRevenue, dateForPay)
        End If

        RunReport = ret

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function



    Private Function DoCruise(ByRef aForm As Object, ByRef aColl1 As Collection, ByRef aColl2 As Collection, ByVal xsd As Date, ByVal xED As Date, ByVal OrderBy As Integer, ByVal doLoadedCost As Integer, ByVal PrintRawData As Integer, ByVal doSummary As Integer, ByVal doInvoiceRevenue As Integer, ByVal dateForPay As Integer) As Integer

        Dim strSQL As String
        Dim crewStr As String
        Dim lngRow As Integer
        Dim i, j As Integer
        Dim c As Integer
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim Totals() As Double
        Dim blTotals() As Double
        Dim grTotals() As Double
        Dim n As Integer
        Dim Page As Integer
        Dim tmpColl1 As Collection
        Dim tmpColl2 As Collection
        Dim fld1 As String
        Dim stb1 As String
        Dim Vdec As Integer
        Dim TCost As Double
        Dim rev As Double
        Dim VolCut As Double
        Dim LoadCount As Double
        Dim vcColl As Collection
        Dim dpColl As Collection
        Dim volCode As String
        Dim depCost As Double
        Dim vColl As Collection
        Dim tmpItem As String
        Dim aPart As clsLib_DataAccess.clsCPart
        Dim ldWhereClause As String
        Dim aVol As clsLib_ReportClasses.clsAdvance
        Dim aKey As String
        Dim dec As Integer
        Dim aDate As String
        Dim SD As Date
        Dim ED As Date
        Dim pi As Double
        Dim xPay As Double
        Dim cstLogging As Double
        Dim revTable As String
        Dim acres As Double
        Dim m As Integer
        Dim volDiff As Integer
        Dim oldvolCode As String
        Dim yMisc As Double
        Dim xMisc As Double
        Dim dateCriteria As String

        If dateForPay <> 0 Then
            dateCriteria = "PAY_DATE"
        Else
            dateCriteria = "DATE_OUT"
        End If

        If (Not basLocals.gVars.gDebugMode) Then On Error GoTo Err_Handler

        gBF.TurnMeOn()


        'Short Circuit selection by only using blocks where we have loads in the date range
        Dim crewSelection As Collection
        crewSelection = aColl1

        Dim crewCodes As String
        Dim Crew As Short
        For Crew = 1 To crewSelection.Count

            crewCodes = crewCodes & "|" & crewSelection(Crew) & "|"
            If Crew <> crewSelection.Count Then
                crewCodes = crewCodes & ","
            End If
        Next Crew



        '-------------------------------------------------
        '    Get the Pay Period in Selected Range
        '-------------------------------------------------

        strSQL = " SELECT PP.* FROM PAY_PERIODS PP INNER JOIN PAY_PERIOD_TYPES PT ON PP.PAY_PERIOD_TYPE_CODE = PT.PAY_PERIOD_TYPE_CODE " & " WHERE DATEADD(|D|,1,END_DATE) > " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI & " AND START_DATE < " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI & " AND PT.GENERIC_TYPE_CODE IN (" & gConst.PP_ALL & "," & gConst.PP_VENDORS & ") " & " AND CODE <> |SPOT| " & " ORDER BY PP.START_DATE "


        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        If (aRS.EOF) Then
            gBF.TurnMeOff()
            MsgBox("No Pay Periods in Selected Date Range!", MsgBoxStyle.Information)
            Exit Function
        Else
            aRS.MoveFirst()
            xsd = CDate(aRS("START_DATE").Value)
            aRS.MoveLast()
            xED = DateAdd(DateInterval.Day, 1, CDate(aRS("END_DATE").Value))
        End If





        strSQL = "SELECT CREW_CODE FROM LOADSLIPS " & " WHERE " & dateCriteria & " <  " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI & " AND " & dateCriteria & " >= " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI & " AND CREW_CODE IN (" & crewCodes & ") " & " GROUP BY CREW_CODE HAVING SUM(LOAD_COUNT) > 0 "

        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        tmpColl1 = New Collection

        Do While Not aRS.EOF
            tmpColl1.Add(aRS("CREW_CODE").Value)
            aRS.MoveNext()
        Loop

        vcColl = New Collection
        dpColl = New Collection

        fld1 = "CREW_CODE"
        stb1 = "Crew"

        If (aColl1.Count = 0) Then
            MsgBox("No " & stb1 & " Selected.", MsgBoxStyle.Information)
            Exit Function
        End If

        Page = 0
        volDiff = 0

        Vdec = gobjTables("LOADSLIPS").Fields("VOLUME").Decimals
        dec = gobjTables("LOADSLIPS").Fields("GROSS").Decimals

        If (doInvoiceRevenue) Then
            revTable = "INVOICE_DETAILS"
        Else
            revTable = "LOADSLIP_REVENUE"
        End If

        aForm.aGrid(Page).rowHeight(0) = 21

        aForm.aGrid(Page).SetText(2, 0, "Crew")
        aForm.aGrid(Page).SetText(3, 0, "Pay Period")
        aForm.aGrid(Page).SetText(4, 0, "PP End Date")
        aForm.aGrid(Page).SetText(5, 0, "Loads")
        aForm.aGrid(Page).SetText(6, 0, "Volume")
        aForm.aGrid(Page).SetText(7, 0, "Measure")
        aForm.aGrid(Page).SetText(8, 0, "Revenue")
        aForm.aGrid(Page).SetText(9, 0, "Avg Rate")
        aForm.aGrid(Page).SetText(10, 0, "Total Cost")
        aForm.aGrid(Page).SetText(11, 0, "Avg Rate")
        aForm.aGrid(Page).SetText(12, 0, "Gr. Profit")
        aForm.aGrid(Page).SetText(13, 0, "Avg Rate")
        aForm.aGrid(Page).SetText(14, 0, "Gr Margin (%)")
        n = 14
        aForm.aGrid(Page).MaxCols = n

        For i = 1 To n
            aForm.aGrid(Page).colWidth(i) = 10
        Next i

        lngRow = -1
        crewStr = ""

        '-------------------------------------------------
        '    Get the Pay Period in Selected Range
        '-------------------------------------------------

        strSQL = " SELECT PP.* FROM PAY_PERIODS PP INNER JOIN PAY_PERIOD_TYPES PT ON PP.PAY_PERIOD_TYPE_CODE = PT.PAY_PERIOD_TYPE_CODE " & " WHERE DATEADD(|D|,1,END_DATE) > " & gVars.gDI & gBF.IntlDateFormat(xsd) & gVars.gDI & " AND START_DATE < " & gVars.gDI & gBF.IntlDateFormat(xED) & gVars.gDI & " AND PT.GENERIC_TYPE_CODE IN (" & gConst.PP_ALL & "," & gConst.PP_VENDORS & ") " & " AND CODE <> |SPOT| " & " ORDER BY PP.START_DATE "

        '
        'BL 7/25/2013 -> Changed to strictly less than because xED has a day already added to it: START_DATE < " & gVars.gDI & gBF.IntlDateFormat(xed) & gVars.gDI &
        '

        'DP added so we close aRS before opening a new one.
        'aRS = Nothing
        aRS.Close1()

        aRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        If (aRS.EOF) Then
            gBF.TurnMeOff()
            MsgBox("No Pay Periods in Selected Date Range!", MsgBoxStyle.Information)
            Exit Function
        End If

        oldvolCode = ""

        For i = 1 To tmpColl1.Count + IIf(tmpColl1.Count > 1, 1, 0) 'The last one is the grand totals row

            If (i <= tmpColl1.Count) Then
                tmpItem = tmpColl1(i)
                crewStr = crewStr & "|" & tmpItem & "|, "
                If tmpColl1.Count = 1 Then
                    'There is only one crew, so summary mode should bring back only one
                Else
                    If doSummary Then GoTo NextCrew
                End If
            ElseIf (i = tmpColl1.Count + 1) Then
                crewStr = Left(crewStr, Len(crewStr) - 3)
                crewStr = Right(crewStr, Len(crewStr) - 1)
                tmpItem = crewStr
            End If


            '--------------------------------------------------------------------
            '   Loop Over Crews
            '--------------------------------------------------------------------
            '--------------------------------------------------------
            'Find the Date Range of Loads
            '--------------------------------------------------------

            lngRow = lngRow + 2
            aForm.aGrid(Page).MaxRows = lngRow

            If (i <= tmpColl1.Count) Then

                lngRow = lngRow + 1
                aForm.aGrid(Page).MaxRows = lngRow
                aForm.aGrid(Page).SetText(1, lngRow, "Crew Code:")
                aForm.aGrid(Page).SetText(2, lngRow, CStr(tmpColl1(i)))
                aForm.aGrid(Page).Col = 1
                aForm.aGrid(Page).Row = lngRow
                aForm.aGrid(Page).AllowCellOverflow = True

                lngRow = lngRow + 1
                aForm.aGrid(Page).MaxRows = lngRow
                aForm.aGrid(Page).SetText(1, lngRow, "Crew Name:")
                aForm.aGrid(Page).SetText(2, lngRow, CStr(gComFn.getAValue("DESCRIPTION", "CREW_CODES", "CODE", tmpColl1(i))))
                aForm.aGrid(Page).Col = 1
                aForm.aGrid(Page).Row = lngRow
                aForm.aGrid(Page).AllowCellOverflow = True

                m = 1

                'If (gobjTables("BLOCKS").fields("FLOAT_1").DisplayOrder > 0) Then
                'lngRow = lngRow + 1
                'm = m + 1
                'aForm.aGrid(Page).MaxRows = lngRow
                'aForm.aGrid(Page).SetText 1, lngRow, gobjTables("BLOCKS").fields("FLOAT_1").DisplayName
                'aForm.aGrid(Page).SetText 2, lngRow, gComFn.getAValue("FLOAT_1", "BLOCKS", "CODE", tmpColl1(i))
                'acres = val(gComFn.getAValue("FLOAT_1", "BLOCKS", "CODE", tmpColl1(i)))
                'FloatCellEnter aForm.aGrid(Page), 2, lngRow, acres, 1, False, PrintRawData
                'aForm.aGrid(Page).Col = 1
                'aForm.aGrid(Page).Row = lngRow
                'aForm.aGrid(Page).AllowCellOverflow = True
                'End If

                'If (gobjTables("BLOCKS").fields("DATE_2").DisplayOrder > 0) Then
                'lngRow = lngRow + 1
                'm = m + 1
                'aForm.aGrid(Page).MaxRows = lngRow
                'aForm.aGrid(Page).SetText 1, lngRow, gobjTables("BLOCKS").fields("DATE_2").DisplayName
                'aDate = CStr("" & gComFn.getAValue("DATE_2", "BLOCKS", "CODE", tmpColl1(i)))
                'If (Len(aDate) = 0) Then
                'aDate = "Unknown"
                'ElseIf (Not IsDate(aDate)) Then
                'aDate = "Unknown"
                'ElseIf CDbl(CDate(aDate)) = 2 Then
                'aDate = "Unknown"
                'End If
                'aForm.aGrid(Page).SetText 2, lngRow, CStr(aDate)
                'aForm.aGrid(Page).Col = 1
                'aForm.aGrid(Page).Row = lngRow
                'aForm.aGrid(Page).AllowCellOverflow = True
                'End If


                'If (gobjTables("BLOCKS").fields("EFFECTIVE_DATE").DisplayOrder > 0) Then
                'lngRow = lngRow + 1
                'm = m + 1
                'aForm.aGrid(Page).MaxRows = lngRow
                'aForm.aGrid(Page).SetText 1, lngRow, gobjTables("BLOCKS").fields("EFFECTIVE_DATE").DisplayName
                'aDate = CStr("" & gComFn.getAValue("EFFECTIVE_DATE", "BLOCKS", "CODE", tmpColl1(i)))
                'If (Len(aDate) = 0) Then
                'aDate = "Unknown"
                'ElseIf (Not IsDate(aDate)) Then
                'aDate = "Unknown"
                'ElseIf CDbl(CDate(aDate)) = 2 Then
                'aDate = "Unknown"
                'End If
                'aForm.aGrid(Page).SetText 2, lngRow, CStr(aDate)
                'aForm.aGrid(Page).Col = 1
                'aForm.aGrid(Page).Row = lngRow
                'aForm.aGrid(Page).AllowCellOverflow = True
                'End If


                'If (gobjTables("BLOCKS").fields("DATE_1").DisplayOrder > 0) Then
                'lngRow = lngRow + 1
                'm = m + 1
                'aForm.aGrid(Page).MaxRows = lngRow
                'aForm.aGrid(Page).SetText 1, lngRow, gobjTables("BLOCKS").fields("DATE_1").DisplayName
                'aDate = CStr("" & gComFn.getAValue("DATE_1", "BLOCKS", "CODE", tmpColl1(i)))
                'If (Len(aDate) = 0) Then
                'aDate = "Unknown"
                'ElseIf (Not IsDate(aDate)) Then
                'aDate = "Unknown"
                'ElseIf CDbl(CDate(aDate)) = 2 Then
                'aDate = "Unknown"
                'End If
                'aForm.aGrid(Page).SetText 2, lngRow, CStr(aDate)
                'aForm.aGrid(Page).Col = 1
                'aForm.aGrid(Page).Row = lngRow
                'aForm.aGrid(Page).AllowCellOverflow = True
                'End If

                aForm.aGrid(Page).Row = lngRow - m
                aForm.aGrid(Page).Row2 = lngRow
                aForm.aGrid(Page).Col = 1
                aForm.aGrid(Page).Col2 = n
                aForm.aGrid(Page).BlockMode = True
                aForm.aGrid(Page).FontBold = True
                aForm.aGrid(Page).BlockMode = False

                'volCode = gComFn.getAValue("VOLUME_UOM_CODE", "BLOCKS", "CODE", tmpColl1(i))
                'If (Len(oldvolCode) > 0 And StrComp(volCode, oldvolCode, vbTextCompare) <> 0) Then
                'volDiff = 1
                'End If
                'oldvolCode = volCode

            Else

                lngRow = lngRow + 1
                aForm.aGrid(Page).MaxRows = lngRow
                aForm.aGrid(Page).SetText(1, lngRow, "Crew Code:")
                aForm.aGrid(Page).SetText(2, lngRow, "All Selected Crews")
                aForm.aGrid(Page).Col = 1
                aForm.aGrid(Page).Row = lngRow
                aForm.aGrid(Page).AllowCellOverflow = True

                aForm.aGrid(Page).Row = lngRow
                aForm.aGrid(Page).Row2 = lngRow
                aForm.aGrid(Page).Col = 1
                aForm.aGrid(Page).Col2 = n
                aForm.aGrid(Page).BlockMode = True
                aForm.aGrid(Page).FontBold = True
                aForm.aGrid(Page).BlockMode = False

                'strSQL = " SELECT DISTINCT VOLUME_UOM_CODE FROM BLOCKS WHERE CODE IN (|" & tmpItem & "|) " 'This is no longer necessary DP 9/3/2015

                'Set xRS = gDatalayer.LoadRecordset(strSQL, gobjErrors)
                'If (gobjErrors.Count = 0) Then
                'Else
                'DisplayErrors
                'Exit Function
                'End If

                'If Not xRS.EOF Then
                'm = 0
                'Do While xRS.EOF = False
                'm = m + 1
                'volCode = xRS(0)
                'xRS.MoveNext
                'Loop
                'If (m > 1) Then
                'volCode = "****"
                'Else
                'volCode = volCode  'Only One
                'End If
                'Set xRS = Nothing
                'Else
                'volCode = "****"
                'End If

            End If

            lngRow = lngRow + 1
            aForm.aGrid(Page).MaxRows = lngRow

            ReDim Totals(n)
            ReDim blTotals(n)


            '-----------------------------------------------------------
            '  Loop Over Pay Periods
            '-----------------------------------------------------------

            Do While aRS.EOF = False

                volCode = ""
                oldvolCode = ""

                SD = aRS("START_DATE").Value
                ED = CDate(aRS("END_DATE").Value).AddDays(1)

                strSQL = " SELECT DISTINCT T1.VOLUME_UOM_CODE FROM LOADSLIPS T0 " & " INNER JOIN BLOCKS T1 ON T0.BLOCK_CODE = T1.CODE " & " WHERE T0.CREW_CODE IN (|" & tmpItem & "|) " & " AND T0." & dateCriteria & " <  " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND T0." & dateCriteria & " >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND (T0.C_PART_13_CODE = |NONE| OR T0.C_PART_13_CODE = |H| OR T0.C_PART_13_CODE IS NULL) " & " GROUP BY T1.VOLUME_UOM_CODE "

                xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                If (gobjErrors.Count = 0) Then
                Else
                    basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                    Exit Function
                End If

                Do While xRS.EOF = False

                    volCode = CStr(xRS("VOLUME_UOM_CODE").Value)

                    If (Len(oldvolCode) > 0 And StrComp(volCode, oldvolCode, CompareMethod.Text) <> 0) Then
                        volCode = "****"
                        volDiff = 1
                    End If
                    oldvolCode = volCode

                    xRS.MoveNext()
                Loop

                'xRS = Nothing
                xRS.Close1()


                strSQL = " SELECT SUM(T0.LOAD_COUNT) AS LOADCOUNT, SUM(T0.GROSS) AS GROSS, SUM(T0.TARE) AS TARE, SUM(T0.NET) AS NET, SUM(T0.VOLUME) AS VOLUME FROM LOADSLIPS T0 " & " INNER JOIN BLOCKS T1 ON T0.BLOCK_CODE = T1.CODE " & " WHERE T0.CREW_CODE IN (|" & tmpItem & "|) " & " AND T0." & dateCriteria & " <  " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND T0." & dateCriteria & " >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND (T0.C_PART_13_CODE = |NONE| OR T0.C_PART_13_CODE = |H| OR T0.C_PART_13_CODE IS NULL) "
                '" GROUP BY T1.VOLUME_UOM_CODE "

                xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                If (gobjErrors.Count = 0) Then
                Else
                    basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                    Exit Function
                End If

                'Do While xRS.EOF = False


                If xRS.EOF = False Then

                    LoadCount = IIf(IsDBNull(xRS("LOADCOUNT").Value), 0, CDbl(xRS("LOADCOUNT").Value))

                    If (LoadCount > 0) Then

                        VolCut = IIf(IsDBNull(xRS("VOLUME").Value), 0, CDbl(xRS("VOLUME").Value))


                        lngRow = lngRow + 1
                        aForm.aGrid(Page).MaxRows = lngRow

                        aForm.aGrid(Page).SetText(3, lngRow, CStr(aRS("CODE").Value))
                        aForm.aGrid(Page).SetText(4, lngRow, CStr(FormatDateTime(aRS("END_DATE").Value, DateFormat.ShortDate)))

                        gBF.FloatCellEnter(aForm.aGrid(Page), 5, lngRow, LoadCount, 0, False, PrintRawData)
                        gBF.FloatCellEnter(aForm.aGrid(Page), 6, lngRow, VolCut, dec, False, PrintRawData)
                        aForm.aGrid(Page).SetText(7, lngRow, "" & volCode)

                        '-------------------------------------------------------------------
                        '    Revenue
                        '-------------------------------------------------------------------

                        strSQL = " SELECT SUM(R0.PAY) AS REVENUE FROM LOADSLIPS LS INNER JOIN " & revTable & " R0 ON R0.LOAD_CODE = LS.CODE " & " WHERE LS.CREW_CODE IN (|" & tmpItem & "|) AND LS." & dateCriteria & " >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND LS." & dateCriteria & " < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI

                        If (doInvoiceRevenue) Then
                            strSQL = strSQL & " AND R0.PAY_TYPE IN (" & gConst.REV_CONTRACT_LOADS & ", " & gConst.REV_CONTRACT_CUSTOMER_LOADS & ")"
                        Else
                            'Do Nada - Loadslip Revenue
                        End If

                        'xRS = Nothing
                        xRS.Close1() 'DP want to close this before we open a new one, even though it resides in an If statement using the prior connection we are done with it at this point.

                        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                        If (gobjErrors.Count = 0) Then
                        Else
                            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                            Exit Function
                        End If

                        rev = IIf(IsDBNull(xRS("REVENUE").Value), 0, CDbl(xRS("REVENUE").Value))

                        '-------------------------------------------------------------------
                        '     Miscellaneous Income
                        '-------------------------------------------------------------------

                        strSQL = " SELECT SUM(PAY) AS PAYOLA FROM MISCELLANEOUS_INCOME LS INNER JOIN INVOICE_DETAILS R0 ON R0.LOAD_CODE = LS.CODE " & " WHERE LS.CREW_CODE IN (|" & tmpItem & "|) AND LS.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND LS.ENTRY_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND R0.PAY_TYPE IN (" & gConst.REV_CONTRACT_MISC & ")"

                        'xRS = Nothing
                        xRS.Close1() 'DP added so we close this before using it again

                        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                        If (gobjErrors.Count = 0) Then
                        Else
                            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                            Exit Function
                        End If

                        yMisc = IIf(IsDBNull(xRS("PAYOLA").Value), 0, CDbl(xRS("PAYOLA").Value))

                        rev = rev + yMisc

                        gBF.FloatCellEnter(aForm.aGrid(Page), 8, lngRow, rev, 2, True, PrintRawData)
                        If (System.Math.Abs(VolCut) > 0) Then
                            gBF.FloatCellEnter(aForm.aGrid(Page), 9, lngRow, rev / VolCut, 2, True, PrintRawData)
                        Else
                            gBF.FloatCellEnter(aForm.aGrid(Page), 9, lngRow, 0, 2, True, PrintRawData)
                        End If




                        '-------------------------------------------------------------------
                        '     Total Costs
                        '-------------------------------------------------------------------

                        strSQL = " SELECT SUM(PAY) AS PAYOLA FROM LOADSLIPS LS INNER JOIN VENDOR_STATEMENT_DETAILS R0 ON R0.LOAD_CODE = LS.CODE " & " WHERE LS.CREW_CODE IN (|" & tmpItem & "|) AND LS." & dateCriteria & " >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND LS." & dateCriteria & " < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND R0.PAY_TYPE IN (" & gConst.LOAD_PAY_1 & ", " & gConst.LOAD_PAY_3 & ", " & gConst.LOAD_PAY_5 & ", " & gConst.LOAD_PAY_30 & ")"

                        'xRS = Nothing
                        xRS.Close1() 'DP added so we close this before using it again

                        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                        If (gobjErrors.Count = 0) Then
                        Else
                            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                            Exit Function
                        End If

                        xPay = IIf(IsDBNull(xRS("PAYOLA").Value), 0, CDbl(xRS("PAYOLA").Value))

                        strSQL = " SELECT SUM(PAY) AS PAYOLA FROM LOADSLIPS LS INNER JOIN EMPLOYEE_STATEMENT_DETAILS R0 ON R0.LOAD_CODE = LS.CODE " & " WHERE LS.CREW_CODE IN (|" & tmpItem & "|) AND LS." & dateCriteria & " >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND LS." & dateCriteria & " < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND R0.PAY_TYPE IN (" & gConst.LOAD_PAY_1 & ", " & gConst.LOAD_PAY_3 & ", " & gConst.LOAD_PAY_5 & ", " & gConst.LOAD_PAY_30 & ")" & " AND R0.ACTIVITY_CODE NOT IN (|DEPLETION|, |STUMPAGE|) "

                        'xRS = Nothing
                        xRS.Close1() 'DP added so we close this before using it again

                        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                        If (gobjErrors.Count = 0) Then
                        Else
                            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                            Exit Function
                        End If

                        xPay = xPay + IIf(IsDBNull(xRS("PAYOLA").Value), 0, CDbl(xRS("PAYOLA").Value))

                        '-------------------------------------------------------------------
                        '     Miscellaneous Expenses
                        '-------------------------------------------------------------------

                        strSQL = " SELECT SUM(PAY) AS PAYOLA FROM MISCELLANEOUS_EXPENSES LS INNER JOIN VENDOR_STATEMENT_DETAILS R0 ON R0.LOAD_CODE = LS.CODE " & " WHERE LS.CREW_CODE IN (|" & tmpItem & "|) AND LS.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND LS.ENTRY_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND R0.PAY_TYPE IN (" & gConst.LOAD_PAY_MISC & ")"

                        'xRS = Nothing
                        xRS.Close1() 'DP added so we close this before using it again

                        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                        If (gobjErrors.Count = 0) Then
                        Else
                            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                            Exit Function
                        End If

                        xMisc = IIf(IsDBNull(xRS("PAYOLA").Value), 0, CDbl(xRS("PAYOLA").Value))

                        strSQL = " SELECT SUM(PAY) AS PAYOLA FROM EMPLOYEE_EXPENSES LS INNER JOIN EMPLOYEE_STATEMENT_DETAILS R0 ON R0.LOAD_CODE = LS.CODE " & " WHERE LS.CREW_CODE IN (|" & tmpItem & "|) AND LS.ENTRY_DATE >= " & gVars.gDI & gBF.IntlDateFormat(SD) & gVars.gDI & " AND LS.ENTRY_DATE < " & gVars.gDI & gBF.IntlDateFormat(ED) & gVars.gDI & " AND R0.PAY_TYPE IN (" & gConst.LOAD_PAY_MISC & ")"

                        'xRS = Nothing
                        xRS.Close1() 'DP added so we close this before using it again

                        xRS = gDataLayer.LoadRecordset(strSQL, gobjErrors)
                        If (gobjErrors.Count = 0) Then
                        Else
                            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
                            Exit Function
                        End If

                        xMisc = xMisc + IIf(IsDBNull(xRS("PAYOLA").Value), 0, CDbl(xRS("PAYOLA").Value))

                        xPay = xPay + xMisc

                        gBF.FloatCellEnter(aForm.aGrid(Page), 10, lngRow, xPay, 2, True, PrintRawData)
                        If (System.Math.Abs(VolCut) > 0) Then
                            gBF.FloatCellEnter(aForm.aGrid(Page), 11, lngRow, xPay / VolCut, 2, True, PrintRawData)
                        Else
                            gBF.FloatCellEnter(aForm.aGrid(Page), 11, lngRow, 0, 2, True, PrintRawData)
                        End If

                        cstLogging = xPay

                        '-------------------------------------------------------------------
                        '     Profit
                        '-------------------------------------------------------------------

                        pi = rev - cstLogging

                        gBF.FloatCellEnter(aForm.aGrid(Page), 12, lngRow, pi, 2, True, PrintRawData)
                        If (System.Math.Abs(VolCut) > 0) Then
                            gBF.FloatCellEnter(aForm.aGrid(Page), 13, lngRow, pi / VolCut, 2, True, PrintRawData)
                        Else
                            gBF.FloatCellEnter(aForm.aGrid(Page), 13, lngRow, 0, 2, True, PrintRawData)
                        End If

                        If (System.Math.Abs(rev) > 0) Then
                            gBF.FloatCellEnter(aForm.aGrid(Page), 14, lngRow, 100 * pi / rev, 2, False, PrintRawData)
                        Else
                            gBF.FloatCellEnter(aForm.aGrid(Page), 14, lngRow, 0, 2, False, PrintRawData)
                        End If

                        blTotals(5) = blTotals(5) + LoadCount
                        blTotals(6) = blTotals(6) + VolCut
                        blTotals(8) = blTotals(8) + rev
                        blTotals(10) = blTotals(10) + cstLogging

                        'xRS = Nothing
                        xRS.Close1() 'DP added so we close this since we are done with it, before moving onto the next Do While step.

                        'Else

                        '                FloatCellEnter aForm.aGrid(Page), 5, lngRow, 0, 0, False, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 6, lngRow, 0, Vdec, False, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 8, lngRow, 0, 2, True, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 9, lngRow, 0, 2, True, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 10, lngRow, 0, 2, True, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 11, lngRow, 0, 2, True, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 12, lngRow, 0, 2, True, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 13, lngRow, 0, 2, True, PrintRawData
                        '                FloatCellEnter aForm.aGrid(Page), 14, lngRow, 0, 2, False, PrintRawData
                    Else
                        'DP added because in case where LoadCount = 0 then it will open another xrs and another and another
                        'xRS = Nothing
                        xRS.Close1()
                    End If



                Else
                    'No Data, go to next period
                End If

                'xRS.MoveNext
                'Loop

                aRS.MoveNext()


            Loop

            lngRow = lngRow + 1
            aForm.aGrid(Page).MaxRows = lngRow

            aForm.aGrid(Page).SetText(2, lngRow, "Crew Totals")

            gBF.FloatCellEnter(aForm.aGrid(Page), 5, lngRow, blTotals(5), 0, False, PrintRawData)
            gBF.FloatCellEnter(aForm.aGrid(Page), 6, lngRow, blTotals(6), Vdec, False, PrintRawData)
            'If volDiff = 1 Then
            'volCode = "TEST2" 'Update the volume code because there are multiple of them
            'Else
            'The volCode should still be set to the only one used on the report and should be printed.
            'End If
            aForm.aGrid(Page).SetText(7, lngRow, CStr(volCode))

            gBF.FloatCellEnter(aForm.aGrid(Page), 8, lngRow, blTotals(8), 2, True, PrintRawData)
            If (System.Math.Abs(blTotals(6)) > 0) Then
                gBF.FloatCellEnter(aForm.aGrid(Page), 9, lngRow, blTotals(8) / blTotals(6), 2, True, PrintRawData)
            Else
                gBF.FloatCellEnter(aForm.aGrid(Page), 9, lngRow, 0, 2, True, PrintRawData)
            End If

            gBF.FloatCellEnter(aForm.aGrid(Page), 10, lngRow, blTotals(10), 2, True, PrintRawData)
            If (System.Math.Abs(blTotals(6)) > 0) Then
                gBF.FloatCellEnter(aForm.aGrid(Page), 11, lngRow, blTotals(10) / blTotals(6), 2, True, PrintRawData)
            Else
                gBF.FloatCellEnter(aForm.aGrid(Page), 11, lngRow, 0, 2, True, PrintRawData)
            End If

            gBF.FloatCellEnter(aForm.aGrid(Page), 12, lngRow, blTotals(8) - blTotals(10), 2, True, PrintRawData)
            If (System.Math.Abs(blTotals(6)) > 0) Then
                gBF.FloatCellEnter(aForm.aGrid(Page), 13, lngRow, (blTotals(8) - blTotals(10)) / blTotals(6), 2, True, PrintRawData)
            Else
                gBF.FloatCellEnter(aForm.aGrid(Page), 13, lngRow, 0, 2, True, PrintRawData)
            End If

            If (System.Math.Abs(blTotals(8)) > 0) Then
                gBF.FloatCellEnter(aForm.aGrid(Page), 14, lngRow, 100 * (blTotals(8) - blTotals(10)) / blTotals(8), 2, False, PrintRawData)
            Else
                gBF.FloatCellEnter(aForm.aGrid(Page), 14, lngRow, 0, 2, False, PrintRawData)
            End If


            aForm.aGrid(Page).Row = lngRow
            aForm.aGrid(Page).Row2 = lngRow
            aForm.aGrid(Page).Col = 1
            aForm.aGrid(Page).Col2 = n
            aForm.aGrid(Page).BlockMode = True
            aForm.aGrid(Page).FontBold = True
            aForm.aGrid(Page).BlockMode = False

            ReDim blTotals(n)

NextCrew:
            aRS.MoveFirst() 'Move to the top of the pay period list

        Next i

        'aRS = Nothing
        aRS.Close1()
        'xRS = Nothing
        xRS.Close1()


        For c = 1 To aForm.aGrid(Page).MaxCols
            If (c <> 2) Then
                aForm.aGrid(Page).colWidth(c) = gBF.Max(aForm.aGrid(Page).MaxTextColWidth(c) + 2, 8)
            End If
        Next c

        '--------------------------------------------------------------------
        '   Print Footer
        '--------------------------------------------------------------------
        If (PrintRawData = 0) Then
            lngRow = lngRow + 2
            aForm.aGrid(Page).MaxRows = lngRow
            aForm.aGrid(Page).SetText(1, lngRow, "Date Range: " & xsd & " to " & IIf(System.Math.Abs(xED.ToOADate() - CLng(xED.ToOADate)) < EPS, System.DateTime.FromOADate(xED.ToOADate - 1), xED))
            aForm.aGrid(Page).Col = 1
            aForm.aGrid(Page).Row = lngRow
            aForm.aGrid(Page).AllowCellOverflow = True
            lngRow = lngRow + 1
            aForm.aGrid(Page).MaxRows = lngRow
            aForm.aGrid(Page).SetText(1, lngRow, FormatDateTime(Now, DateFormat.ShortDate) & " " & FormatDateTime(Now, DateFormat.ShortTime))
            aForm.aGrid(Page).Col = 1
            aForm.aGrid(Page).Row = lngRow
            aForm.aGrid(Page).AllowCellOverflow = True
        End If

        aForm.aGrid(Page).BlockMode = True
        aForm.aGrid(Page).Col = 1
        aForm.aGrid(Page).Col2 = aForm.aGrid(Page).MaxCols
        aForm.aGrid(Page).Row = 1
        aForm.aGrid(Page).Row2 = aForm.aGrid(Page).MaxRows
        aForm.aGrid(Page).Lock = True
        aForm.aGrid(Page).BlockMode = False

        DoCruise = True

        gBF.TurnMeOff()

        Exit Function


Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function



    Function LoadBlockPayActWDepl(ByVal BL As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSQL As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (Not basLocals.gVars.gDebugMode) Then On Error GoTo Err_Handler

        LoadBlockPayActWDepl = New Collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSQL = "SELECT BPA.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM BLOCK_PAY_ACTIVITIES BPA INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE WHERE CREW_CODE = |" & BL & "|" & " AND (IS_PAY <> 0 OR IS_PAY_DRIVER <> 0 OR IS_PAY_OWN_DRIVER <> 0) AND PAY_ACTIVITY_CODE = |DEPLETION| "

        aRS = gDatalayer.LoadRecordset(strSQL, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadBlockPayActWDepl.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        'aRS = Nothing
        aRS.Close1()


        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSQL = "SELECT BPA.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM BLOCK_PAY_ACTIVITIES BPA INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE WHERE CREW_CODE = |" & BL & "|" & " AND (IS_PAY <> 0 OR IS_PAY_DRIVER <> 0 OR IS_PAY_OWN_DRIVER <> 0) AND PAY_ACTIVITY_CODE <> |DEPLETION| ORDER BY PAY_ACTIVITY_CODE "

        aRS = gDatalayer.LoadRecordset(strSQL, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            basLocals.gBF.DisplayErrors(basLocals.gobjErrors, basLocals.gVars.gApp_Title, 0, String.Empty, Nothing)
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadBlockPayActWDepl.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        'aRS = Nothing
        aRS.Close1()

        Exit Function

Err_Handler:
        gBF.TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = gBF.getSource(New StackTrace(eException, True))
        gBF.DisplayErrorsNet(desc, src)

    End Function
End Class
