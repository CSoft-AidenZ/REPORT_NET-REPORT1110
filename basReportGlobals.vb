Option Strict Off
Option Explicit On
Module basReportGlobals
	Public Declare Function LoadCursorFromFile Lib "user32.dll"  Alias "LoadCursorFromFileA"(ByVal lpFileName As String) As Integer
	Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Integer) As Integer
	Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Integer)
	
	Public Declare Function GetPrivateProfileString Lib "Kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Public Declare Function WritePrivateProfileString Lib "Kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	
	Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	Declare Function RegQueryValueExString Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	
	Declare Function RegQueryValueExNULL Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As Integer, ByRef lpcbData As Integer) As Integer
	
	Declare Function RegQueryValueExLong Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer
	
	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_USERS As Integer = &H80000003
	Public Const KEY_QUERY_VALUE As Integer = &H1
    Public Const ERROR_NONE As Integer = 0
    Public Const REG_SZ As Integer = 1
    Public Const REG_DWORD As Integer = 4

    Public Const NA As Integer = -9999
    Public Const VOL_EQUALS_SCALE_VOLUME As Integer = -8888
    Public Const EPS As Double = 0.00001

    Public Const DB_ACCESS As Integer = 0
    Public Const DB_MSSQL As Integer = 1
    Public Const DB_ORACLE As Integer = 2
    Public Const DB_POSTGRESQL As Integer = 3

    Public Const MAX_CPARTS As Integer = 13
    Public Const MAX_PRIME_PARTS As Integer = 3
    Public Const MAX_LUDFS As Integer = 10

    Public Const LOAD_IMPORT_FORMAT As Integer = 1
    Public Const LOAD_REVENUE_IMPORT_FORMAT As Integer = 2
    Public Const LOAD_ACTIVITIES_IMPORT_FORMAT As Integer = 3

    Public Const REV_CONTRACT_NON_REVENUE As Integer = 0
    Public Const REV_CONTRACT_FIXED_FEE As Integer = 6
    Public Const REV_CONTRACT_TIME As Integer = 2
    Public Const REV_CONTRACT_LOADS As Integer = 1
    Public Const REV_CONTRACT_CUSTOMER_LOADS As Integer = 3
    Public Const REV_CONTRACT_PROD As Integer = 4 'Pay by Production
    Public Const REV_CONTRACT_MISC As Integer = 5

    Public Const PP_ALL As Integer = -1
    Public Const PP_EMPLOYEES As Integer = 1
    Public Const PP_VENDORS As Integer = 2
    Public Const PP_INVOICES As Integer = 3
    Public Const PP_REPORTING As Integer = 4
    Public Const PP_CUSTOMER As Integer = 5

    Public Const LOAD_PAY_1 As Integer = 1 'By Load Slip
    Public Const LOAD_PAY_2 As Integer = 2 'Pay By Hour
    Public Const LOAD_PAY_3 As Integer = 3 'By PCT of load Revenue
    Public Const LOAD_PAY_4 As Integer = 4 'Pay by Production
    Public Const LOAD_PAY_MISC As Integer = 5
    Public Const LOAD_PAY_TREE As Integer = 6 'Pay by Trees or Tree Volume
    Public Const LOAD_PAY_5 As Integer = 7 'Pay by Load Value
    Public Const LOAD_PAY_30 As Integer = 8 'By PCT of load Revenue |NEW TO HANDLE TRANSITION

    Public Const I_PAY As Integer = 1
    Public Const I_REVENUE As Integer = 2

    'Used for the Destination C Part for Conversions
    Public Const Destination As Integer = 1000
    Public Const TRUCK_TYPE As Integer = 1001
    Public Const ROUTE As Integer = 1002
    Public Const BLOCK_CODE As Integer = 1003
    Public Const EQUIPMENT_CODE As Integer = 1004

    Public Const IS_PAY_OWN_DRIVER As Integer = 3
    Public Const IS_PAY_CONTRACTOR_DRIVER As Integer = 2
    Public Const IS_PAY_CONTRACTOR As Integer = 1

    Public Const NONE As Integer = 0
    Public Const HRS As Integer = 1
    Public Const XDAY As Integer = 2
    Public Const RATE_SHEET As Integer = 3
    Public Const PRODUCTION As Integer = 4
    Public Const LOAD_REVENUE As Integer = 5
    Public Const LUMP_SUM As Integer = 6
    Public Const TREES As Integer = 7 'Tree Count on Time Slips
    Public Const TREE_VOL As Integer = 8 'Converted Volume from Tree Count
    Public Const DISTANCE As Integer = 9
    Public Const DIST_WGHT As Integer = 10

    Public Enum ConversionType
        NO_CONVERSIONS = 1
        GENERIC_CONVERSIONS = 2
        BLOCK_SPECIFIC_CONVERSIONS = 0
    End Enum

    Public Enum eDatatypes
        dtNumeric = 1
        dtDate = 2
        dtboolean = 3
        dttext = 4
        dtMemo = 5
    End Enum

    '----------------------------------------------------------
    '   Constants for the Grid
    '----------------------------------------------------------
    Public Const TypeHAlignRight As Integer = 1
    Public Const CellTypeFloat As Integer = 2
    Public Const CellTypeCurrency As Integer = 12
    Public Const CellTypePercent As Integer = 14
    Public Const TypeHAlignCenter As Integer = 2
    Public Const CellTypeCheckBox As Integer = 10
    Public Const CellBorderStyleSolid As Integer = 1
    Public Const ActionSetCellBorder As Integer = 16
    Public Const TypeHAlignLeft As Integer = 0
    Public Const BackColorStyleUnderGrid As Integer = 1
	
	
    'Public gVars As CSoft_DataAccess.clsgVars
    'Public gobjTables As New CSoft_DataAccess.clsTables
    'Public gobjErrors As New CSoft_DataAccess.clsErrors
    Public gObjtables As clsLib_DataAccess.clsTables
    Public gObjErrors As clsLib_DataAccess.clsErrors
    Public gDataLayer As clsLib_DataAccess.clsDataLayer
    Public gVars As clsLib_DataAccess.clsVars
    Public gDebugMode As Integer

    Public Function doPageBreak(ByVal aGrid As Object, ByVal Frow As Integer, ByVal Lrow As Integer, ByVal Orientation As Integer, ByVal minRows As Integer) As Integer

        Dim rwPage As Integer
        Dim iPages As Integer
        Dim remainingRows As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If (Frow > Lrow) Then
            Exit Function
        End If

        If (Orientation = 1) Then
            rwPage = Val("" & gVars.gBasicSetup.Value("LINES_PER_PAGE_LANDSCAPE"))
        Else
            rwPage = Val("" & gVars.gBasicSetup.Value("LINES_PER_PAGE_PORTRAIT"))
        End If

        iPages = (Frow - 1) \ rwPage

        'remaining rows
        remainingRows = rwPage - (Frow - (iPages * rwPage))

        'rows need for section
        'If (Lrow - Frow) > remainingRows Then
        If remainingRows >= minRows Then

            ''Stay on same page if at least 10 rows remaining

        Else

            'Do a Page Break
            aGrid.Row = Lrow
            aGrid.RowPageBreak = True
            doPageBreak = True
            aGrid.Col = aGrid.MaxCols
            aGrid.BackColor = System.Drawing.Color.Cyan

        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function PercentCellEnter(ByVal aGrid As Object, ByRef c As Integer, ByRef r As Integer, ByRef val_Renamed As Double, ByRef dec As Integer, ByRef dollars As Boolean, ByRef doRawdata As Integer) As Integer

        Dim xVal As Double
        Dim xFact As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aGrid.Row = r
        aGrid.Row2 = r
        aGrid.Col = c
        aGrid.Col2 = c
        aGrid.BlockMode = True

        'aGrid.CellType = CellTypeFloat
        aGrid.CellType = CellTypePercent

        If (doRawdata <> 0) Then

            aGrid.TypePercentDecPlaces = dec
            aGrid.TypeHAlign = TypeHAlignRight
            'aGrid.TypeFloatMoney = False
            'aGrid.TypeFloatSeparator = False

        Else

            aGrid.TypePercentDecPlaces = dec
            aGrid.TypeHAlign = TypeHAlignRight
            'aGrid.TypeFloatMoney = dollars
            'aGrid.TypeFloatSeparator = True

        End If

        aGrid.BlockMode = False

        'ADDED TO MAKE ROUNDING WORK
        '   IE -- VAL = 3.995 ----THIS WOULD PREVIOUSLY ROUNDED TO 3.99
        '       ID WILL NOW BE 3.995001 AND WILL ROUND TO 4.00

        'val = val + 0.000000001
        '---------------------------------------------------------------

        xFact = Pow10(10.0#, dec)

        xVal = BRound(val_Renamed, xFact)

        'aGrid.SetText c, r, CStr(val)
        aGrid.SetFloat(c, r, xVal / 100.0#)

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function CurrencyCellEnter(ByVal aGrid As Object, ByRef c As Integer, ByRef r As Integer, ByRef val_Renamed As Double, ByRef dec As Integer, ByRef dollars As Boolean, ByRef negativeIndicator As Integer) As Integer

        Dim xVal As Double
        Dim xFact As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If val_Renamed <= 0 Or val_Renamed >= 0 Then
            'nothing
        Else
            val_Renamed = 0
        End If

        aGrid.Row = r
        aGrid.Row2 = r
        aGrid.Col = c
        aGrid.Col2 = c
        aGrid.BlockMode = True

        aGrid.CellType = CellTypeCurrency

        aGrid.TypeFloatDecimalPlaces = dec
        aGrid.TypeHAlign = TypeHAlignRight
        aGrid.TypeCurrencyShowSymbol = dollars
        aGrid.TypeCurrencyShowSep = True
        aGrid.TypeCurrencyNegStyle = negativeIndicator

        aGrid.BlockMode = False

        'ADDED TO MAKE ROUNDING WORK
        '   IE -- VAL = 3.995 ----THIS WOULD PREVIOUSLY ROUNDED TO 3.99
        '       ID WILL NOW BE 3.995001 AND WILL ROUND TO 4.00

        'val = val + 0.000000001
        '---------------------------------------------------------------

        xFact = Pow10(10.0#, dec)

        xVal = BRound(val_Renamed, xFact)

        'aGrid.SetText c, r, CStr(val)
        aGrid.SetFloat(c, r, xVal)

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function FloatCellEnter(ByVal aGrid As Object, ByRef c As Integer, ByRef r As Integer, ByRef val_Renamed As Double, ByRef dec As Integer, ByRef dollars As Boolean, ByRef doRawdata As Integer) As Integer

        Dim xVal As Double
        Dim xFact As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If val_Renamed <= 0 Or val_Renamed >= 0 Then
            'nothing
        Else
            val_Renamed = 0
        End If

        aGrid.Row = r
        aGrid.Row2 = r
        aGrid.Col = c
        aGrid.Col2 = c
        aGrid.BlockMode = True

        aGrid.CellType = CellTypeFloat

        If (doRawdata <> 0) Then

            aGrid.TypeFloatDecimalPlaces = dec
            aGrid.TypeHAlign = TypeHAlignRight
            aGrid.TypeFloatMoney = False
            aGrid.TypeFloatSeparator = False

        Else

            aGrid.TypeFloatDecimalPlaces = dec
            aGrid.TypeHAlign = TypeHAlignRight
            aGrid.TypeFloatMoney = dollars
            aGrid.TypeFloatSeparator = True

        End If

        aGrid.BlockMode = False

        'ADDED TO MAKE ROUNDING WORK
        '   IE -- VAL = 3.995 ----THIS WOULD PREVIOUSLY ROUNDED TO 3.99
        '       ID WILL NOW BE 3.995001 AND WILL ROUND TO 4.00

        'val = val + 0.000000001
        '---------------------------------------------------------------

        xFact = Pow10(10.0#, dec)

        xVal = BRound(val_Renamed, xFact)

        'aGrid.SetText c, r, CStr(val)
        aGrid.SetFloat(c, r, xVal)

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function Pow10(ByRef aVal As Double, ByRef dec As Integer) As Object

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        Pow10 = System.Math.Exp(dec * System.Math.Log(aVal))

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function BRound(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double

        '  For smaller numbers:
        '  BRound = CLng(X * Factor) / Factor

        Dim Temp, FixTemp As Double

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        Temp = (x * Factor)
        FixTemp = Fix(Temp + 0.5 * System.Math.Sign(x) + 0.00000000001)

        ' Handle rounding of .5 in a special manner

        If System.Math.Abs(Temp - Int(Temp) - 0.5) < 0.00000000001 Then
            If FixTemp / 2 <> Int(FixTemp / 2) Then ' Is Temp odd
                ' Reduce Magnitude by 1 to make even
                FixTemp = FixTemp - System.Math.Sign(x)
            End If
        End If

        BRound = FixTemp / Factor

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Sub GetMaxMinTimes(ByRef aLoad As CSOFT_RECORDSET_EXT.clsRecordsetExt, ByRef MinTripTime As Double, ByRef MaxTripTime As Double)

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim JoinStr As String
        Dim OrderByStr As String
        Dim whereStr As String
        Dim i As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        JoinStr = " BLOCK_DISTANCES T0 "
        OrderByStr = "ORDER BY T0.BLOCK_CODE, "
        whereStr = " WHERE T0.BLOCK_CODE = |" & aLoad("BLOCK_CODE").Value & "| AND T0.DESTINATION_CODE = |" & aLoad("DESTINATION_CODE").Value & "|" & " AND T0.EFFECTIVE_DATE <=  " & gVars.gDI & IntlDateFormat(aLoad("DATE_OUT").Value) & gVars.gDI

        For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS

            If (gobjTables("BLOCK_DISTANCES").fields("C_PART_" & i & "_CODE").DisplayOrder > 0) Then
                JoinStr = "(" & JoinStr & " INNER JOIN C_PART_" & i & " AS T" & i & " ON T" & i & ".CODE = T0.C_PART_" & i & "_CODE ) "
                whereStr = whereStr & " AND (C_PART_" & i & "_CODE = |" & aLoad("C_PART_" & i & "_CODE").Value & "| OR C_PART_" & i & "_CODE = |NONE| )"
                OrderByStr = OrderByStr & " T" & i & ".IS_NONE ASC, "
            End If

        Next i

        If (gobjTables("BLOCK_DISTANCES").fields("TRUCK_TYPE_CODE").DisplayOrder > 0) Then
            JoinStr = "(" & JoinStr & " INNER JOIN TRUCK_TYPES AS TT ON TT.CODE = T0.TRUCK_TYPE_CODE ) "
            whereStr = whereStr & " AND (TRUCK_TYPE_CODE = |" & aLoad("TRUCK_TYPE_CODE").Value & "| OR TRUCK_TYPE_CODE = |NONE| )"
            OrderByStr = OrderByStr & " TT.IS_NONE ASC, "
        End If

        If (gobjTables("BLOCK_DISTANCES").fields("ROUTE_CODE").DisplayOrder > 0) Then
            JoinStr = "(" & JoinStr & " INNER JOIN ROUTES AS RO ON RO.CODE = T0.ROUTE_CODE ) "
            whereStr = whereStr & " AND (ROUTE_CODE = |" & aLoad("ROUTE_CODE").Value & "| OR ROUTE_CODE = |NONE| )"
            OrderByStr = OrderByStr & " RO.IS_NONE, "
        End If

        'OrderByStr = VBA.Left(OrderByStr, Len(OrderByStr) - 2)
        OrderByStr = OrderByStr & " T0.EFFECTIVE_DATE DESC "
        strSql = "SELECT MIN_TIME, MAX_TIME FROM " & JoinStr & whereStr & OrderByStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Sub
        End If

        If (aRS.EOF) Then
            MinTripTime = 0
            MaxTripTime = 0
        Else
            MinTripTime = aRS(0).Value
            MaxTripTime = aRS(1).Value
        End If

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub


    Public Function GetTruckHours(ByVal Truck As String, ByVal Block As String, ByVal SD As Date, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim TT As Double
        Dim selStr As String
        Dim i As Integer
        Dim errorStr As String = ""

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        selStr = " BLOCK_CODE, DESTINATION_CODE, DATE_OUT "

        For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS

            If (gObjtables("BLOCK_DISTANCES").Fields("C_PART_" & i & "_CODE").DisplayOrder > 0) Then
                selStr = selStr & ", C_PART_" & i & "_CODE "
            End If

        Next i

        If (gObjtables("BLOCK_DISTANCES").Fields("TRUCK_TYPE_CODE").DisplayOrder > 0) Then
            selStr = selStr & ", TRUCK_TYPE_CODE "
        End If

        If (gObjtables("BLOCK_DISTANCES").Fields("ROUTE_CODE").DisplayOrder > 0) Then
            selStr = selStr & ", ROUTE_CODE "
        End If


        strSql = " SELECT COUNT(LOAD_COUNT), " & selStr & " FROM LOADSLIPS LS " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LS.TRUCK_CODE = |" & Truck & "| AND LS.BLOCK_CODE = |" & Block & "|" & " AND LS.IS_REVERSAL = 0 AND LS.IS_REVERSED = 0 " & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) " & " GROUP BY " & selStr

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Return 0
            Exit Function
        End If

        GetTruckHours = 0
        Do While aRS.EOF = False
            TT = GetTripTime(aRS, errorStr)
            GetTruckHours = GetTruckHours + TT * aRS(0).Value
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function is_Equipment_Company_Owned(ByVal eqCode As String) As Integer

        '0=Not Company Owned
        '1=Is Company Owned

        Dim strSql As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        is_Equipment_Company_Owned = 0

        'Check to see if employee is an Employee
        strSql = " SELECT OWNER FROM EQUIPMENT EQ INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE " & " WHERE EQ.CODE = |" & eqCode & "|"
        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If gobjErrors.Count = 0 Then
        Else
            DisplayErrors()
            Exit Function
        End If

        '0=Not COMPANY OWNED
        '1=Is COMPANY OWNED
        If (aRS.EOF) Then
            is_Equipment_Company_Owned = 0
        Else
            is_Equipment_Company_Owned = aRS(0).Value
        End If

        Exit Function

        TurnMeOff()

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function is_Truck_Company_Owned(ByVal trCode As String) As Integer

        '0=Not Company Owned
        '1=Is Company Owned

        Dim strSql As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        is_Truck_Company_Owned = 0

        'Check to see if employee is an Employee
        strSql = " SELECT OWNER FROM TRUCKS EQ INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE " & " WHERE EQ.CODE = |" & trCode & "|"
        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If gobjErrors.Count = 0 Then
        Else
            DisplayErrors()
            Exit Function
        End If

        '0=Not COMPANY OWNED
        '1=Is COMPANY OWNED
        If (aRS.EOF) Then
            is_Truck_Company_Owned = 0
        Else
            is_Truck_Company_Owned = aRS(0).Value
        End If

        Exit Function

        TurnMeOff()

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivityGroups_NonCosted() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivityGroups_NonCosted = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT DISTINCT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM PAY_ACTIVITY_GROUPS PG INNER JOIN PAY_ACTIVITIES PA ON PA.PAY_ACTIVITY_GROUP_CODE = PG.CODE " & " WHERE PA.IS_NOT_COSTING = 0 " & " ORDER BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            If (StrComp(aRS("PAY_ACT_GROUP").Value, "NONE", CompareMethod.Text) <> 0) Then

                anItem = New clsLib_DataAccess.clsCPart
                anItem.PartCode = "" & aRS("PAY_ACT_GROUP").Value
                anItem.PartNo = aRS("IS_GROUP").Value
                LoadPayActivityGroups_NonCosted.Add(anItem, CStr(aRS("PAY_ACT_GROUP").Value))

            End If
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function
Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivityGroups() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivityGroups = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT DISTINCT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM PAY_ACTIVITY_GROUPS PG INNER JOIN PAY_ACTIVITIES PA ON PA.PAY_ACTIVITY_GROUP_CODE = PG.CODE " & " ORDER BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            If (StrComp(aRS("PAY_ACT_GROUP").Value, "NONE", CompareMethod.Text) <> 0) Then

                anItem = New clsLib_DataAccess.clsCPart
                anItem.PartCode = "" & aRS("PAY_ACT_GROUP").Value
                anItem.PartNo = aRS("IS_GROUP").Value
                LoadPayActivityGroups.Add(anItem, CStr(aRS("PAY_ACT_GROUP").Value))

            End If
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function
Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadPayActivities_NonCosted() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivities_NonCosted = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT CODE FROM PAY_ACTIVITIES WHERE CODE <> |NONE| AND IS_NOT_COSTING = 0 ORDER BY CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            LoadPayActivities_NonCosted.Add(CStr(aRS("CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivities() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivities = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT CODE FROM PAY_ACTIVITIES WHERE CODE <> |NONE| ORDER BY CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            LoadPayActivities.Add(CStr(aRS("CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivitiesByGroup(ByVal aGrp As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivitiesByGroup = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT CODE FROM PAY_ACTIVITIES WHERE PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| " & " OR ( CODE = |" & aGrp & "| AND PAY_ACTIVITY_GROUP_CODE = |NONE| ) " & " AND CODE <> |NONE| ORDER BY CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            LoadPayActivitiesByGroup.Add(CStr(aRS("CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivitiesByGroup_NonCosted(ByVal aGrp As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivitiesByGroup_NonCosted = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT CODE FROM PAY_ACTIVITIES WHERE PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| " & " OR ( CODE = |" & aGrp & "| AND PAY_ACTIVITY_GROUP_CODE = |NONE| ) " & " AND IS_NOT_COSTING = 0 AND CODE <> |NONE| ORDER BY CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            LoadPayActivitiesByGroup_NonCosted.Add(CStr(aRS("CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadPayActivityGroups_NOT_on_Block(ByRef actColl As Collection, ByVal BL As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim aColl As collection

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        aColl = New collection

        '    If (actColl.Count = 0) Then
        '        actStr = ""
        '    Else
        '        actStr = " AND PAY_ACTIVITY_CODE NOT IN ("
        '        For i = 1 To actColl.Count
        '            actStr = actStr & "|" & actColl(i).PartCode & "|, "
        '        Next i
        '        actStr = VBA.Left(actStr, Len(actStr) - 2) & ")"
        '    End If

        actStr = ""

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        strSql = " (SELECT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM " & " (((( BLOCK_COST_DETAIL T0 INNER JOIN EQUIPMENT T1 ON T1.CODE = T0.EQUIPMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1)  ) "

        strSql = strSql & " UNION " & " (SELECT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM " & " (((( BLOCK_COST_DETAIL T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) ) "

        strSql = strSql & " UNION " & " (SELECT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM " & " (((( LOAD_TRUCK_COST T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) ) "

        strSql = strSql & " UNION " & " (SELECT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM " & " ((((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_CONTRACTORS T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " GROUP BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) ) "

        strSql = strSql & " UNION " & " (SELECT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP FROM " & " ((((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_EQUIPMENT T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " GROUP BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION), IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) ) "

        strSql = strSql & " ORDER BY PAY_ACT_GROUP "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Return New Collection
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            If (isObjKeyinVBCollection(aColl, CStr(aRS("PAY_ACT_GROUP").Value))) Or (isObjKeyinVBCollection(actColl, CStr(aRS("PAY_ACT_GROUP").Value))) Then
                'Do Nada
            Else
                anItem.PartCode = "" & aRS("PAY_ACT_GROUP").Value
                anItem.PartDescription = "" & aRS("PAY_ACT_GROUP_DESCRIPTION").Value
                anItem.PartNo = aRS("IS_GROUP").Value
                aColl.Add(anItem, CStr(aRS("PAY_ACT_GROUP").Value))
            End If
            aRS.MoveNext()
        Loop

        LoadPayActivityGroups_NOT_on_Block = aColl

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivityGroups_NOT_on_Block_SP(ByRef actColl As Collection, ByVal BL As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim aColl As collection
        Dim onBlockColl As collection
        Dim offBlockColl As collection
        Dim aGrp As String
        Dim aDescription As String = ""
        Dim isGroup As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        onBlockColl = LoadBlockPayActivities(BL)
        offBlockColl = LoadPayActivities_NOT_on_Block(onBlockColl, BL, SD, ED)

        aColl = New collection

        For i = 1 To offBlockColl.Count

            aGrp = GetActivityGroup(offBlockColl(i).PartCode, aDescription, isGroup)

            anItem = New clsLib_DataAccess.clsCPart
            If (isObjKeyinVBCollection(aColl, CStr(aGrp))) Then
                'Do Nada
            Else
                anItem.PartCode = "" & aGrp
                anItem.PartDescription = "" & aDescription
                anItem.PartNo = isGroup
                aColl.Add(anItem, CStr(aGrp))
            End If
        Next i

        LoadPayActivityGroups_NOT_on_Block_SP = aColl


        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadPayActivityGroups_NOT_on_Block_SP_NonCosted(ByRef actColl As Collection, ByVal BL As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim aColl As collection
        Dim onBlockColl As collection
        Dim offBlockColl As collection
        Dim aGrp As String
        Dim aDescription As String = ""
        Dim isGroup As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        onBlockColl = LoadBlockPayActivities_NonCosted(BL)
        offBlockColl = LoadPayActivities_NOT_on_Block_NonCosted(onBlockColl, BL, SD, ED)

        aColl = New collection

        For i = 1 To offBlockColl.Count

            aGrp = GetActivityGroup(offBlockColl(i).PartCode, aDescription, isGroup)

            anItem = New clsLib_DataAccess.clsCPart
            If (isObjKeyinVBCollection(aColl, CStr(aGrp))) Then
                'Do Nada
            Else
                anItem.PartCode = "" & aGrp
                anItem.PartDescription = "" & aDescription
                anItem.PartNo = isGroup
                aColl.Add(anItem, CStr(aGrp))
            End If
        Next i

        LoadPayActivityGroups_NOT_on_Block_SP_NonCosted = aColl


        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivities_NOT_on_BlockByGroup_NonCosted(ByRef actColl As Collection, ByVal BL As String, ByVal aGrp As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim actStrT As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivities_NOT_on_BlockByGroup_NonCosted = New collection

        If (actColl.Count = 0) Then
            actStr = ""
        Else
            actStr = " AND PAY_ACTIVITY_CODE NOT IN ("
            For i = 1 To actColl.Count
                actStr = actStr & "|" & actColl(i).PartCode & "|, "
            Next i
            actStr = Left(actStr, Len(actStr) - 2) & ")"
        End If

        actStrT = Replace(actStr, "PAY_ACTIVITY_CODE", "ACTIVITY_CODE", 1, -1, CompareMethod.Text)

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        strSql = " (SELECT T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION  FROM " & " (((( BLOCK_COST_DETAIL T0 INNER JOIN EQUIPMENT T1 ON T1.CODE = T0.EQUIPMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND PA.IS_NOT_COSTING = 0 AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM " & " (((( BLOCK_COST_DETAIL T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND PA.IS_NOT_COSTING = 0 AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, PA.DESCRIPTION FROM " & " (((( LOAD_TRUCK_COST T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND PA.IS_NOT_COSTING = 0 AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStrT & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, PA.DESCRIPTION FROM " & " ((((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_CONTRACTORS T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND PA.IS_NOT_COSTING = 0 AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, PA.DESCRIPTION FROM " & " ((((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_EQUIPMENT T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND PA.IS_NOT_COSTING = 0 AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.ACTIVITY_CODE, PA.DESCRIPTION  ) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadPayActivities_NOT_on_BlockByGroup_NonCosted.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadPayActivities_NOT_on_BlockByGroup(ByRef actColl As Collection, ByVal BL As String, ByVal aGrp As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim actStrT As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivities_NOT_on_BlockByGroup = New collection

        If (actColl.Count = 0) Then
            actStr = ""
        Else
            actStr = " AND PAY_ACTIVITY_CODE NOT IN ("
            For i = 1 To actColl.Count
                actStr = actStr & "|" & actColl(i).PartCode & "|, "
            Next i
            actStr = Left(actStr, Len(actStr) - 2) & ")"
        End If

        actStrT = Replace(actStr, "PAY_ACTIVITY_CODE", "ACTIVITY_CODE", 1, -1, CompareMethod.Text)

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        strSql = " (SELECT T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION  FROM " & " (((( BLOCK_COST_DETAIL T0 INNER JOIN EQUIPMENT T1 ON T1.CODE = T0.EQUIPMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM " & " (((( BLOCK_COST_DETAIL T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.PAY_ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, PA.DESCRIPTION FROM " & " (((( LOAD_TRUCK_COST T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStrT & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, PA.DESCRIPTION FROM " & " ((((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_CONTRACTORS T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.ACTIVITY_CODE, PA.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, PA.DESCRIPTION FROM " & " ((((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_EQUIPMENT T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " GROUP BY T0.ACTIVITY_CODE, PA.DESCRIPTION  ) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadPayActivities_NOT_on_BlockByGroup.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivities_NOT_on_Block(ByRef actColl As Collection, ByVal BL As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim actStrT As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivities_NOT_on_Block = New collection

        If (actColl.Count = 0) Then
            actStr = ""
        Else
            actStr = " AND PAY_ACTIVITY_CODE NOT IN ("
            For i = 1 To actColl.Count
                actStr = actStr & "|" & actColl(i).PartCode & "|, "
            Next i
            actStr = Left(actStr, Len(actStr) - 2) & ")"
        End If

        actStrT = Replace(actStr, "PAY_ACTIVITY_CODE", "ACTIVITY_CODE", 1, -1, CompareMethod.Text)

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        strSql = " (SELECT T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION  FROM " & " ((( BLOCK_COST_DETAIL T0 INNER JOIN EQUIPMENT T1 ON T1.CODE = T0.EQUIPMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION FROM " & " ((( BLOCK_COST_DETAIL T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, T2.DESCRIPTION FROM " & " ((( LOAD_TRUCK_COST T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStrT & " GROUP BY T0.ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, T2.DESCRIPTION FROM " & " (((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_CONTRACTORS T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.ACTIVITY_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " GROUP BY T0.ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, T2.DESCRIPTION FROM " & " (((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_EQUIPMENT T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.ACTIVITY_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " GROUP BY T0.ACTIVITY_CODE, T2.DESCRIPTION  ) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadPayActivities_NOT_on_Block.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivities_NOT_on_Block_NonCosted(ByRef actColl As Collection, ByVal BL As String, ByRef SD As Date, ByRef ED As Date) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart
        Dim actStr As String
        Dim actStrT As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivities_NOT_on_Block_NonCosted = New collection

        If (actColl.Count = 0) Then
            actStr = ""
        Else
            actStr = " AND PAY_ACTIVITY_CODE NOT IN ("
            For i = 1 To actColl.Count
                actStr = actStr & "|" & actColl(i).PartCode & "|, "
            Next i
            actStr = Left(actStr, Len(actStr) - 2) & ")"
        End If

        actStrT = Replace(actStr, "PAY_ACTIVITY_CODE", "ACTIVITY_CODE", 1, -1, CompareMethod.Text)

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt

        strSql = " (SELECT T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION  FROM " & " ((( BLOCK_COST_DETAIL T0 INNER JOIN EQUIPMENT T1 ON T1.CODE = T0.EQUIPMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T2.IS_NOT_COSTING = 0 AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION FROM " & " ((( BLOCK_COST_DETAIL T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.PAY_ACTIVITY_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T2.IS_NOT_COSTING = 0 AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStr & " GROUP BY T0.PAY_ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, T2.DESCRIPTION FROM " & " ((( LOAD_TRUCK_COST T0 INNER JOIN TRUCKS T1 ON T1.CODE = T0.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.ACTIVITY_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = T1.OWNER_CODE )" & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T2.IS_NOT_COSTING = 0 AND T1.CODE <> |NONE| AND T0.BLOCK_CODE = |" & BL & "| " & actStrT & " GROUP BY T0.ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, T2.DESCRIPTION FROM " & " (((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_CONTRACTORS T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.ACTIVITY_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND T2.IS_NOT_COSTING = 0 AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " GROUP BY T0.ACTIVITY_CODE, T2.DESCRIPTION  ) "

        strSql = strSql & " UNION " & " (SELECT T0.ACTIVITY_CODE, T2.DESCRIPTION FROM " & " (((( VENDOR_STATEMENT_DETAILS T0 INNER JOIN TIME_DETAILS_EQUIPMENT T1 ON T1.CODE = T0.LOAD_CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN TRUCKS TR ON TR.CODE = T1.TRUCK_CODE) " & " INNER JOIN PAY_ACTIVITIES T2 ON T2.CODE = T0.ACTIVITY_CODE) " & " WHERE T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_TYPE IN (" & LOAD_PAY_2 & "," & LOAD_PAY_TREE & ") AND T1.BLOCK_CODE = |" & BL & "| " & actStr & " AND T2.IS_NOT_COSTING = 0 AND ( EQ.CODE <> |NONE| OR TR.CODE <> |NONE| ) " & " GROUP BY T0.ACTIVITY_CODE, T2.DESCRIPTION  ) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadPayActivities_NOT_on_Block_NonCosted.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetActivityGroup(ByVal Activity As String, ByRef aDescription As Object, ByRef isGroup As Integer) As String

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim xCostField As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        strSql = "SELECT PA.PAY_ACTIVITY_GROUP_CODE, PA.DESCRIPTION AS ACTIVITY_DESC, PAG.DESCRIPTION AS GROUP_DESC FROM PAY_ACTIVITIES PA " & " INNER JOIN PAY_ACTIVITY_GROUPS PAG ON PA.PAY_ACTIVITY_GROUP_CODE = PAG.CODE " & " WHERE PA.CODE = |" & Activity & "|"

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Return ""
            Exit Function
        End If

        If (StrComp(aRS(0).Value, "NONE", CompareMethod.Text) = 0) Then
            GetActivityGroup = Activity
            aDescription = aRS("ACTIVITY_DESC").Value
            isGroup = 0
        Else
            GetActivityGroup = aRS("PAY_ACTIVITY_GROUP_CODE").Value
            aDescription = aRS("GROUP_DESC").Value
            isGroup = 1
        End If

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadBlockPayActivityGroups_NonCosted(ByVal BL As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadBlockPayActivityGroups_NonCosted = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT DISTINCT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP " & " FROM (( BLOCK_PAY_ACTIVITIES BPA " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE ) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE ) " & " WHERE PA.IS_NOT_COSTING = 0 AND BPA.BLOCK_CODE = |" & BL & "| " & " AND (IS_PAY_LIKE_TRUCKING = 0 OR (IS_PAY_LIKE_TRUCKING <> 0 AND IS_TRUCKING <> 0)) " & " ORDER BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACT_GROUP").Value
            anItem.PartDescription = "" & aRS("PAY_ACT_GROUP_DESCRIPTION").Value
            anItem.PartNo = aRS("IS_GROUP").Value
            LoadBlockPayActivityGroups_NonCosted.Add(anItem, CStr(aRS("PAY_ACT_GROUP").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadBlockPayActivityGroups(ByVal BL As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadBlockPayActivityGroups = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT DISTINCT IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) AS PAY_ACT_GROUP, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.DESCRIPTION, PG.DESCRIPTION) AS PAY_ACT_GROUP_DESCRIPTION, IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, 0, 1) AS IS_GROUP " & " FROM (( BLOCK_PAY_ACTIVITIES BPA " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE ) " & " INNER JOIN PAY_ACTIVITY_GROUPS PG ON PG.CODE = PA.PAY_ACTIVITY_GROUP_CODE ) " & " WHERE BPA.BLOCK_CODE = |" & BL & "| " & " AND (IS_PAY_LIKE_TRUCKING = 0 OR (IS_PAY_LIKE_TRUCKING <> 0 AND IS_TRUCKING <> 0)) " & " ORDER BY IIF(PA.PAY_ACTIVITY_GROUP_CODE=|NONE|, PA.CODE, PG.CODE) "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACT_GROUP").Value
            anItem.PartDescription = "" & aRS("PAY_ACT_GROUP_DESCRIPTION").Value
            anItem.PartNo = aRS("IS_GROUP").Value
            LoadBlockPayActivityGroups.Add(anItem, CStr(aRS("PAY_ACT_GROUP").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadBlockPayActivitiesByGroup_NonCosted(ByVal BL As String, ByVal aGrp As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadBlockPayActivitiesByGroup_NonCosted = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT BPA.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM BLOCK_PAY_ACTIVITIES BPA INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE " & " WHERE BLOCK_CODE = |" & BL & "| AND PA.IS_NOT_COSTING = 0 " & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " AND (IS_PAY_LIKE_TRUCKING = 0 OR (IS_PAY_LIKE_TRUCKING <> 0 AND IS_TRUCKING <> 0)) " & " ORDER BY PAY_ACTIVITY_CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadBlockPayActivitiesByGroup_NonCosted.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadBlockPayActivitiesByGroup(ByVal BL As String, ByVal aGrp As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadBlockPayActivitiesByGroup = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT BPA.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM BLOCK_PAY_ACTIVITIES BPA INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE " & " WHERE BLOCK_CODE = |" & BL & "| " & " AND (PA.PAY_ACTIVITY_GROUP_CODE = |" & aGrp & "| OR (PA.PAY_ACTIVITY_GROUP_CODE = |NONE| AND PA.CODE = |" & aGrp & "| ) ) " & " AND (IS_PAY_LIKE_TRUCKING = 0 OR (IS_PAY_LIKE_TRUCKING <> 0 AND IS_TRUCKING <> 0)) " & " ORDER BY PAY_ACTIVITY_CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadBlockPayActivitiesByGroup.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadBlockPayActivities_NonCosted(ByVal BL As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadBlockPayActivities_NonCosted = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT BPA.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM BLOCK_PAY_ACTIVITIES BPA INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE " & " WHERE BLOCK_CODE = |" & BL & "| AND PA.IS_NOT_COSTING = 0 ORDER BY PAY_ACTIVITY_CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadBlockPayActivities_NonCosted.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function LoadBlockPayActivities(ByVal BL As String) As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String
        Dim anItem As clsLib_DataAccess.clsCPart

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadBlockPayActivities = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT BPA.PAY_ACTIVITY_CODE, PA.DESCRIPTION FROM BLOCK_PAY_ACTIVITIES BPA INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = BPA.PAY_ACTIVITY_CODE WHERE BLOCK_CODE = |" & BL & "| ORDER BY PAY_ACTIVITY_CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            anItem = New clsLib_DataAccess.clsCPart
            anItem.PartCode = "" & aRS("PAY_ACTIVITY_CODE").Value
            anItem.PartDescription = "" & aRS("DESCRIPTION").Value
            LoadBlockPayActivities.Add(anItem, CStr(aRS("PAY_ACTIVITY_CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function GetWeightConversion(ByRef FromUOM As String, ByRef ToUOM As String, ByRef errorStr As String) As Double

        Dim strSql As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strWhere As String
        Dim isNoConvert As Integer

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        strSql = "SELECT CONVERSION_FACTOR FROM CONVERSIONS WHERE BLOCK_CODE = |ALL|" & " AND FROM_UOM_CODE = |" & FromUOM & "| AND TO_UOM_CODE = |" & ToUOM & "|"

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If gObjErrors.Count > 0 Then
            DisplayErrors()
            Exit Function
        End If

        If aRS.EOF Then

            'No Factor
            GetWeightConversion = 0
            Exit Function

        Else

            GetWeightConversion = aRS("CONVERSION_FACTOR").Value

        End If

        'Check for conversion in wght to volume
        If (Val(gVars.gBasicSetup.Value("CONVERSION_WGHT_TO_VOL")) = 1) Then
            If (Val(CStr(GetWeightConversion)) > EPS) Then
                GetWeightConversion = 1.0# / GetWeightConversion
            End If
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivityTypeDescriptions() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivityTypeDescriptions = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT DESCRIPTION FROM PAY_ACTIVITY_TYPES WHERE CODE <> |NONE| ORDER BY CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            LoadPayActivityTypeDescriptions.Add(CStr(aRS("DESCRIPTION").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function LoadPayActivityTypes() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim i As Integer
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        LoadPayActivityTypes = New collection

        aRS = New CSOFT_RECORDSET_EXT.clsRecordsetExt
        strSql = "SELECT CODE FROM PAY_ACTIVITY_TYPES WHERE CODE <> |NONE| ORDER BY CODE "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            LoadPayActivityTypes.Add(CStr(aRS("CODE").Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetBudgetRate(ByVal aBlock As String, ByVal Activity As String, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim whereStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        whereStr = " AND BR.ACTIVITY_CODE = |" & Activity & "|"
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("BLOCK_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND ( BR.BLOCK_CODE = |" & aBlock & "| OR BR.BLOCK_CODE = |NONE| )"
        End If
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("C_PART_1_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND BL.C_PART_1_CODE = BR.C_PART_1_CODE "
        End If
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("C_PART_2_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND BL.C_PART_2_CODE = BR.C_PART_2_CODE "
        End If
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("C_PART_3_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND BL.C_PART_3_CODE = BR.C_PART_3_CODE "
        End If

        strSql = "SELECT TOP 1 ACTIVITY_RATE FROM BUDGET_ACTIVITY_RATES BR INNER JOIN BLOCKS BL ON BL.CODE = BR.BLOCK_CODE " & " WHERE BR.EFFECTIVE_DATE <= " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI
        strSql = strSql & whereStr & " ORDER BY BL.IS_NONE ASC, BR.EFFECTIVE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If aRS.EOF Then
            GetBudgetRate = 0
        Else
            GetBudgetRate = aRS(0).Value
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetBudgetPrdHr(ByVal aBlock As String, ByVal Activity As String, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim whereStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        whereStr = " AND BR.ACTIVITY_CODE = |" & Activity & "|"
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("BLOCK_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND ( BR.BLOCK_CODE = |" & aBlock & "| OR BR.BLOCK_CODE = |NONE| )"
        End If
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("C_PART_1_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND BL.C_PART_1_CODE = BR.C_PART_1_CODE "
        End If
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("C_PART_2_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND BL.C_PART_2_CODE = BR.C_PART_2_CODE "
        End If
        If (gObjtables("BUDGET_ACTIVITY_RATES").Fields("C_PART_3_CODE").DisplayOrder > 0) Then
            whereStr = whereStr & " AND BL.C_PART_3_CODE = BR.C_PART_3_CODE "
        End If

        strSql = "SELECT TOP 1 PRODUCTION_PER_HOUR FROM BUDGET_ACTIVITY_RATES BR INNER JOIN BLOCKS BL ON BL.CODE = BR.BLOCK_CODE " & " WHERE BR.EFFECTIVE_DATE <= " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI
        strSql = strSql & whereStr & " ORDER BY BL.IS_NONE ASC, BR.EFFECTIVE_DATE DESC "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If aRS.EOF Then
            GetBudgetPrdHr = NA
        Else
            GetBudgetPrdHr = aRS(0).Value
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function GetBlockLoadVolume_Ext(ByVal BlockCode As String, ByVal SD As Date, ByVal ED As Date, ByVal DestCode As String, ByVal ToUOM As String, ByRef aColl As Collection, ByRef retErrorStr As String) As Double

        '-------------------------------
        'Bob Added Errorstr 4/11/11
        '-------------------------------
        'NOTE BlockCode is assumed to not have doQuote already run
        'Doquote is run in this function

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim groupBy As String
        Dim whereStr As String
        Dim i As Integer
        Dim aPart As clsLib_DataAccess.clsCPart
        Dim cPartColl As collection
        Dim myConversion As Double
        Dim errorStr As String
        Dim IS_CONVERT_USING_NET As Integer
        Dim LOAD_UOM_CODE As String
        Dim strSql As String
        Dim aCode As String
        Dim xED As Date

        If (Not gVars.gDebugMode) Then On Error GoTo Err_Handler

        'xED is used for the call to get the conversion
        'Because the Pay Period has added a day to ED
        xED = DateAdd(DateInterval.Day, -1, ED)

        If (aColl.Count > 0) Then
            whereStr = ""
            For i = 1 To aColl.Count
                aPart = aColl(i)
                whereStr = whereStr & " AND C_PART_" & aPart.PartNo & "_CODE = |" & aPart.PartCode & "|"
            Next i
        Else
            whereStr = ""
        End If

        If (Len(DestCode) > 0) Then
            whereStr = whereStr & " AND DESTINATION_CODE = |" & DestCode & "|"
        End If

        'strSql = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |" & BlockCode & "|"

        strSql = "SELECT  MAX(P.USE_C_PART_1_CODE) AS USE_C_PART_1_CODE,  MAX(P.USE_C_PART_2_CODE) AS USE_C_PART_2_CODE,  MAX(P.USE_C_PART_3_CODE) AS USE_C_PART_3_CODE,  MAX(P.USE_C_PART_4_CODE) AS USE_C_PART_4_CODE, " & "  MAX(P.USE_C_PART_5_CODE) AS USE_C_PART_5_CODE,  MAX(P.USE_C_PART_6_CODE) AS USE_C_PART_6_CODE,  MAX(P.USE_C_PART_7_CODE) AS USE_C_PART_7_CODE,  MAX(P.USE_C_PART_8_CODE) AS USE_C_PART_8_CODE, " & "  MAX(P.USE_C_PART_9_CODE) AS USE_C_PART_9_CODE,  MAX(P.USE_C_PART_10_CODE) AS USE_C_PART_10_CODE,  MAX(P.USE_C_PART_11_CODE) AS USE_C_PART_11_CODE,  MAX(P.USE_C_PART_12_CODE) AS USE_C_PART_12_CODE, " & "  MAX(P.USE_C_PART_13_CODE) AS USE_C_PART_13_CODE,  MAX(P.USE_DESTINATION_CODE) AS USE_DESTINATION_CODE  " & "  FROM  CONVERSION_PARTS P WHERE BLOCK_CODE = |" & BlockCode & "| OR BLOCK_CODE = |ALL|"

        xRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        groupBy = ""
        If (Not xRS.EOF) Then
            groupBy = ""
            For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                    groupBy = groupBy & " C_PART_" & i & "_CODE, "
                End If
            Next i
            If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                groupBy = groupBy & " DESTINATION_CODE, "
            End If
            xRS.MoveFirst()
        End If

        If (Len(groupBy) > 0) Then
            groupBy = Left(groupBy, Len(groupBy) - 2) & " "
        End If

        '-------------------------------------------------------------------------------
        ' First Do Scale Volume where UoM = Target UoM (Tons)
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|
        '-------------------------------------------------------------------------------

        strSql = "SELECT SUM(SCALE_VOLUME) AS SCALE_VOLUME FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND ABS(SCALE_VOLUME) > " & EPS & "  AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE = |" & ToUOM & "|"

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetBlockLoadVolume_Ext = 0

        If (aRS.EOF) Then
            GetBlockLoadVolume_Ext = 0
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolume_Ext = 0
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolume_Ext = 0
                Else
                    GetBlockLoadVolume_Ext = aRS(0).Value
                End If
            End If
        End If


        '-------------------------------------------------------------------------------
        ' Next Grab Scale Volume where UoM <> Target UoM (Tons)
        ' AND the Net Pounds = 0 (otherwise use the net pounds to convert to tons)
        '-------------------------------------------------------------------------------
        strSql = "SELECT SUM(SCALE_VOLUME) AS SCALE_VOLUME, SCALE_VOLUME_UOM_CODE " & IIf(Len(groupBy) > 0, "," & groupBy, "") & " FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND ABS(SCALE_VOLUME) > " & EPS & " AND ABS(NET) < " & EPS & " AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE <> |" & ToUOM & "|"

        If (Len(groupBy) > 0) Then
            strSql = strSql & " GROUP BY " & groupBy & ", SCALE_VOLUME_UOM_CODE "
        Else
            strSql = strSql & " GROUP BY SCALE_VOLUME_UOM_CODE "
        End If

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If


        If (aRS.EOF) Then
            'GetBlockLoadVolume_Ext = 0
        Else
            Do While aRS.EOF = False

                'Build C Part Collection for Conversion
                cPartColl = New collection
                If (Not xRS.EOF) Then

                    For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                        If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                            aCode = aRS("C_PART_" & i & "_CODE").Value
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

                    If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                        aCode = aRS("DESTINATION_CODE").Value
                        If (Len(aCode) > 0) Then
                            '-> Got it
                        Else
                            aCode = "NONE"
                        End If
                    Else
                        aCode = "NONE"
                    End If
                    cPartColl.Add(aCode, CStr(Destination))

                End If

                'Try to Convert weight into cruise basis
                errorStr = ""
                LOAD_UOM_CODE = aRS("SCALE_VOLUME_UOM_CODE").Value
                'IS_CONVERT_USING_NET = getAValue("IS_CONVERT_USING_NET", "BLOCKS", "CODE", BlockCode)

                myConversion = GetConversionRate(BlockCode, xED, LOAD_UOM_CODE, ToUOM, cPartColl, NA, NA, "", errorStr)

                If (Len(errorStr) = 0 And System.Math.Abs(myConversion - NA) > EPS) Then
                    GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext + (myConversion * aRS("SCALE_VOLUME").Value)
                Else
                    GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext
                    '---->  No conversion found, make zero and go to the next one
                    'Conversion = 0
                    retErrorStr = retErrorStr & errorStr & vbCrLf
                    'Exit Function
                End If

                aRS.MoveNext()

            Loop

        End If

        '-------------------------------------------------------------------------------
        ' Add Conversion Amounts for loads that have not been scaled
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|
        '-------------------------------------------------------------------------------
        strSql = "SELECT SUM(NET) AS NET, SUM(ADJ_NET) AS ADJ_NET " & IIf(Len(groupBy) > 0, "," & groupBy, "") & " FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE <> |" & ToUOM & "|"

        If (Len(groupBy) > 0) Then
            strSql = strSql & " GROUP BY " & groupBy
        End If

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (aRS.EOF) Then
            GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext
            Exit Function
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext
                'BL 9/25/2017 Exit Function
                GoTo NextRecord
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext
                    'BL 9/25/2017 Exit Function
                    GoTo NextRecord
                End If
            End If
        End If

        Do While aRS.EOF = False

            'Build C Part Collection for Conversion
            cPartColl = New collection
            If (Not xRS.EOF) Then

                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        aCode = aRS("C_PART_" & i & "_CODE").Value
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

                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    aCode = aRS("DESTINATION_CODE").Value
                    If (Len(aCode) > 0) Then
                        '-> Got it
                    Else
                        aCode = "NONE"
                    End If
                Else
                    aCode = "NONE"
                End If
                cPartColl.Add(aCode, CStr(Destination))

            End If

            'Try to Convert weight into cruise basis
            errorStr = ""
            LOAD_UOM_CODE = getAValue("LOAD_UOM_CODE", "BLOCKS", "CODE", BlockCode)
            IS_CONVERT_USING_NET = getAValue("IS_CONVERT_USING_NET", "BLOCKS", "CODE", BlockCode)

            myConversion = GetConversionRate(BlockCode, xED, LOAD_UOM_CODE, ToUOM, cPartColl, NA, NA, "", errorStr)

            If (Len(errorStr) = 0 And System.Math.Abs(myConversion - NA) > EPS) Then
                If (IS_CONVERT_USING_NET <> 0) Then
                    GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext + (myConversion * aRS("ADJ_NET").Value)
                Else
                    GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext + (myConversion * aRS("NET").Value)
                End If
            Else
                GetBlockLoadVolume_Ext = GetBlockLoadVolume_Ext
                '---->  No conversion found, make zero and go to the next one
                'Conversion = 0
                retErrorStr = retErrorStr & errorStr & vbCrLf
                'Exit Function
            End If

NextRecord:
            aRS.MoveNext()

        Loop

        aRS.Close1()

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function



    Function GetBlockLoadVolumeSplit(ByVal BlockCode As String, ByVal SD As Date, ByVal ED As Date, ByVal DestCode As String, ByVal ToUOM As String, ByRef aColl As Collection) As Double

        'NOTE BlockCode is assumed to not have doQuote already run
        'Doquote is run in this function

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim groupBy As String
        Dim whereStr As String
        Dim i As Integer
        Dim aPart As clsLib_DataAccess.clsCPart
        Dim cPartColl As collection
        Dim myConversion As Double
        Dim errorStr As String
        Dim IS_CONVERT_USING_NET As Integer
        Dim LOAD_UOM_CODE As String
        Dim strSql As String
        Dim aCode As String
        Dim xED As Date

        If (Not gVars.gDebugMode) Then On Error GoTo Err_Handler

        'xED is used for the call to get the conversion
        'Because the Pay Period has added a day to ED
        xED = DateAdd(DateInterval.Day, -1, ED)

        If (aColl.Count > 0) Then
            whereStr = ""
            For i = 1 To aColl.Count
                aPart = aColl(i)
                whereStr = whereStr & " AND C_PART_" & aPart.PartNo & "_CODE = |" & aPart.PartCode & "|"
            Next i
        Else
            whereStr = ""
        End If

        If (Len(DestCode) > 0) Then
            whereStr = whereStr & " AND DESTINATION_CODE = |" & DestCode & "|"
        End If

        strSql = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |" & BlockCode & "|"

        xRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        groupBy = ""
        If (Not xRS.EOF) Then
            groupBy = ""
            For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                    groupBy = groupBy & " C_PART_" & i & "_CODE, "
                End If
            Next i
            If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                groupBy = groupBy & " DESTINATION_CODE, "
            End If
            xRS.MoveFirst()
        End If

        If (Len(groupBy) > 0) Then
            groupBy = Left(groupBy, Len(groupBy) - 2) & " "
        End If

        '-------------------------------------------------------------------------------
        ' First Do Scale Amounts
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|

        strSql = "SELECT SUM(SCALE_VOLUME) AS SCALE_VOLUME FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND C_PART_13_CODE NOT IN (|H|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE = |" & ToUOM & "|"

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetBlockLoadVolumeSplit = 0

        If (aRS.EOF) Then
            GetBlockLoadVolumeSplit = 0
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolumeSplit = 0
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolumeSplit = 0
                Else
                    GetBlockLoadVolumeSplit = aRS(0).Value
                End If
            End If
        End If

        '-------------------------------------------------------------------------------
        ' Add Conversion Amounts for loads that have not been scaled
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|
        strSql = "SELECT SUM(NET) AS NET, SUM(ADJ_NET) AS ADJ_NET " & IIf(Len(groupBy) > 0, "," & groupBy, "") & " FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND C_PART_13_CODE NOT IN (|H|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE <> |" & ToUOM & "|"

        If (Len(groupBy) > 0) Then
            strSql = strSql & " GROUP BY " & groupBy
        End If

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (aRS.EOF) Then
            GetBlockLoadVolumeSplit = GetBlockLoadVolumeSplit
            Exit Function
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolumeSplit = GetBlockLoadVolumeSplit
                'BL 9/25/2017 Exit Function
                GoTo NextRecord
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolumeSplit = GetBlockLoadVolumeSplit
                    'BL 9/25/2017 Exit Function
                    GoTo NextRecord
                End If
            End If
        End If

        Do While aRS.EOF = False

            'Build C Part Collection for Conversion
            cPartColl = New collection
            If (Not xRS.EOF) Then

                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        aCode = aRS("C_PART_" & i & "_CODE").Value
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

                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    aCode = aRS("DESTINATION_CODE").Value
                    If (Len(aCode) > 0) Then
                        '-> Got it
                    Else
                        aCode = "NONE"
                    End If
                Else
                    aCode = "NONE"
                End If
                cPartColl.Add(aCode, CStr(Destination))

            End If

            'Try to Convert weight into cruise basis
            errorStr = ""
            LOAD_UOM_CODE = getAValue("LOAD_UOM_CODE", "BLOCKS", "CODE", BlockCode)
            IS_CONVERT_USING_NET = getAValue("IS_CONVERT_USING_NET", "BLOCKS", "CODE", BlockCode)

            myConversion = GetConversionRate(BlockCode, xED, LOAD_UOM_CODE, ToUOM, cPartColl, NA, NA, "", errorStr)

            If (Len(errorStr) = 0 And System.Math.Abs(myConversion - NA) > EPS) Then
                If (IS_CONVERT_USING_NET <> 0) Then
                    GetBlockLoadVolumeSplit = GetBlockLoadVolumeSplit + (myConversion * aRS("ADJ_NET").Value)
                Else
                    GetBlockLoadVolumeSplit = GetBlockLoadVolumeSplit + (myConversion * aRS("NET").Value)
                End If
            Else
                GetBlockLoadVolumeSplit = GetBlockLoadVolumeSplit
                '---->  No conversion found, make zero and go to the next one
                'Conversion = 0
                'ErrorStr = ErrorStr & vbCrLf & TicketDetails(aLoad)
                'Exit Function
            End If

Nextrecord:
            aRS.MoveNext()

        Loop

        aRS.Close1()

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetBlockLoadVolumeHD(ByVal BlockCode As String, ByVal SD As Date, ByVal ED As Date, ByVal DestCode As String, ByVal ToUOM As String, acoll As Collection) As Double

        'NOTE BlockCode is assumed to not have doQuote already run
        'Doquote is run in this function

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim groupBy As String
        Dim whereStr As String
        Dim i As Long
        Dim aPart As clsLib_DataAccess.clsCPart
        Dim cPartColl As Collection
        Dim Conversion As Double
        Dim errorStr As String
        Dim IS_CONVERT_USING_NET As Long
        Dim LOAD_UOM_CODE As String
        Dim strSQL As String
        Dim aCode As String
        Dim BlockLoadVolume As Double
        Dim xED As Date

        If (Not gVars.gDebugMode) Then On Error GoTo Err_Handler

        'xED is used for the call to get the conversion
        'Because the Pay Period has added a day to ED
        xED = DateAdd(DateInterval.Day, -1, ED)

        If (acoll.Count > 0) Then
            whereStr = ""
            For i = 1 To acoll.Count
                aPart = acoll(i)
                If String.Compare(aPart.PartCode, "NONE", True) = 0 Then
                    'Account for the NONE WildCard
                Else
                    whereStr = whereStr & " AND C_PART_" & aPart.PartNo & "_CODE = |" & aPart.PartCode & "|"
                End If
            Next i
        Else
            whereStr = ""
        End If

        If (Len(DestCode) > 0) Then
            whereStr = whereStr & " AND DESTINATION_CODE = |" & DestCode & "|"
        End If

        '    MsgBox "whereStr:" & whereStr

        '-------------------------------------------------------------------------------
        ' First Do Scale Amounts
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|
        '    strSQL = "SELECT SUM(SCALE_VOLUME) AS SCALE_VOLUME FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & _
        '             " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & _
        '             " AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & _
        '             " " & whereStr & " AND SCALE_VOLUME_UOM_CODE = |" & ToUOM & "|"

        '-------------------------------------------------------------------------------
        ' The Where Clause Now Matches on C_PART_13
        '-------------------------------------------------------------------------------
        strSQL = "SELECT SUM(SCALE_VOLUME) AS SCALE_VOLUME FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & _
                 " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & _
                 " " & whereStr & " AND SCALE_VOLUME_UOM_CODE = |" & ToUOM & "|"

        'MsgBox "strSQL:" & strSQL

        aRS = gObjtables.Datalayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        BlockLoadVolume = 0

        If (aRS.EOF) Then
            BlockLoadVolume = 0
        Else
            If (IsDBNull(aRS(0).Value)) Then
                BlockLoadVolume = 0
            Else
                If (aRS(0).Value = 0) Then
                    BlockLoadVolume = 0
                Else
                    BlockLoadVolume = aRS(0).Value
                End If
            End If
        End If


        strSQL = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |" & BlockCode & "|"

        xRS = gObjtables.Datalayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        groupBy = ""
        If (Not xRS.EOF) Then
            groupBy = ""
            For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                    groupBy = groupBy & " C_PART_" & i & "_CODE, "
                End If
            Next i
            If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                groupBy = groupBy & " DESTINATION_CODE, "
            End If
            xRS.MoveFirst()
        End If

        If (Len(groupBy) > 0) Then
            groupBy = Left(groupBy, Len(groupBy) - 2) & " "
        End If


        '-------------------------------------------------------------------------------
        ' If Len(groupBy) = 0, then there are NO Conversion Parts Set Up at the Block Level
        ' So, Revert to the ALL Block for the Conversion Parts
        ' Bob & Teresa 9/5/2017
        '-------------------------------------------------------------------------------
        If (Len(groupBy) = 0) Then

            strSQL = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |ALL|"

            xRS = gObjtables.Datalayer.LoadRecordset(strSQL, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            groupBy = ""
            If (Not xRS.EOF) Then
                groupBy = ""
                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        groupBy = groupBy & " C_PART_" & i & "_CODE, "
                    End If
                Next i
                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    groupBy = groupBy & " DESTINATION_CODE, "
                End If
                xRS.MoveFirst()
            End If

            If (Len(groupBy) > 0) Then
                groupBy = Left(groupBy, Len(groupBy) - 2) & " "
            End If

        End If


        '-------------------------------------------------------------------------------
        ' Add Conversion Amounts for loads that have not been scaled
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|
        '    strSQL = "SELECT SUM(NET) AS NET, SUM(ADJ_NET) AS ADJ_NET " & IIf(Len(groupBy) > 0, "," & groupBy, "") & " FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & _
        '             " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & _
        '             " AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & _
        '             " " & whereStr & " AND SCALE_VOLUME_UOM_CODE <> |" & ToUOM & "|"

        strSQL = "SELECT SUM(NET) AS NET, SUM(ADJ_NET) AS ADJ_NET " & IIf(Len(groupBy) > 0, "," & groupBy, "") & " FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & _
                 " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & _
                 " " & whereStr & " AND SCALE_VOLUME_UOM_CODE <> |" & ToUOM & "|"

        If (Len(groupBy) > 0) Then
            strSQL = strSQL & " GROUP BY " & groupBy
        End If

        aRS = gObjtables.Datalayer.LoadRecordset(strSQL, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (aRS.EOF) Then
            GetBlockLoadVolumeHD = BlockLoadVolume
            Exit Function
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolumeHD = BlockLoadVolume
                'BL 10/2/2017 Exit Function
                GoTo NextRecord
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolumeHD = BlockLoadVolume
                    'BL 10/2/2017 Exit Function
                    GoTo NextRecord
                End If
            End If
        End If

        Do While aRS.EOF = False

            'Build C Part Collection for Conversion
            cPartColl = New Collection
            If (Not xRS.EOF) Then

                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        aCode = aRS("C_PART_" & i & "_CODE").Value
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

                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    aCode = aRS("DESTINATION_CODE").Value
                    If (Len(aCode) > 0) Then
                        '-> Got it
                    Else
                        aCode = "NONE"
                    End If
                Else
                    aCode = "NONE"
                End If
                cPartColl.Add(aCode, CStr(Destination))

                xRS.MoveFirst()

            End If

            'Try to Convert weight into cruise basis
            errorStr = ""
            LOAD_UOM_CODE = getAValue("LOAD_UOM_CODE", "BLOCKS", "CODE", BlockCode)
            IS_CONVERT_USING_NET = getAValue("IS_CONVERT_USING_NET", "BLOCKS", "CODE", BlockCode)

            Conversion = GetConversionRate(BlockCode, xED, LOAD_UOM_CODE, ToUOM, cPartColl, NA, NA, "", errorStr)

            If (Len(errorStr) = 0 And Math.Abs(Conversion - NA) > EPS) Then
                If (IS_CONVERT_USING_NET <> 0) Then
                    BlockLoadVolume = BlockLoadVolume + (Conversion * aRS("ADJ_NET").Value)
                Else
                    BlockLoadVolume = BlockLoadVolume + (Conversion * aRS("NET").Value)
                End If
            Else
                BlockLoadVolume = BlockLoadVolume
                '---->  No conversion found, make zero and go to the next one
                'Conversion = 0
                'ErrorStr = ErrorStr & vbCrLf & TicketDetails(aLoad)
                'Exit Function
            End If

NextRecord:

            aRS.MoveNext()

        Loop

        aRS.Close1()

        GetBlockLoadVolumeHD = BlockLoadVolume

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Function GetBlockLoadVolume(ByVal BlockCode As String, ByVal SD As Date, ByVal ED As Date, ByVal DestCode As String, ByVal ToUOM As String, ByRef aColl As Collection) As Double

        'NOTE BlockCode is assumed to not have doQuote already run
        'Doquote is run in this function

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim xRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim groupBy As String
        Dim whereStr As String
        Dim i As Integer
        Dim aPart As clsLib_DataAccess.clsCPart
        Dim cPartColl As collection
        Dim myConversion As Double
        Dim errorStr As String
        Dim IS_CONVERT_USING_NET As Integer
        Dim LOAD_UOM_CODE As String
        Dim strSql As String
        Dim aCode As String
        Dim xED As Date

        If (Not gVars.gDebugMode) Then On Error GoTo Err_Handler

        'xED is used for the call to get the conversion
        'Because the Pay Period has added a day to ED
        xED = DateAdd(DateInterval.Day, -1, ED)

        If (aColl.Count > 0) Then
            whereStr = ""
            For i = 1 To aColl.Count
                aPart = aColl(i)
                whereStr = whereStr & " AND C_PART_" & aPart.PartNo & "_CODE = |" & aPart.PartCode & "|"
            Next i
        Else
            whereStr = ""
        End If

        If (Len(DestCode) > 0) Then
            whereStr = whereStr & " AND DESTINATION_CODE = |" & DestCode & "|"
        End If

        strSql = "SELECT * FROM CONVERSION_PARTS WHERE BLOCK_CODE = |" & BlockCode & "|"

        xRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        groupBy = ""
        If (Not xRS.EOF) Then
            groupBy = ""
            For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                    groupBy = groupBy & " C_PART_" & i & "_CODE, "
                End If
            Next i
            If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                groupBy = groupBy & " DESTINATION_CODE, "
            End If
            xRS.MoveFirst()
        End If

        If (Len(groupBy) > 0) Then
            groupBy = Left(groupBy, Len(groupBy) - 2) & " "
        End If

        '-------------------------------------------------------------------------------
        ' First Do Scale Amounts
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|

        strSql = "SELECT SUM(SCALE_VOLUME) AS SCALE_VOLUME FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE = |" & ToUOM & "|"

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetBlockLoadVolume = 0

        If (aRS.EOF) Then
            GetBlockLoadVolume = 0
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolume = 0
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolume = 0
                Else
                    GetBlockLoadVolume = aRS(0).Value
                End If
            End If
        End If

        '-------------------------------------------------------------------------------
        ' Add Conversion Amounts for loads that have not been scaled
        '-------------------------------------------------------------------------------
        ' Bob Added 12/5/05 |AND T0.C_PART_13_CODE NOT IN (|D|, |I|, |S|)|
        strSql = "SELECT SUM(NET) AS NET, SUM(ADJ_NET) AS ADJ_NET " & IIf(Len(groupBy) > 0, "," & groupBy, "") & " FROM LOADSLIPS WHERE BLOCK_CODE = |" & BlockCode & "|" & " AND DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND C_PART_13_CODE NOT IN (|D|, |I|, |S|) " & " " & whereStr & " AND SCALE_VOLUME_UOM_CODE <> |" & ToUOM & "|"

        If (Len(groupBy) > 0) Then
            strSql = strSql & " GROUP BY " & groupBy
        End If

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (aRS.EOF) Then
            GetBlockLoadVolume = GetBlockLoadVolume
            Exit Function
        Else
            If (IsDBNull(aRS(0).Value)) Then
                GetBlockLoadVolume = GetBlockLoadVolume
                'BL 9/25/2017 Exit Function
                GoTo NextRecord
            Else
                If (aRS(0).Value = 0) Then
                    GetBlockLoadVolume = GetBlockLoadVolume
                    'BL 9/25/2017 Exit Function
                    GoTo NextRecord
                End If
            End If
        End If

        Do While aRS.EOF = False

            'Build C Part Collection for Conversion
            cPartColl = New collection
            If (Not xRS.EOF) Then

                For i = MAX_PRIME_PARTS + 1 To MAX_CPARTS
                    If (xRS("USE_C_PART_" & i & "_CODE").Value <> 0) Then
                        aCode = aRS("C_PART_" & i & "_CODE").Value
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

                If (xRS("USE_DESTINATION_CODE").Value <> 0) Then
                    aCode = aRS("DESTINATION_CODE").Value
                    If (Len(aCode) > 0) Then
                        '-> Got it
                    Else
                        aCode = "NONE"
                    End If
                Else
                    aCode = "NONE"
                End If
                cPartColl.Add(aCode, CStr(Destination))

            End If

            'Try to Convert weight into cruise basis
            errorStr = ""
            LOAD_UOM_CODE = getAValue("LOAD_UOM_CODE", "BLOCKS", "CODE", BlockCode)
            IS_CONVERT_USING_NET = getAValue("IS_CONVERT_USING_NET", "BLOCKS", "CODE", BlockCode)

            myConversion = GetConversionRate(BlockCode, xED, LOAD_UOM_CODE, ToUOM, cPartColl, NA, NA, "", errorStr)

            If (Len(errorStr) = 0 And System.Math.Abs(myConversion - NA) > EPS) Then
                If (IS_CONVERT_USING_NET <> 0) Then
                    GetBlockLoadVolume = GetBlockLoadVolume + (myConversion * aRS("ADJ_NET").Value)
                Else
                    GetBlockLoadVolume = GetBlockLoadVolume + (myConversion * aRS("NET").Value)
                End If
            Else
                GetBlockLoadVolume = GetBlockLoadVolume
                '---->  No conversion found, make zero and go to the next one
                'Conversion = 0
                'ErrorStr = ErrorStr & vbCrLf & TicketDetails(aLoad)
                'Exit Function
            End If
NextRecord:
            aRS.MoveNext()

        Loop

        aRS.Close1()

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function



    Public Function GetContractorLoadPayFF(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String, ByVal eqVal As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqField As String
        Dim eqTbl As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for Load Tickets, But NOT for those tickets where the
        '    Equipment Owner is my Company ->  In this case it is in the FF Rate
        '-------------------------------------------------------------------------

        'If (StrComp(Activity, gVars.gTruckingCode, vbTextCompare) = 0) Then
        If (isActivityPayLikeTrucking(Activity) <> 0) Then
            eqField = "TRUCK_CODE"
            'ElseIf (StrComp(Activity, gVars.gTruckingCode, vbTextCompare) <> 0) Then
        Else
            eqField = "EQUIPMENT_CODE"
        End If

        If (isActivityPayLikeTrucking(Activity) = 0) Then
            'If (StrComp(Activity, gVars.gTruckingCode, vbTextCompare) <> 0) Then

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE )" & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND PA.IS_NOT_COSTING = 0 AND OW.OWNER = 0 "

            If (Len(eqVal) > 0) Then
                strSql = strSql & " AND T1." & eqField & " = |" & eqVal & "|"
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

        Else

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN TRUCKS EQ ON EQ.CODE = LS.TRUCK_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE ) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") "

            If Val("" & gVars.gBasicSetup.Value("COST_TRUCKS_LIKE_CONTRACTORS")) <> 0 Then
                'Leave out the following condition
            Else
                strSql = strSql & " AND OW.OWNER = 0 "
            End If

            If (Len(eqVal) > 0) Then
                strSql = strSql & " AND LS." & eqField & " = |" & eqVal & "|"
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

        End If

        If (Not aRS.EOF) Then
            GetContractorLoadPayFF = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorLoadPayFF = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function



    Public Function GetContractorLoadDetailPay(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String, ByVal C_PART_4_CODE As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqField As String
        Dim eqTbl As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for Load Tickets, But NOT for those tickets where the
        '    Equipment Owner is my Company ->  In this case it is in the FF Rate
        '-------------------------------------------------------------------------

        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ") " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND C_PART_13_CODE = |D| AND C_PART_4_CODE = |" & C_PART_4_CODE & "|"

        If (Len(Activity) > 0) Then
            strSql = strSql & " AND T1.ACTIVITY_CODE IN (" & Activity & ")"
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorLoadDetailPay = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorLoadDetailPay = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetContractorLoadHeaderPay(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqField As String
        Dim eqTbl As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for Load Tickets, But NOT for those tickets where the
        '    Equipment Owner is my Company ->  In this case it is in the FF Rate
        '-------------------------------------------------------------------------

        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ") " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND C_PART_13_CODE = |H| "

        If (Len(Activity) > 0) Then
            strSql = strSql & " AND T1.ACTIVITY_CODE IN (" & Activity & ")"
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorLoadHeaderPay = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorLoadHeaderPay = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Function GetFullEqCost(ByVal SD As Date, ByVal ED As Date, ByVal EQUIPMENT_CODE As String, ByVal ACTIVITY_CODE As String, ByVal aBlock As String, ByVal doLoadedCost As Integer, ByVal EqType As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqField As String = ""
        Dim eqTbl As String = ""
        Dim eqCost As String = ""
        Dim eqVal As String = ""

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        'If (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, vbTextCompare) <> 0) Then
        If (StrComp(EqType, "TR", CompareMethod.Text) <> 0) Then
            eqField = "EQUIPMENT_CODE"
            eqTbl = "EQUIPMENT"
            eqCost = "EQUIPMENT_COST"
            eqVal = EQUIPMENT_CODE
            'ElseIf (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, vbTextCompare) = 0) Then
        ElseIf (StrComp(EqType, "TR", CompareMethod.Text) = 0) Then
            eqField = "TRUCK_CODE"
            eqTbl = "TRUCKS"
            eqCost = "TRUCK_COST"
            eqVal = EQUIPMENT_CODE
        End If

        If (doLoadedCost) Then

            strSql = " SELECT (SUM(TOTAL_COST) + SUM(EMPLOYEE_LOAD_COST)) AS TOTAL_COST, SUM(" & eqCost & ") AS EQUIPMENT_COST, SUM(EMPLOYEE_COST) AS EMPLOYEE_COST "

        Else

            strSql = " SELECT SUM(FULLY_FOUNDED_COST) AS TOTAL_COST, SUM(" & eqCost & ") AS EQUIPMENT_COST, SUM(EMPLOYEE_COST) AS EMPLOYEE_COST "

        End If

        strSql = strSql & " FROM ( BLOCK_COST_DETAIL T0 INNER JOIN " & eqTbl & " T1 ON T1.CODE = T0." & eqField & ") " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0." & eqField & " <> |NONE| AND T0.BLOCK_CODE IN (|" & aBlock & "|)  AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "|"

        If (Len(EQUIPMENT_CODE) > 0) Then
            strSql = strSql & " AND " & eqField & " = |" & eqVal & "|"
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetFullEqCost = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetFullEqCost = 0
        End If

        'If activity is Trucking and
        'Equipment is a Truck
        If (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, CompareMethod.Text) = 0 And (StrComp(EqType, "TR", CompareMethod.Text) = 0)) Then

            '-----------------------------------------------------------------------
            '   Fully Founded Truck Cost where cost is based on weight NOT Hours
            '-----------------------------------------------------------------------

            If (doLoadedCost) Then

                strSql = "SELECT SUM(T0.COST) AS TRUCK_COST " & " FROM LOAD_TRUCK_COST T0 INNER JOIN LOADSLIPS LS ON T0.LOAD_CODE = LS.CODE " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LS.BLOCK_CODE IN (|" & aBlock & "|) AND T0.ACTIVITY_CODE = |" & ACTIVITY_CODE & "|"

            Else

                strSql = "SELECT SUM(T0.FULLY_FOUNDED_COST) AS TRUCK_COST " & " FROM LOAD_TRUCK_COST T0 INNER JOIN LOADSLIPS LS ON T0.LOAD_CODE = LS.CODE " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LS.BLOCK_CODE IN (|" & aBlock & "|) AND T0.ACTIVITY_CODE = |" & ACTIVITY_CODE & "|"

            End If


            If (Len(EQUIPMENT_CODE) > 0) Then
                strSql = strSql & " AND T0." & eqField & " = |" & eqVal & "|"
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                GetFullEqCost = GetFullEqCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
                GetFullEqCost = GetFullEqCost
            End If

        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployeeLoadCostRentalEQ(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String, ByVal EQUIPMENT_CODE As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        '-------------------------------------------------------------------------------
        '  Returns -  Employee Pay for Loads;
        '  If LEN(Equipment_Code) > "", retunrs cost for a specific piece of equipment
        '-------------------------------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Employees, For Loads FOR Rental Equipment (Contractor Equipment, My Driver) on Loads
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM (((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = T1.EQUIPMENT_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE )" & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND PA.IS_NOT_COSTING = 0 AND OW.OWNER = 0 AND T1.EQUIPMENT_CODE <> |NONE| AND T1.ACTIVITY_CODE = |" & Activity & "|"

        If (Len(EQUIPMENT_CODE) > 0) Then
            strSql = strSql & " AND T1.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "|"
        End If

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeLoadCostRentalEQ = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeLoadCostRentalEQ = 0
        End If

        'If (StrComp(Activity, gVars.gTruckingCode, vbTextCompare) = 0) Then
        If (isActivityPayLikeTrucking(Activity) <> 0) Then
            '-------------------------------------------------------------------------
            '    Pay Employees, For Loads FOR Rental Trucks (Contractor Truck, My Driver) on Loads
            '-------------------------------------------------------------------------
            strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM (((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN TRUCKS EQ ON EQ.CODE = LS.TRUCK_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE )" & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND PA.IS_NOT_COSTING = 0 AND OW.OWNER = 0 AND LS.TRUCK_CODE <> |NONE| AND T1.ACTIVITY_CODE = |" & Activity & "|"

            If (Len(EQUIPMENT_CODE) > 0) Then
                strSql = strSql & " AND LS.TRUCK_CODE = |" & EQUIPMENT_CODE & "|"
            End If

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                GetEmployeeLoadCostRentalEQ = GetEmployeeLoadCostRentalEQ + IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
            Else
                GetEmployeeLoadCostRentalEQ = GetEmployeeLoadCostRentalEQ + 0
            End If

        End If


        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function





    Public Function GetContractorHourlyPay(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal EmCode As String, ByVal eqCode As String, ByVal Activity As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND TD.EMPLOYEE_CODE = |" & EmCode & "| AND TD.EQUIPMENT_CODE = |" & eqCode & "| " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") " & " GROUP BY T1.BLOCK_CODE, T1.ACTIVITY_CODE "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorHourlyPay = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorHourlyPay = 0
        End If

        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND TD.EMPLOYEE_CODE = |" & EmCode & "| AND TD.EQUIPMENT_CODE = |" & eqCode & "| AND T1.PAY_TYPE = " & LOAD_PAY_2 & " " & " GROUP BY T1.BLOCK_CODE, T1.ACTIVITY_CODE "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorHourlyPay = GetContractorHourlyPay + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorHourlyPay = GetContractorHourlyPay
        End If

        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_EQUIPMENT TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND TD.EQUIPMENT_CODE = |" & eqCode & "| AND T1.PAY_TYPE = " & LOAD_PAY_2 & " " & " GROUP BY T1.BLOCK_CODE, T1.ACTIVITY_CODE "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorHourlyPay = GetContractorHourlyPay + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorHourlyPay = GetContractorHourlyPay
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetContractorOthPay(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByRef Activity As String, ByVal skipTrucking As Integer) As Double

        '------------------------------------------------------------------------
        '  Costs Included Here Are:
        '  4.  Pay Contractors, for Prod on Contractor Production Tickets -> Travel - > Truck & Equipment are NONE
        '------------------------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim activityStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Len(Activity) = 0 Then
            activityStr = ""
        Else
            activityStr = " AND TD.PAY_ACTIVITY_CODE = |" & Activity & "|"
        End If

        GetContractorOthPay = 0

        '-------------------------------------------------------------------------
        '    Pay Contractors, for Time Where Equip & Truck are |NONE| - Travel Time
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM ((VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE IN (" & LOAD_PAY_2 & ") " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorOthPay = GetContractorOthPay + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorOthPay = GetContractorOthPay
        End If

        '-------------------------------------------------------------------------
        '    Pay Contractors, for Time Where Equip & Truck are |NONE| - Travel Time
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM ((VENDOR_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorOthPay = GetContractorOthPay + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorOthPay = GetContractorOthPay
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function



    Public Function GetContractorHourlyPay_FF(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String, ByVal EQUIPMENT_CODE As String, ByVal EQUIPMENT_TYPE As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqFieldColl As collection
        Dim eqTblColl As collection
        Dim eqField As String = ""
        Dim eqTbl As String
        Dim i As Integer
        Dim ii As Integer
        Dim jj As Integer
        Dim equipStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        eqTblColl = New collection
        eqTblColl.Add("EQUIPMENT")
        eqTblColl.Add("TRUCKS")

        eqFieldColl = New collection
        eqFieldColl.Add("EQUIPMENT_CODE")
        eqFieldColl.Add("TRUCK_CODE")

        ii = 1
        jj = 2

        If Len(EQUIPMENT_TYPE) > 0 Then
            If (StrComp(EQUIPMENT_TYPE, "TR", CompareMethod.Text) = 0) Then 'Equipment is a Truck
                ii = 2
            Else
                jj = 1
            End If

        End If

        GetContractorHourlyPay_FF = 0

        For i = ii To jj

            eqTbl = eqTblColl(i)
            eqField = eqFieldColl(i)

            If Len(EQUIPMENT_CODE) > 0 Then
                equipStr = "AND EQ.CODE = |" & EQUIPMENT_CODE & "| "
            Else
                equipStr = "AND EQ.CODE <> |NONE| "
            End If

            '-------------------------------------------------------------------------
            '    Pay Contractors, for time on Contractor Time Tickets
            '    Exclude Details where Equipment is Company-Owned. -> Only Include Contractor Equipment
            '-------------------------------------------------------------------------
1:          strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE ) " & " INNER JOIN " & eqTbl & " EQ ON EQ.CODE = TD." & eqField & " ) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE ) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE )" & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") " & equipStr & " AND PA.IS_NOT_COSTING = 0 AND OW.OWNER = 0 "

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                GetContractorHourlyPay_FF = GetContractorHourlyPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
                GetContractorHourlyPay_FF = GetContractorHourlyPay_FF
            End If

            '--------------------------------------------------------------------------------------------------------------------------------
            '    Pull in Payroll for Time on Employee Time Tickets for Rental Equipment - (This payroll cost is not in fully founded rate)
            '-------------------------------------------------------------------------------------------------------------------------------
2:          strSql = "SELECT SUM(TD.EMPLOYEE_COST+TD.EMPLOYEE_LOAD_COST) AS PAY " & " FROM ((( BLOCK_COST_DETAIL TD INNER JOIN " & eqTbl & " EQ ON EQ.CODE = TD." & eqField & ") " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = TD.PAY_ACTIVITY_CODE )" & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE IN (|" & aBlock & "|) AND TD.PAY_ACTIVITY_CODE = |" & Activity & "|" & " AND PA.IS_NOT_COSTING = 0 AND OW.OWNER = 0 " & equipStr

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                GetContractorHourlyPay_FF = GetContractorHourlyPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
                GetContractorHourlyPay_FF = GetContractorHourlyPay_FF
            End If

        Next i

        If Len(EQUIPMENT_CODE) > 0 Then
            equipStr = " AND TD." & eqField & " = |" & EQUIPMENT_CODE & "|"
        Else
            equipStr = ""
        End If

        '------------------------------------------------------------------------------------------
        '    Pay Contractors, for Equipment Time on Employee Time Tickets  - Rental Equipment
        '------------------------------------------------------------------------------------------
4:      strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (( VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE )" & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE = " & LOAD_PAY_2 & equipStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorHourlyPay_FF = GetContractorHourlyPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorHourlyPay_FF = GetContractorHourlyPay_FF
        End If

        '------------------------------------------------------------------------------------------
        '    Pay Contractors, for Equipment Time on Equipment Time Tickets  - Rental Equipment ON Equipment Slips
        '    Equipment Only - No Trucks
        '------------------------------------------------------------------------------------------
5:      strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM ((VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_EQUIPMENT TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE )" & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) AND T1.ACTIVITY_CODE = |" & Activity & "|" & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE = " & LOAD_PAY_2 & equipStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorHourlyPay_FF = GetContractorHourlyPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorHourlyPay_FF = GetContractorHourlyPay_FF
        End If

        GetContractorHourlyPay_FF = GetContractorHourlyPay_FF + GetContractorLoadPayFF(aBlock, SD, ED, Activity, EQUIPMENT_CODE)

        GetContractorHourlyPay_FF = GetContractorHourlyPay_FF + GetEmployeeLoadCostRentalEQ(aBlock, SD, ED, Activity, EQUIPMENT_CODE)

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetContractorProdPay_FF(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String) As Double

        '----------------------------------------------------------
        ' aBlock is either one block or a list.
        ' DoQuote Has ALREADY been applied
        '----------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqFieldColl As collection
        Dim eqTblColl As collection
        Dim eqField As String
        Dim eqTbl As String
        Dim i As Integer
        Dim activityStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Len(Activity) = 0 Then
            activityStr = ""
        Else
            activityStr = " AND T1.ACTIVITY_CODE = |" & Activity & "|"
        End If

        eqTblColl = New collection
        eqTblColl.Add("EQUIPMENT")
        eqTblColl.Add("TRUCKS")

        eqFieldColl = New collection
        eqFieldColl.Add("EQUIPMENT_CODE")
        eqFieldColl.Add("TRUCK_CODE")

        GetContractorProdPay_FF = 0

        For i = 1 To 2

            eqTbl = eqTblColl(i)
            eqField = eqFieldColl(i)

            '--------------------------------------------------------------------------------------------------------------------------------
            '    Pull in Payroll for Time on Employee Production Tickets for Rental Equipment - (This payroll cost is not in fully founded rate)
            '-------------------------------------------------------------------------------------------------------------------------------
3:          strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM (((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN " & eqTbl & " EQ ON EQ.CODE = TD." & eqField & ") " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE IN (|" & aBlock & "|) " & activityStr & " AND OW.OWNER = 0 AND TD." & eqField & " <> |NONE|" & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE = " & LOAD_PAY_4

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                GetContractorProdPay_FF = GetContractorProdPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
                GetContractorProdPay_FF = GetContractorProdPay_FF
            End If

        Next i

        '-----------------------------------------------------------------------------------------------------------
        '    Pay Contractors, for Equipment Prod on Employee Prod Tickets  - Rental Equipment ON Production Slips
        '    Equipment Only - No Trucks
        '-----------------------------------------------------------------------------------------------------------
6:      strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = TD.EQUIPMENT_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (OW.OWNER = 0 AND EQ.CODE <> |NONE|) " & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorProdPay_FF = GetContractorProdPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorProdPay_FF = GetContractorProdPay_FF
        End If


        '--------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets,
        '    but Only if Contractor Equipment or No Equipment
        '    - Equipment
        '-------------------------------------------------------------------------

7:      strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN EQUIPMENT EQ ON EQ.CODE = TD.EQUIPMENT_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (OW.OWNER = 0 AND EQ.CODE <> |NONE|) " & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorProdPay_FF = GetContractorProdPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorProdPay_FF = GetContractorProdPay_FF
        End If

        '--------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets,
        '    but Only if Contractor Equipment or No Equipment
        '    - Trucks
        '-------------------------------------------------------------------------

8:      strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN TRUCKS EQ ON EQ.CODE = TD.TRUCK_CODE) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (|" & aBlock & "|) " & " AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) AND (OW.OWNER = 0 AND EQ.CODE <> |NONE|) " & " AND PA.IS_NOT_COSTING = 0 AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorProdPay_FF = GetContractorProdPay_FF + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorProdPay_FF = GetContractorProdPay_FF
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function isActivityPayLikeTrucking(ByVal Activity As String) As Integer

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    See if Activity is Pay Like Trucking
        '-------------------------------------------------------------------------
        strSql = "SELECT IS_PAY_LIKE_TRUCKING FROM PAY_ACTIVITIES WHERE CODE = |" & Activity & "|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            isActivityPayLikeTrucking = aRS(0).Value
        Else
            isActivityPayLikeTrucking = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetisPayLikeTruckingCollection_NonCosted() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetisPayLikeTruckingCollection_NonCosted = New collection

        '-------------------------------------------------------------------------
        '    See if Activity is Pay Like Trucking
        '-------------------------------------------------------------------------
        strSql = "SELECT CODE FROM PAY_ACTIVITIES WHERE IS_PAY_LIKE_TRUCKING <> 0 AND IS_NOT_COSTING = 0 ORDER BY IS_TRUCKING DESC, CODE "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            GetisPayLikeTruckingCollection_NonCosted.Add(CStr(aRS(0).Value), CStr(aRS(0).Value))
            aRS.MoveNext()
        Loop

        'UPGRADE_NOTE: Object aRS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetisPayLikeTruckingCollection() As Collection

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetisPayLikeTruckingCollection = New collection

        '-------------------------------------------------------------------------
        '    See if Activity is Pay Like Trucking
        '-------------------------------------------------------------------------
        strSql = "SELECT CODE FROM PAY_ACTIVITIES WHERE IS_PAY_LIKE_TRUCKING <> 0 ORDER BY IS_TRUCKING DESC, CODE "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False
            GetisPayLikeTruckingCollection.Add(CStr(aRS(0).Value), CStr(aRS(0).Value))
            aRS.MoveNext()
        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetContractorMisc(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN MISCELLANEOUS_EXPENSES TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorMisc = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorMisc = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetContractorMisc_NonCosted(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (( VENDOR_STATEMENT_DETAILS T1 INNER JOIN MISCELLANEOUS_EXPENSES TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE)  " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") " & " AND PA.IS_NOT_COSTING = 0 "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorMisc_NonCosted = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorMisc_NonCosted = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetEmployeeMisc(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN EMPLOYEE_EXPENSES TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeMisc = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeMisc = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployeeMisc_NonCosted(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM (( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN EMPLOYEE_EXPENSES TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE)  " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") " & " AND PA.IS_NOT_COSTING = 0 "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeMisc_NonCosted = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeMisc_NonCosted = 0
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetMiscCostRS_NonCosted(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date) As CSOFT_RECORDSET_EXT.clsRecordsetExt

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets
        '-------------------------------------------------------------------------
        strSql = "( SELECT |Employee| as MTYPE, SUM(T1.PAY) AS PAY, AC.DESCRIPTION AS DESCRIPTION " & " FROM ((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN EMPLOYEE_EXPENSES TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN ACCOUNT_CODES AC ON AC.CODE = TD.ACCOUNT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE)  " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND PA.IS_NOT_COSTING = 0 " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") GROUP BY AC.DESCRIPTION ) "

        strSql = strSql & " UNION "
        strSql = strSql & "( SELECT |Contractor| as MTYPE, SUM(T1.PAY) AS PAY, AC.DESCRIPTION AS DESCRIPTION " & " FROM ((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN MISCELLANEOUS_EXPENSES TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN ACCOUNT_CODES AC ON AC.CODE = TD.ACCOUNT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE)  " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND PA.IS_NOT_COSTING = 0 " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") GROUP BY AC.DESCRIPTION ) "

        strSql = strSql & " ORDER BY MTYPE, DESCRIPTION "

        GetMiscCostRS_NonCosted = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetMiscCostRS(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date) As CSOFT_RECORDSET_EXT.clsRecordsetExt

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Production Tickets
        '-------------------------------------------------------------------------
        strSql = "( SELECT |Employee| as MTYPE, SUM(T1.PAY) AS PAY, AC.DESCRIPTION AS DESCRIPTION " & " FROM (( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN EMPLOYEE_EXPENSES TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN ACCOUNT_CODES AC ON AC.CODE = TD.ACCOUNT_CODE) " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") GROUP BY AC.DESCRIPTION ) "

        strSql = strSql & " UNION "
        strSql = strSql & "( SELECT |Contractor| as MTYPE, SUM(T1.PAY) AS PAY, AC.DESCRIPTION AS DESCRIPTION " & " FROM (( VENDOR_STATEMENT_DETAILS T1 INNER JOIN MISCELLANEOUS_EXPENSES TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN ACCOUNT_CODES AC ON AC.CODE = TD.ACCOUNT_CODE) " & " WHERE TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_MISC & ") GROUP BY AC.DESCRIPTION ) "

        strSql = strSql & " ORDER BY MTYPE, DESCRIPTION "

        GetMiscCostRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetContractor_Equip_ActVolume_Where_Equipment_Is_NONE(ByVal SD As Date, ByVal ED As Date, ByVal EQUIPMENT_TYPE_CODE As String, ByVal EQUIPMENT_CODE As String, ByVal aBlock As String, ByVal ACTIVITY_CODE As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetContractor_Equip_ActVolume_Where_Equipment_Is_NONE = 0

        If (StrComp(EQUIPMENT_TYPE_CODE, "TR", CompareMethod.Text) = 0) Then

            strSql = " SELECT SUM(LS.VOLUME) FROM LOADSLIPS LS " & " INNER JOIN VENDOR_STATEMENT_DETAILS SD ON SD.LOAD_CODE = LS.CODE " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LS.TRUCK_CODE = |" & EQUIPMENT_CODE & "|" & " AND LS.BLOCK_CODE = |" & aBlock & "|" & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) " & " AND SD.PAY_TYPE IN (" & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_5 & ", " & LOAD_PAY_30 & ") "

        Else

            strSql = " SELECT SUM(LS.VOLUME * LP.LOAD_PCT) FROM (( LOADSLIP_PAY_ACTIVITIES LP " & " INNER JOIN LOADSLIPS LS ON LS.CODE = LP.LOAD_CODE ) " & " INNER JOIN VENDOR_STATEMENT_DETAILS SD ON SD.LOAD_CODE = LS.CODE AND SD.ACTIVITY_CODE = LP.PAY_ACTIVITY_CODE ) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LP.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "|" & " AND LP.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND LS.BLOCK_CODE = |" & aBlock & "|" & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) " & " AND SD.PAY_TYPE IN (" & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_5 & ", " & LOAD_PAY_30 & ") "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetContractor_Equip_ActVolume_Where_Equipment_Is_NONE = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetActVolumeWhenEquipmentisNONE(ByVal SD As Date, ByVal ED As Date, ByVal EQUIPMENT_TYPE_CODE As String, ByVal aBlock As String, ByVal ACTIVITY_CODE As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim EQUIPMENT_CODE As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        EQUIPMENT_CODE = "NONE"

        GetActVolumeWhenEquipmentisNONE = 0

        If (StrComp(EQUIPMENT_TYPE_CODE, "TR", CompareMethod.Text) = 0) Then

            strSql = " SELECT SUM(LS.VOLUME) FROM LOADSLIPS LS " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LS.TRUCK_CODE = |" & EQUIPMENT_CODE & "|" & " AND LS.BLOCK_CODE = |" & aBlock & "|" & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) "
        Else

            strSql = " SELECT SUM(LS.VOLUME * LP.LOAD_PCT) FROM LOADSLIP_PAY_ACTIVITIES LP " & " INNER JOIN LOADSLIPS LS ON LS.CODE = LP.LOAD_CODE  " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LP.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "|" & " AND LP.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND LS.BLOCK_CODE = |" & aBlock & "|" & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetActVolumeWhenEquipmentisNONE = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetEquipmentActVolume(ByVal SD As Date, ByVal ED As Date, ByVal EQUIPMENT_TYPE_CODE As String, ByVal EQUIPMENT_CODE As String, ByVal aBlock As String, ByVal ACTIVITY_CODE As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If (StrComp(EQUIPMENT_CODE, "NONE", CompareMethod.Text) = 0) Then
            Exit Function
        End If

        GetEquipmentActVolume = 0

        If (StrComp(EQUIPMENT_TYPE_CODE, "TR", CompareMethod.Text) = 0) Then

            strSql = " SELECT SUM(LS.VOLUME) FROM LOADSLIPS LS " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LS.TRUCK_CODE = |" & EQUIPMENT_CODE & "|" & " AND LS.BLOCK_CODE = |" & aBlock & "|" & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) "
        Else

            strSql = " SELECT SUM(LS.VOLUME * LP.LOAD_PCT) FROM LOADSLIP_PAY_ACTIVITIES LP " & " INNER JOIN LOADSLIPS LS ON LS.CODE = LP.LOAD_CODE  " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND LP.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "|" & " AND LP.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND LS.BLOCK_CODE = |" & aBlock & "|" & " AND (LS.C_PART_13_CODE = |NONE| OR LS.C_PART_13_CODE = |H| OR LS.C_PART_13_CODE IS NULL) "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetEquipmentActVolume = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEquipmentHours(ByVal SD As Date, ByVal ED As Date, ByVal EQUIPMENT_CODE As String, ByVal aBlock As String, ByVal ACTIVITY_CODE As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If (StrComp(EQUIPMENT_CODE, "NONE", CompareMethod.Text) = 0) Then
            Exit Function
        End If

        GetEquipmentHours = 0

        '-------------------------------------------------------------------------
        '    Hours
        '-------------------------------------------------------------------------
        'If (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, vbTextCompare) <> 0) Then

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM TIME_DETAILS AS T0 " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE = |" & aBlock & "| " & " AND T0.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "| AND T0.EQUIPMENT_CODE <> |NONE|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEquipmentHours = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEquipmentHours = 0
        End If

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM TIME_DETAILS_CONTRACTORS AS T0 " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE = |" & aBlock & "| " & " AND T0.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "| AND T0.EQUIPMENT_CODE <> |NONE|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEquipmentHours = GetEquipmentHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEquipmentHours = GetEquipmentHours
        End If

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM TIME_DETAILS_EQUIPMENT AS T0 " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE = |" & aBlock & "| " & " AND T0.EQUIPMENT_CODE = |" & EQUIPMENT_CODE & "| AND T0.EQUIPMENT_CODE <> |NONE|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEquipmentHours = GetEquipmentHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEquipmentHours = GetEquipmentHours
        End If

        'ElseIf (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, vbTextCompare) = 0) Then

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM TIME_DETAILS AS T0 " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE = |" & aBlock & "| " & " AND T0.TRUCK_CODE = |" & EQUIPMENT_CODE & "| AND T0.TRUCK_CODE <> |NONE|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEquipmentHours = GetEquipmentHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEquipmentHours = 0
        End If

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM TIME_DETAILS_CONTRACTORS AS T0 " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE = |" & aBlock & "| " & " AND T0.TRUCK_CODE = |" & EQUIPMENT_CODE & "| AND T0.TRUCK_CODE <> |NONE|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEquipmentHours = GetEquipmentHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEquipmentHours = GetEquipmentHours
        End If

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM TIME_DETAILS_EQUIPMENT AS T0 " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE = |" & aBlock & "| " & " AND T0.TRUCK_CODE = |" & EQUIPMENT_CODE & "| AND T0.TRUCK_CODE <> |NONE|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEquipmentHours = GetEquipmentHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEquipmentHours = GetEquipmentHours
        End If

        'End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetActivityHours(ByVal SD As Date, ByVal ED As Date, ByVal aBlock As String, ByVal ACTIVITY_CODE As String, ByRef isProductionOnly As Integer) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetActivityHours = 0

        '-------------------------------------------------------------------------
        '    Hours
        '-------------------------------------------------------------------------

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM (( TIME_DETAILS AS T0 INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE ) " & " INNER JOIN PAY_ACTIVITY_TYPES PT ON PT.CODE = PA.PAY_ACTIVITY_TYPE_CODE ) " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE IN (|" & aBlock & "|) "

        If (isProductionOnly) Then
            strSql = strSql & " AND PRODUCTION <> 0 "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetActivityHours = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetActivityHours = 0
        End If

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM (( TIME_DETAILS_CONTRACTORS AS T0 INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE ) " & " INNER JOIN PAY_ACTIVITY_TYPES PT ON PT.CODE = PA.PAY_ACTIVITY_TYPE_CODE ) " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE IN (|" & aBlock & "|) "

        If (isProductionOnly) Then
            strSql = strSql & " AND PRODUCTION <> 0 "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetActivityHours = GetActivityHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetActivityHours = GetActivityHours
        End If

        strSql = " SELECT SUM(T0.UNITS) AS HOURS FROM (( TIME_DETAILS_EQUIPMENT AS T0 INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T0.PAY_ACTIVITY_CODE ) " & " INNER JOIN PAY_ACTIVITY_TYPES PT ON PT.CODE = PA.PAY_ACTIVITY_TYPE_CODE ) " & " WHERE T0.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| AND T0.BLOCK_CODE IN (|" & aBlock & "|) "

        If (isProductionOnly) Then
            strSql = strSql & " AND PRODUCTION <> 0 "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetActivityHours = GetActivityHours + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetActivityHours = GetActivityHours
        End If


        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetActivityCost_Total_ByGroup(ByVal doByDestination As Integer, ByVal aBlock As String, ByVal doByActivity As Object, ByVal doLoadedCost As Integer, _
                                                  Optional ByVal SD As Date = Nothing, Optional ByVal ED As Date = Nothing) As Collection

        '-------------------------------------
        'doByActivity = 1 -> Group By Activity
        'doByActivity = 0 -> Group by Block
        '-------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim activCost As Double
        Dim a_Coll As Collection
        Dim tabl As Object
        Dim cstColl As Collection
        Dim aVal As clsLib_DataAccess.clsAValue
        Dim grpBy0 As String = ""
        Dim selBY0 As String = ""
        Dim fldStr0 As String = ""
        Dim grpBy1 As String = ""
        Dim selBY1 As String = ""
        Dim fldStr1 As String = ""
        Dim grpBy2 As String = ""
        Dim selBY2 As String = ""
        Dim fldStr2 As String = ""
        Dim grpVal As String

        '-----------------------------------------------------------------------
        '       doByDestination is a Flag that is Used for Costs for a Job, where the Job is the Desintation
        '       and it is purchasing loads.  In this case the Block is the Source (Gravel Pit).
        '       So the cost of gravel is the purchases by the Job -> Destination
        '       Destination is now where the load is used
        '       Block is the source of the gravel
        '-----------------------------------------------------------------------

        If (gDebugMode = 0) Then On Error GoTo Err_Handler


        If (doByActivity > 0) Then
            selBY0 = " T1.ACTIVITY_CODE "
            grpBy0 = " GROUP BY T1.ACTIVITY_CODE ORDER BY T1.ACTIVITY_CODE "
            fldStr0 = "ACTIVITY_CODE"
            selBY1 = " T1.PAY_ACTIVITY_CODE "
            grpBy1 = " GROUP BY T1.PAY_ACTIVITY_CODE ORDER BY T1.PAY_ACTIVITY_CODE "
            fldStr1 = "PAY_ACTIVITY_CODE"
            selBY2 = " T1.ACTIVITY_CODE "
            grpBy2 = " GROUP BY T1.ACTIVITY_CODE ORDER BY T1.ACTIVITY_CODE "
            fldStr2 = "ACTIVITY_CODE"
        ElseIf (doByActivity = 0) Then
            selBY0 = " T1.BLOCK_CODE "
            grpBy0 = " GROUP BY T1.BLOCK_CODE ORDER BY T1.BLOCK_CODE"
            fldStr0 = "BLOCK_CODE"
            selBY1 = " T1.BLOCK_CODE "
            grpBy1 = " GROUP BY T1.BLOCK_CODE ORDER BY T1.BLOCK_CODE"
            fldStr1 = "BLOCK_CODE"
            If doByDestination = 0 Then
                selBY2 = " TD.BLOCK_CODE "
                grpBy2 = " GROUP BY TD.BLOCK_CODE ORDER BY TD.BLOCK_CODE"
                fldStr2 = "BLOCK_CODE"
            Else
                selBY2 = " TD.DESTINATION_CODE "
                grpBy2 = " GROUP BY TD.DESTINATION_CODE ORDER BY TD.DESTINATION_CODE"
                fldStr2 = "DESTINATION_CODE"
            End If
        End If


        cstColl = New Collection
        '-----------------------------------------------------------
        '   Vendor Statement Details
        '-----------------------------------------------------------

        a_Coll = New Collection
        a_Coll.Add("TIME_DETAILS_CONTRACTORS")
        a_Coll.Add("TIME_DETAILS")
        a_Coll.Add("TIME_DETAILS_EQUIPMENT")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, " & selBY0 & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If


            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

                If (doByActivity = 2) Then
                    grpVal = "Subcontractor"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("PRODUCTION_DETAILS_CONTRACTORS")
        a_Coll.Add("PRODUCTION_DETAILS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, " & selBY0 & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

                If (doByActivity = 2) Then
                    grpVal = "Subcontractor"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl


        a_Coll = New Collection
        a_Coll.Add("LOADSLIPS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, " & selBY2 & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_5 & ", " & LOAD_PAY_30 & ") "

            If (Len(aBlock) = 0) Then
            Else
                If doByDestination = 0 Then
                    strSql = strSql & " AND TD.BLOCK_CODE IN (" & aBlock & ") "
                Else
                    strSql = strSql & " AND TD.DESTINATION_CODE IN (" & aBlock & ") "
                End If
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.DATE_OUT >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.DATE_OUT < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy2

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

                If (doByActivity = 2) Then
                    grpVal = "Subcontractor"
                Else
                    grpVal = aRS(fldStr2).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("MISCELLANEOUS_EXPENSES")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, " & selBY0 & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_MISC & " ) "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

                If (doByActivity = 2) Then
                    grpVal = "Miscellaneous"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl


        '-----------------------------------------------------------
        '   EMPLOYEE Statement Details
        '-----------------------------------------------------------

        '-----------------------------------------------------------
        '   NO NEED for Hourly Pay Details -
        '   COSTS ARE IN BLOCK_COST_DETAIL BELOW
        '-----------------------------------------------------------

        a_Coll = New Collection
        a_Coll.Add("TIME_DETAILS")
        a_Coll.Add("TIME_DETAILS_EQUIPMENT")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD, " & selBY0 & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDBNull(aRS(1).Value), 0, aRS(1).Value)
                End If

                If (doByActivity = 2) Then
                    grpVal = "Employee"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("PRODUCTION_DETAILS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD, " & selBY0 & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDBNull(aRS(1).Value), 0, aRS(1).Value)
                End If

                If (doByActivity = 2) Then
                    grpVal = "Employee"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("LOADSLIPS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD, " & selBY0 & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_5 & ", " & LOAD_PAY_30 & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.DATE_OUT >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.DATE_OUT < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDBNull(aRS(1).Value), 0, aRS(1).Value)
                End If

                If (doByActivity = 2) Then
                    grpVal = "Employee"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("EMPLOYEE_EXPENSES")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, " & selBY0 & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_MISC & " ) "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

                If (doByActivity = 2) Then
                    grpVal = "Miscellaneous"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl

        '-----------------------------------------------------------
        '   LOAD_TRUCK_COST
        '-----------------------------------------------------------

        a_Coll = New Collection
        a_Coll.Add("LOADSLIPS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.COST) AS PAY, " & selBY0 & " FROM LOAD_TRUCK_COST T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE 1=1 " 'sic - all records

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.DATE_OUT >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.DATE_OUT < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            strSql = strSql & grpBy0

            aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
            If (gObjErrors.Count = 0) Then
            Else
                Return New Collection
                Exit Function
            End If

            Do While aRS.EOF = False

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)

                If (doByActivity = 2) Then
                    grpVal = "Own Trucks"
                Else
                    grpVal = aRS(fldStr0).Value
                End If

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                aRS.MoveNext()
            Loop

        Next tabl


        '-----------------------------------------------------------
        '   BLOCK_COST_DETAIL
        '-----------------------------------------------------------

        strSql = "SELECT SUM(T1.EMPLOYEE_COST) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD, SUM(EQUIPMENT_COST) AS EQUIP_COST, " & selBY1 & " FROM BLOCK_COST_DETAIL T1 " & " WHERE 1=1 " 'sic

        If (Len(aBlock) = 0) Then
        Else
            strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
        End If

        If (Not IsNothing(SD)) Then
            strSql = strSql & " AND T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
        End If

        strSql = strSql & grpBy1

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Return New Collection
            Exit Function
        End If

        Do While aRS.EOF = False

            If (doByActivity = 2) Then

                '-------------------
                ' Employees
                '-------------------
                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDBNull(aRS(1).Value), 0, aRS(1).Value)
                End If
                grpVal = "Employee"

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

                '-------------------
                ' Equipment
                '-------------------
                activCost = IIf(IsDBNull(aRS(2).Value), 0, aRS(2).Value)
                grpVal = "Own Equipment"

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

            Else

                activCost = IIf(IsDBNull(aRS(0).Value), 0, aRS(0).Value) + IIf(IsDBNull(aRS(2).Value), 0, aRS(2).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDBNull(aRS(1).Value), 0, aRS(1).Value)
                End If
                grpVal = aRS(fldStr1).Value

                If isObjKeyinVBCollection(cstColl, grpVal) Then
                    cstColl(grpVal).dblValue = cstColl(grpVal).dblValue + activCost
                Else
                    aVal = New clsLib_DataAccess.clsAValue
                    aVal.dblValue = activCost
                    aVal.strValue = grpVal
                    cstColl.Add(aVal, grpVal)
                End If

            End If

            aRS.MoveNext()
        Loop

        GetActivityCost_Total_ByGroup = cstColl

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetActivityCost_Total(ByVal doByDestination As Integer, ByVal aBlock As String, ByVal ACTIVITY_CODE As String, ByVal doLoadedCost As Integer, Optional ByRef SD As Object = Nothing, Optional ByRef ED As Object = Nothing) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim activCost As Double
        Dim a_Coll As Collection
        Dim tabl As Object


        '-----------------------------------------------------------------------
        '       doByDestination is a Flag that is Used for Costs for a Job, where the Job is the Desintation
        '       and it is purchasing loads.  In this case the Block is the Source (Gravel Pit).
        '       So the cost of gravel is the purchases by the Job -> Destination
        '       Destination is now where the load is used
        '       Block is the source of the gravel
        '-----------------------------------------------------------------------

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-----------------------------------------------------------
        '   Vendor Statement Details
        '-----------------------------------------------------------

        a_Coll = New Collection
        a_Coll.Add("TIME_DETAILS_CONTRACTORS")
        a_Coll.Add("TIME_DETAILS")
        a_Coll.Add("TIME_DETAILS_EQUIPMENT")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
            End If

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("PRODUCTION_DETAILS_CONTRACTORS")
        a_Coll.Add("PRODUCTION_DETAILS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
            End If

        Next tabl


        a_Coll = New Collection
        a_Coll.Add("LOADSLIPS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_5 & ", " & LOAD_PAY_30 & ") "

            If (Len(aBlock) = 0) Then
            Else
                If doByDestination = 0 Then
                    strSql = strSql & " AND TD.BLOCK_CODE IN (" & aBlock & ") "
                Else
                    strSql = strSql & " AND TD.DESTINATION_CODE IN (" & aBlock & ") "
                End If
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.DATE_OUT >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.DATE_OUT < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
            End If

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("MISCELLANEOUS_EXPENSES")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_MISC & " ) "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
            End If

        Next tabl


        '-----------------------------------------------------------
        '   EMPLOYEE Statement Details
        '-----------------------------------------------------------

        '-----------------------------------------------------------
        '   NO NEED for Hourly Pay Details -
        '   COSTS ARE IN BLOCK_COST_DETAIL BELOW
        '-----------------------------------------------------------

        a_Coll = New Collection
        a_Coll.Add("TIME_DETAILS")
        a_Coll.Add("TIME_DETAILS_EQUIPMENT")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD " & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDbNull(aRS(1).Value), 0, aRS(1).Value)
                End If
            Else
            End If

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("PRODUCTION_DETAILS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD " & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDbNull(aRS(1).Value), 0, aRS(1).Value)
                End If
            Else
            End If

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("LOADSLIPS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD " & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_1 & "," & LOAD_PAY_3 & ", " & LOAD_PAY_5 & ", " & LOAD_PAY_30 & ") "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.DATE_OUT >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.DATE_OUT < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
                If (doLoadedCost) Then
                    activCost = activCost + IIf(IsDbNull(aRS(1).Value), 0, aRS(1).Value)
                End If
            Else
            End If

        Next tabl

        a_Coll = New Collection
        a_Coll.Add("EMPLOYEE_EXPENSES")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE T1.PAY_TYPE IN (" & LOAD_PAY_MISC & " ) "

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.ENTRY_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.ENTRY_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
            End If

        Next tabl

        '-----------------------------------------------------------
        '   LOAD_TRUCK_COST
        '-----------------------------------------------------------

        a_Coll = New Collection
        a_Coll.Add("LOADSLIPS")

        For Each tabl In a_Coll

            strSql = "SELECT SUM(T1.COST) AS PAY " & " FROM LOAD_TRUCK_COST T1 INNER JOIN " & tabl & " TD ON T1.LOAD_CODE = TD.CODE " & " WHERE 1=1 " 'sic - all records

            If (Len(aBlock) = 0) Then
            Else
                strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
            End If

            If (Len(ACTIVITY_CODE) = 0) Then
            Else
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
            End If

            If (Not IsNothing(SD)) Then
                strSql = strSql & " AND TD.DATE_OUT >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND TD.DATE_OUT < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
            End If

        Next tabl


        '-----------------------------------------------------------
        '   BLOCK_COST_DETAIL
        '-----------------------------------------------------------

        strSql = "SELECT SUM(T1.TOTAL_COST) AS PAY, SUM(EMPLOYEE_LOAD_COST) AS EMP_LOAD " & " FROM BLOCK_COST_DETAIL T1 " & " WHERE 1=1 " 'sic

        If (Len(aBlock) = 0) Then
        Else
            strSql = strSql & " AND T1.BLOCK_CODE IN (" & aBlock & ") "
        End If

        If (Len(ACTIVITY_CODE) = 0) Then
        Else
            strSql = strSql & " AND T1.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "| "
        End If

        If (Not IsNothing(SD)) Then
            strSql = strSql & " AND T1.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(CDate(SD)) & gVars.gDI & " AND T1.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(CDate(ED)) & gVars.gDI
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            activCost = activCost + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            If (doLoadedCost) Then
                activCost = activCost + IIf(IsDbNull(aRS(1).Value), 0, aRS(1).Value)
            End If
        Else
        End If

        GetActivityCost_Total = activCost

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetContractorCost_Obsolete(ByVal SD As Date, ByVal ED As Date, ByVal EQUIPMENT_CODE As String, ByVal aBlock As String, ByVal ACTIVITY_CODE As String, ByVal EqType As String) As Double

        '------------------------------------------------------------
        '  Costs Included Here are:
        '  Contractor Hourly for the Piece Of Equipment - TIME_DETAILS_CONTRACTORS
        '  Contractor Hourly Rental - TIME_DETAILS_EQUIPMENT
        '  Contractor Hourly Rental - TIME_DETAILS
        '  Employee Loaded Costs - Hourly - These are Employee Drivers of Contractor Equipment
        '  Contractor Costs Based on Load Tickets
        '  Employee Costs on Loads Where the Equipment is Contyractor
        '------------------------------------------------------------

        '------------------------------------------------------------
        '  Costs NOT Included:
        '  Contractor Costs from Production Slips - PRODUCTION_DETAILS_CONTRACTORS
        '  Contractor Costs from Employee Production Slips - PRODUCTION_DETAILS -> Rental Based on Production
        '  Contractor Costs Where Equipmnet/Truck is NONE
        '------------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim eqField As String
        Dim eqVal As String
        Dim eqTbl As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If (StrComp(EQUIPMENT_CODE, "NONE", CompareMethod.Text) = 0) Then
            Exit Function
        End If

        'If (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, vbTextCompare) <> 0) Then
        If (StrComp(EqType, "TR", CompareMethod.Text) <> 0) Then
            eqField = "EQUIPMENT_CODE"
            eqVal = EQUIPMENT_CODE
            eqTbl = "EQUIPMENT"
            'ElseIf (StrComp(ACTIVITY_CODE, gVars.gTruckingCode, vbTextCompare) = 0) Then
        ElseIf (StrComp(EqType, "TR", CompareMethod.Text) = 0) Then
            eqField = "TRUCK_CODE"
            eqVal = EQUIPMENT_CODE
            eqTbl = "TRUCKS"
        Else
            Exit Function
        End If

        '-------------------------------------------------------------------------
        '    Pay Contractors, for time on Contractor Time Tickets
        '    Exclude Details where Equipment is Company-Owned. -> Only Include Contractor Equipment
        '-------------------------------------------------------------------------
1:      strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM ((( VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_CONTRACTORS TD ON T1.LOAD_CODE = TD.CODE ) " & " INNER JOIN " & eqTbl & " EQ ON EQ.CODE = TD." & eqField & " ) " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE ) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_2 & ", " & LOAD_PAY_TREE & ") AND EQ.CODE <> |NONE| " & " AND EQ.CODE = |" & eqVal & "| AND OW.OWNER = 0 "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorCost_Obsolete = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorCost_Obsolete = 0
        End If

        '------------------------------------------------------------------------------------------
        '    Pay Contractors, for Equipment Time on Employee Time Tickets  - Rental Equipment
        '------------------------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "|" & " AND T1.PAY_TYPE = " & LOAD_PAY_2 & " AND TD." & eqField & " = |" & eqVal & "|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorCost_Obsolete = GetContractorCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorCost_Obsolete = GetContractorCost_Obsolete
        End If

        '------------------------------------------------------------------------------------------
        '    Pay Contractors, for Equipment Time on Equipment Time Tickets  - Rental Equipment ON Equipment Slips
        '------------------------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN TIME_DETAILS_EQUIPMENT TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "|" & " AND T1.PAY_TYPE = " & LOAD_PAY_2 & " AND TD." & eqField & " = |" & eqVal & "|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorCost_Obsolete = GetContractorCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorCost_Obsolete = GetContractorCost_Obsolete
        End If

        '------------------------------------------------------------------------------------------
        '    Pay Contractors, for Equipment Prod on Employee Prod Tickets  - Rental Equipment ON Production Slips
        '------------------------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM VENDOR_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| AND T1.ACTIVITY_CODE = |" & ACTIVITY_CODE & "|" & " AND T1.PAY_TYPE = " & LOAD_PAY_4 & " AND TD." & eqField & " = |" & eqVal & "|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorCost_Obsolete = GetContractorCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorCost_Obsolete = GetContractorCost_Obsolete
        End If

        '--------------------------------------------------------------------------------------------------------------------------------
        '    Pull in Payroll for Time on Employee Time Tickets for Rental Equipment - (This payroll cost is not in fully founded rate)
        '-------------------------------------------------------------------------------------------------------------------------------
2:      strSql = "SELECT SUM(TD.EMPLOYEE_COST+TD.EMPLOYEE_LOAD_COST) AS PAY " & " FROM (( BLOCK_COST_DETAIL TD INNER JOIN " & eqTbl & "  EQ ON EQ.CODE = TD." & eqField & ") " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE = |" & aBlock & "| AND TD.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "|" & " AND OW.OWNER = 0 AND TD." & eqField & " = |" & eqVal & "|"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorCost_Obsolete = GetContractorCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorCost_Obsolete = GetContractorCost_Obsolete
        End If

        '--------------------------------------------------------------------------------------------------------------------------------
        '    Pull in Payroll for Time on Employee Production Tickets for Rental Equipment - (This payroll cost is not in fully founded rate)
        '-------------------------------------------------------------------------------------------------------------------------------
3:      strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM ((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN " & eqTbl & " EQ ON EQ.CODE = TD." & eqField & ") " & " INNER JOIN OWNERS OW ON OW.CODE = EQ.OWNER_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE = |" & aBlock & "| AND TD.PAY_ACTIVITY_CODE = |" & ACTIVITY_CODE & "|" & " AND OW.OWNER = 0 AND TD." & eqField & " = |" & eqVal & "|" & " AND T1.PAY_TYPE = " & LOAD_PAY_4

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetContractorCost_Obsolete = GetContractorCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetContractorCost_Obsolete = GetContractorCost_Obsolete
        End If


        GetContractorCost_Obsolete = GetContractorCost_Obsolete + GetContractorLoadPayFF(aBlock, SD, ED, ACTIVITY_CODE, eqVal)

        GetContractorCost_Obsolete = GetContractorCost_Obsolete + GetEmployeeLoadCostRentalEQ(aBlock, SD, ED, ACTIVITY_CODE, eqVal)

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployee_DetLoadBased_Pay_Where_Equipment_Is_NONE(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByRef Activity As String, ByVal doLoaded As Integer, ByVal skipTrucking As Integer) As String

        'Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim strSql2 As String
        Dim T As String
        Dim selStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler
        '------------------------------------------------------------------------
        '  Costs Included Here Are:
        '  Contractor Pay for Loads where the Equipment is NONE
        '------------------------------------------------------------------------

        If (doLoaded) Then
            T = " T1.PAY + T1.EMPLOYEE_LOAD_COST "
        Else
            T = " T1.PAY "
        End If

        selStr = " SELECT " & T & " AS COST, T1.ACTIVITY_CODE AS ACTIVITY_CODE, ST.EMPLOYEE_CODE AS PAYEE, LS.BLOCK_CODE, LS.DATE_OUT AS DATE_OUT, LS.TICKET_NO AS TICKET, T1.EQUIPMENT_CODE AS EQUIPMENT, T1.EMPLOYEE_CODE AS DRIVER_CODE, T1.PAY_WEIGHT, T1.PAY_RATE, T1.MEASURE_CODE, |Load Slips| AS SOURCE, LS.CODE AS REF "

        strSql = selStr & " FROM (((EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN EMPLOYEE_STATEMENTS ST ON ST.CODE = T1.STATEMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ")" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND T1.EQUIPMENT_CODE = |NONE| AND PA.IS_PAY_LIKE_TRUCKING = 0 " & " AND PA.IS_NOT_COSTING = 0 "

        If (Len(Activity) > 0) Then
            strSql = strSql & " AND T1.ACTIVITY_CODE = |" & Activity & "|"
        End If

        If (skipTrucking) Then

            GetEmployee_DetLoadBased_Pay_Where_Equipment_Is_NONE = strSql

        Else

            selStr = " SELECT " & T & " AS COST, T1.ACTIVITY_CODE AS ACTIVITY_CODE, ST.EMPLOYEE_CODE AS PAYEE, LS.BLOCK_CODE, LS.DATE_OUT AS DATE_OUT, LS.TICKET_NO AS TICKET, LS.TRUCK_CODE AS EQUIPMENT, LS.DRIVER_CODE AS DRIVER_CODE, T1.PAY_WEIGHT, T1.PAY_RATE, T1.MEASURE_CODE, |Load Slips| AS SOURCE, LS.CODE AS REF  "

            strSql2 = selStr & " FROM ((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN EMPLOYEE_STATEMENTS ST ON ST.CODE = T1.STATEMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ")" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND LS.TRUCK_CODE = |NONE| AND PA.IS_PAY_LIKE_TRUCKING <> 0 " & " AND PA.IS_NOT_COSTING = 0 "

            If (Len(Activity) > 0) Then
                strSql2 = strSql2 & " AND T1.ACTIVITY_CODE = |" & Activity & "|"
            End If

            GetEmployee_DetLoadBased_Pay_Where_Equipment_Is_NONE = "(" & strSql & ") UNION ALL (" & strSql2 & ")"

        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployee_LoadBased_Pay_Where_Equipment_Is_NONE(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByRef Activity As String, ByVal doLoaded As Integer, ByVal skipTrucking As Integer) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim T As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '------------------------------------------------------------------------
        '  Costs Included Here Are:
        '  Contractor Pay for Loads where the Equipment is NONE
        '------------------------------------------------------------------------
        If (doLoaded) Then
            T = " SUM(T1.PAY) + SUM(T1.EMPLOYEE_LOAD_COST) "
        Else
            T = " SUM(T1.PAY) "
        End If


        strSql = "SELECT " & T & " AS PAY " & " FROM ((EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ")" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND T1.EQUIPMENT_CODE = |NONE| AND PA.IS_PAY_LIKE_TRUCKING = 0 " & " AND PA.IS_NOT_COSTING = 0 "

        If (Len(Activity) > 0) Then
            strSql = strSql & " AND T1.ACTIVITY_CODE = |" & Activity & "|"
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetEmployee_LoadBased_Pay_Where_Equipment_Is_NONE = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        If (skipTrucking) Then
            'Skip
        Else

            strSql = "SELECT " & T & " AS PAY " & " FROM ((EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ")" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND LS.TRUCK_CODE = |NONE| AND PA.IS_PAY_LIKE_TRUCKING <> 0 " & " AND PA.IS_NOT_COSTING = 0 "

            If (Len(Activity) > 0) Then
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & Activity & "|"
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            GetEmployee_LoadBased_Pay_Where_Equipment_Is_NONE = GetEmployee_LoadBased_Pay_Where_Equipment_Is_NONE + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetContractor_LoadBased_Pay_Where_Equipment_Is_NONE(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByRef Activity As String, ByVal skipTrucking As Integer) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '------------------------------------------------------------------------
        '  Costs Included Here Are:
        '  Contractor Pay for Loads where the Equipment is NONE
        '------------------------------------------------------------------------

        strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM ((VENDOR_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND PA.IS_NOT_COSTING = 0 AND T1.EQUIPMENT_CODE = |NONE| AND PA.IS_PAY_LIKE_TRUCKING = 0 "

        If (Len(Activity) > 0) Then
            strSql = strSql & " AND T1.ACTIVITY_CODE = |" & Activity & "|"
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        GetContractor_LoadBased_Pay_Where_Equipment_Is_NONE = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        If (skipTrucking) Then

            strSql = "SELECT SUM(T1.PAY) AS PAY " & " FROM ((VENDOR_STATEMENT_DETAILS T1 INNER JOIN LOADSLIPS LS ON T1.LOAD_CODE = LS.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE LS.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND LS.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "| " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_1 & ", " & LOAD_PAY_3 & ", " & LOAD_PAY_30 & ", " & LOAD_PAY_5 & ") " & " AND PA.IS_NOT_COSTING = 0 AND LS.TRUCK_CODE = |NONE| AND PA.IS_PAY_LIKE_TRUCKING <> 0 "

            If (Len(Activity) > 0) Then
                strSql = strSql & " AND T1.ACTIVITY_CODE = |" & Activity & "|"
            End If

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            GetContractor_LoadBased_Pay_Where_Equipment_Is_NONE = GetContractor_LoadBased_Pay_Where_Equipment_Is_NONE + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)

        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployeeOtherCost_Obsolete(ByVal aBlock As String, ByVal SD As Date, ByVal ED As Date, ByVal Activity As String) As Double

        '--------------------------------------------------------------------------
        '   Use the Function - GetEmployeeCost_EqIsNone Instead
        '--------------------------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim activityStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Len(Activity) = 0 Then
            activityStr = ""
        Else
            activityStr = " AND TD.PAY_ACTIVITY_CODE = |" & Activity & "|"
        End If

        '-------------------------------------------------------------------------
        '    Pay Employees, for time on Production Tickets
        '    Only if Contractor or None Equipment - Own Eq is in Full Rate
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM (EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR OW.OWNER = 0) " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeOtherCost_Obsolete = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeOtherCost_Obsolete = 0
        End If

        strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM (EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE = |" & aBlock & "|" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & " AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) AND (TD.TRUCK_CODE = |NONE| OR OW.OWNER = 0) " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeOtherCost_Obsolete = GetEmployeeOtherCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeOtherCost_Obsolete = GetEmployeeOtherCost_Obsolete
        End If

        '-------------------------------------------------------------------------
        '    Pay Empoloyees, for Time Where Equip & Truck are 'NONE' - Travel Time
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(TD.EMPLOYEE_COST+TD.EMPLOYEE_LOAD_COST) AS PAY FROM BLOCK_COST_DETAIL TD " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE = |" & aBlock & "|" & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeOtherCost_Obsolete = GetEmployeeOtherCost_Obsolete + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeOtherCost_Obsolete = GetEmployeeOtherCost_Obsolete
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetDetEmployeeCost_EqIsNone(ByVal aBlock As String, ByVal Activity As String, ByVal SD As Date, ByVal ED As Date, ByVal skipTrucking As Integer) As String

        '-------------------------------------------------------------------------
        '*************************************************************************
        'Make sure the function BELOW (GetDetEmployeeCost_EqIsNone) is above remains the same as this one
        '*************************************************************************
        '-------------------------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim strSql2 As String
        Dim strSql3 As String
        Dim activityStr As String
        Dim selStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Len(Activity) = 0 Then
            activityStr = ""
        Else
            activityStr = " AND TD.PAY_ACTIVITY_CODE = |" & Activity & "|"
        End If

        selStr = " ( SELECT T1.PAY+T1.EMPLOYEE_LOAD_COST AS COST, T1.ACTIVITY_CODE AS ACTIVITY_CODE, ST.EMPLOYEE_CODE AS PAYEE, TD.BLOCK_CODE, TD.TIME_SLIP_DATE AS TIME_SLIP_DATE, TD.EQUIPMENT_CODE AS EQUIPMENT, TD.EMPLOYEE_CODE AS EMPLOYEE_CODE, T1.PAY_WEIGHT AS PAY_UNITS, " & " IIF(T1.PAY_WEIGHT <> 0, (T1.PAY+T1.EMPLOYEE_LOAD_COST)/T1.PAY_WEIGHT, 0)  AS PAY_RATE, T1.MEASURE_CODE AS PAY_BASIS, |Production Details Employees| AS SOURCE, TD.CODE AS REF "

        '-------------------------------------------------------------------------
        '    Pay Employees, for time on Production Tickets
        '    Only if Contractor or None Equipment - Own Eq is in Full Rate
        '-------------------------------------------------------------------------
        strSql = selStr & " FROM ((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN EMPLOYEE_STATEMENTS ST ON ST.CODE = T1.STATEMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ")" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 AND PA.IS_PAY_LIKE_TRUCKING = 0 " & activityStr & " ) "

        If (skipTrucking) Then

            'Skip

        Else

            selStr = " SELECT T1.PAY+T1.EMPLOYEE_LOAD_COST AS COST, T1.ACTIVITY_CODE AS ACTIVITY_CODE, ST.EMPLOYEE_CODE AS PAYEE, TD.BLOCK_CODE, TD.TIME_SLIP_DATE AS TIME_SLIP_DATE, TD.TRUCK_CODE AS EQUIPMENT, TD.EMPLOYEE_CODE AS EMPLOYEE_CODE, T1.PAY_WEIGHT AS PAY_UNITS, " & " IIF(T1.PAY_WEIGHT <> 0, (T1.PAY+T1.EMPLOYEE_LOAD_COST)/T1.PAY_WEIGHT, 0)  AS PAY_RATE, T1.MEASURE_CODE AS PAY_BASIS, |Production Details Employees| AS SOURCE, TD.CODE AS REF "

            strSql2 = selStr & " FROM ((( EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN EMPLOYEE_STATEMENTS ST ON ST.CODE = T1.STATEMENT_CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ") " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 AND PA.IS_PAY_LIKE_TRUCKING <> 0 " & activityStr

            strSql = strSql & " UNION ALL (" & strSql2 & ")"

        End If

        '-------------------------------------------------------------------------
        '    Pay Empoloyees, for Time Where Equip & Truck are |NONE| - Travel Time
        '-------------------------------------------------------------------------
        selStr = " SELECT TD.EMPLOYEE_COST+TD.EMPLOYEE_LOAD_COST AS COST, TD.PAY_ACTIVITY_CODE AS ACTIVITY_CODE, TD.EMPLOYEE_CODE AS PAYEE, TD.BLOCK_CODE, TD.TIME_SLIP_DATE AS TIME_SLIP_DATE, TD.EQUIPMENT_CODE AS EQUIPMENT, TD.EMPLOYEE_CODE AS EMPLOYEE_CODE, TD.HOURS AS PAY_UNITS, " & " IIF(TD.HOURS <> 0, (TD.EMPLOYEE_COST+TD.EMPLOYEE_LOAD_COST)/TD.HOURS, 0)  AS PAY_RATE, |HRS| AS PAY_BASIS, |Time Details Employees| AS SOURCE, TD.CODE AS REF "

        strSql3 = selStr & " FROM BLOCK_COST_DETAIL TD " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = TD.PAY_ACTIVITY_CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE IN (" & aBlock & ")" & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 " & activityStr

        If (skipTrucking) Then
            strSql3 = strSql3 & " AND PA.IS_PAY_LIKE_TRUCKING = 0 "
        End If

        strSql = strSql & " UNION ALL (" & strSql3 & ")"

        GetDetEmployeeCost_EqIsNone = strSql

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetEmployeeCost_EqIsNone(ByVal aBlock As String, ByVal Activity As String, ByVal SD As Date, ByVal ED As Date, ByVal skipTrucking As Integer) As Double

        '-------------------------------------------------------------------------
        '*************************************************************************
        'Make sure the function ABOVE (GetDetEmployeeCost_EqIsNone) is above remains the same as this one
        '*************************************************************************
        '-------------------------------------------------------------------------

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String
        Dim activityStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If Len(Activity) = 0 Then
            activityStr = ""
        Else
            activityStr = " AND TD.PAY_ACTIVITY_CODE = |" & Activity & "|"
        End If

        '-------------------------------------------------------------------------
        '    Pay Employees, for time on Production Tickets
        '    Only if Contractor or None Equipment - Own Eq is in Full Rate
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM ((EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ")" & " AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 AND PA.IS_PAY_LIKE_TRUCKING = 0 " & activityStr

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeCost_EqIsNone = IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeCost_EqIsNone = 0
        End If

        If (skipTrucking) Then

            'Skip

        Else

            strSql = "SELECT SUM(T1.PAY+T1.EMPLOYEE_LOAD_COST) AS PAY " & " FROM ((EMPLOYEE_STATEMENT_DETAILS T1 INNER JOIN PRODUCTION_DETAILS TD ON T1.LOAD_CODE = TD.CODE) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = T1.ACTIVITY_CODE) " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T1.BLOCK_CODE IN (" & aBlock & ") " & " AND T1.PAY_TYPE IN (" & LOAD_PAY_4 & ") " & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 AND PA.IS_PAY_LIKE_TRUCKING <> 0 " & activityStr

            aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
            If (gobjErrors.Count = 0) Then
            Else
                DisplayErrors()
                Exit Function
            End If

            If (Not aRS.EOF) Then
                GetEmployeeCost_EqIsNone = GetEmployeeCost_EqIsNone + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
            Else
                GetEmployeeCost_EqIsNone = GetEmployeeCost_EqIsNone
            End If

        End If

        '-------------------------------------------------------------------------
        '    Pay Empoloyees, for Time Where Equip & Truck are |NONE| - Travel Time
        '-------------------------------------------------------------------------
        strSql = "SELECT SUM(TD.EMPLOYEE_COST+TD.EMPLOYEE_LOAD_COST) AS PAY " & " FROM BLOCK_COST_DETAIL TD " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = TD.PAY_ACTIVITY_CODE " & " WHERE TD.TIME_SLIP_DATE >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND TD.TIME_SLIP_DATE < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND TD.BLOCK_CODE IN (" & aBlock & ")" & " AND (TD.TRUCK_CODE = |NONE| OR TD.TRUCK_CODE IS NULL) AND (TD.EQUIPMENT_CODE = |NONE| OR TD.EQUIPMENT_CODE IS NULL) " & " AND PA.IS_NOT_COSTING = 0 " & activityStr

        If (skipTrucking) Then
            strSql = strSql & " AND PA.IS_PAY_LIKE_TRUCKING = 0 "
        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not aRS.EOF) Then
            GetEmployeeCost_EqIsNone = GetEmployeeCost_EqIsNone + IIf(IsDbNull(aRS(0).Value), 0, aRS(0).Value)
        Else
            GetEmployeeCost_EqIsNone = GetEmployeeCost_EqIsNone
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetTruckingRevenue(ByVal LoadCode As String, ByVal isPayLikeTruckingIncluded As Integer) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        If (isPayLikeTruckingIncluded) Then

            strSql = " SELECT SUM(RV.PAY) AS TRUCK_PAY FROM (( LOADSLIPS T0 " & " INNER JOIN LOADSLIP_REVENUE RV ON RV.LOAD_CODE = T0.CODE ) " & " INNER JOIN PAY_ACTIVITIES PA ON PA.CODE = RV.ACTIVITY_CODE ) " & " WHERE PA.IS_PAY_LIKE_TRUCKING <> 0 AND T0.CODE = |" & LoadCode & "|"

        Else

            strSql = " SELECT SUM(RV.PAY) AS TRUCK_PAY FROM LOADSLIPS T0 " & " INNER JOIN LOADSLIP_REVENUE RV ON RV.LOAD_CODE = T0.CODE " & " WHERE RV.ACTIVITY_CODE = |" & gVars.gTruckingCode & "| AND T0.CODE = |" & LoadCode & "|"

        End If

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (IsDbNull(aRS(0).Value)) Then
            GetTruckingRevenue = 0
        Else
            GetTruckingRevenue = aRS(0).Value
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetScaleData(ByRef LogCount As Double, ByRef NetScale As Double, ByRef Defect As Double, ByVal LoadCode As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        strSql = " SELECT SUM(SCALE_VOLUME) AS VOLUME, SUM(CULL) AS CULL, SUM(PIECE_COUNT) AS PIECES FROM LOADSLIPS  " & " WHERE L_UDF_9 = |" & LoadCode & "| AND C_PART_13_CODE = |D| "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (IsDbNull(aRS(0).Value)) Then
            GetScaleData = 0
            NetScale = 0
            Defect = 0
            LogCount = 0
        Else
            GetScaleData = 1
            NetScale = aRS(0).Value
            Defect = aRS(1).Value
            LogCount = aRS(2).Value
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetScaleVolumeForTheLoad(ByVal LoadCode As String) As Double

        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim strSql As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        strSql = " SELECT SUM(SCALE_VOLUME) AS VOLUME FROM LOADSLIPS  " & " WHERE L_UDF_9 = |" & LoadCode & "| AND C_PART_13_CODE = |D| "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (IsDbNull(aRS(0).Value)) Then
            GetScaleVolumeForTheLoad = NA
        Else
            GetScaleVolumeForTheLoad = aRS(0).Value
        End If

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function

    Public Function GetSampleConversionRate(ByVal blStr As String, ByVal ED As Date, ByRef n As Double) As Double

        'n is returned as the count of sample loads with sample data

        Dim strSql As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim tVolume As Double
        Dim tWeight As Double
        Dim aVal As clsLib_ReportClasses.clsRate
        Dim smplStr As String

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        smplStr = "SELECT DISTINCT L_UDF_9 FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blStr & ") AND T0.C_PART_13_CODE = |D| "

        strSql = "SELECT SUM(T0.VOLUME), SUM(LOAD_COUNT) AS LOADCOUNT FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blStr & ") AND T0.C_PART_13_CODE = |H| AND IS_SAMPLE_LOAD <> 0 " & " AND T0.CODE IN (" & smplStr & ")"

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not IsDbNull(aRS(0).Value)) Then
            If (System.Math.Abs(aRS(0).Value) < EPS) Then
                GetSampleConversionRate = NA
                Exit Function
            Else
                tWeight = aRS(0).Value
                n = aRS(1).Value
            End If
        Else
            GetSampleConversionRate = NA
            Exit Function
        End If

        strSql = "SELECT SUM(T0.SCALE_VOLUME) AS VOL FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blStr & ") AND T0.C_PART_13_CODE = |D| "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not IsDbNull(aRS(0).Value)) Then
            If (System.Math.Abs(aRS(0).Value) < EPS) Then
                GetSampleConversionRate = NA
                Exit Function
            Else
                tVolume = aRS(0).Value
            End If
        Else
            GetSampleConversionRate = NA
            Exit Function
        End If

        GetSampleConversionRate = tWeight / tVolume 'Tons/MBF

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Function GetSampleDistribution(ByVal blStr As String, ByVal SD As Date, ByVal ED As Date) As Collection

        Dim strSql As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt
        Dim tVol As Double
        Dim aVal As clsLib_ReportClasses.clsRate

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        GetSampleDistribution = New collection

        strSql = "SELECT SUM(T0.SCALE_VOLUME) AS VOL FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blStr & ") AND T0.C_PART_13_CODE = |D| "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        If (Not IsDbNull(aRS(0).Value)) Then
            If (System.Math.Abs(aRS(0).Value) < EPS) Then
                Exit Function
            Else
                tVol = aRS(0).Value
            End If

        Else
            Exit Function
        End If

        strSql = "SELECT SUM(T0.SCALE_VOLUME) AS VOL, C_PART_4_CODE FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blStr & ") AND T0.C_PART_13_CODE = |D| " & " GROUP BY C_PART_4_CODE ORDER BY C_PART_4_CODE "

        aRS = gDatalayer.LoadRecordset(strSql, gobjErrors)
        If (gobjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Function
        End If

        Do While aRS.EOF = False

            aVal = New clsLib_ReportClasses.clsRate

            aVal.CPART_CODE(4) = aRS("C_PART_4_CODE").Value
            aVal.VOLUME = aRS("VOL").Value

            If (System.Math.Abs(tVol) > 0) Then

                aVal.Rate = aRS("VOL").Value / tVol

            Else

                aVal.Rate = NA

            End If

            GetSampleDistribution.Add(aVal, aVal.CPART_CODE(4))

            aRS.MoveNext()

        Loop

        aRS = Nothing

        Exit Function

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Function


    Public Sub GetConvertedVolumeForWeightOnlyLoads(ByVal blCode As String, ByVal SD As Date, ByVal ED As Date, ByRef nLoads As Double, ByRef tWeight As Double, ByRef ConversionRate As Double, _
                                                    ByRef wgtVolume As Double, ByRef nSmpLoads As Double, ByRef tSmpWeight As Double, ByRef smpVolume As Double, ByRef nSmple As Double)

        Dim strSql As String
        Dim aRS As CSOFT_RECORDSET_EXT.clsRecordsetExt

        If (gDebugMode = 0) Then On Error GoTo Err_Handler

        '-------------------------------------------------
        '  Get Weight and Load Count for Unsampled Loads
        '-------------------------------------------------

        strSql = "SELECT SUM(T0.VOLUME) AS WEIGHT, SUM(LOAD_COUNT) AS LOADCOUNT FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blCode & ") " & " AND (T0.C_PART_13_CODE = |NONE| OR T0.C_PART_13_CODE = |H| OR T0.C_PART_13_CODE IS NULL) " & " AND IS_SAMPLE_LOAD = 0 "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Sub
        End If

        If IsDBNull(aRS(0).Value) Then
            nLoads = 0
            tWeight = 0
        Else
            If (aRS(1).Value = 0) Then
                nLoads = 0
                tWeight = 0
            Else
                nLoads = aRS(1).Value
                tWeight = aRS(0).Value
            End If
        End If

        '-------------------------------------------------
        '  Get Weight and Load Count for Sampled Loads
        '-------------------------------------------------

        strSql = "SELECT SUM(T0.VOLUME) AS WEIGHT, SUM(LOAD_COUNT) AS LOADCOUNT FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blCode & ") " & " AND (T0.C_PART_13_CODE = |NONE| OR T0.C_PART_13_CODE = |H| OR T0.C_PART_13_CODE IS NULL) " & " AND IS_SAMPLE_LOAD <> 0 "

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Sub
        End If

        If IsDBNull(aRS(0).Value) Then
            nSmpLoads = 0
            tSmpWeight = 0
        Else
            If (aRS(1).Value = 0) Then
                nSmpLoads = 0
                tSmpWeight = 0
            Else
                nSmpLoads = aRS(1).Value
                tSmpWeight = aRS(0).Value
            End If
        End If

        ConversionRate = GetSampleConversionRate(blCode, ED, nSmple) 'nSmple is the number of sample loads for all time
        If (System.Math.Abs(ConversionRate - NA) < EPS) Then
            wgtVolume = 0
        Else
            wgtVolume = tWeight / ConversionRate
        End If


        '-------------------------------------------------
        '  Get Actual Volume of Sampled Loads
        '-------------------------------------------------

        strSql = "SELECT SUM(T0.SCALE_VOLUME) AS SCALED_VOLUME FROM LOADSLIPS T0 " & " WHERE T0.DATE_OUT >= " & gVars.gDI & IntlDateFormat(SD) & gVars.gDI & " AND T0.DATE_OUT < " & gVars.gDI & IntlDateFormat(ED) & gVars.gDI & " AND T0.BLOCK_CODE IN (" & blCode & ") " & " AND T0.C_PART_13_CODE = |D|"

        aRS = gDataLayer.LoadRecordset(strSql, gObjErrors)
        If (gObjErrors.Count = 0) Then
        Else
            DisplayErrors()
            Exit Sub
        End If

        If IsDBNull(aRS(0).Value) Then
            smpVolume = 0
        Else
            smpVolume = aRS(0).Value
        End If

        aRS = Nothing

        Exit Sub

Err_Handler:
        TurnMeOff()
        Dim desc As String = Err.Description
        Dim eException As Exception = Err.GetException
        Dim src As String = getSource(New StackTrace(eException, True))
        DisplayErrorsNet(desc, src)

    End Sub

End Module
