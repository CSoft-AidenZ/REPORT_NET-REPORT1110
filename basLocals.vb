Module basLocals

    Public gobjTables As clsLib_DataAccess.clsTables = Nothing
    Public gobjErrors As clsLib_DataAccess.clsErrors = Nothing
    Public gVars As clsLib_DataAccess.clsVars = Nothing
    Public gDataLayer As clsLib_DataAccess.clsDataLayer
    Public gConst As CSOFT_BASIC_FUNCTIONS.clsConstants = Nothing
    Public gBF As CSOFT_BASIC_FUNCTIONS.clsBasicFunctions = Nothing
    Public gComFn As clsBasComnFuncs.clsCommonFuncs = Nothing
    Public gGenFn As clsBasComnFuncs.clsBasGenFuncs = Nothing

    Public basReportGlobals As clsLib_ReportClasses.claBasReportGlobals
    Public basReportDllFunctions As clsLib_ReportClasses.clsBasReportDllFunctions
    Public basGridFuncs As CSOFT_BASIC_FUNCTIONS.clsBasGridFuncs
    Public basGenFuncs As clsBasComnFuncs.clsBasGenFuncs

    Public Const EPS As Double = 0.00001
    Public Const NA As Integer = -9999

    Public Const CellTypeCheckBox As Integer = 10

    Public Enum TypeHAlignConstants
        TypeHAlignLeft = 0
        TypeHAlignRight = 1
        TypeHAlignCenter = 2
    End Enum

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

End Module
