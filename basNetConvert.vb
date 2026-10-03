Imports System.Reflection

Module basNetConvert

    Public Const vbCtrlMask As Integer = 2

    Public Function getSource(ByRef ST As StackTrace) As String

        Dim sf As StackFrame
        Dim methodInfo As MethodInfo
        Dim methodName As String
        Dim sourceStr As String
        Dim ii As Integer

        On Error GoTo Err_Handler

        ii = testStackTrace(ST, 0)
        If ii = 0 Then
            sf = ST.GetFrame(1)
        Else
            sf = ST.GetFrame(0)
        End If

        methodInfo = sf.GetMethod

        methodName = methodInfo.Name

        sourceStr = methodName

        Dim filename As String = sf.GetFileName
        Dim linenumber As Integer = sf.GetFileLineNumber

        If Len(filename) > 0 Then
            ii = filename.LastIndexOf("\")
            If ii <> -1 Then
                filename = filename.Substring(ii + 1, Len(filename) - ii - 1)
            Else
            End If
            sourceStr = sourceStr & ":" & filename
        End If

        If Len(linenumber) > 0 Then
            sourceStr = sourceStr & ":" & linenumber
        End If

        Return sourceStr

Err_Handler:
        Return "???"

    End Function

    Private Function testStackTrace(ByRef ST As StackTrace, ByVal idx As Integer) As Integer

        Dim sf As StackFrame

        On Error GoTo Err_Handler

        sf = ST.GetFrame(0)

        Dim filename As String = sf.GetFileName
        Dim linenumber As Integer = sf.GetFileLineNumber

        If filename IsNot Nothing Then
            testStackTrace = True
        Else
            testStackTrace = False
        End If

        Exit Function
Err_Handler:

    End Function

    Public Sub DisplayErrorsNet(ByVal desc As String, ByVal src As String)

        MsgBox(desc & vbCrLf & src, MsgBoxStyle.Critical, src)

    End Sub

End Module
