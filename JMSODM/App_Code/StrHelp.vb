Imports Microsoft.VisualBasic

Public Class StrHelp

    Public Shared Function GetInt(ByVal _obj As Object) As Integer
        Dim Ret As Integer
        Try
            Ret = Integer.Parse(_obj)
        Catch ex As Exception
            Ret = 0
        End Try

        Return Ret
    End Function




End Class
