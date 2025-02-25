Namespace OutlookAddIn3.Utils
    Public Class OutlookUtils
        Public Shared Function FormatDateTime(dt As DateTime) As String
            Return dt.ToString("yyyy-MM-dd HH:mm:ss")
        End Function

        Public Shared Function SafeGetString(value As Object) As String
            Return If(value IsNot Nothing, value.ToString(), String.Empty)
        End Function
    End Class
End Namespace