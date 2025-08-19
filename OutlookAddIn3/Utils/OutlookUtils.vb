Imports System.Runtime.InteropServices
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace OutlookAddIn3.Utils
    Public Class OutlookUtils
        Public Shared Function FormatDateTime(dt As DateTime) As String
            Return dt.ToString("yyyy-MM-dd HH:mm:ss")
        End Function

        Public Shared Function SafeGetString(value As Object) As String
            Return If(value IsNot Nothing, value.ToString(), String.Empty)
        End Function

        ''' <summary>
        ''' 安全获取邮件项
        ''' </summary>
        ''' <param name="entryId">邮件项的EntryID</param>
        ''' <returns>邮件项对象，如果获取失败则返回Nothing</returns>
        Public Shared Function SafeGetItemFromID(entryId As String) As Object
            Try
                If String.IsNullOrWhiteSpace(entryId) Then
                    Return Nothing
                End If
                
                ' 检查 Outlook 应用程序和会话是否可用
                If Globals.ThisAddIn?.Application?.Session Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("Outlook 应用程序或会话不可用")
                    Return Nothing
                End If
                
                Return Globals.ThisAddIn.Application.Session.GetItemFromID(entryId)
            Catch ex As System.Runtime.InteropServices.COMException
                System.Diagnostics.Debug.WriteLine($"COM异常获取邮件项 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                ' 静默处理，不再抛出异常
                Return Nothing
            Catch ex As System.Runtime.InteropServices.InvalidComObjectException
                System.Diagnostics.Debug.WriteLine($"无效的COM对象异常: {ex.Message}")
                Return Nothing
            Catch ex As System.UnauthorizedAccessException
                System.Diagnostics.Debug.WriteLine($"访问被拒绝异常: {ex.Message}")
                Return Nothing
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"获取邮件项时发生异常 ({ex.GetType().Name}): {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 安全获取邮件项并验证类型
        ''' </summary>
        ''' <typeparam name="T">期望的邮件项类型</typeparam>
        ''' <param name="entryId">邮件项的EntryID</param>
        ''' <returns>指定类型的邮件项，如果类型不匹配或获取失败则返回Nothing</returns>
        Public Shared Function SafeGetItemFromID(Of T As Class)(entryId As String) As T
            Try
                If String.IsNullOrWhiteSpace(entryId) Then
                    Return Nothing
                End If
                
                ' 检查 Outlook 应用程序和会话是否可用
                If Globals.ThisAddIn?.Application?.Session Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("Outlook 应用程序或会话不可用")
                    Return Nothing
                End If

                Dim item = Globals.ThisAddIn.Application.Session.GetItemFromID(entryId)
                If item IsNot Nothing AndAlso TypeOf item Is T Then
                    Return DirectCast(item, T)
                End If

                If item IsNot Nothing Then
                    SafeReleaseComObject(item)
                End If

                Return Nothing
            Catch ex As System.Runtime.InteropServices.COMException
                System.Diagnostics.Debug.WriteLine($"COM异常获取类型{GetType(T).Name}邮件项 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                ' 静默处理，不再抛出异常
                Return Nothing
            Catch ex As System.Runtime.InteropServices.InvalidComObjectException
                System.Diagnostics.Debug.WriteLine($"无效的COM对象异常获取类型{GetType(T).Name}: {ex.Message}")
                Return Nothing
            Catch ex As System.UnauthorizedAccessException
                System.Diagnostics.Debug.WriteLine($"访问被拒绝异常获取类型{GetType(T).Name}: {ex.Message}")
                Return Nothing
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"获取类型{GetType(T).Name}邮件项时发生异常 ({ex.GetType().Name}): {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 安全释放COM对象
        ''' </summary>
        ''' <param name="comObject">要释放的COM对象</param>
        Public Shared Sub SafeReleaseComObject(comObject As Object)
            Try
                If comObject IsNot Nothing Then
                    Marshal.ReleaseComObject(comObject)
                End If
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"释放COM对象时出错: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace