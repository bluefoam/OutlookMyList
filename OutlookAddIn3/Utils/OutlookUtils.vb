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
            Return SafeGetItemFromID(entryId, Nothing)
        End Function

        ''' <summary>
        ''' 安全获取邮件项（带StoreID优化）
        ''' </summary>
        ''' <param name="entryId">邮件项的EntryID</param>
        ''' <param name="storeId">可选的StoreID，提供时可显著提升性能</param>
        ''' <returns>邮件项对象，如果获取失败则返回Nothing</returns>
        Public Shared Function SafeGetItemFromID(entryId As String, storeId As String) As Object
            Try
                If String.IsNullOrWhiteSpace(entryId) Then
                    Return Nothing
                End If
                
                ' 检查 Outlook 应用程序和会话是否可用
                If Globals.ThisAddIn?.Application?.Session Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("Outlook 应用程序或会话不可用")
                    Return Nothing
                End If
                
                ' 使用StoreID可以显著提升GetItemFromID的性能
                If Not String.IsNullOrWhiteSpace(storeId) Then
                    Return Globals.ThisAddIn.Application.Session.GetItemFromID(entryId, storeId)
                Else
                    Return Globals.ThisAddIn.Application.Session.GetItemFromID(entryId)
                End If
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
            Return SafeGetItemFromID(Of T)(entryId, Nothing)
        End Function

        ''' <summary>
        ''' 安全获取邮件项并验证类型（带StoreID优化）
        ''' </summary>
        ''' <typeparam name="T">期望的邮件项类型</typeparam>
        ''' <param name="entryId">邮件项的EntryID</param>
        ''' <param name="storeId">可选的StoreID，提供时可显著提升性能</param>
        ''' <returns>指定类型的邮件项，如果类型不匹配或获取失败则返回Nothing</returns>
        Public Shared Function SafeGetItemFromID(Of T As Class)(entryId As String, storeId As String) As T
            Try
                If String.IsNullOrWhiteSpace(entryId) Then
                    Return Nothing
                End If
                
                ' 检查 Outlook 应用程序和会话是否可用
                If Globals.ThisAddIn?.Application?.Session Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("Outlook 应用程序或会话不可用")
                    Return Nothing
                End If

                Dim item As Object
                ' 使用StoreID可以显著提升GetItemFromID的性能
                If Not String.IsNullOrWhiteSpace(storeId) Then
                    item = Globals.ThisAddIn.Application.Session.GetItemFromID(entryId, storeId)
                Else
                    item = Globals.ThisAddIn.Application.Session.GetItemFromID(entryId)
                End If

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
        ''' 快速打开邮件（针对 Flag 任务优化）
        ''' </summary>
        ''' <param name="entryId">邮件项的EntryID</param>
        ''' <param name="storeId">可选的StoreID，建议在 Flag 任务中提供以提升性能</param>
        ''' <returns>是否成功打开</returns>
        Public Shared Function FastOpenMailItem(entryId As String, Optional storeId As String = Nothing) As Boolean
            Try
                If String.IsNullOrWhiteSpace(entryId) Then
                    System.Diagnostics.Debug.WriteLine("FastOpenMailItem: EntryID为空")
                    Return False
                End If

                ' 检查 Outlook 应用程序和会话是否可用
                If Globals.ThisAddIn?.Application?.Session Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("FastOpenMailItem: Outlook 应用程序或会话不可用")
                    Return False
                End If

                ' 确保在主线程执行，提升 COM 调用性能
                If Threading.Thread.CurrentThread.GetApartmentState() <> Threading.ApartmentState.STA Then
                    System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 不在STA线程中，性能可能受影响")
                End If

                Dim mailItem As Object = SafeGetItemFromID(entryId, storeId)
                If mailItem IsNot Nothing Then
                    ' 直接显示邮件，False 参数表示非模态显示
                    If TypeOf mailItem Is Outlook.MailItem Then
                        DirectCast(mailItem, Outlook.MailItem).Display(False)
                        System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 邮件打开成功")
                        Return True
                    ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                        DirectCast(mailItem, Outlook.AppointmentItem).Display(False)
                        System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 会议项打开成功")
                        Return True
                    ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                        DirectCast(mailItem, Outlook.MeetingItem).Display(False)
                        System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 会议邮件打开成功")
                        Return True
                    ElseIf TypeOf mailItem Is Outlook.TaskItem Then
                        DirectCast(mailItem, Outlook.TaskItem).Display(False)
                        System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 任务项打开成功")
                        Return True
                    Else
                        ' 对于其他类型，尝试通用方法
                        CallByName(mailItem, "Display", CallType.Method, False)
                        System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 项目打开成功")
                        Return True
                    End If
                Else
                    System.Diagnostics.Debug.WriteLine("FastOpenMailItem: 无法获取邮件项")
                    Return False
                End If
            Catch ex As System.Runtime.InteropServices.COMException
                System.Diagnostics.Debug.WriteLine($"FastOpenMailItem COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Return False
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine($"FastOpenMailItem 异常: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取邮件项的短格式EntryID
        ''' </summary>
        ''' <param name="longEntryId">长格式EntryID</param>
        ''' <returns>短格式EntryID，如果获取失败则返回原始EntryID</returns>
        Public Shared Function GetShortEntryID(longEntryId As String) As String
            Try
                If String.IsNullOrWhiteSpace(longEntryId) Then
                    Return longEntryId
                End If
                
                ' 如果已经是短格式（以EF开头），直接返回
                If longEntryId.StartsWith("EF") Then
                    Return longEntryId
                End If
                
                ' 尝试通过长格式EntryID获取邮件项，然后读取其短格式EntryID
                Dim mailItem As Object = SafeGetItemFromID(longEntryId)
                If mailItem IsNot Nothing Then
                    Dim shortEntryId As String = mailItem.EntryID
                    SafeReleaseComObject(mailItem)
                    Return shortEntryId
                End If
                
                ' 如果无法获取，返回原始ID
                Return longEntryId
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"获取短格式EntryID失败: {ex.Message}")
                Return longEntryId
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