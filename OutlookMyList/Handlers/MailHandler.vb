Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics
Imports System.Windows.Forms   ' 添加这行

Namespace OutlookMyList.Handlers
    Public Class MailHandler
    ''' <summary>
    ''' 将ListView项目的Tag转换为EntryID字符串
    ''' </summary>
    ''' <param name="tag">ListView项目的Tag对象</param>
    ''' <returns>EntryID字符串</returns>
    Private Shared Function ConvertEntryIDToString(tag As Object) As String
        Try
            If tag Is Nothing Then
                Return String.Empty
            End If
            
            ' 如果Tag是字节数组（长格式EntryID的二进制数据）
            If TypeOf tag Is Byte() Then
                Dim bytes As Byte() = DirectCast(tag, Byte())
                ' 将字节数组转换为十六进制字符串
                Return BitConverter.ToString(bytes).Replace("-", "")
            End If
            
            ' 如果Tag是字符串，直接返回
            Return tag.ToString()
        Catch ex As System.Exception
            Debug.WriteLine($"ConvertEntryIDToString error: {ex.Message}")
            Return String.Empty
        End Try
    End Function
    Public Shared Function GetPermanentEntryID(item As Object) As String
        Try
            If TypeOf item Is MailItem Then
                Return DirectCast(item, MailItem).EntryID
            ElseIf TypeOf item Is AppointmentItem Then
                Return DirectCast(item, AppointmentItem).EntryID
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"GetPermanentEntryID error: {ex.Message}")
        End Try
        Return String.Empty
    End Function

    Public Shared Sub UpdateHighlight(lvMails As ListView, currentMailEntryID As String)
        lvMails.BeginUpdate()
        Try
            For Each item As ListViewItem In lvMails.Items
                Dim itemEntryID As String = ConvertEntryIDToString(item.Tag).Trim()
                Dim currentEntryID As String = currentMailEntryID.Trim()
                If String.Equals(itemEntryID, currentEntryID, StringComparison.OrdinalIgnoreCase) Then
                    item.BackColor = System.Drawing.Color.FromArgb(255, 255, 200)
                    item.Font = New System.Drawing.Font(lvMails.Font, System.Drawing.FontStyle.Bold)
                    item.Selected = True
                    item.EnsureVisible()
                Else
                    item.BackColor = System.Drawing.SystemColors.Window
                    item.Font = lvMails.Font
                    item.Selected = False
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateHighlight error: {ex.Message}")
        Finally
            lvMails.EndUpdate()
        End Try
    End Sub

    Private Shared Function GetTaskStatus(status As OlTaskStatus) As String
        Select Case status
            Case OlTaskStatus.olTaskComplete
                Return "已完成"
            Case OlTaskStatus.olTaskDeferred
                Return "已推迟"
            Case OlTaskStatus.olTaskInProgress
                Return "进行中"
            Case OlTaskStatus.olTaskNotStarted
                Return "未开始"
            Case OlTaskStatus.olTaskWaiting
                Return "等待中"
            Case Else
                Return "未知状态"
        End Select
    End Function

    Public Shared Function ReplaceTableTag(mailItemHTML As String) As String
        Dim oldTableTag As String
        Dim newTableTag As String

        ' 定义要查找和替换的字符串
        oldTableTag = "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" width=""100%"">"
        newTableTag = "<table class=""hidden-table"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" width=""100%"">"

        ' 检查是否包含旧的表格标签
        If InStr(mailItemHTML, oldTableTag) > 0 Then
            ' 替换第一个匹配的表格标签
            Return Replace(mailItemHTML, oldTableTag, newTableTag, 1, 1)
            ' 输出或处理替换后的HTML
            'Debug.Print resultHTML
        Else
            ' 如果没有找到，输出原始HTML
            'Debug.Print "未找到匹配的表格标签，原始HTML保持不变。"
            'Debug.Print mailItemHTML
            Return mailItemHTML
        End If
    End Function

    Public Shared Function DisplayMailContent(mailEntryID As String) As String
        Try
            Dim mail As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mailEntryID)

            If TypeOf mail Is MailItem Then
                Dim mailItem As MailItem = DirectCast(mail, MailItem)
                Try
                    Dim subject As String = If(String.IsNullOrEmpty(mailItem.Subject), "无主题", mailItem.Subject)
                    Dim senderName As String = If(String.IsNullOrEmpty(mailItem.SenderName), "未知", mailItem.SenderName)
                    Dim receivedTime As String = If(mailItem.ReceivedTime = DateTime.MinValue, "未知", mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim htmlBody As String = If(String.IsNullOrEmpty(mailItem.HTMLBody), "", ReplaceTableTag(mailItem.HTMLBody))
                    
                    Return $"<html><body style='font-family: Arial; padding: 10px; Font-size:12px;'>" &
                           $"<h4 style='color: var(--theme-color, #0078d7);'>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>发件人:</strong> {senderName}<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>时间:</strong> {receivedTime}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid var(--theme-color, #0078d7); padding-top: 10px;'>" &
                           $"<style>.hidden-table {{display: none;}} img {{display: none;}}</style>" &
                           $"{htmlBody}" &
                           $"</div>" &
                           "</body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问邮件属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问邮件属性</body></html>"
                End Try

                       '"<div style='border-top: 1px solid #ccc; padding-top: 10px;'> <style> .hidden-table {display: none;} img {display: none;}</style>"
            ElseIf TypeOf mail Is AppointmentItem Then
                Dim appointment As AppointmentItem = DirectCast(mail, AppointmentItem)
                Try
                    Dim subject As String = If(String.IsNullOrEmpty(appointment.Subject), "无主题", appointment.Subject)
                    Dim organizer As String = If(String.IsNullOrEmpty(appointment.Organizer), "未知", appointment.Organizer)
                    Dim startTime As String = If(appointment.Start = DateTime.MinValue, "未设置", appointment.Start.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim endTime As String = If(appointment.End = DateTime.MinValue, "未设置", appointment.End.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim location As String = If(String.IsNullOrEmpty(appointment.Location), "未设置", appointment.Location)
                    Dim body As String = If(String.IsNullOrEmpty(appointment.Body), "", appointment.Body.Replace(vbCrLf, "<br>").Replace("<br><br>", "<br>"))
                    
                    Return $"<html><body style='font-family: Arial; padding: 10px;Font-size:12px;'>" &
                           $"<h4 style='color: var(--theme-color, #0078d7);'>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>组织者:</strong> {organizer}<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>开始时间:</strong> {startTime}<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>结束时间:</strong> {endTime}<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>地点:</strong> {location}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid var(--theme-color, #0078d7); padding-top: 10px;Font-size:12px;'>" &
                           $"{body}" &
                           $"</div></body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问会议属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问会议属性时发生异常: {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议属性</body></html>"
                End Try
            ElseIf TypeOf mail Is TaskItem Then
                Dim task As TaskItem = DirectCast(mail, TaskItem)
                Try
                    Dim subject As String = If(String.IsNullOrEmpty(task.Subject), "无主题", task.Subject)
                    Dim startDate As String = If(task.StartDate = DateTime.MinValue, "未设置", task.StartDate.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim dueDate As String = If(task.DueDate = DateTime.MinValue, "未设置", task.DueDate.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim percentComplete As Integer = task.PercentComplete
                    Dim status As String = GetTaskStatus(task.Status)
                    Dim body As String = If(String.IsNullOrEmpty(task.Body), "", task.Body.Replace(vbCrLf, "<br>").Replace("<br><br>", "<br>"))
                    
                    Return $"<html><body style='font-family: Arial; padding: 10px;Font-size:12px;'>" &
                           $"<h4 style='color: var(--theme-color, #0078d7);'>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>开始时间:</strong> {startDate}<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>结束时间:</strong> {dueDate}<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>完成度:</strong> {percentComplete}%<br/>" &
                           $"<strong style='color: var(--theme-color, #0078d7);'>状态:</strong> {status}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid var(--theme-color, #0078d7); padding-top: 10px;Font-size:12px;'>" &
                           $"{body}" &
                           $"</div></body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问任务属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问任务属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问任务属性时发生异常: {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问任务属性</body></html>"
                End Try
            ElseIf TypeOf mail Is MeetingItem Then
                Dim meeting As MeetingItem = DirectCast(mail, MeetingItem)

                ' 获取关联的约会项目以获取时间信息
                Dim associatedAppointment As AppointmentItem = meeting.GetAssociatedAppointment(False)

                ' 获取会议状态信息
                Dim meetingStatus As String = "会议邀请"
                Select Case meeting.MessageClass
                    Case "IPM.Schedule.Meeting.Canceled"
                        meetingStatus = "会议已取消"
                        Return "<html><body><p>会议已取消, 无法显示内容</p></body></html>"
                    Case "IPM.Schedule.Meeting.Request"
                        meetingStatus = "会议邀请"
                    Case "IPM.Schedule.Meeting.Resp.Pos"
                        meetingStatus = "已接受"
                    Case "IPM.Schedule.Meeting.Resp.Neg"
                        meetingStatus = "已拒绝"
                    Case "IPM.Schedule.Meeting.Resp.Tent"
                        meetingStatus = "暂定"
                End Select


                Try
                    Dim subject As String = If(String.IsNullOrEmpty(meeting.Subject), "无主题", meeting.Subject)
                    Dim senderName As String = If(String.IsNullOrEmpty(meeting.SenderName), "未知", meeting.SenderName)
                    Dim body As String = If(String.IsNullOrEmpty(meeting.Body), "", meeting.Body.Replace(vbCrLf, "<br>").Replace(" <br><br>", "<br>"))
                    
                    Return $"<html><body style='font-family: Arial; padding: 10px;Font-size:12px;'>" &
                           $"<h4>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                           $"<strong>状态:</strong> {meetingStatus}<br/>" &
                           $"<strong>发件人:</strong> {senderName}<br/>" &
                           $"<strong>开始时间:</strong> {If(associatedAppointment IsNot Nothing AndAlso associatedAppointment.Start <> DateTime.MinValue,
                                                          associatedAppointment.Start.ToString("yyyy-MM-dd HH:mm:ss"),
                                                          "未设置")}<br/>" &
                           $"<strong>结束时间:</strong> {If(associatedAppointment IsNot Nothing AndAlso associatedAppointment.End <> DateTime.MinValue,
                                                          associatedAppointment.End.ToString("yyyy-MM-dd HH:mm:ss"),
                                                          "未设置")}<br/>" &
                           $"<strong>地点:</strong> {If(associatedAppointment IsNot Nothing AndAlso Not String.IsNullOrEmpty(associatedAppointment.Location),
                                                     associatedAppointment.Location,
                                                     "未设置")}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size:12px;'>" &
                           $"{body}" &
                           $"</div></body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问会议属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问会议属性时发生异常: {ex.Message}")
                    Return "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议属性</body></html>"
                End Try

                If 0 Then
                    Return $"<html><body style='font-family: Arial; padding: 10px;Font-size:12px;'>" &
                       $"<h4>{meeting.Subject}</h4>" &
                       $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                       $"<strong>开始时间:</strong> {meeting.Start:yyyy-MM-dd HH:mm:ss}<br/>" &
                       $"<strong>结束时间:</strong> {meeting.End:yyyy-MM-dd HH:mm:ss}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size:12px;'>" &
                       $"{meeting.HTMLBody.Replace(vbCrLf, "<br>").Replace(" <br><br>", "<br>")}" &
                       $"</div></body></html>"
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"显示邮件内容时出错: {ex.Message}")
        End Try
        Return "<html><body><p>无法显示内容</p></body></html>"
    End Function
    Public Shared Sub OpenLink(url As String)
        Try
            Process.Start(New ProcessStartInfo With {
                .FileName = url,
                .UseShellExecute = True
            })
        Catch ex As System.Exception
            Debug.WriteLine($"打开链接出错: {ex.Message}")
            MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。")
        End Try
    End Sub
End Class
End Namespace