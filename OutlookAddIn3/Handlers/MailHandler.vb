Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics
Imports System.Windows.Forms   ' 添加这行

Public Class MailHandler
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
                If String.Equals(item.Tag.ToString().Trim(), currentMailEntryID.Trim(), StringComparison.OrdinalIgnoreCase) Then
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
    Public Shared Function DisplayMailContent(mailEntryID As String) As String
        Try
            Dim mail As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailEntryID)

            If TypeOf mail Is MailItem Then
                Dim mailItem As MailItem = DirectCast(mail, MailItem)
                Return $"<html><body style='font-family: Arial; padding: 10px; Font-size=12px;'>" &
                       $"<h3>{If(String.IsNullOrEmpty(mailItem.Subject), "无主题", mailItem.Subject)}</h3>" &
                       $"<div style='margin-bottom: 10px;Font-size=12px;'>" &
                       $"<strong>发件人:</strong> {If(String.IsNullOrEmpty(mailItem.SenderName), "未知", mailItem.SenderName)}<br/>" &
                       $"<strong>时间:</strong> {If(mailItem.ReceivedTime = DateTime.MinValue, "未知", mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"))}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;' onclick='handleLinks(event)'>" &
                       $"{If(String.IsNullOrEmpty(mailItem.HTMLBody), "", mailItem.HTMLBody)}" &
                       $"</div>" &
                       $"<script>" &
                       "function handleLinks(e) {" &
                       "  if (e.target.tagName === 'A') {" &
                       "    e.preventDefault();" &
                       "    window.external.OpenLink(e.target.href);" &
                       "  }" &
                       "}" &
                       "</script>" &
                       "</body></html>"
            ElseIf TypeOf mail Is AppointmentItem Then
                Dim appointment As AppointmentItem = DirectCast(mail, AppointmentItem)
                Return $"<html><body style='font-family: Arial; padding: 10px;Font-size=12px;'>" &
                       $"<h3>{If(String.IsNullOrEmpty(appointment.Subject), "无主题", appointment.Subject)}</h3>" &
                       $"<div style='margin-bottom: 10px;Font-size=12px;'>" &
                       $"<strong>组织者:</strong> {If(String.IsNullOrEmpty(appointment.Organizer), "未知", appointment.Organizer)}<br/>" &
                       $"<strong>开始时间:</strong> {If(appointment.Start = DateTime.MinValue, "未设置", appointment.Start.ToString("yyyy-MM-dd HH:mm:ss"))}<br/>" &
                       $"<strong>结束时间:</strong> {If(appointment.End = DateTime.MinValue, "未设置", appointment.End.ToString("yyyy-MM-dd HH:mm:ss"))}<br/>" &
                       $"<strong>地点:</strong> {If(String.IsNullOrEmpty(appointment.Location), "未设置", appointment.Location)}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size=12px;'>" &
                       $"{If(String.IsNullOrEmpty(appointment.Body), "", appointment.Body.Replace(vbCrLf, "<br>").Replace("<br><br>", "<br>"))}" &
                       $"</div></body></html>"
            ElseIf TypeOf mail Is TaskItem Then
                Dim task As TaskItem = DirectCast(mail, TaskItem)
                Return $"<html><body style='font-family: Arial; padding: 10px;Font-size=12px;'>" &
                       $"<h3>{If(String.IsNullOrEmpty(task.Subject), "无主题", task.Subject)}</h3>" &
                       $"<div style='margin-bottom: 10px;Font-size=12px;'>" &
                       $"<strong>开始时间:</strong> {If(task.StartDate = DateTime.MinValue, "未设置", task.StartDate.ToString("yyyy-MM-dd HH:mm:ss"))}<br/>" &
                       $"<strong>结束时间:</strong> {If(task.DueDate = DateTime.MinValue, "未设置", task.DueDate.ToString("yyyy-MM-dd HH:mm:ss"))}<br/>" &
                       $"<strong>完成度:</strong> {task.PercentComplete}%<br/>" &
                       $"<strong>状态:</strong> {GetTaskStatus(task.Status)}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size=12px;'>" &
                       $"{If(String.IsNullOrEmpty(task.Body), "", task.Body.Replace(vbCrLf, "<br>").Replace("<br><br>", "<br>"))}" &
                       $"</div></body></html>"
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


                Return $"<html><body style='font-family: Arial; padding: 10px;Font-size=12px;'>" &
                       $"<h3>{If(String.IsNullOrEmpty(meeting.Subject), "无主题", meeting.Subject)}</h3>" &
                       $"<div style='margin-bottom: 10px;Font-size=12px;'>" &
                       $"<strong>状态:</strong> {meetingStatus}<br/>" &
                       $"<strong>发件人:</strong> {If(String.IsNullOrEmpty(meeting.SenderName), "未知", meeting.SenderName)}<br/>" &
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
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size=12px;'>" &
                       $"{If(String.IsNullOrEmpty(meeting.Body), "", meeting.Body.Replace(vbCrLf, "<br>").Replace(" <br><br>", "<br>"))}" &
                       $"</div></body></html>"

                If 0 Then
                    Return $"<html><body style='font-family: Arial; padding: 10px;Font-size=12px;'>" &
                       $"<h3>{meeting.Subject}</h3>" &
                       $"<div style='margin-bottom: 10px;Font-size=12px;'>" &
                       $"<strong>开始时间:</strong> {meeting.Start:yyyy-MM-dd HH:mm:ss}<br/>" &
                       $"<strong>结束时间:</strong> {meeting.End:yyyy-MM-dd HH:mm:ss}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size=12px;'>" &
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