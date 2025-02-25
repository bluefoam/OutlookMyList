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

    Public Shared Sub LoadConversationMails(lvMails As ListView, currentMailEntryID As String)
        lvMails.BeginUpdate()
        Try
            lvMails.Items.Clear()
            Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
            Dim conversation As Conversation = Nothing

            If TypeOf currentItem Is MailItem Then
                conversation = DirectCast(currentItem, MailItem).GetConversation()
            ElseIf TypeOf currentItem Is AppointmentItem Then
                conversation = DirectCast(currentItem, AppointmentItem).GetConversation()
            End If

            If conversation IsNot Nothing Then
                Dim table As Table = conversation.GetTable()
                table.Columns.Add("EntryID")
                table.Columns.Add("SentOn")
                table.Columns.Add("SenderName")
                table.Columns.Add("Subject")
                table.Sort("[SentOn]", False)

                Dim items As New List(Of ListViewItem)()
                Dim highlightIndex As Integer = -1

                Do Until table.EndOfTable
                    Dim row As Row = table.GetNextRow()
                    Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString())
                    Dim entryId As String = GetPermanentEntryID(mailItem)

                    Dim lvi As New ListViewItem With {
                        .Text = If(row("SentOn") IsNot Nothing,
                                 DateTime.Parse(row("SentOn").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                 "Unknown Date"),
                        .Tag = entryId
                    }
                    lvi.SubItems.Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                    lvi.SubItems.Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))

                    If String.Equals(entryId, currentMailEntryID.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        lvi.BackColor = System.Drawing.Color.FromArgb(255, 255, 200)
                        lvi.Font = New System.Drawing.Font(lvMails.Font, System.Drawing.FontStyle.Bold)
                        highlightIndex = items.Count
                    End If

                    items.Add(lvi)
                Loop

                lvMails.Items.AddRange(items.ToArray())

                If highlightIndex >= 0 Then
                    lvMails.Items(highlightIndex).Selected = True
                    lvMails.Items(highlightIndex).EnsureVisible()
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"LoadConversationMails error: {ex.Message}")
        Finally
            lvMails.EndUpdate()
        End Try
    End Sub

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

    Public Shared Function DisplayMailContent(mailEntryID As String) As String
        Try
            Dim mail As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailEntryID)

            If TypeOf mail Is MailItem Then
                Dim mailItem As MailItem = DirectCast(mail, MailItem)
                Return $"<html><body style='font-family: Arial; padding: 10px;'>" &
                       $"<h3>{mailItem.Subject}</h3>" &
                       $"<div style='margin-bottom: 10px;'>" &
                       $"<strong>发件人:</strong> {mailItem.SenderName}<br/>" &
                       $"<strong>时间:</strong> {mailItem.ReceivedTime:yyyy-MM-dd HH:mm:ss}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;' onclick='handleLinks(event)'>" &
                       $"{mailItem.HTMLBody}" &
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
                Return $"<html><body style='font-family: Arial; padding: 10px;'>" &
                       $"<h3>{appointment.Subject}</h3>" &
                       $"<div style='margin-bottom: 10px;'>" &
                       $"<strong>开始时间:</strong> {appointment.Start:yyyy-MM-dd HH:mm:ss}<br/>" &
                       $"<strong>结束时间:</strong> {appointment.End:yyyy-MM-dd HH:mm:ss}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;'>" &
                       $"{appointment.Body}" &
                       $"</div></body></html>"
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