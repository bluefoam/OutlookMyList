Public Class ThisAddIn
    Private WithEvents currentExplorer As Outlook.Explorer
    Private customTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private mailThreadPane As MailThreadPane
    Private taskMonitor As TaskMonitor

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        currentExplorer = Me.Application.ActiveExplorer
        InitializeMailPane()

        ' 初始化任务监视器
        taskMonitor = New TaskMonitor()
        taskMonitor.Initialize()
    End Sub

    Private Sub InitializeMailPane()
        mailThreadPane = New MailThreadPane()
        customTaskPane = Me.CustomTaskPanes.Add(mailThreadPane, "相关邮件")
        customTaskPane.Width = 400
        customTaskPane.Visible = True
        ' Initialize with empty values
        mailThreadPane.UpdateMailList(String.Empty, String.Empty)
    End Sub

    Private Sub currentExplorer_SelectionChange() Handles currentExplorer.SelectionChange
        If mailThreadPane Is Nothing OrElse customTaskPane Is Nothing Then Return

        If currentExplorer.Selection.Count > 0 Then
            Dim selection As Object = currentExplorer.Selection(1)
            If TypeOf selection Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(selection, Outlook.MailItem)
                If mail.ConversationID IsNot Nothing Then
                    mailThreadPane.UpdateMailList(mail.ConversationID, mail.EntryID)
                End If
            ElseIf TypeOf selection Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(selection, Outlook.AppointmentItem)
                If appointment.GlobalAppointmentID IsNot Nothing Then
                    mailThreadPane.UpdateMailList(appointment.GlobalAppointmentID, appointment.EntryID)
                End If
            ElseIf TypeOf selection Is Outlook.MeetingItem Then
                Dim meeting As Outlook.MeetingItem = DirectCast(selection, Outlook.MeetingItem)
                'Dim associatedMail As Outlook.MailItem = meeting.GetAssociatedItem()
                'If associatedMail IsNot Nothing Then
                mailThreadPane.UpdateMailList(String.Empty, meeting.EntryID)
                'End If
            ElseIf TypeOf selection Is Outlook.TaskItem Then
                Dim task As Outlook.TaskItem = DirectCast(selection, Outlook.TaskItem)
                'If task.ConversationID IsNot Nothing Then
                    mailThreadPane.UpdateMailList(String.Empty, task.EntryID)
                'End If
            ElseIf TypeOf selection Is Outlook.ContactItem Then
                Dim contact As Outlook.ContactItem = DirectCast(selection, Outlook.ContactItem)
                mailThreadPane.UpdateMailList(String.Empty, contact.EntryID)
            End If
        End If
    End Sub

    Public Sub ToggleTaskPane()
        If customTaskPane IsNot Nothing Then
            customTaskPane.Visible = Not customTaskPane.Visible
        End If
    End Sub

    Public Sub UpdateMailList()
        If mailThreadPane IsNot Nothing Then
            mailThreadPane.UpdateMailList(String.Empty, String.Empty)
        End If
    End Sub
    
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 清理任务监视器
        If taskMonitor IsNot Nothing Then
            taskMonitor.Cleanup()
        End If
    End Sub
End Class
