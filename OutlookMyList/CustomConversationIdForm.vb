Imports System.Windows.Forms

Public Class CustomConversationIdForm
    Inherits Form

    Private txtCustomId As TextBox

    Public ReadOnly Property EnteredId As String
        Get
            If txtCustomId IsNot Nothing Then
                Return txtCustomId.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Sub New(originalId As String, currentCustomId As String)
        Me.Text = "设置自定义会话ID"
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.ClientSize = New Drawing.Size(480, 220)

        Dim lblOriginal As New Label()
        lblOriginal.AutoSize = True
        lblOriginal.Text = "原始会话ID: " & If(String.IsNullOrEmpty(originalId), "(无)", originalId)
        lblOriginal.Location = New Drawing.Point(12, 12)

        Dim lblCurrent As New Label()
        lblCurrent.AutoSize = True
        lblCurrent.Text = "当前自定义会话ID: " & If(String.IsNullOrEmpty(currentCustomId), "(未设置)", currentCustomId)
        lblCurrent.Location = New Drawing.Point(12, 36)

        Dim lblHint As New Label()
        lblHint.AutoSize = True
        lblHint.Text = "留空并点" & "确定" & "可清除自定义会话ID。"
        lblHint.Location = New Drawing.Point(12, 60)

        Dim lblInput As New Label()
        lblInput.AutoSize = True
        lblInput.Text = "新的自定义会话ID："
        lblInput.Location = New Drawing.Point(12, 90)

        txtCustomId = New TextBox()
        txtCustomId.Location = New Drawing.Point(12, 110)
        txtCustomId.Size = New Drawing.Size(450, 24)
        txtCustomId.Text = currentCustomId

        Dim btnOK As New Button()
        btnOK.Text = "确定"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = New Drawing.Point(282, 160)
        btnOK.Size = New Drawing.Size(80, 28)

        Dim btnCancel As New Button()
        btnCancel.Text = "取消"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = New Drawing.Point(382, 160)
        btnCancel.Size = New Drawing.Size(80, 28)

        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel

        Me.Controls.AddRange(New Control() {lblOriginal, lblCurrent, lblHint, lblInput, txtCustomId, btnOK, btnCancel})
    End Sub
End Class