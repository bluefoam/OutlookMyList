Partial Class OutlookRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        If (container IsNot Nothing) Then
            container.Add(Me)
        End If
    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())
        InitializeComponent()
    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Factory.CreateRibbonTab()
        Me.Group1 = Factory.CreateRibbonGroup()
        Me.ToggleButton1 = Factory.CreateRibbonToggleButton()
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()

        'Tab1
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.ControlId.OfficeId = "TabAddIns"
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "MyList"
        Me.Tab1.Name = "Tab1"

        'Group1
        Me.Group1.Items.Add(Me.ToggleButton1)
        Me.Group1.Label = "视图"
        Me.Group1.Name = "Group1"

        'ToggleButton1
        Me.ToggleButton1.Label = "ShowPane"
        Me.ToggleButton1.Name = "ToggleButton1"
        Me.ToggleButton1.ScreenTip = "切换邮件会话面板的显示状态"

        'OutlookRibbon
        Me.Name = "OutlookRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Group1.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ToggleButton1 As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection
    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon1() As OutlookRibbon
        Get
            Return Me.GetRibbon(Of OutlookRibbon)()
        End Get
    End Property
End Class