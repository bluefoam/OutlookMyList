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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ToggleButton1 = Me.Factory.CreateRibbonToggleButton
        Me.ToggleButtonPagination = Me.Factory.CreateRibbonToggleButton
        Me.ToggleButtonCacheEnabled = Me.Factory.CreateRibbonToggleButton
        Me.ButtonMergeConversation = Me.Factory.CreateRibbonButton
        Me.ButtonSettings = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.ControlId.OfficeId = "TabMail"
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "MyList"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ToggleButton1)
        Me.Group1.Items.Add(Me.ToggleButtonPagination)
        Me.Group1.Items.Add(Me.ToggleButtonCacheEnabled)
        Me.Group1.Items.Add(Me.ButtonMergeConversation)
        Me.Group1.Items.Add(Me.ButtonSettings)
        Me.Group1.Label = "视图"
        Me.Group1.Name = "Group1"
        '
        'ToggleButton1
        '
        Me.ToggleButton1.Label = "ShowPane"
        Me.ToggleButton1.Name = "ToggleButton1"
        Me.ToggleButton1.ScreenTip = "切换邮件会话面板的显示状态"
        '
        'ToggleButtonPagination
        '
        Me.ToggleButtonPagination.Label = "分页"
        Me.ToggleButtonPagination.Name = "ToggleButtonPagination"
        Me.ToggleButtonPagination.ScreenTip = "切换邮件列表分页功能"
        Me.ToggleButtonPagination.Checked = False
        '
        'ToggleButtonCacheEnabled
        '
        Me.ToggleButtonCacheEnabled.Label = "缓存开启"
        Me.ToggleButtonCacheEnabled.Name = "ToggleButtonCacheEnabled"
        Me.ToggleButtonCacheEnabled.ScreenTip = "开启/关闭缓存策略"
        Me.ToggleButtonCacheEnabled.Checked = True

        'ButtonMergeConversation
        Me.ButtonMergeConversation.Label = "合并自定义会话"
        Me.ButtonMergeConversation.Name = "ButtonMergeConversation"
        Me.ButtonMergeConversation.ScreenTip = "将所选邮件合并到同一自定义会话"
        '
        'ButtonSettings
        '
        Me.ButtonSettings.Label = "设置"
        Me.ButtonSettings.Name = "ButtonSettings"
        Me.ButtonSettings.ScreenTip = "配置错误提醒和其他选项"
        '
        'OutlookRibbon
        '
        Me.Name = "OutlookRibbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ToggleButton1 As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ToggleButtonPagination As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ToggleButtonCacheEnabled As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents ButtonMergeConversation As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection
    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon1() As OutlookRibbon
        Get
            Return Me.GetRibbon(Of OutlookRibbon)()
        End Get
    End Property
End Class