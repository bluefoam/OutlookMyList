Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms

Public Class Win32Helper
    ' Win32 API 声明
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindowEx(hwndParent As IntPtr, hwndChildAfter As IntPtr, lpszClass As String, lpszWindow As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetParent(hWndChild As IntPtr, hWndNewParent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetClassName(hWnd As IntPtr, lpClassName As System.Text.StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function EnumChildWindows(hWndParent As IntPtr, lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
    End Function

    Public Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

    ' 常量定义
    Public Const SWP_NOZORDER As UInteger = &H4
    Public Const SWP_NOACTIVATE As UInteger = &H10
    Public Const SWP_SHOWWINDOW As UInteger = &H40
    Public Shared ReadOnly HWND_TOP As IntPtr = New IntPtr(0)

    ' 结构体定义
    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer

        Public ReadOnly Property Width As Integer
            Get
                Return Right - Left
            End Get
        End Property

        Public ReadOnly Property Height As Integer
            Get
                Return Bottom - Top
            End Get
        End Property
    End Structure

    ' 查找Outlook主窗口
    Public Shared Function FindOutlookMainWindow() As IntPtr
        ' 尝试多种可能的Outlook主窗口类名
        Dim outlookWindow As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
        If outlookWindow = IntPtr.Zero Then
            outlookWindow = FindWindow("OutlookGrid", Nothing)
        End If
        If outlookWindow = IntPtr.Zero Then
            outlookWindow = FindWindow("XLMAIN", Nothing) ' Excel有时会被误识别
        End If
        If outlookWindow = IntPtr.Zero Then
            ' 尝试通过窗口标题查找
            outlookWindow = FindWindow(Nothing, "Microsoft Outlook")
        End If
        
        System.Diagnostics.Debug.WriteLine("找到Outlook主窗口: " & outlookWindow.ToString())
        Return outlookWindow
    End Function

    ' 查找阅读窗格
    Public Shared Function FindReadingPane(outlookMainWindow As IntPtr) As IntPtr
        Dim readingPane As IntPtr = IntPtr.Zero
        Dim foundWindows As New List(Of String)
        
        ' 枚举子窗口查找阅读窗格
        EnumChildWindows(outlookMainWindow, 
            Function(hWnd As IntPtr, lParam As IntPtr) As Boolean
                Dim className As New System.Text.StringBuilder(256)
                GetClassName(hWnd, className, className.Capacity)
                Dim classNameStr = className.ToString()
                
                ' 记录找到的窗口类名用于调试
                foundWindows.Add(classNameStr)
                
                ' Outlook阅读窗格的类名可能是这些之一
                If classNameStr.Contains("_WwG") OrElse 
                   classNameStr.Contains("Internet Explorer_Server") OrElse
                   classNameStr.Contains("RichEdit") OrElse
                   classNameStr.Contains("RICHEDIT") OrElse
                   classNameStr.Contains("WordDocument") OrElse
                   classNameStr.Contains("_WwF") Then
                    
                    ' 获取窗口大小，选择最大的作为主要阅读区域
                    Dim rect As RECT
                    If GetWindowRect(hWnd, rect) AndAlso rect.Width > 200 AndAlso rect.Height > 100 Then
                        readingPane = hWnd
                        System.Diagnostics.Debug.WriteLine("找到阅读窗格: " & classNameStr & " (" & rect.Width & "x" & rect.Height & ")")
                        Return False ' 停止枚举
                    End If
                End If
                
                Return True ' 继续枚举
            End Function, IntPtr.Zero)
        
        ' 调试信息：输出所有找到的窗口类名
        System.Diagnostics.Debug.WriteLine("所有子窗口类名: " & String.Join(", ", foundWindows))
        System.Diagnostics.Debug.WriteLine("最终选择的阅读窗格: " & readingPane.ToString())
        
        Return readingPane
    End Function

    ' 在指定窗口下方创建自定义面板
    Public Shared Function CreateBottomPanel(parentWindow As IntPtr, targetWindow As IntPtr, panelControl As Control, ByRef originalParent As IntPtr) As Boolean
        Try
            ' 获取目标窗口的位置和大小
            Dim targetRect As RECT
            If Not GetWindowRect(targetWindow, targetRect) Then
                System.Diagnostics.Debug.WriteLine("无法获取目标窗口矩形")
                Return False
            End If

            ' 获取父窗口的位置
            Dim parentRect As RECT
            If Not GetWindowRect(parentWindow, parentRect) Then
                System.Diagnostics.Debug.WriteLine("无法获取父窗口矩形")
                Return False
            End If

            ' 调试信息
            System.Diagnostics.Debug.WriteLine($"目标窗口位置: ({targetRect.Left}, {targetRect.Top}) 大小: {targetRect.Width}x{targetRect.Height}")
            System.Diagnostics.Debug.WriteLine($"父窗口位置: ({parentRect.Left}, {parentRect.Top}) 大小: {parentRect.Width}x{parentRect.Height}")

            ' 计算面板应该放置的位置（在目标窗口下方）
            Dim panelHeight As Integer = 150
            
            ' 将屏幕坐标转换为父窗口的客户区坐标
            Dim panelX As Integer = targetRect.Left - parentRect.Left
            Dim panelY As Integer = targetRect.Bottom - parentRect.Top - panelHeight
            Dim panelWidth As Integer = targetRect.Width
            
            ' 确保面板在父窗口范围内
            If panelX < 0 Then panelX = 0
            If panelY < 0 Then panelY = targetRect.Top - parentRect.Top + targetRect.Height - panelHeight
            If panelWidth > parentRect.Width Then panelWidth = parentRect.Width
            
            System.Diagnostics.Debug.WriteLine($"计算的面板位置: ({panelX}, {panelY}) 大小: {panelWidth}x{panelHeight}")

            ' 首先尝试调整目标窗口大小，为面板腾出空间
            Dim newTargetHeight As Integer = targetRect.Height - panelHeight
            If newTargetHeight > 100 Then ' 确保目标窗口还有足够空间
                SetWindowPos(targetWindow, HWND_TOP, 
                            targetRect.Left - parentRect.Left, 
                            targetRect.Top - parentRect.Top, 
                            targetRect.Width, 
                            newTargetHeight, 
                            SWP_NOZORDER Or SWP_NOACTIVATE)
                System.Diagnostics.Debug.WriteLine("已调整目标窗口大小")
            End If

            ' 设置面板的父窗口为Outlook主窗口
            originalParent = SetParent(panelControl.Handle, parentWindow)
            System.Diagnostics.Debug.WriteLine($"设置父窗口，旧父窗口: {originalParent}")

            ' 定位面板到计算的位置
            Dim result As Boolean = SetWindowPos(panelControl.Handle, HWND_TOP, 
                        panelX, panelY, panelWidth, panelHeight, 
                        SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
            
            System.Diagnostics.Debug.WriteLine($"面板定位结果: {result}")
            
            ' 确保面板可见
            panelControl.Visible = True
            panelControl.BringToFront()

            Return result
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("创建底部面板失败: " & ex.Message)
            Return False
        End Try
    End Function
End Class