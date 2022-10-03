Imports System.Console
Imports System.ConsoleColor
Imports System.Diagnostics.Process
Imports System.IO
Imports System.IO.Directory
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms.Application
Imports System.Windows.Forms.MessageBox
Imports System.Windows.Forms.MessageBoxButtons
Imports System.Windows.Forms.MessageBoxIcon

Friend Module Common
#Region "Helper"
    ''' <summary>
    ''' Check internet connection.
    ''' </summary>
    ''' <returns>Connection state.</returns>
    Private Function IsInternetAvailable()
        Dim objResp As WebResponse
        Try
            objResp = WebRequest.Create(New Uri(My.Resources.link_base)).GetResponse
            objResp.Close()
            objResp = Nothing
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Check update.
    ''' </summary>
    Friend Sub CheckUpdate()
        If IsInternetAvailable() AndAlso Not (New WebClient).DownloadString(My.Resources.link_ver).Contains(My.Resources.app_ver) Then
            Show($"A newer version of the {My.Resources.app_name} is available to download.", "Update", OK, Information)
            Run(New FrmUpdate)
        End If
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' End process.
    ''' </summary>
    ''' <param name="name">Process name.</param>
    Friend Sub KillProcess(name)
        If GetProcessesByName(name).Count > 0 Then
            For Each item In GetProcessesByName(name)
                item.Kill()
            Next
        End If
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Create directory advanced.
    ''' </summary>
    ''' <param name="path">Folder path.</param>
    Friend Sub CreateDirectoryAdv(path)
        If Not Exists(path) Then
            CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Delete file advanced.
    ''' </summary>
    ''' <param name="path">File path.</param>
    Friend Sub DeleteFileAdv(path)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Publish value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cap">Caption.</param>
    ''' <param name="cell">Cell address.</param>
    Friend Sub PublishValue(xlApp, cap, cell)
        ForegroundColor = Cyan
        Write(vbTab & cap)
        ForegroundColor = White
        Dim value = Val(ReadLine)
        If value > 0 Then
            xlApp.Range(cell).Activate()
            xlApp.ActiveCell.FormulaR1C1 = value
        End If
    End Sub
#End Region

#Region "Timer"
    ''' <summary>
    ''' Start tmr advanced.
    ''' </summary>
    ''' <param name="tmr">Timer.</param>
    <Extension()>
    Friend Sub StartAdv(tmr)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Stop tmr advanced.
    ''' </summary>
    ''' <param name="tmr">Timer.</param>
    <Extension()>
    Friend Sub StopAdv(tmr)
        If tmr.Enabled Then
            tmr.Start()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' Infomation.
    ''' </summary>
    Friend Sub Info()
        Clear()
        ForegroundColor = Blue
        WriteLine(My.Resources.gr_name)
        WriteLine(My.Resources.cc_text)
        ForegroundColor = Green
        WriteLine(vbCrLf & My.Resources.app_true_name)
    End Sub
#End Region
End Module
