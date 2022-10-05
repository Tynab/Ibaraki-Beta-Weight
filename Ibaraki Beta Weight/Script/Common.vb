Imports System.Console
Imports System.ConsoleColor
Imports System.Diagnostics.Process
Imports System.IO
Imports System.IO.Directory
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Threading.Thread
Imports System.Windows.Forms
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
    Private Function IsNetAvail()
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
    Friend Sub ChkUpd()
        If IsNetAvail() AndAlso Not (New WebClient).DownloadString(My.Resources.link_ver).Contains(My.Resources.app_ver) Then
            Show($"A newer version of the {My.Resources.app_name} is available to download.", "Update", OK, Information)
            Run(New FrmUpdate)
        End If
    End Sub

    ''' <summary>
    ''' Fade in form
    ''' </summary>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        While frm.Opacity < 1
            frm.Opacity += 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Fade out form
    ''' </summary>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        While frm.Opacity > 0
            frm.Opacity -= 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' End process.
    ''' </summary>
    ''' <param name="name">Process name.</param>
    Friend Sub KillPrcs(name As String)
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
    Friend Sub CrtDirAdv(path As String)
        If Not Exists(path) Then
            CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Delete file advanced.
    ''' </summary>
    ''' <param name="path">File path.</param>
    Friend Sub DelFileAdv(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Yes/No question (1/0).
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Function YSQ(caption As String)
        PrefWarn(caption)
        ForegroundColor = White
        Dim vu = Val(ReadLine)
        If vu <> 0 Or vu <> 1 Then
            Do Until vu = 0 Or vu = 1
                PrefWarn(caption)
                ForegroundColor = Red
                vu = Val(ReadLine)
            Loop
        End If
        Return vu
    End Function

    ''' <summary>
    ''' Redirect value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="caption">Caption.</param>
    ''' <param name="cell">Cell address.</param>
    Friend Sub RdrVal(xlApp As Microsoft.Office.Interop.Excel.Application, caption As String, cell As String)
        ForegroundColor = Cyan
        Write(vbTab & caption)
        ForegroundColor = White
        Dim value = Val(ReadLine)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub

    ''' <summary>
    ''' Redirect value with detail to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="caption">Caption.</param>
    ''' <param name="detail">Detail.</param>
    ''' <param name="cell">Cell address.</param>
    'Friend Sub RdrValDtl(xlApp As Microsoft.Office.Interop.Excel.Application, caption As String, detail As String, cell As String)
    '    ForegroundColor = Cyan
    '    Write(vbTab & caption)
    '    ForegroundColor = Magenta
    '    Write(detail)
    '    ForegroundColor = White
    '    Dim value = Val(ReadLine)
    '    If value > 0 Then
    '        DctVal(xlApp, cell, value)
    '    End If
    'End Sub

    ''' <summary>
    ''' Direct value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="vu">Value.</param>
    Friend Sub DctVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, vu As Object)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.FormulaR1C1 = vu
    End Sub

    ''' <summary>
    ''' Mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="vu">Value.</param>
    Friend Sub ModVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, vu As Object)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.FormulaR1C1 = vu
        xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
    End Sub

    ''' <summary>
    ''' Direct value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    Friend Sub ClrVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.MergeArea.ClearContents()
    End Sub

    ''' <summary>
    ''' Import value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="caption">Caption.</param>
    ''' <param name="cell">Cell address.</param>
    Friend Sub ImpVal(xlApp As Microsoft.Office.Interop.Excel.Application, caption As String, cell As String)
        PrefImp(caption)
        Dim value = Val(ReadLine)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub
#End Region

#Region "Timer"
    ''' <summary>
    ''' Start timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Stop timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
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

    ''' <summary>
    ''' Prefix input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Sub PrefInp(caption As String)
        ForegroundColor = Cyan
        WriteLine(vbTab & vbTab & caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Prefix input RIP.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    'Friend Sub PrefInpRip(caption As String)
    '    ForegroundColor = Cyan
    '    WriteLine(vbTab & caption)
    '    ForegroundColor = White
    'End Sub

    ''' <summary>
    ''' Prefix input detail.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <param name="detail">Detail.</param>
    Friend Sub PrefInpDtl(caption As String, detail As String)
        ForegroundColor = Cyan
        Write(vbTab & caption)
        ForegroundColor = Magenta
        Write(detail)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Prefix import.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Sub PrefImp(caption As String)
        Info()
        PrefInp(caption)
    End Sub

    ''' <summary>
    ''' Prefix warning.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Sub PrefWarn(caption As String)
        Info()
        ForegroundColor = Yellow
        WriteLine(vbTab & vbTab & caption)
    End Sub
#End Region
End Module
