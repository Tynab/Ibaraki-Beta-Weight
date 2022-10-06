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
    ''' <returns>Answer value.</returns>
    Friend Function YNQ(caption As String)
        PrefWarn(caption)
        ForegroundColor = White
        Dim value = Val(ReadLine)
        If value <> 0 Or value <> 1 Then
            Do Until value = 0 Or value = 1
                PrefWarn(caption)
                ForegroundColor = Red
                value = Val(ReadLine)
            Loop
        End If
        Return value
    End Function

    ''' <summary>
    ''' Direct value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="value">Value.</param>
    Friend Sub DctVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Object)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.FormulaR1C1 = value
    End Sub

    ''' <summary>
    ''' Mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="value">Value.</param>
    Friend Sub ModVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As Object)
        xlApp.Range(cell).Activate()
        xlApp.ActiveCell.FormulaR1C1 = value
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
    ''' Publish value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="cell">Cell address.</param>
    ''' <param name="value">Value.</param>
    Friend Sub PubVal(xlApp As Microsoft.Office.Interop.Excel.Application, cell As String, value As String)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub

    ''' <summary>
    ''' Publish mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubModVal(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Publish mod value to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="title">Title rebar.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubModValFull(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, title As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Publish mod value expansion to excel.
    ''' </summary>
    ''' <param name="xlApp">Excel application.</param>
    ''' <param name="row">Row number.</param>
    ''' <param name="name">Name rebar.</param>
    ''' <param name="weight">Weight rebar.</param>
    ''' <param name="price">Price rebar.</param>
    ''' <param name="number">Number rebar.</param>
    Friend Sub PubModValExp(xlApp As Microsoft.Office.Interop.Excel.Application, row As String, name As String, weight As Double, price As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            ModVal(xlApp, $"CQ{row}", price)
            DctVal(xlApp, $"BA{row}", number)
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
    ''' Title desciption.
    ''' </summary>
    ''' <param name="description">Description.</param>
    Friend Sub TitDecs(description As String)
        ForegroundColor = Magenta
        Write(description)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Detail input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlInp(caption As String)
        ForegroundColor = Cyan
        Write(vbTab & vbTab & caption)
        ForegroundColor = White
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Detail input RIP.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlInpRip(caption As String)
        ForegroundColor = Cyan
        Write(vbTab & caption)
        ForegroundColor = White
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Detail input description.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <param name="description">Description.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlInpDtl(caption As String, description As String)
        ForegroundColor = Cyan
        Write(vbTab & vbTab & caption)
        TitDecs(description)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Detail input RIP description.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <param name="description">Description.</param>
    ''' <returns>Input value.</returns>
    Friend Function DtlInpRipDesc(caption As String, description As String)
        ForegroundColor = Cyan
        Write(vbTab & caption)
        TitDecs(description)
        Return Val(ReadLine)
    End Function

    ''' <summary>
    ''' Header input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Friend Function HdrInp(caption As String)
        Info()
        Return DtlInp(caption)
    End Function

    ''' <summary>
    ''' Header input description.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <param name="description">Description.</param>
    ''' <returns>Input value.</returns>
    Friend Function HdrInpDesc(caption As String, description As String)
        Info()
        Return DtlInpDtl(caption, description)
    End Function

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
