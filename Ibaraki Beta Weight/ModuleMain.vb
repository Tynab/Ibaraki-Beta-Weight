Imports System.Console
Imports System.ConsoleColor
Imports System.Text.Encoding

Public Module ModuleMain
    ''' <summary>
    ''' Main.
    ''' </summary>
    Public Sub Main()
        OutputEncoding = UTF8
        If Not My.Settings.Chk_Key Then
            If InputBox("Enter Serial", "License key") = My.Resources.key_ser Then
                UpdVldLic()
                RunApp()
            Else
                ForegroundColor = Red
                Write("Wrong license. Press any key to exit...")
                ReadKey()
                Environment.Exit(0)
            End If
        Else
            RunApp()
        End If
    End Sub
End Module
