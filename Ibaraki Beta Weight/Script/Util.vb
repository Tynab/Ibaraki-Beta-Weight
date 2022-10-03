Imports System.Console
Imports System.ConsoleColor

Friend Module Util
    ''' <summary>
    ''' 外周/内周GL-150.
    ''' </summary>
    Friend Sub GLNormal150(xlApp)
        ForegroundColor = Yellow
        WriteLine(vbTab & vbTab & "外周/内周GL-150")
        PublishValue(xlApp, "  4G: ", "BA18")
        PublishValue(xlApp, "3.5G: ", "BA19")
        PublishValue(xlApp, "  3G: ", "BA20")
        PublishValue(xlApp, "2.5G: ", "BA21")
        PublishValue(xlApp, "  2G: ", "BA22")
        PublishValue(xlApp, "1.5G: ", "BA23")
        PublishValue(xlApp, "  1G: ", "BA24")
        PublishValue(xlApp, "0.5G: ", "BA25")
    End Sub
End Module
