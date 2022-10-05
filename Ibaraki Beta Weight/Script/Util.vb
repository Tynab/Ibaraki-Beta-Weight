Imports System.Console
Imports Microsoft.Office.Interop.Excel

Friend Module Util
    ''' <summary>
    ''' 外周/内周GL-150.
    ''' </summary>
    ''' ''' <param name="xlApp">Excel Application.</param>
    Friend Sub GL50(xlApp As Application)
        RdrVal(xlApp, "  4G: ", "BA18")
        RdrVal(xlApp, "3.5G: ", "BA19")
        RdrVal(xlApp, "  3G: ", "BA20")
        RdrVal(xlApp, "2.5G: ", "BA21")
        RdrVal(xlApp, "  2G: ", "BA22")
        RdrVal(xlApp, "1.5G: ", "BA23")
        RdrVal(xlApp, "  1G: ", "BA24")
        RdrVal(xlApp, "0.5G: ", "BA25")
    End Sub

    ''' <summary>
    ''' 外周深GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub GL300(xlApp As Application)
        RdrVal(xlApp, "  4G: ", "BA27")
        RdrVal(xlApp, "3.5G: ", "BA28")
        RdrVal(xlApp, "  3G: ", "BA29")
        RdrVal(xlApp, "2.5G: ", "BA30")
        RdrVal(xlApp, "  2G: ", "BA31")
        RdrVal(xlApp, "1.5G: ", "BA32")
        RdrVal(xlApp, "  1G: ", "BA33")
        RdrVal(xlApp, "0.5G: ", "BA34")
    End Sub

    ''' <summary>
    ''' 外周深GL-300/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub GL300Cut(xlApp As Application)
        RdrVal(xlApp, "  3G: ", "BA38")
        RdrVal(xlApp, "2.5G: ", "BA39")
        RdrVal(xlApp, "  2G: ", "BA40")
        RdrVal(xlApp, "1.5G: ", "BA41")
        RdrVal(xlApp, "  1G: ", "BA42")
        RdrVal(xlApp, "0.5G: ", "BA43")
    End Sub

    ''' <summary>
    ''' 外周深GL-400.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub GL400(xlApp As Application)
        RdrVal(xlApp, "  4G: ", "BA45")
        RdrVal(xlApp, "3.5G: ", "BA46")
        RdrVal(xlApp, "  3G: ", "BA47")
        RdrVal(xlApp, "2.5G: ", "BA48")
        RdrVal(xlApp, "  2G: ", "BA49")
        RdrVal(xlApp, "1.5G: ", "BA50")
        RdrVal(xlApp, "  1G: ", "BA51")
        RdrVal(xlApp, "0.5G: ", "BA52")
    End Sub

    ''' <summary>
    ''' 外周深GL-400/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub GL400Cut(xlApp As Application)
        RdrVal(xlApp, "  4G: ", "BA54")
        RdrVal(xlApp, "3.5G: ", "BA55")
        RdrVal(xlApp, "  3G: ", "BA56")
        RdrVal(xlApp, "2.5G: ", "BA57")
        RdrVal(xlApp, "  2G: ", "BA58")
        RdrVal(xlApp, "1.5G: ", "BA59")
        RdrVal(xlApp, "  1G: ", "BA60")
        RdrVal(xlApp, "0.5G: ", "BA61")
    End Sub

    ''' <summary>
    ''' ガレージ外周GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub GL300Gar(xlApp As Application)
        RdrVal(xlApp, "  4G: ", "BA72")
        RdrVal(xlApp, "3.5G: ", "BA73")
        RdrVal(xlApp, "  3G: ", "BA74")
        RdrVal(xlApp, "2.5G: ", "BA75")
        RdrVal(xlApp, "  2G: ", "BA76")
        RdrVal(xlApp, "1.5G: ", "BA77")
        RdrVal(xlApp, "  1G: ", "BA78")
        RdrVal(xlApp, "0.5G: ", "BA79")
    End Sub

    ''' <summary>
    ''' スラブユニット.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub SlapUnit(xlApp As Application)
        RdrVal(xlApp, "  4G: ", "BA99")
        RdrVal(xlApp, "3.5G: ", "BA100")
        RdrVal(xlApp, "  3G: ", "BA101")
        RdrVal(xlApp, "2.5G: ", "BA102")
        RdrVal(xlApp, "  2G: ", "BA103")
        RdrVal(xlApp, "1.5G: ", "BA104")
        RdrVal(xlApp, "  1G: ", "BA105")
    End Sub

    ''' <summary>
    ''' コーナー.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Corner(xlApp As Application)
        RdrVal(xlApp, "D16: ", "BA165")
        RdrVal(xlApp, "D13: ", "BA164")
        RdrVal(xlApp, "D10: ", "BA163")
    End Sub

    ''' <summary>
    ''' ストレート.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Straight(xlApp As Application)
        RdrVal(xlApp, "D16: ", "BA162")
        RdrVal(xlApp, "D13: ", "BA161")
        RdrVal(xlApp, "D10: ", "BA160")
    End Sub

    ''' <summary>
    ''' キャップタイヤ (320).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub CapTire(xlApp As Application)
        RdrVal(xlApp, "D16: ", "BA188")
        RdrVal(xlApp, "D10: ", "BA181")
    End Sub

    ''' <summary>
    ''' ロングコーナー (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub LongCorner(xlApp As Application)
        RdrVal(xlApp, " 750×2250: ", "BA179")
        RdrVal(xlApp, " 750×1750: ", "BA177")
        RdrVal(xlApp, " 750×1250: ", "BA178")
        ' 1250×1250
        PrefInpDtl("1250×1250 ", "[4.1]: ")
        Dim vu = Val(ReadLine)
        If vu > 0 Then
            DctVal(xlApp, "AH167", "1250×1250")
            ModVal(xlApp, "CM167", 4.1)
            DctVal(xlApp, "BA167", vu)
        End If
    End Sub
End Module
