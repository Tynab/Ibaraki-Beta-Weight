Imports Microsoft.Office.Interop.Excel

Friend Module Util
    ''' <summary>
    ''' 運賃(2トン車).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Fare(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            DctVal(xlApp, "BA241", choosen)
        End If
    End Sub

    ''' <summary>
    ''' 外周/内周GL-150.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Unit150(xlApp As Application)
        PubVal(xlApp, "BA18", DtlInpRip("  4G: "))
        PubVal(xlApp, "BA19", DtlInpRip("3.5G: "))
        PubVal(xlApp, "BA20", DtlInpRip("  3G: "))
        PubVal(xlApp, "BA21", DtlInpRip("2.5G: "))
        PubVal(xlApp, "BA22", DtlInpRip("  2G: "))
        PubVal(xlApp, "BA23", DtlInpRip("1.5G: "))
        PubVal(xlApp, "BA24", DtlInpRip("  1G: "))
        PubVal(xlApp, "BA25", DtlInpRip("0.5G: "))
    End Sub

    ''' <summary>
    ''' 外周深GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA27", DtlInpRip("  4G: "))
            PubVal(xlApp, "BA28", DtlInpRip("3.5G: "))
            PubVal(xlApp, "BA29", DtlInpRip("  3G: "))
            PubVal(xlApp, "BA30", DtlInpRip("2.5G: "))
            PubVal(xlApp, "BA31", DtlInpRip("  2G: "))
            PubVal(xlApp, "BA32", DtlInpRip("1.5G: "))
            PubVal(xlApp, "BA33", DtlInpRip("  1G: "))
            PubVal(xlApp, "BA34", DtlInpRip("0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-300/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300Cut(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA38", DtlInpRip("  3G: "))
            PubVal(xlApp, "BA39", DtlInpRip("2.5G: "))
            PubVal(xlApp, "BA40", DtlInpRip("  2G: "))
            PubVal(xlApp, "BA41", DtlInpRip("1.5G: "))
            PubVal(xlApp, "BA42", DtlInpRip("  1G: "))
            PubVal(xlApp, "BA43", DtlInpRip("0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-400.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit400(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA45", DtlInpRip("  4G: "))
            PubVal(xlApp, "BA46", DtlInpRip("3.5G: "))
            PubVal(xlApp, "BA47", DtlInpRip("  3G: "))
            PubVal(xlApp, "BA48", DtlInpRip("2.5G: "))
            PubVal(xlApp, "BA49", DtlInpRip("  2G: "))
            PubVal(xlApp, "BA50", DtlInpRip("1.5G: "))
            PubVal(xlApp, "BA51", DtlInpRip("  1G: "))
            PubVal(xlApp, "BA52", DtlInpRip("0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-400/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit400Cut(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA54", DtlInpRip("  4G: "))
            PubVal(xlApp, "BA55", DtlInpRip("3.5G: "))
            PubVal(xlApp, "BA56", DtlInpRip("  3G: "))
            PubVal(xlApp, "BA57", DtlInpRip("2.5G: "))
            PubVal(xlApp, "BA58", DtlInpRip("  2G: "))
            PubVal(xlApp, "BA59", DtlInpRip("1.5G: "))
            PubVal(xlApp, "BA60", DtlInpRip("  1G: "))
            PubVal(xlApp, "BA61", DtlInpRip("0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' ガレージ外周GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300Gar(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA72", DtlInpRip("  4G: "))
            PubVal(xlApp, "BA73", DtlInpRip("3.5G: "))
            PubVal(xlApp, "BA74", DtlInpRip("  3G: "))
            PubVal(xlApp, "BA75", DtlInpRip("2.5G: "))
            PubVal(xlApp, "BA76", DtlInpRip("  2G: "))
            PubVal(xlApp, "BA77", DtlInpRip("1.5G: "))
            PubVal(xlApp, "BA78", DtlInpRip("  1G: "))
            PubVal(xlApp, "BA79", DtlInpRip("0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブユニット.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub SlabUnit(xlApp As Application)
        PubVal(xlApp, "BA99", DtlInpRip("  4G: "))
        PubVal(xlApp, "BA100", DtlInpRip("3.5G: "))
        PubVal(xlApp, "BA101", DtlInpRip("  3G: "))
        PubVal(xlApp, "BA102", DtlInpRip("2.5G: "))
        PubVal(xlApp, "BA103", DtlInpRip("  2G: "))
        PubVal(xlApp, "BA104", DtlInpRip("1.5G: "))
        PubVal(xlApp, "BA105", DtlInpRip("  1G: "))
    End Sub

    ''' <summary>
    ''' 電気温水器.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="value">Value.</param>
    Friend Sub ElecWtrHtr(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA107", value)
        Else
            ClrVal(xlApp, "BA107")
        End If
    End Sub

    ''' <summary>
    ''' コーナー.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub JtCor(xlApp As Application)
        PubVal(xlApp, "BA165", DtlInpRip("D16: "))
        PubVal(xlApp, "BA164", DtlInpRip("D13: "))
        PubVal(xlApp, "BA163", DtlInpRip("D10: "))
    End Sub

    ''' <summary>
    ''' ストレート.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub JtStr(xlApp As Application)
        PubVal(xlApp, "BA162", DtlInpRip("D16: "))
        PubVal(xlApp, "BA161", DtlInpRip("D13: "))
        PubVal(xlApp, "BA160", DtlInpRip("D10: "))
    End Sub

    ''' <summary>
    ''' キャップタイヤ (320).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub CapTire(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA188", DtlInpRip("D16: "))
            PubVal(xlApp, "BA181", DtlInpRip("D10: "))
        End If
    End Sub

    ''' <summary>
    ''' ロングコーナー (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub LongCor(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA179", DtlInpRip(" 750×2250: "))
            PubVal(xlApp, "BA177", DtlInpRip(" 750×1750: "))
            PubVal(xlApp, "BA178", DtlInpRip(" 750×1250: "))
            PubModVal(xlApp, "167", "1250×1250", 4.1, DtlInpRipDesc("1250×1250 ", "[4.1]: "))
        End If
    End Sub

    ''' <summary>
    ''' クランク.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA171", DtlInpRip("D16 (750×920×750): "))
            PubVal(xlApp, "BA172", DtlInpRip("D10 (500×920×500): "))
            PubVal(xlApp, "BA173", DtlInpRip("D16 (750×460×750): "))
            PubVal(xlApp, "BA174", DtlInpRip("D10 (500×460×500): "))
        End If
    End Sub

    ''' <summary>
    ''' 島 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Island(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA176", DtlInpRip("350×930×350: "))
            PubVal(xlApp, "BA175", DtlInpRip("350×470×350: "))
        End If
    End Sub

    ''' <summary>
    ''' ストレート (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Straight(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubModValExp(xlApp, "169", "4000", 6.3, My.Settings.Pr_D16, DtlInpRipDesc("4000 ", $"[6.3({My.Settings.Pr_D16})]: "))
            PubModValExp(xlApp, "170", "3500", 5.5, My.Settings.Pr_D16, DtlInpRipDesc("3500 ", $"[5.5({My.Settings.Pr_D16})]: "))
            PubVal(xlApp, "BA182", DtlInpRip("3000: "))
            PubVal(xlApp, "BA183", DtlInpRip("2500: "))
            PubVal(xlApp, "BA184", DtlInpRip("2000: "))
        End If
    End Sub

    ''' <summary>
    ''' コーナー3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Corner3(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA166", DtlInpRip("右(750×460×350): "))
            PubVal(xlApp, "BA168", DtlInpRip("左(750×460×350): "))
            PubModVal(xlApp, "187", "750×240×350", 2.2, DtlInpRipDesc("右(750×240×350) ", "[2.2]: "))
            PubModVal(xlApp, "190", "750×240×350", 2.2, DtlInpRipDesc("左(750×240×350) ", "[2.2]: "))
        End If
    End Sub

    ''' <summary>
    ''' クランク3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank3(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubModValFull(xlApp, "189", "（クランク３右）", "750×460×460×350", 3.3, DtlInpRipDesc("右(750×460×460×350) ", "[3.3]: "))
            PubModValFull(xlApp, "188", "（クランク３左）", "750×460×460×350", 3.3, DtlInpRipDesc("左(750×460×460×350) ", "[3.3]: "))
        End If
    End Sub

    ''' <summary>
    ''' コ型3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub UType3(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubModValFull(xlApp, "191", "（コノ字３右）", "750×460×460×350", 3.3, DtlInpRipDesc("右(750×460×460×350) ", "[3.3]: "))
            PubModValFull(xlApp, "196", "（コノ字３左）", "750×460×460×350", 3.3, DtlInpRipDesc("左(750×460×460×350) ", "[3.3]: "))
        End If
    End Sub

    ''' <summary>
    ''' フック (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Hook(xlApp As Application)
        PubModVal(xlApp, "192", "695×160　　フック付", 0.6, DtlInpRipDesc("695×160 ", "[0.6]: "))
        PubModVal(xlApp, "193", "595×160　　フック付", 0.5, DtlInpRipDesc("595×160 ", "[0.5]: "))
        PubModVal(xlApp, "194", "160×160　　フック付", 0.3, DtlInpRipDesc("160×160 ", "[0.3]: "))
        PubVal(xlApp, "BA185", DtlInpRip("435×250: "))
    End Sub

    ''' <summary>
    ''' 主筋補強 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub MainReinf(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA202", DtlInpRip("2500: "))
            PubVal(xlApp, "BA203", DtlInpRip("2000: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ曲 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub SlabBndg(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA115", DtlInpRip("250×5250: "))
            PubVal(xlApp, "BA116", DtlInpRip("250×4750: "))
            PubVal(xlApp, "BA117", DtlInpRip("250×4250: "))
            PubVal(xlApp, "BA118", DtlInpRip("250×3750: "))
            PubVal(xlApp, "BA119", DtlInpRip("250×3250: "))
            PubVal(xlApp, "BA120", DtlInpRip("250×2750: "))
            PubVal(xlApp, "BA121", DtlInpRip("250×2250: "))
            PubVal(xlApp, "BA122", DtlInpRip("250×1750: "))
            PubVal(xlApp, "BA123", DtlInpRip("250×1250: "))
            PubVal(xlApp, "BA124", DtlInpRip("250× 750: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ直 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub SlabStr(xlApp As Application)
        PubVal(xlApp, "BA125", DtlInpRip("5500: "))
        PubVal(xlApp, "BA126", DtlInpRip("5000: "))
        PubVal(xlApp, "BA127", DtlInpRip("4500: "))
        PubVal(xlApp, "BA128", DtlInpRip("4000: "))
        PubVal(xlApp, "BA129", DtlInpRip("3500: "))
        PubVal(xlApp, "BA130", DtlInpRip("3000: "))
        PubVal(xlApp, "BA131", DtlInpRip("2500: "))
        PubVal(xlApp, "BA132", DtlInpRip("2000: "))
        PubVal(xlApp, "BA133", DtlInpRip("1500: "))
        PubVal(xlApp, "BA134", DtlInpRip("1200: "))
        PubVal(xlApp, "BA135", DtlInpRip("1000: "))
        PubVal(xlApp, "BA136", DtlInpRip(" 900: "))
    End Sub

    ''' <summary>
    ''' スラブ補強曲 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub SlabReinfBndg(xlApp As Application, choosen As Integer)
        If choosen = 1 Then
            PubVal(xlApp, "BA137", DtlInpRip("250×5250: "))
            PubVal(xlApp, "BA138", DtlInpRip("250×4750: "))
            PubVal(xlApp, "BA139", DtlInpRip("250×4250: "))
            PubVal(xlApp, "BA140", DtlInpRip("250×3750: "))
            PubVal(xlApp, "BA141", DtlInpRip("250×3250: "))
            PubVal(xlApp, "BA142", DtlInpRip("250×2750: "))
            PubVal(xlApp, "BA143", DtlInpRip("250×2250: "))
            PubVal(xlApp, "BA144", DtlInpRip("250×1750: "))
            PubVal(xlApp, "BA145", DtlInpRip("250×1250: "))
            PubVal(xlApp, "BA146", DtlInpRip("250× 750: "))
        End If
    End Sub
End Module
