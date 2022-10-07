Imports Microsoft.Office.Interop.Excel

Friend Module Util
    ''' <summary>
    ''' 運賃(2トン車).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Fare(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            DctVal(xlApp, "BA243", choosen)
        End If
    End Sub

    ''' <summary>
    ''' 外周/内周GL-150.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Unit150(xlApp As Application)
        PubDVal(xlApp, "BA18", DtlDInp(vbTab & "  4G: "))
        PubDVal(xlApp, "BA19", DtlDInp(vbTab & "3.5G: "))
        PubDVal(xlApp, "BA20", DtlDInp(vbTab & "  3G: "))
        PubDVal(xlApp, "BA21", DtlDInp(vbTab & "2.5G: "))
        PubDVal(xlApp, "BA22", DtlDInp(vbTab & "  2G: "))
        PubDVal(xlApp, "BA23", DtlDInp(vbTab & "1.5G: "))
        PubDVal(xlApp, "BA24", DtlDInp(vbTab & "  1G: "))
        PubDVal(xlApp, "BA25", DtlDInp(vbTab & "0.5G: "))
    End Sub

    ''' <summary>
    ''' 外周深GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA27", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA28", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA29", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA30", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA31", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA32", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA33", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA34", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-300/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300Cut(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA36", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA37", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA38", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA39", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA40", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA41", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA42", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA43", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-400.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit400(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA45", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA46", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA47", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA48", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA49", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA50", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA51", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA52", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-400/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit400Cut(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA54", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA55", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA56", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA57", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA58", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA59", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA60", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA61", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 玄関・勝手口.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub EntrBackDoor(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA63", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA64", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA65", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA66", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA67", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA68", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA69", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA70", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' ガレージ外周GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300Gar(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA72", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA73", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA74", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA75", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA76", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA77", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA78", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA79", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブユニット.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub SlabUnit(xlApp As Application)
        PubDVal(xlApp, "BA99", DtlDInp(vbTab & "  4G: "))
        PubDVal(xlApp, "BA100", DtlDInp(vbTab & "3.5G: "))
        PubDVal(xlApp, "BA101", DtlDInp(vbTab & "  3G: "))
        PubDVal(xlApp, "BA102", DtlDInp(vbTab & "2.5G: "))
        PubDVal(xlApp, "BA103", DtlDInp(vbTab & "  2G: "))
        PubDVal(xlApp, "BA104", DtlDInp(vbTab & "1.5G: "))
        PubDVal(xlApp, "BA105", DtlDInp(vbTab & "  1G: "))
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
        PubDVal(xlApp, "BA165", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA164", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA163", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' ストレート.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub JtStr(xlApp As Application)
        PubDVal(xlApp, "BA162", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA161", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA160", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' キャップタイヤ (320).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub CapTire(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA181", DtlDInp(vbTab & "D10: "))
        End If
    End Sub

    ''' <summary>
    ''' ロングコーナー (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub LongCor(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA178", DtlDInp(vbTab & " 750×1250" & vbTab & ": "))
            PubDVal(xlApp, "BA179", DtlDInp(vbTab & " 750×2250" & vbTab & ": "))
            PubDVal(xlApp, "BA177", DtlDInp(vbTab & " 750×1750" & vbTab & ": "))
            PubDModVal(xlApp, "167", "1250×1250", 4.1, DtlDInpDesc(vbTab & "1250×1250 ", "[4.1]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' クランク.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA171", DtlDInp(vbTab & "D16 (750×920×750): "))
            PubDVal(xlApp, "BA172", DtlDInp(vbTab & "D10 (500×920×500): "))
            PubDVal(xlApp, "BA173", DtlDInp(vbTab & "D16 (750×460×750): "))
            PubDVal(xlApp, "BA174", DtlDInp(vbTab & "D10 (500×460×500): "))
        End If
    End Sub

    ''' <summary>
    ''' 島 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Island(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA176", DtlDInp(vbTab & "350×930×350: "))
            PubDVal(xlApp, "BA175", DtlDInp(vbTab & "350×470×350: "))
        End If
    End Sub

    ''' <summary>
    ''' ストレート (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Straight(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "169", "4000", 6.3, My.Settings.Pr_D16, DtlDInpDesc(vbTab & "4000 ", $"[6.3({My.Settings.Pr_D16})]" & vbTab))
            PubDModVal(xlApp, "170", "3500", 5.5, My.Settings.Pr_D16, DtlDInpDesc(vbTab & "3500 ", $"[5.5({My.Settings.Pr_D16})]" & vbTab))
            PubDVal(xlApp, "BA182", DtlDInp(vbTab & "3000" & vbTab & vbTab & ": "))
            PubDVal(xlApp, "BA183", DtlDInp(vbTab & "2500" & vbTab & vbTab & ": "))
            PubDVal(xlApp, "BA184", DtlDInp(vbTab & "2000" & vbTab & vbTab & ": "))
        End If
    End Sub

    ''' <summary>
    ''' コーナー3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Corner3(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA166", DtlDInp(vbTab & "右(750×460×350)" & vbTab & vbTab & ": "))
            PubDVal(xlApp, "BA168", DtlDInp(vbTab & "左(750×460×350)" & vbTab & vbTab & ": "))
            PubDModVal(xlApp, "187", "750×240×350", 2.2, DtlDInpDesc(vbTab & "右(750×240×350) ", "[2.2]" & vbTab))
            PubDModVal(xlApp, "190", "750×240×350", 2.2, DtlDInpDesc(vbTab & "左(750×240×350) ", "[2.2]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' クランク3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank3(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "189", "（クランク３右）", "750×460×460×350", 3.3, DtlDInpDesc(vbTab & "右(750×460×460×350) ", "[3.3]"))
            PubDModVal(xlApp, "188", "（クランク３左）", "750×460×460×350", 3.3, DtlDInpDesc(vbTab & "左(750×460×460×350) ", "[3.3]"))
        End If
    End Sub

    ''' <summary>
    ''' コ型3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub UType3(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "191", "（コノ字３右）", "750×460×460×350", 3.3, DtlDInpDesc(vbTab & "右(750×460×460×350) ", "[3.3]"))
            PubDModVal(xlApp, "196", "（コノ字３左）", "750×460×460×350", 3.3, DtlDInpDesc(vbTab & "左(750×460×460×350) ", "[3.3]"))
        End If
    End Sub

    ''' <summary>
    ''' フック (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Hook(xlApp As Application)
        PubDModVal(xlApp, "192", "695×160　　フック付", 0.6, DtlDInpDesc(vbTab & "695×160 ", "[0.6]" & vbTab))
        PubDModVal(xlApp, "193", "595×160　　フック付", 0.5, DtlDInpDesc(vbTab & "595×160 ", "[0.5]" & vbTab))
        PubDModVal(xlApp, "194", "160×160　　フック付", 0.3, DtlDInpDesc(vbTab & "160×160 ", "[0.3]" & vbTab))
        PubDVal(xlApp, "BA185", DtlDInp(vbTab & "435×250" & vbTab & vbTab & ": "))
    End Sub

    ''' <summary>
    ''' 主筋補強 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub MainReinf(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA202", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA203", DtlDInp(vbTab & "2000: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ曲 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub SlabBndg(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA115", DtlDInp(vbTab & "250×5250: "))
            PubDVal(xlApp, "BA116", DtlDInp(vbTab & "250×4750: "))
            PubDVal(xlApp, "BA117", DtlDInp(vbTab & "250×4250: "))
            PubDVal(xlApp, "BA118", DtlDInp(vbTab & "250×3750: "))
            PubDVal(xlApp, "BA119", DtlDInp(vbTab & "250×3250: "))
            PubDVal(xlApp, "BA120", DtlDInp(vbTab & "250×2750: "))
            PubDVal(xlApp, "BA121", DtlDInp(vbTab & "250×2250: "))
            PubDVal(xlApp, "BA122", DtlDInp(vbTab & "250×1750: "))
            PubDVal(xlApp, "BA123", DtlDInp(vbTab & "250×1250: "))
            PubDVal(xlApp, "BA124", DtlDInp(vbTab & "250× 750: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ直 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub SlabStr(xlApp As Application)
        PubDVal(xlApp, "BA125", DtlDInp(vbTab & "5500: "))
        PubDVal(xlApp, "BA126", DtlDInp(vbTab & "5000: "))
        PubDVal(xlApp, "BA127", DtlDInp(vbTab & "4500: "))
        PubDVal(xlApp, "BA128", DtlDInp(vbTab & "4000: "))
        PubDVal(xlApp, "BA129", DtlDInp(vbTab & "3500: "))
        PubDVal(xlApp, "BA130", DtlDInp(vbTab & "3000: "))
        PubDVal(xlApp, "BA131", DtlDInp(vbTab & "2500: "))
        PubDVal(xlApp, "BA132", DtlDInp(vbTab & "2000: "))
        PubDVal(xlApp, "BA133", DtlDInp(vbTab & "1500: "))
        PubDVal(xlApp, "BA134", DtlDInp(vbTab & "1200: "))
        PubDVal(xlApp, "BA135", DtlDInp(vbTab & "1000: "))
        PubDVal(xlApp, "BA136", DtlDInp(vbTab & " 900: "))
    End Sub

    ''' <summary>
    ''' スラブ補強曲 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub SlabReinfBndg(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA137", DtlDInp(vbTab & "250×5250: "))
            PubDVal(xlApp, "BA138", DtlDInp(vbTab & "250×4750: "))
            PubDVal(xlApp, "BA139", DtlDInp(vbTab & "250×4250: "))
            PubDVal(xlApp, "BA140", DtlDInp(vbTab & "250×3750: "))
            PubDVal(xlApp, "BA141", DtlDInp(vbTab & "250×3250: "))
            PubDVal(xlApp, "BA142", DtlDInp(vbTab & "250×2750: "))
            PubDVal(xlApp, "BA143", DtlDInp(vbTab & "250×2250: "))
            PubDVal(xlApp, "BA144", DtlDInp(vbTab & "250×1750: "))
            PubDVal(xlApp, "BA145", DtlDInp(vbTab & "250×1250: "))
            PubDVal(xlApp, "BA146", DtlDInp(vbTab & "250× 750: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ補強直 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub SlabReinfStr(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA147", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA148", DtlDInp(vbTab & "5000: "))
            PubDVal(xlApp, "BA149", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "BA150", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "BA151", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA152", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA153", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA154", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "BA155", DtlDInp(vbTab & "1500: "))
            PubDVal(xlApp, "BA156", DtlDInp(vbTab & "1000: "))
        End If
        DctVal(xlApp, "BA157", 2)
        DctVal(xlApp, "BA158", 3)
        DctVal(xlApp, "BA159", 3)
    End Sub

    ''' <summary>
    ''' スリーブ.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="value">Value.</param>
    Friend Sub Sleeve(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA198", value)
            DctVal(xlApp, "BA197", value)
            DctVal(xlApp, "BA199", value)
            DctVal(xlApp, "BA200", value * 2)
            DctVal(xlApp, "BA201", value)
        End If
    End Sub

    ''' <summary>
    ''' 副資材リスト.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Parts(xlApp As Application)
        Dim name = $"{DtlSInp(vbTab & "邸名" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")}様邸"
        DctVal(xlApp, "BJ12", name)
        CType(xlApp.ActiveSheet, Worksheet).Name = name
        PubSVal(xlApp, vbTab & "住所" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BJ13")
        PubSVal(xlApp, vbTab & "邸名コード" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "AD5")
        PubSVal(xlApp, vbTab & "納品日" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BO2")
        Dim ipp = DtlYNQ(vbTab & "運賃(分納)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")
        If ipp = 1 Then
            DctVal(xlApp, "BA244", ipp)
        End If
        PubDVal(xlApp, "BA214", DtlDInpDesc(vbTab & "フラットアンカーボルト (本)", vbTab & vbTab & "[M12×350]" & vbTab))
        PubDVal(xlApp, "BA215", DtlDInpDesc(vbTab & "カットスクリュー・Ⅱ (袋)", vbTab & vbTab & "[M12用]" & vbTab & vbTab))
        PubDVal(xlApp, "BA216", DtlDInp(vbTab & "カットスクリュー・Ⅱ専用ピット (個)" & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA217", DtlDInpDesc(vbTab & "ホールダウンアンカーボルト (本)", vbTab & vbTab & "[M12×700]" & vbTab))
        PubDVal(xlApp, "BA218", DtlDInpDesc(vbTab & "アンカーグリッパーM12用 (箱)", vbTab & vbTab & "[D10 TG1210D]" & vbTab))
        PubDVal(xlApp, "BA219", DtlDInpDesc(vbTab & "アンカーグリッパーM12用 (箱)", vbTab & vbTab & "[D13 TG1213D]" & vbTab))
        PubDVal(xlApp, "BA220", DtlDInpDesc(vbTab & "アンカーグリッパーM12用 (箱)", vbTab & vbTab & "[D16 TG1216D]" & vbTab))
        PubDVal(xlApp, "BA237", DtlDInpDesc(vbTab & "マグネット差し筋アンカー (ｾｯﾄ)", vbTab & vbTab & "[直]" & vbTab & vbTab))
        PubDVal(xlApp, "BA236", DtlDInpDesc(vbTab & "マグネット差し筋アンカー (ｾｯﾄ)", vbTab & vbTab & "[曲]" & vbTab & vbTab))
        PubDVal(xlApp, "BA221", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[60ﾐﾘ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA222", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[80ﾐﾘ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA223", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[60×70×80]" & vbTab))
        PubDVal(xlApp, "BA225", DtlDInpDesc(vbTab & "排水用スリーブホルダー D10用 (袋)", vbTab & "[50Φ・75Φ用]" & vbTab))
        PubDVal(xlApp, "BA226", DtlDInpDesc(vbTab & "給水用スリーブホルダー D10用 (袋)", vbTab & "[50Φ]" & vbTab & vbTab))
        Dim curingShRingTree = DtlDInpDesc(vbTab & "養生シート輪木 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[3.6×5.4]" & vbTab)
        If curingShRingTree > 0 Then
            DctVal(xlApp, "BA227", curingShRingTree)
        Else
            DctVal(xlApp, "BA227", 1)
            ClrVal(xlApp, "BF227")
            ClrVal(xlApp, "CB227")
        End If
        PubDVal(xlApp, "BA228", DtlDInp(vbTab & "Ｍ型鉄筋ベース (個)" & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA229", DtlDInpDesc(vbTab & "樹脂スペーサー改 (ｹｰｽ)", vbTab & vbTab & vbTab & "[300ヶ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA232", DtlDInpDesc(vbTab & "アンカーボルトセット M18 (ｾｯﾄ)", vbTab & vbTab & "[M18×380]" & vbTab))
        PubDVal(xlApp, "BA234", DtlDInpDesc(vbTab & "NSP吊巾止 W160用 (本)", vbTab & vbTab & vbTab & "[200本]" & vbTab & vbTab))
        PubDVal(xlApp, "BA238", DtlDInpDesc(vbTab & "アンカーボルト M16 (本)", vbTab & vbTab & vbTab & "[M16×415]" & vbTab))
        PubDVal(xlApp, "BA239", DtlDInpDesc(vbTab & "ホールダウンアンカーボルト M12 (本)", vbTab & "[M12×498]" & vbTab))
        ' Extend
        PubDVal(xlApp, "BA224", DtlDInpDesc(vbTab & "樹脂スペーサー (個)", vbTab & vbTab & vbTab & "[70×80]" & vbTab & vbTab))
        PubDVal(xlApp, "BA230", DtlDInpDesc(vbTab & "鉄筋スペーサー60用 (個)", vbTab & vbTab & vbTab & "[60ﾖｳ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA231", DtlDInpDesc(vbTab & "鉄筋スペーサー80用 (個)", vbTab & vbTab & vbTab & "[80ﾖｳ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA233", DtlDInpDesc(vbTab & "偏心用鉄筋ベース (個)", vbTab & vbTab & vbTab & "[280×160×60]" & vbTab))
        PubDVal(xlApp, "BA235", DtlDInpDesc(vbTab & "防錆巾止め金具 (本)", vbTab & vbTab & vbTab & "[Fﾊﾟﾈﾙ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA240", DtlDInpDesc(vbTab & "アンカーボルトセットM12 (本)", vbTab & vbTab & "[M12×498]" & vbTab))
        PubDVal(xlApp, "BA241", DtlDInpDesc(vbTab & "アンカーボルトセットM12軸柱用 (本)", vbTab & "[M12×498]" & vbTab))
        PubDVal(xlApp, "BA242", DtlDInpDesc(vbTab & "Ｕボルト (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & "[M8]" & vbTab & vbTab))
    End Sub
End Module
