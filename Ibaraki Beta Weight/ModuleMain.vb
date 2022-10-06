Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Console
Imports System.ConsoleColor

Public Module ModuleMain
    Public Sub Main()
        ChkUpd()
        KillPrcs(XL_NAME)
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim ofd As New OpenFileDialog With {
            .Multiselect = False,
            .Title = "Open Excel Document",
            .Filter = "Excel Document|*.xlsx;*.xls"
        }
        ofd.ShowDialog()
        ' TODO

        Dim filePath = ofd.FileName
        xlApp.Workbooks.Open(filePath)

        'Dim temp As New UShort
        Dim t As String

        ' Fare
        Fare(xlApp, YNQ("運賃(2トン車): "))
        ' Unit GL-150
        PrefWarn("外周/内周GL-150")
        Unit150(xlApp)
        ' Unit GL-300
        Unit300(xlApp, YNQ("外周深GL-300: "))
        ' Unit GL-300/+30
        Unit300Cut(xlApp, YNQ("外周深GL-300/+30: "))
        ' Unit GL-400
        Unit400(xlApp, YNQ("外周深GL-400: "))
        ' Unit GL-400/+30
        Unit400Cut(xlApp, YNQ("外周深GL-400/+30: "))
        ' Unit Garage GL-300
        Unit300Gar(xlApp, YNQ("ガレージ外周GL-300: "))
        ' Slab unit
        PrefWarn("スラブユニット")
        SlabUnit(xlApp)
        ' Electric water heater
        ElecWtrHtr(xlApp, HdrInp("電気温水器: "))
        ' Corner joint
        PrefWarn("コーナー")
        JtCor(xlApp)
        ' Straight joint
        PrefWarn("ストレート")
        JtStr(xlApp)
        ' Cap tire
        CapTire(xlApp, YNQ("キャップタイヤ (320): "))
        ' Edge
        PubVal(xlApp, "BA180", HdrInp("端部(700×350): "))
        ' Long corner
        LongCor(xlApp, YNQ("ロングコーナー (D16): "))
        ' Crank
        Crank(xlApp, YNQ("クランク: "))
        ' Island
        Island(xlApp, YNQ("島 (D16): "))
        ' Straight
        Straight(xlApp, YNQ("ストレート (D16): "))
        ' Haunch
        PubVal(xlApp, "BA180", HdrInp("ハンチ (D16[660×410×660]): "))
        ' Corner 3
        Corner3(xlApp, YNQ("コーナー3 (D16): "))
        ' Crank 3
        Crank3(xlApp, YNQ("クランク3 (D16): "))
        ' U type 3
        UType3(xlApp, YNQ("コ型3 (D16): "))
        ' M type
        PubModVal(xlApp, "195", "350×460×460×350", 2.7, HdrInpDesc("M型 (D16[350×460×460×350]) ", "[2.7]: "))
        ' hook
        PrefWarn("フック (D10)")
        Hook(xlApp)
        ' Main reinforcement
        MainReinf(xlApp, YNQ("主筋補強 (D10): "))
        ' Bending
        SlabBndg(xlApp, YNQ("スラブ曲 (D13): "))
        ' Slab straight
        PrefWarn("スラブ直 (D13)")
        SlabStr(xlApp)
        ' Slab reinforcement bending
        SlabReinfBndg(xlApp, YNQ("スラブ補強曲 (D10): "))

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "スラブ補強直 (D10): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "スラブ補強直 (D10): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "5500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA147").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "5000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA148").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "4500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA149").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "4000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA150").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "3500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA151").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "3000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA152").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "2500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA153").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "2000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA154").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "1500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA155").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "1000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA156").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If
        xlApp.Range("BA157").Activate()
        xlApp.ActiveCell.FormulaR1C1 = 2
        xlApp.Range("BA158").Activate()
        xlApp.ActiveCell.FormulaR1C1 = 3
        xlApp.Range("BA159").Activate()
        xlApp.ActiveCell.FormulaR1C1 = 3

        Info()
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & vbTab & "スリーブ: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA197").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
            xlApp.Range("BA198").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
            xlApp.Range("BA199").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
            xlApp.Range("BA200").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp * 2
            xlApp.Range("BA201").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine(vbTab & vbTab & "副資材リスト")
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "邸名:" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        t = Console.ReadLine.ToString + "様邸"
        xlApp.Range("BJ12").Activate()
        xlApp.ActiveCell.FormulaR1C1 = t
        CType(xlApp.ActiveSheet, Worksheet).Name = t
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "住所:" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        t = Console.ReadLine.ToString
        xlApp.Range("BJ13").Activate()
        xlApp.ActiveCell.FormulaR1C1 = t
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "邸名コード:" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        t = Console.ReadLine.ToString
        xlApp.Range("AD5").Activate()
        xlApp.ActiveCell.FormulaR1C1 = t
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "納品日:" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        t = Console.ReadLine.ToString
        xlApp.Range("BO2").Activate()
        xlApp.ActiveCell.FormulaR1C1 = t
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & "運賃(分納):" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & "運賃(分納):" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab)
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            xlApp.Range("BA242").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "フラットアンカーボルト (本)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[M12×350]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA214").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "カットスクリュー・Ⅱ (袋)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[M12用]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA215").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "カットスクリュー・Ⅱ専用ピット (個):" & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA216").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "ホールダウンアンカーボルト (本)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[M12×700]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA217").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "アンカーグリッパーM12用 (箱)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[D10 TG1210D]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA218").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "アンカーグリッパーM12用 (箱)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[D13 TG1213D]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA219").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "アンカーグリッパーM12用 (箱)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[D16 TG1216D]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA220").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "マグネット差し筋アンカー (セット)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & "[直]:" & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA237").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "マグネット差し筋アンカー (セット)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & "[曲]:" & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA236").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "スペーサーブロック (個)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & vbTab & "[60ﾐﾘ]:" & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA221").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "スペーサーブロック (個)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & vbTab & "[80ﾐﾘ]:" & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA222").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "スペーサーブロック (個)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & vbTab & "[60×70×80]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA223").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "排水用スリーブホルダー D10用 (袋)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & "[50Φ・75Φ用]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA225").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "給水用スリーブホルダー D10用 (袋)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & "[50Φ]:" & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA226").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "養生シート輪木 (セット)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & vbTab & "[3.6×5.4]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA227").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        Else
            xlApp.Range("BA227").Activate()
            xlApp.ActiveCell.FormulaR1C1 = 1
            xlApp.Range("BF227").Activate()
            xlApp.ActiveCell.MergeArea.ClearContents()
            xlApp.Range("CB227").Activate()
            xlApp.ActiveCell.MergeArea.ClearContents()
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "M型鉄筋ベース (個):" & vbTab & vbTab & vbTab & vbTab & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA228").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "樹脂スペーサー改 (ケース)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & "[300ヶ]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA229").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "アンカーボルトセット M18 (セット)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & "[M18×380]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA232").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "NSP吊巾止 W160用 (本)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & vbTab & "[200本]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA234").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "アンカーボルト M16 (本)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & vbTab & vbTab & "[M16×415]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA238").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "ホールダウンアンカーボルト M12 (本)")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(vbTab & "[M12×498]:" & vbTab)
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA239").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        xlApp.ActiveWorkbook.Close(SaveChanges:=True)
        Process.Start(filePath)
    End Sub

End Module
