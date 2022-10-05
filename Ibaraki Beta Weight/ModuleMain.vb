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

        ' 2t
        Dim is2t = YSQ("運賃(2トン車): ")
        If is2t = 1 Then
            DctVal(xlApp, "BA241", is2t)
        End If
        ' GL-150
        PrefWarn("外周/内周GL-150")
        GL50(xlApp)
        ' GL-300
        If YSQ("外周深GL-300: ") = 1 Then
            GL300(xlApp)
        End If
        ' GL-300/+30
        If YSQ("外周深GL-300/+30: ") = 1 Then
            GL300Cut(xlApp)
        End If
        ' GL-400
        If YSQ("外周深GL-400: ") = 1 Then
            GL400(xlApp)
        End If
        ' GL-400/+30
        If YSQ("外周深GL-400/+30: ") = 1 Then
            GL400Cut(xlApp)
        End If
        ' Garage GL-300
        If YSQ("ガレージ外周GL-300: ") = 1 Then
            GL300Gar(xlApp)
        End If
        ' Slap unit
        PrefWarn("スラブユニット")
        SlapUnit(xlApp)
        ' Electric water
        PrefImp("電気温水器: ")
        Dim elecWtr = Val(ReadLine)
        If elecWtr > 0 Then
            DctVal(xlApp, "BA107", elecWtr)
        Else
            ClrVal(xlApp, "BA107")
        End If
        ' Corner
        PrefWarn("コーナー")
        Corner(xlApp)
        ' Straight
        PrefWarn("ストレート")
        Straight(xlApp)
        ' Cap tire
        If YSQ("キャップタイヤ (320): ") = 1 Then
            CapTire(xlApp)
        End If
        ' Edge
        ImpVal(xlApp, "端部(700×350): ", "BA180")
        ' Long corner
        If YSQ("ロングコーナー (D16): ") = 1 Then
            LongCorner(xlApp)
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "クランク: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "クランク: ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "D16 (750×920×750): ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA171").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "D10 (500×920×500): ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA172").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "D16 (750×460×750): ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA173").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "D10 (500×460×500): ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA174").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "島 (D16): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "島 (D16): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "350×930×350: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA176").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "350×470×350: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA175").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "ストレート (D16): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "ストレート (D16): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "4000 ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[6.3(66)]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("AH169").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "4000"
                xlApp.Range("CM169").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 6.3
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("CQ169").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 66
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA169").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "3500 ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[5.5(66)]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("AH170").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "3500"
                xlApp.Range("CM170").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 5.5
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("CQ170").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 66
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA170").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "3000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA182").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "2500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA183").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "2000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA184").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & vbTab & "ハンチ (D16[660×410×660]): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA186").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "コーナー3 (D16): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "コーナー3 (D16): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "右(750×460×350): ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA166").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "左(750×460×350): ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA168").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "右(750×240×350) ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[2.2]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("AH187").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "750×240×350"
                xlApp.Range("CM187").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 2.2
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA187").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "左(750×240×350) ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[2.2]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("AH190").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "750×240×350"
                xlApp.Range("CM190").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 2.2
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA190").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "クランク3 (D16): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "クランク3 (D16): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "右(750×460×460×350) ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[3.3]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("AH189").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "750×460×460×350"
                xlApp.Range("CM189").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 3.3
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA189").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "左(750×460×460×350) ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[3.3]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("X188").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "（クランク３左）"
                xlApp.Range("AH188").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "750×460×460×350"
                xlApp.Range("CM188").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 3.3
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA188").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If

        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "コ型3 (D16): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "コ型3 (D16): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "右(750×460×460×350) ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[3.3]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("X191").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "（コノ字３右）"
                xlApp.Range("AH191").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "750×460×460×350"
                xlApp.Range("CM191").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 3.3
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA191").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "左(750×460×460×350) ")
            Console.ForegroundColor = ConsoleColor.Magenta
            Console.Write("[3.3]: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("X196").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "（コノ字３左）"
                xlApp.Range("AH196").Activate()
                xlApp.ActiveCell.FormulaR1C1 = "750×460×460×350"
                xlApp.Range("CM196").Activate()
                xlApp.ActiveCell.FormulaR1C1 = 3.3
                xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
                xlApp.Range("BA196").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If
        Info()
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & vbTab & "M型 (D16[350×460×460×350]) ")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write("[2.7]: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("AH195").Activate()
            xlApp.ActiveCell.FormulaR1C1 = "350×460×460×350"
            xlApp.Range("CM195").Activate()
            xlApp.ActiveCell.FormulaR1C1 = 2.7
            xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
            xlApp.Range("BA195").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine(vbTab & vbTab & "フック (D10)")
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "695×160 ")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write("[0.6]: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("AH192").Activate()
            xlApp.ActiveCell.FormulaR1C1 = "695×160　　フック付"
            xlApp.Range("CM192").Activate()
            xlApp.ActiveCell.FormulaR1C1 = 0.6
            xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
            xlApp.Range("BA192").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "595×160 ")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write("[0.5]: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("AH193").Activate()
            xlApp.ActiveCell.FormulaR1C1 = "595×160　　フック付"
            xlApp.Range("CM193").Activate()
            xlApp.ActiveCell.FormulaR1C1 = 0.5
            xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
            xlApp.Range("BA193").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "160×160 ")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write("[0.3]: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("AH194").Activate()
            xlApp.ActiveCell.FormulaR1C1 = "160×160　　フック付"
            xlApp.Range("CM194").Activate()
            xlApp.ActiveCell.FormulaR1C1 = 0.3
            xlApp.ActiveCell.Interior.Color = RGB(0, 176, 240)
            xlApp.Range("BA194").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "435×250: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA185").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "主筋補強 (D10): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "主筋補強 (D10): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "2500: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA202").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "2000: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA203").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If
        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "スラブ曲 (D13): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "スラブ曲 (D13): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×5250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA115").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×4750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA116").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×4250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA117").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×3750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA118").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×3250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA119").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×2750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA120").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×2250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA121").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×1750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA122").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×1250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA123").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250× 750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA124").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If
        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine(vbTab & vbTab & "スラブ直 (D13)")
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "5500: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA125").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "5000: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA126").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "4500: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA127").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "4000: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA128").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "3500: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA129").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "3000: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA130").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "2500: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA131").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "2000: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA132").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "1500: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA133").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "1200: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA134").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & "1000: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA135").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write(vbTab & " 900: ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp > 0 Then
            xlApp.Range("BA136").Activate()
            xlApp.ActiveCell.FormulaR1C1 = temp
        End If
        Info()
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.Write(vbTab & vbTab & "スラブ補強曲 (D10): ")
        Console.ForegroundColor = ConsoleColor.White
        temp = Val(Console.ReadLine)
        If temp <> 0 Or temp <> 1 Then
            Do Until temp = 0 Or temp = 1
                Info()
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(vbTab & vbTab & "スラブ補強曲 (D10): ")
                Console.ForegroundColor = ConsoleColor.Red
                temp = Val(Console.ReadLine)
            Loop
        End If
        If temp = 1 Then
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×5250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA137").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×4750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA138").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×4250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA139").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×3750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA140").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×3250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA141").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×2750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA142").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×2250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA143").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×1750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA144").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250×1250: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA145").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.Write(vbTab & "250× 750: ")
            Console.ForegroundColor = ConsoleColor.White
            temp = Val(Console.ReadLine)
            If temp > 0 Then
                xlApp.Range("BA146").Activate()
                xlApp.ActiveCell.FormulaR1C1 = temp
            End If
        End If
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
