Imports Microsoft.Office.Interop.Excel

Friend Module Service
    ''' <summary>
    ''' Weight Ibaraki Beta.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub WtIbarakiBeta(xlApp As Application)
        ' Fare
        Fare(xlApp, HdrYNQ(vbTab & vbTab & "運賃 (2トン車): "))
        ' Unit GL-150
        HdrWrng(vbTab & vbTab & "外周/内周GL-150" & vbCrLf)
        Unit150(xlApp)
        ' Unit GL-300
        Unit300(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-300: "))
        ' Unit GL-300/+30
        Unit300Cut(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-300/+30: "))
        ' Unit GL-400
        Unit400(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-400: "))
        ' Unit GL-400/+30
        Unit400Cut(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-400/+30: "))
        ' Entrance/Back door
        EntrBackDoor(xlApp, HdrYNQ(vbTab & vbTab & "玄関・勝手口: "))
        ' Unit Garage GL-300
        Unit300Gar(xlApp, HdrYNQ(vbTab & vbTab & "ガレージ外周GL-300: "))
        ' Slab unit
        HdrWrng(vbTab & vbTab & "スラブユニット" & vbCrLf)
        SlabUnit(xlApp)
        ' Electric water heater
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))
        ' Corner joint
        HdrWrng(vbTab & vbTab & "コーナー" & vbCrLf)
        JtCor(xlApp)
        ' Straight joint
        HdrWrng(vbTab & vbTab & "ストレート" & vbCrLf)
        JtStr(xlApp)
        ' Cap tire
        CapTire(xlApp, HdrYNQ(vbTab & vbTab & "キャップタイヤ (320): "))
        ' Edge
        PubDVal(xlApp, "BA180", HdrDInp(vbTab & vbTab & "端部(700×350): "))
        ' Long corner
        LongCor(xlApp, HdrYNQ(vbTab & vbTab & "ロングコーナー (D16): "))
        ' Crank
        Crank(xlApp, HdrYNQ(vbTab & vbTab & "クランク: "))
        ' Island
        Island(xlApp, HdrYNQ(vbTab & vbTab & "島 (D16): "))
        ' Straight
        Straight(xlApp, HdrYNQ(vbTab & vbTab & "ストレート (D16): "))
        ' Haunch
        PubDVal(xlApp, "BA180", HdrDInp(vbTab & vbTab & "ハンチ (D16[660×410×660]): "))
        ' Corner 3D
        Corner3d(xlApp, HdrYNQ(vbTab & vbTab & "コーナー3 (D16): "))
        ' Crank 3D
        Crank3d(xlApp, HdrYNQ(vbTab & vbTab & "クランク3 (D16): "))
        ' U type 3D
        UType3d(xlApp, HdrYNQ(vbTab & vbTab & "コ型3 (D16): "))
        ' M type
        PubDModVal(xlApp, "195", "350×460×460×350", 2.7, HdrDInp(vbTab & vbTab & "Ｍ型 (D16[350×460×460×350]): "))
        ' Hook
        HdrWrng(vbTab & vbTab & "フック (D10)" & vbCrLf)
        Hook(xlApp)
        ' Main reinforcement
        MainReinf(xlApp, HdrYNQ(vbTab & vbTab & "主筋補強 (D10): "))
        ' Slab bending
        SlabBndg(xlApp, HdrYNQ(vbTab & vbTab & "スラブ曲 (D13): "))
        ' Slab straight
        HdrWrng(vbTab & vbTab & "スラブ直 (D13)" & vbCrLf)
        SlabStr(xlApp)
        ' Slab reinforcement bending
        SlabReinfBndg(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強曲 (D10): "))
        ' Slab reinforcement straight
        SlabReinfStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強直 (D10): "))
        ' Sleeve
        Sleeve(xlApp, HdrDInp(vbTab & vbTab & "スリーブ: "))
        ' Parts
        HdrWrng(vbTab & vbTab & "副資材リスト" & vbCrLf)
        Parts(xlApp)
    End Sub
End Module
