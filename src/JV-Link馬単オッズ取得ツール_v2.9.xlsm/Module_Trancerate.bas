Attribute VB_Name = "Module_Trancerate"
Function JyoCord(cvt)

If cvt = "01" Then
    JyoCord = "嶥杫"
ElseIf cvt = "02" Then
    JyoCord = "敓娰"
ElseIf cvt = "03" Then
    JyoCord = "暉搰"
ElseIf cvt = "04" Then
    JyoCord = "怴妰"
ElseIf cvt = "05" Then
    JyoCord = "搶嫗"
ElseIf cvt = "06" Then
    JyoCord = "拞嶳"
ElseIf cvt = "07" Then
    JyoCord = "拞嫗"
ElseIf cvt = "08" Then
    JyoCord = "嫗搒"
ElseIf cvt = "09" Then
    JyoCord = "嶃恄"
ElseIf cvt = 10 Then
    JyoCord = "彫憅"
ElseIf cvt = 30 Then
    JyoCord = "栧暿"
ElseIf cvt = 31 Then
    JyoCord = "杒尒"
ElseIf cvt = 32 Then
    JyoCord = "娾尒戲"
ElseIf cvt = 33 Then
    JyoCord = "懷峀"
ElseIf cvt = 34 Then
    JyoCord = "埉愳"
ElseIf cvt = 35 Then
    JyoCord = "惙壀"
ElseIf cvt = 36 Then
    JyoCord = "悈戲"
ElseIf cvt = 37 Then
    JyoCord = "忋嶳"
ElseIf cvt = 38 Then
    JyoCord = "嶰忦"
ElseIf cvt = 39 Then
    JyoCord = "懌棙"
ElseIf cvt = 40 Then
    JyoCord = "塅搒媨"
ElseIf cvt = 41 Then
    JyoCord = "崅嶈"
ElseIf cvt = 42 Then
    JyoCord = "塝榓"
ElseIf cvt = 43 Then
    JyoCord = "慏嫶"
ElseIf cvt = 44 Then
    JyoCord = "戝堜"
ElseIf cvt = 45 Then
    JyoCord = "愳嶈"
ElseIf cvt = 46 Then
    JyoCord = "嬥戲"
ElseIf cvt = 47 Then
    JyoCord = "妢徏"
ElseIf cvt = 48 Then
    JyoCord = "柤屆壆"
ElseIf cvt = 49 Then
    JyoCord = "婭埳嶰堜帥"
ElseIf cvt = 50 Then
    JyoCord = "墍揷"
ElseIf cvt = 51 Then
    JyoCord = "昉楬"
ElseIf cvt = 52 Then
    JyoCord = "塿揷"
ElseIf cvt = 53 Then
    JyoCord = "暉嶳"
ElseIf cvt = 54 Then
    JyoCord = "崅抦"
ElseIf cvt = 55 Then
    JyoCord = "嵅夑"
ElseIf cvt = 56 Then
    JyoCord = "峳旜"
ElseIf cvt = 57 Then
    JyoCord = "拞捗"
ElseIf cvt = 58 Then
    JyoCord = "嶥杫(抧曽)"
ElseIf cvt = 59 Then
    JyoCord = "敓娰(抧曽)"
ElseIf cvt = 60 Then
    JyoCord = "怴妰(抧曽)"
ElseIf cvt = 61 Then
    JyoCord = "拞嫗(抧曽)"
End If

End Function

Function ShortJyo2Jyo(cvt)

If cvt = "嶥" Then
    ShortJyo2Jyo = "嶥杫"
ElseIf cvt = "敓" Then
    ShortJyo2Jyo = "敓娰"
ElseIf cvt = "暉" Then
    ShortJyo2Jyo = "暉搰"
ElseIf cvt = "怴" Then
    ShortJyo2Jyo = "怴妰"
ElseIf cvt = "搶" Then
    ShortJyo2Jyo = "搶嫗"
ElseIf cvt = "拞" Then
    ShortJyo2Jyo = "拞嶳"
ElseIf cvt = "柤" Then
    ShortJyo2Jyo = "拞嫗"
ElseIf cvt = "嫗" Then
    ShortJyo2Jyo = "嫗搒"
ElseIf cvt = "嶃" Then
    ShortJyo2Jyo = "嶃恄"
ElseIf cvt = "彫" Then
    ShortJyo2Jyo = "彫憅"
End If

End Function

Function Jyo2ShortJyo(cvt)

If cvt = "嶥杫" Then
    Jyo2ShortJyo = "嶥"
ElseIf cvt = "敓娰" Then
    Jyo2ShortJyo = "敓"
ElseIf cvt = "暉搰" Then
    Jyo2ShortJyo = "暉"
ElseIf cvt = "怴妰" Then
    Jyo2ShortJyo = "怴"
ElseIf cvt = "搶嫗" Then
    Jyo2ShortJyo = "搶"
ElseIf cvt = "拞嶳" Then
    Jyo2ShortJyo = "拞"
ElseIf cvt = "拞嫗" Then
    Jyo2ShortJyo = "柤"
ElseIf cvt = "嫗搒" Then
    Jyo2ShortJyo = "嫗"
ElseIf cvt = "嶃恄" Then
    Jyo2ShortJyo = "嶃"
ElseIf cvt = "彫憅" Then
    Jyo2ShortJyo = "彫"
End If

End Function

Function JyogyakuCord(cvt)

If cvt = "嶥杫" Then
    JyogyakuCord = "01"
ElseIf cvt = "敓娰" Then
    JyogyakuCord = "02"
ElseIf cvt = "暉搰" Then
    JyogyakuCord = "03"
ElseIf cvt = "怴妰" Then
    JyogyakuCord = "04"
ElseIf cvt = "搶嫗" Then
    JyogyakuCord = "05"
ElseIf cvt = "拞嶳" Then
    JyogyakuCord = "06"
ElseIf cvt = "拞嫗" Then
    JyogyakuCord = "07"
ElseIf cvt = "嫗搒" Then
    JyogyakuCord = "08"
ElseIf cvt = "嶃恄" Then
    JyogyakuCord = "09"
ElseIf cvt = "彫憅" Then
    JyogyakuCord = 10
ElseIf cvt = "栧暿" Then
    JyogyakuCord = 30
ElseIf cvt = "杒尒" Then
    JyogyakuCord = 31
ElseIf cvt = "娾尒戲" Then
    JyogyakuCord = 32
ElseIf cvt = "懷峀" Then
    JyogyakuCord = 33
ElseIf cvt = "埉愳" Then
    JyogyakuCord = 34
ElseIf cvt = "惙壀" Then
    JyogyakuCord = 35
ElseIf cvt = "悈戲" Then
    JyogyakuCord = 36
ElseIf cvt = "忋嶳" Then
    JyogyakuCord = 37
ElseIf cvt = "嶰忦" Then
    JyogyakuCord = 38
ElseIf cvt = "懌棙" Then
    JyogyakuCord = 39
ElseIf cvt = "塅搒媨" Then
    JyogyakuCord = 40
ElseIf cvt = "崅嶈" Then
    JyogyakuCord = 41
ElseIf cvt = "塝榓" Then
    JyogyakuCord = 42
ElseIf cvt = "慏嫶" Then
    JyogyakuCord = 43
ElseIf cvt = "戝堜" Then
    JyogyakuCord = 44
ElseIf cvt = "愳嶈" Then
    JyogyakuCord = 45
ElseIf cvt = "嬥戲" Then
    JyogyakuCord = 46
ElseIf cvt = "妢徏" Then
    JyogyakuCord = 47
ElseIf cvt = "柤屆壆" Then
    JyogyakuCord = 48
ElseIf cvt = "婭埳嶰堜帥" Then
    JyogyakuCord = 49
ElseIf cvt = "墍揷" Then
    JyogyakuCord = 50
ElseIf cvt = "昉楬" Then
    JyogyakuCord = 51
ElseIf cvt = "塿揷" Then
    JyogyakuCord = 52
ElseIf cvt = "暉嶳" Then
    JyogyakuCord = 53
ElseIf cvt = "崅抦" Then
    JyogyakuCord = 54
ElseIf cvt = "嵅夑" Then
    JyogyakuCord = 55
ElseIf cvt = "峳旜" Then
    JyogyakuCord = 56
ElseIf cvt = "拞捗" Then
    JyogyakuCord = 57
ElseIf cvt = "嶥杫(抧曽)" Then
    JyogyakuCord = 58
ElseIf cvt = "敓娰(抧曽)" Then
    JyogyakuCord = 59
ElseIf cvt = "怴妰(抧曽)" Then
    JyogyakuCord = 60
ElseIf cvt = "拞嫗(抧曽)" Then
    JyogyakuCord = 61
End If

End Function

Function TenkoCord(cvt)

If cvt = 0 Then
    TenkoCord = "枹愝掕"
ElseIf cvt = 1 Then
    TenkoCord = "惏"
ElseIf cvt = 2 Then
    TenkoCord = "撥"
ElseIf cvt = 3 Then
    TenkoCord = "塉"
ElseIf cvt = 4 Then
    TenkoCord = "彫塉"
ElseIf cvt = 5 Then
    TenkoCord = "愥"
ElseIf cvt = 6 Then
    TenkoCord = "彫愥"
End If

End Function


Function BabaCord(cvt)

If cvt = 0 Then
    BabaCord = "枹愝掕"
ElseIf cvt = 1 Then
    BabaCord = "椙"
ElseIf cvt = 2 Then
    BabaCord = "鈉廳"
ElseIf cvt = 3 Then
    BabaCord = "廳"
ElseIf cvt = 4 Then
    BabaCord = "晄椙"
End If

End Function


Function SeibetuCord(cvt)

If cvt = 0 Then
    SeibetuCord = "枹愝掕"
ElseIf cvt = 1 Then
    SeibetuCord = "壊攏"
ElseIf cvt = 2 Then
    SeibetuCord = "柲攏"
ElseIf cvt = 3 Then
    SeibetuCord = "僙儞攏"
End If

End Function


Function GradeCord(cvt)

If cvt = "A" Then
    GradeCord = "G1"
ElseIf cvt = "B" Then
    GradeCord = "G2"
ElseIf cvt = "C" Then
    GradeCord = "G3"
ElseIf cvt = "D" Then
    GradeCord = "廳徿"
ElseIf cvt = "E" Then
    GradeCord = "摿暿"
ElseIf cvt = "F" Then
    GradeCord = "J丒G1"
ElseIf cvt = "G" Then
    GradeCord = "J丒G2"
ElseIf cvt = "H" Then
    GradeCord = "J丒G3"
Else
    GradeCord = "-"
End If

End Function

Function YoubiCord(cvt)

If cvt = 0 Then
    YoubiCord = "枹愝掕"
ElseIf cvt = 1 Then
    YoubiCord = "搚"
ElseIf cvt = 2 Then
    YoubiCord = "擔"
ElseIf cvt = 3 Then
    YoubiCord = "廽"
ElseIf cvt = 4 Then
    YoubiCord = "寧"
ElseIf cvt = 5 Then
    YoubiCord = "壩"
ElseIf cvt = 6 Then
    YoubiCord = "悈"
ElseIf cvt = 7 Then
    YoubiCord = "栘"
ElseIf cvt = 8 Then
    YoubiCord = "嬥"
End If

End Function


Function KyososyubetuCord(cvt)

If cvt = "00" Then
    KyososyubetuCord = "枹愝掕"
ElseIf cvt = 11 Then
    KyososyubetuCord = "2嵨"
ElseIf cvt = 12 Then
    KyososyubetuCord = "3嵨"
ElseIf cvt = 13 Then
    KyososyubetuCord = "3嵨埲忋"
ElseIf cvt = 14 Then
    KyososyubetuCord = "4嵨埲忋"
ElseIf cvt = 18 Then
    KyososyubetuCord = "3嵨埲忋"
ElseIf cvt = 19 Then
    KyososyubetuCord = "4嵨埲忋"
ElseIf cvt = 21 Then
    KyososyubetuCord = "2嵨"
ElseIf cvt = 22 Then
    KyososyubetuCord = "3嵨"
ElseIf cvt = 23 Then
    KyososyubetuCord = "3嵨埲忋"
ElseIf cvt = 24 Then
    KyososyubetuCord = "4嵨埲忋"
End If


End Function




Function KyosoKigoCord(cvt)


If cvt = "000" Then
    KyosoKigoCord = "枹愝掕"
ElseIf cvt = "001" Then
    KyosoKigoCord = ""
ElseIf cvt = "002" Then
    KyosoKigoCord = "尒廗偄丒庒庤婻庤"
ElseIf cvt = "003" Then
    KyosoKigoCord = ""
ElseIf cvt = "004" Then
    KyosoKigoCord = ""
ElseIf cvt = "020" Then
    KyosoKigoCord = "柲"
ElseIf cvt = "021" Then
    KyosoKigoCord = "柲"
ElseIf cvt = "023" Then
    KyosoKigoCord = "柲"
ElseIf cvt = "024" Then
    KyosoKigoCord = "柲"
ElseIf cvt = "030" Then
    KyosoKigoCord = "壊丒据"
ElseIf cvt = "031" Then
    KyosoKigoCord = "壊丒据"
ElseIf cvt = "033" Then
    KyosoKigoCord = "壊丒据"
ElseIf cvt = "034" Then
    KyosoKigoCord = "壊丒据"
ElseIf cvt = "040" Then
    KyosoKigoCord = "壊丒柲"
ElseIf cvt = "041" Then
    KyosoKigoCord = "壊丒柲"
ElseIf cvt = "043" Then
    KyosoKigoCord = "壊丒柲"
ElseIf cvt = "044" Then
    KyosoKigoCord = "壊丒柲"
ElseIf cvt = "A00" Then
    KyosoKigoCord = "崿崌"
ElseIf cvt = "A01" Then
    KyosoKigoCord = "崿崌"
ElseIf cvt = "A02" Then
    KyosoKigoCord = "崿崌尒廗偄丒庒庤婻庤"
ElseIf cvt = "A03" Then
    KyosoKigoCord = "崿崌"
ElseIf cvt = "A04" Then
    KyosoKigoCord = "崿崌"
ElseIf cvt = "A10" Then
    KyosoKigoCord = "崿崌壊"
ElseIf cvt = "A11" Then
    KyosoKigoCord = "崿崌壊"
ElseIf cvt = "A13" Then
    KyosoKigoCord = "崿崌壊"
ElseIf cvt = "A14" Then
    KyosoKigoCord = "崿崌壊"
ElseIf cvt = "A20" Then
    KyosoKigoCord = "崿崌柲"
ElseIf cvt = "A21" Then
    KyosoKigoCord = "崿崌柲"
ElseIf cvt = "A23" Then
    KyosoKigoCord = "崿崌柲"
ElseIf cvt = "A24" Then
    KyosoKigoCord = "崿崌柲"
ElseIf cvt = "A30" Then
    KyosoKigoCord = "崿崌壊丒据"
ElseIf cvt = "A31" Then
    KyosoKigoCord = "崿崌壊丒据"
ElseIf cvt = "A33" Then
    KyosoKigoCord = "崿崌壊丒据"
ElseIf cvt = "A34" Then
    KyosoKigoCord = "崿崌壊丒据"
ElseIf cvt = "A40" Then
    KyosoKigoCord = "崿崌壊丒柲"
ElseIf cvt = "A41" Then
    KyosoKigoCord = "崿崌壊丒柲"
Else
    KyosoKigoCord = "枹愝掕"
End If

End Function


Function KyosoJyokenCord(cvt)

If cvt = "000" Then
    KyosoJyokenCord = ""
ElseIf cvt = "001" Then
    KyosoJyokenCord = "100枩墌埲壓"
ElseIf cvt = "002" Then
    KyosoJyokenCord = "200枩墌埲壓"
ElseIf cvt = "003" Then
    KyosoJyokenCord = "300枩墌埲壓"
ElseIf cvt = "099" Then
    KyosoJyokenCord = "9900枩墌埲壓"
ElseIf cvt = "100" Then
    KyosoJyokenCord = "1壄墌埲壓"
ElseIf cvt = "701" Then
    KyosoJyokenCord = "怴攏"
ElseIf cvt = "702" Then
    KyosoJyokenCord = "枹弌憱"
ElseIf cvt = "703" Then
    KyosoJyokenCord = "枹彑棙"
ElseIf cvt = "999" Then
    KyosoJyokenCord = "僆乕僾儞"
Else
    KyosoJyokenCord = ""
End If


End Function


Function JyuryosyubetuCord(cvt)

If cvt = 0 Then
    JyuryosyubetuCord = "枹愝掕"
ElseIf cvt = 1 Then
    JyuryosyubetuCord = "僴儞僨"
ElseIf cvt = 2 Then
    JyuryosyubetuCord = ""
ElseIf cvt = 3 Then
    JyuryosyubetuCord = "攏楊"
ElseIf cvt = 4 Then
    JyuryosyubetuCord = ""
Else
    JyuryosyubetuCord = "枹愝掕"
End If

End Function




Function TrackCord(cvt)

If cvt = "00" Then
    TrackCord = "枹愝掕"
ElseIf cvt = 10 Then
    TrackCord = "幣捈"
ElseIf cvt = 11 Then
    TrackCord = "幣嵍"
ElseIf cvt = 12 Then
    TrackCord = "幣嵍奜"
ElseIf cvt = 13 Then
    TrackCord = "幣嵍撪仺奜"
ElseIf cvt = 14 Then
    TrackCord = "幣嵍奜仺撪"
ElseIf cvt = 15 Then
    TrackCord = "幣嵍撪2廡"
ElseIf cvt = 16 Then
    TrackCord = "幣嵍奜2廡"
ElseIf cvt = 17 Then
    TrackCord = "幣塃"
ElseIf cvt = 18 Then
    TrackCord = "幣塃奜"
ElseIf cvt = 19 Then
    TrackCord = "幣塃撪仺奜"
ElseIf cvt = 20 Then
    TrackCord = "幣塃奜仺撪"
ElseIf cvt = 21 Then
    TrackCord = "幣塃撪2廡"
ElseIf cvt = 22 Then
    TrackCord = "幣塃奜2廡"
ElseIf cvt = 23 Then
    TrackCord = "僟乕僩嵍"
ElseIf cvt = 24 Then
    TrackCord = "僟乕僩塃"
ElseIf cvt = 25 Then
    TrackCord = "僟乕僩嵍撪"
ElseIf cvt = 26 Then
    TrackCord = "僟乕僩塃奜"
ElseIf cvt = 27 Then
    TrackCord = "僒儞僪嵍"
ElseIf cvt = 28 Then
    TrackCord = "僒儞僪塃"
ElseIf cvt = 29 Then
    TrackCord = "僟乕僩捈"
ElseIf cvt = 51 Then
    TrackCord = "忈奞幣鍲"
ElseIf cvt = 52 Then
    TrackCord = "忈奞幣仺僟乕僩"
ElseIf cvt = 53 Then
    TrackCord = "忈奞幣嵍"
ElseIf cvt = 54 Then
    TrackCord = "忈奞幣"
ElseIf cvt = 55 Then
    TrackCord = "忈奞幣奜"
ElseIf cvt = 56 Then
    TrackCord = "忈奞幣奜仺撪"
ElseIf cvt = 57 Then
    TrackCord = "忈奞幣撪仺奜"
ElseIf cvt = 58 Then
    TrackCord = "忈奞幣撪2廡"
ElseIf cvt = 59 Then
    TrackCord = "忈奞幣奜2廡"
Else
    TrackCord = "枹愝掕"
End If

End Function


Function TyakusaCord(cvt)

If cvt = "___" Then
    TyakusaCord = "枹愝掕"
ElseIf cvt = "_12" Then
    TyakusaCord = "1/2攏恎"
ElseIf cvt = "_34" Then
    TyakusaCord = "3/4攏恎"
ElseIf cvt = "1__" Then
    TyakusaCord = "1攏恎"
ElseIf cvt = "112" Then
    TyakusaCord = "1 1/2攏恎"
ElseIf cvt = "114" Then
    TyakusaCord = "1 1/4攏恎"
ElseIf cvt = "134" Then
    TyakusaCord = "1 3/4攏恎"
ElseIf cvt = "2__" Then
    TyakusaCord = "2攏恎"
ElseIf cvt = "212" Then
    TyakusaCord = "2 1/2攏恎"
ElseIf cvt = "3__" Then
    TyakusaCord = "3攏恎"
ElseIf cvt = "312" Then
    TyakusaCord = "3 1/2攏恎"
ElseIf cvt = "4__" Then
    TyakusaCord = "4攏恎"
ElseIf cvt = "5__" Then
    TyakusaCord = "5攏恎"
ElseIf cvt = "6__" Then
    TyakusaCord = "6攏恎"
ElseIf cvt = "7__" Then
    TyakusaCord = "7攏恎"
ElseIf cvt = "8__" Then
    TyakusaCord = "8攏恎"
ElseIf cvt = "9__" Then
    TyakusaCord = "9攏恎"
ElseIf cvt = "A__" Then
    TyakusaCord = "傾僞儅"
ElseIf cvt = "D__" Then
    TyakusaCord = "摨拝"
ElseIf cvt = "H__" Then
    TyakusaCord = "僴僫"
ElseIf cvt = "K__" Then
    TyakusaCord = ""
ElseIf cvt = "T__" Then
    TyakusaCord = "攏恎"
ElseIf cvt = "Z__" Then
    TyakusaCord = "攏恎"
Else
    TyakusaCord = "枹愝掕"
End If

End Function



Function HinsyuCord(cvt)

If cvt = 0 Then
    HinsyuCord = "枹愝掕"
ElseIf cvt = 1 Then
    HinsyuCord = "僒儔"
ElseIf cvt = 2 Then
    HinsyuCord = "僒儔宯"
ElseIf cvt = 3 Then
    HinsyuCord = "弨僒儔"
ElseIf cvt = 4 Then
    HinsyuCord = "寉敿"
ElseIf cvt = 5 Then
    HinsyuCord = "傾傾"
ElseIf cvt = 6 Then
    HinsyuCord = "傾儔宯"
ElseIf cvt = 7 Then
    HinsyuCord = "傾儔僽"
ElseIf cvt = 8 Then
    HinsyuCord = "拞敿"
Else
    HinsyuCord = "枹愝掕"
End If

End Function






