Attribute VB_Name = "Module_Trancerate"
Function JyoCord(cvt)

If cvt = "01" Then
    JyoCord = "�D�y"
ElseIf cvt = "02" Then
    JyoCord = "����"
ElseIf cvt = "03" Then
    JyoCord = "����"
ElseIf cvt = "04" Then
    JyoCord = "�V��"
ElseIf cvt = "05" Then
    JyoCord = "����"
ElseIf cvt = "06" Then
    JyoCord = "���R"
ElseIf cvt = "07" Then
    JyoCord = "����"
ElseIf cvt = "08" Then
    JyoCord = "���s"
ElseIf cvt = "09" Then
    JyoCord = "��_"
ElseIf cvt = 10 Then
    JyoCord = "���q"
ElseIf cvt = 30 Then
    JyoCord = "���"
ElseIf cvt = 31 Then
    JyoCord = "�k��"
ElseIf cvt = 32 Then
    JyoCord = "�〈��"
ElseIf cvt = 33 Then
    JyoCord = "�эL"
ElseIf cvt = 34 Then
    JyoCord = "����"
ElseIf cvt = 35 Then
    JyoCord = "����"
ElseIf cvt = 36 Then
    JyoCord = "����"
ElseIf cvt = 37 Then
    JyoCord = "��R"
ElseIf cvt = 38 Then
    JyoCord = "�O��"
ElseIf cvt = 39 Then
    JyoCord = "����"
ElseIf cvt = 40 Then
    JyoCord = "�F�s�{"
ElseIf cvt = 41 Then
    JyoCord = "����"
ElseIf cvt = 42 Then
    JyoCord = "�Y�a"
ElseIf cvt = 43 Then
    JyoCord = "�D��"
ElseIf cvt = 44 Then
    JyoCord = "���"
ElseIf cvt = 45 Then
    JyoCord = "���"
ElseIf cvt = 46 Then
    JyoCord = "����"
ElseIf cvt = 47 Then
    JyoCord = "�}��"
ElseIf cvt = 48 Then
    JyoCord = "���É�"
ElseIf cvt = 49 Then
    JyoCord = "�I�ɎO�䎛"
ElseIf cvt = 50 Then
    JyoCord = "���c"
ElseIf cvt = 51 Then
    JyoCord = "�P�H"
ElseIf cvt = 52 Then
    JyoCord = "�v�c"
ElseIf cvt = 53 Then
    JyoCord = "���R"
ElseIf cvt = 54 Then
    JyoCord = "���m"
ElseIf cvt = 55 Then
    JyoCord = "����"
ElseIf cvt = 56 Then
    JyoCord = "�r��"
ElseIf cvt = 57 Then
    JyoCord = "����"
ElseIf cvt = 58 Then
    JyoCord = "�D�y(�n��)"
ElseIf cvt = 59 Then
    JyoCord = "����(�n��)"
ElseIf cvt = 60 Then
    JyoCord = "�V��(�n��)"
ElseIf cvt = 61 Then
    JyoCord = "����(�n��)"
End If

End Function

Function ShortJyo2Jyo(cvt)

If cvt = "�D" Then
    ShortJyo2Jyo = "�D�y"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "����"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "����"
ElseIf cvt = "�V" Then
    ShortJyo2Jyo = "�V��"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "����"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "���R"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "����"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "���s"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "��_"
ElseIf cvt = "��" Then
    ShortJyo2Jyo = "���q"
End If

End Function

Function Jyo2ShortJyo(cvt)

If cvt = "�D�y" Then
    Jyo2ShortJyo = "�D"
ElseIf cvt = "����" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "����" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "�V��" Then
    Jyo2ShortJyo = "�V"
ElseIf cvt = "����" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "���R" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "����" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "���s" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "��_" Then
    Jyo2ShortJyo = "��"
ElseIf cvt = "���q" Then
    Jyo2ShortJyo = "��"
End If

End Function

Function JyogyakuCord(cvt)

If cvt = "�D�y" Then
    JyogyakuCord = "01"
ElseIf cvt = "����" Then
    JyogyakuCord = "02"
ElseIf cvt = "����" Then
    JyogyakuCord = "03"
ElseIf cvt = "�V��" Then
    JyogyakuCord = "04"
ElseIf cvt = "����" Then
    JyogyakuCord = "05"
ElseIf cvt = "���R" Then
    JyogyakuCord = "06"
ElseIf cvt = "����" Then
    JyogyakuCord = "07"
ElseIf cvt = "���s" Then
    JyogyakuCord = "08"
ElseIf cvt = "��_" Then
    JyogyakuCord = "09"
ElseIf cvt = "���q" Then
    JyogyakuCord = 10
ElseIf cvt = "���" Then
    JyogyakuCord = 30
ElseIf cvt = "�k��" Then
    JyogyakuCord = 31
ElseIf cvt = "�〈��" Then
    JyogyakuCord = 32
ElseIf cvt = "�эL" Then
    JyogyakuCord = 33
ElseIf cvt = "����" Then
    JyogyakuCord = 34
ElseIf cvt = "����" Then
    JyogyakuCord = 35
ElseIf cvt = "����" Then
    JyogyakuCord = 36
ElseIf cvt = "��R" Then
    JyogyakuCord = 37
ElseIf cvt = "�O��" Then
    JyogyakuCord = 38
ElseIf cvt = "����" Then
    JyogyakuCord = 39
ElseIf cvt = "�F�s�{" Then
    JyogyakuCord = 40
ElseIf cvt = "����" Then
    JyogyakuCord = 41
ElseIf cvt = "�Y�a" Then
    JyogyakuCord = 42
ElseIf cvt = "�D��" Then
    JyogyakuCord = 43
ElseIf cvt = "���" Then
    JyogyakuCord = 44
ElseIf cvt = "���" Then
    JyogyakuCord = 45
ElseIf cvt = "����" Then
    JyogyakuCord = 46
ElseIf cvt = "�}��" Then
    JyogyakuCord = 47
ElseIf cvt = "���É�" Then
    JyogyakuCord = 48
ElseIf cvt = "�I�ɎO�䎛" Then
    JyogyakuCord = 49
ElseIf cvt = "���c" Then
    JyogyakuCord = 50
ElseIf cvt = "�P�H" Then
    JyogyakuCord = 51
ElseIf cvt = "�v�c" Then
    JyogyakuCord = 52
ElseIf cvt = "���R" Then
    JyogyakuCord = 53
ElseIf cvt = "���m" Then
    JyogyakuCord = 54
ElseIf cvt = "����" Then
    JyogyakuCord = 55
ElseIf cvt = "�r��" Then
    JyogyakuCord = 56
ElseIf cvt = "����" Then
    JyogyakuCord = 57
ElseIf cvt = "�D�y(�n��)" Then
    JyogyakuCord = 58
ElseIf cvt = "����(�n��)" Then
    JyogyakuCord = 59
ElseIf cvt = "�V��(�n��)" Then
    JyogyakuCord = 60
ElseIf cvt = "����(�n��)" Then
    JyogyakuCord = 61
End If

End Function

Function TenkoCord(cvt)

If cvt = 0 Then
    TenkoCord = "���ݒ�"
ElseIf cvt = 1 Then
    TenkoCord = "��"
ElseIf cvt = 2 Then
    TenkoCord = "��"
ElseIf cvt = 3 Then
    TenkoCord = "�J"
ElseIf cvt = 4 Then
    TenkoCord = "���J"
ElseIf cvt = 5 Then
    TenkoCord = "��"
ElseIf cvt = 6 Then
    TenkoCord = "����"
End If

End Function


Function BabaCord(cvt)

If cvt = 0 Then
    BabaCord = "���ݒ�"
ElseIf cvt = 1 Then
    BabaCord = "��"
ElseIf cvt = 2 Then
    BabaCord = "�c�d"
ElseIf cvt = 3 Then
    BabaCord = "�d"
ElseIf cvt = 4 Then
    BabaCord = "�s��"
End If

End Function


Function SeibetuCord(cvt)

If cvt = 0 Then
    SeibetuCord = "���ݒ�"
ElseIf cvt = 1 Then
    SeibetuCord = "���n"
ElseIf cvt = 2 Then
    SeibetuCord = "�Ĕn"
ElseIf cvt = 3 Then
    SeibetuCord = "�Z���n"
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
    GradeCord = "�d��"
ElseIf cvt = "E" Then
    GradeCord = "����"
ElseIf cvt = "F" Then
    GradeCord = "J�EG1"
ElseIf cvt = "G" Then
    GradeCord = "J�EG2"
ElseIf cvt = "H" Then
    GradeCord = "J�EG3"
Else
    GradeCord = "-"
End If

End Function

Function YoubiCord(cvt)

If cvt = 0 Then
    YoubiCord = "���ݒ�"
ElseIf cvt = 1 Then
    YoubiCord = "�y"
ElseIf cvt = 2 Then
    YoubiCord = "��"
ElseIf cvt = 3 Then
    YoubiCord = "�j"
ElseIf cvt = 4 Then
    YoubiCord = "��"
ElseIf cvt = 5 Then
    YoubiCord = "��"
ElseIf cvt = 6 Then
    YoubiCord = "��"
ElseIf cvt = 7 Then
    YoubiCord = "��"
ElseIf cvt = 8 Then
    YoubiCord = "��"
End If

End Function


Function KyososyubetuCord(cvt)

If cvt = "00" Then
    KyososyubetuCord = "���ݒ�"
ElseIf cvt = 11 Then
    KyososyubetuCord = "2��"
ElseIf cvt = 12 Then
    KyososyubetuCord = "3��"
ElseIf cvt = 13 Then
    KyososyubetuCord = "3�Έȏ�"
ElseIf cvt = 14 Then
    KyososyubetuCord = "4�Έȏ�"
ElseIf cvt = 18 Then
    KyososyubetuCord = "3�Έȏ�"
ElseIf cvt = 19 Then
    KyososyubetuCord = "4�Έȏ�"
ElseIf cvt = 21 Then
    KyososyubetuCord = "2��"
ElseIf cvt = 22 Then
    KyososyubetuCord = "3��"
ElseIf cvt = 23 Then
    KyososyubetuCord = "3�Έȏ�"
ElseIf cvt = 24 Then
    KyososyubetuCord = "4�Έȏ�"
End If


End Function




Function KyosoKigoCord(cvt)


If cvt = "000" Then
    KyosoKigoCord = "���ݒ�"
ElseIf cvt = "001" Then
    KyosoKigoCord = ""
ElseIf cvt = "002" Then
    KyosoKigoCord = "���K���E���R��"
ElseIf cvt = "003" Then
    KyosoKigoCord = ""
ElseIf cvt = "004" Then
    KyosoKigoCord = ""
ElseIf cvt = "020" Then
    KyosoKigoCord = "��"
ElseIf cvt = "021" Then
    KyosoKigoCord = "��"
ElseIf cvt = "023" Then
    KyosoKigoCord = "��"
ElseIf cvt = "024" Then
    KyosoKigoCord = "��"
ElseIf cvt = "030" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "031" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "033" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "034" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "040" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "041" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "043" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "044" Then
    KyosoKigoCord = "���E��"
ElseIf cvt = "A00" Then
    KyosoKigoCord = "����"
ElseIf cvt = "A01" Then
    KyosoKigoCord = "����"
ElseIf cvt = "A02" Then
    KyosoKigoCord = "�������K���E���R��"
ElseIf cvt = "A03" Then
    KyosoKigoCord = "����"
ElseIf cvt = "A04" Then
    KyosoKigoCord = "����"
ElseIf cvt = "A10" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A11" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A13" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A14" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A20" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A21" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A23" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A24" Then
    KyosoKigoCord = "������"
ElseIf cvt = "A30" Then
    KyosoKigoCord = "�������E��"
ElseIf cvt = "A31" Then
    KyosoKigoCord = "�������E��"
ElseIf cvt = "A33" Then
    KyosoKigoCord = "�������E��"
ElseIf cvt = "A34" Then
    KyosoKigoCord = "�������E��"
ElseIf cvt = "A40" Then
    KyosoKigoCord = "�������E��"
ElseIf cvt = "A41" Then
    KyosoKigoCord = "�������E��"
Else
    KyosoKigoCord = "���ݒ�"
End If

End Function


Function KyosoJyokenCord(cvt)

If cvt = "000" Then
    KyosoJyokenCord = ""
ElseIf cvt = "001" Then
    KyosoJyokenCord = "100���~�ȉ�"
ElseIf cvt = "002" Then
    KyosoJyokenCord = "200���~�ȉ�"
ElseIf cvt = "003" Then
    KyosoJyokenCord = "300���~�ȉ�"
ElseIf cvt = "099" Then
    KyosoJyokenCord = "9900���~�ȉ�"
ElseIf cvt = "100" Then
    KyosoJyokenCord = "1���~�ȉ�"
ElseIf cvt = "701" Then
    KyosoJyokenCord = "�V�n"
ElseIf cvt = "702" Then
    KyosoJyokenCord = "���o��"
ElseIf cvt = "703" Then
    KyosoJyokenCord = "������"
ElseIf cvt = "999" Then
    KyosoJyokenCord = "�I�[�v��"
Else
    KyosoJyokenCord = ""
End If


End Function


Function JyuryosyubetuCord(cvt)

If cvt = 0 Then
    JyuryosyubetuCord = "���ݒ�"
ElseIf cvt = 1 Then
    JyuryosyubetuCord = "�n���f"
ElseIf cvt = 2 Then
    JyuryosyubetuCord = ""
ElseIf cvt = 3 Then
    JyuryosyubetuCord = "�n��"
ElseIf cvt = 4 Then
    JyuryosyubetuCord = ""
Else
    JyuryosyubetuCord = "���ݒ�"
End If

End Function




Function TrackCord(cvt)

If cvt = "00" Then
    TrackCord = "���ݒ�"
ElseIf cvt = 10 Then
    TrackCord = "�Œ�"
ElseIf cvt = 11 Then
    TrackCord = "�ō�"
ElseIf cvt = 12 Then
    TrackCord = "�ō��O"
ElseIf cvt = 13 Then
    TrackCord = "�ō������O"
ElseIf cvt = 14 Then
    TrackCord = "�ō��O����"
ElseIf cvt = 15 Then
    TrackCord = "�ō���2�T"
ElseIf cvt = 16 Then
    TrackCord = "�ō��O2�T"
ElseIf cvt = 17 Then
    TrackCord = "�ŉE"
ElseIf cvt = 18 Then
    TrackCord = "�ŉE�O"
ElseIf cvt = 19 Then
    TrackCord = "�ŉE�����O"
ElseIf cvt = 20 Then
    TrackCord = "�ŉE�O����"
ElseIf cvt = 21 Then
    TrackCord = "�ŉE��2�T"
ElseIf cvt = 22 Then
    TrackCord = "�ŉE�O2�T"
ElseIf cvt = 23 Then
    TrackCord = "�_�[�g��"
ElseIf cvt = 24 Then
    TrackCord = "�_�[�g�E"
ElseIf cvt = 25 Then
    TrackCord = "�_�[�g����"
ElseIf cvt = 26 Then
    TrackCord = "�_�[�g�E�O"
ElseIf cvt = 27 Then
    TrackCord = "�T���h��"
ElseIf cvt = 28 Then
    TrackCord = "�T���h�E"
ElseIf cvt = 29 Then
    TrackCord = "�_�[�g��"
ElseIf cvt = 51 Then
    TrackCord = "��Q���F"
ElseIf cvt = 52 Then
    TrackCord = "��Q�Ł��_�[�g"
ElseIf cvt = 53 Then
    TrackCord = "��Q�ō�"
ElseIf cvt = 54 Then
    TrackCord = "��Q��"
ElseIf cvt = 55 Then
    TrackCord = "��Q�ŊO"
ElseIf cvt = 56 Then
    TrackCord = "��Q�ŊO����"
ElseIf cvt = 57 Then
    TrackCord = "��Q�œ����O"
ElseIf cvt = 58 Then
    TrackCord = "��Q�œ�2�T"
ElseIf cvt = 59 Then
    TrackCord = "��Q�ŊO2�T"
Else
    TrackCord = "���ݒ�"
End If

End Function


Function TyakusaCord(cvt)

If cvt = "___" Then
    TyakusaCord = "���ݒ�"
ElseIf cvt = "_12" Then
    TyakusaCord = "1/2�n�g"
ElseIf cvt = "_34" Then
    TyakusaCord = "3/4�n�g"
ElseIf cvt = "1__" Then
    TyakusaCord = "1�n�g"
ElseIf cvt = "112" Then
    TyakusaCord = "1 1/2�n�g"
ElseIf cvt = "114" Then
    TyakusaCord = "1 1/4�n�g"
ElseIf cvt = "134" Then
    TyakusaCord = "1 3/4�n�g"
ElseIf cvt = "2__" Then
    TyakusaCord = "2�n�g"
ElseIf cvt = "212" Then
    TyakusaCord = "2 1/2�n�g"
ElseIf cvt = "3__" Then
    TyakusaCord = "3�n�g"
ElseIf cvt = "312" Then
    TyakusaCord = "3 1/2�n�g"
ElseIf cvt = "4__" Then
    TyakusaCord = "4�n�g"
ElseIf cvt = "5__" Then
    TyakusaCord = "5�n�g"
ElseIf cvt = "6__" Then
    TyakusaCord = "6�n�g"
ElseIf cvt = "7__" Then
    TyakusaCord = "7�n�g"
ElseIf cvt = "8__" Then
    TyakusaCord = "8�n�g"
ElseIf cvt = "9__" Then
    TyakusaCord = "9�n�g"
ElseIf cvt = "A__" Then
    TyakusaCord = "�A�^�}"
ElseIf cvt = "D__" Then
    TyakusaCord = "����"
ElseIf cvt = "H__" Then
    TyakusaCord = "�n�i"
ElseIf cvt = "K__" Then
    TyakusaCord = ""
ElseIf cvt = "T__" Then
    TyakusaCord = "�n�g"
ElseIf cvt = "Z__" Then
    TyakusaCord = "�n�g"
Else
    TyakusaCord = "���ݒ�"
End If

End Function



Function HinsyuCord(cvt)

If cvt = 0 Then
    HinsyuCord = "���ݒ�"
ElseIf cvt = 1 Then
    HinsyuCord = "�T��"
ElseIf cvt = 2 Then
    HinsyuCord = "�T���n"
ElseIf cvt = 3 Then
    HinsyuCord = "���T��"
ElseIf cvt = 4 Then
    HinsyuCord = "�y��"
ElseIf cvt = 5 Then
    HinsyuCord = "�A�A"
ElseIf cvt = 6 Then
    HinsyuCord = "�A���n"
ElseIf cvt = 7 Then
    HinsyuCord = "�A���u"
ElseIf cvt = 8 Then
    HinsyuCord = "����"
Else
    HinsyuCord = "���ݒ�"
End If

End Function






