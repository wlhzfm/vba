VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�������� �������� ������� �������� ��� �������� �������������
Const Eud0# = 44 ' ���/��
' ����������� ��������, ��
Const const_P0 = 101325 ' ��
' �������� ����� � ������� �  (������ ����������� ������ 340 �/�)
Const const_C0 = 340 ' �/�

' ������ ������� �������
Enum En_classMatter
    �����_1 = 1
    �����_2 = 2
    �����_3 = 3
    �����_4 = 4
End Enum

' ������������� ����������� ������������ �� ������� ����������������
Enum En_classSpaceClutter
    �����_I = 1
    �����_II = 2
    �����_III = 3
    �����_IV = 4
End Enum

' ��������� ����� �������� ������
' ������ ������ �������� ������
Enum En_combustionMode
    �����_�����_1 = 1
    �����_�����_2 = 2
    �����_�����_3 = 3
    �����_�����_4 = 4
    �����_�����_5 = 5
    �����_�����_6 = 6
End Enum

' sigma - ������� ���������� ��������� �������� (��� �����������������
' ������ ����������� ��������� ������ 7, ��� ������������� ������ 4)
Enum En_sigma
    ���������������������� = 7
    ������������������ = 4
End Enum


Private Function Max#(val_1#, val_2#)
    Max = IIf(val_1 > val_2, val_1, val_2)
End Function
Private Function Ln#(val#)
    Ln = Log(val) / Log(epx(1))
End Function

Sub clearImmediate()
    Application.SendKeys "^g ^a {DEL}"
    Application.Wait Now + TimeValue("0:00:01")
End Sub


Function get_combustionModeTbl() As En_combustionMode()
Dim tbl(1 To 4, 1 To 4) As En_combustionMode
Dim R%
    R = 1 ' ������ 1
    tbl(R, 1) = �����_�����_1
        tbl(R, 2) = �����_�����_1
            tbl(R, 3) = �����_�����_2
                tbl(R, 4) = �����_�����_3
    R = 2 ' ������ 2
    tbl(R, 1) = �����_�����_1
        tbl(R, 2) = �����_�����_2
            tbl(R, 3) = �����_�����_3
                tbl(R, 4) = �����_�����_4
    R = 3 ' ������ 3
    tbl(R, 1) = �����_�����_2
        tbl(R, 2) = �����_�����_3
            tbl(R, 3) = �����_�����_4
                tbl(R, 4) = �����_�����_5
    R = 4 ' ������ 4
    tbl(R, 1) = �����_�����_3
        tbl(R, 2) = �����_�����_4
            tbl(R, 3) = �����_�����_5
                tbl(R, 4) = �����_�����_6
    
    get_combustionModeTbl = tbl
End Function


' ��������� ������� ��� ������� ���������� ���� ��������
' ��� �������� ������ ��������:
' - ��� �������� ��������, ������������� � ������
' - ������������ �������� �������� � �����
' - ����������������� ������������ �������� �������� � ��������
' - ����� �������� �������� ��, ������������� � ������,
'   � ������������� ����� ������ � ������� ����������������
'   �������� ��������������� �������
'   ����������� ��������  �� ��������� ������ ����� ��������
'   ��������, ������������� � ������, � ������ ������������ Z
'   ������� �������� �������� �� ������. ��� ���������� ������
'   ����������� Z ����� ���� ������ ������ 0,1.
'   ��� �������� ������������ ��������� �������� ���� ��������  ��
'   ������� ������������ � ������ ������������� �������������
'   ������������ �������� ���� � �����
' - �������� ������� �������� �������� �������� E��
' - �������� ����� � ������� �0 (������ ����������� ������ 340 �/�)
' - ���������� � ������� ���������������� ����������� ������������
' - ����������� ����������� ������� ����� �, ������� ������������ �� �������



' ����������� ����������� ������� ����� E
' (�3.36)
' ground - �� ����������� �����
'   ��� ������� ���������� �������� ������, �������������� �� �����������
'   �����, �������� ������������ ������������ �����������
Function get_E(Mt#, Cg#, Cct#, Eud#, Optional ground As Boolean = True)
Dim E#
    If (C� <= C��) Then
      E = Mt * Eud
    Else
      E = Mt * Eud * Cct / Cg
    End If
    If (ground) Then
        E = 2 * E
    End If
    get_E = E * 10 ^ 6
End Function

' ��� ������ ��������� ��������� ������� �������� ������ �����������
' �������� ���������� ���������� �� ������� ��������, ������������
' ��� ������� ������� ������ ���������������. ��� �������� �������������
' ����������� � ������ �������� �������� ������� ��������  E��0 = 44 ���/��.
' ��� ���� ������� ������� � �������� ������������ �������� ���������������
' E��=betta*E��0.
' ����� betta - ���������������� ��������. ��� ������� ���������� �������
' ������� ������� �������� ��������� betta ������������ � ������� �3.
Function get_Eud(betta)
    get_Eud = Eud0 * betta
End Function


' =========================================================================
' ����������� ���������� ������ �������� ������
' =========================================================================
' ��������� ����� �������� ������ ������������ � ������� ������� �3.3,
' � ����������� �� ������ �������� �������� � ������ ����������������
' ����������� ������������
Function get_combustionMode(classSpaceClutter As En_classSpaceClutter, _
                            classMater As En_classMatter) _
                            As En_combustionMode
Dim combustionModeTbl() As En_combustionMode
    combustionModeTbl = get_combustionModeTbl()
    get_combustionMode = combustionModeTbl(classMater, classSpaceClutter)
End Function
'Sub test_get_combustionMode()
'    Debug.Print get_combustionMode(�����_II, �����_3)
'End Sub

' u �������� ������ �������
' (�3.37) (�3.38)
Function get_u(combustionMode As En_combustionMode, Optional M# = 0)
Dim k1%, k2%, u#
    k1 = 43
    k2 = 26
    If (combustionMode = �����_�����_1) Then
    ' ����� 1 - ��������� ��� ������� �� ��������� ������ �������
    '           500 �/� � �����;
        get_u = 500
    ElseIf (combustionMode = �����_�����_2) Then
    ' ��� ����������� ������������ �������� ������ ������� ���
    ' ������� �������� 2-4 ������� ������������� ��������������
    ' ������� �������� ������ ������� �� ����������� (�3.37).
    ' � ��� ������, ���� ���������� �������� ������
    ' ������������ ��������, ��������������� ������� ������,
    ' ��� ����������� �� ������� (�.3.37)
    
    ' ����� 2 - �����������, �������� ������ ������� 300-500 �/�;
        u = k1 * M ^ (1 / 6)
        get_u = Max(500, u)
    ElseIf (combustionMode = �����_�����_3) Then
    ' ����� 3 - �����������, �������� ������ ������� 200-300 �/�;
        u = k1 * M ^ (1 / 6)
        get_u = Max(300, u)
    ElseIf (combustionMode = �����_�����_4) Then
    ' ����� 4 - �����������, �������� ������ ������� 150-200 �/�;
        u = k1 * M ^ (1 / 6)
        get_u = Max(200, u)
    ElseIf (combustionMode = �����_�����_5) Then
    ' ����� 5 - �����������, �������� ������ ������� ������������ ��
    '           �������:(�3.37)
        get_u = k1 * M ^ (1 / 6)
    ElseIf (combustionMode = �����_�����_6) Then
    ' ����� 6 - �����������, �������� ������ ������� ������������
    '           �� �������:(�3.38)
        get_u = k2 * M ^ (1 / 6)
    End If
End Function


' =========================================================================
' ������ ������������� ����������� �������� � ��������
' ���� ������ ��������� ���� �������� ��� ��������� �������
' =========================================================================
' ��������� ��������� ���� �������� (���������� �������� d� � ������� ����
' ������ I+) � ����������� �� ���������� �� ������ ������ ��������������
' ������ �� ���������� ������ �������� ������
Function get_dP#(R#, combustionMode As En_combustionMode, _
                 E#, u#, _
                 Optional P0# = const_P0, _
                 Optional sigma As En_sigma = ����������������������)
Dim Px#, ln_Px#, Rx#, ln_Rx#, R��1#
Dim dP#
    Rx = get_Rx(R, E, P0)
    
    If (combustionMode = �����_�����_1) Then
        ' �������� ������������� �������� Px � ������� ���� ������ Ix ������������
        '   �� �������� (��� ����������������� ������): (�3.40)  (�3.41)
        ' ������� (�3.40, �3.41) ����������� ��� ��������  Rx > 0,2. � ������,
        '   ����  Rx<0,2, �� Px = ����� 18, � � ������� (�3.41) ������ Rx �������������
        '   ��������  = 0,14.
        If (Rx < 0.2) Then
            Px = 18
        Else
            ln_Rx = Ln(Rx)
            ln_Px = -1.124 - 1.66 * (ln_Rx) + 0.26 * (ln_Rx) ^ 2
            Px = Exp(1) ^ ln_Px
        End If
        ' ��������� �������� ����������� �������� � �������� ���� ������
        ' ������������ �� ��������: (�3.42) (�3.43)
        dP = Px * P0
    Else
        ' �������������� �������� ������������� �������� (Px1) � �������� ����
        ' ������ Ix1 �� ��������: (�3.44) (�3.45) (�3.46)
        ' ������� (�3.44), (�3.45) ����������� ��� �������� Rx ������� ��������
        ' R��1 = 0,34, � ������, ����  Rx<R��1, � ������� (�3.44), (�3.45)
        ' ������ Rx ������������� �������� R��1.
        R��1 = 0.34
        If (Rx < R��1) Then
            Rx = R��1
        End If
        ' � ������ ����������� �������������� ������ �������� ������������
        ' ������������ ���������� �� �����������  (sigma-1)/sigma.
        If (sigma = ������������������ And combustionMode <> �����_�����_1) Then
            E = E * (sigma - 1) / sigma
        End If
        Px1 = (u ^ 2 / const_C0 ^ 2) * (sigma - 1) / sigma * (0.83 / Rx - 0.14 / Rx ^ 2)
        ' ��������� �������� ����������� �������� � �������� ���� ������
        ' ������������ �� �������� (�3.42), (�3.43). ��� ���� � ������� (�3.42), (�3.43)
        ' ������ Px � Ix ������������� �������� Px1 � Ix1.
        dP = Px1 * P0
    End If
    '
    get_dP = dP
End Function
Function get_I#(R#, combustionMode As En_combustionMode, _
                E#, u#, _
                Optional P0# = const_P0, _
                Optional sigma As En_sigma = ����������������������)
Dim Ix#, ln_Ix#, Rx#, ln_Rx#, R��1#, W#
Dim i#
    Rx = get_Rx(R, E, P0)
    
    If (combustionMode = �����_�����_1) Then
        ' �������� ������������� �������� Px � ������� ���� ������ Ix ������������
        '   �� �������� (��� ����������������� ������): (�3.40)  (�3.41)
        ' ������� (�3.40, �3.41) ����������� ��� ��������  Rx > 0,2. � ������,
        '   ����  Rx<0,2, �� Px = ����� 18, � � ������� (�3.41) ������ Rx �������������
        '   ��������  = 0,14.
        If (Rx < 0.2) Then
            Rx = 0.14
        End If
        ln_Rx = Ln(Rx)
        ln_Ix = -3.4217 - 0.898 * (ln_Rx) - 0.0096 * (ln_Rx) ^ 2
        Ix = Exp(1) ^ ln_Ix

        ' ��������� �������� ����������� �������� � �������� ���� ������
        ' ������������ �� ��������: (�3.42) (�3.43)
        i = Ix * P0 ^ (2 / 3) * E ^ (1 / 3) / const_C0
    Else
        ' �������������� �������� ������������� �������� (Px1) � �������� ����
        ' ������ Ix1 �� ��������: (�3.44) (�3.45) (�3.46)
        ' ������� (�3.44), (�3.45) ����������� ��� �������� Rx ������� ��������
        ' R��1 = 0,34, � ������, ����  Rx<R��1, � ������� (�3.44), (�3.45)
        ' ������ Rx ������������� �������� R��1.
        R��1 = 0.34
        If (Rx < R��1) Then
            Rx = R��1
        End If
        W = u / const_C0 * (sigma - 1) / sigma
        ' � ������ ����������� �������������� ������ �������� ������������
        ' ������������ ���������� �� �����������  (sigma-1)/sigma.
        If (sigma = ������������������ And combustionMode <> �����_�����_1) Then
            E = E * (sigma - 1) / sigma
        End If
        Ix1 = W * (1 - 0.4 * W) * (0.06 / Rx + 0.01 / Rx ^ 2 - 0.0025 / Rx ^ 3)
        ' ��������� �������� ����������� �������� � �������� ���� ������
        ' ������������ �� �������� (�3.42), (�3.43). ��� ���� � ������� (�3.42), (�3.43)
        ' ������ Px � Ix ������������� �������� Px1 � Ix1.
         i = Ix1 * P0 ^ (2 / 3) * E ^ (1 / 3) / const_C0
    End If
    '
    get_I = i
End Function

Function get_Rx#(R#, E#, Optional P0# = const_P0)
    get_Rx = R / ((E / P0) ^ (1 / 3))
End Function


Function get_dP_ByR#(R#, M#, classMatter As En_classMatter, betta#, classSpaceClutter As En_classSpaceClutter)
Dim combustionMode As En_combustionMode
Dim u#, Eud#, E#

    combustionMode = get_combustionMode(classSpaceClutter, classMatter)
    u = get_u(combustionMode, M)
    Eud = get_Eud(betta)
    E = get_E(Mt:=M, Cg:=1, Cct:=1, Eud:=Eud, ground:=True)
    get_dP_ByR = get_dP(R:=R, combustionMode:=combustionMode, E:=E, u:=u)
End Function

Function get_I_ByR#(R#, M#, classMatter As En_classMatter, betta#, classSpaceClutter As En_classSpaceClutter)
Dim combustionMode As En_combustionMode
Dim u#, Eud#, E#

    combustionMode = get_combustionMode(classSpaceClutter, classMatter)
    u = get_u(combustionMode, M)
    Eud = get_Eud(betta)
    E = get_E(Mt:=M, Cg:=1, Cct:=1, Eud:=Eud, ground:=True)
    get_I_ByR = get_I(R:=R, combustionMode:=combustionMode, E:=E, u:=u)
End Function


' =========================================================================
' ����������� �������������� ������������� �������� ��������
' =========================================================================




' =========================================================================
' ������ ����������� �����������
' =========================================================================
Function getR#(dP_border#, M#, classMatter As En_classMatter, betta#, _
                    classSpaceClutter As En_classSpaceClutter, _
                    Optional n_max% = 5, Optional n_min% = 0)
Dim dP#, R#, dR#, n%
    
    R = 10 ^ (n_max + 1)
    For i = n_max To n_min Step -1
        dR = 10 ^ i
        Rmin = R - dR * 10
        Rmax = R
        R = Rmin
        Do
            R = R + dR
            dP = get_dP_ByR(R, M, classMatter, betta, classSpaceClutter)
        Loop Until (dP < dP_border Or R = Rmax)
    Next i
    If (Round(R - 10 ^ n_min) = 0) Then
        R = 0
    End If
    
    getR = R
End Function


