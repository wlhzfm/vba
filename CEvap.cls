VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cevap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim tbl_temp(4) As Double
Dim tbl_v(4) As Double
Dim tbl_Eta(4, 4) As Double


Sub load()
    ' �����������, �
    tbl_temp(0) = 10
    tbl_temp(1) = 15
    tbl_temp(2) = 20
    tbl_temp(3) = 30
    tbl_temp(4) = 35
    ' �������� �������� �������, �/�
    tbl_v(0) = 0
    tbl_v(1) = 0.1
    tbl_v(2) = 0.2
    tbl_v(3) = 0.5
    tbl_v(4) = 1#
    ' �����������
    tbl_Eta(0, 0) = 1#
    tbl_Eta(0, 1) = 1#
    tbl_Eta(0, 2) = 1#
    tbl_Eta(0, 3) = 1#
    tbl_Eta(0, 4) = 1#
    '
    tbl_Eta(1, 0) = 3#
    tbl_Eta(1, 1) = 2.6
    tbl_Eta(1, 2) = 2.4
    tbl_Eta(1, 3) = 1.8
    tbl_Eta(1, 4) = 1.6
    '
    tbl_Eta(2, 0) = 4.6
    tbl_Eta(2, 1) = 3.8
    tbl_Eta(2, 2) = 3.5
    tbl_Eta(2, 3) = 2.4
    tbl_Eta(2, 4) = 2.3
    '
    tbl_Eta(3, 0) = 6.6
    tbl_Eta(3, 1) = 5.7
    tbl_Eta(3, 2) = 5.4
    tbl_Eta(3, 3) = 3.6
    tbl_Eta(3, 4) = 3.2
    '
    tbl_Eta(4, 0) = 10#
    tbl_Eta(4, 1) = 8.7
    tbl_Eta(4, 2) = 7.7
    tbl_Eta(4, 3) = 5.6
    tbl_Eta(4, 4) = 4.6
End Sub



Private Function getEta(v As Double, temp As Double) As Double
Dim w_idx As Integer, temp_idx As Integer
Dim Eta As Double
    v_idx = MyMath.nearest(v, tbl_v)
    temp_idx = MyMath.nearest(temp, tbl_temp)
    Eta = tbl_Eta(v_idx, temp_idx)
    '
    getEta = Eta
End Function

Public Function getVapor(v As Double, temp As Double, _
                          Mm As Double, Ps As Double) As Double
    
Dim Eta As Double
Dim W As Double
    '
    load
    '
    Eta = getEta(v, temp)
    W = 0.000001 * Eta * Mm ^ 0.5 * Ps
    '
    getVapor = W
End Function

Public Function getVaporM(S As Double, time As Double, W As Double) As Double
Dim M As Double
    M = W * S * time
    '
    getVaporM = M
End Function

Function getPs(a As Double, B As Double, Ca As Double, temp As Double) As Double
Dim Ps As Double
    Ps = 10 ^ (a - B / (T + Ca))
    '
    getPs = Ps
End Function

' ���
Function getVaporLPG(Mm As Double, L��� As Double, T0 As Double, T� As Double, _
                    lambda_�� As Double, T As Double, a As Double, lambda_� As Double, _
                    v As Double, d As Double, nu_� As Double) As Double
Dim Re As Double
Dim M As Double
    Re = v * d / nu_�
    M = Mm / L��� * (T0 - T�) * (2 * lambd�_�� * (T / (pi * a)) ^ 0.5 + 5.1 * (Re * lambda_� * T) ^ 0.5 / T)
    '
    getVaporLPG = M
End Function



'************************************************************************
'*********************** ���������������� ������� ***********************
'************************************************************************
'
'Dim evap As Cevap
'
'Function ����_�������������������������_��__�2_�(Mm_���������� As Double, _
'                                     Ps_����������_��� As Double, _
'                                     �������������_� As Double, _
'                                     ��������_����� As Double) As Double
'Dim W As Double
'    Set evap = New Cevap
'    W = evap.getVapor(�������������_�, ��������_�����, Mm_����������, Ps_����������_���)
'    '
'    ����_�������������������������_��__�2_� = W
'End Function
'
'Function ����_������������_�����_��(W_�������������_��__�2_� As Double, _
'                                S_�������_�2 As Double, _
'                                t_�����_��� As Double)
'    Set evap = New Cevap
'    ����_������������_�����_�� = evap.getVaporM(S_�������_�2, t_�����_���, W_�������������_��__�2_�)
'End Function
'
'Function ����_����������������_���(a As Double, _
'                                   B As Double, _
'                                   Ca As Double, _
'                                   t_�����������_� As Double) As Double
'    Set evap = New Cevap
'    ����_����������������_��� = evap.getPs(a, B, Ca, t_�����������_�)
'End Function




'Mm 102, 2
'��������� �������
'a 7, 54424
'B 2629, 65
'Ca 384, 195
'temp    10  C
'v   1   �/�
'S   100 �2
'time    3600    ���
'
'Ps  5,01    ���
'W   5,06E-04    ��/(�2*�)
'M   182,3   ��



