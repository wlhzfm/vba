VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ������� ���
Public Function f_GG_R(M As Double, rho As Double, CLFL As Double) As Double
    f_GG_R = 14.6 * (M / (rho * CLFL)) ^ 0.33
End Function

Public Function f_GG_Z(M As Double, rho As Double, CLFL As Double) As Double
    f_GG_Z = 0.33 * (M / (ro * CLFL)) ^ 0.33
End Function

' ���
Public Function f_FL_R(M As Double, T As Double, rho As Double, Ps As Double, CLFL As Double) As Double
    f_FL_R = 3.2 * (T / 3600) ^ 0.5 * (Ps / CLFL) ^ 0.8 * (M / (rho * Ps)) ^ 0.35
End Function

Public Function f_FL_Z(M As Double, T As Double, rho As Double, Ps As Double, CLFL As Double) As Double
    f_FL_Z = 0.12 * (T / 3600) ^ 0.5 * (Ps / CLFL) ^ 0.8 * (M / (rho * Ps)) ^ 0.35
End Function

'************************************************************************
'*********************** ���������������� ������� ***********************
'************************************************************************
'Dim lfl As CLFL
'Function �404_����_��_R(M_����������_�� As Double, _
'                        rho_��������������_��__�3, _
'                        C���� As Double) As Double
'    Set lfl = New CLFL
'    �404_����_��_R = lfl.f_GG_R(M_����������_��, rho_��������������_��__�3, C����)
'End Function
'
'Function �404_����_��_Z(M_����������_�� As Double, _
'                        rho_��������������_��__�3, _
'                        C���� As Double) As Double
'    Set lfl = New CLFL
'    �404_����_��_Z = lfl.f_GG_R(M_����������_��, rho_��������������_��__�3, C����)
'End Function
'
'Function �404_����_���_R(M_����������_�� As Double, _
'                         T_��������������_��� As Double, _
'                         rho_��������������_��__�3 As Double, _
'                         Ps_������������_��__�3 As Double, _
'                         C���� As Double) As Double
'    Set lfl = New CLFL
'    �404_����_���_R = lfl.f_FL_R(M_����������_��, T_��������������_���, _
'                             rho_��������������_��__�3, Ps_������������_��__�3, _
'                             C����)
'End Function
'
'Function �404_����_���_Z(M_����������_�� As Double, _
'                         T_��������������_��� As Double, _
'                         rho_��������������_��__�3 As Double, _
'                         Ps_������������_��__�3 As Double, _
'                         C���� As Double) As Double
'    Set lfl = New CLFL
'    �404_����_���_Z = lfl.f_FL_Z(M_����������_��, T_��������������_���, _
'                             rho_��������������_��__�3, Ps_������������_��__�3, _
'                             C����)
'End Function

