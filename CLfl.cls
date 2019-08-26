VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Горючие газ
Public Function f_GG_R(M As Double, rho As Double, CLFL As Double) As Double
    f_GG_R = 14.6 * (M / (rho * CLFL)) ^ 0.33
End Function

Public Function f_GG_Z(M As Double, rho As Double, CLFL As Double) As Double
    f_GG_Z = 0.33 * (M / (ro * CLFL)) ^ 0.33
End Function

' ЛВЖ
Public Function f_FL_R(M As Double, T As Double, rho As Double, Ps As Double, CLFL As Double) As Double
    f_FL_R = 3.2 * (T / 3600) ^ 0.5 * (Ps / CLFL) ^ 0.8 * (M / (rho * Ps)) ^ 0.35
End Function

Public Function f_FL_Z(M As Double, T As Double, rho As Double, Ps As Double, CLFL As Double) As Double
    f_FL_Z = 0.12 * (T / 3600) ^ 0.5 * (Ps / CLFL) ^ 0.8 * (M / (rho * Ps)) ^ 0.35
End Function

'************************************************************************
'*********************** Пользовательские функции ***********************
'************************************************************************
'Dim lfl As CLFL
'Function М404_НПКР_ГГ_R(M_МассаПаров_кг As Double, _
'                        rho_ПлотностьПаров_кг__м3, _
'                        Cнкпр As Double) As Double
'    Set lfl = New CLFL
'    М404_НПКР_ГГ_R = lfl.f_GG_R(M_МассаПаров_кг, rho_ПлотностьПаров_кг__м3, Cнкпр)
'End Function
'
'Function М404_НПКР_ГГ_Z(M_МассаПаров_кг As Double, _
'                        rho_ПлотностьПаров_кг__м3, _
'                        Cнкпр As Double) As Double
'    Set lfl = New CLFL
'    М404_НПКР_ГГ_Z = lfl.f_GG_R(M_МассаПаров_кг, rho_ПлотностьПаров_кг__м3, Cнкпр)
'End Function
'
'Function М404_НПКР_ЛВЖ_R(M_МассаПаров_кг As Double, _
'                         T_ВремяИспарения_сек As Double, _
'                         rho_ПлотностьПаров_кг__м3 As Double, _
'                         Ps_ДавлНасПаров_кг__м3 As Double, _
'                         Cнкпр As Double) As Double
'    Set lfl = New CLFL
'    М404_НПКР_ЛВЖ_R = lfl.f_FL_R(M_МассаПаров_кг, T_ВремяИспарения_сек, _
'                             rho_ПлотностьПаров_кг__м3, Ps_ДавлНасПаров_кг__м3, _
'                             Cнкпр)
'End Function
'
'Function М404_НПКР_ЛВЖ_Z(M_МассаПаров_кг As Double, _
'                         T_ВремяИспарения_сек As Double, _
'                         rho_ПлотностьПаров_кг__м3 As Double, _
'                         Ps_ДавлНасПаров_кг__м3 As Double, _
'                         Cнкпр As Double) As Double
'    Set lfl = New CLFL
'    М404_НПКР_ЛВЖ_Z = lfl.f_FL_Z(M_МассаПаров_кг, T_ВремяИспарения_сек, _
'                             rho_ПлотностьПаров_кг__м3, Ps_ДавлНасПаров_кг__м3, _
'                             Cнкпр)
'End Function

