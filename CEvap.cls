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
    ' Температура, С
    tbl_temp(0) = 10
    tbl_temp(1) = 15
    tbl_temp(2) = 20
    tbl_temp(3) = 30
    tbl_temp(4) = 35
    ' Скорость движения воздуха, м/с
    tbl_v(0) = 0
    tbl_v(1) = 0.1
    tbl_v(2) = 0.2
    tbl_v(3) = 0.5
    tbl_v(4) = 1#
    ' Коэффициент
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

' СУГ
Function getVaporLPG(Mm As Double, Lисп As Double, T0 As Double, Tж As Double, _
                    lambda_тв As Double, T As Double, a As Double, lambda_в As Double, _
                    v As Double, d As Double, nu_в As Double) As Double
Dim Re As Double
Dim M As Double
    Re = v * d / nu_в
    M = Mm / Lисп * (T0 - Tж) * (2 * lambdа_тв * (T / (pi * a)) ^ 0.5 + 5.1 * (Re * lambda_в * T) ^ 0.5 / T)
    '
    getVaporLPG = M
End Function



'************************************************************************
'*********************** Пользовательские функции ***********************
'************************************************************************
'
'Dim evap As Cevap
'
'Function ГОСТ_ИнтенсивностьИспаренияЛВЖ_кг__м2_с(Mm_МолекМасса As Double, _
'                                     Ps_ДавлНасПар_кПа As Double, _
'                                     СкоростьВетра_м As Double, _
'                                     ТемпВозд_градС As Double) As Double
'Dim W As Double
'    Set evap = New Cevap
'    W = evap.getVapor(СкоростьВетра_м, ТемпВозд_градС, Mm_МолекМасса, Ps_ДавлНасПар_кПа)
'    '
'    ГОСТ_ИнтенсивностьИспаренияЛВЖ_кг__м2_с = W
'End Function
'
'Function ГОСТ_ИспарениеЛВЖ_Масса_кг(W_Интенсивность_кг__м2_с As Double, _
'                                S_Площадь_м2 As Double, _
'                                t_Время_сек As Double)
'    Set evap = New Cevap
'    ГОСТ_ИспарениеЛВЖ_Масса_кг = evap.getVaporM(S_Площадь_м2, t_Время_сек, W_Интенсивность_кг__м2_с)
'End Function
'
'Function ГОСТ_ДавленНасыщПаров_кПа(a As Double, _
'                                   B As Double, _
'                                   Ca As Double, _
'                                   t_Температура_С As Double) As Double
'    Set evap = New Cevap
'    ГОСТ_ДавленНасыщПаров_кПа = evap.getPs(a, B, Ca, t_Температура_С)
'End Function




'Mm 102, 2
'Константы Антуана
'a 7, 54424
'B 2629, 65
'Ca 384, 195
'temp    10  C
'v   1   м/с
'S   100 м2
'time    3600    сек
'
'Ps  5,01    кПа
'W   5,06E-04    кг/(м2*с)
'M   182,3   кг



