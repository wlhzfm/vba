VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'значение удельной теплоты сгорания для типичных углеводородов
Const Eud0# = 44 ' МДж/кг
' атмосферное давление, Па
Const const_P0 = 101325 ' Па
' скорость звука в воздухе С  (обычно принимается равной 340 м/с)
Const const_C0 = 340 ' м/с

' Классы горючих веществ
Enum En_classMatter
    класс_1 = 1
    класс_2 = 2
    класс_3 = 3
    класс_4 = 4
End Enum

' Классификация окружающего пространства по степени загроможденности
Enum En_classSpaceClutter
    класс_I = 1
    класс_II = 2
    класс_III = 3
    класс_IV = 4
End Enum

' Ожидаемый режим сгорания облака
' Классы режима сгорания облака
Enum En_combustionMode
    режим_класс_1 = 1
    режим_класс_2 = 2
    режим_класс_3 = 3
    режим_класс_4 = 4
    режим_класс_5 = 5
    режим_класс_6 = 6
End Enum

' sigma - степень расширения продуктов сгорания (для газопаровоздушных
' смесей допускается принимать равным 7, для пылевоздушных смесей 4)
Enum En_sigma
    газопаровоздушнаяСмесь = 7
    пылевоздушнаяСмесь = 4
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
    R = 1 ' строка 1
    tbl(R, 1) = режим_класс_1
        tbl(R, 2) = режим_класс_1
            tbl(R, 3) = режим_класс_2
                tbl(R, 4) = режим_класс_3
    R = 2 ' строка 2
    tbl(R, 1) = режим_класс_1
        tbl(R, 2) = режим_класс_2
            tbl(R, 3) = режим_класс_3
                tbl(R, 4) = режим_класс_4
    R = 3 ' строка 3
    tbl(R, 1) = режим_класс_2
        tbl(R, 2) = режим_класс_3
            tbl(R, 3) = режим_класс_4
                tbl(R, 4) = режим_класс_5
    R = 4 ' строка 4
    tbl(R, 1) = режим_класс_3
        tbl(R, 2) = режим_класс_4
            tbl(R, 3) = режим_класс_5
                tbl(R, 4) = режим_класс_6
    
    get_combustionModeTbl = tbl
End Function


' Исходными данными для расчета параметров волн давления
' при сгорании облака являются:
' - вид горючего вещества, содержащегося в облаке
' - концентрация горючего вещества в смеси
' - стехиометрическая концентрация горючего вещества с воздухом
' - масса горючего вещества Мт, содержащегося в облаке,
'   с концентрацией между нижним и верхним концентрационным
'   пределом распространения пламени
'   Допускается величину  Мт принимать равной массе горючего
'   вещества, содержащегося в облаке, с учетом коэффициента Z
'   участия горючего вещества во взрыве. При отсутствии данных
'   коэффициент Z может быть принят равным 0,1.
'   При струйном стационарном истечении горючего газа величину  Мт
'   следует рассчитывать с учетом стационарного распределения
'   концентраций горючего газа в струе
' - удельная теплота сгорания горючего вещества Eуд
' - скорость звука в воздухе С0 (обычно принимается равной 340 м/с)
' - информация о степени загроможденности окружающего пространства
' - эффективный энергозапас горючей смеси Е, который определяется по формуле



' эффективный энергозапас горючей смеси E
' (П3.36)
' ground - на поверхности земли
'   При расчете параметров сгорания облака, расположенного на поверхности
'   земли, величина эффективного энергозапаса удваивается
Function get_E(Mt#, Cg#, Cct#, Eud#, Optional ground As Boolean = True)
Dim E#
    If (Cг <= Cст) Then
      E = Mt * Eud
    Else
      E = Mt * Eud * Cct / Cg
    End If
    If (ground) Then
        E = 2 * E
    End If
    get_E = E * 10 ^ 6
End Function

' При оценке масштабов поражения волнами давления должно учитываться
' различие химических соединений по теплоте сгорания, используемой
' для расчета полного запаса энерговыделения. Для типичных углеводородов
' принимается в расчет значение удельной теплоты сгорания  Eуд0 = 44 МДж/кг.
' Для иных горючих веществ в расчетах используется удельное энерговыделение
' Eуд=betta*Eуд0.
' Здесь betta - корректировочный параметр. Для условно выделенных классов
' горючих веществ величины параметра betta представлены в таблице П3.
Function get_Eud(betta)
    get_Eud = Eud0 * betta
End Function


' =========================================================================
' определение ожидаемого режима сгорания облака
' =========================================================================
' Ожидаемый режим сгорания облака определяется с помощью таблицы П3.3,
' в зависимости от класса горючего вещества и класса загроможденности
' окружающего пространства
Function get_combustionMode(classSpaceClutter As En_classSpaceClutter, _
                            classMater As En_classMatter) _
                            As En_combustionMode
Dim combustionModeTbl() As En_combustionMode
    combustionModeTbl = get_combustionModeTbl()
    get_combustionMode = combustionModeTbl(classMater, classSpaceClutter)
End Function
'Sub test_get_combustionMode()
'    Debug.Print get_combustionMode(класс_II, класс_3)
'End Sub

' u скорость фронта пламени
' (П3.37) (П3.38)
Function get_u(combustionMode As En_combustionMode, Optional M# = 0)
Dim k1%, k2%, u#
    k1 = 43
    k2 = 26
    If (combustionMode = режим_класс_1) Then
    ' класс 1 - детонация или горение со скоростью фронта пламени
    '           500 м/с и более;
        get_u = 500
    ElseIf (combustionMode = режим_класс_2) Then
    ' При определении максимальной скорости фронта пламени для
    ' режимов сгорания 2-4 классов дополнительно рассчитывается
    ' видимая скорость фронта пламени по соотношению (П3.37).
    ' В том случае, если полученная величина больше
    ' максимальной скорости, соответствующей данному классу,
    ' она принимается по формуле (П.3.37)
    
    ' класс 2 - дефлаграция, скорость фронта пламени 300-500 м/с;
        u = k1 * M ^ (1 / 6)
        get_u = Max(500, u)
    ElseIf (combustionMode = режим_класс_3) Then
    ' класс 3 - дефлаграция, скорость фронта пламени 200-300 м/с;
        u = k1 * M ^ (1 / 6)
        get_u = Max(300, u)
    ElseIf (combustionMode = режим_класс_4) Then
    ' класс 4 - дефлаграция, скорость фронта пламени 150-200 м/с;
        u = k1 * M ^ (1 / 6)
        get_u = Max(200, u)
    ElseIf (combustionMode = режим_класс_5) Then
    ' класс 5 - дефлаграция, скорость фронта пламени определяется по
    '           формуле:(П3.37)
        get_u = k1 * M ^ (1 / 6)
    ElseIf (combustionMode = режим_класс_6) Then
    ' класс 6 - дефлаграция, скорость фронта пламени определяется
    '           по формуле:(П3.38)
        get_u = k2 * M ^ (1 / 6)
    End If
End Function


' =========================================================================
' расчет максимального избыточного давления и импульса
' фазы сжатия воздушных волн давления для различных режимов
' =========================================================================
' Параметры воздушных волн давления (избыточное давление dР и импульс фазы
' сжатия I+) в зависимости от расстояния от центра облака рассчитываются
' исходя из ожидаемого режима сгорания облака
Function get_dP#(R#, combustionMode As En_combustionMode, _
                 E#, u#, _
                 Optional P0# = const_P0, _
                 Optional sigma As En_sigma = газопаровоздушнаяСмесь)
Dim Px#, ln_Px#, Rx#, ln_Rx#, Rкр1#
Dim dP#
    Rx = get_Rx(R, E, P0)
    
    If (combustionMode = режим_класс_1) Then
        ' Величины безразмерного давления Px и импульс фазы сжатия Ix определяются
        '   по формулам (для газопаровоздушных смесей): (П3.40)  (П3.41)
        ' Формулы (П3.40, П3.41) справедливы для значений  Rx > 0,2. В случае,
        '   если  Rx<0,2, то Px = равно 18, а в формулу (П3.41) вместо Rx подставляется
        '   величина  = 0,14.
        If (Rx < 0.2) Then
            Px = 18
        Else
            ln_Rx = Ln(Rx)
            ln_Px = -1.124 - 1.66 * (ln_Rx) + 0.26 * (ln_Rx) ^ 2
            Px = Exp(1) ^ ln_Px
        End If
        ' Размерные величины избыточного давления и импульса фазы сжатия
        ' определяются по формулам: (П3.42) (П3.43)
        dP = Px * P0
    Else
        ' Рассчитываются величины безразмерного давления (Px1) и импульса фазы
        ' сжатия Ix1 по формулам: (П3.44) (П3.45) (П3.46)
        ' Формулы (П3.44), (П3.45) справедливы для значений Rx больших величины
        ' Rкр1 = 0,34, в случае, если  Rx<Rкр1, в формулы (П3.44), (П3.45)
        ' вместо Rx подставляется величина Rкр1.
        Rкр1 = 0.34
        If (Rx < Rкр1) Then
            Rx = Rкр1
        End If
        ' В случае дефлаграции пылевоздушного облака величина эффективного
        ' энергозапаса умножается на коэффициент  (sigma-1)/sigma.
        If (sigma = пылевоздушнаяСмесь And combustionMode <> режим_класс_1) Then
            E = E * (sigma - 1) / sigma
        End If
        Px1 = (u ^ 2 / const_C0 ^ 2) * (sigma - 1) / sigma * (0.83 / Rx - 0.14 / Rx ^ 2)
        ' Размерные величины избыточного давления и импульса фазы сжатия
        ' определяются по формулам (П3.42), (П3.43). При этом в формулы (П3.42), (П3.43)
        ' вместо Px и Ix подставляются величины Px1 и Ix1.
        dP = Px1 * P0
    End If
    '
    get_dP = dP
End Function
Function get_I#(R#, combustionMode As En_combustionMode, _
                E#, u#, _
                Optional P0# = const_P0, _
                Optional sigma As En_sigma = газопаровоздушнаяСмесь)
Dim Ix#, ln_Ix#, Rx#, ln_Rx#, Rкр1#, W#
Dim i#
    Rx = get_Rx(R, E, P0)
    
    If (combustionMode = режим_класс_1) Then
        ' Величины безразмерного давления Px и импульс фазы сжатия Ix определяются
        '   по формулам (для газопаровоздушных смесей): (П3.40)  (П3.41)
        ' Формулы (П3.40, П3.41) справедливы для значений  Rx > 0,2. В случае,
        '   если  Rx<0,2, то Px = равно 18, а в формулу (П3.41) вместо Rx подставляется
        '   величина  = 0,14.
        If (Rx < 0.2) Then
            Rx = 0.14
        End If
        ln_Rx = Ln(Rx)
        ln_Ix = -3.4217 - 0.898 * (ln_Rx) - 0.0096 * (ln_Rx) ^ 2
        Ix = Exp(1) ^ ln_Ix

        ' Размерные величины избыточного давления и импульса фазы сжатия
        ' определяются по формулам: (П3.42) (П3.43)
        i = Ix * P0 ^ (2 / 3) * E ^ (1 / 3) / const_C0
    Else
        ' Рассчитываются величины безразмерного давления (Px1) и импульса фазы
        ' сжатия Ix1 по формулам: (П3.44) (П3.45) (П3.46)
        ' Формулы (П3.44), (П3.45) справедливы для значений Rx больших величины
        ' Rкр1 = 0,34, в случае, если  Rx<Rкр1, в формулы (П3.44), (П3.45)
        ' вместо Rx подставляется величина Rкр1.
        Rкр1 = 0.34
        If (Rx < Rкр1) Then
            Rx = Rкр1
        End If
        W = u / const_C0 * (sigma - 1) / sigma
        ' В случае дефлаграции пылевоздушного облака величина эффективного
        ' энергозапаса умножается на коэффициент  (sigma-1)/sigma.
        If (sigma = пылевоздушнаяСмесь And combustionMode <> режим_класс_1) Then
            E = E * (sigma - 1) / sigma
        End If
        Ix1 = W * (1 - 0.4 * W) * (0.06 / Rx + 0.01 / Rx ^ 2 - 0.0025 / Rx ^ 3)
        ' Размерные величины избыточного давления и импульса фазы сжатия
        ' определяются по формулам (П3.42), (П3.43). При этом в формулы (П3.42), (П3.43)
        ' вместо Px и Ix подставляются величины Px1 и Ix1.
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
' определение дополнительных характеристик взрывной нагрузки
' =========================================================================




' =========================================================================
' оценка поражающего воздействия
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


