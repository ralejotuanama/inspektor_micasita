Attribute VB_Name = "modsis"
Option Explicit

Public Sub modsis_gs_Carga_TipGrp(p_Combo As ComboBox)
   'Tipo de Grupo de Parámetros
   'Rubro 038
   p_Combo.Clear
   
   p_Combo.AddItem "RESERVADO PARA ADMINISTRADOR"
   p_Combo.ItemData(p_Combo.NewIndex) = 1
   
   p_Combo.AddItem "GESTIONADO POR USUARIO"
   p_Combo.ItemData(p_Combo.NewIndex) = 2
   
   p_Combo.ListIndex = -1
End Sub

Public Sub modsis_gs_Carga_TipPar(p_Combo As ComboBox, ByVal p_FlgOpc As Integer)
   'Tipo de Parámetros
   'Rubro 010
   'Cuando p_FlgOpc sea
   '  1  :  Porcentaje / Cantidad
   '  2  :  Porcentaje / Cantidad / Descriptivo

   p_Combo.Clear
   
   p_Combo.AddItem "PORCENTAJE"
   p_Combo.ItemData(p_Combo.NewIndex) = 1
   
   p_Combo.AddItem "CANTIDAD"
   p_Combo.ItemData(p_Combo.NewIndex) = 2
   
   If p_FlgOpc = 2 Then
      p_Combo.AddItem "DESCRIPTIVO"
      p_Combo.ItemData(p_Combo.NewIndex) = 3
   End If
   
   p_Combo.ListIndex = -1
End Sub

Public Sub modsis_gs_Carga_TipValPar(p_Combo As ComboBox)
   'Tipo de Valor para Parámetros
   'Rubro 043
   p_Combo.Clear
   
   p_Combo.AddItem "VALOR FIJO"
   p_Combo.ItemData(p_Combo.NewIndex) = 1
   
   p_Combo.AddItem "RANGO DE VALORES"
   p_Combo.ItemData(p_Combo.NewIndex) = 2
   
   p_Combo.ListIndex = -1
End Sub

Public Sub modsis_gs_Carga_TipBus(p_Combo As ComboBox)
   p_Combo.Clear
   
   p_Combo.AddItem "1: POR DOC. IDENTIDAD"
   p_Combo.ItemData(p_Combo.NewIndex) = 1
   
   p_Combo.AddItem "2: POR NUMERO SOLICITUD"
   p_Combo.ItemData(p_Combo.NewIndex) = 2
   
   p_Combo.ListIndex = -1
End Sub

Public Sub modsis_gs_Carga_TipBus_1(p_Combo As ComboBox)
   p_Combo.Clear
   
   p_Combo.AddItem "1: POR DOC. IDENTIDAD"
   p_Combo.ItemData(p_Combo.NewIndex) = 1
   
   p_Combo.AddItem "2: POR NUMERO OPERACION"
   p_Combo.ItemData(p_Combo.NewIndex) = 2
   
   p_Combo.ListIndex = -1
End Sub

Public Sub modsis_gs_Carga_Tip_TipCam(p_Combo As ComboBox)
   p_Combo.Clear
   
   p_Combo.AddItem "COMERCIAL"
   p_Combo.ItemData(p_Combo.NewIndex) = 1
   
   p_Combo.ListIndex = -1
End Sub

'Cronogramas Antiguos (Pasados a este módulo por MAIP (07-04-2010)

Public Sub gs_Calcul_COFIDE(p_Arregl() As modcal_g_est_CuoCof, p_TasInt As Double, p_TasCom As Double, p_FlgCuo As Integer, p_FecDes As String, p_NumCuo As Integer, p_MtoPre As Double, p_Tramo As Integer, p_PerGra As Integer, Optional p_UltVct As String)
   'p_Arregl   -  Estructura para generar cronograma
   'p_TasInt   -  Tasa de Interes Anual
   'p_TasCom   -  Tasa de Comisión COFIDE
   'p_DiaPag   -  Día de Pago
   'p_FlgCuo   -  Flag de Cuota Extraordinaria (Multiplicar x 1 ó 2 en los meses de Julio y Diciembre)
   'p_FecDes   -  Fecha de Desembolso
   'p_NumCuo   -  Número de Cuotas
   'p_MtoPre   -  Monto Préstamo
   'p_Tramo    -  1 - Tramo No Concesional (Mensual) / 2 - Tramo Concesional (Semestral)
   'p_PerGra   -  Período de Gracia
   'p_UltVct   -  Fecha de Ultimo Vencimiento de Cronograma No-Concesional
   
   Dim r_int_Contad        As Integer
   Dim r_int_Cont01        As Integer
   Dim r_str_DiaPag        As String
   Dim r_int_MesIni        As Integer
   Dim r_int_AnoIni        As Integer
   Dim r_str_FecVct        As String
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   Dim r_str_VctIni        As String
   Dim r_str_FinMes        As String
   Dim r_dbl_FacMen        As Double
   Dim r_dbl_TasMen        As Double
   Dim r_int_CuoExt        As Integer
   Dim r_dbl_AcuFac        As Double
   Dim r_dbl_ImpPre        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_CupMen        As Double
   Dim r_dbl_Princi        As Double
   Dim r_dbl_AjuPre        As Double
   Dim r_dbl_CuoIni        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
   Dim r_dbl_Capita        As Double
   Dim r_dbl_ComMen        As Double
   Dim r_str_VctHab        As String
   Dim r_int_DiaAdc        As Integer
   
   'Obteniendo si habran Cuotas Extraordinarias
   r_int_CuoExt = p_FlgCuo
   
   'Calculando Tasa de Interés Mensual y Comisión Mensual
   r_dbl_TasMen = ff_IntMen(p_TasInt)
   r_dbl_ComMen = ff_IntMen(p_TasCom)
   
   'Para hallar diferencia de días entre la Fecha de Desembolso y el Ultimo Día Hábil del Mes de Desembolso
   'Sólo válido para Tramo No Concesional Sin Período de Gracia
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
   r_str_VctHab = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
   r_int_DifDia = CInt(CDate(r_str_VctHab) - CDate(p_FecDes))
   
   r_int_DiaAdc = 0
   If p_Tramo = 1 And p_PerGra = 0 And r_int_DifDia > 0 Then
      'Sólo válido para Tramo No Concesional y sin Período de Gracia
      'Si la diferencia entre la Fecha de Desembolso y Ultimo Día Habil del Mes de Desembolso es Mayor a Cero
      'Se guardan los días de diferencia y se asume como Fecha de Inicio de Cronograma
      'el Ultimo Día Hábil del Mes de Desembolso
      r_str_VctIni = r_str_VctHab
      r_int_DiaAdc = r_int_DifDia
   Else
      'Fecha de Inicio de Cronograma siempre será la Fecha de Desembolso
      r_str_VctIni = p_FecDes
   End If
   
   'Obteniendo Primera Fecha de Vencimiento
   r_str_DiaPag = Day(CDate(p_FecDes))
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   'Si Tramo es No Concesional o es Tramo Concesional y hay Período de Gracia
   'Se incrementa el Mes en 1
   If (p_Tramo = 1) Or (p_PerGra > 0 And p_Tramo = 2) Then
      r_int_MesIni = r_int_MesIni + 1
      
      If r_int_MesIni = 13 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = 1
      End If
   Else
      'Si el Tramo es Concesional y no hay Período de Gracia se incrementa el mes en 6
      r_int_MesIni = r_int_MesIni + 6
         
      If r_int_MesIni > 12 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = r_int_MesIni - 12
      End If
   End If
   
   'Calculando Factor de Ajuste de Saldo Capital y Nuevo Saldo Capital
   r_dbl_AjuPre = (1 + ((1 + p_TasInt / 100) * (1 + p_TasCom / 100) - 1)) ^ (1 / 12) - 1
   r_dbl_ImpPre = p_MtoPre
   r_dbl_SalCap = r_dbl_ImpPre
   
   'Inicializando Arreglo
   If p_Tramo = 2 And p_PerGra > 0 Then
      Select Case p_PerGra
         Case 1 To 5:   p_NumCuo = p_NumCuo + 1
         Case 6 To 11:  p_NumCuo = p_NumCuo
         Case Is > 11:  p_NumCuo = p_NumCuo - 1
      End Select
   End If
   ReDim p_Arregl(p_NumCuo)
   
   'Si Período de Gracia es Mayor a Cero
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      
      If p_Tramo = 1 Then        'Tramo No Concesional
         For r_int_Contad = 1 To p_PerGra
            r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
            r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
            
            r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
            r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
            r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
            r_dbl_Comisi = CDbl(Format((1 + r_dbl_ComMen) ^ (r_int_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
            
            r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
            
            p_Arregl(r_int_Contad).CuoCof_FecVct = r_str_FecVct
            p_Arregl(r_int_Contad).CuoCof_DifDia = r_int_DifDia
            p_Arregl(r_int_Contad).CuoCof_AcuDia = r_int_AcuDia
            p_Arregl(r_int_Contad).CuoCof_Capita = 0
            p_Arregl(r_int_Contad).CuoCof_Intere = 0
            p_Arregl(r_int_Contad).CuoCof_Comisi = r_dbl_Comisi
            p_Arregl(r_int_Contad).CuoCof_IntTot = r_dbl_Intere + r_dbl_Comisi
            p_Arregl(r_int_Contad).CuoCof_CupMen = r_dbl_CupMen
            p_Arregl(r_int_Contad).CuoCof_ValCuo = r_dbl_Comisi
            p_Arregl(r_int_Contad).CuoCof_SalCap = r_dbl_SalCap
            p_Arregl(r_int_Contad).CuoCof_IntCap = r_dbl_Intere
            
            r_str_VctIni = r_str_FecVct
            
            r_int_MesIni = r_int_MesIni + 1
            If r_int_MesIni = 13 Then
               r_int_AnoIni = r_int_AnoIni + 1
               r_int_MesIni = 1
            End If
         Next r_int_Contad
         
         r_dbl_ImpPre = CDbl(Format(r_dbl_SalCap, "########0.00"))
      Else                    'Tramo Concesional
         For r_int_Contad = 1 To p_PerGra
            r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
            r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
            
            r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
            r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            r_str_VctIni = r_str_FecVct
            
            r_int_MesIni = r_int_MesIni + 1
            If r_int_MesIni = 13 Then
               r_int_AnoIni = r_int_AnoIni + 1
               r_int_MesIni = 1
            End If
         Next r_int_Contad
         
         r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_AcuDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap + r_dbl_Intere, "########0.00"))
         r_dbl_ImpPre = r_dbl_SalCap
         
         p_Arregl(1).CuoCof_FecVct = r_str_FecVct
         p_Arregl(1).CuoCof_DifDia = r_int_DifDia
         p_Arregl(1).CuoCof_AcuDia = r_int_AcuDia
         p_Arregl(1).CuoCof_Capita = 0
         p_Arregl(1).CuoCof_Intere = 0
         p_Arregl(1).CuoCof_Comisi = 0
         p_Arregl(1).CuoCof_IntTot = r_dbl_Intere
         p_Arregl(1).CuoCof_CupMen = 0
         p_Arregl(1).CuoCof_ValCuo = 0
         p_Arregl(1).CuoCof_SalCap = r_dbl_SalCap
         p_Arregl(1).CuoCof_IntCap = r_dbl_Intere
         
         
         r_int_MesIni = r_int_MesIni + 5
         If r_int_MesIni > 12 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = r_int_MesIni - 12
         End If
         
         p_PerGra = 1
      End If
   End If

   'Calculando Fechas de Vencimiento y Factores Mensuales
   r_int_AcuDia = 0
   r_dbl_AcuFac = 0
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
      r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
      
      'Si es Tramo Concesional validar que la Fecha de Vencimiento no puede superar el período total del préstamo
      If p_Tramo = 2 Then
         If CDate(r_str_FecVct) > CDate(p_UltVct) Then
            r_str_FecVct = p_UltVct
         End If
      End If
      
      r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      'Calculando Factor Mensual
      If (r_int_MesIni = 7 Or r_int_MesIni = 12) And r_int_CuoExt = 2 And p_Tramo = 1 Then
         r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30) * 2
      Else
         r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30)
      End If
      
      'Acumulando Factor
      r_dbl_AcuFac = r_dbl_AcuFac + CDbl(Format(r_dbl_FacMen, "###0.000000000"))
      
      'Pasando datos a Arreglo
      p_Arregl(r_int_Contad).CuoCof_FecVct = r_str_FecVct
      p_Arregl(r_int_Contad).CuoCof_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCof_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCof_Factor = r_dbl_FacMen
      
      r_str_VctIni = r_str_FecVct
      If p_Tramo = 1 Then     'Tramo No Concesional (Mensual)
         r_int_MesIni = r_int_MesIni + 1
         If r_int_MesIni = 13 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = 1
         End If
      Else                    'Tramo Concesional (Semestral)
         r_int_MesIni = r_int_MesIni + 6
         If r_int_MesIni > 12 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = r_int_MesIni - 12
         End If
      End If
   Next r_int_Contad
   
   'Calculando Valor Cuota Inicial
   r_dbl_CuoIni = r_dbl_ImpPre / r_dbl_AcuFac
   r_dbl_SalCap = r_dbl_ImpPre
   
   'Proceso de Iteración para Calcular Cupón Mensual
   For r_int_Cont01 = 1 To 1500
      'Calculando Cupón Mensual
      For r_int_Contad = p_PerGra + 1 To p_NumCuo
         If (Month(CDate(p_Arregl(r_int_Contad).CuoCof_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCof_FecVct)) = 12) And r_int_CuoExt = 2 And p_Tramo = 1 Then
            r_dbl_CupMen = r_dbl_CuoIni * 2
         Else
            r_dbl_CupMen = r_dbl_CuoIni
         End If
         
         r_dbl_Princi = r_dbl_SalCap
         
         r_dbl_Intere = (1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCof_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
         r_dbl_Comisi = (1 + r_dbl_ComMen) ^ (p_Arregl(r_int_Contad).CuoCof_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
         
         r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere - r_dbl_Comisi
         
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         p_Arregl(r_int_Contad).CuoCof_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCof_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCof_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCof_IntTot = r_dbl_Intere + r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCof_CupMen = r_dbl_CupMen
         p_Arregl(r_int_Contad).CuoCof_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCof_SalCap = r_dbl_SalCap
         p_Arregl(r_int_Contad).CuoCof_IntCap = 0
      Next r_int_Contad
      
      'Ajustes de Saldo Capital y Valor de Cuota Inicial
      If r_dbl_SalCap > 0# Then
         Select Case r_dbl_SalCap
            Case 0 To 0.5:    r_dbl_CuoIni = r_dbl_CuoIni + 0.00001
            Case 0.5 To 1:    r_dbl_CuoIni = r_dbl_CuoIni + 0.0001
            Case 1 To 2:      r_dbl_CuoIni = r_dbl_CuoIni + 0.001
            Case 2 To 5:      r_dbl_CuoIni = r_dbl_CuoIni + 0.01
            Case Is > 5:      r_dbl_CuoIni = r_dbl_CuoIni + 0.1
         End Select

         r_dbl_SalCap = r_dbl_ImpPre
      ElseIf r_dbl_SalCap < 0 Then
         Select Case r_dbl_SalCap
            Case Is < -5:     r_dbl_CuoIni = r_dbl_CuoIni - 0.1
            Case -5 To -2:    r_dbl_CuoIni = r_dbl_CuoIni - 0.01
            Case -2 To -1:    r_dbl_CuoIni = r_dbl_CuoIni - 0.001
            Case -1 To -0.5:  r_dbl_CuoIni = r_dbl_CuoIni - 0.0001
            Case -0.5 To -0#: r_dbl_CuoIni = r_dbl_CuoIni - 0.00001
         End Select
         r_dbl_SalCap = r_dbl_ImpPre
      ElseIf r_dbl_SalCap = 0# Then
         Exit For
      End If
   Next r_int_Cont01
   
   'Redondeando Cuota
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni, "######0.00"))
   r_dbl_SalCap = r_dbl_ImpPre

   'Calculando Cupón Mensual Redondeado
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      If (Month(CDate(p_Arregl(r_int_Contad).CuoCof_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCof_FecVct)) = 12) And r_int_CuoExt = 2 Then
         r_dbl_CupMen = r_dbl_CuoIni * 2
      Else
         r_dbl_CupMen = r_dbl_CuoIni
      End If
      
      r_dbl_Princi = r_dbl_SalCap
      
      r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCof_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
      r_dbl_Comisi = CDbl(Format((1 + r_dbl_ComMen) ^ (p_Arregl(r_int_Contad).CuoCof_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
      
      r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      p_Arregl(r_int_Contad).CuoCof_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCof_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCof_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCof_IntTot = r_dbl_Intere + r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCof_CupMen = r_dbl_CupMen
      p_Arregl(r_int_Contad).CuoCof_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCof_SalCap = r_dbl_SalCap
      p_Arregl(r_int_Contad).CuoCof_IntCap = 0
   Next r_int_Contad

   'Recalculando Ultima Cuota ya Redondeada
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni + r_dbl_SalCap, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(p_NumCuo).CuoCof_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
   r_dbl_Comisi = CDbl(Format((1 + r_dbl_ComMen) ^ (p_Arregl(p_NumCuo).CuoCof_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
   
   r_dbl_Capita = r_dbl_CuoIni - r_dbl_Intere - r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCof_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCof_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCof_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo).CuoCof_IntTot = r_dbl_Intere + r_dbl_Comisi
   p_Arregl(p_NumCuo).CuoCof_CupMen = r_dbl_CupMen
   p_Arregl(p_NumCuo).CuoCof_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   p_Arregl(p_NumCuo).CuoCof_SalCap = r_dbl_SalCap
   p_Arregl(p_NumCuo).CuoCof_IntCap = 0
   
   'Para Ajustar Primera Cuota cuando Fecha de Desembolso no coincide con Ultimo Día Habil del Mes
   'Primera Cuota debe ser Mayor que las demás
   If p_PerGra = 0 And p_Tramo = 1 And r_int_DiaAdc > 0 Then
      r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ ((p_Arregl(1).CuoCof_DifDia + r_int_DiaAdc) / 30) * p_MtoPre - p_MtoPre, "######0.00"))
      r_dbl_Comisi = CDbl(Format((1 + r_dbl_ComMen) ^ ((p_Arregl(1).CuoCof_DifDia + r_int_DiaAdc) / 30) * p_MtoPre - p_MtoPre, "######0.00"))
      
      p_Arregl(1).CuoCof_Intere = r_dbl_Intere
      p_Arregl(1).CuoCof_Comisi = r_dbl_Comisi
      p_Arregl(1).CuoCof_IntTot = r_dbl_Intere + r_dbl_Comisi
      p_Arregl(1).CuoCof_ValCuo = p_Arregl(1).CuoCof_Capita + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Calcul_Cliente_1(p_Arregl() As modcal_g_est_CuoCli, p_TasInt As Double, p_FlgCuo As Integer, p_FecDes As String, p_NumCuo As Integer, p_MtoPre As Double, p_Tramo As Integer, p_PerGra As Integer, p_UltVct As String)
   'p_Arregl   -  Estructura para generar cronograma
   'p_TasInt   -  Tasa de Interes Anual
   'p_FlgCuo   -  Flag de Cuota Extraordinaria (Multiplicar x 1 ó 2 en los meses de Julio y Diciembre)
   'p_FecDes   -  Fecha de Desembolso
   'p_NumCuo   -  Número de Cuotas
   'p_MtoPre   -  Monto Préstamo
   'p_Tramo    -  1 - Tramo No Concesional (Mensual) / 2 - Tramo Concesional (Semestral)
   'p_PerGra   -  Período de Gracia
   'p_UltVct   -  Fecha de Ultimo Vencimiento de Cronograma No-Concesional
   
   Dim r_int_Contad        As Integer
   Dim r_int_Cont01        As Integer
   Dim r_str_DiaPag        As String
   Dim r_int_MesIni        As Integer
   Dim r_int_AnoIni        As Integer
   Dim r_str_FecVct        As String
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   Dim r_str_VctIni        As String
   Dim r_str_FinMes        As String
   Dim r_dbl_FacMen        As Double
   Dim r_dbl_TasMen        As Double
   Dim r_int_CuoExt        As Integer
   Dim r_dbl_AcuFac        As Double
   Dim r_dbl_ImpPre        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_CupMen        As Double
   Dim r_dbl_Princi        As Double
   Dim r_dbl_AjuPre        As Double
   Dim r_dbl_CuoIni        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Capita        As Double
   Dim r_str_VctHab        As String
   Dim r_int_DiaAdc        As Integer
   
   'Obteniendo si habran Cuotas Extraordinarias
   r_int_CuoExt = p_FlgCuo
   
   'Calculando Tasa de Interés Mensual y Comisión Mensual
   r_dbl_TasMen = ff_IntMen(p_TasInt)
   
   'Para hallar diferencia de días entre la Fecha de Desembolso y el Ultimo Día Hábil del Mes de Desembolso
   'Sólo válido para Tramo No Concesional Sin Período de Gracia
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
   r_str_VctHab = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
   r_int_DifDia = CInt(CDate(r_str_VctHab) - CDate(p_FecDes))
   
   r_int_DiaAdc = 0
   If p_Tramo = 1 And p_PerGra = 0 And r_int_DifDia > 0 Then
      'Sólo válido para Tramo No Concesional y sin Período de Gracia
      'Si la diferencia entre la Fecha de Desembolso y Ultimo Día Habil del Mes de Desembolso es Mayor a Cero
      'Se guardan los días de diferencia y se asume como Fecha de Inicio de Cronograma
      'el Ultimo Día Hábil del Mes de Desembolso
      r_str_VctIni = r_str_VctHab
      r_int_DiaAdc = r_int_DifDia
   Else
      'Fecha de Inicio de Cronograma siempre será la Fecha de Desembolso
      r_str_VctIni = p_FecDes
   End If
   
   'Obteniendo Primera Fecha de Vencimiento
   r_str_DiaPag = Day(CDate(p_FecDes))
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   'Si Tramo es No Concesional o es Tramo Concesional y hay Período de Gracia
   'Se incrementa el Mes en 1
   If (p_Tramo = 1) Or (p_PerGra > 0 And p_Tramo = 2) Then
      r_int_MesIni = r_int_MesIni + 1
      
      If r_int_MesIni = 13 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = 1
      End If
   Else
      'Si el Tramo es Concesional y no hay Período de Gracia se incrementa el mes en 6
      r_int_MesIni = r_int_MesIni + 6
         
      If r_int_MesIni > 12 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = r_int_MesIni - 12
      End If
   End If
   
   'Calculando Factor de Ajuste de Saldo Capital y Nuevo Saldo Capital
   r_dbl_AjuPre = (1 + ((1 + p_TasInt / 100) - 1)) ^ (1 / 12) - 1
   r_dbl_ImpPre = p_MtoPre
   r_dbl_SalCap = r_dbl_ImpPre
   
   'Inicializando Arreglo
   If p_Tramo = 2 And p_PerGra > 0 Then
      Select Case p_PerGra
         Case 1 To 5:   p_NumCuo = p_NumCuo + 1
         Case 6 To 11:  p_NumCuo = p_NumCuo
         Case Is > 11:  p_NumCuo = p_NumCuo - 1
      End Select
   End If
   ReDim p_Arregl(p_NumCuo)
   
   'Si Período de Gracia es Mayor a Cero
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      
      If p_Tramo = 1 Then        'Tramo No Concesional
         For r_int_Contad = 1 To p_PerGra
            r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
            r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
            
            r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
            r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
            r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
            
            r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
            
            p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
            p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
            p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
            p_Arregl(r_int_Contad).CuoCli_Capita = 0
            p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
            p_Arregl(r_int_Contad).CuoCli_CupMen = r_dbl_CupMen
            p_Arregl(r_int_Contad).CuoCli_ValCuo = 0
            p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
            p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
            
            r_str_VctIni = r_str_FecVct
            
            r_int_MesIni = r_int_MesIni + 1
            If r_int_MesIni = 13 Then
               r_int_AnoIni = r_int_AnoIni + 1
               r_int_MesIni = 1
            End If
         Next r_int_Contad
         
         r_dbl_ImpPre = CDbl(Format(r_dbl_SalCap, "########0.00"))
      Else                    'Tramo Concesional
         For r_int_Contad = 1 To p_PerGra
            r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
            r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
            
            r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
            r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            r_str_VctIni = r_str_FecVct
            
            r_int_MesIni = r_int_MesIni + 1
            If r_int_MesIni = 13 Then
               r_int_AnoIni = r_int_AnoIni + 1
               r_int_MesIni = 1
            End If
         Next r_int_Contad
         
         r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_AcuDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap + r_dbl_Intere, "########0.00"))
         r_dbl_ImpPre = r_dbl_SalCap
         
         p_Arregl(1).CuoCli_FecVct = r_str_FecVct
         p_Arregl(1).CuoCli_DifDia = r_int_DifDia
         p_Arregl(1).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(1).CuoCli_Capita = 0
         p_Arregl(1).CuoCli_Intere = r_dbl_Intere
         p_Arregl(1).CuoCli_CupMen = 0
         p_Arregl(1).CuoCli_ValCuo = 0
         p_Arregl(1).CuoCli_SalCap = r_dbl_SalCap
         p_Arregl(1).CuoCli_IntCap = r_dbl_Intere
         
         
         r_int_MesIni = r_int_MesIni + 5
         If r_int_MesIni > 12 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = r_int_MesIni - 12
         End If
         
         p_PerGra = 1
      End If
   End If

   'Calculando Fechas de Vencimiento y Factores Mensuales
   r_int_AcuDia = 0
   r_dbl_AcuFac = 0
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
      r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
      
      'Si es Tramo Concesional validar que la Fecha de Vencimiento no puede superar el período total del préstamo
      If p_Tramo = 2 Then
         If CDate(r_str_FecVct) > CDate(p_UltVct) Then
            r_str_FecVct = p_UltVct
         End If
      End If
      
      r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      'Calculando Factor Mensual
      If (r_int_MesIni = 7 Or r_int_MesIni = 12) And r_int_CuoExt = 2 And p_Tramo = 1 Then
         r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30) * 2
      Else
         r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30)
      End If
      
      'Acumulando Factor
      r_dbl_AcuFac = r_dbl_AcuFac + CDbl(Format(r_dbl_FacMen, "###0.000000000"))
      
      'Pasando datos a Arreglo
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
      
      r_str_VctIni = r_str_FecVct
      If p_Tramo = 1 Then     'Tramo No Concesional (Mensual)
         r_int_MesIni = r_int_MesIni + 1
         If r_int_MesIni = 13 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = 1
         End If
      Else                    'Tramo Concesional (Semestral)
         r_int_MesIni = r_int_MesIni + 6
         If r_int_MesIni > 12 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = r_int_MesIni - 12
         End If
      End If
   Next r_int_Contad
   
   'Calculando Valor Cuota Inicial
   r_dbl_CuoIni = r_dbl_ImpPre / r_dbl_AcuFac
   r_dbl_SalCap = r_dbl_ImpPre
   
   'Proceso de Iteración para Calcular Cupón Mensual
   For r_int_Cont01 = 1 To 1500
      'Calculando Cupón Mensual
      For r_int_Contad = p_PerGra + 1 To p_NumCuo
         If (Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 12) And r_int_CuoExt = 2 And p_Tramo = 1 Then
            r_dbl_CupMen = r_dbl_CuoIni * 2
         Else
            r_dbl_CupMen = r_dbl_CuoIni
         End If
         
         r_dbl_Princi = r_dbl_SalCap
         
         r_dbl_Intere = (1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
         r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_CupMen = r_dbl_CupMen
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         p_Arregl(r_int_Contad).CuoCli_IntCap = 0
      Next r_int_Contad
      
      'Ajustes de Saldo Capital y Valor de Cuota Inicial
      If r_dbl_SalCap > 0# Then
         Select Case r_dbl_SalCap
            'Case 0# To 1:     r_dbl_CuoIni = r_dbl_CuoIni + 0.00001
            Case 0# To 2:     r_dbl_CuoIni = r_dbl_CuoIni + 0.0001
            Case 2 To 5:      r_dbl_CuoIni = r_dbl_CuoIni + 0.001
            Case Is > 5:      r_dbl_CuoIni = r_dbl_CuoIni + 0.01
         End Select

         r_dbl_SalCap = r_dbl_ImpPre
      ElseIf r_dbl_SalCap < 0 Then
         Select Case r_dbl_SalCap
            Case Is < -5:     r_dbl_CuoIni = r_dbl_CuoIni - 0.01
            Case -5 To -2:    r_dbl_CuoIni = r_dbl_CuoIni - 0.001
            Case -2 To -0#:   r_dbl_CuoIni = r_dbl_CuoIni - 0.0001
            'Case -1 To -0#:   r_dbl_CuoIni = r_dbl_CuoIni - 0.00001
         End Select
         r_dbl_SalCap = r_dbl_ImpPre
      ElseIf r_dbl_SalCap = 0# Then
         Exit For
      End If
   Next r_int_Cont01
   
   'Redondeando Cuota
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni, "######0.00"))
   r_dbl_SalCap = r_dbl_ImpPre

   'Calculando Cupón Mensual Redondeado
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      If (Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 12) And r_int_CuoExt = 2 Then
         r_dbl_CupMen = r_dbl_CuoIni * 2
      Else
         r_dbl_CupMen = r_dbl_CuoIni
      End If
      
      r_dbl_Princi = r_dbl_SalCap
      
      r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
      
      r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_CupMen = r_dbl_CupMen
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      p_Arregl(r_int_Contad).CuoCli_IntCap = 0
   Next r_int_Contad

   'Recalculando Ultima Cuota ya Redondeada
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni + r_dbl_SalCap, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(p_NumCuo).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
   
   r_dbl_Capita = r_dbl_CuoIni - r_dbl_Intere
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_CupMen = r_dbl_CupMen
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
   p_Arregl(p_NumCuo).CuoCli_IntCap = 0
   
   'Para Ajustar Primera Cuota cuando Fecha de Desembolso no coincide con Ultimo Día Habil del Mes
   'Primera Cuota debe ser Mayor que las demás
   If p_PerGra = 0 And p_Tramo = 1 And r_int_DiaAdc > 0 Then
      r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ ((p_Arregl(1).CuoCli_DifDia + r_int_DiaAdc) / 30) * p_MtoPre - p_MtoPre, "######0.00"))
      
      p_Arregl(1).CuoCli_Intere = r_dbl_Intere
      p_Arregl(1).CuoCli_ValCuo = p_Arregl(1).CuoCli_Capita + r_dbl_Intere
   End If
End Sub

Public Sub gs_Calcul_Cliente(p_Arregl() As modcal_g_est_CuoCli, p_TasInt As Double, p_FlgCuo As Integer, p_FecDes As String, p_NumCuo As Integer, p_MtoPre As Double, p_Tramo As Integer, p_PerGra As Integer, p_UltVct As String)
   'p_Arregl   -  Estructura para generar cronograma
   'p_TasInt   -  Tasa de Interes Anual
   'p_FlgCuo   -  Flag de Cuota Extraordinaria (Multiplicar x 1 ó 2 en los meses de Julio y Diciembre)
   'p_FecDes   -  Fecha de Desembolso
   'p_NumCuo   -  Número de Cuotas
   'p_MtoPre   -  Monto Préstamo
   'p_Tramo    -  1 - Tramo No Concesional (Mensual) / 2 - Tramo Concesional (Semestral)
   'p_PerGra   -  Período de Gracia
   'p_UltVct   -  Fecha de Ultimo Vencimiento de Cronograma No-Concesional
   
   Dim r_int_Contad        As Integer
   Dim r_int_Cont01        As Integer
   Dim r_str_DiaPag        As String
   Dim r_int_MesIni        As Integer
   Dim r_int_AnoIni        As Integer
   Dim r_str_FecVct        As String
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   Dim r_str_VctIni        As String
   Dim r_str_FinMes        As String
   Dim r_dbl_FacMen        As Double
   Dim r_dbl_TasMen        As Double
   Dim r_int_CuoExt        As Integer
   Dim r_dbl_AcuFac        As Double
   Dim r_dbl_ImpPre        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_CupMen        As Double
   Dim r_dbl_Princi        As Double
   Dim r_dbl_AjuPre        As Double
   Dim r_dbl_CuoIni        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Capita        As Double
   Dim r_str_VctHab        As String
   Dim r_int_DiaAdc        As Integer
   
   'Obteniendo si habran Cuotas Extraordinarias
   r_int_CuoExt = p_FlgCuo
   
   'Calculando Tasa de Interés Mensual y Comisión Mensual
   r_dbl_TasMen = ff_IntMen(p_TasInt)
   
   'Para hallar diferencia de días entre la Fecha de Desembolso y el Ultimo Día Hábil del Mes de Desembolso
   'Sólo válido para Tramo No Concesional Sin Período de Gracia
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
   r_str_VctHab = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
   r_int_DifDia = CInt(CDate(r_str_VctHab) - CDate(p_FecDes))
   
   r_int_DiaAdc = 0
   If p_Tramo = 1 And p_PerGra = 0 And r_int_DifDia > 0 Then
      'Sólo válido para Tramo No Concesional y sin Período de Gracia
      'Si la diferencia entre la Fecha de Desembolso y Ultimo Día Habil del Mes de Desembolso es Mayor a Cero
      'Se guardan los días de diferencia y se asume como Fecha de Inicio de Cronograma
      'el Ultimo Día Hábil del Mes de Desembolso
      r_str_VctIni = r_str_VctHab
      r_int_DiaAdc = r_int_DifDia
   Else
      'Fecha de Inicio de Cronograma siempre será la Fecha de Desembolso
      r_str_VctIni = p_FecDes
   End If
   
   'Obteniendo Primera Fecha de Vencimiento
   r_str_DiaPag = Day(CDate(p_FecDes))
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   'Si Tramo es No Concesional o es Tramo Concesional y hay Período de Gracia
   'Se incrementa el Mes en 1
   If (p_Tramo = 1) Or (p_PerGra > 0 And p_Tramo = 2) Then
      r_int_MesIni = r_int_MesIni + 1
      
      If r_int_MesIni = 13 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = 1
      End If
   Else
      'Si el Tramo es Concesional y no hay Período de Gracia se incrementa el mes en 6
      r_int_MesIni = r_int_MesIni + 6
         
      If r_int_MesIni > 12 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = r_int_MesIni - 12
      End If
   End If
   
   'Calculando Factor de Ajuste de Saldo Capital y Nuevo Saldo Capital
   r_dbl_AjuPre = (1 + ((1 + p_TasInt / 100) - 1)) ^ (1 / 12) - 1
   r_dbl_ImpPre = p_MtoPre
   r_dbl_SalCap = r_dbl_ImpPre
   
   'Inicializando Arreglo
   If p_Tramo = 2 And p_PerGra > 0 Then
      Select Case p_PerGra
         Case 1 To 5:   p_NumCuo = p_NumCuo + 1
         Case 6 To 11:  p_NumCuo = p_NumCuo
         Case Is > 11:  p_NumCuo = p_NumCuo - 1
      End Select
   End If
   ReDim p_Arregl(p_NumCuo)
   
   'Si Período de Gracia es Mayor a Cero
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      
      If p_Tramo = 1 Then        'Tramo No Concesional
         For r_int_Contad = 1 To p_PerGra
            r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
            r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
            
            r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
            r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
            r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
            
            r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
            
            p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
            p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
            p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
            p_Arregl(r_int_Contad).CuoCli_Capita = 0
            p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
            p_Arregl(r_int_Contad).CuoCli_CupMen = r_dbl_CupMen
            p_Arregl(r_int_Contad).CuoCli_ValCuo = 0
            p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
            p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
            
            r_str_VctIni = r_str_FecVct
            
            r_int_MesIni = r_int_MesIni + 1
            If r_int_MesIni = 13 Then
               r_int_AnoIni = r_int_AnoIni + 1
               r_int_MesIni = 1
            End If
         Next r_int_Contad
         
         r_dbl_ImpPre = CDbl(Format(r_dbl_SalCap, "########0.00"))
      Else                    'Tramo Concesional
         For r_int_Contad = 1 To p_PerGra
            r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
            r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
            
            r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
            r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            r_str_VctIni = r_str_FecVct
            
            r_int_MesIni = r_int_MesIni + 1
            If r_int_MesIni = 13 Then
               r_int_AnoIni = r_int_AnoIni + 1
               r_int_MesIni = 1
            End If
         Next r_int_Contad
         
         r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_AcuDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap + r_dbl_Intere, "########0.00"))
         r_dbl_ImpPre = r_dbl_SalCap
         
         p_Arregl(1).CuoCli_FecVct = r_str_FecVct
         p_Arregl(1).CuoCli_DifDia = r_int_DifDia
         p_Arregl(1).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(1).CuoCli_Capita = 0
         p_Arregl(1).CuoCli_Intere = r_dbl_Intere
         p_Arregl(1).CuoCli_CupMen = 0
         p_Arregl(1).CuoCli_ValCuo = 0
         p_Arregl(1).CuoCli_SalCap = r_dbl_SalCap
         p_Arregl(1).CuoCli_IntCap = r_dbl_Intere
         
         
         r_int_MesIni = r_int_MesIni + 5
         If r_int_MesIni > 12 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = r_int_MesIni - 12
         End If
         
         p_PerGra = 1
      End If
   End If

   'Calculando Fechas de Vencimiento y Factores Mensuales
   r_int_AcuDia = 0
   r_dbl_AcuFac = 0
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
      r_str_FecVct = ff_DiaHabil(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000"))
      
      'Si es Tramo Concesional validar que la Fecha de Vencimiento no puede superar el período total del préstamo
      If p_Tramo = 2 Then
         If CDate(r_str_FecVct) > CDate(p_UltVct) Then
            r_str_FecVct = p_UltVct
         End If
      End If
      
      r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      'Calculando Factor Mensual
      If (r_int_MesIni = 7 Or r_int_MesIni = 12) And r_int_CuoExt = 1 And p_Tramo = 1 Then
         r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30) * 2
      Else
         r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30)
      End If
      
      'Acumulando Factor
      r_dbl_AcuFac = r_dbl_AcuFac + CDbl(Format(r_dbl_FacMen, "###0.000000000"))
      
      'Pasando datos a Arreglo
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
      
      r_str_VctIni = r_str_FecVct
      If p_Tramo = 1 Then     'Tramo No Concesional (Mensual)
         r_int_MesIni = r_int_MesIni + 1
         If r_int_MesIni = 13 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = 1
         End If
      Else                    'Tramo Concesional (Semestral)
         r_int_MesIni = r_int_MesIni + 6
         If r_int_MesIni > 12 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = r_int_MesIni - 12
         End If
      End If
   Next r_int_Contad
   
   'Calculando Valor Cuota Inicial
   r_dbl_CuoIni = CDbl(Format(r_dbl_ImpPre, "#######0.00")) / r_dbl_AcuFac
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni, "#####0.00"))
   
   r_dbl_SalCap = CDbl(Format(r_dbl_ImpPre, "########0.00"))
   
   'Proceso de Iteración para Calcular Cupón Mensual
   For r_int_Cont01 = 1 To 1500
      'Calculando Cupón Mensual
      For r_int_Contad = p_PerGra + 1 To p_NumCuo
         If (Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 12) And r_int_CuoExt = 1 And p_Tramo = 1 Then
            r_dbl_CupMen = r_dbl_CuoIni * 2
         Else
            r_dbl_CupMen = r_dbl_CuoIni
         End If
         
         r_dbl_Princi = r_dbl_SalCap
         
         r_dbl_Intere = (1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
         r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_CupMen = r_dbl_CupMen
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         p_Arregl(r_int_Contad).CuoCli_IntCap = 0
      Next r_int_Contad
      
      'Ajustes de Saldo Capital y Valor de Cuota Inicial
      If r_dbl_SalCap > 0# Then
         r_dbl_CuoIni = r_dbl_CuoIni + 0.01
         
         'Select Case r_dbl_SalCap
         '   Case 0# To 1:     r_dbl_CuoIni = r_dbl_CuoIni + 0.001
         '   Case 1 To 2:      r_dbl_CuoIni = r_dbl_CuoIni + 0.01
         '   Case 2 To 5:      r_dbl_CuoIni = r_dbl_CuoIni + 0.1
         '   Case Is > 5:      r_dbl_CuoIni = r_dbl_CuoIni + 1
         'End Select

         r_dbl_SalCap = r_dbl_ImpPre
      ElseIf r_dbl_SalCap < 0 Then
         r_dbl_CuoIni = r_dbl_CuoIni - 0.01
         
         'Select Case r_dbl_SalCap
         '   Case Is < -5:     r_dbl_CuoIni = r_dbl_CuoIni - 1
         '   Case -5 To -2:    r_dbl_CuoIni = r_dbl_CuoIni - 0.1
         '   Case -2 To -1:    r_dbl_CuoIni = r_dbl_CuoIni - 0.01
         '   Case -1 To -0#:   r_dbl_CuoIni = r_dbl_CuoIni - 0.001
         'End Select
         r_dbl_SalCap = r_dbl_ImpPre
      ElseIf r_dbl_SalCap = 0# Then
         Exit For
      End If
   Next r_int_Cont01
   
   'Redondeando Cuota
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni, "######0.00"))
   r_dbl_SalCap = r_dbl_ImpPre

   'Calculando Cupón Mensual Redondeado
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      If (Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 12) And r_int_CuoExt = 1 Then
         r_dbl_CupMen = r_dbl_CuoIni * 2
      Else
         r_dbl_CupMen = r_dbl_CuoIni
      End If
      
      r_dbl_Princi = r_dbl_SalCap
      
      r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
      
      r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_CupMen = r_dbl_CupMen
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      p_Arregl(r_int_Contad).CuoCli_IntCap = 0
   Next r_int_Contad

   'Recalculando Ultima Cuota ya Redondeada
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni + r_dbl_SalCap, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(p_NumCuo).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
   
   r_dbl_Capita = r_dbl_CuoIni - r_dbl_Intere
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_CupMen = r_dbl_CupMen
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
   p_Arregl(p_NumCuo).CuoCli_IntCap = 0
   
   'Para Ajustar Primera Cuota cuando Fecha de Desembolso no coincide con Ultimo Día Habil del Mes
   'Primera Cuota debe ser Mayor que las demás
   If p_PerGra = 0 And p_Tramo = 1 And r_int_DiaAdc > 0 Then
      r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ ((p_Arregl(1).CuoCli_DifDia + r_int_DiaAdc) / 30) * p_MtoPre - p_MtoPre, "######0.00"))
      
      p_Arregl(1).CuoCli_Intere = r_dbl_Intere
      p_Arregl(1).CuoCli_ValCuo = p_Arregl(1).CuoCli_Capita + r_dbl_Intere
   End If
End Sub

Public Function gf_Interes(ByVal p_TasMen As Double, ByVal p_NumDia As Integer, ByVal p_BasImp As Double) As Double
   gf_Interes = (1 + p_TasMen) ^ (p_NumDia / 30) * p_BasImp - p_BasImp
End Function

Public Sub gs_Calcul_Cliente_VctFijo_MVNCo(p_Arregl() As modcal_g_est_CuoCli, p_TasInt As Double, p_TasDes As Double, p_TasViv As Double, p_Portes As Double, p_FecDes As String, p_NumCuo As Integer, p_MtoPre As Double, p_MtoViv As Double, p_PerGra As Integer)
   'No Concesional
   'p_Arregl   -  Estructura para generar cronograma
   'p_TasInt   -  Tasa de Interes Anual
   'p_FlgCuo   -  Flag de Cuota Extraordinaria (Multiplicar x 1 ó 2 en los meses de Julio y Diciembre)
   'p_FecDes   -  Fecha de Desembolso
   'p_NumCuo   -  Número de Cuotas
   'p_MtoPre   -  Monto Préstamo
   'p_Tramo    -  1 - Tramo No Concesional (Mensual) / 2 - Tramo Concesional (Semestral)
   'p_PerGra   -  Período de Gracia
   'p_UltVct   -  Fecha de Ultimo Vencimiento de Cronograma No-Concesional
   
   Dim r_int_Contad        As Integer
   Dim r_int_Cont01        As Integer
   Dim r_str_DiaPag        As String
   Dim r_int_MesIni        As Integer
   Dim r_int_AnoIni        As Integer
   Dim r_str_FecVct        As String
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   Dim r_str_VctIni        As String
   Dim r_str_FinMes        As String
   Dim r_dbl_FacMen        As Double
   Dim r_dbl_TasMen        As Double
   Dim r_dbl_AcuFac        As Double
   Dim r_dbl_ImpPre        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_CupMen        As Double
   Dim r_dbl_Princi        As Double
   Dim r_dbl_AjuPre        As Double
   Dim r_dbl_CuoIni        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Capita        As Double
   Dim r_str_VctHab        As String
   Dim r_int_DiaAdc        As Integer
   
   Dim r_dbl_DesMen        As Double
   Dim r_dbl_DesDia        As Double
   Dim r_dbl_IntDia        As Double
   
   Dim r_dbl_SegViv        As Double
   Dim r_dbl_SegPre        As Double
   
   
   'Calculando Tasa de Interés Mensual
   'r_dbl_TasMen = ff_IntMen(p_TasInt)
   
   r_dbl_IntDia = ((1 + p_TasInt / 100) ^ (1 / 360)) - 1
   
   r_dbl_DesMen = ((1 + p_TasDes / 100) ^ (1 / 12)) - 1
   r_dbl_DesDia = ((1 + p_TasDes / 100) ^ (1 / 360)) - 1
   
   
   'r_dbl_SegViv = CDbl(Format(p_TasViv / 100 * 0.72 * p_MtoViv, "###,###,##0.00"))
   r_dbl_SegViv = CDbl(Format(p_TasViv / 100 * p_MtoViv, "###,###,##0.00"))
   
   'Obteniendo Primera Fecha de Vencimiento
   r_str_VctIni = p_FecDes
   
   r_str_DiaPag = Format(Day(CDate(p_FecDes)), "00")
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   r_int_MesIni = r_int_MesIni + 1
   
   If r_int_MesIni = 13 Then
      r_int_AnoIni = r_int_AnoIni + 1
      r_int_MesIni = 1
   End If
   
   r_dbl_ImpPre = p_MtoPre
   r_dbl_SalCap = r_dbl_ImpPre
   
   ReDim p_Arregl(p_NumCuo)
   
   'Si Período de Gracia es Mayor a Cero
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_str_FecVct = r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")
         
         r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
         
         r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (r_int_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_ValCuo = 0
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         
         r_str_VctIni = r_str_FecVct
         
         r_int_MesIni = r_int_MesIni + 1
         If r_int_MesIni = 13 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = 1
         End If
      Next r_int_Contad
      
      r_dbl_ImpPre = CDbl(Format(r_dbl_SalCap, "########0.00"))
   End If

   'Calculando Fechas de Vencimiento y Factores Mensuales
   r_int_AcuDia = 0
   r_dbl_AcuFac = 0
   
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_str_FecVct = r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")
      
      r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      'Calculando Factor Mensual
      'r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30)
      
      r_dbl_FacMen = (1 + (r_dbl_IntDia + r_dbl_DesDia)) ^ (-1 * r_int_AcuDia)
      
      'Acumulando Factor
      r_dbl_AcuFac = r_dbl_AcuFac + CDbl(Format(r_dbl_FacMen, "###0.000000"))
      
      'Pasando datos a Arreglo
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
      
      r_str_VctIni = r_str_FecVct
      
      r_int_MesIni = r_int_MesIni + 1
      If r_int_MesIni = 13 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = 1
      End If
   Next r_int_Contad
   
   'Calculando Cuotas
   r_dbl_CuoIni = CDbl(Format((r_dbl_ImpPre / r_dbl_AcuFac) + r_dbl_SegViv + p_Portes, "###,###,##0.00"))
   r_dbl_SalCap = r_dbl_ImpPre
   
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_dbl_CupMen = r_dbl_CuoIni
      r_dbl_Princi = r_dbl_SalCap
      
      r_dbl_Intere = CDbl(Format(r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap, "###,##0.00"))
      r_dbl_SegPre = CDbl(Format(r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap, "###,##0.00"))
      
      'r_dbl_Intere = (1 + r_dbl_TasMen) ^ (p_Arregl(r_int_Contad).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
      'r_dbl_SegPre = (1 + (p_TasDes / 100)) ^ (p_Arregl(r_int_Contad).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
   
      r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere - r_dbl_SegPre - r_dbl_SegViv - p_Portes
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_SegPre
      p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_SegViv
      p_Arregl(r_int_Contad).CuoCli_Portes = p_Portes
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_SegPre + r_dbl_SegViv + p_Portes
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      p_Arregl(r_int_Contad).CuoCli_IntCap = 0
   Next r_int_Contad

   'Recalculando Ultima Cuota ya Redondeada
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni + r_dbl_SalCap, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   'r_dbl_Intere = CDbl(Format((1 + r_dbl_TasMen) ^ (p_Arregl(p_NumCuo).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap, "######0.00"))
   'r_dbl_SegPre = (1 + (p_TasDes / 100)) ^ (p_Arregl(p_NumCuo).CuoCli_DifDia / 30) * r_dbl_SalCap - r_dbl_SalCap
   
   r_dbl_Intere = CDbl(Format(r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap, "###,###,##0.00"))
   r_dbl_SegPre = CDbl(Format(r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap, "###,###,##0.00"))
   
   r_dbl_Capita = r_dbl_CuoIni - r_dbl_Intere - r_dbl_SegPre - r_dbl_SegViv - p_Portes
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SegPre = r_dbl_SegPre
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_SegPre + r_dbl_SegViv + p_Portes
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
   p_Arregl(p_NumCuo).CuoCli_IntCap = 0
End Sub

Public Sub gs_Calcul_Cliente_VctFijo_MVCon(p_Arregl() As modcal_g_est_CuoCli, p_TasInt As Double, p_FecDes As String, p_NumCuo As Integer, p_MtoPre As Double, p_PerGra As Integer, p_UltVct As String)
   'No Concesional
   'p_Arregl   -  Estructura para generar cronograma
   'p_TasInt   -  Tasa de Interes Anual
   'p_FlgCuo   -  Flag de Cuota Extraordinaria (Multiplicar x 1 ó 2 en los meses de Julio y Diciembre)
   'p_FecDes   -  Fecha de Desembolso
   'p_NumCuo   -  Número de Cuotas
   'p_MtoPre   -  Monto Préstamo
   'p_Tramo    -  1 - Tramo No Concesional (Mensual) / 2 - Tramo Concesional (Semestral)
   'p_PerGra   -  Período de Gracia
   'p_UltVct   -  Fecha de Ultimo Vencimiento de Cronograma No-Concesional
   
   Dim r_int_Contad        As Integer
   Dim r_int_Cont01        As Integer
   Dim r_str_DiaPag        As String
   Dim r_int_MesIni        As Integer
   Dim r_int_AnoIni        As Integer
   Dim r_str_FecVct        As String
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   Dim r_str_VctIni        As String
   Dim r_str_FinMes        As String
   Dim r_dbl_FacMen        As Double
   Dim r_dbl_TasMen        As Double
   Dim r_dbl_AcuFac        As Double
   Dim r_dbl_ImpPre        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_CupMen        As Double
   Dim r_dbl_Princi        As Double
   Dim r_dbl_AjuPre        As Double
   Dim r_dbl_CuoIni        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Capita        As Double
   Dim r_str_VctHab        As String
   Dim r_int_DiaAdc        As Integer
   
   Dim r_dbl_IntDia        As Double
   
   
   'Calculando Interés Diario
   r_dbl_IntDia = ((1 + p_TasInt / 100) ^ (1 / 360)) - 1
   
   
   'Obteniendo Primera Fecha de Vencimiento
   r_str_VctIni = p_FecDes
   
   r_str_DiaPag = Format(Day(CDate(p_FecDes)), "00")
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   If p_PerGra > 0 Then
      r_int_MesIni = r_int_MesIni + 1
      
      If r_int_MesIni = 13 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = 1
      End If
   Else
      r_int_MesIni = r_int_MesIni + 6
         
      If r_int_MesIni > 12 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = r_int_MesIni - 12
      End If
   End If
   
   
   r_dbl_ImpPre = p_MtoPre
   r_dbl_SalCap = r_dbl_ImpPre
   
   If p_PerGra > 0 Then
      Select Case p_PerGra
         Case 1 To 5:   p_NumCuo = p_NumCuo + 1
         Case 6 To 11:  p_NumCuo = p_NumCuo
         Case Is > 11:  p_NumCuo = p_NumCuo - 1
      End Select
   End If
   
   ReDim p_Arregl(p_NumCuo)
   
   'Si Período de Gracia es Mayor a Cero
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_str_FecVct = r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")
         
         r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
         
         r_dbl_Intere = CDbl(Format(r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap, "###,##0.00"))
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_ValCuo = 0
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         
         r_str_VctIni = r_str_FecVct
         
         r_int_MesIni = r_int_MesIni + 1
         If r_int_MesIni = 13 Then
            r_int_AnoIni = r_int_AnoIni + 1
            r_int_MesIni = 1
         End If
      Next r_int_Contad
      
      r_dbl_ImpPre = CDbl(Format(r_dbl_SalCap, "########0.00"))
   End If

   'Calculando Fechas de Vencimiento y Factores Mensuales
   r_int_AcuDia = 0
   r_dbl_AcuFac = 0
   
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_str_FecVct = r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")
      
      If CDate(r_str_FecVct) > CDate(p_UltVct) Then
         r_str_FecVct = p_UltVct
      End If
      
      r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      'Calculando Factor Mensual
      r_dbl_FacMen = (1 + r_dbl_IntDia) ^ (-1 * r_int_AcuDia)
      
      'Acumulando Factor
      r_dbl_AcuFac = r_dbl_AcuFac + CDbl(Format(r_dbl_FacMen, "###0.000000"))
      
      'Pasando datos a Arreglo
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecVct
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
      
      r_str_VctIni = r_str_FecVct
      
      r_int_MesIni = r_int_MesIni + 6
      If r_int_MesIni > 12 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = r_int_MesIni - 12
      End If
   Next r_int_Contad
   
   'Calculando Cuotas
   r_dbl_CuoIni = CDbl(Format((r_dbl_ImpPre / r_dbl_AcuFac), "###,###,##0.00"))
   r_dbl_SalCap = r_dbl_ImpPre
   
   For r_int_Contad = p_PerGra + 1 To p_NumCuo
      r_dbl_CupMen = r_dbl_CuoIni
      r_dbl_Princi = r_dbl_SalCap
      
      r_dbl_Intere = CDbl(Format(r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap, "###,##0.00"))
   
      r_dbl_Capita = r_dbl_CupMen - r_dbl_Intere
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      p_Arregl(r_int_Contad).CuoCli_IntCap = 0
   Next r_int_Contad

   'Recalculando Ultima Cuota ya Redondeada
   r_dbl_CuoIni = CDbl(Format(r_dbl_CuoIni + r_dbl_SalCap, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
   
   r_dbl_Capita = r_dbl_CuoIni - r_dbl_Intere
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_Capita + r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
   p_Arregl(p_NumCuo).CuoCli_IntCap = 0
End Sub

Public Sub gs_Calcul_Cliente_MiCasita(p_Arregl() As modcal_g_est_CuoCli, ByVal p_MtoViv As Double, ByVal p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_CuoExt As Integer, ByVal p_TasInt As Double, ByVal p_TasSgD As Double, ByVal p_TipSGV As Integer, ByVal p_TasSgV As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaVct As Integer, ByVal p_PerGra As Integer, Optional ByRef p_NuePre As Double)
   Dim r_str_DiaVct     As String
   Dim r_str_FecIni     As String
   Dim r_str_FecSgt     As String
   Dim r_int_PosMes     As Integer
   Dim r_int_PosAno     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_DifDia     As Integer
   Dim r_int_AcuDia     As Integer
   
   Dim r_dbl_IntDia     As Double
   Dim r_dbl_DesDia     As Double
   Dim r_dbl_Factor     As Double
   Dim r_dbl_FacMen     As Double
   Dim r_dbl_FacAcu     As Double
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_CuoMen     As Double
   Dim r_dbl_ValCuo     As Double
   Dim r_dbl_SalCap     As Double
   Dim r_int_MesDes     As Integer
   Dim r_int_AnoDes     As Integer
   Dim r_str_FecAju     As String
   Dim r_str_PriVct     As String

   ReDim p_Arregl(p_NumCuo)

   p_NuePre = p_MtoPre
   
   'Calculando Factor a Aplicar en Seguro de Vivienda
   If p_TipSGV = 1 Then
      r_dbl_SegViv = p_TasSgV / 100 * p_MtoViv
      r_dbl_SegViv = CDbl(Format(r_dbl_SegViv, "######0.00"))
   Else
      r_dbl_SegViv = p_TasSgV
   End If
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_DesDia = (1 + (p_TasSgD / 100)) ^ (1 / 30) - 1        'Calculando Tasa Diaria de Interes por Seguro Desgravamen

   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   'Obteniendo Mes y Año de Desembolso
   r_int_MesDes = Month(CDate(p_FecDes))
   r_int_AnoDes = Year(CDate(p_FecDes))
   
   r_int_MesDes = r_int_MesDes + 1
   If r_int_MesDes = 13 Then
      r_int_MesDes = 1
      r_int_AnoDes = r_int_AnoDes + 1
   End If
   
   'Obteniendo Fecha 1er Vencimiento
   r_str_PriVct = Format(p_DiaVct, "00") & "/" & Format(r_int_MesDes, "00") & "/" & Format(r_int_AnoDes, "0000")
   
   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_MesDes = r_int_MesDes + 1
      If r_int_MesDes = 13 Then
         r_int_MesDes = 1
         r_int_AnoDes = r_int_AnoDes + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaVct), "00") & "/" & Format(r_int_MesDes, "00") & "/" & Format(r_int_AnoDes, "0000")
   End If
   
   'Ajustando Fecha de Desembolso
   r_int_MesDes = Month(CDate(r_str_PriVct))
   r_int_AnoDes = Year(CDate(r_str_PriVct))
   
   r_int_MesDes = r_int_MesDes - 1
   If r_int_MesDes = 0 Then
      r_int_MesDes = 12
      r_int_AnoDes = r_int_AnoDes - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaVct), "00") & "/" & Format(r_int_MesDes, "00") & "/" & Format(r_int_AnoDes, "0000")
   
   If p_PerGra > 0 Then
      r_int_MesDes = Month(CDate(r_str_FecAju))
      r_int_AnoDes = Year(CDate(r_str_FecAju))
      
      r_int_MesDes = r_int_MesDes + p_PerGra
      If r_int_MesDes > 13 Then
         r_int_MesDes = r_int_MesDes - 12
         r_int_AnoDes = r_int_AnoDes + 1
      End If
      
      r_str_FecAju = Format(CInt(p_DiaVct), "00") & "/" & Format(r_int_MesDes, "00") & "/" & Format(r_int_AnoDes, "0000")
   End If
      
   'Capitalizando desde Fecha de Desembsolso Real hasta Fecha de Ajuste
   r_int_DifDia = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))
   p_MtoPre = (p_MtoPre * (1 + r_dbl_Factor) ^ r_int_DifDia)
   
   If p_PerGra > 0 Then
      p_NuePre = p_MtoPre
   End If
   
   r_str_FecIni = r_str_FecAju

   r_int_PosMes = Month(CDate(r_str_FecIni))
   r_int_PosAno = Year(CDate(r_str_FecIni))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_FecSgt = Format(p_DiaVct, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   
   'Obteniendo Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 To p_NumCuo
      r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      If (r_int_PosMes = 7 Or r_int_PosMes = 12) And p_CuoExt = 1 Then
         r_dbl_FacMen = (1 + r_dbl_Factor) ^ (-1 * r_int_AcuDia) * 2
      Else
         r_dbl_FacMen = (1 + r_dbl_Factor) ^ (-1 * r_int_AcuDia)
      End If
      
      
      r_dbl_FacAcu = r_dbl_FacAcu + r_dbl_FacMen
   
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
   
      'Calculando Siguiente Vencimiento
      r_str_FecIni = r_str_FecSgt
      
      r_int_PosMes = r_int_PosMes + 1
      
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecSgt = Format(p_DiaVct, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   Next r_int_Contad
   
   r_dbl_CuoMen = p_MtoPre / r_dbl_FacAcu
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   r_dbl_CuoMen = r_dbl_CuoMen
   
   r_dbl_SalCap = p_MtoPre
   
   'Generando Cronograma
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_SegDes = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))     'Redondeando
      r_dbl_SegDes = CDbl(Format(r_dbl_SegDes, "######0.00"))     'Redondeando
      
      If (Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 7 Or Month(CDate(p_Arregl(r_int_Contad).CuoCli_FecVct)) = 12) And p_CuoExt = 1 Then
         r_dbl_ValCuo = r_dbl_CuoMen * 2
      Else
         r_dbl_ValCuo = r_dbl_CuoMen
      End If
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_SegDes
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_SegDes
      p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_SegViv
      p_Arregl(r_int_Contad).CuoCli_Portes = p_Portes
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_SegViv + p_Portes
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_SegDes = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))     'Redondeando
   r_dbl_SegDes = CDbl(Format(r_dbl_SegDes, "######0.00"))     'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_SegDes
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SegPre = r_dbl_SegDes
   p_Arregl(p_NumCuo).CuoCli_SegViv = r_dbl_SegViv
   p_Arregl(p_NumCuo).CuoCli_Portes = p_Portes
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_SegViv + p_Portes
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
End Sub

Public Function modcal_gf_Calcul_MtoMax(p_ValCuo As Double, p_TasInt As Double, p_FecDes As String, p_NumCuo As Integer, ByVal p_ValViv As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_Portes As Double, ByVal p_CuoApr As Double, ByVal p_intere As Double, Optional ByRef p_CuoFin As Double) As Double
   Dim r_int_Contad        As Integer
   Dim r_str_DiaPag        As String
   Dim r_str_DiaFij        As String
   Dim r_int_MesIni        As Integer
   Dim r_int_AnoIni        As Integer
   Dim r_str_FecVct        As String
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   Dim r_str_VctIni        As String
   Dim r_dbl_FacMen        As Double
   Dim r_dbl_AcuFac        As Double
   Dim r_dbl_AjuPre        As Double
   Dim r_dbl_MtoMax        As Double
   Dim r_dbl_CuoMen        As Double
   Dim r_arr_CliNco()      As modcal_g_est_CuoCli

   
   p_CuoFin = 0
   modcal_gf_Calcul_MtoMax = 0
   
   'Obteniendo Primera Fecha de Vencimiento
   r_str_DiaFij = Day(CDate(p_FecDes))
   
   r_str_DiaPag = r_str_DiaFij
   r_int_MesIni = Month(CDate(p_FecDes))
   r_int_AnoIni = Year(CDate(p_FecDes))
   
   r_str_VctIni = r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")
   
   r_int_MesIni = r_int_MesIni + 1
   If r_int_MesIni = 13 Then
      r_int_AnoIni = r_int_AnoIni + 1
      r_int_MesIni = 1
   End If
   
   'Calculando Factor de Ajuste de Saldo Capital y Nuevo Saldo Capital
   r_dbl_AjuPre = (1 + ((1 + p_TasInt / 100) - 1)) ^ (1 / 12) - 1
   
   'Calculando Fechas de Vencimiento y Factores Mensuales
   r_int_AcuDia = 0
   r_dbl_AcuFac = 0
   
   For r_int_Contad = 1 To p_NumCuo
      r_str_DiaPag = r_str_DiaFij
      If Not IsDate(r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")) Then
         r_str_DiaPag = Format(ff_Ultimo_Dia_Mes(r_int_MesIni, r_int_AnoIni), "00")
      End If
      r_str_FecVct = r_str_DiaPag & "/" & Format(r_int_MesIni, "00") & "/" & Format(r_int_AnoIni, "0000")
      
      r_int_DifDia = CInt(CDate(r_str_FecVct) - CDate(r_str_VctIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      'Calculando Factor Mensual
      r_dbl_FacMen = (1 + r_dbl_AjuPre) ^ ((-1 * r_int_AcuDia) / 30)
      
      'Acumulando Factor
      r_dbl_AcuFac = r_dbl_AcuFac + CDbl(Format(r_dbl_FacMen, "###0.000000000"))
      
      r_str_VctIni = r_str_FecVct
      r_int_MesIni = r_int_MesIni + 1
      
      If r_int_MesIni = 13 Then
         r_int_AnoIni = r_int_AnoIni + 1
         r_int_MesIni = 1
      End If
   Next r_int_Contad
   
   'Monto Máximo Préstamo
   r_dbl_MtoMax = p_ValCuo * r_dbl_AcuFac
   r_dbl_MtoMax = Format(r_dbl_MtoMax, "######0")
   r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
      
   Call gs_Calcul_Cliente_MiCasita(r_arr_CliNco(), p_ValViv, r_dbl_MtoMax, p_NumCuo, 2, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, 0)
   r_dbl_CuoMen = r_arr_CliNco(1).CuoCli_ValCuo
   
   Do While p_CuoApr < r_dbl_CuoMen
      r_dbl_MtoMax = r_dbl_MtoMax - 100
      
      Call gs_Calcul_Cliente_MiCasita(r_arr_CliNco(), p_ValViv, r_dbl_MtoMax, p_NumCuo, 2, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, 0)
      r_dbl_CuoMen = r_arr_CliNco(1).CuoCli_ValCuo
   Loop
   
   If r_dbl_MtoMax > 0 Then
      r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
      Call gs_Calcul_Cliente_MiCasita(r_arr_CliNco(), p_ValViv, r_dbl_MtoMax, p_NumCuo, 2, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, 0)
      p_CuoFin = r_arr_CliNco(1).CuoCli_ValCuo
   End If
   
   modcal_gf_Calcul_MtoMax = r_dbl_MtoMax
End Function

Public Sub gs_Cronog_MiCasita_PrePag(p_Arregl() As modcal_g_est_CuoCli, ByVal p_MtoViv As Double, ByVal p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_CuoExt As Integer, ByVal p_TasInt As Double, ByVal p_TasSgD As Double, ByVal p_TipSGV As Integer, ByVal p_TasSgV As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByVal p_PriVct As String)
   Dim r_dbl_TasInt_Dia    As Double
   Dim r_dbl_TasSgD_Dia    As Double
   Dim r_dbl_ImpSgV_Men    As Double
   Dim r_dbl_ImpSgD_Men    As Double
   
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_DesDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_CuoMen        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_IntCap        As Double
   
   Dim r_dbl_FacAcu        As Double
   Dim r_dbl_FacMen        As Double
   
   Dim r_dbl_ValCuo        As Double
   Dim r_dbl_Capita        As Double
   
   
   Dim r_int_Contad        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo)
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasSgV / 100 * p_MtoViv
   Else
      r_dbl_ImpSgV_Men = p_TasSgV
   End If
   r_dbl_ImpSgV_Men = CDbl(Format(r_dbl_ImpSgV_Men, "######0.00"))
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_DesDia = (1 + (p_TasSgD / 100)) ^ (1 / 30) - 1        'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))

   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   
   r_str_PriVct = p_PriVct

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = p_FecDes
   r_str_FecSgt = r_str_PriVct

   r_dbl_SalCap = p_MtoPre
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 To p_NumCuo
      r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      r_dbl_FacMen = (1 + r_dbl_Factor) ^ (-1 * r_int_AcuDia)
      r_dbl_FacAcu = r_dbl_FacAcu + r_dbl_FacMen
   
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
   
      'Calculando Siguiente Vencimiento
      r_str_FecIni = r_str_FecSgt
      
      r_int_PosMes = r_int_PosMes + 1
      
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_SalCap / r_dbl_FacAcu
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   
   
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_ImpSgD_Men
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
      p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
      p_Arregl(r_int_Contad).CuoCli_Portes = p_Portes
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_ImpSgV_Men + p_Portes
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
   r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_ImpSgD_Men
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SegPre = r_dbl_ImpSgD_Men
   p_Arregl(p_NumCuo).CuoCli_SegViv = r_dbl_ImpSgV_Men
   p_Arregl(p_NumCuo).CuoCli_Portes = p_Portes
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_ImpSgV_Men + p_Portes
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
End Sub

Public Sub gs_Cronog_CRCPBP_ConMVi_PrePag(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_TasInt As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByVal p_PriVct As String)
   Dim r_dbl_TasInt_Dia    As Double
   
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_CuoMen        As Double
   Dim r_dbl_SalCap        As Double
   
   Dim r_dbl_FacAcu        As Double
   Dim r_dbl_FacMen        As Double
   
   Dim r_dbl_ValCuo        As Double
   Dim r_dbl_Capita        As Double
   
   Dim r_int_Contad        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   'Select Case p_PerGra
   '   Case 1 To 5:   p_NumCuo = p_NumCuo + 1
   '   Case 6 To 11:  p_NumCuo = p_NumCuo
   '   Case Is > 11:  p_NumCuo = p_NumCuo - 1
   'End Select
   
   ReDim p_Arregl(p_NumCuo)
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_Factor = r_dbl_IntDia
   
   
   r_str_FecIni = p_FecDes
   r_str_FecSgt = p_PriVct
   
   r_int_PosMes = Month(CDate(p_PriVct))
   r_int_PosAno = Year(CDate(p_PriVct))
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 To p_NumCuo
      r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      r_dbl_FacMen = (1 + r_dbl_Factor) ^ (-1 * r_int_AcuDia)
      r_dbl_FacAcu = r_dbl_FacAcu + r_dbl_FacMen
   
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
   
      'Calculando Siguiente Vencimiento
      r_str_FecIni = r_str_FecSgt
      
      r_int_PosMes = r_int_PosMes + 6
      
      If r_int_PosMes > 12 Then
         r_int_PosMes = r_int_PosMes - 12
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   Next r_int_Contad
   
   r_dbl_CuoMen = p_MtoPre / r_dbl_FacAcu
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   r_dbl_SalCap = p_MtoPre
   
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
End Sub

Public Sub gs_Cronog_CRCPBP_NC_PrePag(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_PorCon As Double, ByVal p_ValInm As Double, ByVal p_NumCuo As Integer, ByVal p_TasInt As Double, ByVal p_TasDes As Double, ByVal p_TipSGV As Integer, ByVal p_TasViv As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_PriVct As String, ByVal p_DiaPag As Integer, ByRef p_MtoNCo As Double, ByRef p_MtoCon As Double)
   'p_Arregl   -  Estructura para generar cronograma
   'p_MtoPre   -  Monto del Préstamo
   'p_PorCon   -  Porcentaje de Tramo Concesional
   'p_TopCon   -  Tope de Tramo Concesional (Siempre en Soles)
   'p_TipCam   -  Tipo de Cambio
   'p_ValInm   -  Valor de Inmueble
   'p_NumCuo   -  Número de Cuotas
   'p_PerGra   -  Período de Gracia
   'p_TasInt   -  Tasa de Interés
   'p_SegDes   -  Tasa de Seguro de Desgravamen
   'p_TipSGV   -  Tipo de Seguro de Inmueble
   'p_SegViv   -  Tasa de Seguro de Inmueble
   'p_Portes   -  Portes
   'p_FecDes   -  Fecha de Desembolso
   'p_DiaPag   -  Día de Pago
   
   Dim r_dbl_TasInt_Dia    As Double
   Dim r_dbl_TasSgD_Dia    As Double
   Dim r_dbl_ImpSgV_Men    As Double
   Dim r_dbl_ImpSgD_Men    As Double
   
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_DesDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_MtoCon        As Double
   Dim r_dbl_MtoNCo        As Double
   Dim r_dbl_CuoMen        As Double
   Dim r_dbl_SalCap        As Double
   Dim r_dbl_IntCap        As Double
   
   Dim r_dbl_FacAcu        As Double
   Dim r_dbl_FacMen        As Double
   
   Dim r_dbl_ValCuo        As Double
   Dim r_dbl_Capita        As Double
   
   
   Dim r_int_Contad        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo)
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasViv / 100 * p_ValInm
   Else
      r_dbl_ImpSgV_Men = p_TasViv
   End If
   r_dbl_ImpSgV_Men = CDbl(Format(r_dbl_ImpSgV_Men, "######0.00"))
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_DesDia = (1 + (p_TasDes / 100)) ^ (1 / 30) - 1        'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))

   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = p_FecDes
   r_str_FecSgt = p_PriVct

   r_int_PosMes = Month(CDate(p_PriVct))
   r_int_PosAno = Year(CDate(p_PriVct))
   
   r_dbl_SalCap = p_MtoPre
   
   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = r_dbl_SalCap * p_PorCon / 100
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = r_dbl_SalCap - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   p_MtoNCo = r_dbl_MtoNCo
   p_MtoCon = r_dbl_MtoCon
   
   'Recalculando Tasa Diaria de Interes por Seguro Desgravamen
   r_dbl_DesDia = (1 + (p_TasDes / 100 / (r_dbl_MtoNCo / (r_dbl_MtoNCo + r_dbl_MtoCon)))) ^ (1 / 30) - 1
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))
   
   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 To p_NumCuo
      r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      r_dbl_FacMen = (1 + r_dbl_Factor) ^ (-1 * r_int_AcuDia)
      r_dbl_FacAcu = r_dbl_FacAcu + r_dbl_FacMen
   
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = r_dbl_FacMen
   
      'Calculando Siguiente Vencimiento
      r_str_FecIni = r_str_FecSgt
      
      r_int_PosMes = r_int_PosMes + 1
      
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))    'Redondeando
      
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_ImpSgD_Men
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
      p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
      p_Arregl(r_int_Contad).CuoCli_Portes = p_Portes
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_ImpSgV_Men + p_Portes
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ p_Arregl(p_NumCuo).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
   r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_ImpSgD_Men
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo).CuoCli_SegPre = r_dbl_ImpSgD_Men
   p_Arregl(p_NumCuo).CuoCli_SegViv = r_dbl_ImpSgV_Men
   p_Arregl(p_NumCuo).CuoCli_Portes = p_Portes
   p_Arregl(p_NumCuo).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_ImpSgV_Men + p_Portes
   p_Arregl(p_NumCuo).CuoCli_SalCap = r_dbl_SalCap
End Sub


