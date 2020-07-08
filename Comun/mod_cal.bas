Attribute VB_Name = "modcal"
Option Explicit

Public Type modcal_g_est_CuoCli
   CuoCli_FecVct     As String
   CuoCli_DifDia     As Integer
   CuoCli_AcuDia     As Integer
   CuoCli_Factor     As Double
   CuoCli_Capita     As Double
   CuoCli_Intere     As Double
   CuoCli_Comisi     As Double
   CuoCli_CupMen     As Double
   CuoCli_ValCuo     As Double
   CuoCli_SalCap     As Double
   CuoCli_IntCap     As Double
   CuoCli_SegPre     As Double
   CuoCli_SegViv     As Double
   CuoCli_Portes     As Double
End Type

Public Type modcal_g_est_CuoCof
   CuoCof_FecVct     As String
   CuoCof_DifDia     As Integer
   CuoCof_AcuDia     As Integer
   CuoCof_Factor     As Double
   CuoCof_Capita     As Double
   CuoCof_Intere     As Double
   CuoCof_Comisi     As Double
   CuoCof_IntTot     As Double
   CuoCof_CupMen     As Double
   CuoCof_ValCuo     As Double
   CuoCof_SalCap     As Double
   CuoCof_IntCap     As Double
End Type

Public Function modcal_gf_Calcul_MtoMax_CRCPBP(p_ValCuo As Double, p_TasInt As Double, p_FecDes As String, p_NumCuo As Integer, ByVal p_ValViv As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_Portes As Double, ByVal p_CuoApr As Double, ByVal p_intere As Double, ByVal p_PorCon As Double, ByVal p_TopCon As Double, ByVal p_TipCam As Double, Optional ByRef p_CuoFin As Double, Optional ByVal p_PerGra As Integer) As Double
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
   Dim r_dbl_MtoNCo        As Double
   Dim r_dbl_MtoCon        As Double
   Dim r_dbl_IntGra        As Double
   
   modcal_gf_Calcul_MtoMax_CRCPBP = 0
   
   
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
      
   Call gs_Cronog_CRCPBP_NC(r_arr_CliNco(), r_dbl_MtoMax, p_PorCon, p_TopCon, p_TipCam, p_ValViv, p_NumCuo, p_PerGra, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   
   'r_dbl_MtoMax = r_dbl_MtoMax - r_dbl_IntGra
   r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   
   Do While p_CuoApr < r_dbl_CuoMen
      r_dbl_MtoMax = r_dbl_MtoMax - 100
      
      Call gs_Cronog_CRCPBP_NC(r_arr_CliNco(), r_dbl_MtoMax, p_PorCon, p_TopCon, p_TipCam, p_ValViv, p_NumCuo, p_PerGra, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, r_dbl_MtoNCo, r_dbl_MtoCon, , 2)
      r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   Loop
   
   r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
   Call gs_Cronog_CRCPBP_NC(r_arr_CliNco(), r_dbl_MtoMax, p_PorCon, p_TopCon, p_TipCam, p_ValViv, p_NumCuo, 0, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, r_dbl_MtoNCo, r_dbl_MtoCon)
   p_CuoFin = r_arr_CliNco(2).CuoCli_ValCuo
   
   modcal_gf_Calcul_MtoMax_CRCPBP = r_dbl_MtoMax
End Function

Public Function modcal_gf_Calcul_MtoMax_miCasita(p_ValCuo As Double, p_TasInt As Double, p_FecDes As String, p_NumCuo As Integer, ByVal p_ValViv As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_Portes As Double, ByVal p_CuoApr As Double, ByVal p_intere As Double, Optional ByRef p_CuoFin As Double) As Double
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
   modcal_gf_Calcul_MtoMax_miCasita = 0
   
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
      
   Call gs_Cronog_MiCasita(r_arr_CliNco(), p_ValViv, r_dbl_MtoMax, p_NumCuo, 2, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, 0)
   r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   
   Do While p_CuoApr < r_dbl_CuoMen
      r_dbl_MtoMax = r_dbl_MtoMax - 100
      
      Call gs_Cronog_MiCasita(r_arr_CliNco(), p_ValViv, r_dbl_MtoMax, p_NumCuo, 2, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, 0)
      r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   Loop
   
   If r_dbl_MtoMax > 0 Then
      r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
      Call gs_Cronog_MiCasita(r_arr_CliNco(), p_ValViv, r_dbl_MtoMax, p_NumCuo, 2, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, 0)
      
      p_CuoFin = r_arr_CliNco(2).CuoCli_ValCuo
   End If
   
   modcal_gf_Calcul_MtoMax_miCasita = r_dbl_MtoMax
End Function

Public Function modcal_gf_Calcul_MtoMax_MiHogar(p_ValCuo As Double, p_TasInt As Double, p_FecDes As String, p_NumCuo As Integer, ByVal p_ValViv As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_Portes As Double, ByVal p_CuoApr As Double, ByVal p_intere As Double, ByVal p_TopCon As Double, Optional ByRef p_CuoFin As Double, Optional ByVal p_PerGra As Integer) As Double
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
   Dim r_dbl_MtoNCo        As Double
   Dim r_dbl_MtoCon        As Double
   Dim r_dbl_IntGra        As Double
   
   modcal_gf_Calcul_MtoMax_MiHogar = 0
   
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
   r_dbl_MtoMax = Format(r_dbl_MtoMax + p_TopCon, "######0")
   r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
      
   Call gs_Cronog_Mihogar_NC(r_arr_CliNco(), r_dbl_MtoMax, p_TopCon, p_ValViv, p_NumCuo, p_PerGra, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   
   'r_dbl_MtoMax = r_dbl_MtoMax - r_dbl_IntGra
   r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   
   Do While p_CuoApr < r_dbl_CuoMen
      r_dbl_MtoMax = r_dbl_MtoMax - 100
      
      Call gs_Cronog_Mihogar_NC(r_arr_CliNco(), r_dbl_MtoMax, p_TopCon, p_ValViv, p_NumCuo, p_PerGra, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, r_dbl_MtoNCo, r_dbl_MtoCon, , 2)
      
      DoEvents
      r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   Loop
   
   r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
   Call gs_Cronog_Mihogar_NC(r_arr_CliNco(), r_dbl_MtoMax, p_TopCon, p_ValViv, p_NumCuo, 0, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 5, r_dbl_MtoNCo, r_dbl_MtoCon)
   p_CuoFin = r_arr_CliNco(2).CuoCli_ValCuo
   
   modcal_gf_Calcul_MtoMax_MiHogar = r_dbl_MtoMax
End Function

Public Function modcal_gf_Calcul_MtoMax_CME(p_ValCuo As Double, p_TasInt As Double, p_FecDes As String, p_NumCuo As Integer, ByVal p_ValViv As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_Portes As Double, ByVal p_CuoApr As Double, ByVal p_intere As Double, ByVal p_PorCon As Double, ByVal p_TopCon As Double, Optional ByRef p_CuoFin As Double, Optional ByVal p_PerGra As Integer) As Double
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
   Dim r_dbl_MtoNCo        As Double
   Dim r_dbl_MtoCon        As Double
   Dim r_dbl_IntGra        As Double
   
   modcal_gf_Calcul_MtoMax_CME = 0
   
   
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
      
   Call gs_Cronog_CME_NC(r_arr_CliNco(), r_dbl_MtoMax, p_PorCon, p_TopCon, p_ValViv, p_NumCuo, p_PerGra, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 2, r_dbl_MtoNCo, r_dbl_MtoCon, r_dbl_IntGra, 2)
   
   'r_dbl_MtoMax = r_dbl_MtoMax - r_dbl_IntGra
   r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   
   Do While p_CuoApr < r_dbl_CuoMen
      r_dbl_MtoMax = r_dbl_MtoMax - 100
      
      Call gs_Cronog_CME_NC(r_arr_CliNco(), r_dbl_MtoMax, p_PorCon, p_TopCon, p_ValViv, p_NumCuo, p_PerGra, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 2, r_dbl_MtoNCo, r_dbl_MtoCon, , 2)
      r_dbl_CuoMen = r_arr_CliNco(2).CuoCli_ValCuo
   Loop
   
   r_dbl_MtoMax = CDbl(Mid(CStr(r_dbl_MtoMax), 1, Len(CStr(r_dbl_MtoMax)) - 2) & "00")
   Call gs_Cronog_CME_NC(r_arr_CliNco(), r_dbl_MtoMax, p_PorCon, p_TopCon, p_ValViv, p_NumCuo, 0, p_intere, p_SegDes, 1, p_SegViv, p_Portes, p_FecDes, 2, r_dbl_MtoNCo, r_dbl_MtoCon, , 2)
   p_CuoFin = r_arr_CliNco(2).CuoCli_ValCuo
   
   modcal_gf_Calcul_MtoMax_CME = r_dbl_MtoMax
End Function

Public Sub gs_Cronog_MiCasita(p_Arregl() As modcal_g_est_CuoCli, ByVal p_MtoViv As Double, ByVal p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_CuoExt As Integer, ByVal p_TasInt As Double, ByVal p_TasSgD As Double, ByVal p_TipSGV As Integer, ByVal p_TasSgV As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByVal p_PerGra As Integer, Optional ByRef p_IntCap As Double, Optional ByVal p_Cronog As Integer)
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
   p_IntCap = 0
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasSgV / 100 * p_MtoViv
      If p_Cronog <> 1 Then
         r_dbl_ImpSgV_Men = r_dbl_ImpSgV_Men * 0.72
      End If
   Else
      r_dbl_ImpSgV_Men = p_TasSgV
   End If
   r_dbl_ImpSgV_Men = CDbl(Format(r_dbl_ImpSgV_Men, "######0.00"))
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_DesDia = (1 + (p_TasSgD / 100)) ^ (1 / 30) - 1        'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.0000000000"))
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))
   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))   'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + (r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men)
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         'p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         'p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         'p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         'p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_Capita = 0
         'p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
         'p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_Portes = 0
         'p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
      
      p_IntCap = r_dbl_IntCap
   End If
   
   
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

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 Then
      r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      r_dbl_ImpSgD_Men = p_MtoPre * (1 + r_dbl_DesDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
      
      p_Arregl(1).CuoCli_Intere = p_Arregl(1).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1).CuoCli_SegPre = p_Arregl(1).CuoCli_SegPre + r_dbl_ImpSgD_Men
      p_Arregl(1).CuoCli_SegViv = p_Arregl(1).CuoCli_SegViv + r_dbl_ImpSgV_Men
      p_Arregl(1).CuoCli_ValCuo = p_Arregl(1).CuoCli_ValCuo + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
   End If
End Sub

Public Sub gs_Cronog_CME_NC(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_PorCon As Double, ByVal p_TopCon As Double, ByVal p_ValInm As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_TasDes As Double, ByVal p_TipSGV As Integer, ByVal p_TasViv As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByRef p_MtoNCo As Double, ByRef p_MtoCon As Double, Optional ByRef p_IntCap As Double, Optional ByVal p_Cronog As Integer)
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
   
   p_IntCap = 0
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasViv / 100 * p_ValInm
      
      If p_Cronog <> 1 Then
         r_dbl_ImpSgV_Men = r_dbl_ImpSgV_Men * 0.72
      End If
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
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))   'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + (r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men)
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         'p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         'p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         'p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         'p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_Capita = 0
         'p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
         'p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_Portes = 0
         'p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
      
      p_IntCap = r_dbl_IntCap
   End If
   
   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = r_dbl_SalCap * p_PorCon / 100
   
   If r_dbl_MtoCon > p_TopCon Then
      r_dbl_MtoCon = p_TopCon
   End If
   
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

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 Then
      r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      r_dbl_ImpSgD_Men = p_MtoPre * (1 + r_dbl_DesDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
      
      p_Arregl(1).CuoCli_Intere = p_Arregl(1).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1).CuoCli_SegPre = p_Arregl(1).CuoCli_SegPre + r_dbl_ImpSgD_Men
      p_Arregl(1).CuoCli_SegViv = p_Arregl(1).CuoCli_SegViv + r_dbl_ImpSgV_Men
      p_Arregl(1).CuoCli_ValCuo = p_Arregl(1).CuoCli_ValCuo + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
   End If
End Sub

Public Sub gs_Cronog_CME_ConMVi(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer)
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
   
   'Calculando Primer Vencimiento TNC
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento TNC
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   

   'Ajustando Fecha de Inicio de Pagos
   'r_str_FecIni = r_str_FecAju
   r_str_FecIni = p_FecDes
   
   If p_PerGra > 0 Then
      r_int_PosMes = Month(CDate(r_str_FecAju))
      r_int_PosAno = Year(CDate(r_str_FecAju))
      
      r_int_PosMes = r_int_PosMes + p_PerGra
      
      If r_int_PosMes > 12 Then
         r_int_PosMes = r_int_PosMes - 12
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecIni = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If
   
   'Obteniendo Primer Vencimiento TC
   r_int_PosMes = r_int_PosMes + 6
   If r_int_PosMes > 12 Then
      r_int_PosMes = r_int_PosMes - 12
      r_int_PosAno = r_int_PosAno + 1
   End If
   r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   
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

Public Sub gs_Cronog_CME_ConCli(p_Arregl() As modcal_g_est_CuoCli, p_ArrDes() As modcal_g_est_CuoCli, ByVal p_NumCuo As Integer, ByVal p_TasInt As Double)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_int_Contad        As Integer
   
   ReDim p_ArrDes(p_NumCuo)
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   
   'Generando Cronograma
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = (p_Arregl(r_int_Contad).CuoCli_SalCap + p_Arregl(r_int_Contad).CuoCli_Capita) * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - (p_Arregl(r_int_Contad).CuoCli_SalCap + p_Arregl(r_int_Contad).CuoCli_Capita)
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      p_ArrDes(r_int_Contad).CuoCli_FecVct = p_Arregl(r_int_Contad).CuoCli_FecVct
      p_ArrDes(r_int_Contad).CuoCli_DifDia = p_Arregl(r_int_Contad).CuoCli_DifDia
      p_ArrDes(r_int_Contad).CuoCli_AcuDia = p_Arregl(r_int_Contad).CuoCli_AcuDia
      p_ArrDes(r_int_Contad).CuoCli_Factor = p_Arregl(r_int_Contad).CuoCli_Factor
      p_ArrDes(r_int_Contad).CuoCli_Capita = p_Arregl(r_int_Contad).CuoCli_Capita
      p_ArrDes(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_ArrDes(r_int_Contad).CuoCli_ValCuo = p_Arregl(r_int_Contad).CuoCli_Capita + r_dbl_Intere
      p_ArrDes(r_int_Contad).CuoCli_SalCap = p_Arregl(r_int_Contad).CuoCli_SalCap
   Next r_int_Contad
End Sub

Public Sub gs_Cronog_CME_NCMVI(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_PorCon As Double, ByVal p_TopCon As Double, ByVal p_ValInm As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_TasDes As Double, ByVal p_TipSGV As Integer, ByVal p_TasViv As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByRef p_MtoNCo As Double, ByRef p_MtoCon As Double, Optional ByRef p_IntCap As Double, Optional ByVal p_Cronog As Integer)
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
   Dim r_dbl_SalAux        As Double
   
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


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   p_IntCap = 0
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasViv / 100 * p_ValInm
      
      If p_Cronog <> 1 Then
         r_dbl_ImpSgV_Men = r_dbl_ImpSgV_Men * 0.72
      End If
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
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))   'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + (r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men)
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
         p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
         p_Arregl(r_int_Contad).CuoCli_Portes = 0
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
      
      p_IntCap = r_dbl_IntCap
   End If
   
   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = r_dbl_SalCap * p_PorCon / 100
   
   If r_dbl_MtoCon > p_TopCon Then
      r_dbl_MtoCon = p_TopCon
   End If
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = r_dbl_SalCap - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   p_MtoNCo = r_dbl_MtoNCo
   p_MtoCon = r_dbl_MtoCon
   
   'Recalculando Tasa Diaria de Interes por Seguro Desgravamen
   r_dbl_DesDia = (1 + (p_TasDes / 100 / (r_dbl_MtoNCo / (r_dbl_MtoNCo + r_dbl_MtoCon)))) ^ (1 / 30) - 1
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))
   
   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   If p_PerGra > 0 Then
      r_dbl_SalAux = r_dbl_MtoNCo - p_IntCap
   
      For r_int_Contad = 1 To p_PerGra
         r_dbl_SalAux = r_dbl_SalAux + p_Arregl(r_int_Contad).CuoCli_ValCuo
   
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalAux
      Next r_int_Contad
   End If
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SegPre = r_dbl_ImpSgD_Men
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SegViv = r_dbl_ImpSgV_Men
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Portes = p_Portes
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_ImpSgV_Men + p_Portes
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 Then
      r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      r_dbl_ImpSgD_Men = p_MtoPre * (1 + r_dbl_DesDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
      
      p_Arregl(p_PerGra + 1).CuoCli_Intere = p_Arregl(p_PerGra + 1).CuoCli_Intere + r_dbl_Intere
      p_Arregl(p_PerGra + 1).CuoCli_SegPre = p_Arregl(p_PerGra + 1).CuoCli_SegPre + r_dbl_ImpSgD_Men
      p_Arregl(p_PerGra + 1).CuoCli_SegViv = p_Arregl(p_PerGra + 1).CuoCli_SegViv + r_dbl_ImpSgV_Men
      p_Arregl(p_PerGra + 1).CuoCli_ValCuo = p_Arregl(p_PerGra + 1).CuoCli_ValCuo + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
   End If
End Sub

Public Sub gs_Cronog_CME_NCCof(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   'r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.00000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

'   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
'      r_int_PosMes = r_int_PosMes + 1
'      If r_int_PosMes = 13 Then
'         r_int_PosMes = 1
'         r_int_PosAno = r_int_PosAno + 1
'      End If
      
      'Ajustando Fecha de 1er Vencimiento
'      r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
'      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   Else
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_FecAju = ff_DiaHabil(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1000
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         'r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         'r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "#####0.00"))
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Cronog_CME_NCCof_1(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.0000000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
   Else
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_FecAju = ff_DiaHabil(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1000
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         'r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         'r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "#####0.00"))
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Cronog_Mihogar_NC(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_ValInm As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_TasDes As Double, ByVal p_TipSGV As Integer, ByVal p_TasViv As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByRef p_MtoNCo As Double, ByRef p_MtoCon As Double, Optional ByRef p_IntCap As Double, Optional ByVal p_Cronog As Integer)
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

   Dim r_int_NumIte        As Integer

   ReDim p_Arregl(p_NumCuo)
   
   p_IntCap = 0
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasViv / 100 * p_ValInm
      
      If p_Cronog <> 1 Then
         r_dbl_ImpSgV_Men = r_dbl_ImpSgV_Men * 0.72
      End If
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
   
   'Calculando Primer Vencimiento
   'p_DiaPag = Day(CDate(p_FecDes))
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   'If p_DiaPag > 28 Then
   '   p_DiaPag = 28
   'End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   'r_str_FecIni = p_FecDes
   'r_str_FecSgt = r_str_PriVct

   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   p_MtoNCo = r_dbl_MtoNCo
   p_MtoCon = r_dbl_MtoCon
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   'Recalculando Tasa Diaria de Interes por Seguro Desgravamen
   r_dbl_DesDia = (1 + (p_TasDes / 100 / (r_dbl_MtoNCo / (r_dbl_MtoNCo + r_dbl_MtoCon)))) ^ (1 / 30) - 1
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))
   
   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))   'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + (r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men)
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         'p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         'p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         'p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         'p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_Capita = 0
         'p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
         'p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_Portes = 0
         'p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
      
      p_IntCap = r_dbl_IntCap
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   p_MtoNCo = r_dbl_SalCap
   
   'r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
   
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
   
   'Generando Cronograma
   For r_int_NumIte = 1 To 1000
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
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.0000"))   'Redondeando
      
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         r_dbl_CuoMen = r_dbl_CuoMen + 0.01
      ElseIf r_dbl_SalCap < 0 Then
         r_dbl_CuoMen = r_dbl_CuoMen - 0.01
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.0000"))        'Redondeando Valor Cuota
      
      DoEvents
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
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

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 Then
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_ImpSgD_Men = r_dbl_MtoNCo * (1 + r_dbl_DesDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
      
      p_Arregl(1).CuoCli_Intere = p_Arregl(1).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1).CuoCli_SegPre = p_Arregl(1).CuoCli_SegPre + r_dbl_ImpSgD_Men
      p_Arregl(1).CuoCli_SegViv = p_Arregl(1).CuoCli_SegViv + r_dbl_ImpSgV_Men
      p_Arregl(1).CuoCli_ValCuo = p_Arregl(1).CuoCli_ValCuo + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
   End If
End Sub

Public Sub gs_Cronog_Mihogar_NCCof(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.00000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

'   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
'      r_int_PosMes = r_int_PosMes + 1
'      If r_int_PosMes = 13 Then
'         r_int_PosMes = 1
'         r_int_PosAno = r_int_PosAno + 1
'      End If
      
      'Ajustando Fecha de 1er Vencimiento
'      r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
'      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   Else
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_FecAju = ff_DiaHabil(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.0000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1500
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.0000"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.0000"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.0000"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.0000"))   'Redondeando
      
      'Anterior
      'If r_dbl_SalCap = 0# Then
      '   Exit For
      'ElseIf r_dbl_SalCap > 0 Then
      '   If r_dbl_SalCap > 10 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen + 0.1
      '   ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen + 0.001
      '   ElseIf r_dbl_SalCap < 0.5 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
      '   End If
      'ElseIf r_dbl_SalCap < 0 Then
      '   If r_dbl_SalCap < -10 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.01
      '   ElseIf r_dbl_SalCap < -5 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.01
      '   ElseIf r_dbl_SalCap < -0.3 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.001
      '   Else
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
      '   End If
      'End If
      
      'Actual
      'If r_dbl_SalCap = 0# Then
      '   Exit For
      'ElseIf r_dbl_SalCap > 0 Then
      '   If r_dbl_SalCap > 10 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen + 0.1
      '   ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
      '   ElseIf r_dbl_SalCap < 0.5 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen + 0.00001
      '   End If
      'ElseIf r_dbl_SalCap < 0 Then
      '   If r_dbl_SalCap < -10 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.01
      '   ElseIf r_dbl_SalCap < -5 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.01
      '   ElseIf r_dbl_SalCap < -0.3 Then
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
      '   Else
      '      r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
      '   End If
      'End If
      
      'Nuevo 2008
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
      ElseIf r_dbl_SalCap < 0 Then
         r_dbl_CuoMen = r_dbl_CuoMen - 0.0001
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.0000"))        'Redondeando Valor Cuota
      
      DoEvents
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Cronog_Mihogar_NCCof_01(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.00000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

'   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
'      r_int_PosMes = r_int_PosMes + 1
'      If r_int_PosMes = 13 Then
'         r_int_PosMes = 1
'         r_int_PosAno = r_int_PosAno + 1
'      End If
      
      'Ajustando Fecha de 1er Vencimiento
'      r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
'      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   Else
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_FecAju = ff_DiaHabil(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.0000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1500
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
      
      DoEvents
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Cronog_Mihogar_NCCof_02(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.0000000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

'   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
'      r_int_PosMes = r_int_PosMes + 1
'      If r_int_PosMes = 13 Then
'         r_int_PosMes = 1
'         r_int_PosAno = r_int_PosAno + 1
'      End If
      
      'Ajustando Fecha de 1er Vencimiento
'      r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
'      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   Else
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_FecAju = ff_DiaHabil(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.0000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1500
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.00001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
      
      DoEvents
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Cronog_Mihogar_NCCof_03(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   'r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.0000000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

'   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
'      r_int_PosMes = r_int_PosMes + 1
'      If r_int_PosMes = 13 Then
'         r_int_PosMes = 1
'         r_int_PosAno = r_int_PosAno + 1
'      End If
      
      'Ajustando Fecha de 1er Vencimiento
'      r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
'      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   Else
      r_str_PriVct = ff_DiaHabil(r_str_PriVct)
'   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_FecAju = ff_DiaHabil(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.0000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1500
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.00001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
      
      DoEvents
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))        'Redondeando Valor Cuota
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub

Public Sub gs_Cronog_Mihogar_ConCof(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String

   Dim r_dbl_MtoPre        As Double

   
   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   'Calculando Tasas y Factores
   
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   'r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.00000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 6
   If r_int_PosMes > 12 Then
      r_int_PosMes = r_int_PosMes - 12
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   r_str_PriVct = ff_DiaHabil(r_str_PriVct)

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = p_FecDes
   r_str_FecSgt = r_str_PriVct
   
   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 6
         
         If r_int_PosMes > 12 Then
            r_int_PosMes = r_int_PosMes - 12
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   
   r_dbl_MtoPre = r_dbl_SalCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      
      r_str_FecSgt = Format(ff_Ultimo_Dia_Mes(r_int_PosMes, r_int_PosAno), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_FecSgt = ff_DiaHabil(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoPre / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   For r_int_NumIte = 1 To 1000
      r_dbl_SalCap = r_dbl_MtoPre
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         'r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         'r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 4))
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00"))
   r_dbl_SalCap = r_dbl_MtoPre
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap
End Sub

Public Sub gs_Cronog_Mihogar_Con(p_Arregl() As modcal_g_est_CuoCli, p_ArrOrg() As modcal_g_est_CuoCli, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String

   Dim r_dbl_MtoPre        As Double

   
   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      'r_int_PosMes = r_int_PosMes + 1
      'If r_int_PosMes = 13 Then
      '   r_int_PosMes = 1
      '   r_int_PosAno = r_int_PosAno + 1
      'End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   Else
      r_int_PosMes = r_int_PosMes - 1
      If r_int_PosMes = 0 Then
         r_int_PosMes = 12
         r_int_PosAno = r_int_PosAno - 1
      End If
   End If
   
   r_int_PosMes = r_int_PosMes + 6
   If r_int_PosMes > 12 Then
      r_int_PosMes = r_int_PosMes - 12
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = p_FecDes
   r_str_FecSgt = r_str_PriVct
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = 0
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_ValCuo = 0
         p_Arregl(r_int_Contad).CuoCli_SalCap = 0
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 6
         
         If r_int_PosMes > 12 Then
            r_int_PosMes = r_int_PosMes - 12
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
   End If
   
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
      r_int_AcuDia = r_int_AcuDia + r_int_DifDia
      
      p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
      p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
      p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
      p_Arregl(r_int_Contad).CuoCli_Factor = 0
   
      'Calculando Siguiente Vencimiento
      r_str_FecIni = r_str_FecSgt
      
      r_int_PosMes = r_int_PosMes + 6
      
      If r_int_PosMes > 12 Then
         r_int_PosMes = r_int_PosMes - 12
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   Next r_int_Contad
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   
   'Generando Cronograma
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = (p_ArrOrg(r_int_Contad).CuoCli_SalCap + p_ArrOrg(r_int_Contad).CuoCli_Capita) * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - (p_ArrOrg(r_int_Contad).CuoCli_SalCap + p_ArrOrg(r_int_Contad).CuoCli_Capita)
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = p_ArrOrg(r_int_Contad).CuoCli_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_ValCuo = p_ArrOrg(r_int_Contad).CuoCli_Capita + r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_SalCap = p_ArrOrg(r_int_Contad).CuoCli_SalCap
   Next r_int_Contad
End Sub

Public Sub gs_Cronog_Mihogar_Con_1(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String

   Dim r_dbl_MtoPre        As Double

   
   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   'Calculando Tasas y Factores
   
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) - 1)
   
   'Calculando Primer Vencimiento (Verificando Diferencia de días)
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If
   
   r_str_FecIni = r_str_PriVct
   
   'Obteniendo Primer Vencimiento
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes + 6
   If r_int_PosMes > 12 Then
      r_int_PosMes = r_int_PosMes - 12
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   'Ajustando Fecha de Inicio de Pagos
   'r_str_FecIni = r_str_FecIni
   r_str_FecSgt = r_str_PriVct
   
   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_ValCuo = 0
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 6
         
         If r_int_PosMes > 12 Then
            r_int_PosMes = r_int_PosMes - 12
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
   End If
   
   
   r_dbl_MtoPre = r_dbl_SalCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
   
   r_dbl_CuoMen = r_dbl_MtoPre / r_dbl_FacAcu
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   r_dbl_SalCap = r_dbl_MtoPre
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap
End Sub

Public Sub gs_Cronog_CRCPBP_NC(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_PorCon As Double, ByVal p_TopCon As Double, ByVal p_TipCam As Double, ByVal p_ValInm As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_TasDes As Double, ByVal p_TipSGV As Integer, ByVal p_TasViv As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByRef p_MtoNCo As Double, ByRef p_MtoCon As Double, Optional ByRef p_IntCap As Double, Optional ByVal p_Cronog As Integer)
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
   
   p_IntCap = 0
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasViv / 100 * p_ValInm
      
      If p_Cronog <> 1 Then
         r_dbl_ImpSgV_Men = r_dbl_ImpSgV_Men * 0.72
      End If
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
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))   'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + (r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men)
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         'p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         'p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         'p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         'p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_Capita = 0
         'p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         'p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
         'p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_Portes = 0
         'p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         'p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
      
      p_IntCap = r_dbl_IntCap
   End If
   
   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = r_dbl_SalCap * p_PorCon / 100
   
   If r_dbl_MtoCon * p_TipCam > p_TopCon Then
      r_dbl_MtoCon = p_TopCon / p_TipCam
   End If
   
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

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 Then
      r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      r_dbl_ImpSgD_Men = p_MtoPre * (1 + r_dbl_DesDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
      
      p_Arregl(1).CuoCli_Intere = p_Arregl(1).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1).CuoCli_SegPre = p_Arregl(1).CuoCli_SegPre + r_dbl_ImpSgD_Men
      p_Arregl(1).CuoCli_SegViv = p_Arregl(1).CuoCli_SegViv + r_dbl_ImpSgV_Men
      p_Arregl(1).CuoCli_ValCuo = p_Arregl(1).CuoCli_ValCuo + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
   End If
End Sub

Public Sub gs_Cronog_CRCPBP_NCMVI(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_PorCon As Double, ByVal p_TopCon As Double, ByVal p_TipCam As Double, ByVal p_ValInm As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_TasDes As Double, ByVal p_TipSGV As Integer, ByVal p_TasViv As Double, ByVal p_Portes As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer, ByRef p_MtoNCo As Double, ByRef p_MtoCon As Double, Optional ByRef p_IntCap As Double, Optional ByVal p_Cronog As Integer)
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
   Dim r_dbl_SalAux        As Double
   
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


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   p_IntCap = 0
   
   'Calculando Seguro de Inmueble
   If p_TipSGV = 1 Then
      r_dbl_ImpSgV_Men = p_TasViv / 100 * p_ValInm
      
      If p_Cronog <> 1 Then
         r_dbl_ImpSgV_Men = r_dbl_ImpSgV_Men * 0.72
      End If
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
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_FecAju) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   r_dbl_SalCap = p_MtoPre
   
   If p_PerGra > 0 Then
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_ImpSgD_Men = r_dbl_SalCap * (1 + r_dbl_DesDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))   'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + (r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men)
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_SegPre = r_dbl_ImpSgD_Men
         p_Arregl(r_int_Contad).CuoCli_SegViv = r_dbl_ImpSgV_Men
         p_Arregl(r_int_Contad).CuoCli_Portes = 0
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      Next r_int_Contad
      
      p_IntCap = r_dbl_IntCap
   End If
   
   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = r_dbl_SalCap * p_PorCon / 100
   
   If r_dbl_MtoCon * p_TipCam > p_TopCon Then
      r_dbl_MtoCon = p_TopCon / p_TipCam
   End If
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = r_dbl_SalCap - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   p_MtoNCo = r_dbl_MtoNCo
   p_MtoCon = r_dbl_MtoCon
   
   'Recalculando Tasa Diaria de Interes por Seguro Desgravamen
   r_dbl_DesDia = (1 + (p_TasDes / 100 / (r_dbl_MtoNCo / (r_dbl_MtoNCo + r_dbl_MtoCon)))) ^ (1 / 30) - 1
   r_dbl_DesDia = CDbl(Format(r_dbl_DesDia, "##0.00000000000"))
   
   r_dbl_Factor = r_dbl_IntDia + r_dbl_DesDia
   
   If p_PerGra > 0 Then
      r_dbl_SalAux = r_dbl_MtoNCo - p_IntCap
   
      For r_int_Contad = 1 To p_PerGra
         r_dbl_SalAux = r_dbl_SalAux + p_Arregl(r_int_Contad).CuoCli_ValCuo
   
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalAux
      Next r_int_Contad
   End If
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SegPre = r_dbl_ImpSgD_Men
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SegViv = r_dbl_ImpSgV_Men
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Portes = p_Portes
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo + r_dbl_ImpSgV_Men + p_Portes
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 Then
      r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      r_dbl_ImpSgD_Men = p_MtoPre * (1 + r_dbl_DesDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))              'Redondeando
      r_dbl_ImpSgD_Men = CDbl(Format(r_dbl_ImpSgD_Men, "######0.00"))      'Redondeando
      
      p_Arregl(p_PerGra + 1).CuoCli_Intere = p_Arregl(p_PerGra + 1).CuoCli_Intere + r_dbl_Intere
      p_Arregl(p_PerGra + 1).CuoCli_SegPre = p_Arregl(p_PerGra + 1).CuoCli_SegPre + r_dbl_ImpSgD_Men
      p_Arregl(p_PerGra + 1).CuoCli_SegViv = p_Arregl(p_PerGra + 1).CuoCli_SegViv + r_dbl_ImpSgV_Men
      p_Arregl(p_PerGra + 1).CuoCli_ValCuo = p_Arregl(p_PerGra + 1).CuoCli_ValCuo + r_dbl_Intere + r_dbl_ImpSgD_Men + r_dbl_ImpSgV_Men
   End If
End Sub

Public Sub gs_Cronog_CRCPBP_ConMVi(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer)
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
   
   'Calculando Primer Vencimiento TNC
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento TNC
      r_str_PriVct = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(CInt(p_DiaPag), "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   

   'Ajustando Fecha de Inicio de Pagos
   'r_str_FecIni = r_str_FecAju
   r_str_FecIni = p_FecDes
   
   If p_PerGra > 0 Then
      r_int_PosMes = Month(CDate(r_str_FecAju))
      r_int_PosAno = Year(CDate(r_str_FecAju))
      
      r_int_PosMes = r_int_PosMes + p_PerGra
      
      If r_int_PosMes > 12 Then
         r_int_PosMes = r_int_PosMes - 12
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      r_str_FecIni = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   End If
   
   'Obteniendo Primer Vencimiento TC
   r_int_PosMes = r_int_PosMes + 6
   If r_int_PosMes > 12 Then
      r_int_PosMes = r_int_PosMes - 12
      r_int_PosAno = r_int_PosAno + 1
   End If
   r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   
   
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


Public Sub gs_Cronog_CRCPBP_ConCli(p_Arregl() As modcal_g_est_CuoCli, p_ArrDes() As modcal_g_est_CuoCli, ByVal p_NumCuo As Integer, ByVal p_TasInt As Double)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_int_Contad        As Integer
   
   ReDim p_ArrDes(p_NumCuo)
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1       'Calculando Tasa Diaria de Interes
   r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   
   'Generando Cronograma
   For r_int_Contad = 1 To p_NumCuo
      r_dbl_Intere = (p_Arregl(r_int_Contad).CuoCli_SalCap + p_Arregl(r_int_Contad).CuoCli_Capita) * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - (p_Arregl(r_int_Contad).CuoCli_SalCap + p_Arregl(r_int_Contad).CuoCli_Capita)
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      p_ArrDes(r_int_Contad).CuoCli_FecVct = p_Arregl(r_int_Contad).CuoCli_FecVct
      p_ArrDes(r_int_Contad).CuoCli_DifDia = p_Arregl(r_int_Contad).CuoCli_DifDia
      p_ArrDes(r_int_Contad).CuoCli_AcuDia = p_Arregl(r_int_Contad).CuoCli_AcuDia
      p_ArrDes(r_int_Contad).CuoCli_Factor = p_Arregl(r_int_Contad).CuoCli_Factor
      p_ArrDes(r_int_Contad).CuoCli_Capita = p_Arregl(r_int_Contad).CuoCli_Capita
      p_ArrDes(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_ArrDes(r_int_Contad).CuoCli_ValCuo = p_Arregl(r_int_Contad).CuoCli_Capita + r_dbl_Intere
      p_ArrDes(r_int_Contad).CuoCli_SalCap = p_Arregl(r_int_Contad).CuoCli_SalCap
   Next r_int_Contad
End Sub

Public Function ff_IntMen(p_TasInt As Double) As Double
   ff_IntMen = (1 + p_TasInt / 100) ^ (8.33333333333333E-02) - 1
End Function

Public Function ff_DiaHabil(p_FecPag As String) As String
   Dim r_int_DiaSem        As Integer
   Dim r_str_FecPag        As String
   Dim r_int_Contad        As Integer
   Dim r_int_DiaPag        As Integer
   Dim r_int_MesPag        As Integer
   Dim r_int_DiaAux        As Integer
   Dim r_int_MesAux        As Integer
   Dim r_int_FlgFer        As Integer
   Dim r_arr_DiaFer(10)    As String
   Dim r_arr_SemSan()      As String
   
   Dim r_int_Sta_PerAno    As Integer
   Dim r_int_Sta_PerMes    As Integer
   Dim r_int_Sta_Cons01    As Integer
   Dim r_int_Sta_Cons02    As Integer
   Dim r_int_Sta_Varia1    As Integer
   Dim r_int_Sta_Varia2    As Integer
   Dim r_int_Sta_Varia3    As Integer
   Dim r_int_Sta_Varia4    As Integer
   Dim r_int_Sta_Varia5    As Integer
   Dim r_int_Sta_DiaFin    As Integer
   Dim r_int_Sta_MesFin    As Integer
   Dim r_str_Sta_DomRes    As String
      
   'Cargando Días Feriados Constantes
   r_arr_DiaFer(1) = "01/01"
   r_arr_DiaFer(2) = "01/05"
   r_arr_DiaFer(3) = "29/06"
   r_arr_DiaFer(4) = "28/07"
   r_arr_DiaFer(5) = "29/07"
   r_arr_DiaFer(6) = "30/08"
   r_arr_DiaFer(7) = "08/10"
   r_arr_DiaFer(8) = "01/11"
   r_arr_DiaFer(9) = "08/12"
   r_arr_DiaFer(10) = "25/12"
   
   r_int_Sta_PerAno = Year(CDate(p_FecPag))
   r_int_Sta_PerMes = Month(CDate(p_FecPag))
   
   r_int_Sta_DiaFin = 0
   r_int_Sta_MesFin = 0
   
   ReDim r_arr_SemSan(0)
   
   If r_int_Sta_PerMes = 3 Or r_int_Sta_PerMes = 4 Then
      If r_int_Sta_PerAno >= 1583 And r_int_Sta_PerAno <= 1699 Then
         r_int_Sta_Cons01 = 22:  r_int_Sta_Cons02 = 2
      ElseIf r_int_Sta_PerAno >= 1700 And r_int_Sta_PerAno <= 1799 Then
         r_int_Sta_Cons01 = 23:  r_int_Sta_Cons02 = 3
      ElseIf r_int_Sta_PerAno >= 1800 And r_int_Sta_PerAno <= 1899 Then
         r_int_Sta_Cons01 = 23:  r_int_Sta_Cons02 = 4
      ElseIf r_int_Sta_PerAno >= 1900 And r_int_Sta_PerAno <= 2099 Then
         r_int_Sta_Cons01 = 24:  r_int_Sta_Cons02 = 5
      ElseIf r_int_Sta_PerAno >= 2100 And r_int_Sta_PerAno <= 2199 Then
         r_int_Sta_Cons01 = 24:  r_int_Sta_Cons02 = 6
      ElseIf r_int_Sta_PerAno >= 2200 And r_int_Sta_PerAno <= 2299 Then
         r_int_Sta_Cons01 = 25:  r_int_Sta_Cons02 = 0
      End If
      
      r_int_Sta_Varia1 = r_int_Sta_PerAno Mod 19
      r_int_Sta_Varia2 = r_int_Sta_PerAno Mod 4
      r_int_Sta_Varia3 = r_int_Sta_PerAno Mod 7
      r_int_Sta_Varia4 = ((19 * r_int_Sta_Varia1) + r_int_Sta_Cons01) Mod 30
      r_int_Sta_Varia5 = ((2 * r_int_Sta_Varia2) + (4 * r_int_Sta_Varia3) + (6 * r_int_Sta_Varia4) + r_int_Sta_Cons02) Mod 7
      
      If r_int_Sta_Varia4 + r_int_Sta_Varia5 < 10 Then
         r_int_Sta_DiaFin = r_int_Sta_Varia4 + r_int_Sta_Varia5 + 22
         r_int_Sta_MesFin = 3
      Else
         r_int_Sta_DiaFin = r_int_Sta_Varia4 + r_int_Sta_Varia5 - 9
         r_int_Sta_MesFin = 4
      End If
      
      If r_int_Sta_DiaFin = 26 And r_int_Sta_MesFin = 4 Then
         r_int_Sta_DiaFin = 19
      ElseIf r_int_Sta_DiaFin = 25 And r_int_Sta_MesFin = 4 And r_int_Sta_Varia4 = 28 And r_int_Sta_Varia5 = 6 And r_int_Sta_Varia1 > 10 Then
         r_int_Sta_DiaFin = 18
      End If
      
      r_str_Sta_DomRes = Format(r_int_Sta_DiaFin, "00") & "/" & Format(r_int_Sta_MesFin, "00") & "/" & Format(r_int_Sta_PerAno, "0000")
      
      ReDim r_arr_SemSan(4)
      For r_int_Contad = 1 To 4
         r_arr_SemSan(r_int_Contad) = r_str_Sta_DomRes
         
         r_str_Sta_DomRes = CDate(r_str_Sta_DomRes) - CDate(1)
      Next r_int_Contad
   End If
   
   
   r_str_FecPag = p_FecPag
   
   Do While 1 = 1
      r_int_DiaSem = Weekday(CDate(r_str_FecPag))
   
      'Validando que no sea Sábado o Domingo
      If Not (r_int_DiaSem = vbSunday Or r_int_DiaSem = vbSaturday) Then
         'Validando que no sea Feriado
         r_int_DiaPag = Day(CDate(r_str_FecPag))
         r_int_MesPag = Month(CDate(r_str_FecPag))
         
         r_int_FlgFer = 0
         For r_int_Contad = 1 To UBound(r_arr_DiaFer)
            r_int_DiaAux = CInt(Left(r_arr_DiaFer(r_int_Contad), 2))
            r_int_MesAux = CInt(Right(r_arr_DiaFer(r_int_Contad), 2))
            
            If r_int_DiaPag = r_int_DiaAux And r_int_MesPag = r_int_MesAux Then
               r_int_FlgFer = 1
               Exit For
            End If
         Next r_int_Contad
         
         For r_int_Contad = 1 To UBound(r_arr_SemSan)
            If CDate(r_arr_SemSan(r_int_Contad)) = CDate(r_str_FecPag) Then
               r_int_FlgFer = 1
               Exit For
            End If
         Next r_int_Contad
         
         If r_int_FlgFer = 0 Then
            Exit Do
         End If
      End If
      
      r_str_FecPag = Format(CDate(r_str_FecPag) - CDate(1), "dd/mm/yyyy")
   Loop
   
   ff_DiaHabil = Format(r_str_FecPag, "dd/mm/yyyy")
End Function

Public Function ff_DiaHabil_Adelante(p_FecPag As String) As String
   Dim r_int_DiaSem        As Integer
   Dim r_str_FecPag        As String
   Dim r_int_Contad        As Integer
   Dim r_int_DiaPag        As Integer
   Dim r_int_MesPag        As Integer
   Dim r_int_DiaAux        As Integer
   Dim r_int_MesAux        As Integer
   Dim r_int_FlgFer        As Integer
   Dim r_arr_DiaFer(10)    As String
   Dim r_arr_SemSan()      As String
   
   Dim r_int_Sta_PerAno    As Integer
   Dim r_int_Sta_PerMes    As Integer
   Dim r_int_Sta_Cons01    As Integer
   Dim r_int_Sta_Cons02    As Integer
   Dim r_int_Sta_Varia1    As Integer
   Dim r_int_Sta_Varia2    As Integer
   Dim r_int_Sta_Varia3    As Integer
   Dim r_int_Sta_Varia4    As Integer
   Dim r_int_Sta_Varia5    As Integer
   Dim r_int_Sta_DiaFin    As Integer
   Dim r_int_Sta_MesFin    As Integer
   Dim r_str_Sta_DomRes    As String
      
   'Cargando Días Feriados Constantes
   r_arr_DiaFer(1) = "01/01"
   r_arr_DiaFer(2) = "01/05"
   r_arr_DiaFer(3) = "29/06"
   r_arr_DiaFer(4) = "28/07"
   r_arr_DiaFer(5) = "29/07"
   r_arr_DiaFer(6) = "30/08"
   r_arr_DiaFer(7) = "08/10"
   r_arr_DiaFer(8) = "01/11"
   r_arr_DiaFer(9) = "08/12"
   r_arr_DiaFer(10) = "25/12"
   
   r_int_Sta_PerAno = Year(CDate(p_FecPag))
   r_int_Sta_PerMes = Month(CDate(p_FecPag))
   
   r_int_Sta_DiaFin = 0
   r_int_Sta_MesFin = 0
   
   ReDim r_arr_SemSan(0)
   
   If r_int_Sta_PerMes = 3 Or r_int_Sta_PerMes = 4 Then
      If r_int_Sta_PerAno >= 1583 And r_int_Sta_PerAno <= 1699 Then
         r_int_Sta_Cons01 = 22:  r_int_Sta_Cons02 = 2
      ElseIf r_int_Sta_PerAno >= 1700 And r_int_Sta_PerAno <= 1799 Then
         r_int_Sta_Cons01 = 23:  r_int_Sta_Cons02 = 3
      ElseIf r_int_Sta_PerAno >= 1800 And r_int_Sta_PerAno <= 1899 Then
         r_int_Sta_Cons01 = 23:  r_int_Sta_Cons02 = 4
      ElseIf r_int_Sta_PerAno >= 1900 And r_int_Sta_PerAno <= 2099 Then
         r_int_Sta_Cons01 = 24:  r_int_Sta_Cons02 = 5
      ElseIf r_int_Sta_PerAno >= 2100 And r_int_Sta_PerAno <= 2199 Then
         r_int_Sta_Cons01 = 24:  r_int_Sta_Cons02 = 6
      ElseIf r_int_Sta_PerAno >= 2200 And r_int_Sta_PerAno <= 2299 Then
         r_int_Sta_Cons01 = 25:  r_int_Sta_Cons02 = 0
      End If
      
      r_int_Sta_Varia1 = r_int_Sta_PerAno Mod 19
      r_int_Sta_Varia2 = r_int_Sta_PerAno Mod 4
      r_int_Sta_Varia3 = r_int_Sta_PerAno Mod 7
      r_int_Sta_Varia4 = ((19 * r_int_Sta_Varia1) + r_int_Sta_Cons01) Mod 30
      r_int_Sta_Varia5 = ((2 * r_int_Sta_Varia2) + (4 * r_int_Sta_Varia3) + (6 * r_int_Sta_Varia4) + r_int_Sta_Cons02) Mod 7
      
      If r_int_Sta_Varia4 + r_int_Sta_Varia5 < 10 Then
         r_int_Sta_DiaFin = r_int_Sta_Varia4 + r_int_Sta_Varia5 + 22
         r_int_Sta_MesFin = 3
      Else
         r_int_Sta_DiaFin = r_int_Sta_Varia4 + r_int_Sta_Varia5 - 9
         r_int_Sta_MesFin = 4
      End If
      
      If r_int_Sta_DiaFin = 26 And r_int_Sta_MesFin = 4 Then
         r_int_Sta_DiaFin = 19
      ElseIf r_int_Sta_DiaFin = 25 And r_int_Sta_MesFin = 4 And r_int_Sta_Varia4 = 28 And r_int_Sta_Varia5 = 6 And r_int_Sta_Varia1 > 10 Then
         r_int_Sta_DiaFin = 18
      End If
      
      r_str_Sta_DomRes = Format(r_int_Sta_DiaFin, "00") & "/" & Format(r_int_Sta_MesFin, "00") & "/" & Format(r_int_Sta_PerAno, "0000")
      
      ReDim r_arr_SemSan(4)
      For r_int_Contad = 1 To 4
         r_arr_SemSan(r_int_Contad) = r_str_Sta_DomRes
         
         r_str_Sta_DomRes = CDate(r_str_Sta_DomRes) - CDate(1)
      Next r_int_Contad
   End If
   
   r_str_FecPag = p_FecPag
   
   Do While 1 = 1
      r_int_DiaSem = Weekday(CDate(r_str_FecPag))
   
      'Validando que no sea Sábado o Domingo
      If Not (r_int_DiaSem = vbSunday Or r_int_DiaSem = vbSaturday) Then
         'Validando que no sea Feriado
         r_int_DiaPag = Day(CDate(r_str_FecPag))
         r_int_MesPag = Month(CDate(r_str_FecPag))
         
         r_int_FlgFer = 0
         For r_int_Contad = 1 To UBound(r_arr_DiaFer)
            r_int_DiaAux = CInt(Left(r_arr_DiaFer(r_int_Contad), 2))
            r_int_MesAux = CInt(Right(r_arr_DiaFer(r_int_Contad), 2))
            
            If r_int_DiaPag = r_int_DiaAux And r_int_MesPag = r_int_MesAux Then
               r_int_FlgFer = 1
               Exit For
            End If
         Next r_int_Contad
         
         For r_int_Contad = 1 To UBound(r_arr_SemSan)
            If CDate(r_arr_SemSan(r_int_Contad)) = CDate(r_str_FecPag) Then
               r_int_FlgFer = 1
               Exit For
            End If
         Next r_int_Contad
         
         If r_int_FlgFer = 0 Then
            Exit Do
         End If
      End If
      
      r_str_FecPag = Format(CDate(r_str_FecPag) + CDate(1), "dd/mm/yyyy")
   Loop
   
   ff_DiaHabil_Adelante = Format(r_str_FecPag, "dd/mm/yyyy")
End Function

Public Function ff_Ultimo_Dia_Mes(ByVal p_Mes As Integer, p_Ano As Integer) As Integer
   ff_Ultimo_Dia_Mes = 0
   
   If p_Mes = 1 Or p_Mes = 3 Or p_Mes = 5 Or p_Mes = 7 Or p_Mes = 8 Or p_Mes = 10 Or p_Mes = 12 Then
      ff_Ultimo_Dia_Mes = 31
   ElseIf p_Mes = 4 Or p_Mes = 6 Or p_Mes = 9 Or p_Mes = 11 Then
      ff_Ultimo_Dia_Mes = 30
   ElseIf p_Mes = 2 Then
      If p_Ano Mod 4 = 0 Then   'Año Bisiesto
         ff_Ultimo_Dia_Mes = 29
      Else
         ff_Ultimo_Dia_Mes = 28
      End If
   End If
End Function


Private Function ff_Truncar_Numero(ByVal p_Numero As Double, ByVal p_NumDec As Integer) As String
   Dim r_str_Numero  As String
   Dim r_int_PtoDec  As Integer
   Dim r_str_Entero  As String
   Dim r_str_Decima  As String
   
   r_str_Numero = CStr(p_Numero)
   
   r_int_PtoDec = InStr(r_str_Numero, ".")
   
   If r_int_PtoDec > 0 Then
      r_str_Entero = Left(r_str_Numero, r_int_PtoDec - 1)
      r_str_Decima = Mid(r_str_Numero, r_int_PtoDec + 1, p_NumDec)
   Else
      r_str_Entero = r_str_Numero
      r_str_Decima = String(p_NumDec, "0")
   End If
   
   ff_Truncar_Numero = r_str_Entero & "." & r_str_Decima
End Function

Public Function gf_Calculo_CostoEfectivo(p_Arregl() As String, ByVal p_TasInt As Double, ByVal p_MtoPre As Double) As Double
Dim r_int_Contad     As Integer
Dim r_dbl_CuoAcu     As Double
Dim r_dbl_TasInt     As Double
Dim i                As Double
   
   gf_Calculo_CostoEfectivo = 0
   r_dbl_CuoAcu = p_MtoPre
   
   For i = 0.01 To 0.5 Step 0.0001
      r_dbl_CuoAcu = 0
      
      For r_int_Contad = 1 To UBound(p_Arregl)
         r_dbl_CuoAcu = r_dbl_CuoAcu + CDbl(Format(CDbl(p_Arregl(r_int_Contad, 9)) / ((1 + i) ^ (p_Arregl(r_int_Contad, 11) / 360)), "#####0.00"))
      Next r_int_Contad
      
      If r_dbl_CuoAcu < p_MtoPre Then
         Exit For
      End If
   Next i
      
   gf_Calculo_CostoEfectivo = CDbl(Format(i * 100, "##0.0000"))
End Function

Public Function gf_Cronog_CosEfe(p_Arregl() As modcal_g_est_CuoCli, ByVal p_TasInt As Double, ByVal p_MtoPre As Double) As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_CuoAcu     As Double
   Dim r_dbl_TasInt     As Double

   gf_Cronog_CosEfe = 0
   r_dbl_TasInt = p_TasInt
   r_dbl_CuoAcu = p_MtoPre
   
   Do While r_dbl_CuoAcu >= p_MtoPre
      r_dbl_CuoAcu = 0
      
      For r_int_Contad = 1 To UBound(p_Arregl)
         r_dbl_CuoAcu = r_dbl_CuoAcu + CDbl(Format(p_Arregl(r_int_Contad).CuoCli_ValCuo / (1 + (r_dbl_TasInt / 100)) ^ (p_Arregl(r_int_Contad).CuoCli_AcuDia / 360), "#####0.00"))
      Next r_int_Contad
      
      If r_dbl_CuoAcu >= p_MtoPre Then
         r_dbl_TasInt = r_dbl_TasInt + 0.01
      End If
   Loop
   
   gf_Cronog_CosEfe = CDbl(Format(r_dbl_TasInt, "##0.00"))
End Function

Public Function gf_Cronog_CosEfe_MVi(p_Arregl() As modcal_g_est_CuoCli, ByVal p_TasInt As Double, ByVal p_MtoPre As Double) As Double
   Dim r_int_Contad     As Integer
   Dim r_dbl_CuoAcu     As Double
   Dim r_dbl_TasInt     As Double

   gf_Cronog_CosEfe_MVi = 0
   
   r_dbl_TasInt = p_TasInt
   r_dbl_CuoAcu = p_MtoPre
   
   Do While r_dbl_CuoAcu <= p_MtoPre
      r_dbl_CuoAcu = 0
      
      For r_int_Contad = 1 To UBound(p_Arregl)
         r_dbl_CuoAcu = r_dbl_CuoAcu + CDbl(Format(p_Arregl(r_int_Contad).CuoCli_ValCuo / (1 + (r_dbl_TasInt / 100)) ^ (p_Arregl(r_int_Contad).CuoCli_AcuDia / 360), "#####0.00"))
      Next r_int_Contad
      
      If r_dbl_CuoAcu <= p_MtoPre Then
         r_dbl_TasInt = r_dbl_TasInt - 0.01
      End If
   Loop
   
   gf_Cronog_CosEfe_MVi = CDbl(Format(r_dbl_TasInt, "##0.00"))
End Function

Public Sub gs_Cronog_miCasitaCOFIDE_Cof(p_Arregl() As modcal_g_est_CuoCli, p_MtoPre As Double, ByVal p_TopCon As Double, ByVal p_NumCuo As Integer, ByVal p_PerGra As Integer, ByVal p_TasInt As Double, ByVal p_ComCof As Double, ByVal p_FecDes As String, ByVal p_DiaPag As Integer)
   Dim r_dbl_IntDia        As Double
   Dim r_dbl_ComDia        As Double
   Dim r_dbl_Factor        As Double
   Dim r_dbl_Intere        As Double
   Dim r_dbl_Comisi        As Double
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
   Dim r_int_NumIte        As Integer
   Dim r_int_DifDia        As Integer
   Dim r_int_AcuDia        As Integer
   
   Dim r_int_PosMes        As Integer
   Dim r_int_PosAno        As Integer
   Dim r_int_DiaAdc        As Integer
   
   Dim r_str_PriVct        As String
   Dim r_str_FecAju        As String
   Dim r_str_FecIni        As String
   Dim r_str_FecSgt        As String


   ReDim p_Arregl(p_NumCuo + p_PerGra)
   
   
   'Calculando Tasas y Factores
   r_dbl_IntDia = (1 + (p_TasInt / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes
   r_dbl_ComDia = (1 + (p_ComCof / 100)) ^ (1 / 360) - 1     'Calculando Tasa Diaria de Interes por Seguro Desgravamen
   
   'r_dbl_IntDia = CDbl(Format(r_dbl_IntDia, "##0.00000000000"))
   r_dbl_ComDia = CDbl(Format(r_dbl_ComDia, "##0.0000000000000"))

   r_dbl_Factor = ((1 + r_dbl_IntDia) * (1 + r_dbl_ComDia) - 1)
   
   'Calculando Primer Vencimiento
   r_int_PosMes = Month(CDate(p_FecDes))
   r_int_PosAno = Year(CDate(p_FecDes))
   
   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If
   
   r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")

   If CInt(CDate(r_str_PriVct) - CDate(p_FecDes)) < 30 Then
      r_int_PosMes = r_int_PosMes + 1
      If r_int_PosMes = 13 Then
         r_int_PosMes = 1
         r_int_PosAno = r_int_PosAno + 1
      End If
      
      'Ajustando Fecha de 1er Vencimiento
      r_str_PriVct = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
      r_str_PriVct = ff_DiaHabil_Adelante(r_str_PriVct)
   Else
      r_str_PriVct = ff_DiaHabil_Adelante(r_str_PriVct)
   End If

   'Calculando Fecha de Ajuste
   r_int_PosMes = Month(CDate(r_str_PriVct))
   r_int_PosAno = Year(CDate(r_str_PriVct))
   
   r_int_PosMes = r_int_PosMes - 1
   If r_int_PosMes = 0 Then
      r_int_PosMes = 12
      r_int_PosAno = r_int_PosAno - 1
   End If
   
   r_str_FecAju = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
   'r_str_FecAju = ff_DiaHabil_Adelante(r_str_FecAju)
   
   'Días a capitalizar por Diferencia de Fechas entre Fecha de Desembolso y Fecha de Ajuste por Día de Pago escogido por el Cliente
   r_int_DiaAdc = CInt(CDate(r_str_PriVct) - CDate(p_FecDes))

   'Ajustando Fecha de Inicio de Pagos
   r_str_FecIni = r_str_FecAju
   r_str_FecSgt = r_str_PriVct

   r_int_PosMes = r_int_PosMes + 1
   If r_int_PosMes = 13 Then
      r_int_PosMes = 1
      r_int_PosAno = r_int_PosAno + 1
   End If

   'Obteniendo Monto Tramo No Concesional
   r_dbl_MtoCon = p_TopCon
   
   r_dbl_MtoCon = CDbl(Format(r_dbl_MtoCon, "#####0.00"))
   r_dbl_MtoNCo = p_MtoPre - r_dbl_MtoCon
   r_dbl_MtoNCo = CDbl(Format(r_dbl_MtoNCo, "#####0.00"))
   
   r_dbl_SalCap = r_dbl_MtoNCo
   
   If p_PerGra > 0 Then
      r_str_FecIni = p_FecDes
      r_int_AcuDia = 0
      r_dbl_IntCap = 0
      
      For r_int_Contad = 1 To p_PerGra
         r_int_DifDia = CInt(CDate(r_str_FecSgt) - CDate(r_str_FecIni))
         r_int_AcuDia = r_int_AcuDia + r_int_DifDia
            
         'Calculando Interes
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))           'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ r_int_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_SalCap = r_dbl_SalCap + r_dbl_Intere
         r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))           'Redondeando
         
         r_dbl_IntCap = r_dbl_IntCap + r_dbl_Intere
         r_dbl_IntCap = CDbl(Format(r_dbl_IntCap, "######0.00"))           'Redondeando
         
         'Almacenando en arreglo
         p_Arregl(r_int_Contad).CuoCli_FecVct = r_str_FecSgt
         p_Arregl(r_int_Contad).CuoCli_DifDia = r_int_DifDia
         p_Arregl(r_int_Contad).CuoCli_AcuDia = r_int_AcuDia
         p_Arregl(r_int_Contad).CuoCli_IntCap = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Capita = 0
         p_Arregl(r_int_Contad).CuoCli_Intere = 0
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
         
         r_str_FecIni = r_str_FecSgt
      
         'Obteniendo siguiente Vencimiento
         r_int_PosMes = r_int_PosMes + 1
         
         If r_int_PosMes = 13 Then
            r_int_PosMes = 1
            r_int_PosAno = r_int_PosAno + 1
         End If
         
         r_str_FecSgt = Format(p_DiaPag, "00") & "/" & Format(r_int_PosMes, "00") & "/" & Format(r_int_PosAno, "0000")
         r_str_FecSgt = ff_DiaHabil_Adelante(r_str_FecSgt)
      Next r_int_Contad
   End If
   
   r_dbl_MtoNCo = r_dbl_SalCap
   'r_dbl_SalCap = r_dbl_MtoNCo + r_dbl_IntCap
   
   'Calculando Cuota Mensual
   r_int_AcuDia = 0
   r_dbl_FacAcu = 0
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
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
      r_str_FecSgt = ff_DiaHabil_Adelante(r_str_FecSgt)
   Next r_int_Contad
   
   r_dbl_CuoMen = r_dbl_MtoNCo / r_dbl_FacAcu
   'r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   
   'Generando Cronograma
   'Iteraciones
   
   For r_int_NumIte = 1 To 1000
      r_dbl_SalCap = r_dbl_MtoNCo
      
      For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
         r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         'r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
         
         r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
         r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
         
         r_dbl_ValCuo = r_dbl_CuoMen
      
         r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
         r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
         
         'r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
         
         p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
         p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
         p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
         p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
         p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
      Next r_int_Contad
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      If r_dbl_SalCap = 0# Then
         Exit For
      ElseIf r_dbl_SalCap > 0 Then
         If r_dbl_SalCap > 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.1
         ElseIf r_dbl_SalCap > 0.5 And r_dbl_SalCap < 10 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.001
         ElseIf r_dbl_SalCap < 0.5 Then
            r_dbl_CuoMen = r_dbl_CuoMen + 0.0001
         End If
      ElseIf r_dbl_SalCap < 0 Then
         If r_dbl_SalCap < -10 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -5 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.01
         ElseIf r_dbl_SalCap < -0.3 Then
            r_dbl_CuoMen = r_dbl_CuoMen - 0.001
         Else
            r_dbl_CuoMen = r_dbl_CuoMen - 0.00001
         End If
      End If
      
      r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "######0.00000"))        'Redondeando Valor Cuota
   Next r_int_NumIte
   
   'Recalculando con Cuota Final
   'r_dbl_CuoMen = CDbl(ff_Truncar_Numero(r_dbl_CuoMen, 2))
   r_dbl_CuoMen = CDbl(Format(r_dbl_CuoMen, "#####0.00"))
   r_dbl_SalCap = r_dbl_MtoNCo
   
   For r_int_Contad = 1 + p_PerGra To p_NumCuo + p_PerGra
      r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))            'Redondeando
      
      r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(r_int_Contad).CuoCli_DifDia - r_dbl_SalCap
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))    'Redondeando
      
      r_dbl_ValCuo = r_dbl_CuoMen
   
      r_dbl_Capita = r_dbl_ValCuo - r_dbl_Intere - r_dbl_Comisi
      r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
      
      r_dbl_SalCap = CDbl(Format(r_dbl_SalCap, "######0.00"))   'Redondeando
      
      p_Arregl(r_int_Contad).CuoCli_Capita = r_dbl_Capita
      p_Arregl(r_int_Contad).CuoCli_Intere = r_dbl_Intere
      p_Arregl(r_int_Contad).CuoCli_Comisi = r_dbl_Comisi
      p_Arregl(r_int_Contad).CuoCli_ValCuo = r_dbl_ValCuo
      p_Arregl(r_int_Contad).CuoCli_SalCap = r_dbl_SalCap
   Next r_int_Contad
   
   'Ajustando Ultima Cuota
   r_dbl_SalCap = r_dbl_SalCap + r_dbl_Capita
   r_dbl_Capita = r_dbl_SalCap
   
   r_dbl_Intere = r_dbl_SalCap * (1 + r_dbl_IntDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
   r_dbl_Comisi = r_dbl_SalCap * (1 + r_dbl_ComDia) ^ p_Arregl(p_NumCuo + p_PerGra).CuoCli_DifDia - r_dbl_SalCap
      
   r_dbl_Intere = CDbl(Format(r_dbl_Intere, "######0.00"))        'Redondeando
   r_dbl_Comisi = CDbl(Format(r_dbl_Comisi, "######0.00"))        'Redondeando
   
   r_dbl_ValCuo = r_dbl_Capita + r_dbl_Intere + r_dbl_Comisi
   r_dbl_SalCap = r_dbl_SalCap - r_dbl_Capita
   
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Capita = r_dbl_Capita
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Intere = r_dbl_Intere
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_Comisi = r_dbl_Comisi
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_ValCuo = r_dbl_ValCuo
   p_Arregl(p_NumCuo + p_PerGra).CuoCli_SalCap = r_dbl_SalCap

   'Ajustando Primera Cuota
   If r_int_DiaAdc > 0 And p_PerGra = 0 Then
      'r_dbl_Intere = p_MtoPre * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - p_MtoPre
      'r_dbl_Comisi = p_MtoPre * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - p_MtoPre
      
      r_dbl_Intere = r_dbl_MtoNCo * (1 + r_dbl_IntDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      r_dbl_Comisi = r_dbl_MtoNCo * (1 + r_dbl_ComDia) ^ r_int_DiaAdc - r_dbl_MtoNCo
      
      r_dbl_Intere = CDbl(Format(r_dbl_Intere - p_Arregl(1 + p_PerGra).CuoCli_Intere, "######0.00"))            'Redondeando
      r_dbl_Comisi = CDbl(Format(r_dbl_Comisi - p_Arregl(1 + p_PerGra).CuoCli_Comisi, "######0.00"))     'Redondeando
      
      p_Arregl(1 + p_PerGra).CuoCli_Intere = p_Arregl(1 + p_PerGra).CuoCli_Intere + r_dbl_Intere
      p_Arregl(1 + p_PerGra).CuoCli_Comisi = p_Arregl(1 + p_PerGra).CuoCli_Comisi + r_dbl_Comisi
      p_Arregl(1 + p_PerGra).CuoCli_ValCuo = p_Arregl(1 + p_PerGra).CuoCli_ValCuo + r_dbl_Intere + r_dbl_Comisi
   End If
End Sub



