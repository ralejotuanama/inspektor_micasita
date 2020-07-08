Attribute VB_Name = "modtac"
Option Explicit

Public modtac_g_str_CadCon As String

Public Type modtac_tpo_CtaCtb
   CtaCtb_NumCta        As String
   CtaCtb_FlagDH        As Integer
   CtaCtb_ImpoMN        As Double
   CtaCtb_ImpoME        As Double
End Type

Public Type modtac_tpo_CtaAux
   CtaAux_CtaAux        As String
   CtaAux_ImpDeb        As Double
   CtaAux_ImpHab        As Double
   CtaAux_Tipmon        As Integer
End Type

Public Type modtac_tpo_BalCom
   BalCom_CtaAux        As String
   BalCom_ImpDeb        As Double
   BalCom_ImpHab        As Double
   BalCom_TipMon        As Integer
End Type

Public Type modtac_tpo_CtaRcd
   CtaRcd_NumCta        As String
   CtaRcd_DesVar        As String
   CtaRcd_Import        As Double
   CtaRcd_TipMon        As Integer
End Type

Public Type modtac_tpo_Genera
   Genera_Codigo        As String
   Genera_Nombre        As String
End Type

Public Type modtac_tpo_DatRep
   DatRep_RanFec        As String
   DatRep_VeCoMn        As Double
End Type

Public Type modtac_tpo_Ocupac
   Ocupac_CodPmc        As String
   Ocupac_CodPsb        As String
   Ocupac_Descri        As String
End Type

Public Type modtac_tpo_CtbPrv
   CtbPrv_PvCoMv        As Double
   CtbPrv_PvCoBi        As Double
   CtbPrv_PvCoFi        As Double
   CtbPrv_SalMiv        As Double
   CtbPrv_SalBid        As Double
   CtbPrv_SalFid        As Double
End Type

Public modtac_g_int_PerMes       As Integer
Public modtac_g_int_PerAno       As Integer
Public modtac_g_int_Moneda       As Integer
Public modtac_g_str_NumOpe       As String
Public modtac_g_int_Estado       As Integer
Public modtac_g_str_NroInt       As String

'Validacion de Datos en Tipo Cambio
Public Function modtac_gf_ValidaTipCamDia(ByVal p_FecAct As String, ByVal p_FecFin As String, ByVal p_Codigo As Integer) As Integer
   Dim r_str_FecAux     As String
   Dim r_int_ConFec     As Integer
   Dim r_rst_ValTip        As ADODB.Recordset
        
   modtac_gf_ValidaTipCamDia = 1
        
   g_str_Parame = "SELECT * FROM OPE_TIPCAM WHERE "
   g_str_Parame = g_str_Parame & "TIPCAM_CODIGO = " & p_Codigo & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA <= " & Right(p_FecAct, 4) + Mid(p_FecAct, 4, 2) + Left(p_FecAct, 2) & " "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCAM_FECDIA DESC"
      
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_ValTip, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_ValTip.BOF And r_rst_ValTip.EOF) Then
      r_rst_ValTip.MoveFirst
      r_str_FecAux = p_FecFin
      
      Do While Not r_rst_ValTip.EOF
                           
         If r_int_ConFec = 7 Then
            Exit Do
         End If
         
         If CDate(r_str_FecAux) > CDate(gf_FormatoFecha(CStr(Trim((r_rst_ValTip!TIPCAM_FECDIA))))) Then
            modtac_gf_ValidaTipCamDia = 0
            Exit Function
         End If
               
         r_int_ConFec = r_int_ConFec + 1
         r_str_FecAux = CDate(r_str_FecAux) - 1
         r_rst_ValTip.MoveNext
      Loop
   End If
   
   r_rst_ValTip.Close
   Set r_rst_ValTip = Nothing
End Function

'Encriptar Archivo .INI para la conexion con la BD
'Function gf_Seg_EncArc(p_CadEnc As String) As String
'   Dim r_str_CadEnc     As String
'   Dim r_int_Posici     As Integer
'   Dim r_int_Contad     As Integer'
'
'   gf_Seg_EncArc = ""
'   r_str_CadEnc = ""
'
'   For r_int_Contad = 1 To Len(p_CadEnc$)
'      r_int_Posici = InStr(modgen_g_con_CadOri_Arc, Mid$(p_CadEnc, r_int_Contad, 1))
'
'      If r_int_Posici > 0 Then
'         r_str_CadEnc = r_str_CadEnc + Mid(modgen_g_con_CadEnc_Arc, r_int_Posici, 1)
'      End If
'   Next r_int_Contad
'
'   gf_Seg_EncArc = Trim$(r_str_CadEnc)
'End Function

'Desencriptar Archivo .INI para la conexion con la BD
'Function gf_Seg_DesArc(p_CadEnc As String) As String
'   Dim r_str_CadEnc     As String
'   Dim r_int_Posici     As Integer
'   Dim r_int_Contad     As Integer'

'   gf_Seg_DesArc = ""
'   r_str_CadEnc = ""
'
'   For r_int_Contad = 1 To Len(p_CadEnc)
'      r_int_Posici = InStr(modgen_g_con_CadEnc_Arc, Mid$(p_CadEnc, r_int_Contad, 1))
'      If r_int_Posici > 0 Then
'         r_str_CadEnc = r_str_CadEnc + Mid$(modgen_g_con_CadOri_Arc, r_int_Posici, 1)
'      End If
'   Next r_int_Contad
'
'   gf_Seg_DesArc = Trim$(r_str_CadEnc)
'End Function

'Obtener cuenta contable
Public Sub modtac_gs_Carga_CtaCtb(ByVal p_CodEmp As String, p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodCla As Integer, ByVal p_CodNiv As Integer, ByVal p_CodMon As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   If p_CodNiv = 12 Then
      p_CodNiv = 7
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CNTBL_CNTA "
   g_str_Parame = g_str_Parame & " WHERE FLAG_ESTADO = 1 "
   
   If p_CodCla <> 0 Then
      g_str_Parame = g_str_Parame & "AND SUBSTR(CNTA_CTBL, 1, 1) = '" & CStr(p_CodCla) & "' "
   End If
   
   If p_CodNiv <> 0 Then
      g_str_Parame = g_str_Parame & "AND NIV_CNTA = " & CStr(p_CodNiv) & " "
   End If
   
   If p_CodMon = -1 Then
      g_str_Parame = g_str_Parame & "AND SUBSTR(CNTA_CTBL, 3, 1) <> '0' "
   Else
      g_str_Parame = g_str_Parame & "AND SUBSTR(CNTA_CTBL, 3, 1) = '" & CStr(p_CodMon) & " ' "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!CNTA_CTBL) & " - " & Trim(g_rst_Listas!ABREV_CNTA)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!CNTA_CTBL)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!ABREV_CNTA & "")
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function modtac_gf_ObtieneTipCamDia(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String, ByVal p_TipTip As Integer) As Double
   modtac_gf_ObtieneTipCamDia = 0
   
   'Obteniendo Tipo de Cambio del Dia
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_TIPCAM "
   g_str_Parame = g_str_Parame & " WHERE TIPCAM_CODIGO = " & CStr(p_TipCam) & " "
   g_str_Parame = g_str_Parame & "   AND TIPCAM_FECDIA = " & p_FecDia & " "
   g_str_Parame = g_str_Parame & " ORDER BY TIPCAM_HORDIA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      If p_TipTip = 1 Then
         modtac_gf_ObtieneTipCamDia = g_rst_Genera!TIPCAM_VENTAS
      ElseIf p_TipTip = 2 Then
         modtac_gf_ObtieneTipCamDia = g_rst_Genera!TIPCAM_COMPRA
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

'Tipo de Cambio Sunat
Public Function modtac_gf_ObtieneTipCamDia_2(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String, ByVal p_TipTip As Integer) As Double
   
   modtac_gf_ObtieneTipCamDia_2 = 0
   
   g_str_Parame = "SELECT * FROM CALENDARIO WHERE "
   g_str_Parame = g_str_Parame & "FECHA = to_date(" & p_FecDia & ",'yyyy/mm/dd') ORDER BY "
   g_str_Parame = g_str_Parame & "FECHA DESC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
   
      If p_TipTip = 1 Then
         modtac_gf_ObtieneTipCamDia_2 = g_rst_Genera!PROM_SBS
      ElseIf p_TipTip = 2 Then
         modtac_gf_ObtieneTipCamDia_2 = g_rst_Genera!CMP_DOL_PROM
      Else
         modtac_gf_ObtieneTipCamDia_2 = 0
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

'Tipo de Cambio SBS
Public Function modtac_gf_ObtieneTipCamDia_3(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String, ByVal p_TipTip As Integer) As Double
   
   modtac_gf_ObtieneTipCamDia_3 = 0
      
   g_str_Parame = "SELECT * FROM CALENDARIO WHERE "
   g_str_Parame = g_str_Parame & "FECHA <= to_date(" & p_FecDia & ",'yyyy/mm/dd') ORDER BY "
   g_str_Parame = g_str_Parame & "FECHA DESC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
     
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   'g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      If IsNull(g_rst_Genera!PROM_SBS) Then
         modtac_gf_ObtieneTipCamDia_3 = 0
      Else
         modtac_gf_ObtieneTipCamDia_3 = g_rst_Genera!PROM_SBS
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         Exit Function
      End If
      
      g_rst_Genera.MoveNext
   Loop
     
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

'Validacion de Datos en Tipo Cambio
Public Function modtac_gf_ValidaTipCamDia_2(ByVal p_FecAct As String, ByVal p_FecFin As String) As Integer
   Dim r_str_FecAux     As String
   Dim r_int_ConFec     As Integer
   Dim r_rst_ValTip        As ADODB.Recordset
        
   modtac_gf_ValidaTipCamDia_2 = 1
        
   g_str_Parame = "SELECT * FROM CALENDARIO "
   g_str_Parame = g_str_Parame & " WHERE FECHA <= to_date('" & p_FecFin & "', 'dd/mm/yyyy') "
   g_str_Parame = g_str_Parame & " ORDER BY FECHA DESC"
      
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_ValTip, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_ValTip.BOF And r_rst_ValTip.EOF) Then
      r_rst_ValTip.MoveFirst
      r_str_FecAux = p_FecFin
      
      Do While Not r_rst_ValTip.EOF
         If r_int_ConFec = 7 Then
            Exit Do
         End If
         
         If CDate(r_str_FecAux) > CDate(r_rst_ValTip!Fecha) Then
            modtac_gf_ValidaTipCamDia_2 = 0
            Exit Function
         End If
               
         If IsNull(r_rst_ValTip!VTA_DOL_PROM Or r_rst_ValTip!CMP_DOL_PROM Or r_rst_ValTip!PROM_SBS) Then
            modtac_gf_ValidaTipCamDia_2 = 0
            Exit Function
         Else
            If (r_rst_ValTip!VTA_DOL_PROM <= 0 Or r_rst_ValTip!CMP_DOL_PROM <= 0 Or r_rst_ValTip!PROM_SBS <= 0) Then
               modtac_gf_ValidaTipCamDia_2 = 0
               Exit Function
            End If
         End If
         
         r_int_ConFec = r_int_ConFec + 1
         r_str_FecAux = CDate(r_str_FecAux) - 1
         r_rst_ValTip.MoveNext
      Loop
   End If
   
   r_rst_ValTip.Close
   Set r_rst_ValTip = Nothing
End Function

'=======================================================================================================
Public Function modtac_gf_Buscar_SbsCod(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
   modtac_gf_Buscar_SbsCod = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & Trim(p_NumDoc) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      If IsNull(Trim(g_rst_Listas!DATGEN_CODSBS)) Then
         modtac_gf_Buscar_SbsCod = "NO ASIGNADO"
      Else
         modtac_gf_Buscar_SbsCod = Trim(g_rst_Listas!DATGEN_CODSBS)
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modtac_gf_Buscar_TipCla(ByVal p_TipCre As Integer, ByVal p_Codigo As String) As String
   modtac_gf_Buscar_TipCla = ""
   
   g_str_Parame = "SELECT * FROM CTB_TIPCLA WHERE "
   g_str_Parame = g_str_Parame & "TIPCLA_TIPCRE = '" & CStr(p_TipCre) & "' AND "
   g_str_Parame = g_str_Parame & "TIPCLA_CODIGO = '" & CStr(p_Codigo) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modtac_gf_Buscar_TipCla = Trim(g_rst_Listas!TIPCLA_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modtac_gf_Buscar_TipCre(ByVal p_TipCre As Integer) As String
   modtac_gf_Buscar_TipCre = ""
   
   g_str_Parame = "SELECT * FROM CTB_TIPCRE WHERE "
   g_str_Parame = g_str_Parame & "TIPCRE_CODIGO = '" & CStr(p_TipCre) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modtac_gf_Buscar_TipCre = Trim(g_rst_Listas!TIPCRE_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modtac_gf_Buscar_FecCon(ByVal p_NumOpe As String) As String
   modtac_gf_Buscar_FecCon = ""
   
   g_str_Parame = "SELECT * FROM CRE_HIPGAR WHERE "
   g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPGAR_BIEGAR ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      If Not IsNull(g_rst_Listas!HIPGAR_FECCON) Then
            modtac_gf_Buscar_FecCon = gf_FormatoFecha(CStr(g_rst_Listas!HIPGAR_FECCON))
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modtac_gf_Buscar_SecCre(ByVal p_NumSol As String) As String
   modtac_gf_Buscar_SecCre = ""
   
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY EVACRE_NUMSOL ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      If Not IsNull(g_rst_Listas!SEGFECCRE) Then
            modtac_gf_Buscar_SecCre = Trim(g_rst_Listas!SEGUSUACT)
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

'=======================================================================================================
Public Sub modtac_gs_Carga_ProOcu(p_Combo As ComboBox, p_Arregl() As modtac_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CTB_PRFMSB WHERE "
   g_str_Parame = g_str_Parame & "PRFMSB_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PRFMSB_DESCRI ASC"
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PrfMsb_Descri)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PrfMsb_CodPsb)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PrfMsb_CodPsb)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PrfMsb_Descri)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function modtac_gs_Cadena_ExtSal(ByVal p_Saldos As Double) As String
   Dim r_int_Count As Integer
      
   p_Saldos = Trim(CStr(p_Saldos))
   r_int_Count = 1
   
   Do While Len(Trim(Mid(p_Saldos, r_int_Count, 1))) > 0
         
      If Trim(Mid(p_Saldos, r_int_Count, 1)) <> "." Then
         modtac_gs_Cadena_ExtSal = modtac_gs_Cadena_ExtSal + Mid(p_Saldos, r_int_Count, 1)
         r_int_Count = r_int_Count + 1
      Else
         If Mid(Right(Trim(p_Saldos), 2), 1, 1) <> "." Then
            modtac_gs_Cadena_ExtSal = modtac_gs_Cadena_ExtSal + Right(Trim(p_Saldos), 2)
            Exit Function
         Else
            modtac_gs_Cadena_ExtSal = modtac_gs_Cadena_ExtSal + Right(Trim(p_Saldos), 1) + "0"
            Exit Function
         End If
      End If
   Loop
   
   modtac_gs_Cadena_ExtSal = modtac_gs_Cadena_ExtSal + "00"
End Function

Public Function modtac_gs_Genera_NumCar(ByVal p_Saldos As Double, ByVal p_CanDig As Integer) As String
   Dim r_str_Cadena As String
      
   For p_CanDig = 1 To p_CanDig Step 1
      r_str_Cadena = r_str_Cadena & "0"
   Next
  
   modtac_gs_Genera_NumCar = Format(p_Saldos, r_str_Cadena)
End Function

