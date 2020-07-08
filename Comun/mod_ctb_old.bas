Attribute VB_Name = "modctb"
Option Explicit
Public Type modctb_g_tpo_Produc
   Produc_Codigo     As String
   Produc_VctCuo     As Integer
   Produc_VctCre     As Integer
End Type
Public modctb_g_arr_Produc()     As modctb_g_tpo_Produc

Public Type modctb_g_tpo_GasCob
   GasCob_DiaIni     As Integer
   GasCob_DiaFin     As Integer
   GasCob_Import     As Double
End Type
Public modctb_g_tpo_GasDol()     As modctb_g_tpo_GasCob
Public modctb_g_tpo_GasVac()     As modctb_g_tpo_GasCob

Public modctb_g_rst_Tablas       As ADODB.Recordset
Public modctb_g_rst_Genera       As ADODB.Recordset
Public modctb_g_rst_Princi       As ADODB.Recordset
Public modctb_g_rst_Accion       As ADODB.Recordset
Public modctb_g_arr_Tasas()      As moddat_tpo_Genera

Public Sub modctb_CBRP001(ByVal p_FecPro As String)
   'Calculo de Morosidad de Clientes
   
   Dim r_str_CodPrd     As String
   Dim r_str_NumOpe     As String
   Dim r_int_TipMon     As Integer
   Dim r_int_MonDes     As Integer
   Dim r_int_MonViv     As Integer
   Dim r_int_MonOtr     As Integer
   Dim r_int_NumCuo     As Integer
   Dim r_int_DiaTra     As Integer
   Dim r_dbl_TCaMpr     As Double
   Dim r_dbl_TCaDes     As Double
   Dim r_dbl_TCaViv     As Double
   Dim r_dbl_TCaOtr     As Double
   Dim r_dbl_TasInt     As Double
   Dim r_dbl_TasMor     As Double
   Dim r_dbl_TasCom     As Double
   Dim r_dbl_TasInt_Men As Double
   Dim r_dbl_TasMor_Men As Double
   Dim r_dbl_TasCom_Men As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_CapPag     As Double
   Dim r_dbl_CuoPag     As Double
   Dim r_dbl_CuoMor     As Double
   Dim r_dbl_CuoCom     As Double
   Dim r_dbl_GasCob     As Double
   
   
   'Verificar existencia de Tipos de Cambio
   
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_PRXVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 OR "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 3  "
   
   If Not gf_EjecutaSQL(g_str_Parame, modctb_g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If modctb_g_rst_Princi.BOF And modctb_g_rst_Princi.EOF Then
      modctb_g_rst_Princi.Close
      Set modctb_g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   modctb_g_rst_Princi.MoveFirst
   Do While Not modctb_g_rst_Princi.EOF
      r_str_CodPrd = Trim(modctb_g_rst_Princi!HIPMAE_CODPRD)
      r_str_NumOpe = Trim(modctb_g_rst_Princi!HIPMAE_NUMOPE)
      r_int_TipMon = modctb_g_rst_Princi!HIPMAE_MONEDA
      r_int_MonDes = modctb_g_rst_Princi!HIPMAE_MONDES
      r_int_MonViv = modctb_g_rst_Princi!HIPMAE_MONVIV
      r_int_MonOtr = modctb_g_rst_Princi!HIPMAE_OTRMON
      
      r_dbl_TCaMpr = moddat_gf_Obtiene_TipCam(1, r_int_TipMon)
      
      r_dbl_TCaDes = moddat_gf_Obtiene_TipCam(1, r_int_MonDes)
      r_dbl_TCaViv = moddat_gf_Obtiene_TipCam(1, r_int_MonViv)
      r_dbl_TCaOtr = moddat_gf_Obtiene_TipCam(1, r_int_MonOtr)
      
      'Obteniendo Tasas
      r_dbl_TasInt = modctb_g_rst_Princi!HIPMAE_TASINT
      r_dbl_TasInt_Men = ff_IntMen(r_dbl_TasInt)
      
      If moddat_gf_Consulta_ParPrd(modctb_g_arr_Tasas(), modctb_g_rst_Princi!HIPMAE_CODPRD, "003", CStr(r_int_TipMon) & "01") Then
         r_dbl_TasMor = modctb_g_arr_Tasas(1).Genera_Cantid
      End If
      r_dbl_TasMor_Men = ff_IntMen(r_dbl_TasMor)
      
      If moddat_gf_Consulta_ParPrd(modctb_g_arr_Tasas(), modctb_g_rst_Princi!HIPMAE_CODPRD, "003", CStr(r_int_TipMon) & "02") Then
         r_dbl_TasCom = modctb_g_arr_Tasas(1).Genera_Cantid
      End If
      r_dbl_TasCom_Men = ff_IntMen(r_dbl_TasCom)
   
      'Filtrando Cuotas
      g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & r_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2  "
   
      If Not gf_EjecutaSQL(g_str_Parame, modctb_g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (modctb_g_rst_Genera.BOF And modctb_g_rst_Genera.EOF) Then
         modctb_g_rst_Genera.MoveFirst
      
         Do While Not modctb_g_rst_Genera.EOF
            r_int_NumCuo = modctb_g_rst_Genera!HIPCUO_NUMCUO
            
            'Obteniendo Días de Mora
            r_int_DiaTra = CInt(Date - CDate(gf_FormatoFecha(modctb_g_rst_Genera!HIPCUO_FECVCT)))
            
            'Calculando Seguro de Desgravamen
            r_dbl_SegDes = modctb_g_rst_Genera!HIPCUO_DESORG - modctb_g_rst_Genera!HIPCUO_DESPAG
            
            If r_int_TipMon <> r_int_MonDes And r_int_MonDes <> 1 Then
               r_dbl_SegDes = r_dbl_SegDes * r_dbl_TCaDes / r_dbl_TCaMpr
            ElseIf r_int_MonDes = 1 Then
               r_dbl_SegDes = r_dbl_SegDes / r_dbl_TCaMpr
            End If
            
            If r_dbl_SegDes > 0 Then
               r_dbl_SegDes = CDbl(Format(r_dbl_SegDes, "###,###,##0.00"))
            Else
               r_dbl_SegDes = 0
            End If
            
            'Calculando Seguro de Vivienda
            r_dbl_SegViv = modctb_g_rst_Genera!HIPCUO_VIVORG - modctb_g_rst_Genera!HIPCUO_VIVPAG
            If r_int_TipMon <> r_int_MonViv And r_int_MonViv <> 1 Then
               r_dbl_SegViv = r_dbl_SegViv * r_dbl_TCaViv / r_dbl_TCaMpr
            ElseIf r_int_MonViv = 1 Then
               r_dbl_SegViv = r_dbl_SegViv / r_dbl_TCaMpr
            End If
            
            If r_dbl_SegViv > 0 Then
               r_dbl_SegViv = CDbl(Format(r_dbl_SegViv, "###,###,##0.00"))
            Else
               r_dbl_SegViv = 0
            End If
            
            'Calculando Otros Cargos
            r_dbl_OtrCar = modctb_g_rst_Genera!HIPCUO_OTRORG - modctb_g_rst_Genera!HIPCUO_OTRPAG
            If r_int_TipMon <> r_int_MonOtr And r_int_MonOtr <> 1 Then
               r_dbl_OtrCar = r_dbl_OtrCar * r_dbl_TCaOtr / r_dbl_TCaMpr
            ElseIf r_int_MonOtr = 1 Then
               r_dbl_OtrCar = r_dbl_OtrCar / r_dbl_TCaMpr
            End If
            
            If r_dbl_OtrCar > 0 Then
               r_dbl_OtrCar = CDbl(Format(r_dbl_OtrCar, "###,###,##0.00"))
            Else
               r_dbl_OtrCar = 0
            End If
            
            'Obteniendo Capital Pendiente de Pago (Sólo de Cuota)
            r_dbl_CapPag = modctb_g_rst_Genera!HIPCUO_CAPITA + modctb_g_rst_Genera!HIPCUO_CAPBBP - modctb_g_rst_Genera!HIPCUO_CAPPAG
            
            'Obteniendo Importe de Cuota Pendiente de Pago
            r_dbl_CuoPag = modctb_g_rst_Genera!HIPCUO_CAPITA + modctb_g_rst_Genera!HIPCUO_CAPBBP - modctb_g_rst_Genera!HIPCUO_CAPPAG
            r_dbl_CuoPag = r_dbl_CuoPag + modctb_g_rst_Genera!HIPCUO_INTERE + modctb_g_rst_Genera!HIPCUO_INTBBP - modctb_g_rst_Genera!HIPCUO_INTPAG
            r_dbl_CuoPag = r_dbl_CuoPag + r_dbl_SegDes
            r_dbl_CuoPag = r_dbl_CuoPag + r_dbl_SegViv
            r_dbl_CuoPag = r_dbl_CuoPag + r_dbl_OtrCar
         
            'Calcular Interés Moratorio
            r_dbl_CuoMor = gf_Interes(r_dbl_TasMor_Men, r_int_DiaTra, r_dbl_CapPag)
            r_dbl_CuoMor = CDbl(Format(r_dbl_CuoMor, "###,###,##0.00"))
            
            'Calcular Interés Compensatorio
            r_dbl_CuoCom = gf_Interes(r_dbl_TasCom_Men, r_int_DiaTra, r_dbl_CuoPag)
            r_dbl_CuoCom = CDbl(Format(r_dbl_CuoCom, "###,###,##0.00"))
            
            'Calcular Gastos de Cobranzas
            r_dbl_GasCob = modctb_gf_GasCob(r_str_CodPrd, r_int_TipMon, r_int_DiaTra)
            r_dbl_GasCob = CDbl(Format(r_dbl_GasCob, "###,###,##0.00"))
            
            
            'Grabar Datos por cuota
            moddat_g_int_FlgGOK = False
            moddat_g_int_CntErr = 0
   
            Do While moddat_g_int_FlgGOK = False
               g_str_Parame = "USP_CRE_HIPCUO_CALMOR ("
               g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "', "
               g_str_Parame = g_str_Parame & CStr(r_int_NumCuo) & ", "
               g_str_Parame = g_str_Parame & CStr(r_dbl_CuoCom) & ", "
               g_str_Parame = g_str_Parame & CStr(r_dbl_CuoMor) & ", "
               g_str_Parame = g_str_Parame & CStr(r_dbl_GasCob) & ", "
         
               'Datos de Auditoria
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
               g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
               g_str_Parame = g_str_Parame & "1)"
                  
               If Not gf_EjecutaSQL(g_str_Parame, modctb_g_rst_Accion, 2) Then
                  moddat_g_int_CntErr = moddat_g_int_CntErr + 1
               Else
                  moddat_g_int_FlgGOK = True
               End If
         
               If moddat_g_int_CntErr = 6 Then
                  'Grabar en alguna parte LOG de error
               End If
               
               DoEvents
            Loop
         
            modctb_g_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      modctb_g_rst_Genera.Close
      Set modctb_g_rst_Genera = Nothing
   
      modctb_g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   modctb_g_rst_Princi.Close
   Set modctb_g_rst_Princi = Nothing
End Sub

Public Sub modctb_gf_Carga_Produc(p_Arregl() As modctb_g_tpo_Produc)
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PRODUC WHERE "
   g_str_Parame = g_str_Parame & "PRODUC_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PRODUC_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, modctb_g_rst_Tablas, 3) Then
       Exit Sub
   End If
   
   If modctb_g_rst_Tablas.BOF And modctb_g_rst_Tablas.EOF Then
     modctb_g_rst_Tablas.Close
     Set modctb_g_rst_Tablas = Nothing
     
     Exit Sub
   End If
   
   modctb_g_rst_Tablas.MoveFirst
   Do While Not modctb_g_rst_Tablas.EOF
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Produc_Codigo = Trim$(modctb_g_rst_Tablas!Produc_Codigo)
      p_Arregl(UBound(p_Arregl)).Produc_VctCuo = modctb_g_rst_Tablas!Produc_VctCuo
      p_Arregl(UBound(p_Arregl)).Produc_VctCre = modctb_g_rst_Tablas!Produc_VctCre
         
      modctb_g_rst_Tablas.MoveNext
   Loop
   
   modctb_g_rst_Tablas.Close
   Set modctb_g_rst_Tablas = Nothing
End Sub

Public Function modctb_gf_GasCob(ByVal p_CodPrd As String, ByVal p_TipMon As Integer, ByVal p_DiaTra As Integer) As Double
   modctb_gf_GasCob = 0

   g_str_Parame = "SELECT * FROM OPE_GASCOB WHERE "
   g_str_Parame = g_str_Parame & "GASCOB_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "GASCOB_TIPMON = " & CStr(p_TipMon) & " AND "
   g_str_Parame = g_str_Parame & "GASCOB_DIAINI <= " & CStr(p_DiaTra) & " AND "
   g_str_Parame = g_str_Parame & "GASCOB_DIAFIN >= " & CStr(p_DiaTra) & " "

   If Not gf_EjecutaSQL(g_str_Parame, modctb_g_rst_Tablas, 3) Then
       Exit Function
   End If
   
   If modctb_g_rst_Tablas.BOF And modctb_g_rst_Tablas.EOF Then
     modctb_g_rst_Tablas.Close
     Set modctb_g_rst_Tablas = Nothing
     
     Exit Function
   End If
   
   modctb_g_rst_Tablas.MoveFirst
   modctb_gf_GasCob = modctb_g_rst_Tablas!GasCob_Import
   
   modctb_g_rst_Tablas.Close
   Set modctb_g_rst_Tablas = Nothing
End Function


