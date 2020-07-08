Attribute VB_Name = "modopecaj"
Option Explicit

Public Type opecaj_g_MovCaj
   MovCaj_CodBan  As String
   MovCaj_UsuMov  As String
   MovCaj_FecMov  As String
   MovCaj_HorMov  As String
   MovCaj_TipMov  As Integer
   MovCaj_NumOpe  As String
   MovCaj_CodIte  As String
   MovCaj_ImpPag  As Double
   MovCaj_MonPag  As Integer
   MovCaj_TipDoc  As Integer
   MovCaj_NumDoc  As String
   MovCaj_FecDep  As String
   MovCaj_NumCta  As String
   MovCaj_NumCom  As String
   MovCaj_FlgRev  As Integer
   MovCaj_SucOpe  As String
   MovCaj_TipCam  As Double
   MovCaj_ImpMEx  As Double
   MovCaj_ImpMNc  As Double
   MovCaj_ITFPor  As Double
   MovCaj_ITFImp  As Double
   MovCaj_ImpTot  As Double
   MovCaj_TipReg  As Integer
   MovCaj_FecRec  As String
   MovCaj_OfiPag  As String
   MovCaj_ForPag  As String
   MovCaj_CanPag  As String
End Type

Public Type opecaj_g_PagCli
   PagCli_NumCuo  As Integer
   PagCli_Capita  As Double
   PagCli_Intere  As Double
   PagCli_SegDes  As Double
   PagCli_SegViv  As Double
   PagCli_OtrCar  As Double
   PagCli_CapPBP  As Double
   PagCli_IntPBP  As Double
   PagCli_IntCom  As Double
   PagCli_IntMor  As Double
   PagCli_GasCob  As Double
   PagCli_OtrGas  As Double
End Type

Public Type opecaj_g_CuoVct
   CuoVct_NumCuo  As Integer
   CuoVct_FecVct  As String
   CuoVct_TotCuo  As Double
   CuoVct_Situac  As String
End Type

Global opecaj_g_arr_OpeCaj()  As opecaj_g_MovCaj
Global opecaj_g_int_ApeCie    As Integer
Global opecaj_g_str_NumMov    As String
Global opecaj_g_str_UsuMov    As String
Global opecaj_g_str_CodBan    As String
Global opecaj_g_str_CtaBan    As String
Global opecaj_g_str_FecMov    As String
Global opecaj_g_int_FlgAct    As Integer

Public Function opecaj_gf_Inserta_CajMov(ByVal p_CodUsu As String, ByVal p_TipOpe As String, ByVal p_NumOpe As String, ByVal p_CodIte As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, _
                                         ByVal p_CodBan As String, ByVal p_FecDep As String, ByVal p_NumCta As String, ByVal p_NumCom As String, ByVal p_TipMon As Integer, ByVal p_Import As Double, _
                                         ByVal p_FlgRev As Integer, ByVal p_SucOpe As String, ByVal p_TipCam As Double, ByVal p_ImpMEx As Double, ByVal p_ImpMNc As Double, ByVal p_ITFPor As Double, _
                                         ByVal p_ITFImp As Double, ByVal p_ImpTot As Double, ByVal p_OpeRev As Integer, ByVal p_RevNum As String, ByVal p_Operac As String, ByVal p_Numero As Long, _
                                         ByVal p_TipReg As Integer, ByVal p_FecRec As String, ByVal p_OfiPag As String, ByVal p_ForPag As String, ByVal p_CanPag As String, ByVal p_Capita As Double, _
                                         ByVal p_intere As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_OtrCar As Double, ByVal p_CapBBP As Double, ByVal p_IntBBP As Double, _
                                         ByVal p_IntMor As Double, ByVal p_IntCom As Double, ByVal p_GasCob As Double, ByVal p_OtrGas As Double, ByVal p_RevSuc As String, ByVal p_RevFec As String) As Integer
   
   opecaj_gf_Inserta_CajMov = False
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_OPE_CAJMOV ("
   
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & p_TipOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_Numero) & ", "
      g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodIte & "', "
      g_str_Parame = g_str_Parame & CStr(p_Import) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & p_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodBan & "', "
      g_str_Parame = g_str_Parame & p_FecDep & ", "
      g_str_Parame = g_str_Parame & "'" & p_NumCta & "', "
      g_str_Parame = g_str_Parame & "'" & p_NumCom & "', "
      g_str_Parame = g_str_Parame & CStr(p_FlgRev) & ", "
      g_str_Parame = g_str_Parame & "'" & p_SucOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipCam) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ImpMEx) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ImpMNc) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ITFPor) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ITFImp) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ImpTot) & ", "
      g_str_Parame = g_str_Parame & "'" & p_Operac & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipReg) & ", "
      g_str_Parame = g_str_Parame & p_FecRec & ", "
      g_str_Parame = g_str_Parame & "'" & p_OfiPag & "', "
      g_str_Parame = g_str_Parame & "'" & p_ForPag & "', "
      g_str_Parame = g_str_Parame & "'" & p_CanPag & "', "
      g_str_Parame = g_str_Parame & CStr(p_OpeRev) & ", "
      g_str_Parame = g_str_Parame & p_RevNum & ", "
      g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
      g_str_Parame = g_str_Parame & CStr(p_intere) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegDes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegViv) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrCar) & ", "
      g_str_Parame = g_str_Parame & CStr(p_CapBBP) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntBBP) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntMor) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntCom) & ", "
      g_str_Parame = g_str_Parame & CStr(p_GasCob) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
      g_str_Parame = g_str_Parame & "'" & p_RevSuc & "', "
      g_str_Parame = g_str_Parame & p_RevFec & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_OPE_CAJMOV. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   opecaj_gf_Inserta_CajMov = True
End Function

Public Function opecaj_gf_Consulta_ITF(ByVal p_FecOpe As String, ByVal p_TipDat As Integer) As Double
   opecaj_gf_Consulta_ITF = 0
      
   g_str_Parame = "SELECT * FROM OPE_TABITF WHERE "
   g_str_Parame = g_str_Parame & "TABITF_FECINI <= " & p_FecOpe & " AND "
   g_str_Parame = g_str_Parame & "TABITF_FECFIN >= " & p_FecOpe & ""

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If



   g_rst_Genera.MoveFirst
   If p_TipDat = 1 Then
      opecaj_gf_Consulta_ITF = g_rst_Genera!TABITF_PORCEN
   Else
      opecaj_gf_Consulta_ITF = g_rst_Genera!TABITF_IMPORT
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function opecaj_gf_Pago_GasAdm(ByVal p_NumSol As String, ByVal p_CodGas As Integer, ByVal p_TipMon As Integer, ByVal p_Import As Double, ByVal p_PorITF As Double, ByVal p_FecPag As String, ByVal p_Operac As String) As Integer
   opecaj_gf_Pago_GasAdm = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_GASADM_PAGO ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(p_CodGas) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Import) & ", "
      g_str_Parame = g_str_Parame & gf_NueImp_Numero(gf_Truncar_Numero(p_Import * (p_PorITF / 100), 2)) & ", "
      g_str_Parame = g_str_Parame & p_FecPag & ", "
      g_str_Parame = g_str_Parame & "'" & p_Operac & "', "
      g_str_Parame = g_str_Parame & "1, "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_GASADM_PAGO. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   opecaj_gf_Pago_GasAdm = True
End Function

Public Function opecaj_gf_Genera_NumMov() As Long
   Dim r_lng_NumMov     As Long
   Dim r_str_FecSis     As String
   
   opecaj_gf_Genera_NumMov = 0
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   r_str_FecSis = Format(CDate(moddat_g_str_FecSis), "yyyymmdd")
   r_str_FecSis = Left(r_str_FecSis, 4) & "0000"
   
   g_str_Parame = "SELECT * FROM OPE_CAJFOL WHERE "
   g_str_Parame = g_str_Parame & "CAJFOL_CODSUC = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "CAJFOL_FECDIA = " & r_str_FecSis

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumMov = 1
   Else
      r_lng_NumMov = g_rst_Genera!CAJFOL_NUMERO + 1
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_OPE_CAJFOL ("
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & r_str_FecSis & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumMov) & ", "
      
      If r_lng_NumMov = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_OPE_CAJFOL. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   opecaj_gf_Genera_NumMov = r_lng_NumMov
End Function

Public Function opecaj_gf_Pago_Cuotas(ByVal p_NumOpe As String, ByVal p_NumCuo As Integer, ByVal p_FecPag As String, ByVal p_PagMpr As Double, ByVal p_Capita As Double, ByVal p_intere As Double, ByVal p_DesOrg As Double, ByVal p_VivOrg As Double, ByVal p_OtrOrg As Double, ByVal p_IntCom As Double, ByVal p_IntMor As Double, ByVal p_GasCob As Double, ByVal p_OtrGas As Double, ByVal p_TCaDol As Double, ByVal p_TCaMPr As Double, ByVal p_SitCre As Integer, ByVal p_Operac As String, ByVal p_NumMov As Long, ByVal p_SitCuo As Integer, ByVal p_PrxVct As String, ByVal p_CuoPen As Integer, ByVal p_Situac As Integer, ByVal p_SitAnt As Integer, ByVal p_FlgCre As Integer, ByVal p_CapBBP As Double, ByVal p_IntBBP As Double) As Integer
   opecaj_gf_Pago_Cuotas = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPPAG_PAGO ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_NumCuo) & ", "
      g_str_Parame = g_str_Parame & p_FecPag & ", "
      g_str_Parame = g_str_Parame & CStr(p_PagMpr) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
      g_str_Parame = g_str_Parame & CStr(p_intere) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DesOrg) & ", "
      g_str_Parame = g_str_Parame & CStr(p_VivOrg) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrOrg) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntCom) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntMor) & ", "
      g_str_Parame = g_str_Parame & CStr(p_GasCob) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TCaDol) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TCaMPr) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_Operac) & "', "
      g_str_Parame = g_str_Parame & CStr(p_NumMov) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SitCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SitCre) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SitAnt) & ", "
      g_str_Parame = g_str_Parame & p_PrxVct & ", "
      g_str_Parame = g_str_Parame & CStr(p_CuoPen) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Situac) & ", "
      g_str_Parame = g_str_Parame & CStr(p_FlgCre) & ", "
      g_str_Parame = g_str_Parame & CStr(p_CapBBP) & ", "
      g_str_Parame = g_str_Parame & CStr(p_IntBBP) & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_HIPPAG_PAGO. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   opecaj_gf_Pago_Cuotas = True
End Function

Public Function opecaj_gf_ActualizaSaldo(ByVal p_CodBan As String, ByVal p_TipMon As Integer, ByVal p_Import As Double) As Integer
   opecaj_gf_ActualizaSaldo = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_OPE_CAJDIA ("
   
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodBan & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(p_Import) & ", "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "1, "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_OPE_CAJDIA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   opecaj_gf_ActualizaSaldo = True
End Function

Public Sub opecaj_gs_Consulta_CajMov(p_Arregl() As opecaj_g_MovCaj, ByVal p_CodUsu As String, ByVal p_CodBan As String, ByVal p_FecMov As String, ByVal p_NumMov As String)
   'Nivelar en programas para poder quitar p_CodUsu
   Dim r_str_HorMov  As String
   
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & p_FecMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & p_NumMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & p_CodBan & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   p_Arregl(UBound(p_Arregl)).MovCaj_CodBan = g_rst_Princi!CAJMOV_CODBAN
   p_Arregl(UBound(p_Arregl)).MovCaj_UsuMov = g_rst_Princi!CAJMOV_USUMOV
   p_Arregl(UBound(p_Arregl)).MovCaj_FecMov = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
   p_Arregl(UBound(p_Arregl)).MovCaj_HorMov = g_rst_Princi!CAJMOV_HORMOV
   p_Arregl(UBound(p_Arregl)).MovCaj_TipMov = g_rst_Princi!CAJMOV_TIPMOV
   p_Arregl(UBound(p_Arregl)).MovCaj_NumOpe = Trim(g_rst_Princi!CAJMOV_NUMOPE & "")
   p_Arregl(UBound(p_Arregl)).MovCaj_CodIte = Trim(g_rst_Princi!CAJMOV_CODITE & "")
   p_Arregl(UBound(p_Arregl)).MovCaj_ImpPag = g_rst_Princi!CAJMOV_IMPPAG
   p_Arregl(UBound(p_Arregl)).MovCaj_MonPag = g_rst_Princi!CAJMOV_MONPAG
   p_Arregl(UBound(p_Arregl)).MovCaj_TipDoc = g_rst_Princi!CAJMOV_TIPDOC
   p_Arregl(UBound(p_Arregl)).MovCaj_NumDoc = Trim(g_rst_Princi!CAJMOV_NUMDOC & "")
   
   If g_rst_Princi!CAJMOV_FECDEP > 0 Then
      p_Arregl(UBound(p_Arregl)).MovCaj_FecDep = Right(CStr(g_rst_Princi!CAJMOV_FECDEP), 2) & "/" & Mid(CStr(g_rst_Princi!CAJMOV_FECDEP), 5, 2) & "/" & Left(CStr(g_rst_Princi!CAJMOV_FECDEP), 4)
   Else
      p_Arregl(UBound(p_Arregl)).MovCaj_FecDep = "0"
   End If
   
   p_Arregl(UBound(p_Arregl)).MovCaj_NumCta = Trim(g_rst_Princi!CAJMOV_NUMCTA & "")
   p_Arregl(UBound(p_Arregl)).MovCaj_NumCom = Trim(g_rst_Princi!CAJMOV_NUMCOM & "")
   p_Arregl(UBound(p_Arregl)).MovCaj_FlgRev = g_rst_Princi!CAJMOV_FLGREV
   p_Arregl(UBound(p_Arregl)).MovCaj_SucOpe = Trim(g_rst_Princi!CAJMOV_SUCOPE & "")
   p_Arregl(UBound(p_Arregl)).MovCaj_TipCam = g_rst_Princi!CAJMOV_TIPCAM
   p_Arregl(UBound(p_Arregl)).MovCaj_ImpMEx = g_rst_Princi!CAJMOV_IMPMEX
   p_Arregl(UBound(p_Arregl)).MovCaj_ImpMNc = g_rst_Princi!CAJMOV_IMPMNC
   p_Arregl(UBound(p_Arregl)).MovCaj_ITFPor = g_rst_Princi!CAJMOV_ITFPOR
   p_Arregl(UBound(p_Arregl)).MovCaj_ITFImp = g_rst_Princi!CAJMOV_ITFIMP
   p_Arregl(UBound(p_Arregl)).MovCaj_ImpTot = g_rst_Princi!CAJMOV_IMPTOT

   If g_rst_Princi!CAJMOV_TIPREG > 0 Then
      p_Arregl(UBound(p_Arregl)).MovCaj_TipReg = g_rst_Princi!CAJMOV_TIPREG
      p_Arregl(UBound(p_Arregl)).MovCaj_FecRec = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECREC))
      p_Arregl(UBound(p_Arregl)).MovCaj_OfiPag = Trim(g_rst_Princi!CAJMOV_OFIPAG & "")
      p_Arregl(UBound(p_Arregl)).MovCaj_ForPag = Trim(g_rst_Princi!CAJMOV_FORPAG & "")
      p_Arregl(UBound(p_Arregl)).MovCaj_CanPag = Trim(g_rst_Princi!CAJMOV_CANPAG & "")
   Else
      p_Arregl(UBound(p_Arregl)).MovCaj_TipReg = 0
      p_Arregl(UBound(p_Arregl)).MovCaj_FecRec = "0"
      p_Arregl(UBound(p_Arregl)).MovCaj_OfiPag = ""
      p_Arregl(UBound(p_Arregl)).MovCaj_ForPag = ""
      p_Arregl(UBound(p_Arregl)).MovCaj_CanPag = ""
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub opecaj_gs_Imp_GasAdm_Ban(ByVal p_FecMov As String, ByVal p_NumMov As String)
   Dim r_int_MonPre     As Integer
   Dim r_dbl_NetPag     As Double
   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)
   
   'Obteniendo Información del Movimiento de Pago
   g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & p_FecMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & p_NumMov & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Call gs_LinImp(Space(119) & "Fecha: " & gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV)))
   Call gs_LinImp(Space(119) & "Hora:  " & Space(2) & gf_FormatoHora(Format(g_rst_Princi!CAJMOV_HORMOV, "000000")))
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "RUC: 20511904162" & Space(29) & "PAGO DE GASTOS DE CIERRE (EN BANCO)")
   Call gs_LinImp(Space(50) & "-----------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp("Comprobante Nro.      : " & Format(p_NumMov, "00000") & Space(29) & "Nro. de Solicitud     : " & Mid(g_rst_Princi!CAJMOV_NUMOPE, 1, 3) & "-" & Mid(g_rst_Princi!CAJMOV_NUMOPE, 4, 3) & "-" & Mid(g_rst_Princi!CAJMOV_NUMOPE, 7, 2) & "-" & Mid(g_rst_Princi!CAJMOV_NUMOPE, 9, 4))
   Call gs_LinImp("Cliente               : " & CStr(g_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(g_rst_Princi!CAJMOV_NUMDOC) & " / " & moddat_gf_Buscar_NomCli(g_rst_Princi!CAJMOV_TIPDOC, Trim(g_rst_Princi!CAJMOV_NUMDOC)))
   Call gs_LinImp("Fecha de Pago         : " & gf_FormatoFecha(g_rst_Princi!CAJMOV_FECDEP) & " / " & Mid(Trim(g_rst_Princi!CAJMOV_NUMCOM) & Space(15), 1, 15) & Space(6) & "Cuenta de Pago        : " & Trim(g_rst_Princi!CAJMOV_NUMCTA) & " / " & moddat_gf_Consulta_ParDes("505", g_rst_Princi!CAJMOV_CODBAN))
   Call gs_LinImp("Moneda                : " & moddat_gf_Consulta_ParDes("204", g_rst_Princi!CAJMOV_MONPAG))
   Call gs_LinImp("Importe Pagado        : " & gf_FormatoNumero(g_rst_Princi!CAJMOV_IMPTOT, 15))
   Call gs_LinImp("Importe ITF           : " & gf_FormatoNumero(g_rst_Princi!CAJMOV_ITFIMP, 15))
   Call gs_LinImp("Importe Neto Pago     : " & gf_FormatoNumero(g_rst_Princi!CAJMOV_IMPPAG, 15))
   
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(String(30, "-"))
   Call gs_LinImp(Space(5) & "DPTO DE OPERACIONES")
   Call gs_LinImp(Space(5) & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm"))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub opecaj_gs_Imp_CuoHip_Ban(ByVal p_NumOpe As String, ByVal p_FecMov As String, ByVal p_NumMov As String)
   Dim r_int_MonPre     As Integer
   Dim r_str_Linea      As String
   Dim r_dbl_NetPag     As Double
   Dim r_dbl_TotCuo     As Double
   Dim r_int_Contad     As Integer
   Dim r_int_LinBla     As Integer
   
   'Obteniendo Información de la Operación
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   r_int_MonPre = g_rst_Princi!HIPMAE_MONEDA
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Información del Movimiento de Pago
   g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & p_FecMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & p_NumMov & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If moddat_g_int_TipMon = 3 Then
      r_dbl_NetPag = g_rst_Princi!CAJMOV_IMPPAG / g_rst_Princi!CAJMOV_TIPCAM
   Else
      r_dbl_NetPag = g_rst_Princi!CAJMOV_IMPPAG
   End If
   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)
   
   Call gs_LinImp(Space(119) & "Fecha: " & gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV)))
   Call gs_LinImp(Space(119) & "Hora:  " & Space(2) & gf_FormatoHora(Format(g_rst_Princi!CAJMOV_HORMOV, "000000")))
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "RUC: 20511904162" & Space(25) & "PAGO DE CUOTAS CREDITO HIPOTECARIO (EN BANCO)")
   Call gs_LinImp(Space(46) & "---------------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp("Comprobante Nro.      : " & Format(p_NumMov, "00000") & Space(29) & "Nro. de Operación     : " & Mid(p_NumOpe, 1, 3) & "-" & Mid(p_NumOpe, 4, 2) & "-" & Mid(p_NumOpe, 6, 5))
   Call gs_LinImp("Cliente               : " & CStr(g_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(g_rst_Princi!CAJMOV_NUMDOC) & " / " & moddat_gf_Buscar_NomCli(g_rst_Princi!CAJMOV_TIPDOC, Trim(g_rst_Princi!CAJMOV_NUMDOC)))
   Call gs_LinImp("Fecha de Pago         : " & gf_FormatoFecha(g_rst_Princi!CAJMOV_FECDEP) & " / " & Mid(Trim(g_rst_Princi!CAJMOV_NUMCOM) & Space(15), 1, 15) & Space(6) & "Cuenta de Pago        : " & Trim(g_rst_Princi!CAJMOV_NUMCTA) & " / " & moddat_gf_Consulta_ParDes("505", g_rst_Princi!CAJMOV_CODBAN))
   Call gs_LinImp("Moneda                : " & moddat_gf_Consulta_ParDes("204", g_rst_Princi!CAJMOV_MONPAG))
   Call gs_LinImp("Importe Pagado        : " & gf_FormatoNumero(g_rst_Princi!CAJMOV_IMPTOT, 15))
   Call gs_LinImp("Importe ITF           : " & gf_FormatoNumero(g_rst_Princi!CAJMOV_ITFIMP, 15))
   Call gs_LinImp("Importe Neto Pago     : " & gf_FormatoNumero(g_rst_Princi!CAJMOV_IMPPAG, 15))
   
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(53) & "DETALLE DE CUOTAS AMORTIZADAS")
   Call gs_LinImp(Space(13) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Cuota" & Space(2)
   r_str_Linea = r_str_Linea & "Capital" & Space(2)
   r_str_Linea = r_str_Linea & "Interes" & Space(2)
   r_str_Linea = r_str_Linea & "S. Desg." & Space(2)
   r_str_Linea = r_str_Linea & "S. Inm. " & Space(2)
   r_str_Linea = r_str_Linea & " Portes " & Space(2)
   r_str_Linea = r_str_Linea & "Int.Comp." & Space(2)
   r_str_Linea = r_str_Linea & "Int. Mor." & Space(2)
   r_str_Linea = r_str_Linea & "Gto. Cob." & Space(2)
   r_str_Linea = r_str_Linea & "Ot. Gtos." & Space(2)
   r_str_Linea = r_str_Linea & "Total Cuota" & Space(2)
   
   Call gs_LinImp(Space(13) & r_str_Linea)
   Call gs_LinImp(Space(13) & String(110, "-"))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_int_Contad = 0
   
   g_str_Parame = "SELECT * FROM CRE_HIPPAG WHERE "
   g_str_Parame = g_str_Parame & "HIPPAG_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPPAG_NUMMOV = " & p_NumMov & " AND "
   g_str_Parame = g_str_Parame & "HIPPAG_FECMOV = " & p_FecMov & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPPAG_NUMCUO DESC, HIPPAG_FECPAG DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         r_dbl_TotCuo = 0
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_CAPITA, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_INTERE, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_DESORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_VIVORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_OTRORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_INTCOM, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_INTMOR, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_GASCOB, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(g_rst_Princi!HIPPAG_OTRGAS, "###,###,##0.00"))

         r_str_Linea = ""
         r_str_Linea = r_str_Linea & Space(1) & Format(g_rst_Princi!HIPPAG_NUMCUO, "000") & Space(1)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_CAPITA, 8) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_INTERE, 8) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_DESORG, 8) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_VIVORG, 8) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_OTRORG, 8) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_INTCOM, 9) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_INTMOR, 9) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_GASCOB, 9) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPPAG_OTRGAS, 9) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_TotCuo, 11) & Space(2)
         
         Call gs_LinImp(Space(13) & r_str_Linea)
         
         r_int_Contad = r_int_Contad + 1
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'For r_int_LinBla = r_int_Contad To 7
   '   Call gs_LinImp("")
   'Next r_int_LinBla
   
   Call gs_LinImp(Space(13) & String(110, "-"))
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(String(30, "-"))
   Call gs_LinImp(Space(9) & "OPERACIONES")
   Call gs_LinImp(Space(5) & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm"))
End Sub

Public Sub opecaj_gs_ComPagoCab(ByVal p_CodTer As String, ByVal p_NumPag As Integer, ByVal p_NumIte As Integer, ByVal p_CodSuc As String, ByVal p_NumMov As Long, _
                                ByVal p_FecMov As String, ByVal p_TipMov As Integer, ByVal p_FecPag As String, ByVal p_DocIde As String, ByVal p_NomCli As String, _
                                ByVal p_TipNum As Integer, ByVal p_NumSol As String, ByVal p_NumOpe As String, ByVal p_Moneda As String, ByVal p_SimMon As String, _
                                ByVal p_NomBan As String, ByVal p_NumCta As String, ByVal p_Import As Double, ByVal p_ImpITF As Double, ByVal p_ImpTot As Double, _
                                ByVal p_RevMov As String)

   g_str_Parame = "INSERT INTO RPT_COMPGC ("
   g_str_Parame = g_str_Parame & "COMPGC_CODTER, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMPAG, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMITE, "
   g_str_Parame = g_str_Parame & "COMPGC_CODSUC, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMMOV, "
   g_str_Parame = g_str_Parame & "COMPGC_FECMOV, "
   g_str_Parame = g_str_Parame & "COMPGC_TIPMOV, "
   g_str_Parame = g_str_Parame & "COMPGC_FECPAG, "
   g_str_Parame = g_str_Parame & "COMPGC_DOCIDE, "
   g_str_Parame = g_str_Parame & "COMPGC_NOMCLI, "
   g_str_Parame = g_str_Parame & "COMPGC_TIPNUM, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMSOL, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMOPE, "
   g_str_Parame = g_str_Parame & "COMPGC_MONEDA, "
   g_str_Parame = g_str_Parame & "COMPGC_SIMMON, "
   g_str_Parame = g_str_Parame & "COMPGC_NOMBAN, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMCTA, "
   g_str_Parame = g_str_Parame & "COMPGC_IMPORT, "
   g_str_Parame = g_str_Parame & "COMPGC_IMPITF, "
   g_str_Parame = g_str_Parame & "COMPGC_IMPTOT, "
   g_str_Parame = g_str_Parame & "COMPGC_REVMOV, "
   g_str_Parame = g_str_Parame & "COMPGC_NUMCOM) "
   g_str_Parame = g_str_Parame & "VALUES ( "
   g_str_Parame = g_str_Parame & "'" & p_CodTer & "', "
   g_str_Parame = g_str_Parame & CStr(p_NumPag) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
   g_str_Parame = g_str_Parame & "'" & p_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(p_NumMov) & ", "
   g_str_Parame = g_str_Parame & p_FecMov & ", "
   g_str_Parame = g_str_Parame & CStr(p_TipMov) & ", "
   g_str_Parame = g_str_Parame & p_FecPag & ", "
   g_str_Parame = g_str_Parame & "'" & p_DocIde & "', "
   g_str_Parame = g_str_Parame & "'" & p_NomCli & "', "
   g_str_Parame = g_str_Parame & CStr(p_TipNum) & ", "
   g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
   g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
   g_str_Parame = g_str_Parame & "'" & p_Moneda & "', "
   g_str_Parame = g_str_Parame & "'" & p_SimMon & "', "
   g_str_Parame = g_str_Parame & "'" & p_NomBan & "', "
   g_str_Parame = g_str_Parame & "'" & p_NumCta & "', "
   g_str_Parame = g_str_Parame & CStr(p_Import) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ImpITF) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ImpTot) & ", "
   g_str_Parame = g_str_Parame & "'" & p_RevMov & "', "
   g_str_Parame = g_str_Parame & "'" & p_CodSuc & "-" & Mid(p_FecMov, 3, 2) & Format(p_NumMov, "00000") & "')"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
End Sub

Public Sub opecaj_gs_ComPagoDet(ByVal p_CodTer As String, ByVal p_NumPag As Integer, ByVal p_NumIte As Integer, ByVal p_NumLin As Integer, ByVal p_Descri As String, _
                                ByVal p_Import As Double, ByVal p_NumCu1 As Integer, ByVal p_Impor1 As Double, ByVal p_NumCu2 As Integer, ByVal p_Impor2 As Double, _
                                ByVal p_NumCu3 As Integer, ByVal p_Impor3 As Double, ByVal p_NumCu4 As Integer, ByVal p_Impor4 As Double, ByVal p_NumCu5 As Integer, _
                                ByVal p_Impor5 As Double)

   g_str_Parame = "INSERT INTO RPT_COMPGD ("
   g_str_Parame = g_str_Parame & "COMPGD_CODTER, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMPAG, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMITE, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMLIN, "
   g_str_Parame = g_str_Parame & "COMPGD_DESCRI, "
   g_str_Parame = g_str_Parame & "COMPGD_IMPORT, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMCU1, "
   g_str_Parame = g_str_Parame & "COMPGD_IMPOR1, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMCU2, "
   g_str_Parame = g_str_Parame & "COMPGD_IMPOR2, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMCU3, "
   g_str_Parame = g_str_Parame & "COMPGD_IMPOR3, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMCU4, "
   g_str_Parame = g_str_Parame & "COMPGD_IMPOR4, "
   g_str_Parame = g_str_Parame & "COMPGD_NUMCU5, "
   g_str_Parame = g_str_Parame & "COMPGD_IMPOR5) "
   g_str_Parame = g_str_Parame & "VALUES ( "
   g_str_Parame = g_str_Parame & "'" & p_CodTer & "', "
   g_str_Parame = g_str_Parame & CStr(p_NumPag) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumLin) & ", "
   g_str_Parame = g_str_Parame & "'" & p_Descri & "', "
   g_str_Parame = g_str_Parame & CStr(p_Import) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCu1) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Impor1) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCu2) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Impor2) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCu3) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Impor3) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCu4) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Impor4) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCu5) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Impor5) & ") "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
End Sub

Public Sub opecaj_gs_ComPago(ByVal p_SucMov As String, p_NumMov As String, ByVal p_FecMov As String, ByVal p_NumPag As Integer, ByVal p_NumIte As Integer)
   Dim r_arr_PagCli()   As opecaj_g_PagCli
   Dim r_rst_Genera     As ADODB.Recordset
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_str_DocIde     As String
   Dim r_str_NomCli     As String
   Dim r_str_FecPag     As String
   Dim r_str_SimMon     As String
   Dim r_str_Moneda     As String
   Dim r_str_NumSol     As String
   Dim r_str_NumOpe     As String
   Dim r_str_NomBan     As String
   Dim r_str_NumCta     As String
   Dim r_str_CodPrd     As String
   Dim r_str_CodSub     As String
   Dim r_str_NomGas     As String
   Dim r_str_RevMov     As String
   
   Dim r_int_TipNum     As Integer
   Dim r_int_ConLin     As Integer
   Dim r_int_Contad     As Integer
   
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_CapPBP     As Double
   Dim r_dbl_IntPBP     As Double
   Dim r_dbl_IntCom     As Double
   Dim r_dbl_IntMor     As Double
   Dim r_dbl_GasCob     As Double
   Dim r_dbl_OtrGas     As Double
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & p_SucMov & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & p_NumMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & p_FecMov & " "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      If r_rst_Princi!CAJMOV_FECDEP > 0 Then
         r_str_FecPag = CStr(r_rst_Princi!CAJMOV_FECDEP)
      Else
         r_str_FecPag = CStr(r_rst_Princi!CAJMOV_FECREC)
      End If
      
      r_str_DocIde = moddat_gf_Consulta_ParDes("203", CStr(r_rst_Princi!CAJMOV_TIPDOC)) & " - " & Trim(r_rst_Princi!CAJMOV_NUMDOC)
      r_str_NomCli = moddat_gf_Buscar_NomCli(CStr(r_rst_Princi!CAJMOV_TIPDOC), Trim(r_rst_Princi!CAJMOV_NUMDOC))
      r_str_SimMon = moddat_gf_Consulta_ParDes("229", CStr(r_rst_Princi!CAJMOV_MONPAG))
      r_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(r_rst_Princi!CAJMOV_MONPAG))
      
      r_str_NumSol = ""
      r_str_NumOpe = ""
      r_int_TipNum = 0
      
      If r_rst_Princi!CAJMOV_TIPMOV = 1101 Or r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_str_NumSol = Trim(r_rst_Princi!CAJMOV_NUMOPE & "")
         r_int_TipNum = 1
      Else
         r_str_NumOpe = Trim(r_rst_Princi!CAJMOV_NUMOPE & "")
         r_int_TipNum = 2
      End If
      
      r_str_NomBan = ""
      r_str_NumCta = ""
      
      If Len(Trim(r_rst_Princi!CAJMOV_CODBAN)) > 0 And r_rst_Princi!CAJMOV_CODBAN <> "000000" Then
         r_str_NomBan = moddat_gf_Consulta_ParDes("505", Trim(r_rst_Princi!CAJMOV_CODBAN))
         r_str_NumCta = Trim(r_rst_Princi!CAJMOV_NUMCTA & "")
      End If
   
      If r_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         r_str_RevMov = r_rst_Princi!CAJMOV_REVSUC & "-" & Mid(CStr(r_rst_Princi!CAJMOV_REVFEC), 3, 2) & Format(r_rst_Princi!CAJMOV_REVNUM, "00000")
      End If
   
      Call opecaj_gs_ComPagoCab(modgen_g_str_NombPC, p_NumPag, p_NumIte, p_SucMov, p_NumMov, p_FecMov, r_rst_Princi!CAJMOV_TIPMOV, _
                                r_str_FecPag, r_str_DocIde, r_str_NomCli, r_int_TipNum, r_str_NumSol, r_str_NumOpe, r_str_Moneda, r_str_SimMon, r_str_NomBan, _
                                r_str_NumCta, r_rst_Princi!CAJMOV_IMPPAG, r_rst_Princi!CAJMOV_ITFIMP, r_rst_Princi!CAJMOV_IMPTOT, r_str_RevMov)
      
      Select Case r_rst_Princi!CAJMOV_TIPMOV
         Case "1101"    'Pago de Gastos de Cierre
            g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
            g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & r_str_NumSol & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
               Exit Sub
            End If
            
            r_rst_Genera.MoveFirst
            r_str_CodPrd = r_rst_Genera!SOLMAE_CODPRD
            r_str_CodSub = r_rst_Genera!SOLMAE_CODSUB
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
            
            'Buscar en Tabla de Gastos de Cierre
            g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
            g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & r_str_NumSol & "' AND GASADM_OPERAC <> '99999' "
         
            If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
                Exit Sub
            End If
            
            If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
               r_rst_Genera.MoveFirst
               r_int_ConLin = 1
               
               Do While Not r_rst_Genera.EOF
                  r_str_NomGas = ""
                  
                  If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), r_str_CodPrd, r_str_CodSub, "007", Format(r_rst_Genera!GASADM_CODGAS, "00") & Format(r_rst_Genera!GASADM_TIPMON, "0")) Then
                     r_str_NomGas = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
                  End If
                  
                  Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, r_int_ConLin, r_str_NomGas, r_rst_Genera!GASADM_IMPORT, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                  
                  r_rst_Genera.MoveNext
                  r_int_ConLin = r_int_ConLin + 1
               Loop
            End If
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
            
            For r_int_Contad = r_int_ConLin To 11
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, r_int_Contad, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Next r_int_Contad
            
         Case "1102"
            ReDim r_arr_PagCli(0)
         
            g_str_Parame = "SELECT * FROM CRE_HIPPAG WHERE "
            g_str_Parame = g_str_Parame & "HIPPAG_SUCMOV = '" & p_SucMov & "' AND "
            g_str_Parame = g_str_Parame & "HIPPAG_NUMMOV = '" & p_NumMov & "' AND "
            g_str_Parame = g_str_Parame & "HIPPAG_FECMOV = " & p_FecMov & " ORDER BY HIPPAG_NUMCUO ASC "
         
            If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
                Exit Sub
            End If
            
            If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
               r_rst_Genera.MoveFirst
               
               r_dbl_Capita = 0
               r_dbl_Intere = 0
               r_dbl_SegDes = 0
               r_dbl_SegViv = 0
               r_dbl_OtrCar = 0
               r_dbl_CapPBP = 0
               r_dbl_IntPBP = 0
               r_dbl_IntCom = 0
               r_dbl_IntMor = 0
               r_dbl_GasCob = 0
               r_dbl_OtrGas = 0
               
               Do While Not r_rst_Genera.EOF
                  ReDim Preserve r_arr_PagCli(UBound(r_arr_PagCli) + 1)
                  
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_NumCuo = r_rst_Genera("HIPPAG_NUMCUO")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_Capita = r_rst_Genera("HIPPAG_CAPITA")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_Intere = r_rst_Genera("HIPPAG_INTERE")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_SegDes = r_rst_Genera("HIPPAG_DESORG")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_SegViv = r_rst_Genera("HIPPAG_VIVORG")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_OtrCar = r_rst_Genera("HIPPAG_OTRORG")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_CapPBP = r_rst_Genera("HIPPAG_CAPBBP")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_IntPBP = r_rst_Genera("HIPPAG_INTBBP")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_IntMor = r_rst_Genera("HIPPAG_INTMOR")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_IntCom = r_rst_Genera("HIPPAG_INTCOM")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_GasCob = r_rst_Genera("HIPPAG_GASCOB")
                  r_arr_PagCli(UBound(r_arr_PagCli)).PagCli_OtrGas = r_rst_Genera("HIPPAG_OTRGAS")
               
                  r_dbl_Capita = r_dbl_Capita + r_rst_Genera("HIPPAG_CAPITA")
                  r_dbl_Intere = r_dbl_Intere + r_rst_Genera("HIPPAG_INTERE")
                  r_dbl_SegDes = r_dbl_SegDes + r_rst_Genera("HIPPAG_DESORG")
                  r_dbl_SegViv = r_dbl_SegViv + r_rst_Genera("HIPPAG_VIVORG")
                  r_dbl_OtrCar = r_dbl_OtrCar + r_rst_Genera("HIPPAG_OTRORG")
                  r_dbl_CapPBP = r_dbl_CapPBP + r_rst_Genera("HIPPAG_CAPBBP")
                  r_dbl_IntPBP = r_dbl_IntPBP + r_rst_Genera("HIPPAG_INTBBP")
                  r_dbl_IntMor = r_dbl_IntMor + r_rst_Genera("HIPPAG_INTMOR")
                  r_dbl_IntCom = r_dbl_IntCom + r_rst_Genera("HIPPAG_INTCOM")
                  r_dbl_GasCob = r_dbl_GasCob + r_rst_Genera("HIPPAG_GASCOB")
                  r_dbl_OtrGas = r_dbl_OtrGas + r_rst_Genera("HIPPAG_OTRGAS")
               
                  r_rst_Genera.MoveNext
               Loop
               
               ReDim Preserve r_arr_PagCli(5)
               
               For r_int_Contad = UBound(r_arr_PagCli) + 1 To 5
                  r_arr_PagCli(r_int_Contad).PagCli_NumCuo = 0
                  r_arr_PagCli(r_int_Contad).PagCli_Capita = 0
                  r_arr_PagCli(r_int_Contad).PagCli_Intere = 0
                  r_arr_PagCli(r_int_Contad).PagCli_SegDes = 0
                  r_arr_PagCli(r_int_Contad).PagCli_SegViv = 0
                  r_arr_PagCli(r_int_Contad).PagCli_OtrCar = 0
                  r_arr_PagCli(r_int_Contad).PagCli_CapPBP = 0
                  r_arr_PagCli(r_int_Contad).PagCli_IntPBP = 0
                  r_arr_PagCli(r_int_Contad).PagCli_IntMor = 0
                  r_arr_PagCli(r_int_Contad).PagCli_IntCom = 0
                  r_arr_PagCli(r_int_Contad).PagCli_GasCob = 0
                  r_arr_PagCli(r_int_Contad).PagCli_OtrGas = 0
               Next r_int_Contad
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 1, "CAPITAL", r_dbl_Capita, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_Capita, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_Capita, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_Capita, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_Capita, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_Capita)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 2, "INTERES", r_dbl_Intere, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_Intere, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_Intere, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_Intere, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_Intere, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_Intere)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 3, "SEGURO DESGRAVAMEN", r_dbl_SegDes, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_SegDes, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_SegDes, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_SegDes, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_SegDes, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_SegDes)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 4, "SEGURO INMUEBLE", r_dbl_SegViv, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_SegViv, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_SegViv, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_SegViv, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_SegViv, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_SegViv)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 5, "PORTES", r_dbl_OtrCar, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_OtrCar, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_OtrCar, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_OtrCar, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_OtrCar, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_OtrCar)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 6, "CAPITAL PBP", r_dbl_CapPBP, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_CapPBP, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_CapPBP, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_CapPBP, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_CapPBP, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_CapPBP)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 7, "INTERES PBP", r_dbl_IntPBP, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_IntPBP, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_IntPBP, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_IntPBP, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_IntPBP, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_IntPBP)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 8, "INTERES MORATORIO", r_dbl_IntMor, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_IntMor, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_IntMor, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_IntMor, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_IntMor, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_IntMor)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 9, "INTERES COMPENSATORIO", r_dbl_IntCom, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_IntCom, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_IntCom, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_IntCom, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_IntCom, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_IntCom)
               
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 10, "GASTOS DE COBRANZAS", r_dbl_GasCob, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_GasCob, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_GasCob, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_GasCob, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_GasCob, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_GasCob)

               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 11, "OTROS GASTOS", r_dbl_OtrGas, _
                                         r_arr_PagCli(1).PagCli_NumCuo, r_arr_PagCli(1).PagCli_OtrGas, r_arr_PagCli(2).PagCli_NumCuo, r_arr_PagCli(2).PagCli_OtrGas, _
                                         r_arr_PagCli(3).PagCli_NumCuo, r_arr_PagCli(3).PagCli_OtrGas, r_arr_PagCli(4).PagCli_NumCuo, r_arr_PagCli(4).PagCli_OtrGas, _
                                         r_arr_PagCli(5).PagCli_NumCuo, r_arr_PagCli(5).PagCli_OtrGas)
            End If
         
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
         
         Case "1103"
            Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 1, "DESEMBOLSO", r_rst_Princi!CAJMOV_IMPPAG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            
            For r_int_Contad = 2 To 11
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, r_int_Contad, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Next r_int_Contad
            
         Case "2101"
            Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, 1, "EXTORNO PAGO DE GASTOS DE CIERRE", r_rst_Princi!CAJMOV_IMPPAG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            
            For r_int_Contad = 2 To 11
               Call opecaj_gs_ComPagoDet(modgen_g_str_NombPC, p_NumPag, p_NumIte, r_int_Contad, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Next r_int_Contad
         
      End Select
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Public Sub opecaj_gs_EstCta(ByVal p_NumOpe As String, ByVal p_FecIni As String, ByVal p_FecFin As String, Optional ByVal p_NomRpt As String)
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_rst_HipMae     As ADODB.Recordset
   Dim r_rst_HipCuo     As ADODB.Recordset
   Dim r_rst_Grabar     As ADODB.Recordset
   Dim r_rst_HipPag     As ADODB.Recordset
   
   Dim r_arr_CuoVct()   As opecaj_g_CuoVct
   
   Dim r_str_Direcc     As String
   Dim r_str_Distri     As String
   Dim r_str_DiaPag     As String
   Dim r_int_NumCuo     As Integer
   Dim r_int_NumPag     As Integer
   Dim r_int_NumLin     As Integer
   
   If Len(Trim(p_NomRpt & "")) = 0 Then
      p_NomRpt = "OPE_ESTCTA_01.RPT"
   End If
   
   'Obteniendo datos para Cabecera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_HipMae, 3) Then
       Exit Sub
   End If
   
   r_rst_HipMae.MoveFirst
   
   'Buscando Cuotas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_HipCuo, 3) Then
       Exit Sub
   End If
   
   r_int_NumCuo = 0
   r_str_DiaPag = "0"
   
   If Not (r_rst_HipCuo.BOF And r_rst_HipCuo.EOF) Then
      ReDim r_arr_CuoVct(0)
      r_rst_HipCuo.MoveFirst
      
      Do While Not r_rst_HipCuo.EOF
         If r_rst_HipMae!HIPMAE_SITUAC = 9 Then
            If r_rst_HipCuo!HIPCUO_FECVCT > r_rst_HipMae!HIPMAE_FECCAN Then
               Exit Do
            End If
         End If
      
         ReDim Preserve r_arr_CuoVct(UBound(r_arr_CuoVct) + 1)
         r_str_DiaPag = Right(CStr(r_rst_HipCuo!HIPCUO_FECVCT), 2)
         r_arr_CuoVct(UBound(r_arr_CuoVct)).CuoVct_FecVct = gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT))
         r_arr_CuoVct(UBound(r_arr_CuoVct)).CuoVct_NumCuo = r_rst_HipCuo!HIPCUO_NUMCUO
         
         If CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT))) < date Then
            r_arr_CuoVct(UBound(r_arr_CuoVct)).CuoVct_Situac = "VENCIDA"
         Else
            r_arr_CuoVct(UBound(r_arr_CuoVct)).CuoVct_Situac = "X VENCER"
         End If
         
         r_arr_CuoVct(UBound(r_arr_CuoVct)).CuoVct_TotCuo = r_rst_HipCuo!HIPCUO_CAPITA - r_rst_HipCuo!HIPCUO_CAPPAG + r_rst_HipCuo!HIPCUO_INTERE - r_rst_HipCuo!HIPCUO_INTPAG + r_rst_HipCuo!HIPCUO_DESORG - r_rst_HipCuo!HIPCUO_DESPAG + _
                                                            r_rst_HipCuo!HIPCUO_VIVORG - r_rst_HipCuo!HIPCUO_VIVPAG + r_rst_HipCuo!HIPCUO_OTRORG - r_rst_HipCuo!HIPCUO_OTRPAG + r_rst_HipCuo!HIPCUO_INTCOM - r_rst_HipCuo!HIPCUO_ICOPAG + _
                                                            r_rst_HipCuo!HIPCUO_INTMOR - r_rst_HipCuo!HIPCUO_IMOPAG + r_rst_HipCuo!HIPCUO_GASCOB - r_rst_HipCuo!HIPCUO_GCOPAG + r_rst_HipCuo!HIPCUO_OTRGAS - r_rst_HipCuo!HIPCUO_OTGPAG + _
                                                            r_rst_HipCuo!HIPCUO_CAPBBP - r_rst_HipCuo!HIPCUO_CBPPAG + r_rst_HipCuo!HIPCUO_INTBBP - r_rst_HipCuo!HIPCUO_IBPPAG
      
         r_rst_HipCuo.MoveNext
         
         r_int_NumCuo = r_int_NumCuo + 1
         If r_int_NumCuo = 13 Then
            Exit Do
         End If
         
         DoEvents
      Loop
   End If
         
   r_rst_HipCuo.Close
   Set r_rst_HipCuo = Nothing
         
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(r_rst_HipMae!hipmae_numsol, r_str_Direcc, r_str_Distri)
      
   'Insertando Cabecera
   g_str_Parame = "INSERT INTO RPT_ECTACB ("
   g_str_Parame = g_str_Parame & "ECTACB_CODTER, "
   g_str_Parame = g_str_Parame & "ECTACB_NOMRPT, "
   g_str_Parame = g_str_Parame & "ECTACB_NUMOPE, "
   g_str_Parame = g_str_Parame & "ECTACB_PERINI, "
   g_str_Parame = g_str_Parame & "ECTACB_PERFIN, "
   g_str_Parame = g_str_Parame & "ECTACB_PRODUC, "
   g_str_Parame = g_str_Parame & "ECTACB_CODCLI, "
   g_str_Parame = g_str_Parame & "ECTACB_DOCIDE, "
   g_str_Parame = g_str_Parame & "ECTACB_NOMCLI, "
   g_str_Parame = g_str_Parame & "ECTACB_DIRINM, "
   g_str_Parame = g_str_Parame & "ECTACB_DSTINM, "
   g_str_Parame = g_str_Parame & "ECTACB_FECDES, "
   g_str_Parame = g_str_Parame & "ECTACB_MTOPRE, "
   g_str_Parame = g_str_Parame & "ECTACB_INTCAP, "
   g_str_Parame = g_str_Parame & "ECTACB_PERGRA, "
   g_str_Parame = g_str_Parame & "ECTACB_NUMCUO, "
   g_str_Parame = g_str_Parame & "ECTACB_CUOPAG, "
   g_str_Parame = g_str_Parame & "ECTACB_CUOATR, "
   g_str_Parame = g_str_Parame & "ECTACB_MONEDA, "
   g_str_Parame = g_str_Parame & "ECTACB_SIMMON, "
   g_str_Parame = g_str_Parame & "ECTACB_SALCAP, "
   g_str_Parame = g_str_Parame & "ECTACB_SALCON, "
   g_str_Parame = g_str_Parame & "ECTACB_TASINT, "
   g_str_Parame = g_str_Parame & "ECTACB_COSEFE, "
   g_str_Parame = g_str_Parame & "ECTACB_DIAATR, "
   g_str_Parame = g_str_Parame & "ECTACB_DIAPAG, "
   g_str_Parame = g_str_Parame & "ECTACB_SITUAC, "
   g_str_Parame = g_str_Parame & "ECTACB_FECCAN, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN12, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI12, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF12, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS12) "
   
   g_str_Parame = g_str_Parame & "VALUES ( "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & gf_Formato_NumOpe(p_NumOpe) & "', "
   g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(p_FecIni) & "', "
   g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(p_FecFin) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_HipMae!HIPMAE_CODPRD) & "', "
   g_str_Parame = g_str_Parame & "'" & Format(CStr(r_rst_HipMae!HIPMAE_TDOCLI) & Trim(r_rst_HipMae!HIPMAE_NDOCLI), "000000000000") & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(r_rst_HipMae!HIPMAE_TDOCLI) & "-" & Trim(r_rst_HipMae!HIPMAE_NDOCLI) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_HipMae!HIPMAE_TDOCLI, Trim(r_rst_HipMae!HIPMAE_NDOCLI), 1) & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Direcc & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Distri & "', "
   g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(r_rst_HipMae!HIPMAE_FECDES) & "', "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_MTOPRE) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_INTCAP) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_PERGRA) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_NUMCUO) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_CUOPAG) & ", "
   If IsNull(r_rst_HipMae!HIPMAE_CUOATR) Then
      g_str_Parame = g_str_Parame & "0, "
   Else
      g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_CUOATR) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("204", CStr(r_rst_HipMae!HIPMAE_MONEDA)) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_HipMae!HIPMAE_MONEDA)) & "', "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_SALCAP) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_SALCON) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_TASINT) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_COSEFE) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_DIAMOR) & ", "
   g_str_Parame = g_str_Parame & r_str_DiaPag & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("027", CStr(r_rst_HipMae!HIPMAE_SITUAC)) & "', "
   g_str_Parame = g_str_Parame & "'" & IIf(r_rst_HipMae!HIPMAE_FECCAN > 0, gf_FormatoFecha(CStr(r_rst_HipMae!HIPMAE_FECCAN)), "'', ") & "', "
   
   If r_int_NumCuo > 0 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(1).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(1).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(1).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(1).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 1 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(2).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(2).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(2).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(2).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 2 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(3).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(3).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(3).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(3).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 3 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(4).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(4).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(4).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(4).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 4 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(5).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(5).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(5).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(5).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 5 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(6).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(6).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(6).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(6).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 6 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(7).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(7).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(7).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(7).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 7 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(8).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(8).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(8).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(8).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   If r_int_NumCuo > 8 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(9).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(9).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(9).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(9).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   If r_int_NumCuo > 9 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(10).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(10).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(10).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(10).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 10 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(11).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(11).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(11).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(11).CuoVct_Situac & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If

   If r_int_NumCuo > 11 Then
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(12).CuoVct_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_arr_CuoVct(12).CuoVct_TotCuo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(12).CuoVct_FecVct & "', "
      g_str_Parame = g_str_Parame & "'" & r_arr_CuoVct(12).CuoVct_Situac & "') "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'') "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
       Exit Sub
   End If
   
   r_rst_HipMae.Close
   Set r_rst_HipMae = Nothing
   
   'Para obtener los Movimientos
   r_int_NumPag = 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & p_FecIni & "  AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & p_FecFin & "  AND "
   g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV <> 1101 AND "
   g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV <> 1103 AND "
   g_str_Parame = g_str_Parame & "CAJMOV_TIPMOV <> 2101 "
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_FECMOV DESC, CAJMOV_HORMOV DESC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      
      Do While Not r_rst_Princi.EOF
         r_int_NumLin = 1
         
         'Buscando Detalle en CRE_HIPPAG
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPPAG WHERE "
         g_str_Parame = g_str_Parame & "HIPPAG_NUMOPE = '" & p_NumOpe & "' AND "
         g_str_Parame = g_str_Parame & "HIPPAG_SUCMOV = '" & r_rst_Princi!CAJMOV_SUCMOV & "' AND "
         g_str_Parame = g_str_Parame & "HIPPAG_FECMOV = " & CStr(r_rst_Princi!CAJMOV_FECMOV) & " AND "
         g_str_Parame = g_str_Parame & "HIPPAG_NUMMOV = " & CStr(r_rst_Princi!CAJMOV_NUMMOV) & " "
         g_str_Parame = g_str_Parame & "ORDER BY HIPPAG_NUMCUO ASC, HIPPAG_NUMPAG ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_HipPag, 3) Then
             Exit Sub
         End If
         
         If Not (r_rst_HipPag.BOF And r_rst_HipPag.EOF) Then
            r_rst_HipPag.MoveFirst
            
            Do While Not r_rst_HipPag.EOF
               'Grabando Detalle de Movimiento
               
               If r_int_NumLin = 1 Then
                  Call opecaj_gs_EstCtaDet(modgen_g_str_NombPC, p_NomRpt, gf_Formato_NumOpe(p_NumOpe), r_int_NumPag, r_int_NumLin, moddat_gf_Consulta_ParDes("301", CStr(r_rst_Princi!CAJMOV_TIPMOV)), gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECMOV)), _
                                           r_rst_Princi!CAJMOV_SUCMOV & "-" & Mid(CStr(r_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(r_rst_Princi!CAJMOV_NUMMOV, "00000"), gf_FormatoFecha(CStr(r_rst_Princi!CAJMOV_FECDEP)), moddat_gf_Consulta_ParDes("505", CStr(r_rst_Princi!CAJMOV_CODBAN)), _
                                           Trim(r_rst_Princi!CAJMOV_NUMCTA), r_rst_Princi!CAJMOV_IMPPAG, r_rst_Princi!CAJMOV_ITFIMP, r_rst_Princi!CAJMOV_IMPTOT, r_rst_HipPag!HIPPAG_NUMCUO, "", r_rst_HipPag!HIPPAG_CAPITA, r_rst_HipPag!HIPPAG_INTERE, _
                                           r_rst_HipPag!HIPPAG_DESORG, r_rst_HipPag!HIPPAG_VIVORG, r_rst_HipPag!HIPPAG_OTRORG, r_rst_HipPag!HIPPAG_INTCOM, r_rst_HipPag!HIPPAG_INTMOR, r_rst_HipPag!HIPPAG_GASCOB, r_rst_HipPag!HIPPAG_OTRGAS, IIf(Not IsNull(r_rst_HipPag!HIPPAG_CAPBBP), r_rst_HipPag!HIPPAG_CAPBBP, 0), _
                                           IIf(Not IsNull(r_rst_HipPag!HIPPAG_INTBBP), r_rst_HipPag!HIPPAG_INTBBP, 0), "", 0, 0, 0, 0, 0, 0)
               Else
                  Call opecaj_gs_EstCtaDet(modgen_g_str_NombPC, p_NomRpt, gf_Formato_NumOpe(p_NumOpe), r_int_NumPag, r_int_NumLin, "", "", "", "", "", "", 0, 0, 0, r_rst_HipPag!HIPPAG_NUMCUO, "", r_rst_HipPag!HIPPAG_CAPITA, r_rst_HipPag!HIPPAG_INTERE, _
                                           r_rst_HipPag!HIPPAG_DESORG, r_rst_HipPag!HIPPAG_VIVORG, r_rst_HipPag!HIPPAG_OTRORG, r_rst_HipPag!HIPPAG_INTCOM, r_rst_HipPag!HIPPAG_INTMOR, r_rst_HipPag!HIPPAG_GASCOB, r_rst_HipPag!HIPPAG_OTRGAS, IIf(Not IsNull(r_rst_HipPag!HIPPAG_CAPBBP), r_rst_HipPag!HIPPAG_CAPBBP, 0), _
                                           IIf(Not IsNull(r_rst_HipPag!HIPPAG_INTBBP), r_rst_HipPag!HIPPAG_INTBBP, 0), "", 0, 0, 0, 0, 0, 0)
               End If
            
               r_int_NumLin = r_int_NumLin + 1
               
               r_rst_HipPag.MoveNext
               DoEvents
            Loop
         End If
         
         r_rst_HipPag.Close
         Set r_rst_HipPag = Nothing
         
         r_int_NumPag = r_int_NumPag + 1
         r_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Public Sub opecaj_gs_EstCtaDet(ByVal p_CodTer As String, ByVal p_NomRpt As String, ByVal p_NumOpe As String, ByVal p_NumPag As Integer, ByVal p_NumLin As Integer, ByVal p_TipMov As String, ByVal p_FecMov As String, _
                               ByVal p_NumMov As String, ByVal p_FecDep As String, ByVal p_CodBan As String, ByVal p_NumCta As String, ByVal p_ImpPag As Double, ByVal p_ImpITF As Double, ByVal p_ImpTot As Double, _
                               ByVal p_NumCuo As Integer, ByVal p_FecVct As String, ByVal p_Capita As Double, ByVal p_intere As Double, ByVal p_SegDes As Double, ByVal p_SegInm As Double, ByVal p_Portes As Double, _
                               ByVal p_IntCom As Double, ByVal p_IntMor As Double, ByVal p_GasCob As Double, ByVal p_OtrGas As Double, ByVal p_CapPBP As Double, ByVal p_IntPBP As Double, ByVal p_Situac As String, _
                               ByVal p_TotCuo As Double, ByVal p_PenPBP As Double, ByVal p_CobMor As Double, ByVal p_TotPag As Double, ByVal p_SalPen As Double, ByVal p_DiaAtr As Integer)
   
   Dim r_rst_Grabar     As ADODB.Recordset

   g_str_Parame = "INSERT INTO RPT_ECTADT ("
   g_str_Parame = g_str_Parame & "ECTADT_CODTER, "
   g_str_Parame = g_str_Parame & "ECTADT_NOMRPT, "
   g_str_Parame = g_str_Parame & "ECTADT_NUMOPE, "
   g_str_Parame = g_str_Parame & "ECTADT_NUMPAG, "
   g_str_Parame = g_str_Parame & "ECTADT_NUMLIN, "
   g_str_Parame = g_str_Parame & "ECTADT_TIPMOV, "
   g_str_Parame = g_str_Parame & "ECTADT_FECMOV, "
   g_str_Parame = g_str_Parame & "ECTADT_NUMMOV, "
   g_str_Parame = g_str_Parame & "ECTADT_FECPAG, "
   g_str_Parame = g_str_Parame & "ECTADT_BCOREC, "
   g_str_Parame = g_str_Parame & "ECTADT_CTAREC, "
   g_str_Parame = g_str_Parame & "ECTADT_IMPPAG, "
   g_str_Parame = g_str_Parame & "ECTADT_IMPITF, "
   g_str_Parame = g_str_Parame & "ECTADT_IMPORT, "
   g_str_Parame = g_str_Parame & "ECTADT_NUMCUO, "
   g_str_Parame = g_str_Parame & "ECTADT_FECVCT, "
   g_str_Parame = g_str_Parame & "ECTADT_CAPITA, "
   g_str_Parame = g_str_Parame & "ECTADT_INTERE, "
   g_str_Parame = g_str_Parame & "ECTADT_SEGDES, "
   g_str_Parame = g_str_Parame & "ECTADT_SEGINM, "
   g_str_Parame = g_str_Parame & "ECTADT_PORTES, "
   g_str_Parame = g_str_Parame & "ECTADT_INTCOM, "
   g_str_Parame = g_str_Parame & "ECTADT_INTMOR, "
   g_str_Parame = g_str_Parame & "ECTADT_GASCOB, "
   g_str_Parame = g_str_Parame & "ECTADT_OTRGAS, "
   g_str_Parame = g_str_Parame & "ECTADT_CAPPBP, "
   g_str_Parame = g_str_Parame & "ECTADT_INTPBP, "
   g_str_Parame = g_str_Parame & "ECTADT_SITUAC, "
   g_str_Parame = g_str_Parame & "ECTADT_TOTCUO, "
   g_str_Parame = g_str_Parame & "ECTADT_PENPBP, "
   g_str_Parame = g_str_Parame & "ECTADT_COBMOR, "
   g_str_Parame = g_str_Parame & "ECTADT_TOTPAG, "
   g_str_Parame = g_str_Parame & "ECTADT_SALPEN, "
   g_str_Parame = g_str_Parame & "ECTADT_DIAATR) "
   
   g_str_Parame = g_str_Parame & "VALUES ( "
   
   g_str_Parame = g_str_Parame & "'" & p_CodTer & "', "
   g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
   g_str_Parame = g_str_Parame & CStr(p_NumPag) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumLin) & ", "
   g_str_Parame = g_str_Parame & "'" & p_TipMov & "', "
   g_str_Parame = g_str_Parame & "'" & p_FecMov & "', "
   g_str_Parame = g_str_Parame & "'" & p_NumMov & "', "
   g_str_Parame = g_str_Parame & "'" & p_FecDep & "', "
   g_str_Parame = g_str_Parame & "'" & p_CodBan & "', "
   g_str_Parame = g_str_Parame & "'" & p_NumCta & "', "
   g_str_Parame = g_str_Parame & CStr(p_ImpPag) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ImpITF) & ", "
   g_str_Parame = g_str_Parame & CStr(p_ImpTot) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumCuo) & ", "
   g_str_Parame = g_str_Parame & "'" & p_FecVct & "', "
   g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
   g_str_Parame = g_str_Parame & CStr(p_intere) & ", "
   g_str_Parame = g_str_Parame & CStr(p_SegDes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_SegInm) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Portes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_IntCom) & ", "
   g_str_Parame = g_str_Parame & CStr(p_IntMor) & ", "
   g_str_Parame = g_str_Parame & CStr(p_GasCob) & ", "
   g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
   g_str_Parame = g_str_Parame & CStr(p_CapPBP) & ", "
   g_str_Parame = g_str_Parame & CStr(p_IntPBP) & ", "
   g_str_Parame = g_str_Parame & "'" & p_Situac & "', "
   g_str_Parame = g_str_Parame & CStr(p_TotCuo) & ", "
   g_str_Parame = g_str_Parame & CStr(p_PenPBP) & ", "
   g_str_Parame = g_str_Parame & CStr(p_CobMor) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TotPag) & ", "
   g_str_Parame = g_str_Parame & CStr(p_SalPen) & ", "
   g_str_Parame = g_str_Parame & CStr(p_DiaAtr) & ") "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
       Exit Sub
   End If
End Sub

Public Sub opecaj_gs_ResCuo(ByVal p_NumOpe As String)
Dim r_rst_HipMae     As ADODB.Recordset
Dim r_rst_HipCuo     As ADODB.Recordset
Dim r_rst_Grabar     As ADODB.Recordset

Dim r_str_Direcc     As String
Dim r_str_Distri     As String
Dim r_dbl_TotCuo     As Double
Dim r_dbl_PenPBP     As Double
Dim r_dbl_CobMor     As Double
Dim r_dbl_TotPag     As Double
Dim r_dbl_SalPen     As Double
Dim r_str_Situac     As String
Dim r_int_DiaAtr     As Integer
   
   'Obteniendo datos para Cabecera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_HipMae, 3) Then
       Exit Sub
   End If
   
   r_rst_HipMae.MoveFirst
         
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(r_rst_HipMae!hipmae_numsol, r_str_Direcc, r_str_Distri)
      
   'Insertando Cabecera
   g_str_Parame = "INSERT INTO RPT_ECTACB ("
   g_str_Parame = g_str_Parame & "ECTACB_CODTER, "
   g_str_Parame = g_str_Parame & "ECTACB_NOMRPT, "
   g_str_Parame = g_str_Parame & "ECTACB_NUMOPE, "
   g_str_Parame = g_str_Parame & "ECTACB_PERINI, "
   g_str_Parame = g_str_Parame & "ECTACB_PERFIN, "
   g_str_Parame = g_str_Parame & "ECTACB_PRODUC, "
   g_str_Parame = g_str_Parame & "ECTACB_CODCLI, "
   g_str_Parame = g_str_Parame & "ECTACB_DOCIDE, "
   g_str_Parame = g_str_Parame & "ECTACB_NOMCLI, "
   g_str_Parame = g_str_Parame & "ECTACB_DIRINM, "
   g_str_Parame = g_str_Parame & "ECTACB_DSTINM, "
   g_str_Parame = g_str_Parame & "ECTACB_FECDES, "
   g_str_Parame = g_str_Parame & "ECTACB_MTOPRE, "
   g_str_Parame = g_str_Parame & "ECTACB_INTCAP, "
   g_str_Parame = g_str_Parame & "ECTACB_PERGRA, "
   g_str_Parame = g_str_Parame & "ECTACB_NUMCUO, "
   g_str_Parame = g_str_Parame & "ECTACB_CUOPAG, "
   g_str_Parame = g_str_Parame & "ECTACB_CUOATR, "
   g_str_Parame = g_str_Parame & "ECTACB_MONEDA, "
   g_str_Parame = g_str_Parame & "ECTACB_SIMMON, "
   g_str_Parame = g_str_Parame & "ECTACB_SALCAP, "
   g_str_Parame = g_str_Parame & "ECTACB_SALCON, "
   g_str_Parame = g_str_Parame & "ECTACB_TASINT, "
   g_str_Parame = g_str_Parame & "ECTACB_COSEFE, "
   g_str_Parame = g_str_Parame & "ECTACB_DIAATR, "
   g_str_Parame = g_str_Parame & "ECTACB_DIAPAG, "
   g_str_Parame = g_str_Parame & "ECTACB_SITUAC, "
   g_str_Parame = g_str_Parame & "ECTACB_FECCAN, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS01, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS02, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS03, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS04, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS05, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS06, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS07, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS08, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS09, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS10, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS11, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXN12, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXI12, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXF12, "
   g_str_Parame = g_str_Parame & "ECTACB_PRXS12) "
   
   g_str_Parame = g_str_Parame & "VALUES ( "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'OPE_ESTCTA_02.RPT', "
   g_str_Parame = g_str_Parame & "'" & gf_Formato_NumOpe(p_NumOpe) & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_HipMae!HIPMAE_CODPRD) & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(r_rst_HipMae!HIPMAE_TDOCLI) & Trim(r_rst_HipMae!HIPMAE_NDOCLI) & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(r_rst_HipMae!HIPMAE_TDOCLI) & "-" & Trim(r_rst_HipMae!HIPMAE_NDOCLI) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_HipMae!HIPMAE_TDOCLI, Trim(r_rst_HipMae!HIPMAE_NDOCLI)) & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Direcc & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Distri & "', "
   g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(r_rst_HipMae!HIPMAE_FECDES) & "', "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_MTOPRE) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_INTCAP) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_PERGRA) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_NUMCUO) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_CUOPAG) & ", "
   If IsNull(r_rst_HipMae!HIPMAE_CUOATR) Then
      g_str_Parame = g_str_Parame & "0, "
   Else
      g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_CUOATR) & ", "
   End If
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("204", CStr(r_rst_HipMae!HIPMAE_MONEDA)) & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_HipMae!HIPMAE_MONEDA)) & "', "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_SALCAP) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_SALCON) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_TASINT) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_COSEFE) & ", "
   g_str_Parame = g_str_Parame & CStr(r_rst_HipMae!HIPMAE_DIAMOR) & ", "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("027", CStr(r_rst_HipMae!HIPMAE_SITUAC)) & "', "
   g_str_Parame = g_str_Parame & "'" & IIf(r_rst_HipMae!HIPMAE_FECCAN > 0, gf_FormatoFecha(CStr(r_rst_HipMae!HIPMAE_FECCAN)), "") & "', "
   
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "

   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "

   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "

   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "

   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "

   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
      
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "

   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
      
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
      
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
      
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
      
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "0, "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'') "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
      Exit Sub
   End If
   
   r_rst_HipMae.Close
   Set r_rst_HipMae = Nothing
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_HipCuo, 3) Then
      Exit Sub
   End If
   
   r_rst_HipCuo.MoveFirst
      
   Do While Not r_rst_HipCuo.EOF
      r_dbl_TotCuo = r_rst_HipCuo!HIPCUO_CAPITA + r_rst_HipCuo!HIPCUO_INTERE + r_rst_HipCuo!HIPCUO_DESORG + r_rst_HipCuo!HIPCUO_VIVORG + r_rst_HipCuo!HIPCUO_OTRORG
      r_dbl_PenPBP = r_rst_HipCuo!HIPCUO_CAPBBP + r_rst_HipCuo!HIPCUO_INTBBP
      r_dbl_CobMor = r_rst_HipCuo!HIPCUO_INTCOM + r_rst_HipCuo!HIPCUO_INTMOR + r_rst_HipCuo!HIPCUO_GASCOB + r_rst_HipCuo!HIPCUO_OTRGAS
      r_dbl_TotPag = r_rst_HipCuo!HIPCUO_CAPPAG + r_rst_HipCuo!HIPCUO_INTPAG + r_rst_HipCuo!HIPCUO_DESPAG + r_rst_HipCuo!HIPCUO_VIVPAG + r_rst_HipCuo!HIPCUO_OTRPAG + r_rst_HipCuo!HIPCUO_ICOPAG + r_rst_HipCuo!HIPCUO_IMOPAG + r_rst_HipCuo!HIPCUO_GCOPAG + r_rst_HipCuo!HIPCUO_OTGPAG + r_rst_HipCuo!HIPCUO_CBPPAG + r_rst_HipCuo!HIPCUO_IBPPAG
      
      r_dbl_SalPen = r_dbl_TotCuo + r_dbl_PenPBP + r_dbl_CobMor - r_dbl_TotPag
   
      If r_rst_HipCuo!HIPCUO_SITUAC = 1 Then
         r_str_Situac = "PAGADA"
         
         If CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECPAG))) <= CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT))) Then
            r_int_DiaAtr = 0
         Else
            r_int_DiaAtr = CInt(CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT))))
         End If
      Else
         If CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT))) >= date Then
            r_str_Situac = "X VENCER"
            r_int_DiaAtr = 0
         Else
            r_str_Situac = "ATRASADA"
            r_int_DiaAtr = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT))))
         End If
      End If
         
      Call opecaj_gs_EstCtaDet(modgen_g_str_NombPC, "OPE_ESTCTA_02.RPT", gf_Formato_NumOpe(p_NumOpe), r_rst_HipCuo!HIPCUO_NUMCUO, 1, "", "", "", gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECPAG)), "", "", 0, 0, 0, 0, gf_FormatoFecha(CStr(r_rst_HipCuo!HIPCUO_FECVCT)), r_rst_HipCuo!HIPCUO_SALCAP, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                               r_str_Situac, r_dbl_TotCuo, r_dbl_PenPBP, r_dbl_CobMor, r_dbl_TotPag, r_dbl_SalPen, r_int_DiaAtr)
   
      r_rst_HipCuo.MoveNext
   Loop
   
   r_rst_HipCuo.Close
   Set r_rst_HipCuo = Nothing
End Sub

