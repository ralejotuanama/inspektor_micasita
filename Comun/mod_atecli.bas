Attribute VB_Name = "modatecli"
Option Explicit

Global Const modatecli_g_con_IngSol = 10
Global Const modatecli_g_con_EvaCre = 11
Global Const modatecli_g_con_AceIni = 12
Global Const modatecli_g_con_EvaTas = 13
Global Const modatecli_g_con_EvaSeg = 14
Global Const modatecli_g_con_AprCre = 15
Global Const modatecli_g_con_AceFin = 16
Global Const modatecli_g_con_CerNPr = 17
Global Const modatecli_g_con_EvaLeg = 18
Global Const modatecli_g_con_PolSeg = 19
Global Const modatecli_g_con_TraCof = 20
Global Const modatecli_g_con_AutDes = 21
Global Const modatecli_g_con_Desemb = 22
Global Const modatecli_g_con_RecAdm = 91
Global Const modatecli_g_dbl_MtoFin = 140000
Global Const modatecli_g_dbl_BMSTas = 0.04

Global Const modgen_g_dbl_inm_tc = 3.2
Global Const modgen_g_dbl_inm_factor = 0.003
Global Const modgen_g_dbl_inm_ficha = 37

Global Const modgen_g_dbl_est_tc = 3.2
Global Const modgen_g_dbl_est_factor = 0.003
Global Const modgen_g_dbl_est_ficha = 37

Global Const modgen_g_dbl_gar_tc = 3.2
Global Const modgen_g_dbl_gar_factor = 0.0015
Global Const modgen_g_dbl_gar_ficha = 37

Global Const modgen_g_dbl_GasTas = 176
Global Const modgen_g_dbl_GasNot = 570
Global Const modgen_g_int_TipBie = 2

'Public modatecli_g_int_ActPri_Tit   As Integer        'Código de Actividad Económica Principal - Titular
'Public modatecli_g_int_ActSec_Tit   As Integer
'Public modatecli_g_str_CodCiu_Tit   As String         'Código de CIIU Actividad Económica Principal - Titular
'Public modatecli_g_str_GirCom_Tit   As String
'Public modatecli_g_str_SecEco_Tit   As String
'Public modatecli_g_int_TDoEmp_Tit   As Integer
'Public modatecli_g_str_NDoEmp_Tit   As String
'Public modatecli_g_int_ActPri_Cyg   As Integer
'Public modatecli_g_int_ActSec_Cyg   As Integer
'Public modatecli_g_str_CodCiu_Cyg   As String
'Public modatecli_g_str_GirCom_Cyg   As String
'Public modatecli_g_str_SecEco_Cyg   As String
'Public modatecli_g_int_TDoEmp_Cyg   As Integer
'Public modatecli_g_str_NDoEmp_Cyg   As String

'Ingresos - Inmuebles
Public Type modatecli_g_tpo_IngInm
   IngInm_TipInm     As Integer
   IngInm_Direcc     As String
   IngInm_FecAdq     As String
   IngInm_TipMon     As Integer
   IngInm_ImpVal     As Double
End Type
Public modatecli_g_arr_IngresInm()     As modatecli_g_tpo_IngInm

'Gastos - Tarjetas de Crédito
Public Type modatecli_g_tpo_GasTar
   GasTar_InsFin     As String
   GasTar_TipTar     As String
   GasTar_NumTar     As String
   GasTar_TipMon     As Integer
   GasTar_LinCre     As Double
   GasTar_SalPag     As Double
   GasTar_MonMin     As Double
End Type
Public modatecli_g_arr_GastosTar()     As modatecli_g_tpo_GasTar

'Gastos - Deudas Financieras
Public Type modatecli_g_tpo_GasFin
   GasFin_InsFin     As String
   GasFin_NumOpe     As String
   GasFin_TipMon     As Integer
   GasFin_MonOto     As Double
   GasFin_SalPag     As Double
   GasFin_MesPag     As Integer
   GasFin_CuoMen     As Double
End Type
Public modatecli_g_arr_GastosFin()     As modatecli_g_tpo_GasFin

'Gastos - Gastos Mensuales
Public Type modatecli_g_tpo_GasGas
   GasGas_TipGas     As Integer
   GasGas_ImpVal     As Double
End Type
Public modatecli_g_arr_GastosGas()     As modatecli_g_tpo_GasGas

'Flag de Gastos
Public modatecli_g_int_IngRegInm       As Integer
Public modatecli_g_int_GasRegTar       As Integer
Public modatecli_g_int_GasRegFin       As Integer
Public modatecli_g_int_GasRegGas       As Integer

'Referencias
Public Type modatecli_g_tpo_Refere
   Refere_TipPar     As Integer
   Refere_ApePat     As String
   Refere_ApeMat     As String
   Refere_Nombre     As String
   Refere_Telefo     As String
   Refere_Celula     As String
End Type
Public modatecli_g_arr_Refere(4)    As modatecli_g_tpo_Refere

'Datos del Inmueble
Public Type modatecli_g_tpo_DatInm
   DatInm_PriViv           As Integer
   DatInm_InmIde           As Integer
   DatInm_TipInm           As Integer
   DatInm_Modali           As String
   DatInm_TipVia           As Integer
   DatInm_NomVia           As String
   DatInm_Numero           As String
   DatInm_Interi           As String
   DatInm_TipZon           As Integer
   DatInm_NomZon           As String
   DatInm_UbiGeo           As String
   DatInm_Refere           As String
   DatInm_Estaci           As String
   DatInm_FlgEst           As Integer
   DatInm_PryMCs           As Integer
   DatInm_BcoPry           As String
   DatInm_CodPry           As String
   DatInm_NomPry           As String
   DatInm_FlgPro           As Integer
   DatInm_RazSoc_Pro       As String
   DatInm_TipDoc_Pro       As Integer
   DatInm_NumDoc_Pro       As String
   DatInm_TipVia_Pro       As Integer
   DatInm_NomVia_Pro       As String
   DatInm_NumVia_Pro       As String
   DatInm_IntDpt_Pro       As String
   DatInm_TipZon_Pro       As Integer
   DatInm_NomZon_Pro       As String
   DatInm_UbiGeo_Pro       As String
   DatInm_Refere_Pro       As String
   DatInm_Telefo_Pro       As String
   DatInm_FlgCon           As Integer
   DatInm_RazSoc_Con       As String
   DatInm_TipDoc_Con       As Integer
   DatInm_NumDoc_Con       As String
   DatInm_TipVia_Con       As Integer
   DatInm_NomVia_Con       As String
   DatInm_NumVia_Con       As String
   DatInm_IntDpt_Con       As String
   DatInm_TipZon_Con       As Integer
   DatInm_NomZon_Con       As String
   DatInm_UbiGeo_Con       As String
   DatInm_Refere_Con       As String
   DatInm_Telefo_Con       As String
End Type
Public modatecli_g_arr_DatInm(1)    As modatecli_g_tpo_DatInm

'Datos del Crédito
Public Type modatecli_g_tpo_DatCre
   'DatCre_Modali        As String
   DatCre_TipEva        As Integer
   DatCre_TipMon        As Integer
   DatCre_ComVta        As Double
   DatCre_MtoInm        As Double
   DatCre_MtoEst        As Double
   DatCre_ApoPro        As Double
   DatCre_CuoIni        As Double
   DatCre_FmvBbp        As Double
   DatCre_MefPbp        As Double
   DatCre_MtoAFP        As Double
   DatCre_MPSBMS        As Double
   DatCre_MtoBMS        As Double
   DatCre_MtoGCi        As Double
   DatCre_PreMto        As Double
   DatCre_MtoPre        As Double
   DatCre_BMSTas        As Double
   DatCre_TipCam        As Double
   DatCre_ComVta_Sol    As Double
   DatCre_MtoInm_Sol    As Double
   DatCre_MtoEst_Sol    As Double
   DatCre_ApoPro_Sol    As Double
   DatCre_CuoIni_Sol    As Double
   DatCre_FmvBbp_Sol    As Double
   DatCre_MefPbp_Sol    As Double
   DatCre_MtoAFP_Sol    As Double
   DatCre_MPSBMS_Sol    As Double
   DatCre_MtoBMS_Sol    As Double
   DatCre_MtoGCi_Sol    As Double
   DatCre_PreMto_Sol    As Double
   DatCre_MtoPre_Sol    As Double
   DatCre_ComVta_Dol    As Double
   DatCre_MtoInm_Dol    As Double
   DatCre_MtoEst_Dol    As Double
   DatCre_ApoPro_Dol    As Double
   DatCre_CuoIni_Dol    As Double
   DatCre_FmvBbp_Dol    As Double
   DatCre_MefPbp_Dol    As Double
   DatCre_MtoAFP_Dol    As Double
   DatCre_MPSBMS_Dol    As Double
   DatCre_MtoBMS_Dol    As Double
   DatCre_MtoGCi_Dol    As Double
   DatCre_PreMto_Dol    As Double
   DatCre_MtoPre_Dol    As Double
   DatCre_PlaAno        As Integer
   DatCre_PerGra        As Integer
   DatCre_CuoExt        As Integer
   DatCre_ESgDes        As String
   DatCre_TipSeg        As Integer
   DatCre_TasEsp        As Integer 'nuevo item
   DatCre_ESgViv        As String
   DatCre_DiaPag        As Integer
   DatCre_InsFin        As String
   DatCre_MonAho        As Integer
   DatCre_MtoAho        As Double
   DatCre_MesAho        As Integer
   DatCre_PriViv        As Integer
   DatCre_ConHip        As String
   DatCre_EjeSeg        As String
   DatCre_Observ        As String
End Type
Public modatecli_g_arr_DatCre(1)    As modatecli_g_tpo_DatCre

'Documentos del Credito
Public Type modatecli_g_tpo_DocCre
   DocCre_TipDoc  As Integer
   DocCre_CodAct  As Integer
   DocCre_CodGrp  As String
   DocCre_CodIte  As String
End Type
Public modatecli_g_arr_DocCre()     As modatecli_g_tpo_DocCre

'Flags Generales
Public modatecli_g_int_GastosTit       As Integer
Public modatecli_g_int_RefereTit       As Integer
Public modatecli_g_int_DatInmTit       As Integer
Public modatecli_g_int_DatCreTit       As Integer

'Lista de Rechazos de Solicitud
Public Type modatecli_g_tpo_LisRec
   LisRec_NumSol        As String
   LisRec_FecRec        As String
   LisRec_TipRec        As Integer
   LisRec_MotRec        As Integer
   LisRec_TipCli        As String
End Type
Public modatecli_g_arr_LisRec()     As modatecli_g_tpo_LisRec
Public modatecli_g_arr_CygRec()     As modatecli_g_tpo_LisRec

'Lista de Operaciones
Public Type modatecli_g_tpo_CreHip
   CreHip_NumOpe        As String
   CreHip_CodPrd        As String
   CreHip_CodMod        As String
   CreHip_FecAct        As String
   CreHip_Situac        As Integer
   CreHip_TipCli        As String
End Type
Public modatecli_g_arr_TitOpe()     As modatecli_g_tpo_CreHip
Public modatecli_g_arr_CygOpe()     As modatecli_g_tpo_CreHip

Public Sub modatecli_gs_Limpia_Refere(ByVal p_Indice As Integer)
   modatecli_g_arr_Refere(p_Indice).Refere_TipPar = 0
   modatecli_g_arr_Refere(p_Indice).Refere_ApePat = ""
   modatecli_g_arr_Refere(p_Indice).Refere_ApeMat = ""
   modatecli_g_arr_Refere(p_Indice).Refere_Nombre = ""
   modatecli_g_arr_Refere(p_Indice).Refere_Telefo = ""
   modatecli_g_arr_Refere(p_Indice).Refere_Celula = ""
End Sub

Public Sub modatecli_gs_Limpia_DatInm()
   modatecli_g_arr_DatInm(1).DatInm_InmIde = 0
   modatecli_g_arr_DatInm(1).DatInm_TipInm = 0
   modatecli_g_arr_DatInm(1).DatInm_Modali = ""
   modatecli_g_arr_DatInm(1).DatInm_TipVia = 0
   modatecli_g_arr_DatInm(1).DatInm_NomVia = ""
   modatecli_g_arr_DatInm(1).DatInm_Numero = ""
   modatecli_g_arr_DatInm(1).DatInm_Interi = ""
   modatecli_g_arr_DatInm(1).DatInm_TipZon = 0
   modatecli_g_arr_DatInm(1).DatInm_NomZon = ""
   modatecli_g_arr_DatInm(1).DatInm_UbiGeo = ""
   modatecli_g_arr_DatInm(1).DatInm_Refere = ""
   modatecli_g_arr_DatInm(1).DatInm_Estaci = ""
   modatecli_g_arr_DatInm(1).DatInm_PryMCs = 0
   modatecli_g_arr_DatInm(1).DatInm_CodPry = ""
   modatecli_g_arr_DatInm(1).DatInm_BcoPry = ""
   modatecli_g_arr_DatInm(1).DatInm_NomPry = ""
   modatecli_g_arr_DatInm(1).DatInm_FlgPro = 0
   modatecli_g_arr_DatInm(1).DatInm_RazSoc_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_TipDoc_Pro = 0
   modatecli_g_arr_DatInm(1).DatInm_NumDoc_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_TipVia_Pro = 0
   modatecli_g_arr_DatInm(1).DatInm_NomVia_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_NumVia_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_IntDpt_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_TipZon_Pro = 0
   modatecli_g_arr_DatInm(1).DatInm_NomZon_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_Refere_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_Telefo_Pro = ""
   modatecli_g_arr_DatInm(1).DatInm_FlgCon = 0
   modatecli_g_arr_DatInm(1).DatInm_RazSoc_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_TipDoc_Con = 0
   modatecli_g_arr_DatInm(1).DatInm_NumDoc_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_TipVia_Con = 0
   modatecli_g_arr_DatInm(1).DatInm_NomVia_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_NumVia_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_IntDpt_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_TipZon_Con = 0
   modatecli_g_arr_DatInm(1).DatInm_NomZon_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_UbiGeo_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_Refere_Con = ""
   modatecli_g_arr_DatInm(1).DatInm_Telefo_Con = ""
End Sub

Public Sub modatecli_gs_Limpia_DatCre()
   'modatecli_g_arr_DatCre(1).DatCre_Modali = ""
   modatecli_g_arr_DatCre(1).DatCre_TipEva = 0
   modatecli_g_arr_DatCre(1).DatCre_TipMon = 0
   modatecli_g_arr_DatCre(1).DatCre_ComVta = 0
   modatecli_g_arr_DatCre(1).DatCre_ApoPro = 0
   modatecli_g_arr_DatCre(1).DatCre_MtoPre = 0
   modatecli_g_arr_DatCre(1).DatCre_TipCam = 0
   modatecli_g_arr_DatCre(1).DatCre_ComVta_Sol = 0
   modatecli_g_arr_DatCre(1).DatCre_ApoPro_Sol = 0
   modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol = 0
   modatecli_g_arr_DatCre(1).DatCre_ComVta_Dol = 0
   modatecli_g_arr_DatCre(1).DatCre_ApoPro_Dol = 0
   modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol = 0
   modatecli_g_arr_DatCre(1).DatCre_PlaAno = 0
   modatecli_g_arr_DatCre(1).DatCre_PerGra = 0
   modatecli_g_arr_DatCre(1).DatCre_CuoExt = 0
   modatecli_g_arr_DatCre(1).DatCre_ESgDes = ""
   modatecli_g_arr_DatCre(1).DatCre_TipSeg = 0
   modatecli_g_arr_DatCre(1).DatCre_ESgViv = ""
   modatecli_g_arr_DatCre(1).DatCre_DiaPag = 0
   modatecli_g_arr_DatCre(1).DatCre_InsFin = ""
   modatecli_g_arr_DatCre(1).DatCre_MonAho = 0
   modatecli_g_arr_DatCre(1).DatCre_MtoAho = 0
   modatecli_g_arr_DatCre(1).DatCre_MesAho = 0
   modatecli_g_arr_DatCre(1).DatCre_ConHip = ""
   modatecli_g_arr_DatCre(1).DatCre_EjeSeg = ""
   modatecli_g_arr_DatCre(1).DatCre_Observ = ""
End Sub

Public Function modatecli_gf_Rechaz_SolMae(ByVal p_NumSol As String, p_TipRec As Integer, p_MotRec As Integer) As Integer
   modatecli_gf_Rechaz_SolMae = False
   
   'Actualizando en Tabla de Créditos
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_MODIFICA_CRE_SOLMAE_RECHAZ ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipRec) & ", "
      g_str_Parame = g_str_Parame & CStr(p_MotRec) & ", "
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_MODIFICA_CRE_SOLMAE_RECHAZ. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   modatecli_gf_Rechaz_SolMae = True
End Function

Public Function modatecli_gf_ActIns_SolMae(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As Integer
   modatecli_gf_ActIns_SolMae = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_MODIFICA_CRE_SOLMAE_CODINS ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                              'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                               'Instancia
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_MODIFICA_CRE_SOLMAE_CODINS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   modatecli_gf_ActIns_SolMae = True
End Function

Public Function atecli_gf_Buscar_SolVig(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_TipCli As Integer) As Integer
   '1 - Cliente Titular
   '2 - Cónyuge
   
   atecli_gf_Buscar_SolVig = False
   
   'Verificando Solicitudes en Trámite
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   If p_TipCli = 1 Then
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(p_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & p_NumDoc & "' AND "
   Else
      g_str_Parame = g_str_Parame & "SOLMAE_CYGTDO = " & CStr(p_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_CYGNDO = '" & p_NumDoc & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If p_TipCli = 1 Then
         MsgBox "El Cliente ya presenta una Solicitud de Crédito en Trámite como Titular. " & Chr(10) & Chr(13) & "Solicitud Nro. " & Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4), vbExclamation, modgen_g_str_NomPlt
      Else
         MsgBox "El Cliente ya presenta una Solicitud de Crédito en Trámite como Cónyuge. " & Chr(10) & Chr(13) & "Solicitud Nro. " & Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4), vbExclamation, modgen_g_str_NomPlt
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Function
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   atecli_gf_Buscar_SolVig = True
End Function

Public Function atecli_gf_Buscar_BasNeg(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As Integer
   atecli_gf_Buscar_BasNeg = False
   
   'Validar que Cliente no se encuentre en Base Negativa
   g_str_Parame = "SELECT * FROM CRE_BASNEG WHERE "
   g_str_Parame = g_str_Parame & "BASNEG_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "BASNEG_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "BASNEG_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If g_rst_Princi!BASNEG_TIPBAS = 1 Then
         MsgBox "El Cliente o su Cónyuge se encuentra registrado como Negativo.", vbExclamation, modgen_g_str_NomPlt
      End If
      If g_rst_Princi!BASNEG_TIPBAS = 2 Then
         MsgBox "El Cliente o su Cónyuge se encuentra registrado como PEP.", vbExclamation, modgen_g_str_NomPlt
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   atecli_gf_Buscar_BasNeg = True
End Function

Public Sub atecli_gs_Buscar_SolRec(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_TipCli As Integer)
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If p_TipCli = 1 Then
      ReDim modatecli_g_arr_LisRec(0)
   Else
      ReDim modatecli_g_arr_CygRec(0)
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If p_TipCli = 1 Then
            ReDim Preserve modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec) + 1)
            
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_FecRec = Format(g_rst_Princi!SOLMAE_FECREC, "00000000")
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_TipRec = g_rst_Princi!SOLMAE_TIPREC
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_MotRec = g_rst_Princi!SOLMAE_MOTREC
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_TipCli = 1
         Else
            ReDim Preserve modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec) + 1)
            
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_FecRec = Format(g_rst_Princi!SOLMAE_FECREC, "00000000")
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_TipRec = g_rst_Princi!SOLMAE_TIPREC
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_MotRec = g_rst_Princi!SOLMAE_MOTREC
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_TipCli = 1
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Como Cónyuge
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CYGTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CYGNDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If p_TipCli = 1 Then
            ReDim Preserve modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec) + 1)
            
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_FecRec = Format(g_rst_Princi!SOLMAE_FECREC, "00000000")
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_TipRec = g_rst_Princi!SOLMAE_TIPREC
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_MotRec = g_rst_Princi!SOLMAE_MOTREC
            modatecli_g_arr_LisRec(UBound(modatecli_g_arr_LisRec)).LisRec_TipCli = 2
         Else
            ReDim Preserve modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec) + 1)
            
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_FecRec = Format(g_rst_Princi!SOLMAE_FECREC, "00000000")
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_TipRec = g_rst_Princi!SOLMAE_TIPREC
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_MotRec = g_rst_Princi!SOLMAE_MOTREC
            modatecli_g_arr_CygRec(UBound(modatecli_g_arr_CygRec)).LisRec_TipCli = 2
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub atecli_gs_Buscar_CreHip(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_TipCli As Integer)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_TDOCLI = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_NDOCLI = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC <> 3 AND HIPMAE_SITUAC <> 6 AND HIPMAE_SITUAC <> 7 AND HIPMAE_SITUAC <> 8 AND HIPMAE_SITUAC <> 9)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If p_TipCli = 1 Then
      ReDim modatecli_g_arr_TitOpe(0)
   Else
      ReDim modatecli_g_arr_CygOpe(0)
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If p_TipCli = 1 Then
            ReDim Preserve modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe) + 1)
            
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_CodPrd = g_rst_Princi!HIPMAE_CODPRD
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_CodMod = g_rst_Princi!HIPMAE_CODMOD
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_FecAct = Format(g_rst_Princi!HIPMAE_FECACT, "00000000")
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_Situac = g_rst_Princi!HIPMAE_SITUAC
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_TipCli = 1
         Else
            ReDim Preserve modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe) + 1)
            
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_CodPrd = g_rst_Princi!HIPMAE_CODPRD
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_CodMod = g_rst_Princi!HIPMAE_CODMOD
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_FecAct = Format(g_rst_Princi!HIPMAE_FECACT, "00000000")
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_Situac = g_rst_Princi!HIPMAE_SITUAC
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_TipCli = 1
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Como Cónyuge
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_TDOCYG = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_NDOCYG = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC <> 3 AND HIPMAE_SITUAC <> 6 AND HIPMAE_SITUAC <> 7 AND HIPMAE_SITUAC <> 8 AND HIPMAE_SITUAC <> 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If p_TipCli = 1 Then
            ReDim Preserve modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe) + 1)
            
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_CodPrd = g_rst_Princi!HIPMAE_CODPRD
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_CodMod = g_rst_Princi!HIPMAE_CODMOD
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_FecAct = Format(g_rst_Princi!HIPMAE_FECACT, "00000000")
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_Situac = g_rst_Princi!HIPMAE_SITUAC
            modatecli_g_arr_TitOpe(UBound(modatecli_g_arr_TitOpe)).CreHip_TipCli = 2
         Else
            ReDim Preserve modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe) + 1)
            
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_CodPrd = g_rst_Princi!HIPMAE_CODPRD
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_CodMod = g_rst_Princi!HIPMAE_CODMOD
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_FecAct = Format(g_rst_Princi!HIPMAE_FECACT, "00000000")
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_Situac = g_rst_Princi!HIPMAE_SITUAC
            modatecli_g_arr_CygOpe(UBound(modatecli_g_arr_CygOpe)).CreHip_TipCli = 2
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

