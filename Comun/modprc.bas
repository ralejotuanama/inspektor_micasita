Attribute VB_Name = "modprc"
Option Explicit

Public Type modprc_g_tpo_Genera
   Genera_Codigo     As String
   Genera_Nombre     As String
   Genera_TipPar     As Integer
   Genera_TipVal     As Integer
   Genera_Cantid     As Double
   Genera_ValMin     As Double
   Genera_ValMax     As Double
   Genera_DiaVc1     As Integer
   Genera_DiaVc2     As Integer
End Type

Public Type modprc_g_tpo_CreHip
   CreHip_NumOpe     As String
   CreHip_SitCtb     As Integer
   CreHip_SalCap     As Double
   CreHip_SalCon     As Double
   CreHip_CapVig     As Double
   CreHip_CapVen     As Double
   CreHip_IntVen     As Double
   CreHip_IntMor     As Double
   CreHip_IntCom     As Double
   CreHip_GasCob     As Double
   CreHip_OtrGas     As Double
   CreHip_IMoVig     As Double
   CreHip_ICoVig     As Double
   CreHip_GCoVig     As Double
   CreHip_OtGVig     As Double
   CreHip_CuoAtr     As Integer
   CreHip_NCuoTC     As Integer
   CreHip_CapiTC     As Double
   CreHip_InteTC     As Double
   CreHip_ComiTC     As Double
   CreHip_DevPBP     As Double
   CreHip_PrvICo     As Double
   CreHip_PrvCCo     As Double
   CreHip_MtoAde     As Double
   CreHip_IniAde     As String
   CreHip_FecDes     As String
   CreHip_DevAnt     As String
   CreHip_CuoDev     As Integer
   CreHip_MtoPre     As Double
   CreHip_IntDev     As Double
   CreHip_MtoNCo     As Double
   CreHip_UltCap     As Double
   CreHip_CVcMes     As Double
   CreHip_FecPPg     As String
End Type

Public Type modprc_g_tpo_LogPro
   LogPro_CodPro     As String
   LogPro_FInEje     As String
   LogPro_HInEje     As String
   LogPro_NumErr     As Long
End Type

Public Type modprc_g_tpo_DetGar
   DetGar_Codigo     As Integer
   DetGar_ClaGar     As Integer
End Type

Public Type modprc_g_tpo_TipPrv
   TipPrv_TipPrv     As Integer
   TipPrv_CodCla     As Integer
   TipPrv_ClaGar     As Integer
   TipPrv_Porcen     As Double
End Type

Public Type modprc_g_tpo_MatCtb
   MatCtb_CodBan     As String
   MatCtb_NumCta     As String
   MatCtb_TipMon     As Integer
   MatCtb_CtaCtb     As String
   MatCtb_EmpGrp     As String
   MatCtb_FTCSun     As String
   MatCtb_ComSun     As Double
   MatCtb_VenSun     As Double
   MatCtb_CodMat     As String
   MatCtb_DesCab     As String
   MatCtb_CodLib     As Integer
   MatCtb_NumIte     As Integer
   MatCtb_DesDet     As String
   MatCtb_TipCon     As Integer
   MatCtb_ConCtb     As String
   MatCtb_TipTCa     As Integer
   MatCtb_FlgDHb     As Integer
   MatCtb_ConOpe     As String
End Type

Public Type modprc_g_tpo_CtaBan
   CtaBan_CodBan     As String
   CtaBan_NumCta     As String
   CtaBan_CtaCtb     As String
   CtaBan_TipCta     As String
End Type

Public Type modprc_g_tpo_MatDet
   MatDet_CodMat     As String
   MatDet_CodPrd     As String
   MatDet_DesCab     As String
   MatDet_TipMon     As Integer
   MatDet_SitCre     As Integer
   MatDet_DesDet     As String
   MatDet_CtbCon     As String
   MatDet_DebHab     As String
   MatDet_TipCam     As Integer
   MatDet_NroLib     As String
   MatDet_OpeCon     As String
   CtaPrd_CtaCtb     As String
End Type

Public Type modprc_g_tpo_MatPro
   MatPro_CodPrd     As String
   MatPro_Campo      As String
   MatPro_NumCta     As String
   MatPro_DebHab     As String
   MatPro_Glosa      As String
   MatPro_NroLib     As Integer
   MatPro_TipNot     As String
   MatPro_DesNot     As String
   MatPro_DesCab     As String
End Type

Public Type modprc_g_tpo_Matriz
   Matriz_CtaCtb     As String
   Matriz_FlagDH     As String
   Matriz_OpeCon     As String
   Matriz_ConCtb     As String
   Matriz_import     As Double
   Matriz_DesNot     As String
   Matriz_TipCta     As String
End Type

Public Type modprc_g_tpo_CtaPrd
   CtaPrd_CodPrd     As String
   CtaPrd_CtbCon     As String
   CtaPrd_SitCre     As String
   CtaPrd_CtaCtb     As String
End Type

Public Type modprc_g_tpo_CtaCom
   CtaCom_DocIde     As String
   CtaCom_Descri     As String
   CtaCom_CtaCtb     As String
End Type

Public Type modprc_g_tpo_CieEje
   CieEje_DesCab     As String
   CieEje_SitCre     As String
   CieEje_DesDet     As String
   CieEje_CtbCon     As String
   CieEje_DebHab     As Double
   CieEje_MtoCtb     As String
   CieEje_ctactb     As String
   CieEje_NroLib     As String
   CieEje_TipMon     As String
End Type

Public modprc_g_rst_Princi       As ADODB.Recordset
Public modprc_g_rst_Genera       As ADODB.Recordset
Public modprc_g_rst_Auxili       As ADODB.Recordset
Public modprc_g_rst_Grabar       As ADODB.Recordset
Public modprc_g_rst_Listas       As ADODB.Recordset
Public modprc_g_rst_Consul       As ADODB.Recordset
Public modprc_g_str_CodPro       As String
Public modprc_g_str_CadEje       As String
Global Const modprc_g_dbl_ComBcoSol = 2.5
Global Const modprc_g_dbl_ComBcoDol = 0.75

Public Sub modprc_gs_GrabaErrorProceso(ByVal p_CodPrg As String, ByVal p_FecEje As String, ByVal p_HorEje As String, ByVal p_NumErr As Long, ByVal p_DesErr As String)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "INSERT INTO SEG_PROLGD ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGD_CODPRG, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGD_FECEJE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGD_HOREJE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGD_NUMOCU, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGD_DESOCU) "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "VALUES("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CodPrg & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_FecEje & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_HorEje & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumErr) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_DesErr & "')"

   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
      MsgBox "Proces Critico al Grabar sobre SEG_PROLGD.", vbCritical, p_CodPrg
      Exit Sub
   End If
End Sub

Public Sub modprc_gs_GrabaCabeceraLogProceso(ByVal p_CodPrg As String, ByVal p_FecEje As String, ByVal p_HorEje As String, ByVal p_FecFin As String, ByVal p_HorFin As String, ByVal p_RegPro As Long, ByVal p_RegErr As Long, _
                                             ByVal p_CodEmp As String, ByVal p_CodTit As String, ByVal p_PerPar As Integer, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_IniPro As String, ByVal p_FinPro As String)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "INSERT INTO SEG_PROLGC ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_CODPRG, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_FECEJE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_HOREJE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_FECFIN, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_HORFIN, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_REGPRO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_REGERR, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_USUCRE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_PLTCRE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_TERCRE, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_CODEMP, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_CODTIT, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_PERPAR, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_PERMES, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_PERANO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_INIPRO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "PROLGC_FINPRO) "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "VALUES("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CodPrg & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_FecEje & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_HorEje & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_FecFin & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_HorFin & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_RegPro) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_RegErr) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & modgen_g_str_CodUsu & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & UCase(App.EXEName) & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & modgen_g_str_NombPC & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CodEmp & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CodTit & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerPar) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerMes) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerAno) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_IniPro & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & p_FinPro & ") "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
      MsgBox "Proces Critico al Grabar sobre SEG_PROLGC.", vbCritical, p_CodPrg
      Exit Sub
   End If
End Sub

Public Function modprc_gf_InteresMensual(p_TasInt As Double) As Double
   modprc_gf_InteresMensual = (1 + p_TasInt / 100) ^ (8.33333333333333E-02) - 1
End Function

Public Function modprc_gf_CalculaInteres(ByVal p_TasMen As Double, ByVal p_NumDia As Integer, ByVal p_BasImp As Double) As Double
   modprc_gf_CalculaInteres = (1 + p_TasMen) ^ (p_NumDia / 30) * p_BasImp - p_BasImp
End Function

Public Function modprc_gf_Consulta_ParametroSubPrd(p_Arregl() As modprc_g_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodGrp As String, ByVal p_CodIte As String) As Integer
Dim r_str_Parame     As String
   
   modprc_gf_Consulta_ParametroSubPrd = False
   ReDim p_Arregl(0)
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM CRE_PARPRD "
   r_str_Parame = r_str_Parame & " WHERE PARPRD_CODPRD = '" & p_CodPrd & "' "
   r_str_Parame = r_str_Parame & "   AND PARPRD_CODSUB = '" & p_CodSub & "' "
   r_str_Parame = r_str_Parame & "   AND PARPRD_CODGRP = '" & p_CodGrp & "' "
   r_str_Parame = r_str_Parame & "   AND PARPRD_CODITE = '" & p_CodIte & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, modprc_g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If modprc_g_rst_Listas.BOF And modprc_g_rst_Listas.EOF Then
      modprc_g_rst_Listas.Close
      Set modprc_g_rst_Listas = Nothing
      Exit Function
   End If
   
   modprc_g_rst_Listas.MoveFirst
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(modprc_g_rst_Listas!PARPRD_CODITE)
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(modprc_g_rst_Listas!PARPRD_DESCRI)
   p_Arregl(UBound(p_Arregl)).Genera_TipVal = modprc_g_rst_Listas!PARPRD_TIPVAL
   p_Arregl(UBound(p_Arregl)).Genera_TipPar = modprc_g_rst_Listas!PARPRD_TIPPAR
   p_Arregl(UBound(p_Arregl)).Genera_Cantid = modprc_g_rst_Listas!PARPRD_CANTID
   p_Arregl(UBound(p_Arregl)).Genera_ValMin = modprc_g_rst_Listas!PARPRD_VALMIN
   p_Arregl(UBound(p_Arregl)).Genera_ValMax = modprc_g_rst_Listas!PARPRD_VALMAX
   
   modprc_g_rst_Listas.Close
   Set modprc_g_rst_Listas = Nothing
   
   modprc_gf_Consulta_ParametroSubPrd = True
End Function

Public Sub modprc_gs_DestinoCorreo(ByVal p_NomPro As String, ByVal p_TipEnv As Integer, ByVal p_Asunto As String, ByVal p_Contenido As String, p_LisCor() As modprc_g_tpo_Genera, p_Sesion As MAPISession, p_Mensaje As MAPIMessages, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
Dim r_str_DirCor     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_str_Parame     As String
   
   'Generando Reporte
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM SEG_BATCOR "
   r_str_Parame = r_str_Parame & " WHERE BATCOR_CODPRG = '" & p_NomPro & "' "
   r_str_Parame = r_str_Parame & "   AND BATCOR_TIPENV = " & CStr(p_TipEnv)
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      ReDim p_LisCor(0)
      
      r_rst_Princi.MoveFirst
      Do While Not r_rst_Princi.EOF
         r_str_DirCor = modprc_gf_Buscar_CorreoEjecutivo(r_rst_Princi!BATCOR_CODEJE)
         If Not modprc_gf_Buscar_CorreoRepetido(p_LisCor, r_str_DirCor) Then
            ReDim Preserve p_LisCor(UBound(p_LisCor) + 1)
            p_LisCor(UBound(p_LisCor)).Genera_Codigo = r_str_DirCor
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      
      Call modprc_gs_EnviaCorreo(p_Sesion, p_Mensaje, p_LisCor(), p_Asunto, p_Contenido, p_NomFil, p_RutFil)
   End If
End Sub

Public Function modprc_gf_Buscar_CorreoEjecutivo(ByVal p_CodEje As String) As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_str_Parame     As String
   
   modprc_gf_Buscar_CorreoEjecutivo = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM CRE_EJECMC "
   r_str_Parame = r_str_Parame & " WHERE EJECMC_CODEJE = '" & p_CodEje & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      modprc_gf_Buscar_CorreoEjecutivo = Trim(r_rst_Princi!EJECMC_DIRELE & "")
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Function

Public Function modprc_gf_Buscar_CorreoRepetido(p_Arregl() As modprc_g_tpo_Genera, ByVal p_DirEle As String) As Integer
Dim r_int_Contad     As Integer
   
   modprc_gf_Buscar_CorreoRepetido = False
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      If p_Arregl(r_int_Contad).Genera_Codigo = p_DirEle Then
         modprc_gf_Buscar_CorreoRepetido = True
         Exit For
      End If
   Next r_int_Contad
End Function

Public Sub modprc_gs_EnviaCorreo(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, p_Arregl() As modprc_g_tpo_Genera, p_Asunto As String, p_Contenido As String, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
Dim r_int_Contad      As Integer
   
   On Error GoTo modprc_gs_EnviaCorreo_Error

   'Inicializa
   p_Sesion.DownLoadMail = False
   p_Sesion.NewSession = True
   p_Sesion.SignOn
   p_Mensaje.SessionID = p_Sesion.SessionID
  
   'Envío
   p_Mensaje.Compose
   
   For r_int_Contad = 0 To UBound(p_Arregl) - 1
      p_Mensaje.RecipIndex = r_int_Contad
      p_Mensaje.RecipDisplayName = p_Arregl(r_int_Contad + 1).Genera_Codigo
   Next r_int_Contad

   p_Mensaje.MsgSubject = p_Asunto
   p_Mensaje.MsgNoteText = p_Contenido
   
   If Len(Trim(p_NomFil)) > 0 Then
      p_Mensaje.AttachmentIndex = 0
      p_Mensaje.AttachmentName = p_NomFil
      p_Mensaje.AttachmentPathName = p_RutFil & p_NomFil
      p_Mensaje.AttachmentPosition = 0
      p_Mensaje.AttachmentType = mapData
   End If
   
   p_Mensaje.send
   DoEvents
  
   'Cierra la sesión
   p_Sesion.SignOff
   Exit Sub
  
modprc_gs_EnviaCorreo_Error:
   p_Sesion.SignOff
   MsgBox Err.Description, vbCritical
End Sub

Public Sub modprc_gs_CargaSituacion(ByVal p_ClaCre As String, ByRef p_DiaVc1 As Integer, ByRef p_DiaVc2 As Integer, p_LogPro() As modprc_g_tpo_LogPro)
Dim r_rst_Princi     As ADODB.Recordset

   modprc_g_str_CadEje = "SELECT SITCRE_DIAVC1, SITCRE_DIAVC2 " & _
                         "  FROM CTB_SITCRE " & _
                         " WHERE SITCRE_CLACRE = '" & p_ClaCre & "' " & _
                         "   AND SITCRE_CODSIT = '5'"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Princi, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al leer tabla CTB_SITCRE.")
      Exit Sub
   End If

   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      p_DiaVc1 = CInt(r_rst_Princi!SITCRE_DIAVC1)
      p_DiaVc2 = CInt(r_rst_Princi!SITCRE_DIAVC2)
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Public Sub modprc_gs_CargaTiposClasif(ByVal p_ClaCre As String, p_TipCla() As modprc_g_tpo_Genera, p_LogPro() As modprc_g_tpo_LogPro)
Dim r_rst_Princi     As ADODB.Recordset

   ReDim p_TipCla(0)

   modprc_g_str_CadEje = "SELECT TIPCLA_CODIGO, TIPCLA_DIAINI, TIPCLA_DIAFIN " & _
                         "  FROM CTB_TIPCLA " & _
                         " WHERE TIPCLA_TIPCRE = '" & p_ClaCre & "' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Princi, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al leer tabla CTB_TIPCLA.")
      Exit Sub
   End If

   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      Do While Not r_rst_Princi.EOF
         ReDim Preserve p_TipCla(UBound(p_TipCla) + 1)
         p_TipCla(UBound(p_TipCla)).Genera_Codigo = Trim(r_rst_Princi!TIPCLA_CODIGO)
         p_TipCla(UBound(p_TipCla)).Genera_DiaVc1 = r_rst_Princi!TIPCLA_DIAINI
         p_TipCla(UBound(p_TipCla)).Genera_DiaVc2 = r_rst_Princi!TIPCLA_DIAFIN
      
         r_rst_Princi.MoveNext
      Loop
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Public Sub modprc_gs_EvaluacionSituacion_CreHip(p_CreHip() As modprc_g_tpo_CreHip, p_LogPro() As modprc_g_tpo_LogPro, ByVal p_FecPro As String, ByVal p_DiaAtr As Integer, ByVal p_DiaVc1 As Integer, ByVal p_DiaVc2 As Integer)
Dim r_rst_Genera     As ADODB.Recordset
Dim r_dbl_CapVen     As Double
   
   r_dbl_CapVen = 0
   p_CreHip(1).CreHip_CuoAtr = 0
   p_CreHip(1).CreHip_IntVen = 0
   p_CreHip(1).CreHip_IntMor = 0
   p_CreHip(1).CreHip_IntCom = 0
   p_CreHip(1).CreHip_GasCob = 0
   p_CreHip(1).CreHip_OtrGas = 0
   p_CreHip(1).CreHip_IMoVig = 0
   p_CreHip(1).CreHip_ICoVig = 0
   p_CreHip(1).CreHip_GCoVig = 0
   p_CreHip(1).CreHip_OtGVig = 0
   
   If p_DiaAtr <= p_DiaVc1 Then
      'Vigente
      p_CreHip(1).CreHip_SitCtb = 1
      p_CreHip(1).CreHip_CapVig = p_CreHip(1).CreHip_SalCap + p_CreHip(1).CreHip_SalCon
      p_CreHip(1).CreHip_CapVen = 0
      
   ElseIf p_DiaAtr > p_DiaVc1 And p_DiaAtr <= p_DiaVc2 Then
      'Vencido (Parcial)
      p_CreHip(1).CreHip_SitCtb = 4
      
      'Para leer Cuotas Vencidas
      modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCUO " & _
                            " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                            "   AND HIPCUO_TIPCRO = 1 " & _
                            "   AND HIPCUO_FECVCT < " & Format(CDate(p_FecPro), "yyyymmdd") & " " & _
                            "   AND HIPCUO_SITUAC = 2 "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
         p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
         Exit Sub
      End If
         
      If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
         r_rst_Genera.MoveFirst
         Do While Not r_rst_Genera.EOF
            p_CreHip(1).CreHip_CuoAtr = p_CreHip(1).CreHip_CuoAtr + 1
               
            If CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT)))) > p_DiaVc1 Then
               r_dbl_CapVen = r_dbl_CapVen + r_rst_Genera!HIPCUO_CAPITA - r_rst_Genera!HIPCUO_CAPPAG
               p_CreHip(1).CreHip_IntVen = p_CreHip(1).CreHip_IntVen + CDbl(Format(r_rst_Genera!HIPCUO_INTERE - r_rst_Genera!HIPCUO_INTPAG, "#######0.00"))
               p_CreHip(1).CreHip_IntMor = p_CreHip(1).CreHip_IntMor + CDbl(Format(r_rst_Genera!HIPCUO_INTMOR - r_rst_Genera!HIPCUO_IMOPAG, "#######0.00"))
               p_CreHip(1).CreHip_IntCom = p_CreHip(1).CreHip_IntCom + CDbl(Format(r_rst_Genera!HIPCUO_INTCOM - r_rst_Genera!HIPCUO_ICOPAG, "#######0.00"))
               p_CreHip(1).CreHip_GasCob = p_CreHip(1).CreHip_GasCob + CDbl(Format(r_rst_Genera!HIPCUO_GASCOB - r_rst_Genera!HIPCUO_GCOPAG, "#######0.00"))
               p_CreHip(1).CreHip_OtrGas = p_CreHip(1).CreHip_OtrGas + CDbl(Format(r_rst_Genera!HIPCUO_OTRGAS - r_rst_Genera!HIPCUO_OTGPAG, "#######0.00"))
            Else
               p_CreHip(1).CreHip_IMoVig = p_CreHip(1).CreHip_IMoVig + CDbl(Format(r_rst_Genera!HIPCUO_INTMOR - r_rst_Genera!HIPCUO_IMOPAG, "#######0.00"))
               p_CreHip(1).CreHip_ICoVig = p_CreHip(1).CreHip_ICoVig + CDbl(Format(r_rst_Genera!HIPCUO_INTCOM - r_rst_Genera!HIPCUO_ICOPAG, "#######0.00"))
               p_CreHip(1).CreHip_GCoVig = p_CreHip(1).CreHip_GCoVig + CDbl(Format(r_rst_Genera!HIPCUO_GASCOB - r_rst_Genera!HIPCUO_GCOPAG, "#######0.00"))
               p_CreHip(1).CreHip_OtGVig = p_CreHip(1).CreHip_OtGVig + CDbl(Format(r_rst_Genera!HIPCUO_OTRGAS - r_rst_Genera!HIPCUO_OTGPAG, "#######0.00"))
            End If
            
            r_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   
      p_CreHip(1).CreHip_CapVig = p_CreHip(1).CreHip_SalCap + p_CreHip(1).CreHip_SalCon - r_dbl_CapVen
      p_CreHip(1).CreHip_CapVen = r_dbl_CapVen
      
   ElseIf p_DiaAtr > p_DiaVc2 Then
      'Vencido (Total)
      p_CreHip(1).CreHip_SitCtb = 4
      p_CreHip(1).CreHip_CapVig = 0
      p_CreHip(1).CreHip_CapVen = p_CreHip(1).CreHip_SalCap + p_CreHip(1).CreHip_SalCon
      
      'Para leer Cuotas Vencidas
      modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCUO " & _
                            " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                            "   AND HIPCUO_TIPCRO = 1 " & _
                            "   AND HIPCUO_FECVCT < " & Format(CDate(p_FecPro), "yyyymmdd") & " " & _
                            "   AND HIPCUO_SITUAC = 2 "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
         p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
         Exit Sub
      End If
         
      If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
         r_rst_Genera.MoveFirst
         Do While Not r_rst_Genera.EOF
            p_CreHip(1).CreHip_CuoAtr = p_CreHip(1).CreHip_CuoAtr + 1
            p_CreHip(1).CreHip_IntVen = p_CreHip(1).CreHip_IntVen + CDbl(Format(r_rst_Genera!HIPCUO_INTERE - r_rst_Genera!HIPCUO_INTPAG, "#######0.00"))
            p_CreHip(1).CreHip_IntMor = p_CreHip(1).CreHip_IntMor + CDbl(Format(r_rst_Genera!HIPCUO_INTMOR - r_rst_Genera!HIPCUO_IMOPAG, "#######0.00"))
            p_CreHip(1).CreHip_IntCom = p_CreHip(1).CreHip_IntCom + CDbl(Format(r_rst_Genera!HIPCUO_INTCOM - r_rst_Genera!HIPCUO_ICOPAG, "#######0.00"))
            p_CreHip(1).CreHip_GasCob = p_CreHip(1).CreHip_GasCob + CDbl(Format(r_rst_Genera!HIPCUO_GASCOB - r_rst_Genera!HIPCUO_GCOPAG, "#######0.00"))
            p_CreHip(1).CreHip_OtrGas = p_CreHip(1).CreHip_OtrGas + CDbl(Format(r_rst_Genera!HIPCUO_OTRGAS - r_rst_Genera!HIPCUO_OTGPAG, "#######0.00"))
            
            r_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If
End Sub

Public Sub modprc_gs_DevengadoPBP_CreHip(p_CreHip() As modprc_g_tpo_CreHip, p_LogPro() As modprc_g_tpo_LogPro, ByVal p_FecPro As String)
Dim r_rst_Genera     As ADODB.Recordset
Dim r_int_Entero     As Integer
Dim r_int_Decima     As Integer
Dim r_str_TipCro     As String

   p_CreHip(1).CreHip_NCuoTC = 0
   p_CreHip(1).CreHip_CapiTC = 0
   p_CreHip(1).CreHip_InteTC = 0
   p_CreHip(1).CreHip_ComiTC = 0
   p_CreHip(1).CreHip_DevPBP = 0

   'Para determinar que Cuota TC corresponde según el Cronograma del TNC
   modprc_g_str_CadEje = "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO " & _
                         " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                         "   AND HIPCUO_TIPCRO = 1 " & _
                         "   AND HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & " " & _
                         " ORDER BY HIPCUO_NUMCUO DESC "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
      Exit Sub
   End If

   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      r_int_Entero = (r_rst_Genera!HIPCUO_NUMCUO \ 6) + 1
      r_int_Decima = r_rst_Genera!HIPCUO_NUMCUO Mod 6
      
      If r_int_Decima = 0 Then
         p_CreHip(1).CreHip_NCuoTC = r_int_Entero - 1
      Else
         p_CreHip(1).CreHip_NCuoTC = r_int_Entero
      End If
   
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
         
      'Para obtener Interés, Comisiones y Saldo de Cuota TC
      If Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "006" Then
         r_str_TipCro = 2
      Else
         r_str_TipCro = 4
      End If
      
      modprc_g_str_CadEje = "SELECT HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF FROM CRE_HIPCUO " & _
                            " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                            "   AND HIPCUO_TIPCRO = " & CStr(r_str_TipCro) & " " & _
                            "   AND HIPCUO_NUMCUO = " & CStr(p_CreHip(1).CreHip_NCuoTC) & " " & _
                            " ORDER BY HIPCUO_NUMCUO DESC "
         
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
         p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
         Exit Sub
      End If
   
      r_rst_Genera.MoveFirst
      p_CreHip(1).CreHip_CapiTC = r_rst_Genera!HIPCUO_CAPITA
      p_CreHip(1).CreHip_InteTC = r_rst_Genera!HIPCUO_INTERE
      p_CreHip(1).CreHip_ComiTC = r_rst_Genera!HIPCUO_COMCOF
      p_CreHip(1).CreHip_DevPBP = CDbl(Format(r_rst_Genera!HIPCUO_INTERE / 6, "###,###,##0.00"))
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modprc_gs_ProvisionAdeudado_CreHip(p_CreHip() As modprc_g_tpo_CreHip, p_LogPro() As modprc_g_tpo_LogPro, ByVal p_FecPro As String, ByVal p_TasCof As Double, ByVal p_ComCof As Double)
Dim r_rst_Genera     As ADODB.Recordset
Dim r_str_TipCro     As String
Dim r_int_DifDia     As Integer

   p_CreHip(1).CreHip_PrvICo = 0
   p_CreHip(1).CreHip_PrvCCo = 0
   p_CreHip(1).CreHip_MtoAde = 0
   p_CreHip(1).CreHip_IniAde = "0"

   If Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "003" Then
      r_str_TipCro = 5
   ElseIf Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "004" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "007" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "009" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "010" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "013" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "014" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "015" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "016" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "017" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "018" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "019" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "021" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "022" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "023" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "024" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "025" Then
      r_str_TipCro = 3
   End If

   'Para obtener Saldo Capital y Fecha de Vencimiento correspondiente a Cuota a Pagar en el mes
   modprc_g_str_CadEje = "SELECT HIPCUO_SALCAP, HIPCUO_FECVCT FROM CRE_HIPCUO " & _
                         " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                         "   AND HIPCUO_TIPCRO = " & CStr(r_str_TipCro) & " " & _
                         "   AND HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & " " & _
                         " ORDER BY HIPCUO_NUMCUO DESC "
      
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      r_int_DifDia = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT))))
      p_CreHip(1).CreHip_PrvICo = CDbl(Format(r_rst_Genera!HIPCUO_SALCAP * (1 + p_TasCof / 100) ^ (r_int_DifDia / 360) - r_rst_Genera!HIPCUO_SALCAP, "###,###,##0.00"))
      p_CreHip(1).CreHip_PrvCCo = CDbl(Format(r_rst_Genera!HIPCUO_SALCAP * (1 + p_ComCof / 100) ^ (r_int_DifDia / 360) - r_rst_Genera!HIPCUO_SALCAP, "###,###,##0.00"))
      p_CreHip(1).CreHip_MtoAde = r_rst_Genera!HIPCUO_SALCAP
      p_CreHip(1).CreHip_IniAde = CStr(r_rst_Genera!HIPCUO_FECVCT)
   Else
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   
      modprc_g_str_CadEje = "SELECT HIPMAE_IMPNCO, HIPMAE_IMPCON, EVACOF_FECDES " & _
                            "  FROM CRE_HIPMAE A, TRA_EVACOF B " & _
                            " WHERE HIPMAE_NUMSOL = EVACOF_NUMSOL " & _
                            "   AND HIPMAE_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
         p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
         Exit Sub
      End If
      
      r_rst_Genera.MoveFirst
      r_int_DifDia = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(r_rst_Genera!EVACOF_FECDES))))
      p_CreHip(1).CreHip_PrvICo = CDbl(Format(((r_rst_Genera!HIPMAE_IMPNCO + r_rst_Genera!HIPMAE_IMPCON) * (1 + p_TasCof / 100) ^ (r_int_DifDia / 360) - (r_rst_Genera!HIPMAE_IMPNCO + r_rst_Genera!HIPMAE_IMPCON)), "###,###,##0.00"))
      p_CreHip(1).CreHip_PrvCCo = CDbl(Format(((r_rst_Genera!HIPMAE_IMPNCO + r_rst_Genera!HIPMAE_IMPCON) * (1 + p_ComCof / 100) ^ (r_int_DifDia / 360) - (r_rst_Genera!HIPMAE_IMPNCO + r_rst_Genera!HIPMAE_IMPCON)), "###,###,##0.00"))
      p_CreHip(1).CreHip_MtoAde = r_rst_Genera!HIPMAE_IMPNCO + r_rst_Genera!HIPMAE_IMPCON
      p_CreHip(1).CreHip_IniAde = CStr(r_rst_Genera!EVACOF_FECDES)
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modprc_gs_UltimoCapPago_CreHip(p_CreHip() As modprc_g_tpo_CreHip, p_LogPro() As modprc_g_tpo_LogPro)
   Dim r_rst_Genera     As ADODB.Recordset
   p_CreHip(1).CreHip_UltCap = 0

   'Para obtener Saldo Capital y Fecha de Vencimiento correspondiente a Cuota a Pagar en el mes
   modprc_g_str_CadEje = "SELECT HIPCUO_CAPPAG FROM CRE_HIPCUO " & _
                         " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                         "   AND HIPCUO_TIPCRO = 1 " & _
                         "   AND HIPCUO_SITUAC = 1 " & _
                         "   AND HIPCUO_CAPPAG > 0 " & _
                         " ORDER BY HIPCUO_NUMCUO DESC "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      p_CreHip(1).CreHip_UltCap = r_rst_Genera!HIPCUO_CAPPAG
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modprc_gs_Devengado_CreHip(p_CreHip() As modprc_g_tpo_CreHip, p_LogPro() As modprc_g_tpo_LogPro, ByVal p_FecIni As String, ByVal p_FecPro As String, ByVal p_TasInt As Double)
Dim r_rst_Genera     As ADODB.Recordset
Dim r_rst_Gener1     As ADODB.Recordset
Dim r_int_DifDi1     As Integer
Dim r_dbl_IntDv1     As Double
Dim r_int_DifDi2     As Integer
Dim r_dbl_IntDv2     As Double
Dim r_int_DifDi3     As Integer
Dim r_dbl_IntDv3     As Double
Dim r_dbl_MtoPre     As Double
   
   p_CreHip(1).CreHip_CuoDev = 0
   p_CreHip(1).CreHip_IntDev = 0
   r_dbl_IntDv1 = 0
   r_dbl_IntDv2 = 0
   r_dbl_IntDv3 = 0
   r_int_DifDi1 = 0
   r_int_DifDi2 = 0
   r_int_DifDi3 = 0

   If Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "002" Or Mid(p_CreHip(1).CreHip_NumOpe, 1, 3) = "011" Then
      r_dbl_MtoPre = p_CreHip(1).CreHip_MtoPre
   Else
      r_dbl_MtoPre = p_CreHip(1).CreHip_MtoNCo
   End If

   'Para obtener Saldo Capital y Fecha de Vencimiento correspondiente a Cuota a Pagar en el mes
   modprc_g_str_CadEje = "SELECT HIPCUO_NUMCUO, HIPCUO_SALCAP, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_INTPAG, HIPCUO_FECVCT, HIPCUO_FECPAG " & _
                         "  FROM CRE_HIPCUO " & _
                         " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                         "   AND HIPCUO_TIPCRO = 1 " & _
                         "   AND HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & " " & _
                         " ORDER BY HIPCUO_NUMCUO DESC "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      p_CreHip(1).CreHip_CuoDev = r_rst_Genera!HIPCUO_NUMCUO + 1
      r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT))))
      r_dbl_IntDv1 = CDbl(Format(r_rst_Genera!HIPCUO_SALCAP * (1 + p_TasInt / 100) ^ (r_int_DifDi1 / 360) - r_rst_Genera!HIPCUO_SALCAP, "###,###,##0.00"))
      
      'Si Cliente no pago cuota
      If r_rst_Genera!HIPCUO_INTERE - r_rst_Genera!HIPCUO_INTPAG <> 0 Then
         modprc_g_str_CadEje = "SELECT HIPCUO_SALCAP, HIPCUO_INTERE, HIPCUO_INTPAG, HIPCUO_FECVCT, HIPCUO_FECPAG " & _
                               "  FROM CRE_HIPCUO " & _
                               " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                               "   AND HIPCUO_TIPCRO = 1 " & _
                               "   AND HIPCUO_NUMCUO = " & CStr((r_rst_Genera!HIPCUO_NUMCUO - 1)) & " " & _
                               " ORDER BY HIPCUO_NUMCUO DESC "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Gener1, 3) Then
            p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
            Exit Sub
         End If
         
         If Not (r_rst_Gener1.BOF And r_rst_Gener1.EOF) Then
            If CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))) >= CDate(p_FecIni) Then     'Si hubo Pre-Pago en Mes Actual
               r_int_DifDi2 = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT))) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))))
               r_dbl_IntDv2 = CDbl(Format((r_rst_Genera!HIPCUO_SALCAP + r_rst_Genera!HIPCUO_CAPITA) * (1 + p_TasInt / 100) ^ (r_int_DifDi2 / 360) - (r_rst_Genera!HIPCUO_SALCAP + r_rst_Genera!HIPCUO_CAPITA), "###,###,##0.00"))
            Else
               If CLng(p_CreHip(1).CreHip_FecPPg) > 0 Then
                  If CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_FecPPg))) > CDate(gf_FormatoFecha(CStr(r_rst_Gener1!HIPCUO_FECVCT))) And CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_FecPPg))) < CDate(p_FecIni) Then
                     r_dbl_IntDv2 = r_rst_Genera!HIPCUO_INTERE - r_rst_Genera!HIPCUO_INTPAG
                  Else
                     r_int_DifDi3 = CInt(CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))) - CDate(gf_FormatoFecha(CStr(r_rst_Gener1!HIPCUO_FECVCT))))
                     r_dbl_IntDv3 = CDbl(Format(r_rst_Gener1!HIPCUO_SALCAP * (1 + p_TasInt / 100) ^ (r_int_DifDi3 / 360) - r_rst_Gener1!HIPCUO_SALCAP, "###,###,##0.00"))
                     
                     r_int_DifDi2 = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT))) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))))
                     r_dbl_IntDv2 = CDbl(Format((r_rst_Gener1!HIPCUO_SALCAP + r_dbl_IntDv3) * (1 + p_TasInt / 100) ^ (r_int_DifDi2 / 360) - (r_rst_Gener1!HIPCUO_SALCAP + r_dbl_IntDv3), "###,###,##0.00"))
                  End If
               Else
                  r_int_DifDi3 = CInt(CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))) - CDate(gf_FormatoFecha(CStr(r_rst_Gener1!HIPCUO_FECVCT))))
                  r_dbl_IntDv3 = CDbl(Format(r_rst_Gener1!HIPCUO_SALCAP * (1 + p_TasInt / 100) ^ (r_int_DifDi3 / 360) - r_rst_Gener1!HIPCUO_SALCAP, "###,###,##0.00"))
                  
                  r_int_DifDi2 = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT))) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))))
                  r_dbl_IntDv2 = CDbl(Format((r_rst_Gener1!HIPCUO_SALCAP + r_dbl_IntDv3) * (1 + p_TasInt / 100) ^ (r_int_DifDi2 / 360) - (r_rst_Gener1!HIPCUO_SALCAP + r_dbl_IntDv3), "###,###,##0.00"))
               End If
            End If
         Else
            'Si no hay Cuota Anterior se considera Fecha de Desembolso
            r_int_DifDi3 = CInt(CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_FecDes))))
            r_dbl_IntDv3 = CDbl(Format(r_dbl_MtoPre * (1 + p_TasInt / 100) ^ (r_int_DifDi3 / 360) - r_dbl_MtoPre, "###,###,##0.00"))
            
            r_int_DifDi2 = CInt(CDate(gf_FormatoFecha(CStr(r_rst_Genera!HIPCUO_FECVCT))) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))))
            r_dbl_IntDv2 = CDbl(Format((r_dbl_MtoPre + r_dbl_IntDv3) * (1 + p_TasInt / 100) ^ (r_int_DifDi2 / 360) - (r_dbl_MtoPre + r_dbl_IntDv3), "###,###,##0.00"))
         End If
         
         r_rst_Gener1.Close
         Set r_rst_Gener1 = Nothing
      Else
         If p_CreHip(1).CreHip_DevAnt > 0 Then
            If CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))) >= CDate(p_FecIni) Then        'Si hubo Pre-Pago en Mes Actual
               modprc_g_str_CadEje = "SELECT HIPCUO_SALCAP, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_INTPAG, HIPCUO_FECVCT, HIPCUO_FECPAG " & _
                                     "  FROM CRE_HIPCUO " & _
                                     " WHERE HIPCUO_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' " & _
                                     "   AND HIPCUO_TIPCRO = 1 " & _
                                     "   AND HIPCUO_NUMCUO = " & CStr(p_CreHip(1).CreHip_CuoDev) & " " & _
                                     " ORDER BY HIPCUO_NUMCUO DESC "
               
               If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Gener1, 3) Then
                  p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
                  Exit Sub
               End If
               
               r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))))
               r_dbl_IntDv1 = CDbl(Format((r_rst_Gener1!HIPCUO_SALCAP + r_rst_Gener1!HIPCUO_CAPITA) * (1 + p_TasInt / 100) ^ (r_int_DifDi1 / 360) - (r_rst_Gener1!HIPCUO_SALCAP + r_rst_Gener1!HIPCUO_CAPITA), "###,###,##0.00"))
               
               r_rst_Gener1.Close
               Set r_rst_Gener1 = Nothing
            End If
         End If
      End If
   Else
      p_CreHip(1).CreHip_CuoDev = 1
      
      'If p_CreHip(1).CreHip_DevAnt = "0" Then
         r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_FecDes))))
      'Else
      '   r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(CStr(p_CreHip(1).CreHip_DevAnt))))
      'End If
      
      r_dbl_IntDv1 = CDbl(Format(r_dbl_MtoPre * (1 + p_TasInt / 100) ^ (r_int_DifDi1 / 360) - r_dbl_MtoPre, "###,###,##0.00"))
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   p_CreHip(1).CreHip_IntDev = r_dbl_IntDv1 + r_dbl_IntDv2
End Sub

Public Sub modprc_gs_Devengado_CreHip2(p_CreHip() As modprc_g_tpo_CreHip, p_LogPro() As modprc_g_tpo_LogPro, ByVal p_FecIni As String, ByVal p_FecPro As String, ByVal p_TasInt As Double, ByVal p_DiaAtr As Integer, ByRef p_AcuDvg As Double)
Dim r_rst_Genera     As ADODB.Recordset
Dim r_dbl_IntDv1     As Double
Dim r_int_DifDi1     As Double
Dim r_str_FecPpg     As String

   p_CreHip(1).CreHip_IntDev = 0
   r_str_FecPpg = "0"
   r_dbl_IntDv1 = 0
   r_int_DifDi1 = 0
   p_AcuDvg = 0
   
   'Obtiene fecha de prepago (si hubo)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT NVL(A.PPGCAB_FECPPG, '0') AS FECHA_PREPAGO "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CRE_PPGCAB A "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE A.PPGCAB_NUMOPE = '" & p_CreHip(1).CreHip_NumOpe & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND A.PPGCAB_FECPPG >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND A.PPGCAB_FECPPG <= " & Format(CDate(p_FecPro), "yyyymmdd") & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " ORDER BY A.PPGCAB_FECPPG DESC "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      r_str_FecPpg = CStr(r_rst_Genera!FECHA_PREPAGO)
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   'Calculo del devengado
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT CASE WHEN (SELECT B.HIPCUO_INTERE FROM CRE_HIPCUO B "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE B.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND B.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND B.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND B.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") IS NOT NULL "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            THEN (SELECT C.HIPCUO_INTERE FROM CRE_HIPCUO C "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE C.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND C.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND C.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND C.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            ELSE 0 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       END AS INTERE_CUOTA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       CASE WHEN (SELECT D.HIPCUO_SALCAP FROM CRE_HIPCUO D "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE D.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND D.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND D.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND D.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") IS NOT NULL "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            THEN (SELECT E.HIPCUO_SALCAP FROM CRE_HIPCUO E "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE E.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND E.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND E.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND E.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            ELSE A.HIPMAE_SALCAP "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       END AS SALCAP_CUOTA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       CASE WHEN (SELECT F.HIPCUO_FECVCT FROM CRE_HIPCUO F "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE F.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND F.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND F.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND F.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") IS NOT NULL "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            THEN "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                 (SELECT SUBSTR(G.HIPCUO_FECVCT,7,2)||'/'||SUBSTR(G.HIPCUO_FECVCT,5,2)||'/'||SUBSTR(G.HIPCUO_FECVCT,1,4) FROM CRE_HIPCUO G "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE G.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND G.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND G.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND G.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            ELSE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                 SUBSTR(A.HIPMAE_FECDES,7,2)||'/'||SUBSTR(A.HIPMAE_FECDES,5,2)||'/'||SUBSTR(A.HIPMAE_FECDES,1,4) "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       END AS VCTO_CUOTA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       CASE WHEN (SELECT H.HIPCUO_SITUAC FROM CRE_HIPCUO H "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE H.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND H.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND H.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND H.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") IS NOT NULL "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            THEN (SELECT I.HIPCUO_SITUAC FROM CRE_HIPCUO I "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                   WHERE I.HIPCUO_NUMOPE = A.HIPMAE_NUMOPE AND I.HIPCUO_TIPCRO = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "                     AND I.HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND I.HIPCUO_FECVCT <= " & Format(CDate(p_FecPro), "yyyymmdd") & ") "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "            ELSE 2 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       END AS SITUAC_CUOTA "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CRE_HIPMAE A "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE A.HIPMAE_NUMOPE  = '" & p_CreHip(1).CreHip_NumOpe & "' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & p_CreHip(1).CreHip_NumOpe & " .")
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      If r_str_FecPpg = "0" Then
         r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(CStr(r_rst_Genera!VCTO_CUOTA)))
      Else
         If CStr(r_str_FecPpg) > CStr(r_rst_Genera!VCTO_CUOTA) Then
            r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(gf_FormatoFecha(r_str_FecPpg)))
         Else
            r_int_DifDi1 = CInt(CDate(p_FecPro) - CDate(CStr(r_rst_Genera!VCTO_CUOTA)))
         End If
      End If
      r_dbl_IntDv1 = CDbl(Format(r_rst_Genera!SALCAP_CUOTA * (1 + p_TasInt / 100) ^ (r_int_DifDi1 / 360) - r_rst_Genera!SALCAP_CUOTA, "###,###,##0.00"))
      
      If p_DiaAtr < 31 Then
         p_CreHip(1).CreHip_IntDev = r_dbl_IntDv1
         If r_rst_Genera!SITUAC_CUOTA = 2 Then
            p_AcuDvg = r_dbl_IntDv1 + r_rst_Genera!INTERE_CUOTA
         Else
            p_AcuDvg = r_dbl_IntDv1
         End If
      End If
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Function modprc_gf_ClaCred(p_TipCla() As modprc_g_tpo_Genera, ByVal p_DiaAtr As Integer) As Integer
Dim r_int_Contad        As Integer
   
   modprc_gf_ClaCred = -1

   For r_int_Contad = 1 To UBound(p_TipCla)
      If p_DiaAtr >= p_TipCla(r_int_Contad).Genera_DiaVc1 And p_DiaAtr <= p_TipCla(r_int_Contad).Genera_DiaVc2 Then
         modprc_gf_ClaCred = CInt(p_TipCla(r_int_Contad).Genera_Codigo)
         Exit For
      End If
   Next r_int_Contad
End Function

Public Function modprc_gf_PorcenProv(p_Arregl() As modprc_g_tpo_TipPrv, ByVal p_TipPrv As Integer, ByVal p_ClaCre As Integer, ByVal p_ClaGar As Integer) As Double
Dim r_int_Contad     As Integer

   modprc_gf_PorcenProv = 0
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      If p_Arregl(r_int_Contad).TipPrv_TipPrv = p_TipPrv And p_Arregl(r_int_Contad).TipPrv_CodCla = p_ClaCre And p_Arregl(r_int_Contad).TipPrv_ClaGar = p_ClaGar Then
         modprc_gf_PorcenProv = p_Arregl(r_int_Contad).TipPrv_Porcen
         Exit For
      End If
   Next r_int_Contad
End Function

Private Sub fs_Inserta_RCCCab(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecPro As String, ByVal p_CodSbs As String, _
                              ByVal p_NumEmp As Integer, ByVal p_DeuNor As Double, ByVal p_DeuCpp As Double, ByVal p_DeuDef As Double, ByVal p_DeuDud As Double, ByVal p_DeuPer As Double)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CLI_RCCCAB ("
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ","
      g_str_Parame = g_str_Parame & "'" & p_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
      g_str_Parame = g_str_Parame & p_FecPro & ", "
      g_str_Parame = g_str_Parame & "'" & p_CodSbs & "', "
      g_str_Parame = g_str_Parame & CStr(p_NumEmp) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DeuNor) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DeuCpp) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DeuDef) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DeuDud) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DeuPer) & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Sub fs_Inserta_RCCDet(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_NumIte As Integer, ByVal p_CodEmp As String, _
                              ByVal p_DiaAtr As Integer, ByVal p_MonDeu As Integer, ByVal p_MtoSol As Double, ByVal p_MtoDol As Double, ByVal p_Clasif As Integer, ByVal p_TipDeu As Integer, ByVal p_CtaCtb As String)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CLI_RCCDET ("
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ","
      g_str_Parame = g_str_Parame & "'" & p_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
      g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
      g_str_Parame = g_str_Parame & p_CodEmp & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipDeu) & ", "
      g_str_Parame = g_str_Parame & CStr(p_DiaAtr) & ", "
      g_str_Parame = g_str_Parame & CStr(p_MonDeu) & ", "
      g_str_Parame = g_str_Parame & CStr(p_MtoSol) & ", "
      g_str_Parame = g_str_Parame & CStr(p_MtoDol) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Clasif) & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(p_CtaCtb) & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Function ff_Valida_RCCCab(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_PerMes As Integer, ByVal p_PerAno) As Integer
   ff_Valida_RCCCab = True
   
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM CLI_RCCCAB "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE RCCCAB_TIPDOC = " & CStr(p_TipDoc) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND RCCCAB_NUMDOC = '" & p_NumDoc & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND RCCCAB_PERMES = " & CStr(p_PerMes) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND RCCCAB_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      ff_Valida_RCCCab = False
   End If
   
   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
End Function

Public Function ff_Actualiza_CodigoSBS(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_CodSbs As String, ByVal p_TipPer As Integer) As Integer
Dim r_int_ActSBS           As Integer
   
   ff_Actualiza_CodigoSBS = False
   r_int_ActSBS = 2
   
   If p_TipPer = 1 Then
      'Persona Natural
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM CLI_DATGEN "
      modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DATGEN_TIPDOC = " & CStr(p_TipDoc) & " "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DATGEN_NUMDOC = '" & p_NumDoc & "' "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
         Exit Function
      End If
      
      If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
         modprc_g_rst_Genera.MoveFirst
         If Len(Trim(modprc_g_rst_Genera!DATGEN_CODSBS & "")) = 0 Then
            r_int_ActSBS = 1
         End If
      End If
               
      modprc_g_rst_Genera.Close
      Set modprc_g_rst_Genera = Nothing
      
   ElseIf p_TipPer = 4 Then
      'ET
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM TPR_MAEETE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MAEETE_TIPDOC = " & CStr(p_TipDoc) & " "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAEETE_NUMDOC = '" & p_NumDoc & "' "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
         Exit Function
      End If
      
      If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
         modprc_g_rst_Genera.MoveFirst
         If Len(Trim(modprc_g_rst_Genera!MAEETE_CODSBS & "")) = 0 Then
            r_int_ActSBS = 1
         End If
      End If
      
      modprc_g_rst_Genera.Close
      Set modprc_g_rst_Genera = Nothing
      
   Else
      'Persona Juridica
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM EMP_DATGEN "
      modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DATGEN_EMPTDO = " & CStr(p_TipDoc) & " "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DATGEN_EMPNDO = '" & p_NumDoc & "' "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
         Exit Function
      End If
      
      If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
         modprc_g_rst_Genera.MoveFirst
         If Len(Trim(modprc_g_rst_Genera!DATGEN_CODSBS & "")) = 0 Then
            r_int_ActSBS = 1
         End If
      End If
   
      modprc_g_rst_Genera.Close
      Set modprc_g_rst_Genera = Nothing
   End If
                  
   'Para actualizar Código SBS
   If r_int_ActSBS = 1 Then
      If p_TipPer = 1 Then
         'Persona Natural
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CLI_DATGEN SET DATGEN_CODSBS = '" & p_CodSbs & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DATGEN_TIPDOC = " & CStr(p_TipDoc) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DATGEN_NUMDOC = '" & p_NumDoc & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 2) Then
            Exit Function
         End If
      
      ElseIf p_TipPer = 4 Then
         'ET
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TPR_MAEETE SET MAEETE_CODSBS = '" & p_CodSbs & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MAEETE_TIPDOC = " & CStr(p_TipDoc) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAEETE_NUMDOC = '" & p_NumDoc & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 2) Then
            Exit Function
         End If
         
      Else
         'Persona Juridica
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE EMP_DATGEN SET DATGEN_CODSBS = '" & p_CodSbs & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DATGEN_EMPTDO = " & CStr(p_TipDoc) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DATGEN_EMPNDO = '" & p_NumDoc & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 2) Then
            Exit Function
         End If
      End If
   End If
   
   ff_Actualiza_CodigoSBS = True
End Function

Public Sub modprc_infp8001(p_BarPro As SSPanel, p_NumPro As SSPanel, p_NumErr As SSPanel, ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String)
'Código Proceso   :  INFP8001
'Descripción      :  Nivelación de Cuotas Devengadas y Cuotas Pagadas en cre_hipmae
'F. Creación      :  22-Ene-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg  As Long
Dim r_lng_TotReg  As Long
Dim r_lng_NumErr  As Long
Dim r_str_FFnEje  As String
Dim r_str_HFnEje  As String
Dim r_int_MesDes  As Integer
Dim r_int_AnoDes  As Integer
Dim r_int_MesFin  As Integer
Dim r_int_AnoFin  As Integer
Dim r_int_CuoDev  As Integer
Dim r_int_CuoPag  As Integer
Dim r_int_DiaVc1  As Integer
Dim r_int_DiaVc2  As Integer
   
   p_BarPro.FloodPercent = 0
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_lng_NumErr = 0
   
   'Para determinar Total de Registros a Procesar
   modprc_g_str_CadEje = "SELECT COUNT(*) AS TOTREG FROM CRE_HIPMAE WHERE HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   If r_lng_TotReg > 0 Then
      r_int_MesFin = Month(date)
      r_int_AnoFin = Year(date)
      
      modprc_g_str_CadEje = "SELECT * FROM CRE_HIPMAE WHERE HIPMAE_SITUAC = 2 "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      Do While Not modprc_g_rst_Princi.EOF
         r_int_MesDes = Mid(CStr(modprc_g_rst_Princi!HIPMAE_FECDES), 5, 2)
         r_int_AnoDes = Mid(CStr(modprc_g_rst_Princi!HIPMAE_FECDES), 1, 4)
         r_int_CuoDev = 0
         r_int_CuoPag = 0
         
         'Contando Nro. de Procesos de Devengados Mensuales
         Do While Format(r_int_AnoDes, "0000") & Format(r_int_MesDes, "00") < Format(r_int_AnoFin, "0000") & Format(r_int_MesFin, "00")
            r_int_CuoDev = r_int_CuoDev + 1
            r_int_MesDes = r_int_MesDes + 1
            
            If r_int_MesDes = 13 Then
               r_int_AnoDes = r_int_AnoDes + 1
               r_int_MesDes = 1
            End If
         Loop
         
         'Obteniendo Cuotas Pagadas
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT COUNT(*) AS CUOPAG FROM CRE_HIPCUO "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCUO_NUMOPE = '" & modprc_g_rst_Princi!HIPMAE_NUMOPE & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_TIPCRO = 1 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_SITUAC = 1 "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
             Exit Sub
         End If
         
         modprc_g_rst_Genera.MoveFirst
         r_int_CuoPag = modprc_g_rst_Genera!CUOPAG
         
         modprc_g_rst_Genera.Close
         Set modprc_g_rst_Genera = Nothing
         
         'Grabando en CRE_HIPMAE
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_CUODEV = " & CStr(r_int_CuoDev) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_CUOPAG = " & CStr(r_int_CuoPag) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPMAE_NUMOPE & "' "
      
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
             Exit Sub
         End If
      
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         p_NumPro.Caption = CStr(r_lng_NumReg) & " "
         p_NumErr.Caption = CStr(r_lng_NumErr) & " "
         
         DoEvents
      Loop
   
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Else
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontraron datos en la tabla CRE_HIPMAE.")
   End If
   
   If r_lng_TotReg > 0 Then
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   End If
   
   p_NumPro.Caption = CStr(r_lng_NumReg) & " "
   p_NumErr.Caption = CStr(r_lng_NumErr) & " "
   DoEvents
   
   r_str_FFnEje = Format(date, "yyyymmdd")
   r_str_HFnEje = Format(Time, "hhmmss")
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_NumReg, r_lng_NumErr, "", "", 0, 0, 0, 0, 0)
End Sub

Public Sub modprc_cbrp5001(ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String, ByRef p_NumErr As Long)
'Código Proceso   :  CBRP5001
'Descripción      :  Generación de Cartera de Cobranzas
'F. Creación      :  06-Feb-09
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg     As Long
Dim r_lng_TotReg     As Long
Dim r_lng_NumErr     As Long
Dim r_str_FFnEje     As String
Dim r_str_HFnEje     As String
Dim r_str_FecPro     As String
Dim r_str_NumOpe     As String
Dim r_str_CodPrd     As String
Dim r_str_CodSub     As String
Dim r_arr_ParSub()   As modprc_g_tpo_Genera
Dim r_int_CuoAtr     As Integer
Dim r_int_AtrMax     As Integer
Dim r_int_DiaMor     As Integer
Dim r_int_DiaAtr     As Integer
Dim r_dbl_TasMor     As Double
Dim r_dbl_IntMor     As Double
Dim r_dbl_GasCob     As Double
   
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_lng_NumErr = 0
   
   'Fecha de Proceso
   'r_str_FecPro = Format(Date, "dd/mm/yyyy")
   r_str_FecPro = Format("31/10/2009", "dd/mm/yyyy")
   
   'Inicializando Días de Mora
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_DIAMOR = 0, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_CUOATR = 0 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_SITUAC = 2"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Inicializar Días de Mora en tabla CRE_HIPMAE.")
      
      r_str_FFnEje = Format(date, "yyyymmdd")
      r_str_HFnEje = Format(Time, "hhmmss")
      Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_NumReg, r_lng_NumErr, "000001", "", 0, 0, 0, Format(CDate(r_str_FecPro), "yyyymmdd"), Format(CDate(r_str_FecPro), "yyyymmdd"))
      Exit Sub
   End If
   
   'Leyendo Cursor Principal
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM CRE_HIPMAE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_PRXVCT < " & Format(CDate(r_str_FecPro), "yyyymmdd") & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         r_str_NumOpe = modprc_g_rst_Princi!HIPMAE_NUMOPE
         r_str_CodPrd = modprc_g_rst_Princi!HIPMAE_CODPRD
         r_str_CodSub = modprc_g_rst_Princi!HIPMAE_CODSUB
         
         r_int_AtrMax = CInt(CDate(r_str_FecPro) - CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!HIPMAE_PRXVCT))))
         r_int_CuoAtr = 0
         
         'Obteniendo Tasa de Interés Moratorio
         r_dbl_TasMor = 0
         
         If modprc_gf_Consulta_ParametroSubPrd(r_arr_ParSub, r_str_CodPrd, r_str_CodSub, "002", "201") Then
            r_dbl_TasMor = r_arr_ParSub(1).Genera_Cantid
         End If
         
         'Convirtiendo Tasa Moratoria Anual a Mensual
         r_dbl_TasMor = modprc_gf_InteresMensual(r_dbl_TasMor)
         
         'Obteniendo Día de Inicio de Aplicación de Interes Moratorio
         r_int_DiaMor = 0
         
         If modprc_gf_Consulta_ParametroSubPrd(r_arr_ParSub, r_str_CodPrd, r_str_CodSub, "002", "202") Then
            r_int_DiaMor = r_arr_ParSub(1).Genera_Cantid
         End If
         
         'Buscando Cuotas Atrasadas
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM CRE_HIPCUO "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCUO_NUMOPE = '" & r_str_NumOpe & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_FECVCT < " & Format(CDate(r_str_FecPro), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_SITUAC = 2 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_TIPCRO = 1 "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
            r_lng_NumErr = r_lng_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPCUO (Operación: " & r_str_NumOpe & "). ")
         End If
         
         If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
            modprc_g_rst_Genera.MoveFirst
            
            Do While Not modprc_g_rst_Genera.EOF
               r_int_CuoAtr = r_int_CuoAtr + 1
               r_dbl_IntMor = 0
               r_dbl_GasCob = 0
               r_int_DiaAtr = CInt(CDate(r_str_FecPro) - CDate(gf_FormatoFecha(CStr(modprc_g_rst_Genera!HIPCUO_FECVCT))))
            
               'Calculando Interes Moratorio
               If r_int_DiaAtr >= r_int_DiaMor Then
                  r_dbl_IntMor = modprc_gf_CalculaInteres(r_dbl_TasMor, r_int_DiaAtr, modprc_g_rst_Genera!HIPCUO_CAPITA)
               End If
            
               'Leyendo Gastos de Cobranzas según días de atraso
               modprc_g_str_CadEje = ""
               modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM OPE_GASCOB "
               modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE GASCOB_CODPRD = '" & r_str_CodPrd & "' "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASCOB_CODSUB = '" & r_str_CodSub & "' "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASCOB_DIAINI <= " & CStr(r_int_DiaAtr) & " "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASCOB_DIAFIN >= " & CStr(r_int_DiaAtr) & " "
               
               If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Auxili, 3) Then
                  r_lng_NumErr = r_lng_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer OPE_GASCOB (Operación: " & r_str_NumOpe & "). ")
               End If
               
               If Not (modprc_g_rst_Auxili.BOF And modprc_g_rst_Auxili.EOF) Then
                  r_dbl_GasCob = modprc_g_rst_Auxili!GasCob_Import
               End If
               
               modprc_g_rst_Auxili.Close
               Set modprc_g_rst_Auxili = Nothing
               
               'Actualizando en CRE_HIPCUO
               modprc_g_str_CadEje = ""
               modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPCUO "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPCUO_INTMOR = " & Format(r_dbl_IntMor, "#######0.00") & ", "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPCUO_GASCOB = " & Format(r_dbl_GasCob, "#######0.00") & " "
               modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCUO_NUMOPE = '" & r_str_NumOpe & "' "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_TIPCRO = 1 "
               modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCUO_NUMCUO = " & CStr(modprc_g_rst_Genera!HIPCUO_NUMCUO) & " "
            
               If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
                  r_lng_NumErr = r_lng_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al actualizar en CRE_HIPCUO (Operación: " & r_str_NumOpe & " Cuota: " & CStr(modprc_g_rst_Genera!HIPCUO_NUMCUO) & ").")
               End If
            
               modprc_g_rst_Genera.MoveNext
               DoEvents
            Loop
         End If
      
         modprc_g_rst_Genera.Close
         Set modprc_g_rst_Genera = Nothing
      
         'Actualizando en CRE_HIPMAE
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_DIAMOR = " & CStr(r_int_AtrMax) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_CUOATR = " & CStr(r_int_CuoAtr) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & r_str_NumOpe & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
            r_lng_NumErr = r_lng_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al actualizar en CRE_HIPMAE (Operación: " & r_str_NumOpe & " Cuota: " & CStr(modprc_g_rst_Genera!HIPCUO_NUMCUO) & ").")
         End If
      
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
      Loop
   Else
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontro registros en tabla CRE_HIPMAE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   r_str_FFnEje = Format(date, "yyyymmdd")
   r_str_HFnEje = Format(Time, "hhmmss")
   p_NumErr = r_lng_NumErr
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_NumReg, r_lng_NumErr, "000001", "", 0, 0, 0, Format(CDate(r_str_FecPro), "yyyymmdd"), Format(CDate(r_str_FecPro), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1002(ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String, ByRef p_NumErr As Long, ByVal p_CodEmp As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1002
'Descripción      :  Cierre de Operaciones Mensuales - Créditos Hipotecarios (Elminiación de Proceso)
'Resumen          :  Proceso que elimina Movimientos de la tabla CRE_HIPCIE
'F. Creación      :  22-04-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg     As Long
Dim r_lng_TotReg     As Long
Dim r_lng_NumErr     As Long
Dim r_str_FFnEje     As String
Dim r_str_HFnEje     As String
Dim r_str_FecPro     As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer
Dim r_str_Period        As String
   
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_lng_NumErr = 0
   p_BarPro.FloodPercent = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Período Vigente
   r_str_Period = moddat_gf_ConsultaPerMesActivo(p_CodEmp, 1, r_str_FecIni, r_str_FecFin, r_int_PerMes, r_int_PerAno)
   
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DELETE FROM CRE_HIPCIE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCIE_CODEMP = '" & p_CodEmp & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCIE_PERMES = " & CStr(r_int_PerMes) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCIE_PERANO = " & CStr(r_int_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontro registros en tabla CRE_HIPCIE.")
   End If
      
   p_BarPro.FloodPercent = 100
   r_str_FFnEje = Format(date, "yyyymmdd")
   r_str_HFnEje = Format(Time, "hhmmss")
   
   p_NumErr = r_lng_NumErr
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_NumReg, r_lng_NumErr, "000001", "", 0, 0, 0, Format(CDate(r_str_FecPro), "yyyymmdd"), Format(CDate(r_str_FecPro), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1003(ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String, ByRef p_NumErr As Long, ByVal p_ArcRCC As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1003
'Descripción      :  Carga de Archivo RCC
'Resumen          :  Proceso que carga la información del RCC de los clientes de Créditos Hipotecarios
'F. Creación      :  04-05-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_lng_NumErr        As Long
Dim r_str_FFnEje        As String
Dim r_str_HFnEje        As String
Dim r_str_FecPro        As String
Dim r_int_NumFil        As Integer
Dim r_str_LineaL        As String
Dim r_int_PosSp1        As Integer
Dim r_int_PosSp2        As Integer
Dim r_int_PosSp3        As Integer
Dim r_int_PosSp4        As Integer
Dim r_int_PosSp5        As Integer
Dim r_int_PosSp6        As Integer
Dim r_int_PosSp7        As Integer
Dim r_str_CodSbs        As String
Dim r_str_FecRep        As String
Dim r_str_DocTri        As String
Dim r_str_NumRuc        As String
Dim r_str_TipDoc        As String
Dim r_str_NumDoc        As String
Dim r_str_TipPer        As String
Dim r_str_Evalua()      As String
Dim r_int_ConTem        As Integer
Dim r_int_NumIte        As Integer
Dim r_arr_PerNat()      As moddat_tpo_Genera
Dim r_arr_PerJur()      As moddat_tpo_Genera
Dim r_dbl_DeuNor        As Double
Dim r_dbl_DeuCpp        As Double
Dim r_dbl_DeuDef        As Double
Dim r_dbl_DeuDud        As Double
Dim r_dbl_DeuPer        As Double
Dim r_str_EmpRep        As String
Dim r_int_MonDeu        As Integer
Dim r_int_DiaAtr        As Integer
Dim r_int_TipDeu        As Integer
Dim r_dbl_SalDeu        As Double
Dim r_int_ClaDeu        As Integer
Dim r_str_CtaCtb        As String
Dim r_int_FlgEnc        As Integer
Dim r_int_Contad        As Integer
Dim r_str_CarPos        As String
Dim r_int_NumEmp        As Integer
Dim r_str_Refere        As String

   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_lng_NumErr = 0
   r_int_NumIte = 0
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_PerNat(0)
   ReDim r_arr_PerJur(0)
   ReDim r_str_Evalua(0)
   
   'Cargando en Arreglo a Personas Naturales (Titular)
   modprc_g_str_CadEje = "SELECT DISTINCT HIPMAE_TDOCLI, HIPMAE_NDOCLI FROM CRE_HIPMAE WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SUBSTR(HIPMAE_FECDES,1,6) <= " & p_PerAno & Format(p_PerMes, "00")
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPMAE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerNat(UBound(r_arr_PerNat) + 1)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_TipDoc = modprc_g_rst_Genera!HIPMAE_TDOCLI
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_NumDoc = Trim(modprc_g_rst_Genera!HIPMAE_NDOCLI)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Cantid = 0
      
         r_lng_TotReg = r_lng_TotReg + 1
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando en Arreglo a Personas Naturales (Cónyuge)
   modprc_g_str_CadEje = "SELECT DISTINCT HIPMAE_TDOCYG, HIPMAE_NDOCYG FROM CRE_HIPMAE WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_TDOCYG > 0 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SUBSTR(HIPMAE_FECDES,1,6) <= " & p_PerAno & Format(p_PerMes, "00")
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPMAE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerNat(UBound(r_arr_PerNat) + 1)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_TipDoc = modprc_g_rst_Genera!HIPMAE_TDOCYG
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_NumDoc = Trim(modprc_g_rst_Genera!HIPMAE_NDOCYG)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Cantid = 0
      
         r_lng_TotReg = r_lng_TotReg + 1
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando a Personas Juridicas (Vendedor)
   modprc_g_str_CadEje = "SELECT DISTINCT DATGEN_VENTDO, DATGEN_VENNDO FROM PRY_DATGEN WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_PRYMCS = 1 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_SITUAC = 1"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer PRY_DATGEN.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerJur(UBound(r_arr_PerJur) + 1)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_TipDoc = modprc_g_rst_Genera!DATGEN_VENTDO
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_NumDoc = Trim(modprc_g_rst_Genera!DATGEN_VENNDO)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Cantid = 0
         
         modprc_g_rst_Genera.MoveNext
         r_lng_TotReg = r_lng_TotReg + 1
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando a Personas Juridicas (Constructor)
   modprc_g_str_CadEje = "SELECT DISTINCT DATGEN_CONTDO, DATGEN_CONNDO FROM PRY_DATGEN WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_PRYMCS = 1 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_SITUAC = 1"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer PRY_DATGEN.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         r_int_FlgEnc = 2
         If UBound(r_arr_PerJur) > 0 Then
            For r_int_Contad = 1 To UBound(r_arr_PerJur)
               If r_arr_PerJur(r_int_Contad).Genera_TipDoc = modprc_g_rst_Genera!DATGEN_CONTDO And r_arr_PerJur(r_int_Contad).Genera_NumDoc = Trim(modprc_g_rst_Genera!DATGEN_CONNDO) Then
                  r_int_FlgEnc = 1
               End If
            Next r_int_Contad
         End If
      
         If r_int_FlgEnc = 2 Then
            ReDim Preserve r_arr_PerJur(UBound(r_arr_PerJur) + 1)
            r_arr_PerJur(UBound(r_arr_PerJur)).Genera_TipDoc = modprc_g_rst_Genera!DATGEN_CONTDO
            r_arr_PerJur(UBound(r_arr_PerJur)).Genera_NumDoc = Trim(modprc_g_rst_Genera!DATGEN_CONNDO)
            r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Cantid = 0
            r_lng_TotReg = r_lng_TotReg + 1
         End If
         
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   '************************************ACTUALIZADO 23/08/2017 ***********************************************
   'Cargando en Arreglo a Personas Naturales (ENTIDAD TÉCNICA)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT DISTINCT MAEETE_TIPDOC, MAEETE_NUMDOC FROM TPR_MAEETE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE MAEETE_SITUAC = 1"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MAEETE_TIPDOC <> 6 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPMAE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerNat(UBound(r_arr_PerNat) + 1)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_TipDoc = modprc_g_rst_Genera!MAEETE_TIPDOC
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_NumDoc = Trim(modprc_g_rst_Genera!MAEETE_NUMDOC)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Cantid = 0
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Refere = "ET"
         r_lng_TotReg = r_lng_TotReg + 1
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando a Personas Juridicas (ENTIDAD TÉCNICA)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT DISTINCT MAEETE_TIPDOC, MAEETE_NUMDOC FROM TPR_MAEETE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE MAEETE_SITUAC = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MAEETE_TIPDOC = 6 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer TPR_MAEETE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerJur(UBound(r_arr_PerJur) + 1)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_TipDoc = modprc_g_rst_Genera!MAEETE_TIPDOC
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_NumDoc = Trim(modprc_g_rst_Genera!MAEETE_NUMDOC)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Cantid = 0
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Refere = "ET"
         
         modprc_g_rst_Genera.MoveNext
         r_lng_TotReg = r_lng_TotReg + 1
         DoEvents
      Loop
   End If
   '*****************************************************************************************************************
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Abriendo Archivo RCC
   r_int_NumFil = FreeFile
   Open p_ArcRCC For Input As r_int_NumFil
   Line Input #r_int_NumFil, r_str_LineaL
   
   Do While Not EOF(r_int_NumFil)
      
      If Mid(r_str_LineaL, 1, 1) = "1" Then
         r_int_PosSp1 = InStr(1, r_str_LineaL, "|")                           'Código SBS
         r_int_PosSp2 = InStr(r_int_PosSp1 + 1, r_str_LineaL, "|")            'Fecha Reporte
         r_int_PosSp3 = InStr(r_int_PosSp2 + 1, r_str_LineaL, "|")            'Tipo Documento Tributario
         r_int_PosSp4 = InStr(r_int_PosSp3 + 1, r_str_LineaL, "|")            'RUC
         r_int_PosSp5 = InStr(r_int_PosSp4 + 1, r_str_LineaL, "|")            'Tipo Documento de Identidad
         r_int_PosSp6 = InStr(r_int_PosSp5 + 1, r_str_LineaL, "|")            'Número Documento de Identidad
         r_int_PosSp7 = InStr(r_int_PosSp6 + 1, r_str_LineaL, "|")            'Tipo de Persona
         
         r_str_CodSbs = Mid(r_str_LineaL, 2, r_int_PosSp1 - 2)
         r_str_FecRep = Mid(r_str_LineaL, r_int_PosSp1 + 1, r_int_PosSp2 - 1 - r_int_PosSp1)
         r_str_DocTri = Mid(r_str_LineaL, r_int_PosSp2 + 1, r_int_PosSp3 - 1 - r_int_PosSp2)
         r_str_NumRuc = Mid(r_str_LineaL, r_int_PosSp3 + 1, r_int_PosSp4 - 1 - r_int_PosSp3)
         r_str_TipDoc = Mid(r_str_LineaL, r_int_PosSp4 + 1, r_int_PosSp5 - 1 - r_int_PosSp4)
         r_str_NumDoc = Mid(r_str_LineaL, r_int_PosSp5 + 1, r_int_PosSp6 - 1 - r_int_PosSp5)
         r_str_TipPer = Mid(r_str_LineaL, r_int_PosSp6 + 1, r_int_PosSp7 - 1 - r_int_PosSp6)
                  
         If Len(Trim(r_str_TipPer)) = 0 Then
            r_str_TipPer = "0"
         End If
         
         If Len(Trim(r_str_TipDoc)) > 0 And CInt(r_str_TipPer) = 1 Then
            'Persona Natural
            r_int_FlgEnc = 2
            For r_int_Contad = 1 To UBound(r_arr_PerNat)
               If CInt(r_str_TipDoc) = r_arr_PerNat(r_int_Contad).Genera_TipDoc And r_str_NumDoc = r_arr_PerNat(r_int_Contad).Genera_NumDoc Then
                  r_arr_PerNat(r_int_Contad).Genera_Cantid = 1
                  r_str_Refere = r_arr_PerNat(r_int_Contad).Genera_Refere
                  r_int_FlgEnc = 1
                  Exit For
               End If
               DoEvents
            Next r_int_Contad
            
            'Si encontro Cliente del Archivo RCC en Base de Datos de Clientes (Personas Naturales)
            If r_int_FlgEnc = 1 Then
               
               'Si cliente ya no ha sido registrado anteriormente en RCCCAB
               If ff_Valida_RCCCab(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno) Then
                  r_lng_NumReg = r_lng_NumReg + 1
                  r_dbl_DeuNor = 0:    r_dbl_DeuCpp = 0:    r_dbl_DeuDef = 0:    r_dbl_DeuDud = 0:    r_dbl_DeuPer = 0
                  
                  Line Input #r_int_NumFil, r_str_LineaL
                  DoEvents
                  
                  r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                  r_int_NumEmp = 0
                  r_int_NumIte = 0
                  
                  Erase r_str_Evalua()
                  ReDim r_str_Evalua(0)
                  
                  Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                     If r_str_Refere = "ET" Then
                        'RETIRADO => Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428" And
                        'ADICIONADO => (8104, 7205) SOLES COMO DOLARES
                        If ((Mid(r_str_LineaL, 19, 2) = "14" Or Mid(r_str_LineaL, 19, 2) = "71") And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                             Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429") Or (Mid(r_str_LineaL, 19, 4) = "8113" Or _
                             Mid(r_str_LineaL, 19, 4) = "8123" Or _
                             Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124" Or _
                             Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225") Then
                             
                              r_int_NumIte = r_int_NumIte + 1
                              
                              If UBound(r_str_Evalua) = 0 Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              Else
                                 For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                    If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                       Exit For
                                    End If
                                 Next
                                 If r_int_ConTem > UBound(r_str_Evalua) Then
                                    ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                    r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                    r_int_NumEmp = r_int_NumEmp + 1
                                 End If
                              End If
                              
                              r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                              r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                              r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                              
                              If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                                 r_int_TipDeu = 9                                   'Tarjeta de Crédito
                              Else
                                 r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                              End If
                              
                              r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                              r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                              
                              Select Case r_int_ClaDeu
                                 Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                                 Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                                 Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                                 Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                                 Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                              End Select
                              
                              'Insertar en Base de Datos CLI_RCCDET
                              Call fs_Inserta_RCCDet(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     Else
                        'RETIRADO =>  Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428"
                        'ADICIONADO => (8104)(7101,7103,7104,7205) TANTO SOLES COMO DOLARES
                        If (Mid(r_str_LineaL, 19, 2) = "14" And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                            Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429") Or _
                           (Mid(r_str_LineaL, 19, 6) = "811302" Or Mid(r_str_LineaL, 19, 6) = "812302" Or _
                            Mid(r_str_LineaL, 19, 6) = "811925" Or Mid(r_str_LineaL, 19, 6) = "812925" Or _
                            Mid(r_str_LineaL, 19, 6) = "811922" Or Mid(r_str_LineaL, 19, 6) = "812922" Or _
                            Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124" Or _
                            Mid(r_str_LineaL, 19, 4) = "7111" Or Mid(r_str_LineaL, 19, 4) = "7121" Or _
                            Mid(r_str_LineaL, 19, 4) = "7113" Or Mid(r_str_LineaL, 19, 4) = "7123" Or _
                            Mid(r_str_LineaL, 19, 4) = "7114" Or Mid(r_str_LineaL, 19, 4) = "7124" Or _
                            Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225") Then
                           
                           r_int_NumIte = r_int_NumIte + 1
                           
                           If UBound(r_str_Evalua) = 0 Then
                              ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                              r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_NumEmp = r_int_NumEmp + 1
                           Else
                              For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                 If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                    Exit For
                                 End If
                              Next
                              If r_int_ConTem > UBound(r_str_Evalua) Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              End If
                           End If
                           
                           r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                           r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                           r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                           r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                           
                           If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                              r_int_TipDeu = 9                                   'Tarjeta de Crédito
                           Else
                              r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                           End If
                           
                           r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                           r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                           
                           Select Case r_int_ClaDeu
                              Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                              Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                              Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                              Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                              Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                           End Select
                           
                           'Insertar en Base de Datos CLI_RCCDET
                           Call fs_Inserta_RCCDet(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     End If
                     Line Input #r_int_NumFil, r_str_LineaL
                     DoEvents
                  Loop
               
                  If r_int_NumEmp > 0 Then
                     'Insertar en Base de Datos CLI_RCCCAB
                     Call fs_Inserta_RCCCab(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), r_str_CodSbs, r_int_NumEmp, r_dbl_DeuNor, r_dbl_DeuCpp, r_dbl_DeuDef, r_dbl_DeuDud, r_dbl_DeuPer)
                  End If
                  
                  If r_str_Refere = "ET" Then
                     r_str_TipPer = 4
                  End If
                  
                  'Actualizando Código SBS
                  If Not ff_Actualiza_CodigoSBS(CInt(r_str_TipDoc), r_str_NumDoc, r_str_CodSbs, r_str_TipPer) Then '1
                     r_lng_NumErr = r_lng_NumErr + 1
                     Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al actualizar Código SBS. (" & r_str_TipDoc & "-" & r_str_NumDoc & ")")
                  End If
                  
               Else
                  r_lng_NumErr = r_lng_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Cliente ya fue registrado (" & r_str_TipDoc & "-" & r_str_NumDoc & ").")
                  
                  Line Input #r_int_NumFil, r_str_LineaL
                  DoEvents
               End If
            Else
               Line Input #r_int_NumFil, r_str_LineaL
               DoEvents
            End If                     'Fin Persona Natural (r_int_FlgEnc = 1)
            
         ElseIf Len(Trim(r_str_NumRuc)) = 11 Then
            'Persona Juridica
            r_int_FlgEnc = 2
            For r_int_Contad = 1 To UBound(r_arr_PerJur)
               'If r_arr_PerJur(r_int_Contad).Genera_TipDoc = 7 And r_str_NumRuc = r_arr_PerJur(r_int_Contad).Genera_NumDoc Then
               If (r_arr_PerJur(r_int_Contad).Genera_TipDoc = 6 Or r_arr_PerJur(r_int_Contad).Genera_TipDoc = 7) And r_str_NumRuc = r_arr_PerJur(r_int_Contad).Genera_NumDoc Then
                  r_arr_PerJur(r_int_Contad).Genera_Cantid = 1
                  r_str_Refere = r_arr_PerJur(r_int_Contad).Genera_Refere
                  r_int_FlgEnc = 1
                  Exit For
                  
               End If
               
               DoEvents
            Next r_int_Contad
            
            'Si encontro Cliente del Archivo RCC en Base de Datos de Clientes (Personas Juridicas)
            If r_int_FlgEnc = 1 Then
               'Si cliente ya no ha sido registrado anteriormente en RCCCAB
               If ff_Valida_RCCCab(7, r_str_NumRuc, p_PerMes, p_PerAno) Then
                  r_lng_NumReg = r_lng_NumReg + 1
                  r_dbl_DeuNor = 0:    r_dbl_DeuCpp = 0:    r_dbl_DeuDef = 0:    r_dbl_DeuDud = 0:    r_dbl_DeuPer = 0
               
                  Line Input #r_int_NumFil, r_str_LineaL
                  
                  r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                  r_int_NumEmp = 0
                  r_int_NumIte = 0
                  
                  Erase r_str_Evalua
                  ReDim r_str_Evalua(0)
                  
                  Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                     If r_str_Refere = "ET" Then
                        'RETIRADO => Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428" And
                        'ADICIONADO => (8104, 7205)
                        If ((Mid(r_str_LineaL, 19, 2) = "14" Or Mid(r_str_LineaL, 19, 2) = "71") And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                            Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429") Or (Mid(r_str_LineaL, 19, 4) = "8113" Or _
                            Mid(r_str_LineaL, 19, 4) = "8123" Or _
                            Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124" Or _
                            Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225") Then
                           
                              r_int_NumIte = r_int_NumIte + 1
                              
                              If UBound(r_str_Evalua) = 0 Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              Else
                                 For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                    If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                       Exit For
                                    End If
                                 Next
                                 
                                 If r_int_ConTem > UBound(r_str_Evalua) Then
                                    ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                    r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                    r_int_NumEmp = r_int_NumEmp + 1
                                 End If
                              End If
                              
                              r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                              r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                              r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                              
                              If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                                 r_int_TipDeu = 9                                   'Tarjeta de Crédito
                              Else
                                 r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                              End If
                              
                              r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                              r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                              
                              Select Case r_int_ClaDeu
                                 Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                                 Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                                 Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                                 Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                                 Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                              End Select
                              
                              'Insertar en Base de Datos CLI_RCCDET
                              Call fs_Inserta_RCCDet(7, r_str_NumRuc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     Else
                        'RETIRADO => And Mid(r_str_LineaL, 19, 4) <> "1418" AND Mid(r_str_LineaL, 19, 4) <> "1428"
                        'ADICIONADO => (7101, 7103, 7104, 7205, 8104)
                        If Mid(r_str_LineaL, 19, 2) = "14" And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                           Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429" Or _
                           (Mid(r_str_LineaL, 19, 6) = "811302" Or Mid(r_str_LineaL, 19, 6) = "812302" Or _
                           Mid(r_str_LineaL, 19, 6) = "812925" Or Mid(r_str_LineaL, 19, 6) = "811925" Or _
                           Mid(r_str_LineaL, 19, 6) = "812922" Or Mid(r_str_LineaL, 19, 6) = "811922" Or _
                           Mid(r_str_LineaL, 19, 4) = "7111" Or Mid(r_str_LineaL, 19, 4) = "7121" Or _
                           Mid(r_str_LineaL, 19, 4) = "7113" Or Mid(r_str_LineaL, 19, 4) = "7123" Or _
                           Mid(r_str_LineaL, 19, 4) = "7114" Or Mid(r_str_LineaL, 19, 4) = "7124" Or _
                           Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225" Or _
                           Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124") Then
                           
                           r_int_NumIte = r_int_NumIte + 1
                           
                           If UBound(r_str_Evalua) = 0 Then
                              ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                              r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_NumEmp = r_int_NumEmp + 1
                           Else
                              For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                 If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                    Exit For
                                 End If
                              Next
                              
                              If r_int_ConTem > UBound(r_str_Evalua) Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              End If
                           End If
                           
                           r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                           r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                           r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                           r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                           
                           If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                              r_int_TipDeu = 9                                   'Tarjeta de Crédito
                           Else
                              r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                           End If
                           
                           r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                           r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                           
                           Select Case r_int_ClaDeu
                              Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                              Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                              Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                              Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                              Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                           End Select
                           
                           'Insertar en Base de Datos CLI_RCCDET
                           Call fs_Inserta_RCCDet(7, r_str_NumRuc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     End If
                     Line Input #r_int_NumFil, r_str_LineaL
                     DoEvents
                  Loop
               
                  If r_int_NumEmp > 0 Then
                     'Insertar en Base de Datos CLI_RCCCAB
                     Call fs_Inserta_RCCCab(7, r_str_NumRuc, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), r_str_CodSbs, r_int_NumEmp, r_dbl_DeuNor, r_dbl_DeuCpp, r_dbl_DeuDef, r_dbl_DeuDud, r_dbl_DeuPer)
                  End If
                  
                  If r_str_Refere = "ET" Then
                     r_str_TipPer = 4
                  End If
                  
                  'Actualizando Código SBS
                  'If Not ff_Actualiza_CodigoSBS(7, r_str_NumRuc, r_str_CodSbs, 2) Then
                  If Not ff_Actualiza_CodigoSBS(IIf(r_str_TipPer = 4, 6, 7), r_str_NumRuc, r_str_CodSbs, r_str_TipPer) Then
                     r_lng_NumErr = r_lng_NumErr + 1
                     Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al actualizar Código SBS. (7-" & r_str_NumRuc & ")")
                  End If
               Else
                  r_lng_NumErr = r_lng_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Cliente ya fue registrado (7-" & r_str_NumRuc & ").")
               
                  Line Input #r_int_NumFil, r_str_LineaL
                  DoEvents
               End If
            Else
               Line Input #r_int_NumFil, r_str_LineaL
               DoEvents
            End If            'Fin Persona Juridica (r_int_FlgEnc = 1)
         Else
            'Otro Tipo de Persona No Identificada
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
         End If
      Else
         'Si es Línea de Detalle
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
      End If
      
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Loop
         
   'Verificando si existen clientes en Base de Datos miCasita y que no hay información en Archivo RCC (Personas Naturales)
   For r_int_Contad = 1 To UBound(r_arr_PerNat)
      If r_arr_PerNat(r_int_Contad).Genera_Cantid = 0 Then
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontro cliente. (" & CStr(r_arr_PerNat(r_int_Contad).Genera_TipDoc) & "-" & Trim(r_arr_PerNat(r_int_Contad).Genera_NumDoc) & ")")
      End If
   Next r_int_Contad
         
   'Verificando si existen clientes en Base de Datos miCasita y que no hay información en Archivo RCC (Personas Juridicas)
   For r_int_Contad = 1 To UBound(r_arr_PerJur)
      If r_arr_PerJur(r_int_Contad).Genera_Cantid = 0 Then
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontro cliente. (" & CStr(r_arr_PerJur(r_int_Contad).Genera_TipDoc) & "-" & Trim(r_arr_PerJur(r_int_Contad).Genera_NumDoc) & ")")
      End If
   Next r_int_Contad
   
   'Cerrando Archivo RCC
   Close #r_int_NumFil
   
   'Grabando en LOG información de Proceso
   r_str_FFnEje = Format(date, "yyyymmdd")
   r_str_HFnEje = Format(Time, "hhmmss")
   
   p_NumErr = r_lng_NumErr
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_TotReg, r_lng_NumErr, "", "", 0, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), Format(CDate(r_str_FecPro), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1003_OLD(ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String, ByRef p_NumErr As Long, ByVal p_ArcRCC As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1003
'Descripción      :  Carga de Archivo RCC
'Resumen          :  Proceso que carga la información del RCC de los clientes de Créditos Hipotecarios
'F. Creación      :  04-05-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_lng_NumErr        As Long
Dim r_str_FFnEje        As String
Dim r_str_HFnEje        As String
Dim r_str_FecPro        As String
Dim r_int_NumFil        As Integer
Dim r_str_LineaL        As String
Dim r_int_PosSp1        As Integer
Dim r_int_PosSp2        As Integer
Dim r_int_PosSp3        As Integer
Dim r_int_PosSp4        As Integer
Dim r_int_PosSp5        As Integer
Dim r_int_PosSp6        As Integer
Dim r_int_PosSp7        As Integer
Dim r_str_CodSbs        As String
Dim r_str_FecRep        As String
Dim r_str_DocTri        As String
Dim r_str_NumRuc        As String
Dim r_str_TipDoc        As String
Dim r_str_NumDoc        As String
Dim r_str_TipPer        As String
Dim r_str_Evalua()      As String
Dim r_int_ConTem        As Integer
Dim r_int_NumIte        As Integer
Dim r_arr_PerNat()      As moddat_tpo_Genera
Dim r_arr_PerJur()      As moddat_tpo_Genera
Dim r_dbl_DeuNor        As Double
Dim r_dbl_DeuCpp        As Double
Dim r_dbl_DeuDef        As Double
Dim r_dbl_DeuDud        As Double
Dim r_dbl_DeuPer        As Double
Dim r_str_EmpRep        As String
Dim r_int_MonDeu        As Integer
Dim r_int_DiaAtr        As Integer
Dim r_int_TipDeu        As Integer
Dim r_dbl_SalDeu        As Double
Dim r_int_ClaDeu        As Integer
Dim r_str_CtaCtb        As String
Dim r_int_FlgEnc        As Integer
Dim r_int_Contad        As Integer
Dim r_str_CarPos        As String
Dim r_int_NumEmp        As Integer
Dim r_str_Refere        As String

   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_lng_NumErr = 0
   r_int_NumIte = 0
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_PerNat(0)
   ReDim r_arr_PerJur(0)
   ReDim r_str_Evalua(0)
   
   'Cargando en Arreglo a Personas Naturales (Titular)
   modprc_g_str_CadEje = "SELECT DISTINCT HIPMAE_TDOCLI, HIPMAE_NDOCLI FROM CRE_HIPMAE WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SUBSTR(HIPMAE_FECDES,1,6) <= " & p_PerAno & Format(p_PerMes, "00")
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPMAE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerNat(UBound(r_arr_PerNat) + 1)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_TipDoc = modprc_g_rst_Genera!HIPMAE_TDOCLI
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_NumDoc = Trim(modprc_g_rst_Genera!HIPMAE_NDOCLI)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Cantid = 0
      
         r_lng_TotReg = r_lng_TotReg + 1
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando en Arreglo a Personas Naturales (Cónyuge)
   modprc_g_str_CadEje = "SELECT DISTINCT HIPMAE_TDOCYG, HIPMAE_NDOCYG FROM CRE_HIPMAE WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_TDOCYG > 0 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SUBSTR(HIPMAE_FECDES,1,6) <= " & p_PerAno & Format(p_PerMes, "00")
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPMAE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerNat(UBound(r_arr_PerNat) + 1)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_TipDoc = modprc_g_rst_Genera!HIPMAE_TDOCYG
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_NumDoc = Trim(modprc_g_rst_Genera!HIPMAE_NDOCYG)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Cantid = 0
      
         r_lng_TotReg = r_lng_TotReg + 1
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando a Personas Juridicas (Vendedor)
   modprc_g_str_CadEje = "SELECT DISTINCT DATGEN_VENTDO, DATGEN_VENNDO FROM PRY_DATGEN WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_PRYMCS = 1 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_SITUAC = 1"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer PRY_DATGEN.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerJur(UBound(r_arr_PerJur) + 1)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_TipDoc = modprc_g_rst_Genera!DATGEN_VENTDO
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_NumDoc = Trim(modprc_g_rst_Genera!DATGEN_VENNDO)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Cantid = 0
         
         modprc_g_rst_Genera.MoveNext
         r_lng_TotReg = r_lng_TotReg + 1
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando a Personas Juridicas (Constructor)
   modprc_g_str_CadEje = "SELECT DISTINCT DATGEN_CONTDO, DATGEN_CONNDO FROM PRY_DATGEN WHERE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_PRYMCS = 1 AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DATGEN_SITUAC = 1"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer PRY_DATGEN.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         r_int_FlgEnc = 2
         If UBound(r_arr_PerJur) > 0 Then
            For r_int_Contad = 1 To UBound(r_arr_PerJur)
               If r_arr_PerJur(r_int_Contad).Genera_TipDoc = modprc_g_rst_Genera!DATGEN_CONTDO And r_arr_PerJur(r_int_Contad).Genera_NumDoc = Trim(modprc_g_rst_Genera!DATGEN_CONNDO) Then
                  r_int_FlgEnc = 1
               End If
            Next r_int_Contad
         End If
      
         If r_int_FlgEnc = 2 Then
            ReDim Preserve r_arr_PerJur(UBound(r_arr_PerJur) + 1)
            r_arr_PerJur(UBound(r_arr_PerJur)).Genera_TipDoc = modprc_g_rst_Genera!DATGEN_CONTDO
            r_arr_PerJur(UBound(r_arr_PerJur)).Genera_NumDoc = Trim(modprc_g_rst_Genera!DATGEN_CONNDO)
            r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Cantid = 0
            r_lng_TotReg = r_lng_TotReg + 1
         End If
         
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   '************************************ACTUALIZADO 23/08/2017 ***********************************************
   'Cargando en Arreglo a Personas Naturales (ENTIDAD TÉCNICA)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT DISTINCT MAEETE_TIPDOC, MAEETE_NUMDOC FROM TPR_MAEETE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE MAEETE_SITUAC = 1"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MAEETE_TIPDOC <> 6 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer CRE_HIPMAE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerNat(UBound(r_arr_PerNat) + 1)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_TipDoc = modprc_g_rst_Genera!MAEETE_TIPDOC
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_NumDoc = Trim(modprc_g_rst_Genera!MAEETE_NUMDOC)
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Cantid = 0
         r_arr_PerNat(UBound(r_arr_PerNat)).Genera_Refere = "ET"
         r_lng_TotReg = r_lng_TotReg + 1
         modprc_g_rst_Genera.MoveNext
         DoEvents
      Loop
   End If

   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
   
   'Cargando a Personas Juridicas (ENTIDAD TÉCNICA)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT DISTINCT MAEETE_TIPDOC, MAEETE_NUMDOC FROM TPR_MAEETE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE MAEETE_SITUAC = 1 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MAEETE_TIPDOC = 6 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al leer TPR_MAEETE.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      Do While Not modprc_g_rst_Genera.EOF
         ReDim Preserve r_arr_PerJur(UBound(r_arr_PerJur) + 1)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_TipDoc = modprc_g_rst_Genera!MAEETE_TIPDOC
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_NumDoc = Trim(modprc_g_rst_Genera!MAEETE_NUMDOC)
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Cantid = 0
         r_arr_PerJur(UBound(r_arr_PerJur)).Genera_Refere = "ET"
         
         modprc_g_rst_Genera.MoveNext
         r_lng_TotReg = r_lng_TotReg + 1
         DoEvents
      Loop
   End If
   '*****************************************************************************************************************
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Abriendo Archivo RCC
   r_int_NumFil = FreeFile
   Open p_ArcRCC For Input As r_int_NumFil
   Line Input #r_int_NumFil, r_str_LineaL
   
   Do While Not EOF(r_int_NumFil)
      
      If Mid(r_str_LineaL, 1, 1) = "1" Then
         r_int_PosSp1 = InStr(1, r_str_LineaL, "|")                           'Código SBS
         r_int_PosSp2 = InStr(r_int_PosSp1 + 1, r_str_LineaL, "|")            'Fecha Reporte
         r_int_PosSp3 = InStr(r_int_PosSp2 + 1, r_str_LineaL, "|")            'Tipo Documento Tributario
         r_int_PosSp4 = InStr(r_int_PosSp3 + 1, r_str_LineaL, "|")            'RUC
         r_int_PosSp5 = InStr(r_int_PosSp4 + 1, r_str_LineaL, "|")            'Tipo Documento de Identidad
         r_int_PosSp6 = InStr(r_int_PosSp5 + 1, r_str_LineaL, "|")            'Número Documento de Identidad
         r_int_PosSp7 = InStr(r_int_PosSp6 + 1, r_str_LineaL, "|")            'Tipo de Persona
         
         r_str_CodSbs = Mid(r_str_LineaL, 2, r_int_PosSp1 - 2)
         r_str_FecRep = Mid(r_str_LineaL, r_int_PosSp1 + 1, r_int_PosSp2 - 1 - r_int_PosSp1)
         r_str_DocTri = Mid(r_str_LineaL, r_int_PosSp2 + 1, r_int_PosSp3 - 1 - r_int_PosSp2)
         r_str_NumRuc = Mid(r_str_LineaL, r_int_PosSp3 + 1, r_int_PosSp4 - 1 - r_int_PosSp3)
         r_str_TipDoc = Mid(r_str_LineaL, r_int_PosSp4 + 1, r_int_PosSp5 - 1 - r_int_PosSp4)
         r_str_NumDoc = Mid(r_str_LineaL, r_int_PosSp5 + 1, r_int_PosSp6 - 1 - r_int_PosSp5)
         r_str_TipPer = Mid(r_str_LineaL, r_int_PosSp6 + 1, r_int_PosSp7 - 1 - r_int_PosSp6)
                  
         If Len(Trim(r_str_TipPer)) = 0 Then
            r_str_TipPer = "0"
         End If
         
         If Len(Trim(r_str_TipDoc)) > 0 And CInt(r_str_TipPer) = 1 Then
            'Persona Natural
            r_int_FlgEnc = 2
            For r_int_Contad = 1 To UBound(r_arr_PerNat)
               If CInt(r_str_TipDoc) = r_arr_PerNat(r_int_Contad).Genera_TipDoc And r_str_NumDoc = r_arr_PerNat(r_int_Contad).Genera_NumDoc Then
                  r_arr_PerNat(r_int_Contad).Genera_Cantid = 1
                  r_str_Refere = r_arr_PerNat(r_int_Contad).Genera_Refere
                  r_int_FlgEnc = 1
                  Exit For
               End If
               DoEvents
            Next r_int_Contad
            
            'Si encontro Cliente del Archivo RCC en Base de Datos de Clientes (Personas Naturales)
            If r_int_FlgEnc = 1 Then
               
               'Si cliente ya no ha sido registrado anteriormente en RCCCAB
               If ff_Valida_RCCCab(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno) Then
                  r_lng_NumReg = r_lng_NumReg + 1
                  r_dbl_DeuNor = 0:    r_dbl_DeuCpp = 0:    r_dbl_DeuDef = 0:    r_dbl_DeuDud = 0:    r_dbl_DeuPer = 0
                  
                  Line Input #r_int_NumFil, r_str_LineaL
                  DoEvents
                  
                  r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                  r_int_NumEmp = 0
                  r_int_NumIte = 0
                  
                  Erase r_str_Evalua()
                  ReDim r_str_Evalua(0)
                  
                  Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                     If r_str_Refere = "ET" Then
                        'RETIRADO => Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428" And
                        'ADICIONADO => (8104, 7205) SOLES COMO DOLARES
                        If ((Mid(r_str_LineaL, 19, 2) = "14" Or Mid(r_str_LineaL, 19, 2) = "71") And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                             Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429") Or (Mid(r_str_LineaL, 19, 4) = "8113" Or _
                             Mid(r_str_LineaL, 19, 4) = "8123" Or _
                             Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124" Or _
                             Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225") Then
                             
                              r_int_NumIte = r_int_NumIte + 1
                              
                              If UBound(r_str_Evalua) = 0 Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              Else
                                 For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                    If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                       Exit For
                                    End If
                                 Next
                                 If r_int_ConTem > UBound(r_str_Evalua) Then
                                    ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                    r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                    r_int_NumEmp = r_int_NumEmp + 1
                                 End If
                              End If
                              
                              r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                              r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                              r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                              
                              If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                                 r_int_TipDeu = 9                                   'Tarjeta de Crédito
                              Else
                                 r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                              End If
                              
                              r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                              r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                              
                              Select Case r_int_ClaDeu
                                 Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                                 Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                                 Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                                 Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                                 Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                              End Select
                              
                              'Insertar en Base de Datos CLI_RCCDET
                              Call fs_Inserta_RCCDet(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     Else
                        'RETIRADO =>  Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428"
                        'ADICIONADO => (8104)(7101,7103,7104,7205) TANTO SOLES COMO DOLARES
                        If (Mid(r_str_LineaL, 19, 2) = "14" And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                            Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429") Or _
                           (Mid(r_str_LineaL, 19, 6) = "811302" Or Mid(r_str_LineaL, 19, 6) = "812302" Or _
                            Mid(r_str_LineaL, 19, 6) = "811925" Or Mid(r_str_LineaL, 19, 6) = "812925" Or _
                            Mid(r_str_LineaL, 19, 6) = "811922" Or Mid(r_str_LineaL, 19, 6) = "812922" Or _
                            Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124" Or _
                            Mid(r_str_LineaL, 19, 4) = "7111" Or Mid(r_str_LineaL, 19, 4) = "7121" Or _
                            Mid(r_str_LineaL, 19, 4) = "7113" Or Mid(r_str_LineaL, 19, 4) = "7123" Or _
                            Mid(r_str_LineaL, 19, 4) = "7114" Or Mid(r_str_LineaL, 19, 4) = "7124" Or _
                            Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225") Then
                           
                           r_int_NumIte = r_int_NumIte + 1
                           
                           If UBound(r_str_Evalua) = 0 Then
                              ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                              r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_NumEmp = r_int_NumEmp + 1
                           Else
                              For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                 If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                    Exit For
                                 End If
                              Next
                              If r_int_ConTem > UBound(r_str_Evalua) Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              End If
                           End If
                           
                           r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                           r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                           r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                           r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                           
                           If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                              r_int_TipDeu = 9                                   'Tarjeta de Crédito
                           Else
                              r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                           End If
                           
                           r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                           r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                           
                           Select Case r_int_ClaDeu
                              Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                              Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                              Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                              Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                              Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                           End Select
                           
                           'Insertar en Base de Datos CLI_RCCDET
                           Call fs_Inserta_RCCDet(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     End If
                     Line Input #r_int_NumFil, r_str_LineaL
                     DoEvents
                  Loop
               
                  If r_int_NumEmp > 0 Then
                     'Insertar en Base de Datos CLI_RCCCAB
                     Call fs_Inserta_RCCCab(CInt(r_str_TipDoc), r_str_NumDoc, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), r_str_CodSbs, r_int_NumEmp, r_dbl_DeuNor, r_dbl_DeuCpp, r_dbl_DeuDef, r_dbl_DeuDud, r_dbl_DeuPer)
                  End If
                  
                  If r_str_Refere = "ET" Then
                     r_str_TipPer = 4
                  End If
                  
                  'Actualizando Código SBS
                  If Not ff_Actualiza_CodigoSBS(CInt(r_str_TipDoc), r_str_NumDoc, r_str_CodSbs, r_str_TipPer) Then '1
                     r_lng_NumErr = r_lng_NumErr + 1
                     Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al actualizar Código SBS. (" & r_str_TipDoc & "-" & r_str_NumDoc & ")")
                  End If
                  
               Else
                  r_lng_NumErr = r_lng_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Cliente ya fue registrado (" & r_str_TipDoc & "-" & r_str_NumDoc & ").")
                  
                  Line Input #r_int_NumFil, r_str_LineaL
                  DoEvents
               End If
            Else
               Line Input #r_int_NumFil, r_str_LineaL
               DoEvents
            End If                     'Fin Persona Natural (r_int_FlgEnc = 1)
            
         ElseIf Len(Trim(r_str_NumRuc)) = 11 Then
            'Persona Juridica
            r_int_FlgEnc = 2
            For r_int_Contad = 1 To UBound(r_arr_PerJur)
               'If r_arr_PerJur(r_int_Contad).Genera_TipDoc = 7 And r_str_NumRuc = r_arr_PerJur(r_int_Contad).Genera_NumDoc Then
               If (r_arr_PerJur(r_int_Contad).Genera_TipDoc = 6 Or r_arr_PerJur(r_int_Contad).Genera_TipDoc = 7) And r_str_NumRuc = r_arr_PerJur(r_int_Contad).Genera_NumDoc Then
                  r_arr_PerJur(r_int_Contad).Genera_Cantid = 1
                  r_str_Refere = r_arr_PerJur(r_int_Contad).Genera_Refere
                  r_int_FlgEnc = 1
                  Exit For
                  
               End If
               
               DoEvents
            Next r_int_Contad
            
            'Si encontro Cliente del Archivo RCC en Base de Datos de Clientes (Personas Juridicas)
            If r_int_FlgEnc = 1 Then
               'Si cliente ya no ha sido registrado anteriormente en RCCCAB
               If ff_Valida_RCCCab(7, r_str_NumRuc, p_PerMes, p_PerAno) Then
                  r_lng_NumReg = r_lng_NumReg + 1
                  r_dbl_DeuNor = 0:    r_dbl_DeuCpp = 0:    r_dbl_DeuDef = 0:    r_dbl_DeuDud = 0:    r_dbl_DeuPer = 0
               
                  Line Input #r_int_NumFil, r_str_LineaL
                  
                  r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                  r_int_NumEmp = 0
                  r_int_NumIte = 0
                  
                  Erase r_str_Evalua
                  ReDim r_str_Evalua(0)
                  
                  Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
                     If r_str_Refere = "ET" Then
                        'RETIRADO => Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428" And
                        'ADICIONADO => (8104, 7205)
                        If ((Mid(r_str_LineaL, 19, 2) = "14" Or Mid(r_str_LineaL, 19, 2) = "71") And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                            Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429") Or (Mid(r_str_LineaL, 19, 4) = "8113" Or _
                            Mid(r_str_LineaL, 19, 4) = "8123" Or _
                            Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124" Or _
                            Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225") Then
                           
                              r_int_NumIte = r_int_NumIte + 1
                              
                              If UBound(r_str_Evalua) = 0 Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              Else
                                 For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                    If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                       Exit For
                                    End If
                                 Next
                                 
                                 If r_int_ConTem > UBound(r_str_Evalua) Then
                                    ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                    r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                    r_int_NumEmp = r_int_NumEmp + 1
                                 End If
                              End If
                              
                              r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                              r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                              r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                              
                              If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                                 r_int_TipDeu = 9                                   'Tarjeta de Crédito
                              Else
                                 r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                              End If
                              
                              r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                              r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                              
                              Select Case r_int_ClaDeu
                                 Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                                 Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                                 Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                                 Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                                 Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                              End Select
                              
                              'Insertar en Base de Datos CLI_RCCDET
                              Call fs_Inserta_RCCDet(7, r_str_NumRuc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     Else
                        'RETIRADO => And Mid(r_str_LineaL, 19, 4) <> "1418" AND Mid(r_str_LineaL, 19, 4) <> "1428"
                        'ADICIONADO => (7101, 7103, 7104, 7205, 8104)
                        If Mid(r_str_LineaL, 19, 2) = "14" And CStr(CLng(Mid(r_str_LineaL, 12, 5))) <> 240 And _
                           Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429" Or _
                           (Mid(r_str_LineaL, 19, 6) = "811302" Or Mid(r_str_LineaL, 19, 6) = "812302" Or _
                           Mid(r_str_LineaL, 19, 6) = "812925" Or Mid(r_str_LineaL, 19, 6) = "811925" Or _
                           Mid(r_str_LineaL, 19, 6) = "812922" Or Mid(r_str_LineaL, 19, 6) = "811922" Or _
                           Mid(r_str_LineaL, 19, 4) = "7111" Or Mid(r_str_LineaL, 19, 4) = "7121" Or _
                           Mid(r_str_LineaL, 19, 4) = "7113" Or Mid(r_str_LineaL, 19, 4) = "7123" Or _
                           Mid(r_str_LineaL, 19, 4) = "7114" Or Mid(r_str_LineaL, 19, 4) = "7124" Or _
                           Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225" Or _
                           Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124") Then
                           
                           r_int_NumIte = r_int_NumIte + 1
                           
                           If UBound(r_str_Evalua) = 0 Then
                              ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                              r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_NumEmp = r_int_NumEmp + 1
                           Else
                              For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                                 If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                    Exit For
                                 End If
                              Next
                              
                              If r_int_ConTem > UBound(r_str_Evalua) Then
                                 ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                                 r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                                 r_int_NumEmp = r_int_NumEmp + 1
                              End If
                           End If
                           
                           r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                           r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                           r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                           r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                           
                           If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                              r_int_TipDeu = 9                                   'Tarjeta de Crédito
                           Else
                              r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
                           End If
                           
                           r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                           r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                           
                           Select Case r_int_ClaDeu
                              Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                              Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                              Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                              Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                              Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                           End Select
                           
                           'Insertar en Base de Datos CLI_RCCDET
                           Call fs_Inserta_RCCDet(7, r_str_NumRuc, p_PerMes, p_PerAno, r_int_NumIte, r_str_EmpRep, r_int_DiaAtr, r_int_MonDeu, IIf(r_int_MonDeu = 1, r_dbl_SalDeu, 0), IIf(r_int_MonDeu = 2, r_dbl_SalDeu, 0), r_int_ClaDeu, r_int_TipDeu, r_str_CtaCtb)
                        End If
                     End If
                     Line Input #r_int_NumFil, r_str_LineaL
                     DoEvents
                  Loop
               
                  If r_int_NumEmp > 0 Then
                     'Insertar en Base de Datos CLI_RCCCAB
                     Call fs_Inserta_RCCCab(7, r_str_NumRuc, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), r_str_CodSbs, r_int_NumEmp, r_dbl_DeuNor, r_dbl_DeuCpp, r_dbl_DeuDef, r_dbl_DeuDud, r_dbl_DeuPer)
                  End If
                  
                  If r_str_Refere = "ET" Then
                     r_str_TipPer = 4
                  End If
                  
                  'Actualizando Código SBS
                  'If Not ff_Actualiza_CodigoSBS(7, r_str_NumRuc, r_str_CodSbs, 2) Then
                  If Not ff_Actualiza_CodigoSBS(IIf(r_str_TipPer = 4, 6, 7), r_str_NumRuc, r_str_CodSbs, r_str_TipPer) Then
                     r_lng_NumErr = r_lng_NumErr + 1
                     Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al actualizar Código SBS. (7-" & r_str_NumRuc & ")")
                  End If
               Else
                  r_lng_NumErr = r_lng_NumErr + 1
                  Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Cliente ya fue registrado (7-" & r_str_NumRuc & ").")
               
                  Line Input #r_int_NumFil, r_str_LineaL
                  DoEvents
               End If
            Else
               Line Input #r_int_NumFil, r_str_LineaL
               DoEvents
            End If            'Fin Persona Juridica (r_int_FlgEnc = 1)
         Else
            'Otro Tipo de Persona No Identificada
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
         End If
      Else
         'Si es Línea de Detalle
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
      End If
      
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Loop
         
   'Verificando si existen clientes en Base de Datos miCasita y que no hay información en Archivo RCC (Personas Naturales)
   For r_int_Contad = 1 To UBound(r_arr_PerNat)
      If r_arr_PerNat(r_int_Contad).Genera_Cantid = 0 Then
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontro cliente. (" & CStr(r_arr_PerNat(r_int_Contad).Genera_TipDoc) & "-" & Trim(r_arr_PerNat(r_int_Contad).Genera_NumDoc) & ")")
      End If
   Next r_int_Contad
         
   'Verificando si existen clientes en Base de Datos miCasita y que no hay información en Archivo RCC (Personas Juridicas)
   For r_int_Contad = 1 To UBound(r_arr_PerJur)
      If r_arr_PerJur(r_int_Contad).Genera_Cantid = 0 Then
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "No se encontro cliente. (" & CStr(r_arr_PerJur(r_int_Contad).Genera_TipDoc) & "-" & Trim(r_arr_PerJur(r_int_Contad).Genera_NumDoc) & ")")
      End If
   Next r_int_Contad
   
   'Cerrando Archivo RCC
   Close #r_int_NumFil
   
   'Grabando en LOG información de Proceso
   r_str_FFnEje = Format(date, "yyyymmdd")
   r_str_HFnEje = Format(Time, "hhmmss")
   
   p_NumErr = r_lng_NumErr
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_TotReg, r_lng_NumErr, "", "", 0, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), Format(CDate(r_str_FecPro), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1004(ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String, ByRef p_NumErr As Long, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1004
'Descripción      :  Anulación de Carga de Archivo RCC
'Resumen          :  Proceso que carga la información del RCC de los clientes de Créditos Hipotecarios
'F. Creación      :  06-05-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg     As Long
Dim r_lng_TotReg     As Long
Dim r_lng_NumErr     As Long
Dim r_str_FFnEje     As String
Dim r_str_HFnEje     As String
Dim r_str_FecPro     As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer
Dim r_str_Period        As String
   
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_lng_NumErr = 0
   p_BarPro.FloodPercent = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Borrando Cabecera
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DELETE FROM CLI_RCCCAB "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE RCCCAB_PERMES = " & CStr(p_PerMes) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND RCCCAB_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al borrar en CLI_RCCCAB.")
   End If
      
   p_BarPro.FloodPercent = 50
   
   'Borrando Detalle
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DELETE FROM CLI_RCCDET "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE RCCDET_PERMES = " & CStr(p_PerMes) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND RCCDET_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al borrar en CLI_RCCDET.")
   End If
      
   p_BarPro.FloodPercent = 100
   r_str_FFnEje = Format(date, "yyyymmdd")
   r_str_HFnEje = Format(Time, "hhmmss")
   
   p_NumErr = r_lng_NumErr
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FFnEje, r_str_HFnEje, r_lng_NumReg, r_lng_NumErr, "", "", 0, p_PerMes, p_PerAno, Format(CDate(r_str_FecPro), "yyyymmdd"), Format(CDate(r_str_FecPro), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1005(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerIni As String, ByVal p_PerFin As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_CtbIni As String, ByVal p_CtbFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1005
'Descripción      :  Registro de Pagos de Cuotas de Créditos Hipotecarios
'Resumen          :  Contabilización de Pago de Cuotas de Créditos Hipotecarios
'F. Creación      :  16-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :  13-03-2010
'U. Actualización :  Jorge Luis Tacuche Mesia

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_arr_CtaBan()      As modprc_g_tpo_CtaBan
Dim r_arr_MatPro()      As modprc_g_tpo_MatPro
Dim r_arr_Matriz()      As modprc_g_tpo_Matriz
Dim r_arr_MatDet()      As modprc_g_tpo_MatDet
Dim r_arr_CtaPrd()      As modprc_g_tpo_CtaPrd
Dim r_str_FecPro        As String
Dim r_str_FecVct        As String
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_int_NumAsi        As Integer
Dim r_int_PosMat        As Integer
Dim r_int_AuxCon        As Integer
Dim r_rst_Genera        As ADODB.Recordset
Dim r_rst_Cuotas        As ADODB.Recordset
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_rst_CtaBan        As ADODB.Recordset
Dim r_rst_TipMat        As ADODB.Recordset
Dim r_rst_DifCam        As ADODB.Recordset
Dim r_rst_MatCtb        As ADODB.Recordset
Dim r_rst_MatCab        As ADODB.Recordset
Dim r_rst_MatDet        As ADODB.Recordset
Dim r_rst_CtaPrd        As ADODB.Recordset
Dim r_dbl_AcuDvg        As Double
Dim r_dbl_AcuDvc        As Double
Dim r_dbl_AcuDif        As Double
Dim r_dbl_TipCam        As Double
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_dbl_TipSbs        As Double
Dim r_dbl_TipSun        As Double
Dim r_int_OpeGnb        As Integer
Dim r_dbl_Asi_CapVig    As Double
Dim r_dbl_Asi_CapVen    As Double
Dim r_dbl_Asi_IntEfe    As Double
Dim r_dbl_Asi_IntDev    As Double
Dim r_dbl_Asi_IntVen    As Double
Dim r_dbl_Asi_IntDif    As Double
Dim r_dbl_Asi_SegDes    As Double
Dim r_dbl_Asi_SegInm    As Double

Dim r_dbl_Asi_SegDes_Ven As Double 'Vencidos
Dim r_dbl_Asi_SegInm_Ven As Double '
Dim r_dbl_Asi_SegDes_Liq As Double 'Adelantados (Por Liquidar)
Dim r_dbl_Asi_SegInm_Liq As Double '
Dim r_str_Asi_CtaDeb     As String
Dim r_str_Asi_CtaHab     As String

Dim r_dbl_Asi_Portes    As Double
Dim r_dbl_Asi_CapPBP    As Double
Dim r_dbl_Asi_IntPBP    As Double
Dim r_dbl_Asi_IntMor    As Double
Dim r_dbl_Asi_IntCom    As Double
Dim r_dbl_Asi_GasCob    As Double
Dim r_dbl_Asi_OtrGas    As Double
Dim r_dbl_Asi_ImpITF    As Double
Dim r_dbl_Asi_IMoVig    As Double
Dim r_dbl_Asi_IMoVen    As Double
Dim r_dbl_Asi_CVgGar    As Double
Dim r_dbl_Asi_CVgSGa    As Double
Dim r_dbl_Mes_IntEfe    As Double
Dim r_dbl_Mes_IntDev    As Double
Dim r_dbl_Mes_IntVen    As Double
Dim r_dbl_Mes_IntDif    As Double
Dim r_dbl_Mes_CapVig    As Double
Dim r_dbl_Mes_CapVen    As Double
Dim r_dbl_Mes_IMoVig    As Double
Dim r_dbl_Mes_IMoVen    As Double
Dim r_dbl_Mes_CVgGar    As Double
Dim r_dbl_Mes_CVgSGa    As Double
Dim r_int_TipGar        As Integer
Dim r_int_Cont01        As Integer
Dim r_int_SitCre        As Integer
Dim r_int_Refina        As Integer
Dim r_int_Judici        As Integer
Dim r_int_Castig        As Integer
Dim r_dbl_Totdeb_sol    As Double
Dim r_dbl_Totdeb_dol    As Double
Dim r_dbl_Tothab_sol    As Double
Dim r_dbl_Tothab_dol    As Double
Dim r_str_CueGan        As String
Dim r_str_CuePer        As String
Dim r_str_Cuenta        As String
Dim r_str_FlagDH        As String
Dim r_str_FecAsi        As String
Dim r_dbl_Asi_ManCta    As Double
Dim r_dbl_Asi_CtaBnc    As Double
   
   r_lng_NumReg = 0
   r_lng_TotReg = ff_ConAmo(p_FecIni, p_FecFin)
   p_BarPro.FloodPercent = 0
      
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1005"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
        
   '1.= Para obtener Cuentas de Diferencia de Cambio (Ganancia o Pérdida)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAGEN "
   g_str_Parame = g_str_Parame & " WHERE CTAGEN_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & " ORDER BY CTAGEN_CTACTB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_DifCam, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_DifCam.BOF And r_rst_DifCam.EOF) Then
      r_rst_DifCam.MoveFirst
      Do While Not r_rst_DifCam.EOF
         If r_rst_DifCam!CTAGEN_CODIDE = "01" Then
            r_str_CueGan = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         ElseIf r_rst_DifCam!CTAGEN_CODIDE = "02" Then
            r_str_CuePer = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         End If
         
         r_rst_DifCam.MoveNext
      Loop
   End If
   
   r_rst_DifCam.Close
   Set r_rst_DifCam = Nothing
   
   '2.= Para leer cuentas por cada Cuenta Bancaria
   ReDim r_arr_CtaBan(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN "
   g_str_Parame = g_str_Parame & " ORDER BY CTABAN_CODBAN, CTABAN_NUMCTA ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaBan, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaBan.BOF And r_rst_CtaBan.EOF) Then
      r_rst_CtaBan.MoveFirst
      Do While Not r_rst_CtaBan.EOF
         ReDim Preserve r_arr_CtaBan(UBound(r_arr_CtaBan) + 1)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CodBan = r_rst_CtaBan!CtaBan_CodBan
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_NumCta = Trim(r_rst_CtaBan!CtaBan_NumCta)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CtaCtb = Trim(r_rst_CtaBan!CtaBan_CtaCtb)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_TipCta = r_rst_CtaBan!CtaBan_TipCta
         
         r_rst_CtaBan.MoveNext
      Loop
   End If
   
   r_rst_CtaBan.Close
   Set r_rst_CtaBan = Nothing
   
   '3.= Para leer Matrices Contables de todos los Producto
   ReDim r_arr_MatDet(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATCAB "
   g_str_Parame = g_str_Parame & " WHERE MATCAB_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & "   AND MATCAB_TIPMAT = '100004' "
   g_str_Parame = g_str_Parame & " ORDER BY MATCAB_SITCRE, SUBSTR(MATCAB_CODMAT,1,3) ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatCab, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_MatCab.BOF And r_rst_MatCab.EOF) Then
      r_rst_MatCab.MoveFirst
      Do While Not r_rst_MatCab.EOF
         
         '4.= Matriz por cada producto
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATDET "
         g_str_Parame = g_str_Parame & " WHERE MATDET_CODMAT = '" & r_rst_MatCab!MATCAB_CODMAT & "' "
         g_str_Parame = g_str_Parame & " ORDER BY MATDET_CODMAT, MATDET_NUMITE ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatDet, 3) Then
            Exit Sub
         End If
         
         r_rst_MatDet.MoveFirst
         Do While Not r_rst_MatDet.EOF
            ReDim Preserve r_arr_MatDet(UBound(r_arr_MatDet) + 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodMat = r_rst_MatCab!MATCAB_CODMAT
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodPrd = Trim(Mid(r_rst_MatCab!MATCAB_CODMAT, 1, 3))
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesCab = Trim(r_rst_MatCab!MATCAB_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipMon = r_rst_MatCab!MATCAB_TIPMON
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_SitCre = r_rst_MatCab!MATCAB_SITCRE
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesDet = Trim(r_rst_MatDet!MATDET_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CtbCon = Trim(r_rst_MatDet!MATDET_CONCTB)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DebHab = Left(moddat_gf_Consulta_ParDes("255", CStr(r_rst_MatDet!MATDET_FLGDHB)), 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipCam = CInt(r_rst_MatDet!MATDET_TIPTCA)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_NroLib = r_rst_MatCab!MATCAB_CODLIB
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_OpeCon = Trim(r_rst_MatDet!MATDET_CONOPE) 'mnt_pardes 64 - conceptos operativos
            
            r_rst_MatDet.MoveNext
         Loop
         
         r_rst_MatDet.Close
         Set r_rst_MatDet = Nothing
         r_rst_MatCab.MoveNext
      Loop
   End If
   
   r_rst_MatCab.Close
   Set r_rst_MatCab = Nothing
   
   '5.= Para leer cuentas Contables de todos los productos
   ReDim r_arr_CtaPrd(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRD "
   g_str_Parame = g_str_Parame & " ORDER BY CTAPRD_CODPRD ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaPrd, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaPrd.BOF And r_rst_CtaPrd.EOF) Then
      r_rst_CtaPrd.MoveFirst
      Do While Not r_rst_CtaPrd.EOF
         ReDim Preserve r_arr_CtaPrd(UBound(r_arr_CtaPrd) + 1)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CodPrd = Trim(r_rst_CtaPrd!CtaPrd_CodPrd)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtbCon = Trim(r_rst_CtaPrd!CTAPRD_CONCTB) 'CONCEPTO CONTABLE
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_SitCre = Trim(r_rst_CtaPrd!CTAPRD_TIPCRE)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtaCtb = Trim(r_rst_CtaPrd!CtaPrd_CtaCtb)
         
         r_rst_CtaPrd.MoveNext
      Loop
   End If
   
   r_rst_CtaPrd.Close
   Set r_rst_CtaPrd = Nothing
   
   '6.= Consulta principal de amotizaciones por procesar
   modprc_g_str_CadEje = "SELECT CAJMOV_SUCMOV, CAJMOV_USUMOV, CAJMOV_FECMOV, CAJMOV_FECDEP, CAJMOV_NUMMOV, CAJMOV_NUMOPE, CAJMOV_ITFIMP, CAJMOV_MONPAG, CAJMOV_IMPTOT, CAJMOV_CODBAN, CAJMOV_NUMCTA, CAJMOV_IMPPAG FROM OPE_CAJMOV " & _
                         " WHERE CAJMOV_TIPMOV = 1102 " & _
                         "   AND CAJMOV_CTBFLG = 0 " & _
                         "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " " & _
                         "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         " ORDER BY CAJMOV_NUMMOV ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Fecha de Proceso
         r_str_FecPro = Format(p_FecFin, "dd/mm/yyyy")
         If CDate(gf_FormatoFecha(modprc_g_rst_Princi!CAJMOV_FECMOV)) > CDate(p_CtbFin) Then
            r_str_FecAsi = p_CtbFin
         Else
            r_str_FecAsi = gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))
         End If
         
         r_int_OpeGnb = 0
         r_dbl_Asi_CapVig = 0
         r_dbl_Asi_CVgGar = 0
         r_dbl_Asi_CVgSGa = 0
         r_dbl_Asi_CapVen = 0
         r_dbl_Asi_IntEfe = 0
         r_dbl_Asi_IntDev = 0
         r_dbl_Asi_IntVen = 0
         r_dbl_Asi_IntDif = 0
         r_dbl_Asi_SegDes = 0
         r_dbl_Asi_SegInm = 0
         
         r_dbl_Asi_SegDes_Ven = 0 'Vencidos
         r_dbl_Asi_SegInm_Ven = 0
         r_dbl_Asi_SegDes_Liq = 0 'Adelantados (Por Liquidar)
         r_dbl_Asi_SegInm_Liq = 0
         
         r_dbl_Asi_Portes = 0
         r_dbl_Asi_CapPBP = 0
         r_dbl_Asi_IntPBP = 0
         r_dbl_Asi_IntMor = 0
         r_dbl_Asi_IntCom = 0
         r_dbl_Asi_GasCob = 0
         r_dbl_Asi_OtrGas = 0
         r_dbl_Asi_IMoVig = 0
         r_dbl_Asi_IMoVen = 0
         r_dbl_Asi_ImpITF = modprc_g_rst_Princi!CAJMOV_ITFIMP
         r_dbl_Tothab_sol = 0
         r_dbl_Totdeb_sol = 0
         
         '7.= Para obtener Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = "SELECT HIPMAE_ACUDVG, HIPMAE_ACUDVC, HIPMAE_ACUDIF, HIPMAE_TIPGAR, HIPMAE_SITACT, HIPMAE_REFINA, HIPMAE_JUDICI, HIPMAE_CASTIG FROM CRE_HIPMAE " & _
                               " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         r_dbl_AcuDvg = r_rst_Genera!HIPMAE_ACUDVG
         r_dbl_AcuDvc = r_rst_Genera!HIPMAE_ACUDVC
         r_dbl_AcuDif = r_rst_Genera!HIPMAE_ACUDIF
         r_int_TipGar = r_rst_Genera!HIPMAE_TIPGAR
         r_int_SitCre = r_rst_Genera!HIPMAE_SITACT
         r_int_Refina = r_rst_Genera!HIPMAE_REFINA
         r_int_Judici = r_rst_Genera!HIPMAE_JUDICI
         r_int_Castig = r_rst_Genera!HIPMAE_CASTIG
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         
         '8.= Para obtener Cuotas Pagadas de la operacion
         modprc_g_str_CadEje = "SELECT * FROM CRE_HIPPAG " & _
                               " WHERE HIPPAG_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' " & _
                               "   AND HIPPAG_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' " & _
                               "   AND HIPPAG_NUMMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV) & " " & _
                               "   AND HIPPAG_FECMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_FECMOV) & " " & _
                               " ORDER BY HIPPAG_NUMCUO ASC"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPPAG - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         Do While Not r_rst_Genera.EOF
            r_dbl_Mes_IntEfe = 0
            r_dbl_Mes_IntDev = 0
            r_dbl_Mes_IntVen = 0
            r_dbl_Mes_IntDif = 0
            r_dbl_Mes_CapVen = 0
            r_dbl_Mes_CapVig = 0
            r_dbl_Mes_IMoVig = 0
            r_dbl_Mes_IMoVen = 0
            r_dbl_Mes_CVgGar = 0
            r_dbl_Mes_CVgSGa = 0
            
            '9.= Para obtener Fecha de Vencimiento de Cuota que corresponde
            modprc_g_str_CadEje = "SELECT HIPCUO_FECVCT FROM CRE_HIPCUO " & _
                                  " WHERE HIPCUO_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' " & _
                                  "   AND HIPCUO_TIPCRO = 1 " & _
                                  "   AND HIPCUO_NUMCUO = " & CStr(r_rst_Genera!HIPPAG_NUMCUO) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Cuotas, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
            End If
            
            r_rst_Cuotas.MoveFirst
            r_str_FecVct = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECVCT))
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
         
            '10.= Determinando Tipo de Pago (Mes, Atrasado, Adelantado) y distribuyendo Interés Pagado (HIPPAG_INTERE)
            If CDate(r_str_FecVct) >= CDate(p_CtbIni) And CDate(r_str_FecVct) <= CDate(p_CtbFin) Then
               'Pago Mes
               r_dbl_Asi_SegDes = r_dbl_Asi_SegDes + r_rst_Genera!HIPPAG_DESORG
               r_dbl_Asi_SegInm = r_dbl_Asi_SegInm + r_rst_Genera!HIPPAG_VIVORG
               
               If r_dbl_AcuDvc = 0 Then      'Acumulado Devengado Vencido
                  If r_rst_Genera!HIPPAG_INTERE >= r_dbl_AcuDvg Then
                     r_dbl_Mes_IntDev = r_dbl_AcuDvg
                     r_dbl_Mes_IntEfe = r_rst_Genera!HIPPAG_INTERE - r_dbl_Mes_IntDev
                     r_dbl_AcuDvg = 0
                  Else
                     r_dbl_Mes_IntDev = r_rst_Genera!HIPPAG_INTERE
                     r_dbl_AcuDvg = r_dbl_AcuDvg - r_dbl_Mes_IntDev
                  End If

               'MODIFICACIÒN AL 03/11/2016 - PARA OBTENER EL INTERES EFECTIVO DEL MES - CUADRE DE ASIENTOS
               Else
                  r_dbl_Mes_IntEfe = r_rst_Genera!HIPPAG_INTERE
               End If
            ElseIf CDate(r_str_FecVct) < CDate(p_CtbIni) Then
               'Pago Atrasado
               r_dbl_Asi_SegDes_Ven = r_dbl_Asi_SegDes_Ven + r_rst_Genera!HIPPAG_DESORG 'Vencidos
               r_dbl_Asi_SegInm_Ven = r_dbl_Asi_SegInm_Ven + r_rst_Genera!HIPPAG_VIVORG
               
               If r_dbl_AcuDvc = 0 Then
                  If r_rst_Genera!HIPPAG_INTERE >= r_dbl_AcuDvg Then
                     r_dbl_Mes_IntDev = r_dbl_AcuDvg
                     r_dbl_Mes_IntEfe = r_rst_Genera!HIPPAG_INTERE - r_dbl_Mes_IntDev
                     r_dbl_AcuDvg = 0
                  Else
                     r_dbl_Mes_IntDev = r_rst_Genera!HIPPAG_INTERE
                     r_dbl_AcuDvg = r_dbl_AcuDvg - r_dbl_Mes_IntDev
                  End If
                  
                  r_dbl_Mes_IMoVig = r_dbl_Mes_IMoVig + r_rst_Genera!HIPPAG_INTMOR
               Else
                  'Si Crédito esta vencido
                  If r_rst_Genera!HIPPAG_INTERE >= r_dbl_AcuDvc Then
                     r_dbl_Mes_IntVen = r_dbl_AcuDvc
                     r_dbl_AcuDvc = 0
                  Else
                     r_dbl_Mes_IntVen = r_rst_Genera!HIPPAG_INTERE
                     r_dbl_AcuDvc = r_dbl_AcuDvc - r_dbl_Mes_IntVen
                  End If
                  
                  r_dbl_Mes_CapVen = r_rst_Genera!HIPPAG_CAPITA
                  r_dbl_Mes_IMoVen = r_dbl_Mes_IMoVen + r_rst_Genera!HIPPAG_INTMOR
               End If
            ElseIf CDate(r_str_FecVct) > CDate(p_CtbFin) Then
               'Pago Adelantado
               r_dbl_Asi_SegDes_Liq = r_dbl_Asi_SegDes_Liq + r_rst_Genera!HIPPAG_DESORG 'Adelantados (Por Liquidar)
               r_dbl_Asi_SegInm_Liq = r_dbl_Asi_SegInm_Liq + r_rst_Genera!HIPPAG_VIVORG
            
               r_dbl_Mes_IntDif = r_rst_Genera!HIPPAG_INTERE
               r_dbl_AcuDif = r_dbl_AcuDif + r_dbl_Mes_IntDif
            Else
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al determinar Tipo de Pago - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
            End If
            
            r_dbl_Mes_CapVig = r_rst_Genera!HIPPAG_CAPITA - r_dbl_Mes_CapVen
            
            'Para determinar Tipo de Garantia
            If r_int_TipGar = 1 Or r_int_TipGar = 2 Then
               r_dbl_Mes_CVgGar = r_dbl_Mes_CapVig
            Else
               r_dbl_Mes_CVgSGa = r_dbl_Mes_CapVig
            End If
            
            'Actualizar en CRE_HIPPAG distribución de Intereses
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPPAG "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPPAG_DEVVIG = " & CStr(r_dbl_Mes_IntDev) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_DEVVEN = " & CStr(r_dbl_Mes_IntVen) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_INTDIF = " & CStr(r_dbl_Mes_IntDif) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_INTEFE = " & CStr(r_dbl_Mes_IntEfe) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CAPVIG = " & CStr(r_dbl_Mes_CapVig) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CAPVEN = " & CStr(r_dbl_Mes_CapVen) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_IMOVEN = " & CStr(r_dbl_Mes_IMoVen) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_IMOVIG = " & CStr(r_dbl_Mes_IMoVig) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CVGGAR = " & CStr(r_dbl_Mes_CVgGar) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CVGSGA = " & CStr(r_dbl_Mes_CVgSGa) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPPAG_NUMOPE = '" & modprc_g_rst_Princi!CAJMOV_NUMOPE & "' AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_NUMCUO = " & CStr(r_rst_Genera!HIPPAG_NUMCUO) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_NUMMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_FECMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_FECMOV) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPPAG - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
            End If
            
            r_dbl_Asi_CapVig = r_dbl_Asi_CapVig + r_dbl_Mes_CapVig
            r_dbl_Asi_CVgGar = r_dbl_Asi_CVgGar + r_dbl_Mes_CVgGar
            r_dbl_Asi_CVgSGa = r_dbl_Asi_CVgSGa + r_dbl_Mes_CVgSGa
            r_dbl_Asi_CapVen = r_dbl_Asi_CapVen + r_dbl_Mes_CapVen
            r_dbl_Asi_IntEfe = r_dbl_Asi_IntEfe + r_dbl_Mes_IntEfe
            r_dbl_Asi_IntDev = r_dbl_Asi_IntDev + r_dbl_Mes_IntDev
            r_dbl_Asi_IntVen = r_dbl_Asi_IntVen + r_dbl_Mes_IntVen
            r_dbl_Asi_IntDif = r_dbl_Asi_IntDif + r_dbl_Mes_IntDif
            'r_dbl_Asi_SegDes = r_dbl_Asi_SegDes + r_rst_Genera!HIPPAG_DESORG
            'r_dbl_Asi_SegInm = r_dbl_Asi_SegInm + r_rst_Genera!HIPPAG_VIVORG
            r_dbl_Asi_Portes = r_dbl_Asi_Portes + r_rst_Genera!HIPPAG_OTRORG
            r_dbl_Asi_CapPBP = r_dbl_Asi_CapPBP + r_rst_Genera!HIPPAG_CAPBBP
            r_dbl_Asi_IntPBP = r_dbl_Asi_IntPBP + r_rst_Genera!HIPPAG_INTBBP
            r_dbl_Asi_IntMor = r_dbl_Asi_IntMor + r_rst_Genera!HIPPAG_INTMOR
            r_dbl_Asi_IntCom = r_dbl_Asi_IntCom + r_rst_Genera!HIPPAG_INTCOM
            r_dbl_Asi_GasCob = r_dbl_Asi_GasCob + r_rst_Genera!HIPPAG_GASCOB
            r_dbl_Asi_OtrGas = r_dbl_Asi_OtrGas + r_rst_Genera!HIPPAG_OTRGAS
            r_dbl_Asi_IMoVig = r_dbl_Asi_IMoVig + r_dbl_Mes_IMoVig
            r_dbl_Asi_IMoVen = r_dbl_Asi_IMoVen + r_dbl_Mes_IMoVen
            
            r_rst_Genera.MoveNext
         Loop
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         
         'Actualizando en CRE_HIPMAE Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_ACUDVG = " & CStr(r_dbl_AcuDvg) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDVC = " & CStr(r_dbl_AcuDvc) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDIF = " & CStr(r_dbl_AcuDif) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_CAPVEN = HIPMAE_CAPVEN - " & CStr(r_dbl_Asi_CapVen) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         If (r_int_SitCre = 1) Then
            r_int_SitCre = 1
         ElseIf (r_int_SitCre = 5) And (r_int_Refina = 0) And (r_int_Judici = 0) And (r_int_Castig = 0) Then
            r_int_SitCre = 1
         End If
         
         If (r_int_Refina = 1) Then
            r_int_SitCre = 4
         ElseIf (r_int_Judici = 1) Then
            r_int_SitCre = 6
         ElseIf (r_int_Castig = 1) Then
            r_int_SitCre = 3
         End If
         
         'GENERACION DE ASIENTOS CONTABLES
         If ff_Valida_OperacionGNB(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE), Format(r_str_FecVct, "YYYYMMDD")) Then
            'Operaciones GNB y OTROS
            r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
            r_dbl_ImpSol = modprc_g_rst_Princi!CAJMOV_IMPTOT
            
            'Obtiene correlativo
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, "LM", 6)
            
            'Ingresa Cabecera
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, "001", r_dbl_TipSbs, "O", "DEPOSITO POR DEVOLVER A GNB " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000"), r_str_FecAsi, "1")
            
            'Ingresa Detalle
            r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(modprc_g_rst_Princi!CAJMOV_IMPTOT / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 1, "111301060201", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB", 1, 60), "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 2, "291807010103", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB", 1, 60), "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            
            r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(modprc_g_dbl_ComBcoSol / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 3, "291807010103", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB - COMISION", 1, 60), "D", modprc_g_dbl_ComBcoSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 4, "111301060201", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB - COMISION", 1, 60), "H", modprc_g_dbl_ComBcoSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
         Else
            ReDim r_arr_Matriz(0)
            r_int_PosMat = 0
            
            For r_int_Contad = 1 To UBound(r_arr_MatDet)
               
               'Comparacion de Situacion de Credito con la Situacion de la Matriz
               If r_arr_MatDet(r_int_Contad).MatDet_SitCre = r_int_SitCre Then
                  
                  'Filtra las matrices del producto de la operacion en evaluacion
                  If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE), 1, 3) Then 'mismo producto
                     r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecAsi, "yyyymmdd"), 1)
                     r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
                     
                     If r_int_PosMat = 0 Then
                        r_int_PosMat = r_int_Contad
                     End If
                     
                     ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
                     For r_int_AuxCon = 1 To UBound(r_arr_CtaPrd)
                        If Trim(r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CodPrd) = Trim(r_arr_MatDet(r_int_Contad).MatDet_CodPrd) Then 'codigo producto
                           If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtbCon = r_arr_MatDet(r_int_Contad).MatDet_CtbCon Then 'conceptos contables
                              If r_int_SitCre = r_arr_MatDet(r_int_Contad).MatDet_SitCre Then 'situacion credito
                                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtaCtb
                              End If
                           End If
                        End If
                     Next r_int_AuxCon
                     
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = r_arr_MatDet(r_int_Contad).MatDet_DebHab
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_OpeCon = r_arr_MatDet(r_int_Contad).MatDet_OpeCon 'CONCEPTOS OPERATIVOS
                     
                     For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                        If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_TipCta = r_arr_CtaBan(r_int_Cont01).CtaBan_TipCta
                           Exit For
                        End If
                     Next r_int_Cont01
                     
                     '(100320 - AMORT.-COMISION RECAUDO) (100321 - AMORT.-CTA. BANCO)
                     If r_arr_MatDet(r_int_Contad).MatDet_OpeCon = "100320" Or r_arr_MatDet(r_int_Contad).MatDet_OpeCon = "100321" Then
                        If r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_TipCta = "000002" Then 'CUENTA DE AHORROS-000002) ==> 510
                           If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
                              r_dbl_Asi_ManCta = modprc_g_dbl_ComBcoDol
                              r_dbl_Asi_CtaBnc = modprc_g_dbl_ComBcoDol
                           Else
                              r_dbl_Asi_ManCta = modprc_g_dbl_ComBcoSol
                              r_dbl_Asi_CtaBnc = modprc_g_dbl_ComBcoSol
                           End If
                        Else
                           r_dbl_Asi_ManCta = 0
                           r_dbl_Asi_CtaBnc = 0
                        End If
                     End If
                     
                     Select Case r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                        Case "100301": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapVig 'Capital Vigente (AMORT.-CAPITAL VIGENTE-100301)
                        Case "100317": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CVgGar 'Capital Vigente con Garantia (AMORT.-CAPITAL VIG. CON GARANT.-100317)
                        Case "100318": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CVgSGa 'Capital Vigente sin Garantia (AMORT.-CAPITAL VIG. SIN GARANT.-100318)
                        Case "100303": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapVen 'Capital vencido (AMORT.-CAPITAL VENCIDO-100303)
                        Case "100305": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntEfe 'interes efectivo (AMORT.-INTERES EFECTIVO-100305)
                        Case "100304": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntDev 'interes devengado (AMORT.-INTERES DEVENGADO-100304)
                        Case "100306": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntVen 'interes vencido (AMORT.-INTERES VENCIDO-100306)
                        Case "100316": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntDif 'interes diferido (AMORT.-INT. DIFERIDO-100316)
                        '====================
                        Case "100312": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegDes 'seguro desgravamen (AMORT.-SEG. DESGRAVAMEN-100312)
                        Case "100313": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegInm 'seguro de inmueble (AMORT.-SEG. INMUEBLE-100313)
                        Case "100322": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegDes_Ven '100322 AMORT.-SEG. ATRAZADO
                        Case "100323": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegDes_Liq '100323 AMORT.-SEG. ADELANTADOS
                        Case "100324": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegInm_Ven '100324 AMORT.-SEG. ATRAZADO
                        Case "100325": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegInm_Liq '100325 AMORT.-SEG. ADELANTADOS
                        
                        Case "100310": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_Portes 'portes (AMORT.-PORTES-100310)
                        Case "100314": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapPBP 'capital pbp (AMORT.-CAPITAL PBP-100314)
                        Case "100315": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntPBP 'interes pbp (AMORT.-INTERES PBP-100315)
                        Case "100307": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IMoVig 'interes moratorio vigente (AMORT-INT. MORAT. VIGENTE-100307)
                        Case "100308": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IMoVen 'interes moratorio vencido (AMORT.-INT MORAT. VENCIDO-100308)
                        Case "100319": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntCom 'interes compensatorio (AMORT.-INTERES COMPENSATORIO-100319)
                        Case "100309": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_GasCob 'gastos por cobranza (AMORT.-GASTOS COBRANZAS-100309)
                        Case "100311": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_OtrGas 'Otros Gastos (AMORT.-OTROS GASTOS-100311)
                        Case "100202": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ImpITF 'Importe de ITF (ITF DEPOSITADO-100202)
                        Case "100320": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ManCta 'Mantenimiento de Cuenta (AMORT.-COMISION RECAUDO-100320)
                        Case "100321": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CtaBnc 'Cuenta Bancaria (AMORT.-CTA. BANCO-100321)
                        Case "100201" '(IMPORTE DEPOSITADO-100201)
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = modprc_g_rst_Princi!CAJMOV_IMPTOT 'Importe Depositado
                           For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                              If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaBan(r_int_Cont01).CtaBan_CtaCtb
                                 Exit For
                              End If
                           Next r_int_Cont01
                     End Select
                     
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
                  End If
               End If
            Next r_int_Contad
            
            'Generar Numero de Asiento
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, "LM", r_arr_MatDet(r_int_PosMat).MatDet_NroLib)
            
            'Ingresar Cabecera
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_PosMat).MatDet_NroLib, r_int_NumAsi, Format(modprc_g_rst_Princi!CAJMOV_MONPAG, "000"), r_dbl_TipSbs, "O", Mid(r_arr_MatDet(r_int_PosMat).MatDet_DesCab + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000"), 1, 60), r_str_FecAsi, "1")
            
            'Ingresar Detalles
            r_int_NumIte = 1
            
            For r_int_Contad = 1 To UBound(r_arr_Matriz)
               If modprc_g_rst_Princi!CAJMOV_MONPAG = 1 Then
                  If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100202" Then
                     'Operacion de Truncado a 2 digitos por ITF
                     r_dbl_ImpSol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
                     r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))    'Truncar
                  Else
                     r_dbl_ImpSol = r_arr_Matriz(r_int_Contad).Matriz_import
                     r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSbs, "#######0.00"))
                  End If
                  
               ElseIf modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
                  If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100202" Then
                     'Operacion de Truncado a 2 digitos por ITF
                     r_dbl_ImpSol = CDbl(Format(gf_Truncar_Numero(gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import) * r_dbl_TipSun, 2), "########0.00"))     'Truncar
                     r_dbl_ImpDol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
                  Else
                     r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import * r_dbl_TipSbs, "########0.00"))
                     r_dbl_ImpDol = r_arr_Matriz(r_int_Contad).Matriz_import
                  End If
               End If
               
               'Acumulacion de Haber y Debe
               If r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "H" Then
                  r_dbl_Tothab_sol = r_dbl_Tothab_sol + r_dbl_ImpSol
               ElseIf r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "D" Then
                  r_dbl_Totdeb_sol = r_dbl_Totdeb_sol + r_dbl_ImpSol
               End If
               
               'Inserccion de las Cuentas al Detalle del Asiento
               If r_dbl_ImpSol <> 0 And r_dbl_ImpDol <> 0 Then
                  If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100201" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100320" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100321" Then
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, _
                                                          CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + r_arr_Matriz(r_int_Contad).Matriz_DesNot, 1, 60), r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV)))) 'Contabilizacion de Fecha en el campo Auxiliar
                  Else
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, _
                                                          r_str_FecAsi, Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + r_arr_Matriz(r_int_Contad).Matriz_DesNot, 1, 60), r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))                                                   'Contabilizacion de Fecha en el campo Auxiliar
                  End If
                  
                  r_int_NumIte = r_int_NumIte + 1
               End If
            Next r_int_Contad
            
            If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
               r_dbl_ImpSol = 0
               
               'Diferencia entre el DEBE y HABER para agregar la cuenta Diferencia por Tipo de Cambio
               If r_dbl_Totdeb_sol > r_dbl_Tothab_sol Then
                  r_str_Cuenta = r_str_CueGan
                  r_str_FlagDH = "H"
                  r_dbl_ImpSol = CDbl(Format(r_dbl_Totdeb_sol - r_dbl_Tothab_sol, "#####0.00"))
               End If
               
               If r_dbl_Totdeb_sol < r_dbl_Tothab_sol Then
                  r_str_Cuenta = r_str_CuePer
                  r_str_FlagDH = "D"
                  r_dbl_ImpSol = CDbl(Format(r_dbl_Tothab_sol - r_dbl_Totdeb_sol, "######0.00"))
               End If
               
               'Inserccion de la Cuenta por Diferencia de Tipo de Cambio en el Detalle del Asiento
               If r_dbl_ImpSol > 0 Then
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad - 1).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_str_Cuenta, r_str_FecAsi, Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + "DIF. TIP. CAM.", 1, 60), r_str_FlagDH, r_dbl_ImpSol, 0, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
               End If
            End If
         End If
  
         'actualizando tabla CRE_HIPMAE
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_ACUDVG = " & CStr(r_dbl_AcuDvg) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDVC = " & CStr(r_dbl_AcuDvc) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDIF = " & CStr(r_dbl_AcuDif) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_CAPVEN = HIPMAE_CAPVEN - " & CStr(r_dbl_Asi_CapVen) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         'Actualizando en OPE_CAJMOV para diferenciar los movimientos procesados en el mismo dia
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE OPE_CAJMOV "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET CAJMOV_CTBFLG = 1 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_USUMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_USUMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_NUMMOV = " & modprc_g_rst_Princi!CAJMOV_NUMMOV & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla OPE_CAJMOV - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         'Leyendo siguiente Movimiento
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Loop
   
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1005_OLD(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerIni As String, ByVal p_PerFin As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_CtbIni As String, ByVal p_CtbFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1005
'Descripción      :  Registro de Pagos de Cuotas de Créditos Hipotecarios
'Resumen          :  Contabilización de Pago de Cuotas de Créditos Hipotecarios
'F. Creación      :  16-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :  13-03-2010
'U. Actualización :  Jorge Luis Tacuche Mesia

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_arr_CtaBan()      As modprc_g_tpo_CtaBan
Dim r_arr_MatPro()      As modprc_g_tpo_MatPro
Dim r_arr_Matriz()      As modprc_g_tpo_Matriz
Dim r_arr_MatDet()      As modprc_g_tpo_MatDet
Dim r_arr_CtaPrd()      As modprc_g_tpo_CtaPrd
Dim r_str_FecPro        As String
Dim r_str_FecVct        As String
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_int_NumAsi        As Integer
Dim r_int_PosMat        As Integer
Dim r_int_AuxCon        As Integer
Dim r_rst_Genera        As ADODB.Recordset
Dim r_rst_Cuotas        As ADODB.Recordset
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_rst_CtaBan        As ADODB.Recordset
Dim r_rst_TipMat        As ADODB.Recordset
Dim r_rst_DifCam        As ADODB.Recordset
Dim r_rst_MatCtb        As ADODB.Recordset
Dim r_rst_MatCab        As ADODB.Recordset
Dim r_rst_MatDet        As ADODB.Recordset
Dim r_rst_CtaPrd        As ADODB.Recordset
Dim r_dbl_AcuDvg        As Double
Dim r_dbl_AcuDvc        As Double
Dim r_dbl_AcuDif        As Double
Dim r_dbl_TipCam        As Double
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_dbl_TipSbs        As Double
Dim r_dbl_TipSun        As Double
Dim r_int_OpeGnb        As Integer
Dim r_dbl_Asi_CapVig    As Double
Dim r_dbl_Asi_CapVen    As Double
Dim r_dbl_Asi_IntEfe    As Double
Dim r_dbl_Asi_IntDev    As Double
Dim r_dbl_Asi_IntVen    As Double
Dim r_dbl_Asi_IntDif    As Double
Dim r_dbl_Asi_SegDes    As Double
Dim r_dbl_Asi_SegInm    As Double
Dim r_dbl_Asi_Portes    As Double
Dim r_dbl_Asi_CapPBP    As Double
Dim r_dbl_Asi_IntPBP    As Double
Dim r_dbl_Asi_IntMor    As Double
Dim r_dbl_Asi_IntCom    As Double
Dim r_dbl_Asi_GasCob    As Double
Dim r_dbl_Asi_OtrGas    As Double
Dim r_dbl_Asi_ImpITF    As Double
Dim r_dbl_Asi_IMoVig    As Double
Dim r_dbl_Asi_IMoVen    As Double
Dim r_dbl_Asi_CVgGar    As Double
Dim r_dbl_Asi_CVgSGa    As Double
Dim r_dbl_Mes_IntEfe    As Double
Dim r_dbl_Mes_IntDev    As Double
Dim r_dbl_Mes_IntVen    As Double
Dim r_dbl_Mes_IntDif    As Double
Dim r_dbl_Mes_CapVig    As Double
Dim r_dbl_Mes_CapVen    As Double
Dim r_dbl_Mes_IMoVig    As Double
Dim r_dbl_Mes_IMoVen    As Double
Dim r_dbl_Mes_CVgGar    As Double
Dim r_dbl_Mes_CVgSGa    As Double
Dim r_int_TipGar        As Integer
Dim r_int_Cont01        As Integer
Dim r_int_SitCre        As Integer
Dim r_int_Refina        As Integer
Dim r_int_Judici        As Integer
Dim r_int_Castig        As Integer
Dim r_dbl_Totdeb_sol    As Double
Dim r_dbl_Totdeb_dol    As Double
Dim r_dbl_Tothab_sol    As Double
Dim r_dbl_Tothab_dol    As Double
Dim r_str_CueGan        As String
Dim r_str_CuePer        As String
Dim r_str_Cuenta        As String
Dim r_str_FlagDH        As String
Dim r_str_FecAsi        As String
Dim r_dbl_Asi_ManCta    As Double
Dim r_dbl_Asi_CtaBnc    As Double
   
   r_lng_NumReg = 0
   r_lng_TotReg = ff_ConAmo(p_FecIni, p_FecFin)
   p_BarPro.FloodPercent = 0
      
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1005"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
        
   '1.= Para obtener Cuentas de Diferencia de Cambio (Ganancia o Pérdida)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAGEN "
   g_str_Parame = g_str_Parame & " WHERE CTAGEN_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & " ORDER BY CTAGEN_CTACTB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_DifCam, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_DifCam.BOF And r_rst_DifCam.EOF) Then
      r_rst_DifCam.MoveFirst
      Do While Not r_rst_DifCam.EOF
         If r_rst_DifCam!CTAGEN_CODIDE = "01" Then
            r_str_CueGan = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         ElseIf r_rst_DifCam!CTAGEN_CODIDE = "02" Then
            r_str_CuePer = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         End If
         
         r_rst_DifCam.MoveNext
      Loop
   End If
   
   r_rst_DifCam.Close
   Set r_rst_DifCam = Nothing
   
   '2.= Para leer cuentas por cada Cuenta Bancaria
   ReDim r_arr_CtaBan(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN "
   g_str_Parame = g_str_Parame & " ORDER BY CTABAN_CODBAN, CTABAN_NUMCTA ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaBan, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaBan.BOF And r_rst_CtaBan.EOF) Then
      r_rst_CtaBan.MoveFirst
      Do While Not r_rst_CtaBan.EOF
         ReDim Preserve r_arr_CtaBan(UBound(r_arr_CtaBan) + 1)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CodBan = r_rst_CtaBan!CtaBan_CodBan
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_NumCta = Trim(r_rst_CtaBan!CtaBan_NumCta)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CtaCtb = Trim(r_rst_CtaBan!CtaBan_CtaCtb)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_TipCta = r_rst_CtaBan!CtaBan_TipCta
         
         r_rst_CtaBan.MoveNext
      Loop
   End If
   
   r_rst_CtaBan.Close
   Set r_rst_CtaBan = Nothing
   
   '3.= Para leer Matrices Contables de todos los Producto
   ReDim r_arr_MatDet(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATCAB "
   g_str_Parame = g_str_Parame & " WHERE MATCAB_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & "   AND MATCAB_TIPMAT = '100004' "
   g_str_Parame = g_str_Parame & " ORDER BY MATCAB_SITCRE, SUBSTR(MATCAB_CODMAT,1,3) ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatCab, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_MatCab.BOF And r_rst_MatCab.EOF) Then
      r_rst_MatCab.MoveFirst
      Do While Not r_rst_MatCab.EOF
         
         '4.= Matriz por cada producto
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATDET "
         g_str_Parame = g_str_Parame & " WHERE MATDET_CODMAT = '" & r_rst_MatCab!MATCAB_CODMAT & "' "
         g_str_Parame = g_str_Parame & " ORDER BY MATDET_CODMAT, MATDET_NUMITE ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatDet, 3) Then
            Exit Sub
         End If
         
         r_rst_MatDet.MoveFirst
         Do While Not r_rst_MatDet.EOF
            ReDim Preserve r_arr_MatDet(UBound(r_arr_MatDet) + 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodMat = r_rst_MatCab!MATCAB_CODMAT
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodPrd = Trim(Mid(r_rst_MatCab!MATCAB_CODMAT, 1, 3))
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesCab = Trim(r_rst_MatCab!MATCAB_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipMon = r_rst_MatCab!MATCAB_TIPMON
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_SitCre = r_rst_MatCab!MATCAB_SITCRE
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesDet = Trim(r_rst_MatDet!MATDET_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CtbCon = Trim(r_rst_MatDet!MATDET_CONCTB)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DebHab = Left(moddat_gf_Consulta_ParDes("255", CStr(r_rst_MatDet!MATDET_FLGDHB)), 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipCam = CInt(r_rst_MatDet!MATDET_TIPTCA)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_NroLib = r_rst_MatCab!MATCAB_CODLIB
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_OpeCon = Trim(r_rst_MatDet!MATDET_CONOPE)
            
            r_rst_MatDet.MoveNext
         Loop
         
         r_rst_MatDet.Close
         Set r_rst_MatDet = Nothing
         r_rst_MatCab.MoveNext
      Loop
   End If
   
   r_rst_MatCab.Close
   Set r_rst_MatCab = Nothing
   
   '5.= Para leer cuentas Contables de todos los productos
   ReDim r_arr_CtaPrd(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRD "
   g_str_Parame = g_str_Parame & " ORDER BY CTAPRD_CODPRD ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaPrd, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaPrd.BOF And r_rst_CtaPrd.EOF) Then
      r_rst_CtaPrd.MoveFirst
      Do While Not r_rst_CtaPrd.EOF
         ReDim Preserve r_arr_CtaPrd(UBound(r_arr_CtaPrd) + 1)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CodPrd = Trim(r_rst_CtaPrd!CtaPrd_CodPrd)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtbCon = Trim(r_rst_CtaPrd!CTAPRD_CONCTB)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_SitCre = Trim(r_rst_CtaPrd!CTAPRD_TIPCRE)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtaCtb = Trim(r_rst_CtaPrd!CtaPrd_CtaCtb)
         
         r_rst_CtaPrd.MoveNext
      Loop
   End If
   
   r_rst_CtaPrd.Close
   Set r_rst_CtaPrd = Nothing
      
   '6.= Consulta principal de amotizaciones por procesar
   modprc_g_str_CadEje = "SELECT CAJMOV_SUCMOV, CAJMOV_USUMOV, CAJMOV_FECMOV, CAJMOV_FECDEP, CAJMOV_NUMMOV, CAJMOV_NUMOPE, CAJMOV_ITFIMP, CAJMOV_MONPAG, CAJMOV_IMPTOT, CAJMOV_CODBAN, CAJMOV_NUMCTA, CAJMOV_IMPPAG FROM OPE_CAJMOV " & _
                         " WHERE CAJMOV_TIPMOV = 1102 " & _
                         "   AND CAJMOV_CTBFLG = 0 " & _
                         "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " " & _
                         "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         " ORDER BY CAJMOV_NUMMOV ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Fecha de Proceso
         r_str_FecPro = Format(p_FecFin, "dd/mm/yyyy")
         If CDate(gf_FormatoFecha(modprc_g_rst_Princi!CAJMOV_FECMOV)) > CDate(p_CtbFin) Then
            r_str_FecAsi = p_CtbFin
         Else
            r_str_FecAsi = gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))
         End If
         
         r_int_OpeGnb = 0
         r_dbl_Asi_CapVig = 0
         r_dbl_Asi_CVgGar = 0
         r_dbl_Asi_CVgSGa = 0
         r_dbl_Asi_CapVen = 0
         r_dbl_Asi_IntEfe = 0
         r_dbl_Asi_IntDev = 0
         r_dbl_Asi_IntVen = 0
         r_dbl_Asi_IntDif = 0
         r_dbl_Asi_SegDes = 0
         r_dbl_Asi_SegInm = 0
         r_dbl_Asi_Portes = 0
         r_dbl_Asi_CapPBP = 0
         r_dbl_Asi_IntPBP = 0
         r_dbl_Asi_IntMor = 0
         r_dbl_Asi_IntCom = 0
         r_dbl_Asi_GasCob = 0
         r_dbl_Asi_OtrGas = 0
         r_dbl_Asi_IMoVig = 0
         r_dbl_Asi_IMoVen = 0
         r_dbl_Asi_ImpITF = modprc_g_rst_Princi!CAJMOV_ITFIMP
         r_dbl_Tothab_sol = 0
         r_dbl_Totdeb_sol = 0
         
         '7.= Para obtener Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = "SELECT HIPMAE_ACUDVG, HIPMAE_ACUDVC, HIPMAE_ACUDIF, HIPMAE_TIPGAR, HIPMAE_SITACT, HIPMAE_REFINA, HIPMAE_JUDICI, HIPMAE_CASTIG FROM CRE_HIPMAE " & _
                               " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         r_dbl_AcuDvg = r_rst_Genera!HIPMAE_ACUDVG
         r_dbl_AcuDvc = r_rst_Genera!HIPMAE_ACUDVC
         r_dbl_AcuDif = r_rst_Genera!HIPMAE_ACUDIF
         r_int_TipGar = r_rst_Genera!HIPMAE_TIPGAR
         r_int_SitCre = r_rst_Genera!HIPMAE_SITACT
         r_int_Refina = r_rst_Genera!HIPMAE_REFINA
         r_int_Judici = r_rst_Genera!HIPMAE_JUDICI
         r_int_Castig = r_rst_Genera!HIPMAE_CASTIG
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         
         '8.= Para obtener Cuotas Pagadas de la operacion
         modprc_g_str_CadEje = "SELECT * FROM CRE_HIPPAG " & _
                               " WHERE HIPPAG_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' " & _
                               "   AND HIPPAG_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' " & _
                               "   AND HIPPAG_NUMMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV) & " " & _
                               "   AND HIPPAG_FECMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_FECMOV) & " " & _
                               " ORDER BY HIPPAG_NUMCUO ASC"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPPAG - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         Do While Not r_rst_Genera.EOF
            r_dbl_Mes_IntEfe = 0
            r_dbl_Mes_IntDev = 0
            r_dbl_Mes_IntVen = 0
            r_dbl_Mes_IntDif = 0
            r_dbl_Mes_CapVen = 0
            r_dbl_Mes_CapVig = 0
            r_dbl_Mes_IMoVig = 0
            r_dbl_Mes_IMoVen = 0
            r_dbl_Mes_CVgGar = 0
            r_dbl_Mes_CVgSGa = 0
            
            '9.= Para obtener Fecha de Vencimiento de Cuota que corresponde
            modprc_g_str_CadEje = "SELECT HIPCUO_FECVCT FROM CRE_HIPCUO " & _
                                  " WHERE HIPCUO_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' " & _
                                  "   AND HIPCUO_TIPCRO = 1 " & _
                                  "   AND HIPCUO_NUMCUO = " & CStr(r_rst_Genera!HIPPAG_NUMCUO) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Cuotas, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
            End If
            
            r_rst_Cuotas.MoveFirst
            r_str_FecVct = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECVCT))
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
            
            '10.= Determinando Tipo de Pago (Mes, Atrasado, Adelantado) y distribuyendo Interés Pagado (HIPPAG_INTERE)
            If CDate(r_str_FecVct) >= CDate(p_CtbIni) And CDate(r_str_FecVct) <= CDate(p_CtbFin) Then
               'Pago Mes
               If r_dbl_AcuDvc = 0 Then      'Acumulado Devengado Vencido
                  If r_rst_Genera!HIPPAG_INTERE >= r_dbl_AcuDvg Then
                     r_dbl_Mes_IntDev = r_dbl_AcuDvg
                     r_dbl_Mes_IntEfe = r_rst_Genera!HIPPAG_INTERE - r_dbl_Mes_IntDev
                     r_dbl_AcuDvg = 0
                  Else
                     r_dbl_Mes_IntDev = r_rst_Genera!HIPPAG_INTERE
                     r_dbl_AcuDvg = r_dbl_AcuDvg - r_dbl_Mes_IntDev
                  End If

           'MODIFICACIÒN AL 03/11/2016 - PARA OBTENER EL INTERES EFECTIVO DEL MES - CUADRE DE ASIENTOS
          Else
       r_dbl_Mes_IntEfe = r_rst_Genera!HIPPAG_INTERE
               End If
               
            ElseIf CDate(r_str_FecVct) < CDate(p_CtbIni) Then
               'Pago Atrasado
               If r_dbl_AcuDvc = 0 Then
                  If r_rst_Genera!HIPPAG_INTERE >= r_dbl_AcuDvg Then
                     r_dbl_Mes_IntDev = r_dbl_AcuDvg
                     r_dbl_Mes_IntEfe = r_rst_Genera!HIPPAG_INTERE - r_dbl_Mes_IntDev
                     r_dbl_AcuDvg = 0
                  Else
                     r_dbl_Mes_IntDev = r_rst_Genera!HIPPAG_INTERE
                     r_dbl_AcuDvg = r_dbl_AcuDvg - r_dbl_Mes_IntDev
                  End If
                  
                  r_dbl_Mes_IMoVig = r_dbl_Mes_IMoVig + r_rst_Genera!HIPPAG_INTMOR
               Else
                  'Si Crédito esta vencido
                  If r_rst_Genera!HIPPAG_INTERE >= r_dbl_AcuDvc Then
                     r_dbl_Mes_IntVen = r_dbl_AcuDvc
                     r_dbl_AcuDvc = 0
                  Else
                     r_dbl_Mes_IntVen = r_rst_Genera!HIPPAG_INTERE
                     r_dbl_AcuDvc = r_dbl_AcuDvc - r_dbl_Mes_IntVen
                  End If
                  
                  r_dbl_Mes_CapVen = r_rst_Genera!HIPPAG_CAPITA
                  r_dbl_Mes_IMoVen = r_dbl_Mes_IMoVen + r_rst_Genera!HIPPAG_INTMOR
               End If
               
            ElseIf CDate(r_str_FecVct) > CDate(p_CtbFin) Then
               'Pago Adelantado
               r_dbl_Mes_IntDif = r_rst_Genera!HIPPAG_INTERE
               r_dbl_AcuDif = r_dbl_AcuDif + r_dbl_Mes_IntDif
            Else
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al determinar Tipo de Pago - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
            End If
            
            r_dbl_Mes_CapVig = r_rst_Genera!HIPPAG_CAPITA - r_dbl_Mes_CapVen
            
            'Para determinar Tipo de Garantia
            If r_int_TipGar = 1 Or r_int_TipGar = 2 Then
               r_dbl_Mes_CVgGar = r_dbl_Mes_CapVig
            Else
               r_dbl_Mes_CVgSGa = r_dbl_Mes_CapVig
            End If
            
            'Actualizar en CRE_HIPPAG distribución de Intereses
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPPAG "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPPAG_DEVVIG = " & CStr(r_dbl_Mes_IntDev) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_DEVVEN = " & CStr(r_dbl_Mes_IntVen) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_INTDIF = " & CStr(r_dbl_Mes_IntDif) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_INTEFE = " & CStr(r_dbl_Mes_IntEfe) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CAPVIG = " & CStr(r_dbl_Mes_CapVig) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CAPVEN = " & CStr(r_dbl_Mes_CapVen) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_IMOVEN = " & CStr(r_dbl_Mes_IMoVen) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_IMOVIG = " & CStr(r_dbl_Mes_IMoVig) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CVGGAR = " & CStr(r_dbl_Mes_CVgGar) & ", "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_CVGSGA = " & CStr(r_dbl_Mes_CVgSGa) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPPAG_NUMOPE = '" & modprc_g_rst_Princi!CAJMOV_NUMOPE & "' AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_NUMCUO = " & CStr(r_rst_Genera!HIPPAG_NUMCUO) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_NUMMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPPAG_FECMOV = " & CStr(modprc_g_rst_Princi!CAJMOV_FECMOV) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPPAG - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
            End If
            
            r_dbl_Asi_CapVig = r_dbl_Asi_CapVig + r_dbl_Mes_CapVig
            r_dbl_Asi_CVgGar = r_dbl_Asi_CVgGar + r_dbl_Mes_CVgGar
            r_dbl_Asi_CVgSGa = r_dbl_Asi_CVgSGa + r_dbl_Mes_CVgSGa
            r_dbl_Asi_CapVen = r_dbl_Asi_CapVen + r_dbl_Mes_CapVen
            r_dbl_Asi_IntEfe = r_dbl_Asi_IntEfe + r_dbl_Mes_IntEfe
            r_dbl_Asi_IntDev = r_dbl_Asi_IntDev + r_dbl_Mes_IntDev
            r_dbl_Asi_IntVen = r_dbl_Asi_IntVen + r_dbl_Mes_IntVen
            r_dbl_Asi_IntDif = r_dbl_Asi_IntDif + r_dbl_Mes_IntDif
            r_dbl_Asi_SegDes = r_dbl_Asi_SegDes + r_rst_Genera!HIPPAG_DESORG
            r_dbl_Asi_SegInm = r_dbl_Asi_SegInm + r_rst_Genera!HIPPAG_VIVORG
            r_dbl_Asi_Portes = r_dbl_Asi_Portes + r_rst_Genera!HIPPAG_OTRORG
            r_dbl_Asi_CapPBP = r_dbl_Asi_CapPBP + r_rst_Genera!HIPPAG_CAPBBP
            r_dbl_Asi_IntPBP = r_dbl_Asi_IntPBP + r_rst_Genera!HIPPAG_INTBBP
            r_dbl_Asi_IntMor = r_dbl_Asi_IntMor + r_rst_Genera!HIPPAG_INTMOR
            r_dbl_Asi_IntCom = r_dbl_Asi_IntCom + r_rst_Genera!HIPPAG_INTCOM
            r_dbl_Asi_GasCob = r_dbl_Asi_GasCob + r_rst_Genera!HIPPAG_GASCOB
            r_dbl_Asi_OtrGas = r_dbl_Asi_OtrGas + r_rst_Genera!HIPPAG_OTRGAS
            r_dbl_Asi_IMoVig = r_dbl_Asi_IMoVig + r_dbl_Mes_IMoVig
            r_dbl_Asi_IMoVen = r_dbl_Asi_IMoVen + r_dbl_Mes_IMoVen
            
            r_rst_Genera.MoveNext
         Loop
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         
         'Actualizando en CRE_HIPMAE Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_ACUDVG = " & CStr(r_dbl_AcuDvg) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDVC = " & CStr(r_dbl_AcuDvc) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDIF = " & CStr(r_dbl_AcuDif) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_CAPVEN = HIPMAE_CAPVEN - " & CStr(r_dbl_Asi_CapVen) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         If (r_int_SitCre = 1) Then
            r_int_SitCre = 1
         ElseIf (r_int_SitCre = 5) And (r_int_Refina = 0) And (r_int_Judici = 0) And (r_int_Castig = 0) Then
            r_int_SitCre = 1
         End If
         
         If (r_int_Refina = 1) Then
            r_int_SitCre = 4
         ElseIf (r_int_Judici = 1) Then
            r_int_SitCre = 6
         ElseIf (r_int_Castig = 1) Then
            r_int_SitCre = 3
         End If
         
         If ff_Valida_OperacionGNB(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE), Format(r_str_FecVct, "YYYYMMDD")) Then
            'Operaciones GNB
            r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
            r_dbl_ImpSol = modprc_g_rst_Princi!CAJMOV_IMPTOT
            
            'Obtiene correlativo
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, "LM", 6)
            
            'Ingresa Cabecera
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, "001", r_dbl_TipSbs, "O", "DEPOSITO POR DEVOLVER A GNB " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000"), r_str_FecAsi, "1")
            
            'Ingresa Detalle
            r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(modprc_g_rst_Princi!CAJMOV_IMPTOT / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 1, "111301060201", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB", 1, 60), "D", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 2, "291807010103", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB", 1, 60), "H", r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            
            r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(modprc_g_dbl_ComBcoSol / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 3, "291807010103", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB - COMISION", 1, 60), "D", modprc_g_dbl_ComBcoSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, 6, r_int_NumAsi, 4, "111301060201", CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - DEVOLUCION GNB - COMISION", 1, 60), "H", modprc_g_dbl_ComBcoSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
         Else
            ReDim r_arr_Matriz(0)
            r_int_PosMat = 0
            
            For r_int_Contad = 1 To UBound(r_arr_MatDet)
               
               'Comparacion de Situacion de Credito con la Situacion de la Matriz
               If r_arr_MatDet(r_int_Contad).MatDet_SitCre = r_int_SitCre Then
                  
                  'Filtra las matrices del producto de la operacion en evaluacion
                  If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE), 1, 3) Then
                     r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecAsi, "yyyymmdd"), 1)
                     r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
                     
                     If r_int_PosMat = 0 Then
                        r_int_PosMat = r_int_Contad
                     End If
                     
                     ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
                     For r_int_AuxCon = 1 To UBound(r_arr_CtaPrd)
                        If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CodPrd = r_arr_MatDet(r_int_Contad).MatDet_CodPrd Then
                           If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtbCon = r_arr_MatDet(r_int_Contad).MatDet_CtbCon Then
                              If r_int_SitCre = r_arr_MatDet(r_int_Contad).MatDet_SitCre Then
                                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtaCtb
                              End If
                           End If
                        End If
                     Next r_int_AuxCon
                     
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = r_arr_MatDet(r_int_Contad).MatDet_DebHab
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_OpeCon = r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                     
                     For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                        If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_TipCta = r_arr_CtaBan(r_int_Cont01).CtaBan_TipCta
                           Exit For
                        End If
                     Next r_int_Cont01
                     
                     If r_arr_MatDet(r_int_Contad).MatDet_OpeCon = "100320" Or r_arr_MatDet(r_int_Contad).MatDet_OpeCon = "100321" Then
                        If r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_TipCta = "000002" Then
                           If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
                              r_dbl_Asi_ManCta = modprc_g_dbl_ComBcoDol
                              r_dbl_Asi_CtaBnc = modprc_g_dbl_ComBcoDol
                           Else
                              r_dbl_Asi_ManCta = modprc_g_dbl_ComBcoSol
                              r_dbl_Asi_CtaBnc = modprc_g_dbl_ComBcoSol
                           End If
                        Else
                           r_dbl_Asi_ManCta = 0
                           r_dbl_Asi_CtaBnc = 0
                        End If
                     End If
                     
                     Select Case r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                        Case "100301": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapVig 'Capital Vigente
                        Case "100317": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CVgGar 'Capital Vigente con Garantia
                        Case "100318": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CVgSGa 'Capital Vigente sin Garantia
                        Case "100303": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapVen 'Capital vencido
                        Case "100305": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntEfe 'interes efectivo
                        Case "100304": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntDev 'interes devengado
                        Case "100306": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntVen 'interes vencido
                        Case "100316": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntDif 'interes diferido
                        Case "100312": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegDes 'seguro desgravamen
                        Case "100313": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_SegInm 'seguro de inmueble
                        Case "100310": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_Portes 'portes
                        Case "100314": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapPBP 'capital pbp
                        Case "100315": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntPBP 'interes pbp
                        Case "100307": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IMoVig 'interes moratorio vigente
                        Case "100308": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IMoVen 'interes moratorio vencido
                        Case "100319": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_IntCom 'interes compensatorio
                        Case "100309": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_GasCob 'gastos por cobranza
                        Case "100311": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_OtrGas 'Otros Gastos
                        Case "100202": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ImpITF 'Importe de ITF
                        Case "100320": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ManCta 'Mantenimiento de Cuenta
                        Case "100321": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CtaBnc 'Cuenta Bancaria
                        Case "100201"
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = modprc_g_rst_Princi!CAJMOV_IMPTOT 'Importe Depositado
                           For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                              If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaBan(r_int_Cont01).CtaBan_CtaCtb
                                 Exit For
                              End If
                           Next r_int_Cont01
                     End Select
                     
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
                  End If
               End If
            Next r_int_Contad
            
            'Generar Numero de Asiento
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, "LM", r_arr_MatDet(r_int_PosMat).MatDet_NroLib)
            
            'Ingresar Cabecera
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_PosMat).MatDet_NroLib, r_int_NumAsi, Format(modprc_g_rst_Princi!CAJMOV_MONPAG, "000"), r_dbl_TipSbs, "O", Mid(r_arr_MatDet(r_int_PosMat).MatDet_DesCab + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000"), 1, 60), r_str_FecAsi, "1")
            
            'Ingresar Detalles
            r_int_NumIte = 1
            
            For r_int_Contad = 1 To UBound(r_arr_Matriz)
               If modprc_g_rst_Princi!CAJMOV_MONPAG = 1 Then
                  If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100202" Then
                     'Operacion de Truncado a 2 digitos por ITF
                     r_dbl_ImpSol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
                     r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))    'Truncar
                  Else
                     r_dbl_ImpSol = r_arr_Matriz(r_int_Contad).Matriz_import
                     r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSbs, "#######0.00"))
                  End If
                  
               ElseIf modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
                  If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100202" Then
                     'Operacion de Truncado a 2 digitos por ITF
                     r_dbl_ImpSol = CDbl(Format(gf_Truncar_Numero(gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import) * r_dbl_TipSun, 2), "########0.00"))     'Truncar
                     r_dbl_ImpDol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
                  Else
                     r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import * r_dbl_TipSbs, "########0.00"))
                     r_dbl_ImpDol = r_arr_Matriz(r_int_Contad).Matriz_import
                  End If
               End If
               
               'Acumulacion de Haber y Debe
               If r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "H" Then
                  r_dbl_Tothab_sol = r_dbl_Tothab_sol + r_dbl_ImpSol
               ElseIf r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "D" Then
                  r_dbl_Totdeb_sol = r_dbl_Totdeb_sol + r_dbl_ImpSol
               End If
               
               'Inserccion de las Cuentas al Detalle del Asiento
               If r_dbl_ImpSol <> 0 And r_dbl_ImpDol <> 0 Then
                  If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100201" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100320" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100321" Then
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + r_arr_Matriz(r_int_Contad).Matriz_DesNot, 1, 60), r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV)))) 'Contabilizacion de Fecha en el campo Auxiliar
                  Else
                     Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, r_str_FecAsi, Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + r_arr_Matriz(r_int_Contad).Matriz_DesNot, 1, 60), r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))                                                   'Contabilizacion de Fecha en el campo Auxiliar
                  End If
                  
                  r_int_NumIte = r_int_NumIte + 1
               End If
            Next r_int_Contad
            
            If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
               r_dbl_ImpSol = 0
               
               'Diferencia entre el DEBE y HABER para agregar la cuenta Diferencia por Tipo de Cambio
               If r_dbl_Totdeb_sol > r_dbl_Tothab_sol Then
                  r_str_Cuenta = r_str_CueGan
                  r_str_FlagDH = "H"
                  r_dbl_ImpSol = CDbl(Format(r_dbl_Totdeb_sol - r_dbl_Tothab_sol, "#####0.00"))
               End If
               
               If r_dbl_Totdeb_sol < r_dbl_Tothab_sol Then
                  r_str_Cuenta = r_str_CuePer
                  r_str_FlagDH = "D"
                  r_dbl_ImpSol = CDbl(Format(r_dbl_Tothab_sol - r_dbl_Totdeb_sol, "######0.00"))
               End If
               
               'Inserccion de la Cuenta por Diferencia de Tipo de Cambio en el Detalle del Asiento
               If r_dbl_ImpSol > 0 Then
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad - 1).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_str_Cuenta, r_str_FecAsi, Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + "DIF. TIP. CAM.", 1, 60), r_str_FlagDH, r_dbl_ImpSol, 0, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
               End If
            End If
         End If
         
         'Actualizando en OPE_CAJMOV para diferenciar los movimientos procesados en el mismo dia
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE OPE_CAJMOV "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET CAJMOV_CTBFLG = 1 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_USUMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_USUMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_NUMMOV = " & modprc_g_rst_Princi!CAJMOV_NUMMOV & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla OPE_CAJMOV - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         'Leyendo siguiente Movimiento
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Loop
   
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1006(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerIni As String, ByVal p_PerFin As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1006
'Descripción      :  Traslado de Ingresos Diferidos (Asientos en Edpymebank)
'Resumen          :  Contabilización de Ingresos Diferidos (Asientos en Edpymebank)
'F. Creación      :  17-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg           As Long
Dim r_lng_TotReg           As Long
Dim r_arr_LogPro()         As modprc_g_tpo_LogPro
Dim r_str_FecPro           As String
Dim r_dbl_TipCam           As Double
Dim r_rst_Genera           As ADODB.Recordset
Dim r_rst_Grabar           As ADODB.Recordset
Dim r_dbl_AcuDvg           As Double
Dim r_dbl_AcuDif           As Double
Dim r_dbl_Mes_IntEfe       As Double
Dim r_dbl_Mes_IntDev       As Double
Dim r_dbl_Mes_IntDif       As Double
Dim r_int_LibCon           As Integer
Dim r_str_CtaDeb           As String
Dim r_str_CtaHab           As String
Dim r_str_CtaHb1           As String
Dim r_str_Origen           As String
Dim r_str_TipNot           As String
Dim r_int_NumAsi           As Integer
Dim r_int_TipMon           As Integer
Dim r_dbl_Lin_ImpDeb_Sol   As Double
Dim r_dbl_Lin_ImpHab_Sol   As Double
Dim r_dbl_Lin_ImpHb1_Sol   As Double
Dim r_dbl_Lin_ImpDeb_Dol   As Double
Dim r_dbl_Lin_ImpHab_Dol   As Double
Dim r_dbl_Lin_ImpHb1_Dol   As Double
Dim r_str_CodPrd           As String
Dim r_dbl_Imp_AjuSol       As Double
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   'Validar que p_FecIni y p_FecFin pertenezca al Período Activo
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1006"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_PerFin), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Leyendo Cursor Principal
   modprc_g_str_CadEje = "SELECT * " & _
                         "  FROM CRE_HIPCUO " & _
                         " WHERE HIPCUO_FECVCT >= " & Format(CDate(p_FecIni), "yyyymmdd") & " " & _
                         "   AND HIPCUO_FECVCT <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         "   AND HIPCUO_FECPAG <  " & Format(CDate(p_PerIni), "yyyymmdd") & " " & _
                         "   AND HIPCUO_FECPAG >  0 AND HIPCUO_INTPAG > 0 AND HIPCUO_TIPCRO = 1 " & _
                         " ORDER BY HIPCUO_NUMOPE ASC, HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         'Para obtener Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = "SELECT HIPMAE_ACUDVG, HIPMAE_ACUDVC, HIPMAE_ACUDIF, HIPMAE_MONEDA, HIPMAE_CODPRD " & _
                               "  FROM CRE_HIPMAE " & _
                               " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!HIPCUO_NUMOPE) & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPCUO_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         r_dbl_AcuDvg = r_rst_Genera!HIPMAE_ACUDVG
         r_dbl_AcuDif = r_rst_Genera!HIPMAE_ACUDIF
         r_int_TipMon = r_rst_Genera!HIPMAE_MONEDA
         r_str_CodPrd = Trim(r_rst_Genera!HIPMAE_CODPRD)
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
            
         r_dbl_Mes_IntEfe = 0
         r_dbl_Mes_IntDev = 0
      
         'Distribuyendo Interés Diferido
         If modprc_g_rst_Princi!HIPCUO_INTPAG >= r_dbl_AcuDvg Then
            r_dbl_Mes_IntDev = r_dbl_AcuDvg
            r_dbl_Mes_IntEfe = modprc_g_rst_Princi!HIPCUO_INTPAG - r_dbl_Mes_IntDev
            r_dbl_AcuDvg = 0
         Else
            r_dbl_Mes_IntDev = modprc_g_rst_Princi!HIPCUO_INTPAG
            r_dbl_AcuDvg = r_dbl_AcuDvg - r_dbl_Mes_IntDev
         End If
         
         r_dbl_AcuDif = r_dbl_AcuDif - modprc_g_rst_Princi!HIPCUO_INTPAG
         
         'Actualizando en CRE_HIPMAE Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_ACUDVG = " & CStr(r_dbl_AcuDvg) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPMAE_ACUDIF = " & CStr(r_dbl_AcuDif) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPCUO_NUMOPE & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         If modprc_g_rst_Princi!HIPCUO_INTPAG > 0 Then
            'Generar Asiento Contable
            r_str_CtaDeb = ""
            r_str_CtaHab = ""
         
            Select Case r_str_CodPrd
               Case "001"
                  r_str_CtaDeb = "292102010101"
                  r_str_CtaHab = "512401042401"
                  r_str_CtaHb1 = "142804010101"
               
               Case "002"
                  r_str_CtaDeb = "292102010101"
                  r_str_CtaHab = "512401040601"
                  r_str_CtaHb1 = "142804010101"
               
               Case "003"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042501"
                  r_str_CtaHb1 = "141804010101"
                  
               Case "004"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042301"
                  r_str_CtaHb1 = "141804010101"
               
               Case "006"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401040601"
                  r_str_CtaHb1 = "141804010101"
               
               Case "007"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "009"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042305"
                  r_str_CtaHb1 = "141804010101"
               
               Case "010"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042304"
                  r_str_CtaHb1 = "141804010101"
                  
               Case "011"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401040601"
                  r_str_CtaHb1 = "141804010101"
               
               Case "012"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401040601"
                  r_str_CtaHb1 = "141804010101"
               
               Case "013"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "014"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "015"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "016"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "017"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"

               Case "018"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"

               Case "019"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "021"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"

               Case "022"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"

               Case "023"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"

               Case "024"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"

               Case "025"
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
            End Select
         
            Select Case r_int_TipMon
               Case 1
                  r_dbl_Lin_ImpDeb_Sol = modprc_g_rst_Princi!HIPCUO_INTPAG
                  r_dbl_Lin_ImpHab_Sol = r_dbl_Mes_IntEfe
                  r_dbl_Lin_ImpHb1_Sol = r_dbl_Mes_IntDev
                  
                  r_dbl_Lin_ImpDeb_Dol = CDbl(Format(modprc_g_rst_Princi!HIPCUO_INTPAG / r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHab_Dol = CDbl(Format(r_dbl_Mes_IntEfe / r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHb1_Dol = CDbl(Format(r_dbl_Mes_IntDev / r_dbl_TipCam, "#####0.00"))
                  
               Case 2
                  r_dbl_Lin_ImpDeb_Sol = CDbl(Format(modprc_g_rst_Princi!HIPCUO_INTPAG * r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHab_Sol = CDbl(Format(r_dbl_Mes_IntEfe * r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHb1_Sol = CDbl(Format(r_dbl_Mes_IntDev * r_dbl_TipCam, "#####0.00"))
                  
                  r_dbl_Lin_ImpDeb_Dol = modprc_g_rst_Princi!HIPCUO_INTPAG
                  r_dbl_Lin_ImpHab_Dol = r_dbl_Mes_IntEfe
                  r_dbl_Lin_ImpHb1_Dol = r_dbl_Mes_IntDev
            End Select
         
            r_str_Origen = "LM"
            r_int_LibCon = 1
            r_str_TipNot = "O"
            
            'Obteniendo Nro. de Asiento
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_LibCon)
            
            'Insertar en CNTBL_ASIENTO
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_LibCon, r_int_NumAsi, "001", r_dbl_TipCam, r_str_TipNot, modprc_g_rst_Princi!HIPCUO_NUMOPE & " - DIFERIDO: " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), p_FecFin, "1")
            
            'Insertar en CNTBL_ASIENTO_DET
            Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_LibCon, r_int_NumAsi, 1, r_str_CtaDeb, p_FecFin, modprc_g_rst_Princi!HIPCUO_NUMOPE & " - DIFERIDO " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "D", r_dbl_Lin_ImpDeb_Sol, r_dbl_Lin_ImpDeb_Dol)
            Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_LibCon, r_int_NumAsi, 2, r_str_CtaHab, p_FecFin, modprc_g_rst_Princi!HIPCUO_NUMOPE & " - DIFERIDO " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "H", r_dbl_Lin_ImpHab_Sol, r_dbl_Lin_ImpHab_Dol)
            Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_LibCon, r_int_NumAsi, 3, r_str_CtaHb1, p_FecFin, modprc_g_rst_Princi!HIPCUO_NUMOPE & " - DIFERIDO " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "H", r_dbl_Lin_ImpHb1_Sol, r_dbl_Lin_ImpHb1_Dol)
            
            'Ajuste por Diferencia de Tipo de Cambio
            If r_int_TipMon = 2 Then
               If r_dbl_Lin_ImpDeb_Sol > (r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol) Then
                  r_dbl_Imp_AjuSol = CDbl(Format(r_dbl_Lin_ImpDeb_Sol - (r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol), "######0.00"))
                  
                  If r_dbl_Imp_AjuSol > 0 Then
                     Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_LibCon, r_int_NumAsi, 4, "512804090101", p_FecFin, "AJUSTE DIF. TIPO CAMBIO", "H", r_dbl_Imp_AjuSol, 0)
                  End If
               ElseIf r_dbl_Lin_ImpDeb_Sol < (r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol) Then
                  r_dbl_Imp_AjuSol = CDbl(Format((r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol) - r_dbl_Lin_ImpDeb_Sol, "######0.00"))
               
                  If r_dbl_Imp_AjuSol > 0 Then
                     Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_LibCon, r_int_NumAsi, 4, "412804090101", p_FecFin, "AJUSTE DIF. TIPO CAMBIO", "D", r_dbl_Imp_AjuSol, 0)
                  End If
               End If
            End If
         End If
      
         'Leyendo siguiente cuota
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
      Loop
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1007(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1007
'Descripción      :  Proceso Unico de Nivelación de Ingresos Diferidos
'Resumen          :  Nivelación de Ingresos Diferidos
'F. Creación      :  17-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_FecPro        As String
Dim r_rst_Grabar        As ADODB.Recordset
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1007"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Leyendo Cursor Principal
   modprc_g_str_CadEje = "SELECT HIPPAG_NUMOPE, SUM(HIPPAG_INTERE) AS INTDIF FROM CRE_HIPPAG A, CRE_HIPCUO B, CRE_HIPMAE C " & _
                         " WHERE HIPPAG_NUMOPE = HIPCUO_NUMOPE AND HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPPAG_NUMCUO = HIPCUO_NUMCUO " & _
                         "   AND HIPMAE_SITUAC = 2 AND HIPCUO_TIPCRO = 1 AND HIPPAG_INTERE > 0 " & _
                         "   AND HIPPAG_FECPAG <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         "   AND HIPCUO_FECVCT >= " & Format(CDate(p_FecFin) + CDate(1), "yyyymmdd") & " " & _
                         " GROUP BY HIPPAG_NUMOPE " & _
                         " ORDER BY HIPPAG_NUMOPE ASC "
   
'   modprc_g_str_CadEje = "SELECT HIPPAG_NUMOPE, SUM(HIPPAG_INTERE) AS INTDIF FROM CRE_HIPPAG A, CRE_HIPCUO B, CRE_HIPMAE C WHERE " & _
                         "HIPPAG_NUMOPE = HIPCUO_NUMOPE AND HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPPAG_NUMCUO = HIPCUO_NUMCUO AND " & _
                         "HIPMAE_SITUAC = 2 AND HIPCUO_TIPCRO = 1 AND HIPPAG_INTERE >  0 AND " & _
                         "HIPPAG_FECPAG >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND " & _
                         "HIPPAG_FECPAG <= " & Format(CDate(p_FecFin), "yyyymmdd") & " AND " & _
                         "HIPCUO_FECVCT >= " & Format(CDate(p_FecFin) + CDate(1), "yyyymmdd") & " " & _
                         "GROUP BY HIPPAG_NUMOPE " & _
                         "ORDER BY HIPPAG_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPPAG, CRE_HIPCUO, CRE_HIPMAE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         'Actualizando en CRE_HIPMAE Saldo Acumulado de Interés Diferido
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_ACUDIF = " & CStr(modprc_g_rst_Princi!INTDIF) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPPAG_NUMOPE & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPPAG_NUMOPE) & " .")
         End If
         
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
      Loop
      
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tablas CRE_HIPPAG, CRE_HIPCUO, CRE_HIPMAE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, 0, 0, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1008(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1008
'Descripción      :  Calificación de Clientes
'Resumen          :  Calificación de Clientes
'F. Creación      :  31-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_FecPro        As String
Dim r_rst_CreHip        As ADODB.Recordset
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_rst_RccDet        As ADODB.Recordset
Dim r_rst_ReFina        As ADODB.Recordset
Dim r_dbl_TipCam_Dol    As Double
Dim r_dbl_TotDeu        As Double
Dim r_dbl_TotAli        As Double
Dim r_int_ClaCli        As Integer
Dim r_int_ClaAli        As Integer
Dim r_int_ClaPrv        As Integer
Dim r_int_MesAnt        As Integer
Dim r_int_AnoAnt        As Integer
Dim r_int_ValCla        As Integer
Dim r_str_FecIni        As String
Dim r_str_FecFin        As String
Dim r_int_FlgPag        As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1008"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   If p_PerMes = 1 Then
      r_int_MesAnt = 12
      r_int_AnoAnt = p_PerAno - 1
   Else
      r_int_MesAnt = p_PerMes - 1
      r_int_AnoAnt = p_PerAno
   End If
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   '************************************
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = "SELECT COUNT(DATGEN_TIPDOC) AS TOTREG FROM CLI_DATGEN WHERE DATGEN_TIPCLI = 1 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CLI_DATGEN.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   '*********************************************
   'Leyendo Cursor Principal (Personas Naturales)
   modprc_g_str_CadEje = "SELECT DATGEN_TIPDOC, DATGEN_NUMDOC FROM CLI_DATGEN WHERE DATGEN_TIPCLI = 1 "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CLI_DATGEN.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         r_int_ClaCli = -1
         r_dbl_TotDeu = 0
         r_int_ValCla = 0
         
         '******************************************
         'Buscando Créditos Hipotecarios del Cliente para Clasificacion Interna
         modprc_g_str_CadEje = "SELECT HIPCIE_NUMOPE, HIPCIE_CLACRE, HIPCIE_TIPMON, HIPCIE_SALCAP, HIPCIE_SALCON " & _
                               "  FROM CRE_HIPCIE " & _
                               " WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " " & _
                               "   AND HIPCIE_PERANO = " & CStr(p_PerAno) & " " & _
                               "   AND HIPCIE_TDOCLI = " & CStr(modprc_g_rst_Princi!DATGEN_TIPDOC) & " " & _
                               "   AND HIPCIE_NDOCLI = '" & Trim(modprc_g_rst_Princi!DATGEN_NUMDOC) & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_CreHip, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE - Cliente: " & CStr(modprc_g_rst_Princi!DATGEN_TIPDOC) & "-" & Trim(modprc_g_rst_Princi!DATGEN_NUMDOC))
            r_rst_CreHip.Close
            Set r_rst_CreHip = Nothing
            Exit Sub
         End If
         
         If Not (r_rst_CreHip.BOF And r_rst_CreHip.EOF) Then
            r_rst_CreHip.MoveFirst
            
            'Determina saldo de la deuda en soles y la peor clasificacion del cliente
            Do While Not r_rst_CreHip.EOF
               If r_rst_CreHip!HIPCIE_CLACRE > r_int_ClaCli Then
                  r_int_ClaCli = r_rst_CreHip!HIPCIE_CLACRE
               End If
               
               If r_rst_CreHip!HIPCIE_TIPMON = 1 Then
                  r_dbl_TotDeu = r_dbl_TotDeu + r_rst_CreHip!HIPCIE_SALCAP + r_rst_CreHip!HIPCIE_SALCON
               ElseIf r_rst_CreHip!HIPCIE_TIPMON = 2 Then
                  r_dbl_TotDeu = r_dbl_TotDeu + CDbl(Format((r_rst_CreHip!HIPCIE_SALCAP + r_rst_CreHip!HIPCIE_SALCON) * r_dbl_TipCam_Dol, "#########0.00"))
               End If
               
               'Actualiza Clasificacion para Clientes Refinanciados - ABAD LOLI
               If r_rst_CreHip!HIPCIE_NUMOPE = "0030800036" Then
                  'Obtiene Clasificacion del periodo anterior
                  modprc_g_str_CadEje = "SELECT HIPCIE_CLACLI " & _
                                        "  FROM CRE_HIPCIE " & _
                                        " WHERE HIPCIE_PERMES = " & CStr(r_int_MesAnt) & " " & _
                                        "   AND HIPCIE_PERANO = " & CStr(r_int_AnoAnt) & " " & _
                                        "   AND HIPCIE_NUMOPE = '" & r_rst_CreHip!HIPCIE_NUMOPE & "' "
                  
                  If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_ReFina, 3) Then
                     r_rst_ReFina.Close
                     Set r_rst_ReFina = Nothing
                  End If
                  
                  If Not (r_rst_ReFina.BOF And r_rst_ReFina.EOF) Then
                     r_rst_ReFina.MoveFirst
                     r_int_ClaCli = r_rst_ReFina!HIPCIE_CLACLI
                  End If
                  
                  r_rst_ReFina.Close
                  Set r_rst_ReFina = Nothing
                  
                  'La Fecha de Refinanciacion es 02/04/2012
                  If p_PerMes = 5 Or p_PerMes = 11 Then
                     'Determina rango de fechas
                     If p_PerMes = 5 Then
                        r_str_FecIni = CStr(CInt(p_PerAno - 1)) & "1101"
                        r_str_FecFin = CStr(CInt(p_PerAno)) & "0430"
                     End If
                     If p_PerMes = 11 Then
                        r_str_FecIni = CStr(CInt(p_PerAno)) & "0501"
                        r_str_FecFin = CStr(CInt(p_PerAno)) & "1031"
                     End If
                     
                     'Obtiene cuotas a evaluar
                     modprc_g_str_CadEje = "SELECT HIPCUO_NUMCUO, HIPCUO_SITUAC, HIPCUO_FECVCT, HIPCUO_FECPAG " & _
                                           "  FROM CRE_HIPCUO " & _
                                           " WHERE HIPCUO_NUMOPE = " & r_rst_CreHip!HIPCIE_NUMOPE & " " & _
                                           "   AND HIPCUO_TIPCRO = 1 " & _
                                           "   AND HIPCUO_FECVCT >= " & r_str_FecIni & " " & _
                                           "   AND HIPCUO_FECVCT <= " & r_str_FecFin & " "
                     
                     If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_ReFina, 3) Then
                        r_rst_ReFina.Close
                        Set r_rst_ReFina = Nothing
                     End If
                     
                     If Not (r_rst_ReFina.BOF And r_rst_ReFina.EOF) Then
                        r_rst_ReFina.MoveFirst
                        r_int_FlgPag = 0
                        
                        Do While Not r_rst_ReFina.EOF
                           If r_rst_ReFina!HIPCUO_FECPAG > r_rst_ReFina!HIPCUO_FECVCT Then
                              r_int_FlgPag = 1
                           End If
                           r_rst_ReFina.MoveNext
                        Loop
                        
                        If r_int_FlgPag = 0 Then
                           If r_int_ClaCli > 0 Then
                              r_int_ClaCli = r_int_ClaCli - 1
                           Else
                              r_int_ClaCli = 0
                           End If
                        End If
                        If r_int_FlgPag = 1 Then
                           If r_int_ClaCli < 4 Then
                              r_int_ClaCli = r_int_ClaCli + 1
                           Else
                              r_int_ClaCli = 4
                           End If
                        End If
                     End If
                     
                     r_rst_ReFina.Close
                     Set r_rst_ReFina = Nothing
                  End If
               End If
               
               r_rst_CreHip.MoveNext
            Loop
         End If
         
         r_rst_CreHip.Close
         Set r_rst_CreHip = Nothing
         
         If r_int_ClaCli > -1 Then
            r_int_ClaAli = r_int_ClaCli
            r_dbl_TotAli = r_dbl_TotDeu
            
            '******************************************************
            'Buscando en RCC para Clasificar con Alineación Externa
            modprc_g_str_CadEje = "SELECT RCCDET_CODEMP, RCCDET_CLASIF, SUM(RCCDET_MTOSOL) AS DEUSOL, SUM(RCCDET_MTODOL) AS DEUDOL " & _
                                  "  FROM CLI_RCCDET " & _
                                  " WHERE RCCDET_TIPDOC = " & CStr(modprc_g_rst_Princi!DATGEN_TIPDOC) & " " & _
                                  "   AND RCCDET_NUMDOC = '" & Trim(modprc_g_rst_Princi!DATGEN_NUMDOC) & "' " & _
                                  "   AND RCCDET_PERMES = " & CStr(r_int_MesAnt) & " " & _
                                  "   AND RCCDET_PERANO = " & CStr(r_int_AnoAnt) & " " & _
                                  "   AND RCCDET_CLASIF <> 8 " & _
                                  " GROUP BY RCCDET_CODEMP, RCCDET_CLASIF " & _
                                  " ORDER BY RCCDET_CODEMP ASC, RCCDET_CLASIF DESC "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_RccDet, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CLI_RCCDET - Cliente: " & CStr(modprc_g_rst_Princi!DATGEN_TIPDOC) & "-" & Trim(modprc_g_rst_Princi!DATGEN_NUMDOC))
               r_rst_RccDet.Close
               Set r_rst_RccDet = Nothing
               Exit Sub
            End If
            
            If Not (r_rst_RccDet.BOF And r_rst_RccDet.EOF) Then
               'Sumariza las Deudas del cliente en el Sistema Financiero
               r_rst_RccDet.MoveFirst
               Do While Not r_rst_RccDet.EOF
                  r_dbl_TotAli = r_dbl_TotAli + r_rst_RccDet!DEUSOL + r_rst_RccDet!DEUDOL
                  r_rst_RccDet.MoveNext
               Loop
               
               'Obtiene la peor clasificacion reportada del cliente
               r_rst_RccDet.MoveFirst
               Do While Not r_rst_RccDet.EOF
                  If r_rst_RccDet!RCCDET_CLASIF > r_int_ClaAli Then
                     If CDbl(Format((r_rst_RccDet!DEUSOL + r_rst_RccDet!DEUDOL) / r_dbl_TotAli * 100, "##0.00")) >= 20 And r_rst_RccDet!DEUSOL + r_rst_RccDet!DEUDOL > 100 Then
                        r_int_ClaAli = r_rst_RccDet!RCCDET_CLASIF
                     End If
                  End If
                  
                  r_rst_RccDet.MoveNext
               Loop
            End If
            
            r_rst_RccDet.Close
            Set r_rst_RccDet = Nothing
            
            'Determina Clasificacion para la provision
            If r_int_ClaCli = r_int_ClaAli Then
               r_int_ClaPrv = r_int_ClaCli
            ElseIf r_int_ClaCli > r_int_ClaAli Then
               r_int_ClaPrv = r_int_ClaCli
            Else
               If r_int_ClaAli = 0 Or r_int_ClaAli = 1 Or r_int_ClaAli = 2 Then
                  r_int_ClaPrv = r_int_ClaCli
               Else
                  If r_int_ClaCli = 0 Then
                     r_int_ClaPrv = r_int_ClaAli - 1
                  Else
                     r_int_ClaPrv = r_int_ClaAli
                  End If
               End If
            End If
            
            '**********************************
            'If r_int_ClaAli > r_int_ClaCli Then
            '   r_int_ClaPrv = r_int_ClaAli - 1
            'Else
            '   r_int_ClaPrv = r_int_ClaAli
            'End If
            '**********************************
            
            'Actualizando Clasificación de Cliente en CRE_HIPCIE
            modprc_g_str_CadEje = "UPDATE CRE_HIPCIE " & _
                                  "   SET HIPCIE_CLACLI = " & CStr(r_int_ClaCli) & ", " & _
                                  "       HIPCIE_CLAALI = " & CStr(r_int_ClaAli) & ", " & _
                                  "       HIPCIE_CLAPRV = " & CStr(r_int_ClaPrv) & "  " & _
                                  " WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " " & _
                                  "   AND HIPCIE_PERANO = " & CStr(p_PerAno) & " " & _
                                  "   AND HIPCIE_TDOCLI = " & CStr(modprc_g_rst_Princi!DATGEN_TIPDOC) & " " & _
                                  "   AND HIPCIE_NDOCLI = '" & Trim(modprc_g_rst_Princi!DATGEN_NUMDOC) & "' "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPCIE - Cliente: " & CStr(modprc_g_rst_Princi!DATGEN_TIPDOC) & "-" & Trim(modprc_g_rst_Princi!DATGEN_NUMDOC))
            End If
         End If
         
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
      
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CLI_DATGEN.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1009(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1009
'Descripción      :  Cálculo de Provisiones
'Resumen          :  Calculo de Provisiones
'F. Creación      :  31-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_arr_DetGar()      As modprc_g_tpo_DetGar
Dim r_arr_TipPrv()      As modprc_g_tpo_TipPrv
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_rst_PrvTim        As ADODB.Recordset
Dim r_int_ClaGar        As Integer
Dim r_int_Contad        As Integer
Dim r_int_MesAnt        As Integer
Dim r_int_AnoAnt        As Integer
Dim r_dbl_PrvVol        As Double
Dim r_dbl_PrvGen        As Double
Dim r_dbl_PrvGen_RC     As Double
Dim r_dbl_PrvEsp        As Double
Dim r_dbl_PrvCic        As Double
Dim r_dbl_PrvCic_RC     As Double
Dim r_dbl_PrvCam        As Double
Dim r_dbl_GtoJud        As Double
Dim r_dbl_OtrGas        As Double
Dim r_dbl_PerFmv        As Double
Dim r_int_MonCom        As Integer
Dim r_dbl_ValCom        As Double
Dim r_dbl_MtoCga        As Double
Dim r_dbl_MtoSga        As Double
Dim r_dbl_CobFmv        As Double
Dim r_dbl_CobFmv_RC     As Double
Dim r_dbl_Porce1        As Double
Dim r_dbl_Porce2        As Double
Dim r_dbl_TerSal        As Double
Dim r_dbl_TerPer        As Double
Dim r_int_TipGar        As Integer
Dim r_int_OrgGar        As Integer
Dim r_dbl_ValRea        As Double

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1009"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   '*********************************************************
   '1.= Leer Tablas de Provisiones para Créditos Hipotecarios
   modprc_g_str_CadEje = "SELECT * FROM CTB_TIPPRV WHERE TIPPRV_CLACRE = '13' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CTB_TIPPRV.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   modprc_g_rst_Princi.MoveFirst
   ReDim r_arr_TipPrv(0)
   Do While Not modprc_g_rst_Princi.EOF
      ReDim Preserve r_arr_TipPrv(UBound(r_arr_TipPrv) + 1)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_TipPrv = CInt(modprc_g_rst_Princi!TipPrv_TipPrv)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_CodCla = CInt(modprc_g_rst_Princi!TIPPRV_CLFCRE)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_ClaGar = CInt(modprc_g_rst_Princi!TipPrv_ClaGar)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_Porcen = modprc_g_rst_Princi!TipPrv_Porcen
      modprc_g_rst_Princi.MoveNext
   Loop
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   '***************************
   '2.= Leer Tabla de Garantías
   modprc_g_str_CadEje = "SELECT * FROM CTB_DETGAR "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CTB_DETGAR.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   modprc_g_rst_Princi.MoveFirst
   ReDim r_arr_DetGar(0)
   Do While Not modprc_g_rst_Princi.EOF
      ReDim Preserve r_arr_DetGar(UBound(r_arr_DetGar) + 1)
      r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_Codigo = CInt(modprc_g_rst_Princi!DetGar_Codigo)
      r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_ClaGar = CInt(modprc_g_rst_Princi!DetGar_ClaGar)
      modprc_g_rst_Princi.MoveNext
   Loop
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   '****************************************
   '3.= Para sacar Total de Registros a leer
   modprc_g_str_CadEje = "SELECT COUNT(HIPCIE_NUMOPE) AS TOTREG FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "

   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE (1).")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   '****************************************************
   '4.= Leyendo Cursor Principal (Créditos Hipotecarios)
   modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "
   'modprc_g_str_CadEje = modprc_g_str_CadEje & " AND HIPCIE_NUMOPE = '0221700520'"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE (2).")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         r_int_TipGar = modprc_g_rst_Princi!HIPCIE_TIPGAR
         r_int_OrgGar = modprc_g_rst_Princi!HIPCIE_TIPGAR
         r_dbl_ValRea = 0
         
         '********** OBTIENE EL VALOR DE REALIZACION **********
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM TRA_EVATAS "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE EVATAS_NUMSOL = (SELECT HIPMAE_NUMSOL FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "')"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_PrvTim, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al obtener la tasación")
         End If
         
         If Not (r_rst_PrvTim.BOF And r_rst_PrvTim.EOF) Then
            r_rst_PrvTim.MoveFirst
            r_dbl_ValRea = r_rst_PrvTim!EVATAS_VALREA_INM + r_rst_PrvTim!EVATAS_VALREA_ES1 + r_rst_PrvTim!EVATAS_VALREA_ES2 + r_rst_PrvTim!EVATAS_VALREA_DEP
         End If
         
         r_rst_PrvTim.Close
         Set r_rst_PrvTim = Nothing
         
         
         '********** DETERMINA SI TIENE CLASICACION DUDOSA POR MAS DE 36 MESES **********
         If modprc_g_rst_Princi!HIPCIE_CLACLI = 3 Then
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT COUNT(*) AS CONTADOR "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM (SELECT HIPCIE_CLACLI "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                  FROM CRE_HIPCIE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                 WHERE HIPCIE_PERMES > 0 "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                   AND HIPCIE_PERANO > 2014 "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                   AND HIPCIE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "' "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "         WHERE ROWNUM < 37) "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCIE_CLACLI = 3 "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_PrvTim, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al calcular el tiempo de provision DUDOSA")
            End If
            
            If Not (r_rst_PrvTim.BOF And r_rst_PrvTim.EOF) Then
               r_rst_PrvTim.MoveFirst
               If r_rst_PrvTim!CONTADOR = 36 Then
                  r_int_TipGar = 5
               End If
            End If
            
            r_rst_PrvTim.Close
            Set r_rst_PrvTim = Nothing
         End If
         
         '********** DETERMINA SI TIENE CLASICACION PERDIDA POR MAS DE 24 MESES **********
         If modprc_g_rst_Princi!HIPCIE_CLACLI = 4 Then
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT COUNT(*) AS CONTADOR "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM (SELECT HIPCIE_CLACLI "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                  FROM CRE_HIPCIE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                 WHERE HIPCIE_PERMES > 0 "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                   AND HIPCIE_PERANO > 2015 "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                   AND HIPCIE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "' "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "         WHERE ROWNUM < 25) "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCIE_CLACLI = 4 "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_PrvTim, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al calcular el tiempo de provision DUDOSA")
            End If
            
            If Not (r_rst_PrvTim.BOF And r_rst_PrvTim.EOF) Then
               r_rst_PrvTim.MoveFirst
               If r_rst_PrvTim!CONTADOR = 24 Then
                  r_int_TipGar = 5
               End If
            End If
            
            r_rst_PrvTim.Close
            Set r_rst_PrvTim = Nothing
         End If
         
         'Si Tipo de Garantia es Bloqueo validar que tenga menos de 90 dias de bloqueo
         If r_int_TipGar = 2 Then
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT EVALEG_FECBLQ_INM AS FECHA_BLQ "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM TRA_EVALEG "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE EVALEG_NUMSOL = (SELECT HIPMAE_NUMSOL FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "') "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_PrvTim, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al calcular el tiempo de provision DUDOSA")
            End If
            
            If DateDiff("d", gf_FormatoFecha(r_rst_PrvTim!FECHA_BLQ), p_FecFin) > 90 Then
               r_int_TipGar = 3
            End If
         
            r_rst_PrvTim.Close
            Set r_rst_PrvTim = Nothing
         End If
         
         'Determinando Clase de Garantía según Tipo de Garantía
         r_int_ClaGar = 0
         For r_int_Contad = 1 To UBound(r_arr_DetGar)
            If r_arr_DetGar(r_int_Contad).DetGar_Codigo = r_int_TipGar Then
               r_int_ClaGar = r_arr_DetGar(r_int_Contad).DetGar_ClaGar
               Exit For
            End If
         Next r_int_Contad
         
         '**************************************************
         'INICIO - NUEVA FORMA DE CALCULO DE LAS PROVISIONES
         r_dbl_GtoJud = 0
         r_dbl_MtoCga = 0
         r_dbl_MtoSga = 0
         
         'Determina gastos judiciales
         'r_dbl_GtoJud = modprc_ff_CalculaGastosJudicial(modprc_g_rst_Princi!HIPCIE_CODPRD, modprc_g_rst_Princi!HIPCIE_CODSUB, (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON), modprc_g_rst_Princi!HIPCIE_TIPMON, modprc_g_rst_Princi!HIPCIE_TIPCAM, r_int_TipGar, modprc_g_rst_Princi!HIPCIE_MTOGAR, modprc_g_rst_Princi!HIPCIE_MONGAR)
         r_dbl_GtoJud = modprc_ff_CalculaGastosJudicial(modprc_g_rst_Princi!HIPCIE_CODPRD, modprc_g_rst_Princi!HIPCIE_CODSUB, (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON), modprc_g_rst_Princi!HIPCIE_TIPMON, modprc_g_rst_Princi!HIPCIE_TIPCAM, r_int_TipGar, r_dbl_ValRea, modprc_g_rst_Princi!HIPCIE_MONGAR)
         
         'Determina montos base para el calculo de la provision
         Call modprc_ff_CalculaMontosBaseProv2(r_dbl_MtoCga, r_dbl_MtoSga, r_dbl_GtoJud, modprc_g_rst_Princi!HIPCIE_SALCAP, modprc_g_rst_Princi!HIPCIE_SALCON, modprc_g_rst_Princi!HIPCIE_CODPRD, modprc_g_rst_Princi!HIPCIE_TIPMON, modprc_g_rst_Princi!HIPCIE_TIPCAM, modprc_g_rst_Princi!HIPCIE_FECDES, r_int_TipGar, modprc_g_rst_Princi!HIPCIE_MTOGAR, modprc_g_rst_Princi!HIPCIE_MONGAR, modprc_g_rst_Princi!HIPCIE_INTDIF, 1)
         
         r_dbl_MtoCga = Format(r_dbl_MtoCga, "###,##0.00")
         r_dbl_MtoSga = Format(r_dbl_MtoSga, "###,##0.00")
         r_dbl_PrvGen = 0
         r_dbl_PrvGen_RC = 0
         r_dbl_PrvCic = 0
         r_dbl_PrvCic_RC = 0
         r_dbl_PrvCam = 0
         r_dbl_PrvEsp = 0
         
         'Si Clasificacion del deudor es normal
         If modprc_g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            '*******************  PROVISION GENERICA  ***************************
            r_dbl_Porce1 = (modprc_gf_PorcenProv(r_arr_TipPrv, 1, modprc_g_rst_Princi!HIPCIE_CLAPRV, r_int_ClaGar) / 100)
            r_dbl_PrvGen = r_dbl_Porce1 * (r_dbl_MtoCga + r_dbl_MtoSga)
            
            '************  PROVISION GENERICA - RIESGO CONTRAPARTE **************
            If modprc_g_rst_Princi!HIPCIE_FECDES > 20100630 Then
               If r_int_TipGar = 1 Or r_int_TipGar = 2 Or r_int_TipGar = 9 Then
                  If modprc_g_rst_Princi!HIPCIE_CODPRD = "007" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "009" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "010" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "013" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "014" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "015" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "016" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "017" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "018" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "019" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "021" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "022" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "023" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "024" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "025" Then
                     r_dbl_PrvGen_RC = ((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) / 3) * (0.7 / 100)
                  End If
               End If
            End If
            
            '*******************  PROVISION RIESGO CAMBIARIO  *******************
            If modprc_g_rst_Princi!HIPCIE_TIPMON = 2 Then
               r_dbl_Porce1 = modprc_gf_PorcenProv(r_arr_TipPrv, 4, modprc_g_rst_Princi!HIPCIE_CLAPRV, r_int_ClaGar) / 100
               r_dbl_PrvCam = r_dbl_Porce1 * (r_dbl_MtoCga + r_dbl_MtoSga)
            End If
         End If
         
         'Si Clasificacion del deudor es diferente de normal
         If modprc_g_rst_Princi!HIPCIE_CLAPRV <> 0 Then
            '*********************  PROVISION ESPECIFICA  ***********************
            r_dbl_Porce1 = (modprc_gf_PorcenProv(r_arr_TipPrv, 2, modprc_g_rst_Princi!HIPCIE_CLAPRV, 1) / 100)
            r_dbl_Porce2 = (modprc_gf_PorcenProv(r_arr_TipPrv, 2, modprc_g_rst_Princi!HIPCIE_CLAPRV, 2) / 100)
            r_dbl_PrvEsp = (r_dbl_Porce1 * r_dbl_MtoSga) + (r_dbl_Porce2 * r_dbl_MtoCga)
         
            '***  RIESGO CONTRAPARTE - PRODUCTOS FMV A PARTIR DEL 01/07/2010  ***
            If (modprc_g_rst_Princi!HIPCIE_FECDES > 20100630) Then
               If (r_int_TipGar = 1 Or r_int_TipGar = 2 Or r_int_TipGar = 9) Then
                  If modprc_g_rst_Princi!HIPCIE_CODPRD = "007" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "009" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "010" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "013" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "014" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "015" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "016" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "017" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "018" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "019" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "021" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "022" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "023" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "024" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "025" Then
                     r_dbl_PrvGen_RC = ((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) / 3) * (0.7 / 100)
                  End If
               End If
            End If
         
            '***  LAS CARTAS FIANZAS SE COMPORTAN COMO GARANTIA PREFERIDA (HIPOTECAS) - RESOLUCION SBS
            If r_int_TipGar = 4 Then
               r_dbl_PrvEsp = (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) * r_dbl_Porce2
            End If
         End If
         
         r_dbl_PrvEsp = CDbl(Format(r_dbl_PrvEsp, "######0.00"))
         r_dbl_PrvCam = CDbl(Format(r_dbl_PrvCam, "######0.00"))
         r_dbl_PrvGen = CDbl(Format(r_dbl_PrvGen, "######0.00"))
         r_dbl_PrvGen_RC = CDbl(Format(r_dbl_PrvGen_RC, "######0.00"))
         r_dbl_PrvCic = CDbl(Format(r_dbl_PrvCic, "######0.00"))
         r_dbl_PrvCic_RC = CDbl(Format(r_dbl_PrvCic_RC, "######0.00"))
         'FINAL - NUEVA FORMA DE CALCULO DE LAS PROVISIONES
         
         '****************************************
         'INICIO - CALCULO DE LA COBERTURA DEL FMV
         r_dbl_TerSal = 0
         r_dbl_TerPer = 0
         r_dbl_CobFmv = 0
         r_dbl_CobFmv_RC = 0
         
         'Antes del 01/07/2010 - Cobertura
         If modprc_g_rst_Princi!HIPCIE_FECDES < 20100701 Then
         
            'If modprc_g_rst_Princi!HIPCIE_TIPGAR = 1 Or modprc_g_rst_Princi!HIPCIE_TIPGAR = 2 Or modprc_g_rst_Princi!HIPCIE_TIPGAR = 4 Or modprc_g_rst_Princi!HIPCIE_TIPGAR = 9 Then
            If r_int_TipGar = 1 Or r_int_TipGar = 2 Or r_int_TipGar = 4 Or r_int_TipGar = 9 Then
               If modprc_g_rst_Princi!HIPCIE_CODPRD = "001" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "003" Then
                  If ((modprc_g_rst_Princi!HIPCIE_MTOGAR * 2 / 3) - r_dbl_GtoJud) >= (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) Then
                     r_dbl_CobFmv = 0
                  Else
                     r_dbl_TerSal = (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) / 3
                     r_dbl_TerPer = ((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) - ((modprc_g_rst_Princi!HIPCIE_MTOGAR * 2 / 3) - r_dbl_GtoJud)) / 3
                     
                     If r_dbl_TerSal > r_dbl_TerPer Then
                        r_dbl_CobFmv = r_dbl_TerPer
                     Else
                        r_dbl_CobFmv = r_dbl_TerSal
                     End If
                  End If
               End If
               Select Case modprc_g_rst_Princi!HIPCIE_CODPRD
                      Case "004", "007", "009", "010", "012", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023", "024", "025"
                           r_dbl_CobFmv = (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) / 3
               End Select
            End If
         End If
         r_dbl_CobFmv = CDbl(Format(r_dbl_CobFmv, "######0.00"))
         
         'Despues del 01/07/2010 - Riesgo Contraparte FMV
         If modprc_g_rst_Princi!HIPCIE_FECDES >= 20100701 Then
             'If modprc_g_rst_Princi!HIPCIE_TIPGAR = 1 Or modprc_g_rst_Princi!HIPCIE_TIPGAR = 2 Or modprc_g_rst_Princi!HIPCIE_TIPGAR = 9 Then
             If r_int_TipGar = 1 Or r_int_TipGar = 2 Or r_int_TipGar = 9 Then
                Select Case modprc_g_rst_Princi!HIPCIE_CODPRD
                       Case "004", "007", "009", "010", "012", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023", "024", "025"
                            r_dbl_CobFmv_RC = Format((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON - modprc_g_rst_Princi!HIPCIE_INTDIF) / 3, "###,###,##0.00")
                End Select
            End If
         End If
         'FINAL - CALCULO DE LA COBERTURA DEL FMV
         
         '**********************************************************************
         'INICIO - PROVISION VOLUNTARIAS (05/02/2015) y PROCICLICAS (30/06/2015)
         r_dbl_PrvVol = 0
         If p_PerMes = 1 Then
            r_int_MesAnt = 12
            r_int_AnoAnt = p_PerAno - 1
         Else
            r_int_MesAnt = p_PerMes - 1
            r_int_AnoAnt = p_PerAno
         End If
         
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT HIPCIE_PRVVOL, HIPCIE_PRVCIC, HIPCIE_PRVCIC_RC "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CRE_HIPCIE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCIE_PERMES = " & r_int_MesAnt & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCIE_PERANO = " & r_int_AnoAnt & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCIE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_PrvTim, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al obtener la provision voluntaria")
         End If
         
         If Not (r_rst_PrvTim.BOF And r_rst_PrvTim.EOF) Then
            r_rst_PrvTim.MoveFirst
            r_dbl_PrvVol = r_rst_PrvTim!HIPCIE_PRVVOL
            r_dbl_PrvCic = r_rst_PrvTim!HIPCIE_PRVCIC
            r_dbl_PrvCic_RC = r_rst_PrvTim!HIPCIE_PRVCIC_RC
         End If
         
         r_rst_PrvTim.Close
         Set r_rst_PrvTim = Nothing
         'FINAL - PROVISION VOLUNTARIAS 05/02/2015
         
         'Los creditos con GARANTIA HIPOTECARIA no generan RC
         If r_int_OrgGar = 9 Then
            r_dbl_PrvGen_RC = 0
            r_dbl_PrvCic_RC = 0
            r_dbl_CobFmv_RC = 0
            r_dbl_CobFmv = 0
         End If
         
         'Actualizando Provisión en CRE_HIPCIE
         modprc_g_str_CadEje = "UPDATE CRE_HIPCIE " & _
                               "   SET HIPCIE_PRVESP    = " & CStr(r_dbl_PrvEsp) & ", " & _
                               "       HIPCIE_PRVCAM    = " & CStr(r_dbl_PrvCam) & ", " & _
                               "       HIPCIE_PRVVOL    = " & CStr(r_dbl_PrvVol) & ", " & _
                               "       HIPCIE_PRVGEN    = " & CStr(r_dbl_PrvGen) & ", " & _
                               "       HIPCIE_PRVGEN_RC = " & CStr(r_dbl_PrvGen_RC) & ", " & _
                               "       HIPCIE_PRVCIC    = " & CStr(r_dbl_PrvCic) & ", " & _
                               "       HIPCIE_PRVCIC_RC = " & CStr(r_dbl_PrvCic_RC) & ", " & _
                               "       HIPCIE_CBRFMV    = " & CStr(r_dbl_CobFmv) & ", " & _
                               "       HIPCIE_CBRFMV_RC = " & CStr(r_dbl_CobFmv_RC) & " " & _
                               " WHERE HIPCIE_NUMOPE    = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "'" & _
                               "   AND HIPCIE_PERMES    = " & CStr(p_PerMes) & " " & _
                               "   AND HIPCIE_PERANO    = " & CStr(p_PerAno) & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPCIE - Operación: " & modprc_g_rst_Princi!HIPCIE_NUMOPE)
         End If
         
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
      
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1010(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1010
'Descripción      :  Actualización en CRE_HIPMAE de Saldos de Devengados desde la CRE_HIPCIE
'Resumen          :  Actualización en CRE_HIPMAE de Devengados (Vigentes y Vencidos) desde la CRE_HIPCIE
'F. Creación      :  03-08-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_FecPro        As String
Dim r_rst_Grabar        As ADODB.Recordset
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1010"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = "SELECT COUNT(HIPCIE_NUMOPE) AS TOTREG FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   'Leyendo Cursor Principal (Créditos Hipotecarios)
   modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         'Acumulando Temporalmente en CRE_HIPMAE
         modprc_g_str_CadEje = "UPDATE CRE_HIPMAE SET "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_ACUDVG = " & CStr(modprc_g_rst_Princi!HIPCIE_ACUDVG) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_ACUDVC = " & CStr(modprc_g_rst_Princi!HIPCIE_ACUDVC) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_ULTDEV = " & CStr(modprc_g_rst_Princi!HIPCIE_FECDEV) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "WHERE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPCIE_NUMOPE) & " .")
         End If
            
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Else
      
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1011(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, ByVal p_FecIni As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1011
'Descripción      :  Cierre de Operaciones Mensuales - Créditos Hipotecarios
'Resumen          :  Proceso que traslada los Saldos de Fin de Mes de la tabla CRE_HIPMAE hacia la CRE_HIPCIE
'F. Creación      :  21-04-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :  15-01-2012
'U. Actualización :  Rafael Durand
'Resumen Actual.  :  Adicion de proceso que traslada las cuotas de Fin de Mes de la tabla CRE_HIPCUO hacia la CRE_CUOCIE

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_FecPro        As String
Dim r_str_IntDif        As Double
Dim r_dbl_TipCam_Dol    As Double
Dim r_dbl_TipCam        As Double
Dim r_dbl_ComVta        As Double
Dim r_dbl_ApoPro        As Double
Dim r_int_ActEco        As Integer
Dim r_int_CodCiu        As Integer
Dim r_int_SecEco        As Integer
Dim r_int_TipEva        As Integer
Dim r_int_DiaAtr        As Integer
Dim r_str_PrxVct        As String
Dim r_int_DiaVc1        As Integer
Dim r_int_DiaVc2        As Integer
Dim r_arr_CreHip()      As modprc_g_tpo_CreHip
Dim r_int_TipCre        As Integer
Dim r_dbl_PerPbp        As Double
Dim r_dbl_AcuDvg        As Double
Dim r_dbl_AcuDvc        As Double
Dim r_arr_TipCla()      As modprc_g_tpo_Genera
Dim r_int_ClaCre        As Integer
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_int_Exporc        As Integer
Dim r_dbl_CapPag_TNC    As Double
Dim r_dbl_CapPag_TCO    As Double
Dim r_dbl_SalCap        As Double
Dim r_dbl_SalCon        As Double
Dim r_int_FlgRef        As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1011"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   r_lng_NumReg = 0
   r_lng_TotReg = 0
   r_dbl_PerPbp = 0
   r_int_FlgRef = 0
   p_BarPro.FloodPercent = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Período Vigente
   'r_str_Period = moddat_gf_ConsultaPerMesActivo(p_CodEmp, 1, r_str_FecIni, r_str_FecFin, r_int_PerMes, r_int_PerAno)
   
   'Leyendo Cursor Principal
   modprc_g_str_CadEje = "SELECT COUNT(*) AS TOTREG " & _
                         "  FROM CRE_HIPMAE " & _
                         " WHERE HIPMAE_PROCRE = '" & p_CodEmp & "' " & _
                         "   AND HIPMAE_SITUAC = 2"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   'Cargando Dias para pase a Vencido (Parcial y Total)
   Call modprc_gs_CargaSituacion("13", r_int_DiaVc1, r_int_DiaVc2, r_arr_LogPro())
   
   'Cargando Tipos de Clasificación
   Call modprc_gs_CargaTiposClasif("13", r_arr_TipCla(), r_arr_LogPro())
   
   'Leyendo Cursor Principal
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CRE_HIPMAE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_PROCRE = '" & p_CodEmp & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPMAE_SITUAC = 2 "
   'modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPMAE_NUMOPE = '0040700003' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         ReDim r_arr_CreHip(0)
         ReDim r_arr_CreHip(1)
         
         r_arr_CreHip(1).CreHip_NumOpe = modprc_g_rst_Princi!HIPMAE_NUMOPE
         r_arr_CreHip(1).CreHip_SalCap = modprc_g_rst_Princi!HIPMAE_SALCAP
         r_arr_CreHip(1).CreHip_SalCon = modprc_g_rst_Princi!HIPMAE_SALCON
         r_arr_CreHip(1).CreHip_IniAde = "0"
         r_arr_CreHip(1).CreHip_FecDes = CStr(modprc_g_rst_Princi!HIPMAE_FECDES)
         r_arr_CreHip(1).CreHip_DevAnt = CStr(modprc_g_rst_Princi!HIPMAE_ULTDEV)
         r_arr_CreHip(1).CreHip_MtoPre = modprc_g_rst_Princi!HIPMAE_MTOPRE
         r_arr_CreHip(1).CreHip_MtoNCo = modprc_g_rst_Princi!HIPMAE_IMPNCO
         r_arr_CreHip(1).CreHip_FecPPg = CStr(modprc_g_rst_Princi!HIPMAE_FECPPG)
         r_int_FlgRef = modprc_g_rst_Princi!HIPMAE_REFINA
         
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Or modprc_g_rst_Princi!HIPMAE_MONEDA = 2 Then
            r_dbl_TipCam = r_dbl_TipCam_Dol
         Else
            r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, modprc_g_rst_Princi!HIPMAE_MONEDA, Format(CDate(p_FecFin), "yyyymmdd"), 2)
         End If
         
         'Obteniendo Valor de Compra Venta en Moneda de Préstamo
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_dbl_ComVta = modprc_g_rst_Princi!HIPMAE_CVTSOL
            r_dbl_ApoPro = modprc_g_rst_Princi!HIPMAE_APOSOL
         Else
            r_dbl_ComVta = modprc_g_rst_Princi!HIPMAE_CVTDOL
            r_dbl_ApoPro = modprc_g_rst_Princi!HIPMAE_APODOL
         End If
         
         'Buscando Actividad Económica Principal y CIIU del Cliente
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CLI_DATGEN "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE DATGEN_TIPDOC = " & CStr(modprc_g_rst_Princi!HIPMAE_TDOCLI) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND DATGEN_NUMDOC = '" & Trim(modprc_g_rst_Princi!HIPMAE_NDOCLI) & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CLI_DATGEN - Operación Nro.: " & modprc_g_rst_Princi!HIPMAE_NUMOPE & " .")
         End If
         
         r_int_ActEco = 0
         r_int_CodCiu = 0
         r_int_Exporc = 0
         
         If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
            modprc_g_rst_Genera.MoveFirst
            r_int_ActEco = modprc_g_rst_Genera!DATGEN_OCUPAC
            r_int_CodCiu = modprc_g_rst_Genera!DATGEN_CODCIU
            r_int_Exporc = modprc_g_rst_Genera!DATGEN_EXPORC
         End If
         
         modprc_g_rst_Genera.Close
         Set modprc_g_rst_Genera = Nothing
         
         'Buscando Sector Económico del CIIU
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM MNT_SECCIU "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE SECCIU_CODCIU = " & CStr(r_int_CodCiu) & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla MNT_SECCIU - Operación Nro.: " & modprc_g_rst_Princi!HIPMAE_NUMOPE & " .")
         End If
         
         r_int_SecEco = 0
         
         If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
            modprc_g_rst_Genera.MoveFirst
            r_int_SecEco = modprc_g_rst_Genera!SECCIU_CODSEC
         End If
         
         modprc_g_rst_Genera.Close
         Set modprc_g_rst_Genera = Nothing
         
         'Buscando Tipo de Evaluación Crediticia
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT SOLMAE_TIPEVA "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CRE_SOLMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE SOLMAE_NUMERO = '" & modprc_g_rst_Princi!HIPMAE_NUMSOL & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al leer Tabla CRE_SOLMAE - Operación Nro.: " & modprc_g_rst_Princi!HIPMAE_NUMOPE & " .")
         End If
         
         modprc_g_rst_Genera.MoveFirst
         r_int_TipEva = modprc_g_rst_Genera!SOLMAE_TIPEVA
         
         modprc_g_rst_Genera.Close
         Set modprc_g_rst_Genera = Nothing
         
         'Determinando Días de Atraso
         r_str_PrxVct = gf_FormatoFecha(modprc_g_rst_Princi!HIPMAE_PRXVCT)
         r_int_DiaAtr = 0
         If CDate(r_str_PrxVct) < CDate(p_FecFin) Then
            r_int_DiaAtr = CInt(CDate(p_FecFin) - CDate(r_str_PrxVct))
         End If
         
         'Clasificando Crédito según Días de Atraso
         r_int_ClaCre = modprc_gf_ClaCred(r_arr_TipCla, r_int_DiaAtr)
         
         'Buscando Cuotas Atrasadas y Cargos por Cobranzas Atrasadas y Situación Contable de Crédito
         Call modprc_gs_EvaluacionSituacion_CreHip(r_arr_CreHip, r_arr_LogPro, p_FecFin, r_int_DiaAtr, r_int_DiaVc1, r_int_DiaVc2)
         
         'Actualiza situacion contable del credito
         If modprc_g_rst_Princi!HIPMAE_REFINA = 1 Then
            r_arr_CreHip(1).CreHip_SitCtb = 6
         End If
         If modprc_g_rst_Princi!HIPMAE_JUDICI = 1 Then
            r_arr_CreHip(1).CreHip_SitCtb = 5
         End If
         
         'Determinando Tipo de Crédito
         If r_arr_CreHip(1).CreHip_SitCtb = 4 Then
            r_int_TipCre = 5
         Else
            If modprc_g_rst_Princi!HIPMAE_REFINA = 0 And modprc_g_rst_Princi!HIPMAE_JUDICI = 0 And modprc_g_rst_Princi!HIPMAE_CASTIG = 0 Then
               r_int_TipCre = 1
            ElseIf modprc_g_rst_Princi!HIPMAE_REFINA = 1 Then
               r_int_TipCre = 4
            ElseIf modprc_g_rst_Princi!HIPMAE_JUDICI = 1 Then
               r_int_TipCre = 6
            End If
         End If
         
         'Calculando Devengado de PBP (Productos con Devolución de Dinero en favor de miCasita)
         If modprc_g_rst_Princi!HIPMAE_CODPRD = "001" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "003" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "006" Then
            Call modprc_gs_DevengadoPBP_CreHip(r_arr_CreHip, r_arr_LogPro, p_FecFin)
         End If
         
         'Calculando Provisión de Adeudado (Productos con Fondo Mivivienda-COFIDE)
         If modprc_g_rst_Princi!HIPMAE_CODPRD = "003" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "004" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "007" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "009" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "010" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "013" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "014" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "015" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "016" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "017" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "018" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "019" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "021" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "022" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "023" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "024" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "025" Then
            Call modprc_gs_ProvisionAdeudado_CreHip(r_arr_CreHip, r_arr_LogPro, p_FecFin, modprc_g_rst_Princi!HIPMAE_TASCOF, modprc_g_rst_Princi!HIPMAE_COMCOF)
         End If
         
         'Calculando Interés Devengado del Mes
         Call modprc_gs_Devengado_CreHip(r_arr_CreHip, r_arr_LogPro, p_FecIni, p_FecFin, modprc_g_rst_Princi!HIPMAE_TASINT)
         Call modprc_gs_Devengado_CreHip2(r_arr_CreHip, r_arr_LogPro, p_FecIni, p_FecFin, modprc_g_rst_Princi!HIPMAE_TASINT, r_int_DiaAtr, r_dbl_AcuDvg)
         
'         'Moviendo Intereses y Devengados según Situación
'         If modprc_g_rst_Princi!HIPMAE_ACUDVG - r_arr_CreHip(1).CreHip_IntVen < 0 Then
'            r_dbl_AcuDvg = 0
'         Else
'            'Acumulado Interés Devengado Vigente
'            r_dbl_AcuDvg = modprc_g_rst_Princi!HIPMAE_ACUDVG - r_arr_CreHip(1).CreHip_IntVen
'         End If
'
'         r_dbl_AcuDvg = r_dbl_AcuDvg + r_arr_CreHip(1).CreHip_IntDev                            'Acumulando Interés Devengado Vigente
         r_dbl_AcuDvc = r_arr_CreHip(1).CreHip_IntVen                                           'Acumulando Interés Vencido
         
         'Obteniendo por Diferencia Interés Vencido del Mes
         If r_dbl_AcuDvc - modprc_g_rst_Princi!HIPMAE_ACUDVC < 0 Then
            r_arr_CreHip(1).CreHip_IntVen = 0
         Else
            r_arr_CreHip(1).CreHip_IntVen = r_dbl_AcuDvc - modprc_g_rst_Princi!HIPMAE_ACUDVC
         End If
         
'         If r_int_DiaAtr = 0 Then
'            r_dbl_AcuDvg = r_arr_CreHip(1).CreHip_IntDev
'         ElseIf r_int_DiaAtr > 30 Then
         If r_int_DiaAtr > 30 Then
            r_dbl_AcuDvg = 0
            r_dbl_AcuDvc = modprc_gf_Calcula_InteresVencido(modprc_g_rst_Princi!HIPMAE_NUMOPE, p_FecFin, modprc_g_rst_Princi!HIPMAE_TASINT)
         End If
         
         'Obteniendo por Diferencia Capital Vencido del Mes
         If r_arr_CreHip(1).CreHip_CapVen - modprc_g_rst_Princi!HIPMAE_CAPVEN < 0 Then
            r_arr_CreHip(1).CreHip_CVcMes = 0
         Else
            r_arr_CreHip(1).CreHip_CVcMes = r_arr_CreHip(1).CreHip_CapVen - modprc_g_rst_Princi!HIPMAE_CAPVEN
         End If
         
         'Determina el PBP Perdido
         r_dbl_PerPbp = modprc_gf_Calcula_PBPPerdido(modprc_g_rst_Princi!HIPMAE_NUMOPE)
         
         'Determinado interes diferido del mes
         r_str_IntDif = modprc_gf_Calcula_InteresDiferido(modprc_g_rst_Princi!HIPMAE_NUMOPE, p_FecFin)
         
         'RDB - INI - 20121108
         'Para determinar Capital Pagado por Tramos TNC y TC
         r_dbl_CapPag_TNC = 0
         If modprc_g_rst_Princi!HIPMAE_CODPRD = "002" Or modprc_g_rst_Princi!HIPMAE_CODPRD = "011" Then
            r_dbl_CapPag_TNC = modprc_g_rst_Princi!HIPMAE_TOTPRE - modprc_g_rst_Princi!HIPMAE_SALCAP
         Else
            r_dbl_CapPag_TNC = modprc_g_rst_Princi!HIPMAE_IMPNCO - modprc_g_rst_Princi!HIPMAE_SALCAP
         End If
         
         r_dbl_CapPag_TCO = 0
         If modprc_g_rst_Princi!HIPMAE_CODPRD <> "002" Or modprc_g_rst_Princi!HIPMAE_CODPRD <> "011" Then
            r_dbl_CapPag_TCO = modprc_g_rst_Princi!HIPMAE_IMPCON - (modprc_g_rst_Princi!HIPMAE_SALCON + r_dbl_PerPbp)
         End If
         'RDB - FIN - 20121108
         
         'RDB - INI - 20120224
         If modprc_g_rst_Princi!HIPMAE_REFINA = 1 Then
            r_dbl_AcuDvg = 0
            r_dbl_AcuDvc = r_arr_CreHip(1).CreHip_IntDev
         End If
         
         If r_int_DiaAtr > 90 Then
            r_arr_CreHip(1).CreHip_CapVen = r_arr_CreHip(1).CreHip_CapVen + r_dbl_PerPbp
         Else
            r_arr_CreHip(1).CreHip_CapVig = r_arr_CreHip(1).CreHip_CapVig + r_dbl_PerPbp
         End If
         'RDB - FIN - 20120224
         
         'Buscando Ultimo Pago efectuado sobre Capital
         Call modprc_gs_UltimoCapPago_CreHip(r_arr_CreHip, r_arr_LogPro)
         
         'Reclasificacion de Riesgo Cambiario crediticio
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_int_Exporc = 1
         End If
         
         'Insertando en CRE_HIPCIE
         g_str_Parame = "USP_CRE_HIPCIE ("
         g_str_Parame = g_str_Parame & CStr(p_PerMes) & ","
         g_str_Parame = g_str_Parame & CStr(p_PerAno) & ","
         g_str_Parame = g_str_Parame & "'" & modprc_g_rst_Princi!HIPMAE_NUMOPE & "', "                   'Numero de Operación
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_PROCRE & "") & "', "        'Propietario de Cartera
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_CODTIT & "") & "', "        'Código de Titulización
         g_str_Parame = g_str_Parame & Format(CDate(r_str_FecPro), "yyyymmdd") & ", "                    'Fecha de Proceso
         g_str_Parame = g_str_Parame & CStr(r_dbl_TipCam) & ", "                                         'Tipo de Cambio
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_MONEDA) & ", "                    'Moneda Original de Crèdito
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_MTOPRE) & ", "                    'Monto de Préstamo
         g_str_Parame = g_str_Parame & CStr(r_dbl_ComVta) & ", "                                         'Valor de Compra Venta
         g_str_Parame = g_str_Parame & CStr(r_dbl_ApoPro) & ", "                                         'Aporte Propio
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_INTCAP) & ", "                    'Interés Capitalizado
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_TOTPRE) & ", "                    'Total de Préstamo
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_PLAANO) & ", "                    'Plazo en Años
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_PERGRA) & ", "                    'Período de Gracia
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_NUMCUO) & ", "                    'Número de Cuotas
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_CLACRE) & ", "                    'Clase de Producto
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) & "', "             'Código de Producto
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_CODSUB) & "', "             'Código de Sub-Producto
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_CODMOD) & "', "             'Código de Modalidad
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_UBIGEO) & "', "             'Ubicación Geográfica
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_PRYINM & "") & "', "        'Código de Proyecto Inmobiliario
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_PRYMCS) & ", "                    'Flag de Proyecto miCasita
         g_str_Parame = g_str_Parame & CStr(r_int_ActEco) & ", "                                         'Código de Actividad Económica
         g_str_Parame = g_str_Parame & CStr(r_int_CodCiu) & ", "                                         'Código de CIIU de Actividad Económica Principal
         g_str_Parame = g_str_Parame & CStr(r_int_SecEco) & ", "                                         'Sector Económico
         g_str_Parame = g_str_Parame & CStr(r_int_TipEva) & ", "                                         'Tipo de Evaluación Crediticia
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_TASINT) & ", "                    'Tasa de Interés
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_APLPRE) & ", "                    'Tipo de Aplicación Seguro Desgravamen
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_FOIPRE) & ", "                    'Tasa Seguro Desgravamen
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_APLVIV) & ", "                    'Tipo de Aplicación Seguro Inmueble
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_FOIVIV) & ", "                    'Tasa Seguro Inmueble
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_OTRIMP) & ", "                    'Portes
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_TASMVI) & ", "                    'Tasa Interés Mivivienda
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_TASCOF) & ", "                    'Tasa Interés Cofide
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_COMCOF) & ", "                    'Comisión Cofide
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_COMCRC) & ", "                    'Comisión CRC
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_COMPBP) & ", "                    'Comisión PBP
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_COSEFE) & ", "                    'Costo Efectivo
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_TIPGAR) & ", "                    'Tipo de Garantía
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_MONGAR) & ", "                    'Moneda de Garantía
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_MTOGAR) & ", "                    'Monto de Garantía
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_FECDES) & ", "                    'Fecha de Desembolso
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_SITUAC) & ", "                    'Situación Crédito
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_SALCAP) & ", "                    'Saldo Capital
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_SALCON + r_dbl_PerPbp) & ", "     'Saldo Concesional
         g_str_Parame = g_str_Parame & CStr(r_int_DiaAtr) & ", "                                         'Días de Atraso
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_CuoAtr) & ", "                        'Cuotas de Atraso
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_CUOPEN) & ", "                    'Cuotas Pendientes de Pago
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_CUOPAG) & ", "                    'Cuotas Pagadas
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_SitCtb) & ", "                        'Situación Crédito SBS
         g_str_Parame = g_str_Parame & CStr(r_int_TipCre) & ", "                                         'Tipo Crédito SBS
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_REFINA) & ", "                    'Flag de Refinanciación
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_JUDICI) & ", "                    'Flag de Judicial
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_CASTIG) & ", "                    'Flag de Castigado
         g_str_Parame = g_str_Parame & CStr(r_int_ClaCre) & ", "                                         'Clasificación Crédito
         g_str_Parame = g_str_Parame & "0, "                                                             'Clasificación Cliente
         g_str_Parame = g_str_Parame & "0, "                                                             'Clasificación Alineada
         g_str_Parame = g_str_Parame & "0, "                                                             'Provisión Generica
         g_str_Parame = g_str_Parame & "0, "                                                             'Provisión Específica
         g_str_Parame = g_str_Parame & "0, "                                                             'Provisión Riesgo Cambiario
         g_str_Parame = g_str_Parame & "0, "                                                             'Provisión Pro-Ciclica
         g_str_Parame = g_str_Parame & "0, "                                                             'Provisión Adicional
         g_str_Parame = g_str_Parame & Format(CDate(p_FecFin), "yyyymmdd") & ","                         'Fecha Devengo
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_CuoDev) & ", "                        'Cuota Devengada
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_IntDev) & ", "                        'Devengado Vigente (Mes)
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_IntVen) & ", "                        'Devengado Vencido (Mes)
         'g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_ACUDIF) & ", "                    'Interés Diferido
         g_str_Parame = g_str_Parame & CStr(r_str_IntDif) & ", "                                         'Interés Diferido
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_CapVen) & ", "                        'Capital Vencido (Total)
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_CapVig) & ", "                        'Capital Vigente (Total) (Incluye TNC, TC)
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_VCTANT) & ", "                    'Fecha de Vencimiento Anterior
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_PRXVCT) & ", "                    'Fecha de Próximo Vencimiento
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_ULTPAG) & ", "                    'Fecha de Ultimo Pago
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_ULTDEV) & ", "                    'Fecha de Ultimo Devengo
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_IntCom) & ", "                        'Interés Compensatorio Vencido
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_IntMor) & ", "                        'Interés Moratorio Vencido
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_GasCob) & ", "                        'Gastos de Cobranzas Vencido
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_OtrGas) & ", "                        'Otros Gastos Vencido
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_UltCap) & ", "                        'Ultimo Capital Pagado
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_IMPNCO) & ", "                    'Monto de Préstamo TNC
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_IMPCON) & ", "                    'Monto de Préstamo TC
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_IMoVig) & ", "                        'Interés Moratorio Vigente
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_ICoVig) & ", "                        'Interés Compensatorio Vigente
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_GCoVig) & ", "                        'Gastos de Cobranzas Vigente
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_OtGVig) & ", "                        'Otros Gastos Vigente
         g_str_Parame = g_str_Parame & CStr(r_dbl_AcuDvg) & ", "                                         'Acumulado Interés Devengado Vigente
         g_str_Parame = g_str_Parame & CStr(r_dbl_AcuDvc) & ", "                                         'Acumulado Interés Vencido
         'g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_ACUDIF) & ", "                    'Acumulado Interés Diferido
         g_str_Parame = g_str_Parame & CStr(r_str_IntDif) & ", "                                         'Acumulado Interés Diferido
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_TDOCLI) & ", "                    'Tipo Documento Identidad Cliente
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_NDOCLI) & "', "             'Tipo Documento Identidad Cliente
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_DevPBP) & ", "                        'Devengado Premio Buen Pagador
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_PrvICo) & ", "                        'Provisión de Interés COFIDE
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_PrvCCo) & ", "                        'Provisión de Comisión COFIDE
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_NCuoTC) & ", "                        'Nro. de Cuota Tramo Concesional
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_CapiTC) & ", "                        'Capital Cuota TC
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_InteTC) & ", "                        'Interés Cuota TC
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_ComiTC) & ", "                        'Comisión Cuota TC
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_MtoAde) & ", "                        'Monto Adeudado
         g_str_Parame = g_str_Parame & r_arr_CreHip(1).CreHip_IniAde & ", "                              'Fecha Inicio para Cálculo Provisiones de Adeudado
         g_str_Parame = g_str_Parame & "0, "                                                             'Clasificación para Provisión
         g_str_Parame = g_str_Parame & CStr(r_int_Exporc) & ", "                                         'Flag de Exposición al Riesgo Cambiario
         g_str_Parame = g_str_Parame & CStr(r_arr_CreHip(1).CreHip_CVcMes) & ", "                        'Capital Vencido del Mes
         g_str_Parame = g_str_Parame & CStr(r_dbl_CapPag_TNC) & ", "                                     'Capital Amortizado TNC
         g_str_Parame = g_str_Parame & CStr(r_dbl_CapPag_TCO) & ", "                                     'Capital Amortizado TC
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_FECACT) & ", "                    'Fecha Activación Crédito
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_UlTVCT) & ", "                    'Fecha Ultimo Vencimiento
         g_str_Parame = g_str_Parame & CStr(modprc_g_rst_Princi!HIPMAE_NUECRE) & ", "                    'Nueva Clasificacion de Credito
         g_str_Parame = g_str_Parame & CStr(r_dbl_PerPbp) & ", "                                         'PBP Perdido en el periodo
         g_str_Parame = g_str_Parame & "0, "                                                             'Cobertura de Riesgo del FMV

         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al grabar en tabla CRE_HIPCIE - Operación Nro.: " & modprc_g_rst_Princi!HIPMAE_NUMOPE & " .")
         End If

         'Acumulando Temporalmente en CRE_HIPMAE
         modprc_g_str_CadEje = "UPDATE CRE_HIPMAE SET "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_ACUDVG = " & CStr(r_dbl_AcuDvg) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_ACUDVC = " & CStr(r_dbl_AcuDvc) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_CAPVEN = " & CStr(r_arr_CreHip(1).CreHip_CapVen) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_ULTDEV = " & Format(CDate(p_FecFin), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "WHERE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPMAE_NUMOPE & "'"

         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPMAE_NUMOPE) & " .")
         End If

         'Graba cuotas del credito a la tabla cierre de cuotas
         g_str_Parame = "USP_CRE_CUOCIE ("
         g_str_Parame = g_str_Parame & CStr(p_PerMes) & ","
         g_str_Parame = g_str_Parame & CStr(p_PerAno) & ","
         g_str_Parame = g_str_Parame & "'" & Trim(modprc_g_rst_Princi!HIPMAE_NUMOPE) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_CIECUO - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPMAE_NUMOPE) & " .")
         End If
                  
         'Siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
      p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPMAE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1012(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1012
'Descripción      :  Contabilización de Devengados (Asientos en Edpymebank)
'Resumen          :  Contabilización de Devengados (Asientos en Edpymebank)
'F. Creación      :  04-08-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg              As Long
Dim r_lng_TotReg              As Long
Dim r_arr_LogPro()            As modprc_g_tpo_LogPro
Dim r_str_FecPro              As String
Dim r_rst_Grabar              As ADODB.Recordset
Dim r_dbl_TipCam_Dol          As Double
Dim r_dbl_Tot_ImpDeb_Sol      As Double
Dim r_dbl_Tot_ImpHab_Sol      As Double
Dim r_dbl_Tot_ImpDeb_Dol      As Double
Dim r_dbl_Tot_ImpHab_Dol      As Double
Dim r_str_CtaDeb              As String
Dim r_str_CtaHab              As String
Dim r_dbl_Lin_ImpDeb_Sol      As Double
Dim r_dbl_Lin_ImpHab_Sol      As Double
Dim r_dbl_Lin_ImpDeb_Dol      As Double
Dim r_dbl_Lin_ImpHab_Dol      As Double
Dim r_str_Origen              As String
Dim r_int_libro               As Integer
Dim r_str_TipNot              As String
Dim r_int_NumAsi              As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1012"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = "SELECT COUNT(HIPCIE_NUMOPE) AS TOTREG FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   'Leyendo Cursor Principal (Créditos Hipotecarios)
   modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         r_dbl_Tot_ImpDeb_Sol = 0
         r_dbl_Tot_ImpHab_Sol = 0
         r_dbl_Tot_ImpDeb_Dol = 0
         r_dbl_Tot_ImpHab_Dol = 0
         r_str_CtaDeb = ""
         r_str_CtaHab = ""
      
         Select Case modprc_g_rst_Princi!HIPCIE_CODPRD
            Case "001"  'Producto CRC-PBP
               r_str_CtaDeb = "142804010101"
               r_str_CtaHab = "512401042401"
            
            Case "002"  'Producto miCasita Dolares
               r_str_CtaDeb = "142804010101"
               r_str_CtaHab = "512401040601"
            
            Case "003"  'Producto CME
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042501"
               
            Case "004"  'Producto miHogar
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042301"
            
            Case "006"  'Producto miCasita PBP
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401040601"
            
            Case "007"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
            
            Case "009"  'Producto UNION ANDINA
               r_str_CtaDeb = "141804010101"
'20120509-1 - INICIO
              'r_str_CtaHab = "511401042303"
               r_str_CtaHab = "511401042305"
'20120509-1 - FIN

            Case "010"  'Producto miVivienda Peruanos en el Extranjero
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042304"
               
            Case "011"  'Producto miCasita en Soles
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401040601"
               
            Case "012"  'Producto UNION ANDINA miCasita
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401040601"
            
            Case "013"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "014"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "015"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
            
            Case "016"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
            
            Case "017"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
            
            Case "018"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "019"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
            
            Case "021"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "022"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "023"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "024"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
            Case "025"  'Producto NUEVO miVivienda
               r_str_CtaDeb = "141804010101"
               r_str_CtaHab = "511401042302"
               
         End Select
         
         Select Case modprc_g_rst_Princi!HIPCIE_TIPMON
            Case 1
               r_dbl_Lin_ImpDeb_Sol = modprc_g_rst_Princi!HIPCIE_DEVVIG
               r_dbl_Lin_ImpHab_Sol = modprc_g_rst_Princi!HIPCIE_DEVVIG
               r_dbl_Lin_ImpDeb_Dol = CDbl(Format(modprc_g_rst_Princi!HIPCIE_DEVVIG / r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpHab_Dol = CDbl(Format(modprc_g_rst_Princi!HIPCIE_DEVVIG / r_dbl_TipCam_Dol, "#####0.00"))
               
            Case 2
               r_dbl_Lin_ImpDeb_Sol = CDbl(Format(modprc_g_rst_Princi!HIPCIE_DEVVIG * r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpHab_Sol = CDbl(Format(modprc_g_rst_Princi!HIPCIE_DEVVIG * r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpDeb_Dol = modprc_g_rst_Princi!HIPCIE_DEVVIG
               r_dbl_Lin_ImpHab_Dol = modprc_g_rst_Princi!HIPCIE_DEVVIG
         End Select
         
         r_dbl_Tot_ImpDeb_Sol = r_dbl_Tot_ImpDeb_Sol + r_dbl_Lin_ImpDeb_Sol
         r_dbl_Tot_ImpHab_Sol = r_dbl_Tot_ImpHab_Sol + r_dbl_Lin_ImpHab_Sol
         r_dbl_Tot_ImpDeb_Dol = r_dbl_Tot_ImpDeb_Dol + r_dbl_Lin_ImpDeb_Dol
         r_dbl_Tot_ImpHab_Dol = r_dbl_Tot_ImpHab_Dol + r_dbl_Lin_ImpHab_Dol
         
         r_str_Origen = "LM"
         r_int_libro = 1
         r_str_TipNot = "O"
         
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_libro)
         
         'Insertar en CNTBL_ASIENTO
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, "001", r_dbl_TipCam_Dol, r_str_TipNot, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DEVENG. " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), p_FecFin, "1")
         
         'Insertar en CNTBL_ASIENTO_DET
         Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 1, r_str_CtaDeb, p_FecFin, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DEVENG. " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "D", r_dbl_Lin_ImpDeb_Sol, r_dbl_Lin_ImpDeb_Dol)
         Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 2, r_str_CtaHab, p_FecFin, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DEVENG. " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "H", r_dbl_Lin_ImpHab_Sol, r_dbl_Lin_ImpHab_Dol)
         
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         'p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
      
      'p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1013(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerIni As String, ByVal p_PerFin As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_CtbIni As String, ByVal p_CtbFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1013
'Descripción      :  Registro de Gastos de Cierre de Créditos Hipotecarios
'Resumen          :  Contabilización de Pago de Gastos de Cierre de Créditos Hipotecarios
'F. Creación      :  25-11-2009
'U. Creación      :  Jorge Luis Tacuche Mesia
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_arr_CtaBan()      As modprc_g_tpo_CtaBan
Dim r_arr_MatPro()      As modprc_g_tpo_MatPro
Dim r_arr_Matriz()      As modprc_g_tpo_Matriz
Dim r_arr_MatDet()      As modprc_g_tpo_MatDet
Dim r_arr_CtaPrd()      As modprc_g_tpo_CtaPrd
Dim r_rst_Genera        As ADODB.Recordset
Dim r_rst_Cuotas        As ADODB.Recordset
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_rst_CtaBan        As ADODB.Recordset
Dim r_rst_TipMat        As ADODB.Recordset
Dim r_rst_DifCam        As ADODB.Recordset
Dim r_rst_MatCtb        As ADODB.Recordset
Dim r_rst_MatCab        As ADODB.Recordset
Dim r_rst_MatDet        As ADODB.Recordset
Dim r_rst_CtaPrd        As ADODB.Recordset
Dim r_str_FecPro        As String
Dim r_str_FecVct        As String
Dim r_str_FecAsi        As String
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_int_NumAsi        As Integer
Dim r_int_PosMat        As Integer
Dim r_int_AuxCon        As Integer
Dim r_int_SitCre        As Integer
Dim r_int_Refina        As Integer
Dim r_int_Judici        As Integer
Dim r_int_Castig        As Integer
Dim r_int_Cont01        As Integer
Dim r_dbl_Totdeb_sol    As Double
Dim r_dbl_Totdeb_dol    As Double
Dim r_dbl_Tothab_sol    As Double
Dim r_dbl_Tothab_dol    As Double
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_dbl_TipCam        As Double
Dim r_dbl_TipSbs        As Double
Dim r_dbl_TipSun        As Double
Dim r_str_CueGan        As String
Dim r_str_CuePer        As String
Dim r_str_Cuenta        As String
Dim r_str_FlagDH        As String
Dim r_dbl_Asi_CapVig    As Double
Dim r_dbl_Asi_ImpITF    As Double
Dim r_dbl_Asi_ManCta    As Double
Dim r_dbl_Asi_CtaBnc    As Double
   
   r_lng_NumReg = 0
   r_lng_TotReg = ff_ConGas(p_FecIni, p_FecFin)
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1013"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
        
   'Para Obtener Cuentas de Diferencia x Tipo de Cambio (Ganancia o Pérdida)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAGEN "
   g_str_Parame = g_str_Parame & " WHERE CTAGEN_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & " ORDER BY CTAGEN_CTACTB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_DifCam, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_DifCam.BOF And r_rst_DifCam.EOF) Then
      r_rst_DifCam.MoveFirst
      Do While Not r_rst_DifCam.EOF
         If r_rst_DifCam!CTAGEN_CODIDE = "01" Then
            r_str_CueGan = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         ElseIf r_rst_DifCam!CTAGEN_CODIDE = "02" Then
            r_str_CuePer = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         End If
                   
         r_rst_DifCam.MoveNext
      Loop
   End If
   
   r_rst_DifCam.Close
   Set r_rst_DifCam = Nothing
   
   'Para leer cuentas por cada Cuenta Bancaria
   ReDim r_arr_CtaBan(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN "
   g_str_Parame = g_str_Parame & " ORDER BY CTABAN_CODBAN, CTABAN_NUMCTA ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaBan, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaBan.BOF And r_rst_CtaBan.EOF) Then
      r_rst_CtaBan.MoveFirst
      Do While Not r_rst_CtaBan.EOF
         ReDim Preserve r_arr_CtaBan(UBound(r_arr_CtaBan) + 1)
          
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CodBan = r_rst_CtaBan!CtaBan_CodBan
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_NumCta = Trim(r_rst_CtaBan!CtaBan_NumCta)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CtaCtb = Trim(r_rst_CtaBan!CtaBan_CtaCtb)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_TipCta = r_rst_CtaBan!CtaBan_TipCta
         r_rst_CtaBan.MoveNext
      Loop
   End If
   
   r_rst_CtaBan.Close
   Set r_rst_CtaBan = Nothing
   
   'Para leer Matrices Contables por Cada Producto, por cada Moneda
   ReDim r_arr_MatDet(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATCAB "
   g_str_Parame = g_str_Parame & " WHERE MATCAB_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & "   AND MATCAB_TIPMAT = '110001' "
   g_str_Parame = g_str_Parame & " ORDER BY MATCAB_SITCRE, SUBSTR(MATCAB_CODMAT,1,3) ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatCab, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_MatCab.BOF And r_rst_MatCab.EOF) Then
      'Por cada producto
      r_rst_MatCab.MoveFirst
      Do While Not r_rst_MatCab.EOF
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATDET "
         g_str_Parame = g_str_Parame & " WHERE MATDET_CODMAT = '" & r_rst_MatCab!MATCAB_CODMAT & "' "
         g_str_Parame = g_str_Parame & " ORDER BY MATDET_CODMAT, MATDET_NUMITE ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatDet, 3) Then
            Exit Sub
         End If
               
         'Bucle por Detalle de Matriz
         r_rst_MatDet.MoveFirst
         Do While Not r_rst_MatDet.EOF
            ReDim Preserve r_arr_MatDet(UBound(r_arr_MatDet) + 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodMat = r_rst_MatCab!MATCAB_CODMAT
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodPrd = Trim(Mid(r_rst_MatCab!MATCAB_CODMAT, 1, 3))
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesCab = Trim(r_rst_MatCab!MATCAB_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipMon = r_rst_MatCab!MATCAB_TIPMON
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_SitCre = r_rst_MatCab!MATCAB_SITCRE
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesDet = Trim(r_rst_MatDet!MATDET_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CtbCon = Trim(r_rst_MatDet!MATDET_CONCTB)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DebHab = Left(moddat_gf_Consulta_ParDes("255", CStr(r_rst_MatDet!MATDET_FLGDHB)), 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipCam = CInt(r_rst_MatDet!MATDET_TIPTCA)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_NroLib = CInt(r_rst_MatCab!MATCAB_CODLIB)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_OpeCon = Trim(r_rst_MatDet!MATDET_CONOPE)
            r_rst_MatDet.MoveNext
         Loop
         
         r_rst_MatDet.Close
         Set r_rst_MatDet = Nothing
      
         r_rst_MatCab.MoveNext
      Loop
   End If
   
   r_rst_MatCab.Close
   Set r_rst_MatCab = Nothing
      
   'Para leer cuentas Contables
   ReDim r_arr_CtaPrd(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRD "
   g_str_Parame = g_str_Parame & " ORDER BY CTAPRD_CODPRD ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaPrd, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaPrd.BOF And r_rst_CtaPrd.EOF) Then
      r_rst_CtaPrd.MoveFirst
      Do While Not r_rst_CtaPrd.EOF
         ReDim Preserve r_arr_CtaPrd(UBound(r_arr_CtaPrd) + 1)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CodPrd = Trim(r_rst_CtaPrd!CtaPrd_CodPrd)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtbCon = Trim(r_rst_CtaPrd!CTAPRD_CONCTB)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_SitCre = Trim(r_rst_CtaPrd!CTAPRD_TIPCRE)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtaCtb = Trim(r_rst_CtaPrd!CtaPrd_CtaCtb)
         r_rst_CtaPrd.MoveNext
      Loop
   End If
   
   r_rst_CtaPrd.Close
   Set r_rst_CtaPrd = Nothing
   
   'LEYENDO CURSOR PRINCIPAL
   modprc_g_str_CadEje = "SELECT CAJMOV_SUCMOV, CAJMOV_USUMOV, CAJMOV_FECMOV, CAJMOV_FECDEP, CAJMOV_NUMMOV, CAJMOV_TIPDOC, CAJMOV_NUMDOC, " & _
                         "       CAJMOV_NUMOPE, CAJMOV_ITFIMP, CAJMOV_MONPAG, CAJMOV_IMPTOT, CAJMOV_IMPPAG, CAJMOV_CODBAN, CAJMOV_NUMCTA  " & _
                         "  FROM OPE_CAJMOV " & _
                         " WHERE CAJMOV_TIPMOV = 1101 " & _
                         "   AND CAJMOV_CTBFLG = 0 " & _
                         "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " " & _
                         "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         " ORDER BY CAJMOV_NUMMOV ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Fecha de Proceso
         r_str_FecPro = Format(p_FecFin, "dd/mm/yyyy")
                  
         If CDate(gf_FormatoFecha(modprc_g_rst_Princi!CAJMOV_FECMOV)) > CDate(p_CtbFin) Then
            r_str_FecAsi = p_CtbFin
         Else
            r_str_FecAsi = gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))
         End If
         
         r_dbl_Asi_CapVig = modprc_g_rst_Princi!CAJMOV_IMPPAG
         r_dbl_Asi_ImpITF = modprc_g_rst_Princi!CAJMOV_ITFIMP
         r_dbl_Tothab_sol = 0
         r_dbl_Totdeb_sol = 0
         
         ReDim r_arr_Matriz(0)
         r_int_PosMat = 0
         
         For r_int_Contad = 1 To UBound(r_arr_MatDet)
                  
            'Comparacion de la Matriz con el Producto r_arr_MatCtb(UBound(r_arr_MatCtb)).MatCtb_CodPrd
            If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE), 1, 3) Then
               r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecAsi, "yyyymmdd"), 1)
               r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
                              
               If r_int_PosMat = 0 Then
                  r_int_PosMat = r_int_Contad
               End If
               
               ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
                  
               For r_int_AuxCon = 1 To UBound(r_arr_CtaPrd)
                  If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CodPrd = r_arr_MatDet(r_int_Contad).MatDet_CodPrd Then
                     If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtbCon = r_arr_MatDet(r_int_Contad).MatDet_CtbCon Then
                        r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtaCtb
                     End If
                  End If
               Next r_int_AuxCon
                  
               r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = r_arr_MatDet(r_int_Contad).MatDet_DebHab
               r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_OpeCon = r_arr_MatDet(r_int_Contad).MatDet_OpeCon
               
               For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                  If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_TipCta = r_arr_CtaBan(r_int_Cont01).CtaBan_TipCta
                     Exit For
                  End If
               Next r_int_Cont01
               
               If r_arr_MatDet(r_int_Contad).MatDet_OpeCon = "100422" Or r_arr_MatDet(r_int_Contad).MatDet_OpeCon = "100423" Then
                  If r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_TipCta = "000002" Then
                     If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
                        r_dbl_Asi_ManCta = modprc_g_dbl_ComBcoDol
                        r_dbl_Asi_CtaBnc = modprc_g_dbl_ComBcoDol
                     Else
                        r_dbl_Asi_ManCta = modprc_g_dbl_ComBcoSol
                        r_dbl_Asi_CtaBnc = modprc_g_dbl_ComBcoSol
                     End If
                  Else
                     r_dbl_Asi_ManCta = 0
                     r_dbl_Asi_CtaBnc = 0
                  End If
               End If
               
               Select Case r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                  Case "100602": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapVig 'Capital Vigente
                  Case "100413": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ImpITF 'Importe de ITF
                  Case "100422": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ManCta 'Comision Recaudo
                  Case "100423": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CtaBnc 'Cuenta Bancaria
                  Case "100201"
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = modprc_g_rst_Princi!CAJMOV_IMPTOT 'Importe Depositado
                     For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                        If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaBan(r_int_Cont01).CtaBan_CtaCtb
                           Exit For
                        End If
                     Next r_int_Cont01
               End Select
                  
               r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
            End If
         Next r_int_Contad
                  
         'Generar Numero de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, "LM", r_arr_MatDet(r_int_PosMat).MatDet_NroLib)
         
         'Ingresar Cabecera
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_PosMat).MatDet_NroLib, r_int_NumAsi, Format(modprc_g_rst_Princi!CAJMOV_MONPAG, "000"), r_dbl_TipSbs, "O", Mid((Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + Trim(r_arr_MatDet(r_int_PosMat).MatDet_DesCab)), 1, 60), r_str_FecAsi, "1")
                  
         'Ingresar Detalles
         r_int_NumIte = 1
         For r_int_Contad = 1 To UBound(r_arr_Matriz)
            
            If modprc_g_rst_Princi!CAJMOV_MONPAG = 1 Then
               If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100413" Then
                  'Operacion de Truncado a 2 digitos por ITF
                  r_dbl_ImpSol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
                  r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))      'Truncar
               Else
                  r_dbl_ImpSol = r_arr_Matriz(r_int_Contad).Matriz_import
                  r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSbs, "#######0.00"))
               End If
            ElseIf modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
               If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100413" Then
                  'Operacion de Truncado a 2 digitos por ITF
                  r_dbl_ImpSol = CDbl(Format(gf_Truncar_Numero(gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import) * r_dbl_TipSun, 2), "########0.00"))      'Truncar
                  r_dbl_ImpDol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
               Else
                  r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import * r_dbl_TipSbs, "########0.00"))
                  r_dbl_ImpDol = r_arr_Matriz(r_int_Contad).Matriz_import
               End If
            End If
                           
            'Acumulacion de Haber y Debe
            If r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "H" Then
               r_dbl_Tothab_sol = r_dbl_Tothab_sol + r_dbl_ImpSol
            ElseIf r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "D" Then
               r_dbl_Totdeb_sol = r_dbl_Totdeb_sol + r_dbl_ImpSol
            End If
            
            'Inserccion de las Cuentas al Detalle del Asiento
            If r_dbl_ImpSol <> 0 And r_dbl_ImpDol <> 0 Then
               If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100201" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100422" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100423" Then
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Mid((Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + Trim(r_arr_Matriz(r_int_Contad).Matriz_DesNot)), 1, 60), r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))  'Contabilizacion de Fecha en el campo Auxiliar
               Else
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, r_str_FecAsi, Mid((Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + Trim(r_arr_Matriz(r_int_Contad).Matriz_DesNot)), 1, 60), r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV)))) 'Contabilizacion de Fecha en el campo Auxiliar
               End If
               r_int_NumIte = r_int_NumIte + 1
            End If
            
         Next r_int_Contad
         
         If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
            r_dbl_ImpSol = 0
         
            'Diferencia entre el DEBE y HABER para agregar la cuenta Diferencia por Tipo de Cambio
            If r_dbl_Totdeb_sol > r_dbl_Tothab_sol Then
               r_str_Cuenta = r_str_CueGan
               r_str_FlagDH = "H"
               r_dbl_ImpSol = CDbl(Format(r_dbl_Totdeb_sol - r_dbl_Tothab_sol, "#####0.00"))
            End If
            
            If r_dbl_Totdeb_sol < r_dbl_Tothab_sol Then
               r_str_Cuenta = r_str_CuePer
               r_str_FlagDH = "D"
               r_dbl_ImpSol = CDbl(Format(r_dbl_Tothab_sol - r_dbl_Totdeb_sol, "######0.00"))
            End If
            
            'Inserccion de la Cuenta por Diferencia de Tipo de Cambio en el Detalle del Asiento
            If r_dbl_ImpSol > 0 Then
               Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad - 1).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_str_Cuenta, r_str_FecAsi, Mid((Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + "DIF. TIP. CAM."), 1, 60), r_str_FlagDH, r_dbl_ImpSol, 0, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            End If
         End If
         
         'Actualizando en OPE_CAJMOV para diferenciar los movimientos procesados en el mismo dia
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE OPE_CAJMOV "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET CAJMOV_CTBFLG = 1 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_USUMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_USUMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_NUMMOV = " & modprc_g_rst_Princi!CAJMOV_NUMMOV & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla OPE_CAJMOV - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         'Leyendo siguiente Movimiento
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
         
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Loop
   
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1014(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerIni As String, ByVal p_PerFin As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_CtbIni As String, ByVal p_CtbFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1007
'Descripción      :  Registro de Desembolso de Créditos Hipotecarios
'Resumen          :  Contabilización de Desembolsos de Créditos Hipotecarios
'F. Creación      :  24-03-2010
'U. Creación      :  Jorge Luis Tacuche Mesia
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_arr_CtaBan()      As modprc_g_tpo_CtaBan
Dim r_arr_MatPro()      As modprc_g_tpo_MatPro
Dim r_arr_Matriz()      As modprc_g_tpo_Matriz
Dim r_arr_MatDet()      As modprc_g_tpo_MatDet
Dim r_arr_CtaPrd()      As modprc_g_tpo_CtaPrd
Dim r_arr_CtaCom()      As modprc_g_tpo_CtaCom
Dim r_rst_Genera        As ADODB.Recordset
Dim r_rst_Cuotas        As ADODB.Recordset
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_rst_CtaBan        As ADODB.Recordset
Dim r_rst_TipMat        As ADODB.Recordset
Dim r_rst_DifCam        As ADODB.Recordset
Dim r_rst_MatCtb        As ADODB.Recordset
Dim r_rst_MatCab        As ADODB.Recordset
Dim r_rst_MatDet        As ADODB.Recordset
Dim r_rst_CtaPrd        As ADODB.Recordset
Dim r_rst_CtaCom        As ADODB.Recordset
Dim r_str_FecPro        As String
Dim r_str_FecVct        As String
Dim r_str_FecAsi        As String
Dim r_str_FecDes        As String
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_int_NumAsi        As Integer
Dim r_int_PosMat        As Integer
Dim r_int_AuxCon        As Integer
Dim r_int_SitCre        As Integer
Dim r_int_Refina        As Integer
Dim r_int_Judici        As Integer
Dim r_int_Castig        As Integer
Dim r_int_Cont01        As Integer
Dim r_dbl_Totdeb_sol    As Double
Dim r_dbl_Totdeb_dol    As Double
Dim r_dbl_Tothab_sol    As Double
Dim r_dbl_Tothab_dol    As Double
Dim r_dbl_ImpSol        As Double
Dim r_dbl_ImpDol        As Double
Dim r_dbl_TipCam        As Double
Dim r_str_CueGan        As String
Dim r_str_CuePer        As String
Dim r_str_Cuenta        As String
Dim r_str_FlagDH        As String
Dim r_dbl_TipSbs        As Double
Dim r_dbl_TipSun        As Double
Dim r_dbl_Asi_CapVig    As Double
Dim r_dbl_Asi_ImpITF    As Double
Dim r_int_TipVen        As Integer
Dim r_str_NumVen        As String
Dim r_str_CodMod        As String
Dim r_str_CodPry        As String
Dim r_int_CtaCom        As Integer
Dim r_int_PryMcs        As Integer
Dim r_str_DesCom        As String
Dim r_str_Cadena        As String
   
   r_lng_NumReg = 0
   r_lng_TotReg = ff_ConDes(p_FecIni, p_FecFin)
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1014"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
        
   'Para Obtener Cuentas de Diferencia x Tipo de Cambio (Ganancia o Pérdida)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAGEN "
   g_str_Parame = g_str_Parame & " WHERE CTAGEN_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & " ORDER BY CTAGEN_CTACTB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_DifCam, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_DifCam.BOF And r_rst_DifCam.EOF) Then
      r_rst_DifCam.MoveFirst
      Do While Not r_rst_DifCam.EOF
         If r_rst_DifCam!CTAGEN_CODIDE = "01" Then
            r_str_CueGan = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         ElseIf r_rst_DifCam!CTAGEN_CODIDE = "02" Then
            r_str_CuePer = Trim(r_rst_DifCam!CTAGEN_CTACTB)
         End If
                   
         r_rst_DifCam.MoveNext
      Loop
   End If
   
   r_rst_DifCam.Close
   Set r_rst_DifCam = Nothing
   
   'Para leer cuentas por cada Cuenta Bancaria
   ReDim r_arr_CtaBan(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN "
   g_str_Parame = g_str_Parame & " ORDER BY CTABAN_CODBAN, CTABAN_NUMCTA ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaBan, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaBan.BOF And r_rst_CtaBan.EOF) Then
      r_rst_CtaBan.MoveFirst
      Do While Not r_rst_CtaBan.EOF
         ReDim Preserve r_arr_CtaBan(UBound(r_arr_CtaBan) + 1)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CodBan = r_rst_CtaBan!CtaBan_CodBan
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_NumCta = Trim(r_rst_CtaBan!CtaBan_NumCta)
         r_arr_CtaBan(UBound(r_arr_CtaBan)).CtaBan_CtaCtb = Trim(r_rst_CtaBan!CtaBan_CtaCtb)
         r_rst_CtaBan.MoveNext
      Loop
   End If
   
   r_rst_CtaBan.Close
   Set r_rst_CtaBan = Nothing
   
   'Para leer Matrices Contables por Cada Producto, por cada Moneda
   ReDim r_arr_MatDet(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATCAB "
   g_str_Parame = g_str_Parame & " WHERE MATCAB_CODEMP = '000001' "
   g_str_Parame = g_str_Parame & "   AND MATCAB_TIPMAT = '100002' "
   g_str_Parame = g_str_Parame & " ORDER BY MATCAB_SITCRE, SUBSTR(MATCAB_CODMAT,1,3) ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatCab, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_MatCab.BOF And r_rst_MatCab.EOF) Then
      'Por cada producto
      r_rst_MatCab.MoveFirst
      Do While Not r_rst_MatCab.EOF
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATDET "
         g_str_Parame = g_str_Parame & " WHERE MATDET_CODMAT = '" & r_rst_MatCab!MATCAB_CODMAT & "' "
         g_str_Parame = g_str_Parame & " ORDER BY MATDET_CODMAT, MATDET_NUMITE ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_MatDet, 3) Then
            Exit Sub
         End If
               
         'Bucle por Detalle de Matriz
         r_rst_MatDet.MoveFirst
         Do While Not r_rst_MatDet.EOF
            ReDim Preserve r_arr_MatDet(UBound(r_arr_MatDet) + 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodMat = r_rst_MatCab!MATCAB_CODMAT
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodPrd = Trim(Mid(r_rst_MatCab!MATCAB_CODMAT, 1, 3))
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesCab = Trim(r_rst_MatCab!MATCAB_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipMon = r_rst_MatCab!MATCAB_TIPMON
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesDet = Trim(r_rst_MatDet!MATDET_DESCRI)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CtbCon = Trim(r_rst_MatDet!MATDET_CONCTB)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DebHab = Left(moddat_gf_Consulta_ParDes("255", CStr(r_rst_MatDet!MATDET_FLGDHB)), 1)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipCam = CInt(r_rst_MatDet!MATDET_TIPTCA)
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_NroLib = r_rst_MatCab!MATCAB_CODLIB
            r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_OpeCon = Trim(r_rst_MatDet!MATDET_CONOPE)
                                   
            r_rst_MatDet.MoveNext
         Loop
         
         r_rst_MatDet.Close
         Set r_rst_MatDet = Nothing
         r_rst_MatCab.MoveNext
      Loop
   End If
   
   r_rst_MatCab.Close
   Set r_rst_MatCab = Nothing
      
   'Para leer cuentas Contables
   ReDim r_arr_CtaPrd(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRD "
   g_str_Parame = g_str_Parame & " ORDER BY CTAPRD_CODPRD ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaPrd, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaPrd.BOF And r_rst_CtaPrd.EOF) Then
      r_rst_CtaPrd.MoveFirst
      Do While Not r_rst_CtaPrd.EOF
         ReDim Preserve r_arr_CtaPrd(UBound(r_arr_CtaPrd) + 1)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CodPrd = Trim(r_rst_CtaPrd!CtaPrd_CodPrd)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtbCon = Trim(r_rst_CtaPrd!CTAPRD_CONCTB)
         r_arr_CtaPrd(UBound(r_arr_CtaPrd)).CtaPrd_CtaCtb = Trim(r_rst_CtaPrd!CtaPrd_CtaCtb)
         
         r_rst_CtaPrd.MoveNext
      Loop
   End If
   
   r_rst_CtaPrd.Close
   Set r_rst_CtaPrd = Nothing
   
   'Para leer Cuentas de Constructores
   ReDim r_arr_CtaCom(0)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTACOM "
   g_str_Parame = g_str_Parame & " ORDER BY CTACOM_NROCOD, CTACOM_TIPMON ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaCom, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_CtaCom.BOF And r_rst_CtaCom.EOF) Then
      r_rst_CtaCom.MoveFirst
      Do While Not r_rst_CtaCom.EOF
         ReDim Preserve r_arr_CtaCom(UBound(r_arr_CtaCom) + 1)
         r_arr_CtaCom(UBound(r_arr_CtaCom)).CtaCom_DocIde = CStr(r_rst_CtaCom!CTACOM_TDOCLI) & Trim(r_rst_CtaCom!CTACOM_NDOCLI)
         r_arr_CtaCom(UBound(r_arr_CtaCom)).CtaCom_Descri = Trim(r_rst_CtaCom!CtaCom_Descri)
         r_arr_CtaCom(UBound(r_arr_CtaCom)).CtaCom_CtaCtb = Trim(r_rst_CtaCom!CtaCom_CtaCtb)
          
         r_rst_CtaCom.MoveNext
      Loop
   End If
   
   'LEYENDO CURSOR PRINCIPAL
   modprc_g_str_CadEje = "SELECT CAJMOV_SUCMOV, CAJMOV_USUMOV, CAJMOV_FECMOV, CAJMOV_FECDEP, CAJMOV_NUMMOV, CAJMOV_NUMOPE, CAJMOV_ITFIMP, CAJMOV_MONPAG, CAJMOV_IMPTOT, CAJMOV_IMPPAG, CAJMOV_CODBAN, CAJMOV_NUMCTA FROM OPE_CAJMOV " & _
                         " WHERE CAJMOV_TIPMOV = 1103 " & _
                         "   AND CAJMOV_CTBFLG = 0 " & _
                         "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " " & _
                         "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         " ORDER BY CAJMOV_NUMMOV ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Fecha de Proceso
         r_str_FecPro = Format(p_FecFin, "dd/mm/yyyy")
         
         If CDate(gf_FormatoFecha(modprc_g_rst_Princi!CAJMOV_FECMOV)) > CDate(p_CtbFin) Then
            r_str_FecAsi = p_CtbFin
         Else
            r_str_FecAsi = gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))
         End If
         
         r_dbl_Asi_CapVig = modprc_g_rst_Princi!CAJMOV_IMPPAG
         r_dbl_Asi_ImpITF = modprc_g_rst_Princi!CAJMOV_ITFIMP
         r_dbl_Tothab_sol = 0
         r_dbl_Totdeb_sol = 0
         
         'Para obtener Datos del Proyecto y Cosntructor
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT HIPMAE_NUMOPE, HIPMAE_CODPRD, HIPMAE_FECDES, HIPMAE_CODSUB, HIPMAE_CODMOD, HIPMAE_PRYMCS, HIPMAE_PRYINM "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
         g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & "' "
         g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_NUMOPE ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_CtaBan, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_CtaBan.BOF And r_rst_CtaBan.EOF) Then
            r_rst_CtaBan.MoveFirst
            If r_rst_CtaBan!HIPMAE_PRYMCS = 1 Then
               r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = " & r_rst_CtaBan!HIPMAE_PRYINM & " "
               
               If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Listas, 3) Then
                  Exit Sub
               End If
               
               r_int_TipVen = g_rst_Listas!DATGEN_VENTDO
               r_str_NumVen = Trim(g_rst_Listas!DATGEN_VENNDO)
               r_str_CodMod = r_rst_CtaBan!HIPMAE_CODMOD
               r_str_CodPry = r_rst_CtaBan!HIPMAE_PRYINM
            End If
            
            r_int_PryMcs = r_rst_CtaBan!HIPMAE_PRYMCS
            r_str_FecDes = gf_FormatoFecha(CStr(r_rst_CtaBan!HIPMAE_FECDES))
            r_rst_CtaBan.MoveNext
         End If
         
         r_rst_CtaBan.Close
         Set r_rst_CtaBan = Nothing
            
         ReDim r_arr_Matriz(0)
         r_int_PosMat = 0
         
         For r_int_Contad = 1 To UBound(r_arr_MatDet)
            'Comparacion de la Matriz con el Producto r_arr_MatCtb(UBound(r_arr_MatCtb)).MatCtb_CodPrd
            If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = Mid(Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE), 1, 3) Then
               r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecDes, "yyyymmdd"), 1)
               r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecDes, "yyyymmdd"), 2)
               
               If r_int_PosMat = 0 Then
                  r_int_PosMat = r_int_Contad
               End If
               
               ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
                  
               For r_int_AuxCon = 1 To UBound(r_arr_CtaPrd)
                  'Verifica que tengan del mismo producto
                  If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CodPrd = r_arr_MatDet(r_int_Contad).MatDet_CodPrd Then
                  
                     'Verifica que tengan el mismo concepto contable
                     If r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtbCon = r_arr_MatDet(r_int_Contad).MatDet_CtbCon Then
                        
                        'Proyecto No Vinculado
                        If r_int_PryMcs = 2 Then
                           If r_arr_MatDet(r_int_Contad).MatDet_CtbCon = "000025" Then
                              If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = "001" Or r_arr_MatDet(r_int_Contad).MatDet_CodPrd = "002" Then
                                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = "252419010101"
                              Else
                                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = "251419010101"
                              End If
                           Else
                              r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtaCtb
                           End If
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_OpeCon = r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                           r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_ConCtb = r_arr_MatDet(r_int_Contad).MatDet_CtbCon
                           
                        Else
                        'Proyecto Vinculado
                           If r_arr_MatDet(r_int_Contad).MatDet_CtbCon = "000025" Then
                              
                              For r_int_CtaCom = 1 To UBound(r_arr_CtaCom)
                                 
                                 If Trim(r_arr_CtaCom(r_int_CtaCom).CtaCom_DocIde) = Trim(CStr(r_int_TipVen) & r_str_NumVen) Then
                                    If Mid(r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtaCtb, 3, 1) = Mid(r_arr_CtaCom(r_int_CtaCom).CtaCom_CtaCtb, 3, 1) Then
                                       r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaCom(r_int_CtaCom).CtaCom_CtaCtb
                                       r_str_DesCom = r_arr_CtaCom(r_int_CtaCom).CtaCom_Descri
                                       r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_OpeCon = r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                                       r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_ConCtb = r_arr_MatDet(r_int_Contad).MatDet_CtbCon
                                    End If
                                 End If
                                 
                              Next r_int_CtaCom
                           Else
                              r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaPrd(r_int_AuxCon).CtaPrd_CtaCtb
                              r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_OpeCon = r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                           End If
                        End If
                     End If
                  End If
               Next r_int_AuxCon
               
               r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = r_arr_MatDet(r_int_Contad).MatDet_DebHab
               
               Select Case r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                  Case "100101": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_CapVig 'Capital Vigente
                  Case "100202": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ImpITF 'Importe de ITF
                  Case "100104": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Asi_ImpITF 'Importe de ITF
                  Case "100201": r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = modprc_g_rst_Princi!CAJMOV_IMPTOT 'Importe Depositado
                        
                  For r_int_Cont01 = 1 To UBound(r_arr_CtaBan)
                     If r_arr_CtaBan(r_int_Cont01).CtaBan_CodBan = modprc_g_rst_Princi!CAJMOV_CODBAN And r_arr_CtaBan(r_int_Cont01).CtaBan_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA) Then
                        r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_CtaBan(r_int_Cont01).CtaBan_CtaCtb
                        Exit For
                     End If
                  Next r_int_Cont01
               End Select
               
               If r_arr_MatDet(r_int_Contad).MatDet_CtbCon = "000025" Then
                  If r_str_DesCom = "" Then
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
                  Else
                     r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_str_DesCom)
                  End If
               Else
                  r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
               End If
               
            End If
         Next r_int_Contad
                                             
         'Generar Numero de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, "LM", r_arr_MatDet(r_int_PosMat).MatDet_NroLib)
         
         'Ingresar Cabecera
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_PosMat).MatDet_NroLib, r_int_NumAsi, Format(modprc_g_rst_Princi!CAJMOV_MONPAG, "000"), r_dbl_TipSbs, "O", r_arr_MatDet(r_int_PosMat).MatDet_DesCab + " - " + Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000"), r_str_FecDes, "1")
         
         'Ingresar Detalles
         r_int_NumIte = 1
         
         For r_int_Contad = 1 To UBound(r_arr_Matriz)
            
            If modprc_g_rst_Princi!CAJMOV_MONPAG = 1 Then
               If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100202" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100104" Then
                  'Operacion de Truncado a 2 digitos por ITF
                  r_dbl_ImpSol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
                  r_dbl_ImpDol = CDbl(Format(gf_Truncar_Numero(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSun, "######0.000000"), 2), "########0.00"))     'Truncar
               Else
                  r_dbl_ImpSol = r_arr_Matriz(r_int_Contad).Matriz_import
                  r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import / r_dbl_TipSbs, "#######0.00"))
               End If
               
            ElseIf modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
               If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100202" Or r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100104" Then
                  'Operacion de Truncado a 2 digitos por ITF
                  r_dbl_ImpSol = CDbl(Format(gf_Truncar_Numero(gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import) * r_dbl_TipSun, 2), "########0.00"))      'Truncar
                  r_dbl_ImpDol = gf_NueImp_Numero(r_arr_Matriz(r_int_Contad).Matriz_import)
               Else
                  r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_Contad).Matriz_import * r_dbl_TipSbs, "########0.00"))
                  r_dbl_ImpDol = r_arr_Matriz(r_int_Contad).Matriz_import
               End If
               
            End If
                           
            'Acumulacion de Haber y Debe
            If r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "H" Then
               r_dbl_Tothab_sol = r_dbl_Tothab_sol + r_dbl_ImpSol
            ElseIf r_arr_Matriz(r_int_Contad).Matriz_FlagDH = "D" Then
               r_dbl_Totdeb_sol = r_dbl_Totdeb_sol + r_dbl_ImpSol
            End If
            
            'Inserccion de las Cuentas al Detalle del Asiento
            If r_dbl_ImpSol <> 0 And r_dbl_ImpDol <> 0 Then
               If r_arr_Matriz(r_int_Contad).Matriz_OpeCon = "100201" Then
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, r_str_FecDes, Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + r_arr_Matriz(r_int_Contad).Matriz_DesNot, r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV)))) 'Contabilizacion de Fecha en el campo Auxiliar
               Else
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_Contad).Matriz_CtaCtb, r_str_FecDes, Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + r_arr_Matriz(r_int_Contad).Matriz_DesNot, r_arr_Matriz(r_int_Contad).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV)))) 'Contabilizacion de Fecha en el campo Auxiliar
               End If
               
               r_int_NumIte = r_int_NumIte + 1
            End If
         
         Next r_int_Contad
         
         If modprc_g_rst_Princi!CAJMOV_MONPAG = 2 Then
            r_dbl_ImpSol = 0
         
            'Diferencia entre el DEBE y HABER para agregar la cuenta Diferencia por Tipo de Cambio
            If r_dbl_Totdeb_sol > r_dbl_Tothab_sol Then
               r_str_Cuenta = r_str_CueGan
               r_str_FlagDH = "H"
               r_dbl_ImpSol = CDbl(Format(r_dbl_Totdeb_sol - r_dbl_Tothab_sol, "#####0.00"))
            End If
            
            If r_dbl_Totdeb_sol < r_dbl_Tothab_sol Then
               r_str_Cuenta = r_str_CuePer
               r_str_FlagDH = "D"
               r_dbl_ImpSol = CDbl(Format(r_dbl_Tothab_sol - r_dbl_Totdeb_sol, "######0.00"))
            End If
            
            'Inserccion de la Cuenta por Diferencia de Tipo de Cambio en el Detalle del Asiento
            If r_dbl_ImpSol > 0 Then
               Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, "LM", p_PerAno, p_PerMes, r_arr_MatDet(r_int_Contad - 1).MatDet_NroLib, r_int_NumAsi, r_int_NumIte, r_str_Cuenta, r_str_FecAsi, Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) + " - " + "001-" + Mid(modprc_g_rst_Princi!CAJMOV_FECMOV, 3, 2) + Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") + " - " + "DIF. TIP. CAM.", r_str_FlagDH, r_dbl_ImpSol, 0, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
            End If
         End If
         
         'Actualizando en OPE_CAJMOV para diferenciar los movimientos procesados en el mismo dia
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE OPE_CAJMOV "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET CAJMOV_CTBFLG = 1 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAJMOV_USUMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_USUMOV) & "' AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAJMOV_NUMMOV = " & modprc_g_rst_Princi!CAJMOV_NUMMOV & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla OPE_CAJMOV - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         'Leyendo siguiente Movimiento
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
         
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Loop
   
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1015(ByVal p_CodEmp As String, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerIni As String, ByVal p_PerFin As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_CtbIni As String, ByVal p_CtbFin As String, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP1013
'Descripción      :  Contabilización de Plan de Ahorros
'Resumen          :  Contabilización de Plan de Ahorros
'F. Creación      :  15-08-2014
'U. Creación      :  Rafael Durand Banda
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_rst_Grabar        As ADODB.Recordset
Dim r_str_FecPro        As String
Dim r_str_FecAsi        As String
Dim r_dbl_TipCam        As Double
Dim r_dbl_TipSbs        As Double
Dim r_dbl_TipSun        As Double
Dim r_int_NumLib        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumIte        As Integer
Dim r_int_NumAsi        As Integer
Dim r_str_CtaImpDeb     As String
Dim r_str_CtaImpHab     As String
Dim r_str_CtaComDeb     As String
Dim r_str_CtaComHab     As String
Dim r_dbl_ImpDebSol     As Double
Dim r_dbl_ImpHabSol     As Double
Dim r_dbl_ComDebSol     As Double
Dim r_dbl_ComHabSol     As Double
Dim r_dbl_ImpDebDol     As Double
Dim r_dbl_ImpHabDol     As Double
Dim r_dbl_ComDebDol     As Double
Dim r_dbl_ComHabDol     As Double

   r_lng_NumReg = 0
   r_lng_TotReg = ff_ConAho(p_FecIni, p_FecFin)
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1015"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'LEYENDO CURSOR PRINCIPAL
   modprc_g_str_CadEje = "SELECT CAJMOV_SUCMOV, CAJMOV_USUMOV, CAJMOV_FECMOV, CAJMOV_FECDEP, CAJMOV_NUMMOV, CAJMOV_TIPDOC, CAJMOV_NUMDOC, " & _
                         "       CAJMOV_NUMOPE, CAJMOV_ITFIMP, CAJMOV_MONPAG, CAJMOV_IMPTOT, CAJMOV_IMPPAG, CAJMOV_CODBAN, CAJMOV_NUMCTA  " & _
                         "  FROM OPE_CAJMOV " & _
                         " WHERE CAJMOV_TIPMOV = 1105 " & _
                         "   AND CAJMOV_CTBFLG = 0 " & _
                         "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " " & _
                         "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " " & _
                         " ORDER BY CAJMOV_NUMMOV ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Fecha de Proceso
         r_str_FecPro = Format(p_FecFin, "dd/mm/yyyy")
         If CDate(gf_FormatoFecha(modprc_g_rst_Princi!CAJMOV_FECMOV)) > CDate(p_CtbFin) Then
            r_str_FecAsi = p_CtbFin
         Else
            r_str_FecAsi = gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))
         End If
         
         'Tipo de cambio
         r_dbl_TipSbs = modtac_gf_ObtieneTipCamDia_3(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), Format(r_str_FecAsi, "yyyymmdd"), 1)
         r_dbl_TipSun = modtac_gf_ObtieneTipCamDia_2(2, CStr(modprc_g_rst_Princi!CAJMOV_MONPAG), CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
         
         r_dbl_ImpDebSol = 0
         r_dbl_ImpHabSol = 0
         r_dbl_ComDebSol = 0
         r_dbl_ComHabSol = 0
         r_dbl_ImpDebDol = 0
         r_dbl_ImpHabDol = 0
         r_dbl_ComDebDol = 0
         r_dbl_ComHabDol = 0
         r_str_CtaImpDeb = "111301060201"
         r_str_CtaImpHab = "251419010113"  '"291807010111"
         r_str_CtaComDeb = "421201010101"
         r_str_CtaComHab = "111301060201"
         
         Select Case modprc_g_rst_Princi!CAJMOV_MONPAG
            Case 1
               r_dbl_ImpDebSol = modprc_g_rst_Princi!CAJMOV_IMPPAG
               r_dbl_ImpHabSol = modprc_g_rst_Princi!CAJMOV_IMPPAG
               r_dbl_ImpDebDol = CDbl(Format(modprc_g_rst_Princi!CAJMOV_IMPPAG / r_dbl_TipSbs, "#####0.00"))
               r_dbl_ImpHabDol = CDbl(Format(modprc_g_rst_Princi!CAJMOV_IMPPAG / r_dbl_TipSbs, "#####0.00"))
               r_dbl_ComDebSol = modprc_g_dbl_ComBcoSol
               r_dbl_ComHabSol = modprc_g_dbl_ComBcoSol
               r_dbl_ComDebDol = CDbl(Format(modprc_g_dbl_ComBcoSol / r_dbl_TipSbs, "#####0.00"))
               r_dbl_ComHabDol = CDbl(Format(modprc_g_dbl_ComBcoSol / r_dbl_TipSbs, "#####0.00"))
            Case 2
               r_dbl_ImpDebSol = CDbl(Format(modprc_g_rst_Princi!CAJMOV_IMPPAG * r_dbl_TipSbs, "#####0.00"))
               r_dbl_ImpHabSol = CDbl(Format(modprc_g_rst_Princi!CAJMOV_IMPPAG * r_dbl_TipSbs, "#####0.00"))
               r_dbl_ImpDebDol = modprc_g_rst_Princi!CAJMOV_IMPPAG
               r_dbl_ImpHabDol = modprc_g_rst_Princi!CAJMOV_IMPPAG
               r_dbl_ComDebSol = CDbl(Format(modprc_g_dbl_ComBcoDol * r_dbl_TipSbs, "#####0.00"))
               r_dbl_ComHabSol = CDbl(Format(modprc_g_dbl_ComBcoDol * r_dbl_TipSbs, "#####0.00"))
               r_dbl_ComDebDol = modprc_g_dbl_ComBcoDol
               r_dbl_ComHabDol = modprc_g_dbl_ComBcoDol
         End Select
         
         r_str_Origen = "LM"
         r_int_NumLib = 12
         r_str_TipNot = "B"
         
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_NumLib)
         
         'Insertar en CNTBL_ASIENTO
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_NumLib, r_int_NumAsi, Format(modprc_g_rst_Princi!CAJMOV_MONPAG, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " - " & Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - PLAN AHORRO", r_str_FecAsi, "1")
         
         'Insertar en CNTBL_ASIENTO_DET
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_NumLib, r_int_NumAsi, 1, r_str_CtaImpDeb, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " - " & Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - PLAN AHORRO", "D", r_dbl_ImpDebSol, r_dbl_ImpDebDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_NumLib, r_int_NumAsi, 2, r_str_CtaImpHab, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " - " & Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - PLAN AHORRO", "H", r_dbl_ImpHabSol, r_dbl_ImpHabDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_NumLib, r_int_NumAsi, 3, r_str_CtaComDeb, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " - " & Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - PLAN AHORRO", "D", r_dbl_ComDebSol, r_dbl_ComDebDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
         Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_NumLib, r_int_NumAsi, 4, r_str_CtaComHab, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECDEP))), Trim(modprc_g_rst_Princi!CAJMOV_TIPDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMDOC) & " - " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " - " & Format(CStr(modprc_g_rst_Princi!CAJMOV_NUMMOV), "00000") & " - PLAN AHORRO", "H", r_dbl_ComHabSol, r_dbl_ComHabDol, 1, CDate(gf_FormatoFecha(CStr(modprc_g_rst_Princi!CAJMOV_FECMOV))))
         
         'Actualizando en OPE_CAJMOV para diferenciar los movimientos procesados en el mismo dia
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE OPE_CAJMOV "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET CAJMOV_CTBFLG = 1 "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_SUCMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_SUCMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_USUMOV = '" & Trim(modprc_g_rst_Princi!CAJMOV_USUMOV) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_NUMMOV = " & modprc_g_rst_Princi!CAJMOV_NUMMOV & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla OPE_CAJMOV - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
         
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Loop
      
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, Format(CDate(p_FecIni), "yyyymmdd"), Format(CDate(p_FecFin), "yyyymmdd"))
End Sub

Public Sub modprc_ctbp1016(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
Dim r_lng_NumReg              As Long
Dim r_lng_TotReg              As Long
Dim r_arr_LogPro()            As modprc_g_tpo_LogPro
Dim r_str_FecPro              As String
Dim r_rst_Grabar              As ADODB.Recordset
Dim r_dbl_TipCam_Dol          As Double
Dim r_dbl_Tot_ImpDeb_Sol      As Double
Dim r_dbl_Tot_ImpHab_Sol      As Double
Dim r_dbl_Tot_ImpDeb_Dol      As Double
Dim r_dbl_Tot_ImpHab_Dol      As Double
Dim r_str_CtaDeb              As String
Dim r_str_CtaHab              As String
Dim r_dbl_Lin_ImpDeb_Sol      As Double
Dim r_dbl_Lin_ImpHab_Sol      As Double
Dim r_dbl_Lin_ImpDeb_Dol      As Double
Dim r_dbl_Lin_ImpHab_Dol      As Double
Dim r_str_Origen              As String
Dim r_int_libro               As Integer
Dim r_str_TipNot              As String
Dim r_int_NumAsi              As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1016"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   p_BarPro.FloodPercent = 0
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT COUNT(*) TOTREG  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE AND HIPMAE_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 1  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT >= " & Mid(Format(CDate(p_FecFin), "yyyymm"), 1, 6) & "01"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT <= " & Format(CDate(p_FecFin), "yyyymmdd")
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_VIVORG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_VIVORG - HIPCUO_VIVPAG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   'consulta de seguros de inmueble
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT (HIPCUO_VIVORG - HIPCUO_VIVPAG) SEGINM_DEU, A.HIPCUO_NUMOPE, A.HIPCUO_NUMCUO, HIPMAE_MONEDA,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        HIPMAE_TDOCLI || '-' || TRIM(HIPMAE_NDOCLI) AS NUMDOC  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE AND HIPMAE_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 1  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT >= " & Mid(Format(CDate(p_FecFin), "yyyymm"), 1, 6) & "01"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT <= " & Format(CDate(p_FecFin), "yyyymmdd")
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_VIVORG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_VIVORG - HIPCUO_VIVPAG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
    
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCIE.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   r_str_Origen = "LM"
   r_int_libro = 6
   r_str_TipNot = "O"
         
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_libro)
                  
         
         'Insertar en CNTBL_ASIENTO
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, "001", r_dbl_TipCam_Dol, r_str_TipNot, _
                                       Mid("SEG. INM. REGULARIZACION " & modprc_g_rst_Princi!HIPCUO_NUMOPE & " - " & modprc_g_rst_Princi!HIPCUO_NUMCUO & " - " & _
                                       Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), p_FecFin, "1")
                                                
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Then
               r_dbl_Lin_ImpDeb_Sol = modprc_g_rst_Princi!SEGINM_DEU
               r_dbl_Lin_ImpHab_Sol = modprc_g_rst_Princi!SEGINM_DEU
               r_dbl_Lin_ImpDeb_Dol = CDbl(Format(modprc_g_rst_Princi!SEGINM_DEU / r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpHab_Dol = CDbl(Format(modprc_g_rst_Princi!SEGINM_DEU / r_dbl_TipCam_Dol, "#####0.00"))
         Else
               r_dbl_Lin_ImpDeb_Sol = CDbl(Format(modprc_g_rst_Princi!SEGINM_DEU * r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpHab_Sol = CDbl(Format(modprc_g_rst_Princi!SEGINM_DEU * r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpDeb_Dol = modprc_g_rst_Princi!SEGINM_DEU
               r_dbl_Lin_ImpHab_Dol = modprc_g_rst_Princi!SEGINM_DEU
         End If
         
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_str_CtaDeb = "191807010112"
            r_str_CtaHab = "251602010104"
         Else
            r_str_CtaDeb = "192807010112"
            r_str_CtaHab = "252602010104"
         End If
         
         'Insertar en CNTBL_ASIENTO_DET
         Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 1, r_str_CtaDeb, p_FecFin, _
                                       Mid(modprc_g_rst_Princi!HIPCUO_NUMOPE & " - " & modprc_g_rst_Princi!HIPCUO_NUMCUO & " -  SEG. INM. REGULARIZACION " & _
                                       Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), "D", r_dbl_Lin_ImpDeb_Sol, r_dbl_Lin_ImpDeb_Dol)
         Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 2, r_str_CtaHab, p_FecFin, _
                                       Mid(modprc_g_rst_Princi!HIPCUO_NUMOPE & " - " & modprc_g_rst_Princi!HIPCUO_NUMCUO & " -  SEG. INM. REGULARIZACION " & _
                                       Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), "H", r_dbl_Lin_ImpHab_Sol, r_dbl_Lin_ImpHab_Dol)
         
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1017(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
Dim r_lng_NumReg              As Long
Dim r_lng_TotReg              As Long
Dim r_arr_LogPro()            As modprc_g_tpo_LogPro
Dim r_str_FecPro              As String
Dim r_rst_Grabar              As ADODB.Recordset
Dim r_dbl_TipCam_Dol          As Double
Dim r_dbl_Tot_ImpDeb_Sol      As Double
Dim r_dbl_Tot_ImpHab_Sol      As Double
Dim r_dbl_Tot_ImpDeb_Dol      As Double
Dim r_dbl_Tot_ImpHab_Dol      As Double
Dim r_str_CtaDeb              As String
Dim r_str_CtaHab              As String
Dim r_dbl_Lin_ImpDeb_Sol      As Double
Dim r_dbl_Lin_ImpHab_Sol      As Double
Dim r_dbl_Lin_ImpDeb_Dol      As Double
Dim r_dbl_Lin_ImpHab_Dol      As Double
Dim r_str_Origen              As String
Dim r_int_libro               As Integer
Dim r_str_TipNot              As String
Dim r_int_NumAsi              As Integer
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1017"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   p_BarPro.FloodPercent = 0
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT COUNT(*) TOTREG  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE AND HIPMAE_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 1  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT >= " & Mid(Format(CDate(p_FecFin), "yyyymm"), 1, 6) & "01"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT <= " & Format(CDate(p_FecFin), "yyyymmdd")
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_DESORG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_DESORG - HIPCUO_DESPAG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = modprc_g_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   'consulta de seguros de inmueble
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT (HIPCUO_DESORG - HIPCUO_DESPAG) SEGDES_DEU, A.HIPCUO_NUMOPE, A.HIPCUO_NUMCUO, HIPMAE_MONEDA,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        HIPMAE_TDOCLI || '-' || TRIM(HIPMAE_NDOCLI) AS NUMDOC  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCUO_NUMOPE AND HIPMAE_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 1  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT >= " & Mid(Format(CDate(p_FecFin), "yyyymm"), 1, 6) & "01"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_FECVCT <= " & Format(CDate(p_FecFin), "yyyymmdd")
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_DESORG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_DESORG - HIPCUO_DESPAG > 0  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
    
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   r_str_Origen = "LM"
   r_int_libro = 6
   r_str_TipNot = "O"
         
   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      Do While Not modprc_g_rst_Princi.EOF
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_libro)
         
         'Insertar en CNTBL_ASIENTO
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, "001", r_dbl_TipCam_Dol, r_str_TipNot, _
                                       Mid("SEG. DES. REGULARIZACION " & modprc_g_rst_Princi!HIPCUO_NUMOPE & " - " & modprc_g_rst_Princi!HIPCUO_NUMCUO & " - " & _
                                       Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), p_FecFin, "1")
                                       
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Then
               r_dbl_Lin_ImpDeb_Sol = modprc_g_rst_Princi!SEGDES_DEU
               r_dbl_Lin_ImpHab_Sol = modprc_g_rst_Princi!SEGDES_DEU
               r_dbl_Lin_ImpDeb_Dol = CDbl(Format(modprc_g_rst_Princi!SEGDES_DEU / r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpHab_Dol = CDbl(Format(modprc_g_rst_Princi!SEGDES_DEU / r_dbl_TipCam_Dol, "#####0.00"))
         Else
               r_dbl_Lin_ImpDeb_Sol = CDbl(Format(modprc_g_rst_Princi!SEGDES_DEU * r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpHab_Sol = CDbl(Format(modprc_g_rst_Princi!SEGDES_DEU * r_dbl_TipCam_Dol, "#####0.00"))
               r_dbl_Lin_ImpDeb_Dol = modprc_g_rst_Princi!SEGDES_DEU
               r_dbl_Lin_ImpHab_Dol = modprc_g_rst_Princi!SEGDES_DEU
         End If
         
         If modprc_g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_str_CtaDeb = "191807010112"
            r_str_CtaHab = "251602010103"
         Else
            r_str_CtaDeb = "192807010112"
            r_str_CtaHab = "252602010103"
         End If
         
         'Insertar en CNTBL_ASIENTO_DET
         Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 1, r_str_CtaDeb, p_FecFin, _
                                       Mid(modprc_g_rst_Princi!HIPCUO_NUMOPE & " - " & modprc_g_rst_Princi!HIPCUO_NUMCUO & " -  SEG. DES. REGULARIZACION " & _
                                       Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), "D", r_dbl_Lin_ImpDeb_Sol, r_dbl_Lin_ImpDeb_Dol)
         Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 2, r_str_CtaHab, p_FecFin, _
                                       Mid(modprc_g_rst_Princi!HIPCUO_NUMOPE & " - " & modprc_g_rst_Princi!HIPCUO_NUMCUO & " -  SEG. DES. REGULARIZACION " & _
                                       Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), "H", r_dbl_Lin_ImpHab_Sol, r_dbl_Lin_ImpHab_Dol)
                                       
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1018(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
Dim r_lng_NumReg              As Long
Dim r_lng_TotReg              As Long
Dim r_arr_LogPro()            As modprc_g_tpo_LogPro
Dim r_str_FecPro              As String
Dim r_rst_Grabar              As ADODB.Recordset
Dim r_dbl_TipCam_Dol          As Double
Dim r_dbl_ImpSol              As Double
Dim r_dbl_ImpDol              As Double
Dim r_str_Origen              As String
Dim r_int_libro               As Integer
Dim r_str_TipNot              As String
Dim r_int_NumAsi              As Integer
Dim r_rst_Princi              As ADODB.Recordset
Dim r_rst_Cuotas              As ADODB.Recordset
Dim r_rst_MatDet              As ADODB.Recordset
Dim r_arr_MatDet()            As modprc_g_tpo_MatDet
Dim r_int_SitCre              As Integer
Dim r_int_Refina              As Integer
Dim r_int_Judici              As Integer
Dim r_int_Castig              As Integer
Dim r_int_Contad              As Integer
Dim r_arr_Matriz()            As modprc_g_tpo_Matriz
Dim r_bol_Estado              As Boolean
Dim r_bol_EstPrd              As Boolean
Dim r_int_NumIte              As Integer
Dim r_int_AuxCon              As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1018"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_MatDet(0)
   
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT MATCAB_CODMAT, SUBSTR(MATCAB_CODMAT,1,3) CODPRD, MATCAB_TIPMON,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        MATCAB_SITCRE, TRIM(MATDET_DESCRI) MATDET_DESCRI,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        SUBSTR(C.PARDES_DESCRI,1,1) MATDET_FLGDHB, MATCAB_CODLIB,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        TRIM(MATDET_CONOPE) MATDET_CONOPE, D.CTAPRD_CTACTB  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CTB_MATCAB A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CTB_MATDET B ON A.MATCAB_CODMAT = B.MATDET_CODMAT  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 255 AND C.PARDES_CODITE = MATDET_FLGDHB  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CTB_CTAPRD D ON D.CTAPRD_CODPRD = SUBSTR(MATCAB_CODMAT,1,3) AND D.CTAPRD_CONCTB = MATDET_CONCTB  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE MATCAB_CODEMP = '000001'  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MATCAB_TIPMAT = '100004'  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MATDET_CONOPE IN ('100326','100327','100328','100329','100330')  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY MATCAB_SITCRE, SUBSTR(MATCAB_CODMAT,1,3), MATDET_CONOPE ASC  "
          
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_MatDet, 3) Then
      Exit Sub
   End If
      
   r_rst_MatDet.MoveFirst
   Do While Not r_rst_MatDet.EOF
      ReDim Preserve r_arr_MatDet(UBound(r_arr_MatDet) + 1)
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodMat = r_rst_MatDet!MATCAB_CODMAT
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodPrd = r_rst_MatDet!CODPRD
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipMon = r_rst_MatDet!MATCAB_TIPMON
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_SitCre = r_rst_MatDet!MATCAB_SITCRE
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesDet = Trim(r_rst_MatDet!MATDET_DESCRI & "")
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DebHab = r_rst_MatDet!MATDET_FLGDHB
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_NroLib = r_rst_MatDet!MATCAB_CODLIB
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_OpeCon = Trim(r_rst_MatDet!MATDET_CONOPE & "") 'mnt_pardes 64 - conceptos operativos
      r_arr_MatDet(UBound(r_arr_MatDet)).CtaPrd_CtaCtb = r_rst_MatDet!CtaPrd_CtaCtb
      
      r_rst_MatDet.MoveNext
   Loop
   
   r_rst_MatDet.Close
   Set r_rst_MatDet = Nothing
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT COUNT(*) TOTREG  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE B ON HIPMAE_NUMOPE = HIPCUO_NUMOPE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "         AND (HIPMAE_SITUAC = 2 OR (HIPMAE_SITUAC = 6 AND HIPMAE_FECCAN = " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00")) & "))"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 3  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND SUBSTR(HIPCUO_FECVCT,1,6) =  " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00"))
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = r_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'consulta de seguros de inmueble
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT HIPCUO_NUMOPE, HIPCUO_NUMCUO, B.HIPMAE_CODPRD, HIPCUO_FECVCT, HIPCUO_FECPAG, B.HIPMAE_MONEDA,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        HIPMAE_SITACT, B.HIPMAE_REFINA, B.HIPMAE_JUDICI, B.HIPMAE_CASTIG  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE B ON HIPMAE_NUMOPE = HIPCUO_NUMOPE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "         AND (HIPMAE_SITUAC = 2 OR (HIPMAE_SITUAC = 6 AND HIPMAE_FECCAN = " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00")) & "))"
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 3  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND SUBSTR(HIPCUO_FECVCT,1,6) =  " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00"))
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
    
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
   End If

   r_str_Origen = "LM"
   r_int_libro = 6
   r_str_TipNot = "O"
         
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      Do While Not r_rst_Princi.EOF
         If InStr(moddat_g_str_AgrCME, r_rst_Princi!HIPMAE_CODPRD) Or InStr(moddat_g_str_AgrMIHG, r_rst_Princi!HIPMAE_CODPRD) Or InStr(moddat_g_str_AgrTFMV, r_rst_Princi!HIPMAE_CODPRD) Then
            'SOLO CME, MIHOGAR, FMV
            r_int_SitCre = r_rst_Princi!HIPMAE_SITACT
            r_int_Refina = r_rst_Princi!HIPMAE_REFINA
            r_int_Judici = r_rst_Princi!HIPMAE_JUDICI
            r_int_Castig = r_rst_Princi!HIPMAE_CASTIG
             
            If (r_int_SitCre = 1) Then
               r_int_SitCre = 1
            ElseIf (r_int_SitCre = 5) And (r_int_Refina = 0) And (r_int_Judici = 0) And (r_int_Castig = 0) Then
               r_int_SitCre = 1
            End If
             
            If (r_int_Refina = 1) Then
               r_int_SitCre = 4
            ElseIf (r_int_Judici = 1) Then
               r_int_SitCre = 6
            ElseIf (r_int_Castig = 1) Then
               r_int_SitCre = 3
            End If
   
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT HIPCUO_NUMCUO, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF,  "
            'modprc_g_str_CadEje = modprc_g_str_CadEje & "        DECODE(SUBSTR(HIPCUO_NUMOPE,1,3),'003', HIPCUO_COMCOF,0) AS HIPCUO_COMFMV  "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "        0 AS HIPCUO_COMFMV  "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO  "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_NUMOPE = '" & Trim(r_rst_Princi!HIPCUO_NUMOPE) & "' "
            If r_rst_Princi!HIPMAE_CODPRD = "003" Then
               modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_TIPCRO = 5 "
            Else
               modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_TIPCRO = 3 "
            End If
            'modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_NUMCUO = " & CStr(r_rst_Princi!HIPCUO_NUMCUO) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND SUBSTR(HIPCUO_FECVCT,1,6) =  " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00"))
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Cuotas, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
               r_rst_Cuotas.Close
               Set r_rst_Cuotas = Nothing
               Exit Sub
            End If
           
            r_bol_EstPrd = False
            If (r_rst_Cuotas!HIPCUO_CAPITA + r_rst_Cuotas!HIPCUO_INTERE + r_rst_Cuotas!HIPCUO_COMCOF + r_rst_Cuotas!HIPCUO_COMFMV) > 0 Then
                ReDim r_arr_Matriz(0)
                
                For r_int_Contad = 1 To UBound(r_arr_MatDet)
                    If r_bol_EstPrd = True Then
                       If r_arr_MatDet(r_int_Contad).MatDet_CodPrd <> r_rst_Princi!HIPMAE_CODPRD Then
                          Exit For
                       End If
                    End If
                    
                    If r_arr_MatDet(r_int_Contad).MatDet_SitCre = r_int_SitCre Then
                       If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = r_rst_Princi!HIPMAE_CODPRD Then
                          r_bol_EstPrd = True
                          r_bol_Estado = False
                          
                          Select Case r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                                 Case "100326": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- CAPITAL PASIVO. MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_CAPITA
                                                r_bol_Estado = True
                                 Case "100327": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT. - INTERES PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_INTERE
                                                r_bol_Estado = True
                                 Case "100328": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- COMISION COF. PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_COMCOF
                                                r_bol_Estado = True
                                 Case "100329": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- COMISION FMV. PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_COMFMV
                                                r_bol_Estado = True
                                 Case "100330": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- PROV. PAG. PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_CAPITA + r_rst_Cuotas!HIPCUO_INTERE + _
                                                                                                   r_rst_Cuotas!HIPCUO_COMCOF + r_rst_Cuotas!HIPCUO_COMFMV
                                                r_bol_Estado = True
                          End Select
                          If r_bol_Estado = True Then
                             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_MatDet(r_int_Contad).CtaPrd_CtaCtb
                             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
                             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = r_arr_MatDet(r_int_Contad).MatDet_DebHab
                          End If
                       End If
                    End If
                Next
                If r_bol_EstPrd = True Then
                   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_libro)
                       
                   'Ingresar Cabecera  r_rst_Princi!HIPMAE_MONEDA
                   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, Format(1, "000"), _
                                                 r_dbl_TipCam_Dol, "O", Mid("CUOTA MES PASIVO COFIDE FMV" + " - " + Trim(r_rst_Princi!HIPCUO_NUMOPE) & " - " & _
                                                 r_rst_Cuotas!HIPCUO_NUMCUO & " - " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), p_FecFin, "1")
                   r_int_NumIte = 1
                   For r_int_AuxCon = 1 To UBound(r_arr_Matriz)
                       r_dbl_ImpSol = 0
                       r_dbl_ImpDol = 0
                       If r_arr_Matriz(r_int_AuxCon).Matriz_import > 0 Then
                          If r_rst_Princi!HIPMAE_MONEDA = 1 Then
                             r_dbl_ImpSol = r_arr_Matriz(r_int_AuxCon).Matriz_import
                             r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_AuxCon).Matriz_import / r_dbl_TipCam_Dol, "#######0.00"))
                          ElseIf r_rst_Princi!HIPMAE_MONEDA = 2 Then
                             r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_AuxCon).Matriz_import * r_dbl_TipCam_Dol, "########0.00"))
                             r_dbl_ImpDol = r_arr_Matriz(r_int_AuxCon).Matriz_import
                          End If
                           
                          Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_AuxCon).Matriz_CtaCtb, _
                                                               p_FecFin, Mid(Trim(r_rst_Princi!HIPCUO_NUMOPE) & " - " & r_rst_Cuotas!HIPCUO_NUMCUO & " - " + _
                                                               UCase(r_arr_Matriz(r_int_AuxCon).Matriz_DesNot) & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), _
                                                               r_arr_Matriz(r_int_AuxCon).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(p_FecFin))
           
                          r_int_NumIte = r_int_NumIte + 1
                       Else
                          r_int_NumIte = r_int_NumIte
                       End If
                   Next
                End If
            End If
         End If
         'Leyendo siguiente registro
         r_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
      
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Public Sub modprc_ctbp1018_OLD(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
Dim r_lng_NumReg              As Long
Dim r_lng_TotReg              As Long
Dim r_arr_LogPro()            As modprc_g_tpo_LogPro
Dim r_str_FecPro              As String
Dim r_rst_Grabar              As ADODB.Recordset
Dim r_dbl_TipCam_Dol          As Double
Dim r_dbl_ImpSol              As Double
Dim r_dbl_ImpDol              As Double
'Dim r_str_CtaDeb              As String
'Dim r_str_CtaHab              As String
Dim r_str_Origen              As String
Dim r_int_libro               As Integer
Dim r_str_TipNot              As String
Dim r_int_NumAsi              As Integer
Dim r_rst_Princi              As ADODB.Recordset
Dim r_rst_Cuotas              As ADODB.Recordset
Dim r_rst_MatDet              As ADODB.Recordset
Dim r_arr_MatDet()            As modprc_g_tpo_MatDet
Dim r_int_SitCre              As Integer
Dim r_int_Refina              As Integer
Dim r_int_Judici              As Integer
Dim r_int_Castig              As Integer
Dim r_int_Contad              As Integer
Dim r_arr_Matriz()            As modprc_g_tpo_Matriz
Dim r_bol_Estado              As Boolean
Dim r_bol_EstPrd              As Boolean
Dim r_int_NumIte              As Integer
Dim r_int_AuxCon              As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1018"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   p_BarPro.FloodPercent = 0
   
   ReDim r_arr_MatDet(0)
   
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT MATCAB_CODMAT, SUBSTR(MATCAB_CODMAT,1,3) CODPRD, MATCAB_TIPMON,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        MATCAB_SITCRE, TRIM(MATDET_DESCRI) MATDET_DESCRI,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        SUBSTR(C.PARDES_DESCRI,1,1) MATDET_FLGDHB, MATCAB_CODLIB,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        TRIM(MATDET_CONOPE) MATDET_CONOPE, D.CTAPRD_CTACTB  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CTB_MATCAB A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CTB_MATDET B ON A.MATCAB_CODMAT = B.MATDET_CODMAT  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 255 AND C.PARDES_CODITE = MATDET_FLGDHB  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CTB_CTAPRD D ON D.CTAPRD_CODPRD = SUBSTR(MATCAB_CODMAT,1,3) AND D.CTAPRD_CONCTB = MATDET_CONCTB  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE MATCAB_CODEMP = '000001'  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MATCAB_TIPMAT = '100004'  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND MATDET_CONOPE IN ('100326','100327','100328','100329','100330')  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY MATCAB_SITCRE, SUBSTR(MATCAB_CODMAT,1,3), MATDET_CONOPE ASC  "
          
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_MatDet, 3) Then
      Exit Sub
   End If
      
   r_rst_MatDet.MoveFirst
   Do While Not r_rst_MatDet.EOF
      ReDim Preserve r_arr_MatDet(UBound(r_arr_MatDet) + 1)
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodMat = r_rst_MatDet!MATCAB_CODMAT
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_CodPrd = r_rst_MatDet!CODPRD
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_TipMon = r_rst_MatDet!MATCAB_TIPMON
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_SitCre = r_rst_MatDet!MATCAB_SITCRE
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DesDet = Trim(r_rst_MatDet!MATDET_DESCRI & "")
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_DebHab = r_rst_MatDet!MATDET_FLGDHB
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_NroLib = r_rst_MatDet!MATCAB_CODLIB
      r_arr_MatDet(UBound(r_arr_MatDet)).MatDet_OpeCon = Trim(r_rst_MatDet!MATDET_CONOPE & "") 'mnt_pardes 64 - conceptos operativos
      r_arr_MatDet(UBound(r_arr_MatDet)).CtaPrd_CtaCtb = r_rst_MatDet!CtaPrd_CtaCtb
      
      r_rst_MatDet.MoveNext
   Loop
   
   r_rst_MatDet.Close
   Set r_rst_MatDet = Nothing
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Para sacar Total de Registros a leer
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT COUNT(*) TOTREG  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE B ON HIPMAE_NUMOPE = HIPCUO_NUMOPE AND HIPMAE_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 1  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND SUBSTR(HIPCUO_FECVCT,1,6) =  " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00"))
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_lng_TotReg = r_rst_Princi!TOTREG
   r_lng_NumReg = 0
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'consulta de seguros de inmueble
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT HIPCUO_NUMOPE, HIPCUO_NUMCUO, B.HIPMAE_CODPRD, HIPCUO_FECVCT, HIPCUO_FECPAG, B.HIPMAE_MONEDA,  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "        HIPMAE_SITACT, B.HIPMAE_REFINA, B.HIPMAE_JUDICI, B.HIPMAE_CASTIG  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO A  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  INNER JOIN CRE_HIPMAE B ON HIPMAE_NUMOPE = HIPCUO_NUMOPE AND HIPMAE_SITUAC = 2  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_TIPCRO = 1  "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND SUBSTR(HIPCUO_FECVCT,1,6) =  " & CStr(p_PerAno) & CStr(Format(p_PerMes, "00"))
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  ORDER BY HIPCUO_NUMOPE ASC  "
    
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
   End If

   r_str_Origen = "LM"
   r_int_libro = 6
   r_str_TipNot = "O"
         
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      Do While Not r_rst_Princi.EOF
         If InStr(moddat_g_str_AgrCME, r_rst_Princi!HIPMAE_CODPRD) Or InStr(moddat_g_str_AgrMIHG, r_rst_Princi!HIPMAE_CODPRD) Or InStr(moddat_g_str_AgrTFMV, r_rst_Princi!HIPMAE_CODPRD) Then
            'SOLO CME, MIHOGAR, FMV
            r_int_SitCre = r_rst_Princi!HIPMAE_SITACT
            r_int_Refina = r_rst_Princi!HIPMAE_REFINA
            r_int_Judici = r_rst_Princi!HIPMAE_JUDICI
            r_int_Castig = r_rst_Princi!HIPMAE_CASTIG
             
            If (r_int_SitCre = 1) Then
               r_int_SitCre = 1
            ElseIf (r_int_SitCre = 5) And (r_int_Refina = 0) And (r_int_Judici = 0) And (r_int_Castig = 0) Then
               r_int_SitCre = 1
            End If
             
            If (r_int_Refina = 1) Then
               r_int_SitCre = 4
            ElseIf (r_int_Judici = 1) Then
               r_int_SitCre = 6
            ElseIf (r_int_Castig = 1) Then
               r_int_SitCre = 3
            End If
   
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & " SELECT HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF,  "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "        DECODE(SUBSTR(HIPCUO_NUMOPE,1,3),'003', HIPCUO_COMCOF,0) AS HIPCUO_COMFMV  "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   FROM CRE_HIPCUO  "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "  WHERE HIPCUO_NUMOPE = '" & Trim(r_rst_Princi!HIPCUO_NUMOPE) & "' "
            If r_rst_Princi!HIPMAE_CODPRD = "003" Then
               modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_TIPCRO = 5 "
            Else
               modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_TIPCRO = 3 "
            End If
            modprc_g_str_CadEje = modprc_g_str_CadEje & "    AND HIPCUO_NUMCUO = " & CStr(r_rst_Princi!HIPCUO_NUMCUO) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Cuotas, 3) Then
               r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
               Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
               r_rst_Cuotas.Close
               Set r_rst_Cuotas = Nothing
               Exit Sub
            End If
           
            r_bol_EstPrd = False
            If (r_rst_Cuotas!HIPCUO_CAPITA + r_rst_Cuotas!HIPCUO_INTERE + r_rst_Cuotas!HIPCUO_COMCOF + r_rst_Cuotas!HIPCUO_COMFMV) > 0 Then
                ReDim r_arr_Matriz(0)
                
                For r_int_Contad = 1 To UBound(r_arr_MatDet)
                    If r_bol_EstPrd = True Then
                       If r_arr_MatDet(r_int_Contad).MatDet_CodPrd <> r_rst_Princi!HIPMAE_CODPRD Then
                          Exit For
                       End If
                    End If
                    
                    If r_arr_MatDet(r_int_Contad).MatDet_SitCre = r_int_SitCre Then
                       If r_arr_MatDet(r_int_Contad).MatDet_CodPrd = r_rst_Princi!HIPMAE_CODPRD Then
                          r_bol_EstPrd = True
                          r_bol_Estado = False
                          
                          Select Case r_arr_MatDet(r_int_Contad).MatDet_OpeCon
                                 Case "100326": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- CAPITAL PASIVO. MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_CAPITA
                                                r_bol_Estado = True
                                 Case "100327": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT. - INTERES PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_INTERE
                                                r_bol_Estado = True
                                 Case "100328": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- COMISION COF. PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_COMCOF
                                                r_bol_Estado = True
                                 Case "100329": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- COMISION FMV. PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_COMFMV
                                                r_bol_Estado = True
                                 Case "100330": ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1) 'AMORT.- PROV. PAG. PASIVO MES
                                                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_rst_Cuotas!HIPCUO_CAPITA + r_rst_Cuotas!HIPCUO_INTERE + _
                                                                                                   r_rst_Cuotas!HIPCUO_COMCOF + r_rst_Cuotas!HIPCUO_COMFMV
                                                r_bol_Estado = True
                          End Select
                          If r_bol_Estado = True Then
                             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_arr_MatDet(r_int_Contad).CtaPrd_CtaCtb
                             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = UCase(r_arr_MatDet(r_int_Contad).MatDet_DesDet)
                             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = r_arr_MatDet(r_int_Contad).MatDet_DebHab
                          End If
                       End If
                    End If
                Next
                If r_bol_EstPrd = True Then
                   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, p_PerAno, p_PerMes, r_str_Origen, r_int_libro)
                       
                   'Ingresar Cabecera  r_rst_Princi!HIPMAE_MONEDA
                   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, Format(1, "000"), _
                                                 r_dbl_TipCam_Dol, "O", Mid("CUOTA MES PASIVO COFIDE FMV" + " - " + Trim(r_rst_Princi!HIPCUO_NUMOPE) & " - " & _
                                                 r_rst_Princi!HIPCUO_NUMCUO & " - " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), p_FecFin, "1")
                   r_int_NumIte = 1
                   For r_int_AuxCon = 1 To UBound(r_arr_Matriz)
                       r_dbl_ImpSol = 0
                       r_dbl_ImpDol = 0
                       If r_arr_Matriz(r_int_AuxCon).Matriz_import > 0 Then
                          If r_rst_Princi!HIPMAE_MONEDA = 1 Then
                             r_dbl_ImpSol = r_arr_Matriz(r_int_AuxCon).Matriz_import
                             r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_AuxCon).Matriz_import / r_dbl_TipCam_Dol, "#######0.00"))
                          ElseIf r_rst_Princi!HIPMAE_MONEDA = 2 Then
                             r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_AuxCon).Matriz_import * r_dbl_TipCam_Dol, "########0.00"))
                             r_dbl_ImpDol = r_arr_Matriz(r_int_AuxCon).Matriz_import
                          End If
                           
                          Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, r_int_NumIte, r_arr_Matriz(r_int_AuxCon).Matriz_CtaCtb, _
                                                               p_FecFin, Mid(Trim(r_rst_Princi!HIPCUO_NUMOPE) & " - " & r_rst_Princi!HIPCUO_NUMCUO & " - " + _
                                                               UCase(r_arr_Matriz(r_int_AuxCon).Matriz_DesNot) & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), 1, 60), _
                                                               r_arr_Matriz(r_int_AuxCon).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(p_FecFin))
           
                          r_int_NumIte = r_int_NumIte + 1
                       Else
                          r_int_NumIte = r_int_NumIte
                       End If
                   Next
                End If
            End If
         End If
         'Leyendo siguiente registro
         r_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Loop
                                       
      
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub

Private Sub modprc_fs_Inserta_DetAsi(p_LogPro() As modprc_g_tpo_LogPro, ByVal p_Origen As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_NumLib As Integer, p_NumAsi As Integer, ByVal p_NumIte As Integer, _
                                     ByVal p_CtaCtb As String, ByVal p_FecCtb As String, ByVal p_Glosa As String, ByVal p_FlagDH As String, ByVal p_ImpSol As Double, ByVal p_ImpDol As Double)
Dim r_rst_Grabar     As ADODB.Recordset
                              
   modprc_g_str_CadEje = "INSERT INTO CNTBL_ASIENTO_DET ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ORIGEN, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ANO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "MES, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_LIBRO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_ASIENTO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ITEM, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "CNTA_CTBL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FECHA_CNTBL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DET_GLOSA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FLAG_DEBHAB, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "IMP_MOVSOL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "IMP_MOVDOL) "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "VALUES ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_Origen & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerAno) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerMes) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumLib) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumAsi) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumIte) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CtaCtb & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "to_date ('" & p_FecCtb & "', 'DD/MM/YYYY'), "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_Glosa & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_FlagDH & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_ImpSol) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_ImpDol) & ") "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CNTBL_ASIENTO_DET.")
   End If
End Sub

Public Sub modprc_fs_Inserta_DetAsi_PagVar(p_LogPro() As modprc_g_tpo_LogPro, ByVal p_Origen As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_NumLib As Integer, p_NumAsi As Integer, ByVal p_NumIte As Integer, _
                                            ByVal p_CtaCtb As String, ByVal p_FecCtb As String, ByVal p_Glosa As String, ByVal p_FlagDH As String, ByVal p_ImpSol As Double, ByVal p_ImpDol As Double, ByVal p_DifPag, ByVal p_FecAux As String)
Dim r_rst_Grabar     As ADODB.Recordset
   
   modprc_g_str_CadEje = "INSERT INTO CNTBL_ASIENTO_DET ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ORIGEN, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ANO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "MES, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_LIBRO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_ASIENTO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ITEM, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "CNTA_CTBL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FECHA_CNTBL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DET_GLOSA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FLAG_DEBHAB, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "IMP_MOVSOL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "IMP_MOVDOL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_DOCREF3, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "CTB_FECAUX) "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "VALUES ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_Origen & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerAno) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerMes) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumLib) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumAsi) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumIte) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CtaCtb & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "to_date ('" & p_FecCtb & "', 'DD/MM/YYYY'), "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_Glosa & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_FlagDH & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_ImpSol) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_ImpDol) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_DifPag & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "to_date ('" & p_FecAux & "', 'DD/MM/YYYY')) "

   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CNTBL_ASIENTO_DET.")
   End If
End Sub

Public Sub modprc_fs_Inserta_CabAsi(p_LogPro() As modprc_g_tpo_LogPro, ByVal p_Origen As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_NumLib As Integer, ByVal p_NumAsi As Integer, _
                                     ByVal p_CodMon As String, ByVal p_TipCam As Double, ByVal p_TipNot As String, ByVal p_Glosa As String, ByVal p_FecCtb As String, ByVal p_FlgEst As String)
Dim r_rst_Grabar     As ADODB.Recordset
                              
   modprc_g_str_CadEje = "INSERT INTO CNTBL_ASIENTO ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ORIGEN, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "ANO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "MES, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_LIBRO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_ASIENTO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "COD_MONEDA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "TASA_CAMBIO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "TIPO_NOTA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "DESC_GLOSA, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FECHA_CNTBL, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FEC_REGISTRO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "COD_USR, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "FLAG_ESTADO, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "TOT_SOLDEB, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "TOT_SOLHAB, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "TOT_DOLDEB, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "TOT_DOLHAB) "
   
   modprc_g_str_CadEje = modprc_g_str_CadEje & "VALUES ("
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_Origen & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerAno) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerMes) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumLib) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_NumAsi) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_CodMon & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & CDbl(p_TipCam) & ", "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_TipNot & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & Trim(p_Glosa) & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "to_date ('" & p_FecCtb & "', 'DD/MM/YYYY'), "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "to_date ('" & Format(date, "dd/mm/yyyy") & "', 'DD/MM/YYYY'), "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'jtacuc', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_FlgEst & "', "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "0, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "0, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "0, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "0) "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CNTBL_ASIENTO.")
   End If
End Sub

Public Function modprc_ff_NumAsi(p_LogPro() As modprc_g_tpo_LogPro, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_Origen As String, ByVal p_LibCon As Integer) As Integer
Dim r_rst_Genera     As ADODB.Recordset
Dim r_int_InsUpd     As Integer
   
   modprc_ff_NumAsi = 0
   r_int_InsUpd = 0
   
   'Obteniendo Nro. de Asiento
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CNTBL_LIBRO_MES "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE ANO = " & CStr(p_PerAno) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MES = " & CStr(p_PerMes) & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND ORIGEN = '" & p_Origen & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND NRO_LIBRO = " & CStr(p_LibCon) & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CNTBL_LIBRO_MES.")
   End If
   
   If Not (modprc_g_rst_Genera.BOF And modprc_g_rst_Genera.EOF) Then
      modprc_g_rst_Genera.MoveFirst
      If IsNull(modprc_g_rst_Genera!NRO_ASIENTO) Then
         modprc_ff_NumAsi = 0
      Else
         modprc_ff_NumAsi = modprc_g_rst_Genera!NRO_ASIENTO
      End If
      r_int_InsUpd = 1
   End If
   
   modprc_g_rst_Genera.Close
   Set modprc_g_rst_Genera = Nothing
         
   modprc_ff_NumAsi = modprc_ff_NumAsi + 1
         
   'Actualizando Folio
   If r_int_InsUpd = 0 Then
      'Insert
      modprc_g_str_CadEje = "INSERT INTO CNTBL_LIBRO_MES ("
      modprc_g_str_CadEje = modprc_g_str_CadEje & "ANO, "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "MES, "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "ORIGEN, "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_LIBRO, "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_ASIENTO) "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "VALUES ("
      modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerAno) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_PerMes) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & p_Origen & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(p_LibCon) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CStr(modprc_ff_NumAsi) & ") "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 2) Then
         p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al insertar Tabla CNTBL_LIBRO_MES.")
      End If
      
   Else
      'Update
      modprc_g_str_CadEje = "UPDATE CNTBL_LIBRO_MES SET "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_ASIENTO = " & CStr(modprc_ff_NumAsi) & " "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "ANO = " & CStr(p_PerAno) & " AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "MES = " & CStr(p_PerMes) & " AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "ORIGEN = '" & p_Origen & "' AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "NRO_LIBRO = " & CStr(p_LibCon) & " "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 2) Then
         p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CNTBL_LIBRO_MES.")
      End If
   End If
End Function

Public Function ff_ConGas(ByVal p_FecIni As String, ByVal p_FecFin As String) As Integer
   ff_ConGas = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT NVL(COUNT(*),0) AS TOTREG FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1101 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_CTBFLG = 0 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      ff_ConGas = g_rst_Listas!TOTREG
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function ff_ConAmo(ByVal p_FecIni As String, ByVal p_FecFin As String) As Integer
   ff_ConAmo = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT NVL(COUNT(*),0) AS TOTREG FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1102 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_CTBFLG = 0 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      ff_ConAmo = g_rst_Listas!TOTREG
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function ff_ConDes(ByVal p_FecIni As String, ByVal p_FecFin As String) As Integer
   ff_ConDes = 0
      
   g_str_Parame = "   "
   g_str_Parame = g_str_Parame & "SELECT NVL(COUNT(*),0) AS TOTREG FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1103 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_CTBFLG = 0 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      ff_ConDes = g_rst_Listas!TOTREG
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function ff_ConAho(ByVal p_FecIni As String, ByVal p_FecFin As String) As Integer
   ff_ConAho = 0
      
   g_str_Parame = "   "
   g_str_Parame = g_str_Parame & "SELECT NVL(COUNT(*),0) AS TOTREG FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1105 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_CTBFLG = 0 "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(p_FecIni), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(p_FecFin), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      ff_ConAho = g_rst_Listas!TOTREG
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub modprc_ctbp2001(ByVal p_CodEmp As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_PerFin As String)
'Código Proceso   :  CTBP2001
'Descripción      :  Contabilización de Pago de Gastos de Cierre (Asientos en GesCtb)
'Resumen          :  Contabilización de Pago de Gastos de Cierre (Asientos en GesCtb)
'F. Creación      :  16-09-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg           As Long
Dim r_arr_LogPro()         As modprc_g_tpo_LogPro
Dim r_arr_CtaBco()         As modprc_g_tpo_MatCtb
Dim r_arr_TCaSun()         As modprc_g_tpo_MatCtb
Dim r_arr_MatCtb()         As modprc_g_tpo_MatCtb
Dim r_str_FecCtb           As String
Dim r_str_FecPro           As String
Dim r_rst_Genera           As ADODB.Recordset
Dim r_dbl_TCaSBS_Vta       As Double
Dim r_dbl_TCaSBS_Com       As Double
Dim r_dbl_TipCam_Sun       As Double
Dim r_int_MonPag           As Integer
Dim r_dbl_TotDep           As Double
Dim r_dbl_ITFDep           As Double
Dim r_str_SucMov           As String
Dim r_str_NumSol           As String
Dim r_str_CodBan           As String
Dim r_str_NumCta           As String
Dim r_int_PosIni           As Integer
Dim r_int_PosFin           As Integer
Dim r_int_Contad           As Integer
Dim r_lng_NumAsi           As Long
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP2001"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Fecha de Contabilización
   If date > CDate(gf_FormatoFecha(p_PerFin)) Then
      r_str_FecCtb = gf_FormatoFecha(p_PerFin)
   Else
      r_str_FecCtb = Format(date, "dd/mm/yyyy")
   End If
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Cargando Tipo de Cambio SUNAT (Sino existe Tipo de Cambio para algún día Termina el Proceso)
   If modprc_gf_Carga_TipCambioSunat(r_arr_TCaSun, 1101, p_FecIni, p_FecFin, r_arr_LogPro) = 1 Then
      Exit Sub
   End If
   
   'Obteniendo Tipo de Cambio SBS para Fecha Contable
   r_dbl_TCaSBS_Vta = modprc_gf_TipoCambio(2, 1, 2, Format(CDate(r_str_FecCtb), "yyyymmdd"))
   r_dbl_TCaSBS_Com = modprc_gf_TipoCambio(2, 2, 2, Format(CDate(r_str_FecCtb), "yyyymmdd"))
   
   'Cargando Cuentas Contables para todas las Cuentas Bancarias
   Call modprc_gs_Carga_CuentaCtbBanco(r_arr_CtaBco)
   
   'Cargando Matrices para Gastos de Cierre
   Call modprc_gs_Carga_MatCtb(r_arr_MatCtb, "110001", "000001", r_arr_LogPro)
   
   'Leyendo Cursor Principal (Créditos Hipotecarios)
   r_lng_NumReg = 0
   
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT * FROM OPE_CAJMOV "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_TIPMOV = 1101 "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV >= " & p_FecIni & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV <= " & p_FecFin & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         r_int_MonPag = modprc_g_rst_Princi!CAJMOV_MONPAG
         r_dbl_TotDep = modprc_g_rst_Princi!CAJMOV_IMPTOT
         r_dbl_ITFDep = modprc_g_rst_Princi!CAJMOV_ITFIMP
         r_str_SucMov = modprc_g_rst_Princi!CAJMOV_SUCMOV
         r_str_NumSol = Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE)
         r_str_CodBan = Trim(modprc_g_rst_Princi!CAJMOV_CODBAN)
         r_str_NumCta = Trim(modprc_g_rst_Princi!CAJMOV_NUMCTA)
         
         'Obteniendo Tipo de Cambio SUNAT
         r_dbl_TipCam_Sun = moddat_gf_ObtieneTipCamDia(3, 2, CStr(modprc_g_rst_Princi!CAJMOV_FECDEP), 2)
         
         'Buscando Detalle de Gastos de Cierre (TRA_GASADM)
         modprc_g_str_CadEje = "SELECT * FROM TRA_GASADM WHERE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "GASADM_NUMSOL = '" & r_str_NumSol & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla TRA_GASADM - Solicitud " & r_str_NumSol & ".")
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
            Exit Sub
         End If
         
         If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
            r_rst_Genera.MoveFirst
            r_int_PosIni = 0
            r_int_PosFin = 0
            
            'Para Ubicar Rango de Ubicación de Matriz x Moneda
            For r_int_Contad = 1 To UBound(r_arr_MatCtb)
               If r_int_MonPag = r_arr_MatCtb(r_int_Contad).MatCtb_TipMon And r_int_PosIni = 0 Then
                  r_int_PosIni = r_int_Contad
               ElseIf r_int_MonPag <> r_arr_MatCtb(r_int_Contad).MatCtb_TipMon And r_int_PosIni > 0 And r_int_PosFin = 0 Then
                  r_int_PosFin = r_int_Contad
               End If
            Next r_int_Contad
            
            For r_int_Contad = r_int_PosIni To r_int_PosFin
               Select Case r_arr_MatCtb(r_int_Contad).MatCtb_TipCon
                  Case 3      'Cuenta de Bancos
                     r_arr_MatCtb(r_int_Contad).MatCtb_CtaCtb = modprc_gf_BuscaArreglo_CuentaCtbBanco(r_arr_CtaBco, r_str_CodBan, r_str_NumCta, r_int_MonPag, r_arr_MatCtb(r_int_Contad).MatCtb_ConCtb, "000001")
               End Select
            Next r_int_Contad
            
            'Obteniendo Nro. de Asiento
            'r_lng_NumAsi = modctb_gf_Genera_NumAsi("000001", r_str_SucMov, p_PerAno, p_PerMes, r_arr_MatCtb(r_int_PosIni).MatCtb_CodLib)
            
            Do While Not r_rst_Genera.EOF
               'Insertar en CNTBL_ASIENTO_DET
               'Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 1, r_str_CtaDeb, p_FecFin, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DEVENG. " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "D", r_dbl_Lin_ImpDeb_Sol, r_dbl_Lin_ImpDeb_Dol)
               'Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, 2, r_str_CtaHab, p_FecFin, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DEVENG. " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), "H", r_dbl_Lin_ImpHab_Sol, r_dbl_Lin_ImpHab_Dol)
               r_rst_Genera.MoveNext
            Loop
            
            'Insertar en CNTBL_ASIENTO
            'Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, p_PerAno, p_PerMes, r_int_libro, r_int_NumAsi, "001", r_dbl_TipCam_Dol, r_str_TipNot, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DEVENG. " & Format(p_PerMes, "00") & "-" & CStr(p_PerAno), p_FecFin, "1")
         End If
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
            
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
      Loop
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, p_PerMes, p_PerAno, p_FecIni, p_FecFin)
End Sub

Public Sub modprc_ctbp4001(ByVal p_CodPro As String, ByVal p_FInEje As String, ByVal p_HInEje As String, ByVal p_CodEmp As String, ByVal p_CodSuc As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_TotReg As Long, Optional p_BarPro As SSPanel)
'Código Proceso   :  CTBP4001
'Descripción      :  Cuotas de Créditos Hipotecarios
'Resumen          :  Anexos 7, 16 y 16B
'F. Creación      :  24-06-2011
'U. Creación      :  Daniel F. Seclén Reyes
'F. Actualización :
'U. Actualización :

Dim r_dbl_Evalua(223)   As Double
Dim r_str_Fechas(26)    As String
Dim r_str_FecPro        As String
Dim r_str_FecRpt        As String
Dim l_int_mes           As Integer
Dim l_int_ano           As Integer
Dim l_str_fec           As String
Dim l_str_aux           As String
Dim l_dat_fec           As Date
Dim l_dat_aux           As Date
Dim l_int_con           As Integer
Dim r_int_ConAux        As Integer
Dim r_int_Contad        As Integer
Dim r_int_ConTem        As Integer
Dim r_int_AuxTem        As Integer
Dim r_lng_NumReg        As Long
Dim r_lng_NumErr        As Long
Dim r_str_FeFiEj        As String
Dim r_str_HoFiEj        As String
Dim r_dbl_TipCam        As Double

   Erase r_str_Fechas
   Erase r_dbl_Evalua

   r_str_FecRpt = "01/" & Format(p_PerMes, "00") & "/" & p_PerAno

   l_int_mes = p_PerMes
   l_int_ano = p_PerAno
   l_str_fec = "01/" & l_int_mes & "/" & l_int_ano
   l_dat_fec = modsec_gf_Fin_Del_Mes(CDate(l_str_fec))
   l_str_aux = modsec_gf_Fin_Del_Mes(CDate(l_str_fec))
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, modsec_gf_GenFec(l_dat_fec), 2)
   r_lng_NumReg = -1

    r_str_Fechas(0) = CDate(l_dat_fec + 1)
    r_str_Fechas(1) = CDate(l_dat_fec + 7)
    r_str_Fechas(2) = CDate(l_dat_fec + 8)
    r_str_Fechas(3) = CDate(l_dat_fec + 15)
    r_str_Fechas(4) = CDate(l_dat_fec + 16)
    l_str_fec = l_dat_fec + 1
    r_str_Fechas(5) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4))))
    r_str_Fechas(6) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4)))) + 1
    l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4)))) + 1
    r_str_Fechas(7) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
    r_str_Fechas(8) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    r_str_Fechas(9) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
    r_str_Fechas(10) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    r_str_Fechas(11) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
    r_str_Fechas(12) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
    r_str_Fechas(13) = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    r_str_Fechas(14) = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    l_dat_fec = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
       r_str_Fechas(15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    Else
       r_str_Fechas(15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    End If
   
    If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
       r_str_Fechas(16) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    Else
       r_str_Fechas(16) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) + 1
    End If
    l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    r_str_Fechas(17) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    r_str_Fechas(18) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    r_str_Fechas(19) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    r_str_Fechas(20) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    r_str_Fechas(21) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    r_str_Fechas(22) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   
    For l_int_con = 1 To 5
       l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    Next l_int_con
   
    If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
       r_str_Fechas(23) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    Else
       r_str_Fechas(23) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    End If
    If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
       r_str_Fechas(24) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    Else
       r_str_Fechas(24) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) + 1
    End If
   
    For l_int_con = 1 To 10
       l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    Next l_int_con
   
    r_str_Fechas(25) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
    r_str_Fechas(26) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
         
    'For l_int_con = 1 To 20
    '   l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
    'Next l_int_con
   
    '.Cells(3, 15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   
   '**********************************************************************************************************************************************************
   
   'CUENTAS POR PAGAR
   For r_int_Contad = 0 To 26 Step 2
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_TIPCRO = 1 AND "
      If r_int_Contad = 26 Then
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      modprc_g_str_CadEje = modprc_g_str_CadEje & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Leer Tabla CRE_HIPMAE, CRE_HIPCUO.")
      End If
      
      If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
         modprc_g_rst_Princi.MoveFirst
         Do While Not modprc_g_rst_Princi.EOF
         
            If Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "001" Then
               r_dbl_Evalua((r_int_Contad / 2) + 2) = r_dbl_Evalua((r_int_Contad / 2) + 2) + modprc_g_rst_Princi!CAPITAL
              
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "002" Then
               r_dbl_Evalua((r_int_Contad / 2) + 18) = r_dbl_Evalua((r_int_Contad / 2) + 18) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "003" Then
               r_dbl_Evalua((r_int_Contad / 2) + 34) = r_dbl_Evalua((r_int_Contad / 2) + 34) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 50) = r_dbl_Evalua((r_int_Contad / 2) + 50) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "006" Then
               r_dbl_Evalua((r_int_Contad / 2) + 66) = r_dbl_Evalua((r_int_Contad / 2) + 66) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "007" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "013" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "014" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "015" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "016" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "017" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "018" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "009" Then
               r_dbl_Evalua((r_int_Contad / 2) + 98) = r_dbl_Evalua((r_int_Contad / 2) + 98) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "010" Then
               r_dbl_Evalua((r_int_Contad / 2) + 114) = r_dbl_Evalua((r_int_Contad / 2) + 114) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "011" Then
               r_dbl_Evalua((r_int_Contad / 2) + 130) = r_dbl_Evalua((r_int_Contad / 2) + 130) + modprc_g_rst_Princi!CAPITAL
               
            End If
            
            p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / p_TotReg * 100, "##0.00"))
            r_lng_NumReg = r_lng_NumReg + 1
            modprc_g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Next
   
   For r_int_Contad = 0 To 26 Step 2
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_TIPCRO = 2 AND "
      If r_int_Contad = 26 Then
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      'modprc_g_str_CadEje = modprc_g_str_CadEje & "(HIPMAE_NUMOPE <> '0040700001' AND HIPMAE_NUMOPE <> '0040700002' AND "
      'modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_NUMOPE <> '0040700003' AND HIPMAE_NUMOPE <> '0040700004') "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Leer Tabla CRE_HIPMAE, CRE_HIPCUO.")
      End If
      
      If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
         modprc_g_rst_Princi.MoveFirst
         Do While Not modprc_g_rst_Princi.EOF
         
            If Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "001" Then
               r_dbl_Evalua((r_int_Contad / 2) + 2) = r_dbl_Evalua((r_int_Contad / 2) + 2) + modprc_g_rst_Princi!CAPITAL
              
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "002" Then
               r_dbl_Evalua((r_int_Contad / 2) + 18) = r_dbl_Evalua((r_int_Contad / 2) + 18) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "003" Then
               r_dbl_Evalua((r_int_Contad / 2) + 34) = r_dbl_Evalua((r_int_Contad / 2) + 34) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 50) = r_dbl_Evalua((r_int_Contad / 2) + 50) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "006" Then
               r_dbl_Evalua((r_int_Contad / 2) + 66) = r_dbl_Evalua((r_int_Contad / 2) + 66) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "007" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "013" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "014" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "015" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "016" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "017" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "018" Then
               r_dbl_Evalua((r_int_Contad / 2) + 82) = r_dbl_Evalua((r_int_Contad / 2) + 82) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "009" Then
               r_dbl_Evalua((r_int_Contad / 2) + 98) = r_dbl_Evalua((r_int_Contad / 2) + 98) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "010" Then
               r_dbl_Evalua((r_int_Contad / 2) + 114) = r_dbl_Evalua((r_int_Contad / 2) + 114) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "011" Then
               r_dbl_Evalua((r_int_Contad / 2) + 130) = r_dbl_Evalua((r_int_Contad / 2) + 130) + modprc_g_rst_Princi!CAPITAL
               
            End If
            
            p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / p_TotReg * 100, "##0.00"))
            r_lng_NumReg = r_lng_NumReg + 1
            modprc_g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
      
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Next
      
   For r_int_Contad = 0 To 26 Step 2
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_TIPCRO = 4 AND "
      If r_int_Contad = 26 Then
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      'modprc_g_str_CadEje = modprc_g_str_CadEje & "(hipmae_numope = '0040700001' or hipmae_numope = '0040700002' or "
      'modprc_g_str_CadEje = modprc_g_str_CadEje & "hipmae_numope = '0040700003' or hipmae_numope = '0040700004') "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Leer Tabla CRE_HIPMAE, CRE_HIPCUO.")
      End If
      
      If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
         modprc_g_rst_Princi.MoveFirst
         Do While Not modprc_g_rst_Princi.EOF
                    
            If Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 50) = r_dbl_Evalua((r_int_Contad / 2) + 50) + modprc_g_rst_Princi!CAPITAL
               p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / p_TotReg * 100, "##0.00"))
               r_lng_NumReg = r_lng_NumReg + 1
            End If
            
            modprc_g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
      
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Next
   
   'CUENTAS POR COBRAR
   For r_int_Contad = 0 To 26 Step 2
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_TIPCRO = 3 AND "
      If r_int_Contad = 26 Then
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      modprc_g_str_CadEje = modprc_g_str_CadEje & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Leer Tabla CRE_HIPMAE, CRE_HIPCUO.")
      End If
      
      If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
         modprc_g_rst_Princi.MoveFirst
         Do While Not modprc_g_rst_Princi.EOF
         
            If Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 162) = r_dbl_Evalua((r_int_Contad / 2) + 162) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "007" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "013" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "014" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "015" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "016" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "017" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "018" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "009" Then
               r_dbl_Evalua((r_int_Contad / 2) + 194) = r_dbl_Evalua((r_int_Contad / 2) + 194) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "010" Then
               r_dbl_Evalua((r_int_Contad / 2) + 210) = r_dbl_Evalua((r_int_Contad / 2) + 210) + modprc_g_rst_Princi!CAPITAL
               
            End If
            
            p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / p_TotReg * 100, "##0.00"))
            r_lng_NumReg = r_lng_NumReg + 1
            modprc_g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Next
   
   For r_int_Contad = 0 To 26 Step 2
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_TIPCRO = 5 AND "
      If r_int_Contad = 26 Then
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      modprc_g_str_CadEje = modprc_g_str_CadEje & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Leer Tabla CRE_HIPMAE, CRE_HIPCUO.")
      End If
      
      If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
         modprc_g_rst_Princi.MoveFirst
         Do While Not modprc_g_rst_Princi.EOF
         
            If Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "003" Then
               r_dbl_Evalua((r_int_Contad / 2) + 146) = r_dbl_Evalua((r_int_Contad / 2) + 146) + modprc_g_rst_Princi!CAPITAL
            End If
            
            p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / p_TotReg * 100, "##0.00"))
            r_lng_NumReg = r_lng_NumReg + 1
            modprc_g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Next
   
   For r_int_Contad = 0 To 26 Step 2
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPMAE_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_SITUAC = 2 AND "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_TIPCRO = 4 AND "
      If r_int_Contad = 26 Then
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      modprc_g_str_CadEje = modprc_g_str_CadEje & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Leer Tabla CRE_HIPMAE, CRE_HIPCUO.")
      End If
      
      If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
         modprc_g_rst_Princi.MoveFirst
         Do While Not modprc_g_rst_Princi.EOF
         
            If Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "004" Then
               r_dbl_Evalua((r_int_Contad / 2) + 162) = r_dbl_Evalua((r_int_Contad / 2) + 162) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "007" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "013" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "014" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "015" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "016" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "017" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "018" Then
               r_dbl_Evalua((r_int_Contad / 2) + 178) = r_dbl_Evalua((r_int_Contad / 2) + 178) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "009" Then
               r_dbl_Evalua((r_int_Contad / 2) + 194) = r_dbl_Evalua((r_int_Contad / 2) + 194) + modprc_g_rst_Princi!CAPITAL
               
            ElseIf Trim(modprc_g_rst_Princi!HIPMAE_CODPRD) = "010" Then
               r_dbl_Evalua((r_int_Contad / 2) + 210) = r_dbl_Evalua((r_int_Contad / 2) + 210) + modprc_g_rst_Princi!CAPITAL
               
            End If
                        
            p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / p_TotReg * 100, "##0.00"))
            r_lng_NumReg = r_lng_NumReg + 1
            modprc_g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
   Next
   
   'Eliminando registros de la tabla RPT_ANEXOS
   g_str_Parame = "DELETE FROM RPT_ANEXOS WHERE "
   g_str_Parame = g_str_Parame & "ANEXOS_PERMES = " & p_PerMes & " AND "
   g_str_Parame = g_str_Parame & "ANEXOS_PERANO = " & p_PerAno & " AND "
   g_str_Parame = g_str_Parame & "ANEXOS_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "ANEXOS_CODSUC = '" & p_CodSuc & "' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_lng_NumErr = r_lng_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al Eliminar registro en RPT_ANEXOS.")
   End If
   
   'Insertando en la tabla RPT_ANEXOS
   r_int_ConTem = 0
   For r_int_Contad = 1 To 9 Step 1
      modprc_g_str_CadEje = "USP_RPT_ANEXOS ("
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & Format(p_CodEmp, "000000") & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & Format(p_CodSuc, "000") & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & 0 & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & 0 & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & modgen_g_str_NombPC & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & modgen_g_str_CodUsu & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(p_PerMes) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(p_PerAno) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(1) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(IIf(r_int_Contad < 5, r_int_Contad, IIf(r_int_Contad < 7, r_int_Contad + 1, r_int_Contad + 2))) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & (r_dbl_TipCam) & ", "
      
      For r_int_ConAux = 0 To 15 Step 1
         modprc_g_str_CadEje = modprc_g_str_CadEje & r_dbl_Evalua(r_int_ConTem) & ", "
         r_int_ConTem = r_int_ConTem + 1
      Next

      modprc_g_str_CadEje = Left(modprc_g_str_CadEje, Len(modprc_g_str_CadEje) - 2) & ") "

      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al ejecutar el Procedimiento USP_RPT_ANEXOS.")
      End If
   Next
   
   r_int_ConTem = 144
   For r_int_Contad = 1 To 5 Step 1
      modprc_g_str_CadEje = "USP_RPT_ANEXOS ("
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & Format(p_CodEmp, "000000") & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & Format(p_CodSuc, "000") & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & 0 & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & 0 & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & modgen_g_str_NombPC & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "'" & modgen_g_str_CodUsu & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(p_PerMes) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(p_PerAno) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(2) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & CInt(IIf(r_int_Contad <= 2, r_int_Contad + 2, IIf(r_int_Contad = 3, 7, r_int_Contad + 5))) & ", "
      modprc_g_str_CadEje = modprc_g_str_CadEje & (r_dbl_TipCam) & ", "
      
      For r_int_ConAux = 0 To 15 Step 1
         modprc_g_str_CadEje = modprc_g_str_CadEje & r_dbl_Evalua(r_int_ConTem) & ", "
         r_int_ConTem = r_int_ConTem + 1
      Next

      modprc_g_str_CadEje = Left(modprc_g_str_CadEje, Len(modprc_g_str_CadEje) - 2) & ") "

      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
         r_lng_NumErr = r_lng_NumErr + 1
         Call modprc_gs_GrabaErrorProceso(p_CodPro, p_FInEje, p_HInEje, r_lng_NumErr, "Error al ejecutar el Procedimiento USP_RPT_ANEXOS.")
      End If
   Next
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Fecha y Hora Fin de la Ejecución
   r_str_FeFiEj = Format(Now, "YYYYMMDD")
   r_str_HoFiEj = Format(Now, "HHMMSS")
   
   Call modprc_gs_GrabaCabeceraLogProceso(p_CodPro, p_FInEje, p_HInEje, r_str_FeFiEj, r_str_HoFiEj, r_lng_NumReg, r_lng_NumErr, p_CodEmp, "", 0, 0, 0, p_FInEje, r_str_FeFiEj)
End Sub

Public Sub modprc_gs_Carga_CuentaCtbBanco(p_Arregl() As modprc_g_tpo_MatCtb)
Dim r_str_DirCor     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_str_Parame     As String
   
   ReDim p_Arregl(0)
  
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM CTB_CTABCO "
   r_str_Parame = r_str_Parame & " ORDER BY CTABCO_EMPGRP ASC, CTABCO_CODBAN ASC, CTABCO_NUMCTA ASC, CTABCO_TIPMON ASC, CTABCO_CONCTB ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      
      Do While Not r_rst_Princi.EOF
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).MatCtb_CodBan = Trim(r_rst_Princi!CTABCO_CODBAN)
         p_Arregl(UBound(p_Arregl)).MatCtb_NumCta = Trim(r_rst_Princi!CTABCO_NUMCTA)
         p_Arregl(UBound(p_Arregl)).MatCtb_TipMon = r_rst_Princi!CTABCO_TIPMON
         p_Arregl(UBound(p_Arregl)).MatCtb_ConCtb = Trim(r_rst_Princi!CTABCO_CONCTB)
         p_Arregl(UBound(p_Arregl)).MatCtb_EmpGrp = Trim(r_rst_Princi!CTABCO_EMPGRP)
         p_Arregl(UBound(p_Arregl)).MatCtb_CtaCtb = Trim(r_rst_Princi!CTABCO_CTACTB)
      
         g_rst_Princi.MoveNext
      Loop
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Public Function modprc_gf_BuscaArreglo_CuentaCtbBanco(p_Arregl() As modprc_g_tpo_MatCtb, ByVal p_CodBan As String, ByVal p_NumCta As String, ByVal p_TipMon As Integer, ByVal p_ConCtb As String, ByVal p_EmpGrp As String) As String
Dim r_int_Contad     As Integer
   
   modprc_gf_BuscaArreglo_CuentaCtbBanco = ""
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      If p_Arregl(r_int_Contad).MatCtb_CodBan = p_CodBan And p_Arregl(r_int_Contad).MatCtb_NumCta = p_NumCta And p_Arregl(r_int_Contad).MatCtb_TipMon = p_TipMon And _
         p_Arregl(r_int_Contad).MatCtb_EmpGrp = p_EmpGrp And p_Arregl(r_int_Contad).MatCtb_ConCtb = p_ConCtb Then
         
         modprc_gf_BuscaArreglo_CuentaCtbBanco = p_Arregl(r_int_Contad).MatCtb_CtaCtb
         Exit For
      End If
   Next r_int_Contad
End Function

Public Function modprc_gf_TipoCambio(ByVal p_TipCam As Integer, ByVal p_TipTip As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String) As Double
   'TipCam = 1 - Comercial / 2 - SBS / 3 - Sunat / 4 - BCR
   'TipTip = 1 - Venta / 2 - Compra
   modprc_gf_TipoCambio = 0
   
   'Obteniendo Tipo de Cambio del Dia
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_TIPCAM "
   g_str_Parame = g_str_Parame & " WHERE TIPCAM_CODIGO = " & CStr(p_TipCam) & " "
   g_str_Parame = g_str_Parame & "   AND TIPCAM_TIPMON = " & CStr(p_TipMon) & " "
   g_str_Parame = g_str_Parame & "   AND TIPCAM_FECDIA <= " & p_FecDia & " "
   g_str_Parame = g_str_Parame & " ORDER BY TIPCAM_FECDIA DESC, TIPCAM_HORDIA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
   
      If p_TipTip = 1 Then
         modprc_gf_TipoCambio = g_rst_Genera!TIPCAM_VENTAS
      ElseIf p_TipTip = 2 Then
         modprc_gf_TipoCambio = g_rst_Genera!TIPCAM_COMPRA
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function modprc_gf_Carga_TipCambioSunat(p_Arregl() As modprc_g_tpo_MatCtb, ByVal p_TipMov As Integer, ByVal p_FecIni As String, ByVal p_FecFin As String, _
                                               p_LogPro() As modprc_g_tpo_LogPro) As Integer
Dim r_rst_Genera     As ADODB.Recordset
Dim r_dbl_TCaSun     As Integer

   modprc_gf_Carga_TipCambioSunat = 0
   ReDim p_Arregl(0)
   
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT DISTINCT CAJMOV_FECDEP FROM OPE_CAJMOV "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAJMOV_TIPMOV = " & p_TipMov & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV >= " & p_FecIni & " "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CAJMOV_FECMOV <= " & p_FecFin & " "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tabla OPE_CAJMOV.")
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         r_dbl_TCaSun = moddat_gf_ObtieneTipCamDia(3, 2, CStr(r_rst_Genera!CAJMOV_FECDEP), 2)
         
         If r_dbl_TCaSun = 0 Then
            p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "No se encontró Tipo de Cambio Sunat (Compra) para el día " & gf_FormatoFecha(CStr(r_rst_Genera!CAJMOV_FECDEP)) & ".")
            modprc_gf_Carga_TipCambioSunat = 1
         Else
            p_Arregl(UBound(p_Arregl)).MatCtb_FTCSun = CStr(r_rst_Genera!CAJMOV_FECDEP)
            p_Arregl(UBound(p_Arregl)).MatCtb_ComSun = r_dbl_TCaSun
         End If
      
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Sub modprc_gs_Carga_MatCtb(p_Arregl() As modprc_g_tpo_MatCtb, ByVal p_TipMat As String, ByVal p_CodEmp As String, p_LogPro() As modprc_g_tpo_LogPro)
Dim r_rst_Genera     As ADODB.Recordset

   ReDim p_Arregl(0)
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "SELECT MATCAB_CODMAT, MATCAB_DESCRI, MATCAB_TIPMON, MATCAB_CODLIB, MATDET_NUMITE, MATDET_DESCRI, MATDET_TIPCON, MATDET_CONCTB, MATDET_CTACTB, "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "       MATDET_TIPTCA, MATDET_FLGDHB, MATDET_CONOPE "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "  FROM CTB_MATCAB A, CTB_MATDET B "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MATCAB_CODMAT = MATDET_CODMAT "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MATCAB_TIPMAT = '" & p_TipMat & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MATCAB_CODEMP = '" & p_CodEmp & "' "
   modprc_g_str_CadEje = modprc_g_str_CadEje & " ORDER BY MATCAB_CODMAT ASC, MATDET_NUMITE ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Leer Tablas CTB_MATCAB, CTB_MATDET.")
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).MatCtb_CodMat = Trim(r_rst_Genera!MATCAB_CODMAT)
         p_Arregl(UBound(p_Arregl)).MatCtb_DesCab = Trim(r_rst_Genera!MATCAB_DESCRI)
         p_Arregl(UBound(p_Arregl)).MatCtb_TipMon = r_rst_Genera!MATCAB_TIPMON
         p_Arregl(UBound(p_Arregl)).MatCtb_CodLib = r_rst_Genera!MATCAB_CODLIB
         p_Arregl(UBound(p_Arregl)).MatCtb_NumIte = r_rst_Genera!MATDET_NUMITE
         p_Arregl(UBound(p_Arregl)).MatCtb_DesDet = Trim(r_rst_Genera!MATDET_DESCRI)
         p_Arregl(UBound(p_Arregl)).MatCtb_TipCon = r_rst_Genera!MATDET_TIPCON
         p_Arregl(UBound(p_Arregl)).MatCtb_ConCtb = Trim(r_rst_Genera!MATDET_CONCTB)
         p_Arregl(UBound(p_Arregl)).MatCtb_CtaCtb = Trim(r_rst_Genera!MATDET_CTACTB)
         p_Arregl(UBound(p_Arregl)).MatCtb_TipTCa = r_rst_Genera!MATDET_TIPTCA
         p_Arregl(UBound(p_Arregl)).MatCtb_FlgDHb = r_rst_Genera!MATDET_FLGDHB
         p_Arregl(UBound(p_Arregl)).MatCtb_ConOpe = Trim(r_rst_Genera!MATDET_CONOPE)
         
         r_rst_Genera.MoveNext
      Loop
   End If

   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modprc_AsiCtb_Cab(ByVal p_CodEmp As String, ByVal p_CodSuc As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_LibCon As Integer, ByVal p_NumAsi As Long, ByVal p_FecCtb As String, ByVal p_MonCtb As Integer, _
                             ByVal p_TipCam As Double, ByVal p_GloCab As String, p_CuoAsi As Integer, ByVal p_TotDeb_MN As Double, ByVal p_TotHab_MN As Double, ByVal p_TotDeb_ME As Double, ByVal p_TotHab_ME As Double, _
                             ByVal p_TipReg As Integer, ByVal p_CodPro As String, ByVal p_FecPro As String, ByVal p_HorPro As String, ByVal p_FlgGrb As Integer, p_LogPro() As modprc_g_tpo_LogPro)
Dim r_rst_Grabar     As ADODB.Recordset
   
   g_str_Parame = "USP_CTB_ASICAB ("
   
   'Datos Principales
   g_str_Parame = g_str_Parame & "'" & p_CodEmp & "', "
   g_str_Parame = g_str_Parame & "'" & p_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
   g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_LibCon) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumAsi) & ", "
   g_str_Parame = g_str_Parame & p_FecCtb & ", "
   g_str_Parame = g_str_Parame & CStr(p_MonCtb) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TipCam) & ", "
   g_str_Parame = g_str_Parame & "'" & p_GloCab & "', "
   g_str_Parame = g_str_Parame & CStr(p_CuoAsi) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TotDeb_MN) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TotHab_MN) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TotDeb_ME) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TotHab_ME) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TipReg) & ", "
   g_str_Parame = g_str_Parame & "'" & p_CodPro & "', "
   g_str_Parame = g_str_Parame & p_FecPro & ", "
   g_str_Parame = g_str_Parame & p_HorPro & ", "

   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(p_FlgGrb) & ")"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Grabar en CTB_ASICAB.")
   End If
End Sub

Public Sub modprc_AsiCtb_Det(ByVal p_CodEmp As String, ByVal p_CodSuc As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_LibCon As Integer, ByVal p_NumAsi As Long, ByVal p_NumIte As Integer, ByVal p_CtaCtb As String, _
                             ByVal p_FecCtb As String, ByVal p_DebHab As Integer, ByVal p_GloDet As String, ByVal p_Import_MN As Double, ByVal p_Import_ME As Double, p_LogPro() As modprc_g_tpo_LogPro)
Dim r_rst_Grabar     As ADODB.Recordset
                             
   g_str_Parame = "USP_CTB_ASIDET ("
   
   'Datos Principales
   g_str_Parame = g_str_Parame & "'" & p_CodEmp & "', "
   g_str_Parame = g_str_Parame & "'" & p_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
   g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_LibCon) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumAsi) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
   
   'Datos de Linea
   g_str_Parame = g_str_Parame & "'" & p_CtaCtb & "',"
   g_str_Parame = g_str_Parame & p_FecCtb & ", "
   g_str_Parame = g_str_Parame & CStr(p_DebHab) & ","
   g_str_Parame = g_str_Parame & CStr(p_Import_MN) & ", "
   g_str_Parame = g_str_Parame & CStr(p_Import_ME) & ", "
   g_str_Parame = g_str_Parame & "'" & p_GloDet & "',"
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Grabar en CTB_ASIDET.")
   End If
End Sub

Public Sub modprc_AsiCtb_Doc(ByVal p_CodEmp As String, ByVal p_CodSuc As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_LibCon As Integer, ByVal p_NumAsi As Long, ByVal p_NumIte As Integer, ByVal p_DocTip As Integer, ByVal p_DocSer As String, _
                             ByVal p_DocNum As String, ByVal p_DocFec As String, ByVal p_MovSuc As String, ByVal p_MovNum As String, ByVal p_MovFec As String, ByVal p_RefTip As Integer, ByVal p_RefOpe As String, ByVal p_RefSol As String, ByVal p_IdeTip As Integer, _
                             ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_BcoTip As Integer, ByVal p_BcoCod As String, ByVal p_BcoCta As String, ByVal p_BcoNum As String, ByVal p_BcoFec As String, ByVal p_OrgMon As Integer, ByVal p_OrgMto As Double, p_LogPro() As modprc_g_tpo_LogPro)
Dim r_rst_Grabar     As ADODB.Recordset
   
   g_str_Parame = "USP_CTB_ASIDOC ("
   
   'Datos Principales
   g_str_Parame = g_str_Parame & "'" & p_CodEmp & "', "
   g_str_Parame = g_str_Parame & "'" & p_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
   g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_LibCon) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumAsi) & ", "
   g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
   
   'Documento de Referencia
   g_str_Parame = g_str_Parame & CStr(p_DocTip) & ", "
   g_str_Parame = g_str_Parame & "'" & p_DocSer & "', "
   g_str_Parame = g_str_Parame & "'" & p_DocNum & "', "
   g_str_Parame = g_str_Parame & Format(CDate(p_DocFec), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & "'" & p_MovSuc & "', "
   g_str_Parame = g_str_Parame & p_MovNum & ", "
   g_str_Parame = g_str_Parame & Format(CDate(p_MovFec), "yyyymmdd") & ", "
   
   'Tipo de Operación
   g_str_Parame = g_str_Parame & CStr(p_RefTip) & ", "
   g_str_Parame = g_str_Parame & "'" & p_RefOpe & "', "
   g_str_Parame = g_str_Parame & "'" & p_RefSol & "', "
   
   'Tipo de Persona
   g_str_Parame = g_str_Parame & CStr(p_IdeTip) & ", "
   g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
   g_str_Parame = g_str_Parame & "'" & p_NumDoc & "', "
   
   'Movimiento Bancario
   g_str_Parame = g_str_Parame & CStr(p_BcoTip) & ", "
   g_str_Parame = g_str_Parame & "'" & p_BcoCod & "', "
   g_str_Parame = g_str_Parame & "'" & p_BcoCta & "', "
   g_str_Parame = g_str_Parame & "'" & p_BcoNum & "', "
   g_str_Parame = g_str_Parame & Format(CDate(p_BcoFec), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & CStr(p_OrgMon) & ", "
   g_str_Parame = g_str_Parame & CStr(p_OrgMto) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
      p_LogPro(1).LogPro_NumErr = p_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(p_LogPro(1).LogPro_CodPro, p_LogPro(1).LogPro_FInEje, p_LogPro(1).LogPro_HInEje, p_LogPro(1).LogPro_NumErr, "Error al Grabar en CTB_ASIDOC.")
   End If
End Sub

Public Sub modprc_ctbp9006(ByVal p_CodEmp As String)
'Código Proceso   :  CTBP1006
'Descripción      :  Traslado de Ingresos Diferidos (Asientos en Edpymebank)
'Resumen          :  Contabilización de Ingresos Diferidos (Asientos en Edpymebank)
'F. Creación      :  17-07-2009
'U. Creación      :  Miguel Angel Ikehara Punk
'F. Actualización :
'U. Actualización :

Dim r_lng_NumReg           As Long
Dim r_lng_TotReg           As Long
Dim r_arr_LogPro()         As modprc_g_tpo_LogPro
Dim r_str_FecPro           As String
Dim r_dbl_TipCam           As Double
Dim r_rst_Genera           As ADODB.Recordset
Dim r_rst_Grabar           As ADODB.Recordset
Dim r_dbl_AcuDvg           As Double
Dim r_dbl_AcuDif           As Double
Dim r_dbl_IntPag           As Double
Dim r_dbl_Mes_IntEfe       As Double
Dim r_dbl_Mes_IntDev       As Double
Dim r_dbl_Mes_IntDif       As Double
Dim r_int_LibCon           As Integer
Dim r_str_CtaDeb           As String
Dim r_str_CtaHab           As String
Dim r_str_CtaHb1           As String
Dim r_str_Origen           As String
Dim r_str_TipNot           As String
Dim r_int_NumAsi           As Integer
Dim r_int_TipMon           As Integer
Dim r_dbl_Lin_ImpDeb_Sol   As Double
Dim r_dbl_Lin_ImpHab_Sol   As Double
Dim r_dbl_Lin_ImpHb1_Sol   As Double
Dim r_dbl_Lin_ImpDeb_Dol   As Double
Dim r_dbl_Lin_ImpHab_Dol   As Double
Dim r_dbl_Lin_ImpHb1_Dol   As Double
Dim r_str_CodPrd           As String
Dim r_dbl_Imp_AjuSol       As Double
      
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   'Validar que p_FecIni y p_FecFin pertenezca al Período Activo
   r_arr_LogPro(1).LogPro_CodPro = "CTBP9006"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate("31/10/2009"), "yyyymmdd"), 2)
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Leyendo Cursor Principal
   modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCIE " & _
                         " WHERE HIPCIE_PERMES = 9 AND HIPCIE_PERANO = 0 " & _
                         "   AND HIPCIE_ACUDIF > 0 " & _
                         " ORDER BY HIPCIE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPCUO.")
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         modprc_g_str_CadEje = "SELECT HIPCUO_INTERE FROM CRE_HIPCUO WHERE HIPCUO_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "' AND HIPCUO_TIPCRO = 1 AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCUO_FECVCT >= 20091001 AND HIPCUO_FECVCT <= 20091031"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPCUO_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         r_dbl_IntPag = r_rst_Genera!HIPCUO_INTERE
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
         
         'Para obtener Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = "SELECT HIPMAE_ACUDVG, HIPMAE_ACUDVC, HIPMAE_ACUDIF, HIPMAE_MONEDA, HIPMAE_CODPRD FROM CRE_HIPMAE " & _
                               " WHERE HIPMAE_NUMOPE = '" & Trim(modprc_g_rst_Princi!HIPCIE_NUMOPE) & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Genera, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!HIPCUO_NUMOPE) & " .")
         End If
         
         r_rst_Genera.MoveFirst
         r_dbl_AcuDvg = r_rst_Genera!HIPMAE_ACUDVG
         r_int_TipMon = r_rst_Genera!HIPMAE_MONEDA
         r_str_CodPrd = Trim(r_rst_Genera!HIPMAE_CODPRD)
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
            
         r_dbl_Mes_IntEfe = 0
         r_dbl_Mes_IntDev = 0
      
         'Distribuyendo Interés Diferido
         If r_dbl_IntPag >= r_dbl_AcuDvg Then
            r_dbl_Mes_IntDev = r_dbl_AcuDvg
            r_dbl_Mes_IntEfe = r_dbl_IntPag - r_dbl_Mes_IntDev
            r_dbl_AcuDvg = 0
         Else
            r_dbl_Mes_IntDev = r_dbl_IntPag
            r_dbl_AcuDvg = r_dbl_AcuDvg - r_dbl_Mes_IntDev
         End If
         
         'r_dbl_AcuDif = r_dbl_AcuDif - modprc_g_rst_Princi!HIPCUO_INTPAG
         
         'Actualizando en CRE_HIPMAE Saldo Acumulado de Devengado Vigente, Devengado Vencido, Interés Diferido
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPMAE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPMAE_ACUDVG = " & CStr(r_dbl_AcuDvg) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPMAE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "'"
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al actualizar Tabla CRE_HIPMAE - Operación Nro.: " & Trim(modprc_g_rst_Princi!CAJMOV_NUMOPE) & " .")
         End If
         
         If r_dbl_IntPag > 0 Then
            'Generar Asiento Contable
            r_str_CtaDeb = ""
            r_str_CtaHab = ""
         
            Select Case r_str_CodPrd
               Case "001"  'Producto CRC-PBP
                  r_str_CtaDeb = "292102010101"
                  r_str_CtaHab = "512401042401"
                  r_str_CtaHb1 = "142804010101"
               
               Case "002"  'Producto miCasita Dolares
                  r_str_CtaDeb = "292102010101"
                  r_str_CtaHab = "512401040601"
                  r_str_CtaHb1 = "142804010101"
               
               Case "003"  'Producto CME
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042501"
                  r_str_CtaHb1 = "141804010101"
                  
               Case "004"  'Producto miHogar
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042301"
                  r_str_CtaHb1 = "141804010101"
               
               Case "006"  'Producto miCasita PBP
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401040601"
                  r_str_CtaHb1 = "141804010101"
               
               Case "007"  'Producto NUEVO miVivienda
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042302"
                  r_str_CtaHb1 = "141804010101"
               
               Case "009"  'Producto Union Andina
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042303"
                  r_str_CtaHb1 = "141804010101"
               
               Case "010"  'Mivivienda Peruanos en el Extranjero
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401042304"
                  r_str_CtaHb1 = "141804010101"
               
               Case "011"  'miCasita Soles
                  r_str_CtaDeb = "291102010101"
                  r_str_CtaHab = "511401040601"
                  r_str_CtaHb1 = "141804010101"

            End Select
         
            Select Case r_int_TipMon
               Case 1
                  r_dbl_Lin_ImpDeb_Sol = r_dbl_IntPag
                  r_dbl_Lin_ImpHab_Sol = r_dbl_Mes_IntEfe
                  r_dbl_Lin_ImpHb1_Sol = r_dbl_Mes_IntDev
                  
                  r_dbl_Lin_ImpDeb_Dol = CDbl(Format(r_dbl_IntPag / r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHab_Dol = CDbl(Format(r_dbl_Mes_IntEfe / r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHb1_Dol = CDbl(Format(r_dbl_Mes_IntDev / r_dbl_TipCam, "#####0.00"))
                  
               Case 2
                  r_dbl_Lin_ImpDeb_Sol = CDbl(Format(r_dbl_IntPag * r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHab_Sol = CDbl(Format(r_dbl_Mes_IntEfe * r_dbl_TipCam, "#####0.00"))
                  r_dbl_Lin_ImpHb1_Sol = CDbl(Format(r_dbl_Mes_IntDev * r_dbl_TipCam, "#####0.00"))
                  
                  r_dbl_Lin_ImpDeb_Dol = r_dbl_IntPag
                  r_dbl_Lin_ImpHab_Dol = r_dbl_Mes_IntEfe
                  r_dbl_Lin_ImpHb1_Dol = r_dbl_Mes_IntDev
            End Select
            
            r_str_Origen = "LM"
            r_int_LibCon = 1
            r_str_TipNot = "O"
            
            'Obteniendo Nro. de Asiento
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, 2009, 10, r_str_Origen, r_int_LibCon)
            
            'Insertar en CNTBL_ASIENTO
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, 2009, 10, r_int_LibCon, r_int_NumAsi, "001", r_dbl_TipCam, r_str_TipNot, modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DIFERIDO: " & Format(10, "00") & "-" & CStr(2009), "31/10/2009", "1")
            
            'Insertar en CNTBL_ASIENTO_DET
            Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, 2009, 10, r_int_LibCon, r_int_NumAsi, 1, r_str_CtaDeb, "31/10/2009", modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DIFERIDO " & Format(10, "00") & "-" & CStr(2009), "D", r_dbl_Lin_ImpDeb_Sol, r_dbl_Lin_ImpDeb_Dol)
            Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, 2009, 10, r_int_LibCon, r_int_NumAsi, 2, r_str_CtaHab, "31/10/2009", modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DIFERIDO " & Format(10, "00") & "-" & CStr(2009), "H", r_dbl_Lin_ImpHab_Sol, r_dbl_Lin_ImpHab_Dol)
            Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, 2009, 10, r_int_LibCon, r_int_NumAsi, 3, r_str_CtaHb1, "31/10/2009", modprc_g_rst_Princi!HIPCIE_NUMOPE & " - DIFERIDO " & Format(10, "00") & "-" & CStr(2009), "H", r_dbl_Lin_ImpHb1_Sol, r_dbl_Lin_ImpHb1_Dol)
            
            'Ajuste por Diferencia de Tipo de Cambio
            If r_int_TipMon = 2 Then
               If r_dbl_Lin_ImpDeb_Sol > (r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol) Then
                  r_dbl_Imp_AjuSol = CDbl(Format(r_dbl_Lin_ImpDeb_Sol - (r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol), "######0.00"))
                  
                  If r_dbl_Imp_AjuSol > 0 Then
                     Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, 2009, 10, r_int_LibCon, r_int_NumAsi, 4, "512804090101", "31/10/2009", "AJUSTE DIF. TIPO CAMBIO", "H", r_dbl_Imp_AjuSol, 0)
                  End If
               ElseIf r_dbl_Lin_ImpDeb_Sol < (r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol) Then
                  r_dbl_Imp_AjuSol = CDbl(Format((r_dbl_Lin_ImpHab_Sol + r_dbl_Lin_ImpHb1_Sol) - r_dbl_Lin_ImpDeb_Sol, "######0.00"))
               
                  If r_dbl_Imp_AjuSol > 0 Then
                     Call modprc_fs_Inserta_DetAsi(r_arr_LogPro, r_str_Origen, 2009, 10, r_int_LibCon, r_int_NumAsi, 4, "412804090101", "31/10/2009", "AJUSTE DIF. TIPO CAMBIO", "D", r_dbl_Imp_AjuSol, 0)
                  End If
               End If
            End If
         End If
         
         'Leyendo siguiente cuota
         modprc_g_rst_Princi.MoveNext
         r_lng_NumReg = r_lng_NumReg + 1
         DoEvents
      Loop
   Else
      r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "No existen registros en tabla OPE_CAJMOV.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, p_CodEmp, "", 0, 10, 2009, Format(CDate("01/10/2009"), "yyyymmdd"), Format(CDate("31/10/2009"), "yyyymmdd"))
End Sub

Public Function modprc_ff_CalculaGastosJudicial(ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_SalPre As Double, ByVal p_MonPre As Integer, ByVal p_TipCam As Double, _
                                                ByVal p_TipGar As Integer, ByVal p_MtoGar As Double, ByVal p_MonGar As Integer) As Double
Dim r_dbl_MtoCal        As Double
Dim r_dbl_MtoGar        As Double
Dim r_dbl_Tempo1        As Double
Dim r_rst_Genera        As ADODB.Recordset

   modprc_ff_CalculaGastosJudicial = 0
   r_dbl_MtoCal = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD "
   g_str_Parame = g_str_Parame & " WHERE PARPRD_CODPRD = '" & p_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_CODSUB = '" & p_CodSub & "' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_CODGRP = '055' "
   g_str_Parame = g_str_Parame & "   AND PARPRD_CODITE > '000' "
   g_str_Parame = g_str_Parame & " ORDER BY PARPRD_CODPRD, PARPRD_CODSUB, PARPRD_CODGRP, PARPRD_CODITE "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   '* 1. calcula comision por ejecucion de garantias, solo si tiene registrada hipoteca
   r_dbl_Tempo1 = 0
   If p_TipGar = 1 Then
      If p_MonPre = 1 Then
         If p_MonGar = 1 Then
            r_dbl_MtoCal = p_MtoGar
         Else
            r_dbl_MtoCal = p_MtoGar * p_TipCam
         End If
         r_dbl_MtoGar = r_dbl_MtoCal / p_TipCam
      Else
         If p_MonGar = 1 Then
            r_dbl_MtoCal = p_MtoGar / p_TipCam
         Else
            r_dbl_MtoCal = p_MtoGar
         End If
         r_dbl_MtoGar = r_dbl_MtoCal
      End If
      r_dbl_MtoCal = r_dbl_MtoCal * (2 / 3)
      
      r_rst_Genera.MoveFirst
      Do While Not r_rst_Genera.EOF
         If r_dbl_MtoGar < 20001 Then
            If r_rst_Genera!PARPRD_CODITE = "001" Then
               r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
               Exit Do
            End If
         Else
            If r_rst_Genera!PARPRD_CODITE = "002" Then
               r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
               Exit Do
            End If
         End If
         r_rst_Genera.MoveNext
      Loop
      
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_MtoCal * (r_dbl_Tempo1 / 100)), 2)
   End If
   
   '* 2. calcula comision por presentacion de demanda
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "004" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
   
   If p_MonPre = 1 Then
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round(r_dbl_Tempo1, 2)
   Else
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_Tempo1 / p_TipCam), 2)
   End If
   
   '* 3. calcula comision por remate judicial
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "006" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
   
   If p_MonPre = 1 Then
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round(r_dbl_Tempo1, 2)
   Else
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_Tempo1 / p_TipCam), 2)
   End If
   
   '* 4. calcula publicaciones (el peruano)
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "008" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
   
   If p_MonPre = 1 Then
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round(r_dbl_Tempo1, 2)
   Else
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_Tempo1 / p_TipCam), 2)
   End If
   
   '* 5. calcula costo peritos
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "010" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
      
   If p_MonPre = 1 Then
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round(r_dbl_Tempo1, 2)
   Else
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_Tempo1 / p_TipCam), 2)
   End If
      
   '* 6. calcula costo martillero publico
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "011" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_VALMAX
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
   
   modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_MtoCal * (r_dbl_Tempo1 / 100)), 2)
   
   '* 7. calcula costo cedula de notificacion
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "012" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
   
   If p_MonPre = 1 Then
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round(r_dbl_Tempo1, 2)
   Else
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_Tempo1 / p_TipCam), 2)
   End If
   
   '* 8. costo de arbritraje
   r_dbl_Tempo1 = 0
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      If r_rst_Genera!PARPRD_CODITE = "013" Then
         r_dbl_Tempo1 = r_rst_Genera!PARPRD_CANTID
         Exit Do
      End If
      r_rst_Genera.MoveNext
   Loop
   
   If p_MonPre = 1 Then
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round(r_dbl_Tempo1, 2)
   Else
      modprc_ff_CalculaGastosJudicial = modprc_ff_CalculaGastosJudicial + Round((r_dbl_Tempo1 / p_TipCam), 2)
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Sub modprc_ff_CalculaMontosBaseProv(ByRef p_MtoCga As Double, ByRef p_MtoSga As Double, ByVal p_MtoGto As Double, ByVal p_SalCap As Double, ByVal p_SalCon As Double, ByVal p_CodPrd As String, _
                                           ByVal p_MonPre As Double, ByVal p_TipCam As Double, ByVal p_FecDes As String, ByVal p_TipGar As Integer, ByVal p_MtoGar As Double, ByVal p_MonGar As Integer, ByVal p_IntDif As Double)
Dim r_dbl_MtoBas        As Double
Dim r_dbl_MtoSal        As Double
Dim r_dbl_Gastos        As Double
Dim r_dbl_RemGan        As Double
Dim r_dbl_RemPer        As Double
Dim r_dbl_PorPer        As Double
Dim r_dbl_MtoPer        As Double
Dim r_dbl_MtoMen        As Double
Dim r_dbl_FmvCob        As Double
Dim r_dbl_RieCon        As Double
Dim r_dbl_MtoTer        As Double
Dim r_dbl_MtoTot        As Double
Dim r_dbl_Saldo2        As Double

   r_dbl_MtoBas = 0
   r_dbl_MtoSal = 0
   r_dbl_Gastos = 0
   r_dbl_RemGan = 0
   r_dbl_RemPer = 0
   r_dbl_PorPer = 0
   r_dbl_MtoPer = 0
   r_dbl_MtoMen = 0
   r_dbl_FmvCob = 0
   r_dbl_RieCon = 0
   r_dbl_MtoTer = 0
   r_dbl_MtoTot = 0
   r_dbl_Saldo2 = 0
   
   '* Determina cobertura del FMV
   If p_TipGar = 1 Or p_TipGar = 2 Then
      r_dbl_FmvCob = (p_SalCap + p_SalCon - p_IntDif) * (1 / 3)
         
      '* Determina el monto base de las garantias
      If p_MtoGar > 0 Then
         If p_MonPre = 1 Then
            If p_MonGar = 1 Then
               r_dbl_MtoBas = p_MtoGar
            Else
               r_dbl_MtoBas = p_MtoGar * p_TipCam
            End If
         Else
            If p_MonGar = 1 Then
               r_dbl_MtoBas = p_MtoGar / p_TipCam
            Else
               r_dbl_MtoBas = p_MtoGar
            End If
         End If
      Else
         r_dbl_MtoBas = p_SalCap + p_SalCon
      End If
      r_dbl_MtoBas = r_dbl_MtoBas * (2 / 3)
      
   Else
      
      '* determina los montos directamente
      p_MtoSga = p_SalCap + p_SalCon - p_IntDif
      p_MtoCga = 0
      Exit Sub
   End If
   
   '* Determina montos saldo y gastos judiciales
   r_dbl_MtoSal = p_SalCap + p_SalCon - p_IntDif
   r_dbl_Gastos = p_MtoGto
   
   '* Determina ganancia despues del remate
   If r_dbl_MtoSal + r_dbl_Gastos - r_dbl_MtoBas - p_IntDif < 0 Then
      r_dbl_RemGan = r_dbl_MtoBas + p_IntDif - r_dbl_MtoSal - r_dbl_Gastos
   End If
   
   '* Determina perdida despues del remate
   If r_dbl_MtoSal + r_dbl_Gastos - r_dbl_MtoBas - p_IntDif > 0 Then
      r_dbl_RemPer = r_dbl_MtoSal + r_dbl_Gastos - r_dbl_MtoBas - p_IntDif
   End If
   
   '* Determina perdida esperada de la empresa
   If p_CodPrd = "002" Or p_CodPrd = "006" Or p_CodPrd = "011" Then
      '**************  MICASITA SOLES Y DOLARES **************
      If r_dbl_RemGan > 0 Then
         r_dbl_MtoPer = 0
      End If
      If r_dbl_RemPer > 0 Then
         r_dbl_MtoPer = r_dbl_RemPer
      End If
      
   ElseIf p_CodPrd = "001" Or p_CodPrd = "003" Then
      '********************  CRC-PBP Y CME *******************
      If r_dbl_RemGan > 0 Then
         r_dbl_PorPer = 0
         r_dbl_MtoPer = 0
      End If
      If r_dbl_RemPer > 0 Then
         r_dbl_PorPer = r_dbl_RemPer * (1 / 3)
      End If
      
      '* determina el monto menor
      If r_dbl_FmvCob > 0 Then
         If r_dbl_FmvCob > r_dbl_PorPer Then
            r_dbl_MtoMen = r_dbl_PorPer
         Else
            r_dbl_MtoMen = r_dbl_FmvCob
         End If
         r_dbl_MtoPer = r_dbl_RemPer - r_dbl_MtoMen
      Else
         r_dbl_MtoPer = 0
      End If
      
   Else
      '****************  MIVIVIENDA Y MIHOGAR ****************
      If p_FecDes < "20100701" Then
         'INICIO - CALCULO ANTERIOR
         'If r_dbl_RemGan > 0 Then
         '   r_dbl_MtoTer = 0
         '   r_dbl_MtoPer = 0
         'End If
         'If r_dbl_FmvCob > 0 Then
         '   If r_dbl_RemPer > 0 Then
         '      If r_dbl_FmvCob > r_dbl_RemPer Then
         '         r_dbl_MtoTer = r_dbl_RemPer
         '      Else
         '         r_dbl_MtoTer = 0
         '      End If
         '      r_dbl_MtoPer = r_dbl_RemPer - r_dbl_MtoTer
         '   End If
         'Else
         '   r_dbl_MtoTer = 0
         '   r_dbl_MtoPer = 0
         'End If
         'FINAL - CALCULO ANTERIOR
         
         'INICIO - CALCULO NUEVO
         r_dbl_Saldo2 = r_dbl_MtoSal * (2 / 3)
         
         If r_dbl_MtoBas > r_dbl_Saldo2 Then
            r_dbl_MtoPer = r_dbl_MtoSal - r_dbl_Saldo2
         Else
            r_dbl_MtoPer = r_dbl_Saldo2 - r_dbl_MtoBas
         End If
         'FINAL - CALCULO NUEVO
      Else
         If r_dbl_FmvCob >= r_dbl_RemPer Then
            r_dbl_RieCon = r_dbl_RemPer
            r_dbl_MtoPer = r_dbl_RemPer - r_dbl_RieCon
         Else
            r_dbl_MtoPer = 0
         End If
      End If
   End If
   
   '*Determina el total
   r_dbl_MtoTot = r_dbl_MtoSal - r_dbl_RieCon - r_dbl_MtoMen - r_dbl_MtoTer
   
   '* determina monto sin garantia
   p_MtoSga = r_dbl_MtoPer
   
   '* determina monto con garantia
   p_MtoCga = r_dbl_MtoTot - p_MtoSga
End Sub

Public Sub modprc_ff_CalculaMontosBaseProv2(ByRef p_MtoCga As Double, ByRef p_MtoSga As Double, ByVal p_MtoGto As Double, ByVal p_SalCap As Double, ByVal p_SalCon As Double, ByVal p_CodPrd As String, ByVal p_MonPre As Double, _
                                            ByVal p_TipCam As Double, ByVal p_FecDes As String, ByVal p_TipGar As Integer, ByVal p_MtoGar As Double, ByVal p_MonGar As Integer, ByVal p_IntDif As Double, ByVal p_Origen As Integer)
Dim r_dbl_MtoRemate     As Double
Dim r_dbl_MtoCobTercio  As Double
Dim r_dbl_GastosEjecu   As Double
Dim r_dbl_MtoSalDeuda   As Double
Dim r_dbl_SaldoRemate   As Double
Dim r_dbl_AplicaDeuda   As Double
Dim r_dbl_PerdidaEspe   As Double
Dim r_dbl_CoberturaFMV  As Double
Dim r_dbl_PerdidaMic    As Double

   r_dbl_MtoRemate = 0
   r_dbl_MtoCobTercio = 0
   r_dbl_GastosEjecu = 0
   r_dbl_MtoSalDeuda = 0
   r_dbl_SaldoRemate = 0
   r_dbl_AplicaDeuda = 0
   r_dbl_PerdidaEspe = 0
   r_dbl_CoberturaFMV = 0
   r_dbl_PerdidaMic = 0
   
   '* Saldo total de deuda
   r_dbl_MtoSalDeuda = p_SalCap + p_SalCon - p_IntDif
   
   '* Si no tiene garantia determina los montos directamente
   If Not (p_TipGar = 1 Or p_TipGar = 2 Or p_TipGar = 9) Then
      p_MtoSga = r_dbl_MtoSalDeuda
      p_MtoCga = 0
      Exit Sub
   End If
   
   '* Monto(Base) remate 2/3 garantia
   r_dbl_MtoRemate = Format(p_MtoGar * (2 / 3), "########0.00")
   
   '* Gastos de ejecucion
   r_dbl_GastosEjecu = p_MtoGto
   
   '* Cobertura 1/3 - FMV
   If Not (p_CodPrd = "002" Or p_CodPrd = "006" Or p_CodPrd = "011") Then
      r_dbl_MtoCobTercio = Format(r_dbl_MtoSalDeuda * (1 / 3), "########0.00")
   End If
   
   '* Saldo del Remate
   r_dbl_SaldoRemate = r_dbl_MtoRemate - r_dbl_GastosEjecu
   
   '* Aplicacion de la deuda
   If (p_CodPrd = "002" Or p_CodPrd = "006" Or p_CodPrd = "011") Or (p_CodPrd = "001" Or p_CodPrd = "003") Then
      '* Productos CME, CRC y MICASITA
      If r_dbl_MtoSalDeuda < r_dbl_SaldoRemate Then
         r_dbl_AplicaDeuda = r_dbl_MtoSalDeuda
      Else
         r_dbl_AplicaDeuda = r_dbl_SaldoRemate
      End If
   Else
      '* Productos FMV
      If (r_dbl_MtoSalDeuda * 2 / 3) < r_dbl_SaldoRemate Then
         r_dbl_AplicaDeuda = r_dbl_MtoSalDeuda * (2 / 3)
      Else
         r_dbl_AplicaDeuda = r_dbl_SaldoRemate
      End If
   End If
   
   '* Perdida esperada
   If r_dbl_AplicaDeuda - r_dbl_MtoSalDeuda >= 0 Then
      r_dbl_PerdidaEspe = 0
   Else
      r_dbl_PerdidaEspe = r_dbl_AplicaDeuda - r_dbl_MtoSalDeuda
   End If
   
   '* Cobertura del FMV
   If (p_CodPrd = "002" Or p_CodPrd = "006" Or p_CodPrd = "011") Then
      r_dbl_CoberturaFMV = 0
   ElseIf (p_CodPrd = "001" Or p_CodPrd = "003") Then
      r_dbl_CoberturaFMV = Format(Abs(r_dbl_PerdidaEspe) / 3, "########0.00")
   Else
      r_dbl_CoberturaFMV = r_dbl_MtoCobTercio
   End If
   
   '* Perdida micasita
   If r_dbl_MtoSalDeuda - r_dbl_AplicaDeuda - r_dbl_CoberturaFMV < 0 Then
      r_dbl_PerdidaMic = 0
   Else
      r_dbl_PerdidaMic = r_dbl_MtoSalDeuda - r_dbl_AplicaDeuda - r_dbl_CoberturaFMV
   End If
   
   '* Determina monto con y sin garantia
   p_MtoSga = r_dbl_PerdidaMic
   p_MtoCga = r_dbl_AplicaDeuda
   
   If p_Origen = 2 Then
      p_MtoSga = Abs(r_dbl_PerdidaEspe)
   End If
End Sub

Public Function modprc_gf_Calcula_PBPPerdido(ByVal p_NumOpe As String) As Double
Dim r_rst_Genera        As ADODB.Recordset

   modprc_gf_Calcula_PBPPerdido = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(HIPCUO_CAPBBP) - SUM(HIPCUO_CBPPAG) AS TOTAL FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_CAPBBP > 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If

   r_rst_Genera.MoveFirst
   If IsNull(r_rst_Genera!total) Then
      modprc_gf_Calcula_PBPPerdido = 0
   Else
      modprc_gf_Calcula_PBPPerdido = r_rst_Genera!total
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Function modprc_gf_Calcula_InteresDiferido(ByVal p_NumOpe As String, ByVal p_FecCie As String) As Double
Dim r_rst_Genera        As ADODB.Recordset

   modprc_gf_Calcula_InteresDiferido = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(HIPCUO_INTPAG) AS TOTAL FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT > " & Format(CDate(p_FecCie), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   r_rst_Genera.MoveFirst
   If IsNull(r_rst_Genera!total) Then
      modprc_gf_Calcula_InteresDiferido = 0
   Else
      modprc_gf_Calcula_InteresDiferido = r_rst_Genera!total
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Function modprc_gf_Calcula_InteresVencido(ByVal p_NumOpe As String, ByVal p_FecCie As String, ByVal p_TasInt As Double) As Double
Dim r_rst_Genera        As ADODB.Recordset
Dim r_dbl_DevMes        As Double

   modprc_gf_Calcula_InteresVencido = 0
   
   'obtiene los interes de las cuotas vencidas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(HIPCUO_INTERE - HIPCUO_INTPAG) AS TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & Format(CDate(p_FecCie), "yyyymmdd") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   r_rst_Genera.MoveFirst
   If IsNull(r_rst_Genera!total) Then
      modprc_gf_Calcula_InteresVencido = 0
   Else
      modprc_gf_Calcula_InteresVencido = r_rst_Genera!total
   End If
   
   'obtiene el monto devengado del ultimo mes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_SALCAP, HIPCUO_FECVCT "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & Format(CDate(p_FecCie), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   r_dbl_DevMes = r_rst_Genera!HIPCUO_SALCAP * (1 + (p_TasInt / 100)) ^ ((Format(CDate(p_FecCie), "yyyymmdd") - r_rst_Genera!HIPCUO_FECVCT) / 360) - r_rst_Genera!HIPCUO_SALCAP
   
   'acumula el devengado del mes al interes vencido
   modprc_gf_Calcula_InteresVencido = modprc_gf_Calcula_InteresVencido + Format(r_dbl_DevMes, "########0.00")
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Sub modprc_fs_Actualiza_Proceso(ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_TipPro As Integer)
Dim r_str_Cadena     As String
Dim r_int_NumVec     As Integer
Dim r_rst_Record     As ADODB.Recordset

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS NUM_EJEC "
   r_str_Cadena = r_str_Cadena & "  FROM CTB_PERPRO "
   r_str_Cadena = r_str_Cadena & " WHERE PERPRO_CODANO = " & CStr(p_PerAno) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_CODMES = " & CStr(p_PerMes) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = " & CStr(p_TipPro) & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Record, 3) Then
      Exit Sub
   End If
   
   r_rst_Record.MoveFirst
   r_int_NumVec = r_rst_Record!NUM_EJEC
   
   r_rst_Record.Close
   Set r_rst_Record = Nothing
   
   If r_int_NumVec = 0 Then
      'Inserta
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "INSERT INTO CTB_PERPRO ("
      r_str_Cadena = r_str_Cadena & " PERPRO_CODANO, PERPRO_CODMES, PERPRO_TIPPRO, PERPRO_FECPRO, PERPRO_INDEJE, "
      r_str_Cadena = r_str_Cadena & " SEGUSUCRE, SEGFECCRE, SEGHORCRE, SEGPLTCRE, SEGTERCRE, SEGSUCCRE) "
      r_str_Cadena = r_str_Cadena & "VALUES("
      r_str_Cadena = r_str_Cadena & " " & CStr(p_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & " " & CStr(p_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & " " & CStr(p_TipPro) & ", "
      r_str_Cadena = r_str_Cadena & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      r_str_Cadena = r_str_Cadena & " " & 1 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      r_str_Cadena = r_str_Cadena & " " & Format(Time, "HHMMSS") & ", "
      r_str_Cadena = r_str_Cadena & "'" & UCase(App.EXEName) & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'001'" & ") "
      
      If Not gf_EjecutaSQL(r_str_Cadena, modprc_g_rst_Grabar, 2) Then
         Exit Sub
      End If
      
   Else
      'Actualiza
      r_int_NumVec = r_int_NumVec + 1
      
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "UPDATE CTB_PERPRO "
      r_str_Cadena = r_str_Cadena & "   SET PERPRO_INDEJE = PERPRO_INDEJE + 1, "
      r_str_Cadena = r_str_Cadena & "       SEGUSUACT     = '" & Trim(modgen_g_str_CodUsu) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGFECACT     = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      r_str_Cadena = r_str_Cadena & "       SEGHORACT     = " & Format(Time, "HHMMSS") & ", "
      r_str_Cadena = r_str_Cadena & "       SEGPLTACT     = '" & Trim(UCase(App.EXEName)) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGTERACT     = '" & Trim(modgen_g_str_NombPC) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGSUCACT     = '001' "
      r_str_Cadena = r_str_Cadena & " WHERE PERPRO_CODANO = " & CStr(p_PerAno) & " "
      r_str_Cadena = r_str_Cadena & "   AND PERPRO_CODMES = " & CStr(p_PerMes) & " "
      r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = " & CStr(p_TipPro) & " "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Record, 2) Then
         Exit Sub
      End If
   
   End If
End Sub

Public Function ff_Valida_OperacionGNB(ByVal p_NumOpe As String, ByVal p_FecVct As Long) As Boolean
Dim rst_Operac    As ADODB.Recordset
Dim str_Cadena    As String
Dim lng_FecCan    As Long
   
   ff_Valida_OperacionGNB = False
   
   str_Cadena = ""
   str_Cadena = str_Cadena & "SELECT HIPMAE_SITUAC, HIPMAE_FECCAN "
   str_Cadena = str_Cadena & "  FROM CRE_HIPMAE "
   str_Cadena = str_Cadena & " WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "'"

   If Not gf_EjecutaSQL(str_Cadena, rst_Operac, 3) Then
      Exit Function
   End If
   
   If Not (rst_Operac.BOF And rst_Operac.EOF) Then
      rst_Operac.MoveFirst
      
      If rst_Operac!HIPMAE_SITUAC = 6 Then
         lng_FecCan = CDbl(IIf(IsNull(rst_Operac!HIPMAE_FECCAN) = True, 0, rst_Operac!HIPMAE_FECCAN))
         If Not (p_FecVct <= lng_FecCan) Then
            ff_Valida_OperacionGNB = True
         End If
      End If
   End If
   
   rst_Operac.Close
   Set rst_Operac = Nothing
End Function
