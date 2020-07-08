Attribute VB_Name = "bat_moddat"
Option Explicit

Public Type moddat_tpo_Genera
   Genera_Codigo  As String
   Genera_Nombre  As String
   Genera_TipPar  As Integer
   Genera_TipVal  As Integer
   Genera_Cantid  As Double
   Genera_ValMin  As Double
   Genera_ValMax  As Double
   Genera_PlzMin  As Double
   Genera_PlzMax  As Double
   Genera_TipMon  As Integer
   Genera_Prefij  As String
   Genera_VenTDo  As Integer
   Genera_VenNDo  As String
   Genera_ConTDo  As Integer
   Genera_ConNDo  As String
   Genera_FlgAso  As Integer
   Genera_TipVia  As Integer
   Genera_NomVia  As String
   Genera_NumVia  As String
   Genera_IntDpt  As String
   Genera_TipZon  As Integer
   Genera_NomZon  As String
   Genera_Refere  As String
   Genera_UbiGeo  As String
   Genera_NumSol  As String
   Genera_TipDoc  As Integer
   Genera_NumDoc  As String
   Genera_NomCli  As String
   Genera_ConHip  As String
   Genera_EjeSeg  As String
   Genera_CodIns  As Integer
End Type

Public moddat_g_int_FlgGOK    As Integer
Public moddat_g_int_CntErr    As Integer
Public moddat_g_str_NomEsq    As String
Public moddat_g_str_EntDat    As String
Public moddat_g_str_ClaDat    As String
Public moddat_g_str_FecSis    As String
Public moddat_g_arr_Genera()  As moddat_tpo_Genera

Public Sub moddat_gs_FecSis()
   moddat_g_str_FecSis = Format(Date, "dd/mm/yyyy")
   
   'Obteniendo Fecha del Sistema
   g_str_Parame = "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy') AS VS_FECSIS FROM DUAL"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   moddat_g_str_FecSis = g_rst_Genera!VS_FECSIS
   DoEvents
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_ParVal(ByVal p_CodGrp As String, ByVal p_CodIte As String) As Double
   moddat_gf_Consulta_ParVal = 0
   
   p_CodIte = Format(p_CodIte, "000")
   
   g_str_Parame = "SELECT * FROM MNT_PARVAL WHERE "
   g_str_Parame = g_str_Parame & "PARVAL_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARVAL_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Consulta_ParVal = g_rst_Genera!PARVAL_CANTID
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Consulta_ParSubPrd(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodGrp As String, ByVal p_CodIte As String) As Integer
   moddat_gf_Consulta_ParSubPrd = False
   
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & p_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Function
   End If
   
   g_rst_Genera.MoveFirst
      
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
   p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
   p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
   p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
   p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
   p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_gf_Consulta_ParSubPrd = True
End Function


Public Function moddat_gf_Verifica_DirEle(p_Arregl() As moddat_tpo_Genera, ByVal p_Codigo As String) As Integer
   Dim r_int_Contad     As Integer
   
   moddat_gf_Verifica_DirEle = False
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      If p_Arregl(r_int_Contad).Genera_Codigo = p_Codigo Then
         moddat_gf_Verifica_DirEle = True
         Exit For
      End If
   Next r_int_Contad
End Function

Public Function moddat_gf_Buscar_DirEle_CodEje(ByVal p_CodEje As String) As String
   moddat_gf_Buscar_DirEle_CodEje = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJECMC WHERE "
   g_str_Parame = g_str_Parame & "EJECMC_CODEJE = '" & p_CodEje & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      
      moddat_gf_Buscar_DirEle_CodEje = Trim(g_rst_Genera!EJECMC_DIRELE & "")
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gs_EnvCor(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, p_Arregl() As moddat_tpo_Genera, p_Asunto As String, p_Contenido As String, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
   Dim r_int_Contad      As Integer
   
   On Error GoTo moddat_gf_EnvCor

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
   
   p_Mensaje.Send
   DoEvents
  
  'Cierra la sesión
  p_Sesion.SignOff
  Exit Sub
  
moddat_gf_EnvCor:
   p_Sesion.SignOff
   MsgBox Err.Description, vbCritical

End Sub


