Attribute VB_Name = "bat_modsrv"
Option Explicit

Public g_cls_General    As bat_ClsCnx
Public g_str_CadCnx     As String
Public g_str_RutRpt     As String
Public g_str_RutLog     As String
Public g_rst_Genera     As ADODB.Recordset
Public g_rst_Princi     As ADODB.Recordset
Public g_rst_Listas     As ADODB.Recordset
Public g_rst_GenAux     As ADODB.Recordset
Public g_str_Parame     As String

Public Sub gs_Conecta_Servidor()
   Dim r_bol_FlgRpt     As Boolean
   Dim r_str_NomArc     As String
   Dim r_int_Posici     As Integer

   'Obteniendo Cadenas de Conexión y Rutas
   g_str_CadCnx = gf_LeeInis(modgen_g_str_RutWin & modgen_g_con_ArcIni, "CONEXION", "CADENA")
   g_str_RutRpt = gf_LeeInis(modgen_g_str_RutWin & modgen_g_con_ArcIni, "RUTA_SRV", "PATH_RPT")
   g_str_RutLog = gf_LeeInis(modgen_g_str_RutWin & modgen_g_con_ArcIni, "RUTA_SRV", "PATH_LOG")
   
   Set g_cls_General = New bat_ClsCnx
   r_bol_FlgRpt = g_cls_General.gf_AbreBaseDatos(g_str_CadCnx)

   If r_bol_FlgRpt = False Then
      MsgBox "No se pudo establecer la Conexión." & Chr(13) & Chr(13) & "Error: " & CStr(Err.Number) & " - " & Trim$(Err.Description), vbCritical, modgen_g_str_NomPlt
      End
   End If
   
   'Para obtener Datos de Conexión en Base de Datos
   r_int_Posici = InStr(g_str_CadCnx, "User ID=")
   moddat_g_str_EntDat = Mid(g_str_CadCnx, r_int_Posici + 8, InStr(r_int_Posici + 8, g_str_CadCnx, ";") - (r_int_Posici + 8))
   
   r_int_Posici = InStr(g_str_CadCnx, "Password=")
   moddat_g_str_ClaDat = Mid(g_str_CadCnx, r_int_Posici + 9, InStr(r_int_Posici + 9, g_str_CadCnx, ";") - (r_int_Posici + 9))
   
   r_int_Posici = InStr(g_str_CadCnx, "Data Source=")
   moddat_g_str_NomEsq = Mid(g_str_CadCnx, r_int_Posici + 12, InStr(r_int_Posici + 12, g_str_CadCnx, ";") - (r_int_Posici + 12))
End Sub

Public Function gf_EjecutaSQL(p_Procedimiento As String, p_RecordSet As ADODB.Recordset, p_TipoProceso As Integer) As Boolean
   'True  : Todo Ok.
   'False : Hubo algun error
   
   Dim r_str_Proced     As String
    
   gf_EjecutaSQL = False
   Select Case g_cls_General.gf_EjecutaComandoSQL(p_Procedimiento, p_RecordSet, p_TipoProceso)
      Case -1
         r_str_Proced = gf_ObtieneProcedimiento(p_Procedimiento)
         Call gs_ErrorSQL(r_str_Proced)
         Exit Function
          
      Case Else
         gf_EjecutaSQL = True
         Exit Function
    End Select
End Function

Private Sub gs_ErrorSQL(p_Procedimiento As String)
    Dim r_int_TotErr    As Integer
    Dim r_str_Mensaj    As String
    Dim r_int_Contad    As Integer
    
    r_str_Mensaj = ""
    
    r_int_TotErr = g_cls_General.g_cnx_Conexion.Errors.Count
    
    r_str_Mensaj = r_str_Mensaj & "ERROR al ejecutar un procedimiento." & Chr(13) & Chr(10)
    r_str_Mensaj = r_str_Mensaj & "Store Procedure : " & p_Procedimiento & Chr(13) & Chr(10)
    r_str_Mensaj = r_str_Mensaj & "Total Errores   : " & Format(r_int_TotErr, "00") & Chr(13) & Chr(10)
    r_str_Mensaj = r_str_Mensaj & Chr(13) & Chr(10)
    
    For r_int_Contad = 0 To r_int_TotErr - 1
        r_str_Mensaj = r_str_Mensaj & "Número Error : " & Trim(Str(g_cls_General.g_cnx_Conexion.Errors.Item(r_int_Contad).Number)) & Chr(13) & Chr(10)
        r_str_Mensaj = r_str_Mensaj & "Descripcion  : " & Trim(g_cls_General.g_cnx_Conexion.Errors.Item(r_int_Contad).Description) & Chr(13) & Chr(10)
        'r_str_Mensaj = r_str_Mensaj & "SQLState     : " & Trim(Str(g_cls_General.g_cnx_Conexion.Errors.Item(r_int_Contad).SQLState)) & Chr(13) & Chr(10)
        r_str_Mensaj = r_str_Mensaj & "Source       : " & Trim(g_cls_General.g_cnx_Conexion.Errors.Item(r_int_Contad).Source) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    Next r_int_Contad
    
    MsgBox r_str_Mensaj, vbExclamation, modgen_g_str_NomPlt
End Sub

Private Function gf_ObtieneProcedimiento(p_Procedimiento As String) As String
    Dim r_int_Maximo          As Integer
    Dim r_int_Contad          As Integer
    Dim r_int_TopeMx          As Integer
    Dim r_str_Proced          As String
    
    gf_ObtieneProcedimiento = ""
    r_str_Proced = ""
    
    r_int_TopeMx = 0
    p_Procedimiento = Trim(p_Procedimiento)
    r_int_Maximo = Len(p_Procedimiento)
    
    For r_int_Contad = 1 To r_int_Maximo
        If Mid(p_Procedimiento, r_int_Contad, 1) = " " Then
            r_int_TopeMx = r_int_Contad
            Exit For
        End If
    Next r_int_Contad
    
    If r_int_TopeMx >= 1 Then
        r_str_Proced = Mid(p_Procedimiento, 1, r_int_TopeMx)
    Else
        r_str_Proced = p_Procedimiento
    End If
    
    gf_ObtieneProcedimiento = r_str_Proced
End Function

Public Sub gs_Desconecta_Servidor()
   If Not g_cls_General.gf_CierraBaseDatos() Then
      MsgBox "No se pudo desconectar la Conexión." & Chr(13) & Chr(13) & "Error: " & CStr(Err.Number) & " - " & Trim$(Err.Description), vbCritical, modgen_g_str_NomPlt
      End
   End If
End Sub


