Attribute VB_Name = "modgue"
Option Explicit

Public modgue_g_str_CodUsu As String
Public modgue_g_str_Nombre As String
Public modgue_g_str_TipHor As String
Public modgue_g_int_NroObs As Integer
Public modgue_g_int_ObsMes As Integer
Public modgue_g_int_TabAsi As Integer
Public modgue_g_str_Justif As String
Public modgue_g_int_FlgIng As Integer
Public modgue_g_int_FgAlSa As Integer

Public Function ff_NumObs(ByVal l_str_PerDia As String, ByVal l_str_PerMes As String, ByVal l_str_PerAno As String, ByVal p_CodUsu As String) As String
  
   ff_NumObs = 0
      
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & "SELECT MAX(OBSASI_NROOBS) AS NROOBS FROM ADM_OBSASI WHERE "
   g_str_Parame = g_str_Parame & "OBSASI_CODUSU = '" & Trim(p_CodUsu) & "' AND "
   g_str_Parame = g_str_Parame & "OBSASI_PERDIA = " & l_str_PerDia & " AND "
   g_str_Parame = g_str_Parame & "OBSASI_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "OBSASI_PERANO = " & l_str_PerAno & " "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If IsNull(g_rst_Listas!NROOBS) Then
      ff_NumObs = 1
   Else
      ff_NumObs = g_rst_Listas!NROOBS + 1
   End If

   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Function

