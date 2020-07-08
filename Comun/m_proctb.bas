Attribute VB_Name = "m_proctb"
Option Explicit

Public Sub proctb_CalculaDevengado(ByVal p_NumOpe As String, p_FecPro As String)
   'Declaracion de Variables
   Dim r_rst_Principal        As ADODB.Recordset
   Dim r_rst_Cuotas           As ADODB.Recordset
   
   Dim r_str_CadenaSQL        As String
   Dim r_str_UltDev           As String
   Dim r_str_FecDes           As String
   Dim r_str_FecIni           As String
   Dim r_str_FecFin           As String
   
   Dim r_dbl_MtoPre           As Double
   Dim r_dbl_TasInt           As Double
   
   
   
   
   r_str_CadenaSQL = "SELECT * FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "'"
   
   If Not gf_EjecutaSQL(r_str_CadenaSQL, r_rst_Principal, 3) Then
       Exit Sub
   End If

   If Not (r_rst_Principal.BOF And r_rst_Principal.EOF) Then
      r_rst_Principal.MoveFirst
      
      r_dbl_MtoPre = r_rst_Principal!HIPMAE_MTOPRE
      
      r_str_UltDev = ""
      If r_rst_Principal!HIPMAE_ULTDEV > 0 Then
         r_str_UltDev = gf_FormatoFecha(CStr(r_rst_Principal!HIPMAE_ULTDEV))
      End If
      
      r_str_FecDes = gf_FormatoFecha(CStr(r_rst_Principal!HIPMAE_FECDES))
      r_dbl_TasInt = r_rst_Principal!HIPMAE_TASINT
      
      If Len(r_str_UltDev) = 0 Then
         r_str_FecIni = r_str_FecDes
      Else
         r_str_FecIni = r_str_UltDev
      End If
      
      r_str_FecFin = p_FecPro
      
      
      'Para obtener Cuota Actual
      r_str_CadenaSQL = "SELECT * FROM CRE_HIPCUO WHERE "
      r_str_CadenaSQL = r_str_CadenaSQL & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
      r_str_CadenaSQL = r_str_CadenaSQL & "HIPCUO_TIPCRO = 1  AND "
      r_str_CadenaSQL = r_str_CadenaSQL & "HIPCUO_FECVCT  > " & Format(CDate(r_str_FecIni), "yyyymmdd") & " AND "
      r_str_CadenaSQL = r_str_CadenaSQL & "HIPCUO_FECVCT <= " & Format(CDate(r_str_FecFin), "yyyymmdd") & " "
   
      If Not gf_EjecutaSQL(r_str_CadenaSQL, r_rst_Cuotas, 3) Then
          Exit Sub
      End If
   
      If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
         r_rst_Cuotas.MoveFirst
      End If
   End If
   
   r_rst_Principal.Close
   Set r_rst_Principal = Nothing
End Sub

