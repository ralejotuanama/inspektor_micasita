Attribute VB_Name = "batcbr"
Option Explicit
Public cbr_rst_Princi      As ADODB.Recordset
Public cbr_rst_Accion      As ADODB.Recordset

Public Sub batch_Cbrb001()
   'Cálculo de Cartera Morosa
   Dim r_str_NumOpe     As String
   Dim r_str_FecPro     As String
   
   'Abrir Archivo de LOG
   
   r_str_FecPro = Format(CDate(moddat_g_str_FecSis), "yyyymmdd")
   
   'Inicializando Dias de Morosidad en Cero
   g_str_Parame = "UPDATE CRE_HIPMAE SET HIPMAE_DIAMOR = 0 "
   If Not gf_EjecutaSQL(g_str_Parame, cbr_rst_Princi, 2) Then
      Exit Sub
   End If
   
   
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_PRXVCT <= " & r_str_FecPro & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 OR "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 3  "
   
   If Not gf_EjecutaSQL(g_str_Parame, cbr_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If cbr_rst_Princi.BOF And cbr_rst_Princi.EOF Then
      cbr_rst_Princi.Close
      Set cbr_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   cbr_rst_Princi.MoveFirst
   Do While Not cbr_rst_Princi.EOF
      r_str_NumOpe = Trim(cbr_rst_Princi!HIPMAE_NUMOPE)

      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
   
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CBR0001_CREMOR ("
         g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & r_str_FecPro & ", "
         g_str_Parame = g_str_Parame & "1) "
                  
         If Not gf_EjecutaSQL(g_str_Parame, cbr_rst_Accion, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr = 6 Then
            'Grabar en alguna parte LOG de error
            
            MsgBox "Error Operacion: " & r_str_NumOpe
         End If
         
         DoEvents
      Loop
   
      cbr_rst_Princi.MoveNext
      DoEvents
   Loop

   cbr_rst_Princi.Close
   Set cbr_rst_Princi = Nothing
End Sub


