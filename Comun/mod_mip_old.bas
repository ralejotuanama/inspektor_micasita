Attribute VB_Name = "modmip"
Option Explicit

Public Type modmip_tpo_EstCbr
   EstCbr_Descri        As String
   EstCbr_NumCre        As Integer
   EstCbr_ProAtr        As Integer
   EstCbr_SalCap        As Double
   EstCbr_GarCon        As Double
   EstCbr_GarOtr        As Double
End Type

Public modmip_g_int_PaiRes       As Integer
Public modmip_g_int_CodSex       As Integer
Public modmip_g_int_TDoEmp       As Integer
Public modmip_g_str_NDoEmp       As String
Public modmip_g_str_TDoEmp       As String
Public modmip_g_int_FlgGrb       As Integer
Public modmip_g_int_FlgGrb_1     As Integer
Public modmip_g_int_FlgGrb_2     As Integer
Public modmip_g_int_FlgAct_1     As Integer
Public modmip_g_int_FlgAct_2     As Integer
Public modmip_g_int_FlgAct_3     As Integer
Public modmip_g_int_TipCli       As Integer
Public modmip_g_int_OrdAct       As Integer
Public modmip_g_int_DiaMor       As Integer

Public Sub modmip_gs_Carga_CiuExt(p_Arregl() As moddat_tpo_Genera, p_Combo As ComboBox, ByVal p_CodPai As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_Arregl    -  Arreglo donde se almacenan los datos de las ciudades
'                 p_Combo     -  ComboBox donde se almacenan los datos a mostrar en el formulario que invoca el procedimiento
'                 p_CodPai    -  Código del País para realizar el filtro de ciudades

'F. Creación   :  20-12-2009
'Objetivo      :  Permite cargar las ciudades de un país diferente a Perú, según el país enviado como parámetro
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
Dim r_str_Parame        As String
Dim r_rst_Listas        As ADODB.Recordset

   p_Combo.Clear
   ReDim p_Arregl(0)
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM MNT_CIUEXT WHERE "
   r_str_Parame = r_str_Parame & "CIUEXT_CODPAI = '" & (p_CodPai) & "' "
   r_str_Parame = r_str_Parame & "ORDER BY CIUEXT_DESCRI ASC"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Listas.BOF And r_rst_Listas.EOF) Then
      r_rst_Listas.MoveFirst
      
      Do While Not r_rst_Listas.EOF
         p_Combo.AddItem Trim(r_rst_Listas!CIUEXT_DESCRI)
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Listas!CIUEXT_CODCIU)
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(r_rst_Listas!CIUEXT_DESCRI)
      
         r_rst_Listas.MoveNext
      Loop
   End If
   
   r_rst_Listas.Close
   Set r_rst_Listas = Nothing
End Sub

Public Sub modmip_gs_Carga_CiuExt_Arregl(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPai As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_Arregl    -  Arreglo donde se almacenan los datos de las ciudades
'                 p_Combo     -  ComboBox donde se almacenan los datos a mostrar en el formulario que invoca el procedimiento
'                 p_CodPai    -  Código del País para realizar el filtro de ciudades

'F. Creación   :  20-12-2009
'Objetivo      :  Permite cargar las ciudades de un país diferente a Perú, según el país enviado como parámetro
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
Dim r_str_Parame        As String
Dim r_rst_Listas        As ADODB.Recordset
   
   ReDim p_Arregl(0)
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM MNT_CIUEXT WHERE "
   r_str_Parame = r_str_Parame & "CIUEXT_CODPAI = '" & (p_CodPai) & "' "
   r_str_Parame = r_str_Parame & "ORDER BY CIUEXT_DESCRI ASC"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Listas.BOF And r_rst_Listas.EOF) Then
      r_rst_Listas.MoveFirst
      
      Do While Not r_rst_Listas.EOF
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Listas!CIUEXT_CODCIU)
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(r_rst_Listas!CIUEXT_DESCRI)
      
         r_rst_Listas.MoveNext
      Loop
   End If
   
   r_rst_Listas.Close
   Set r_rst_Listas = Nothing
End Sub

Public Function modmip_gf_Consulta_CiuExt(ByVal p_CodPai As String, ByVal p_CodCiu As String) As String
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodPai    -  Código del País que se desea buscar
'                 p_CodCiu    -  Código de la Ciudad que se desea buscar

'F. Creación   :  20-12-2009
'Objetivo      :  Envía el nombre de la ciudad buscada según el Código de País y el Código de la Ciudad
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
Dim r_str_Parame        As String
Dim r_rst_Listas        As ADODB.Recordset
   
   modmip_gf_Consulta_CiuExt = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM MNT_CIUEXT WHERE "
   r_str_Parame = r_str_Parame & "CIUEXT_CODPAI = '" & (p_CodPai) & "' AND "
   r_str_Parame = r_str_Parame & "CIUEXT_CODCIU = '" & (p_CodCiu) & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Listas.BOF And r_rst_Listas.EOF) Then
      r_rst_Listas.MoveFirst
      modmip_gf_Consulta_CiuExt = Trim(r_rst_Listas!CIUEXT_DESCRI)
   End If
   
   r_rst_Listas.Close
   Set r_rst_Listas = Nothing
End Function

Public Sub modmip_gs_RelLab(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByRef p_RelLab As Integer, ByRef p_TDoVin As Integer, ByRef p_NDoVin As String, ByRef p_TipVin As Integer)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_TipDoc    -  Tipo de DOI de Persona a buscar
'                 p_NumDoc    -  Número de DOI de Persona a buscar
'                 p_RelLab    -  Flag de Retorno si persona tiene Relación Laboral
'                 p_TDoVin    -  Tipo de DOI de Persona Vinculada
'                 p_NDoVin    -  Número de DOI de Persona Vinculada
'                 p_TipVin    -  Tipo de Vinculación

'F. Creación   :  20-12-2009
'Objetivo      :  Busca Vinculación de una Persona con miCasita
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_str_Parame     As String
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_int_FlgTra     As Integer
   Dim r_int_FlgDHb     As Integer
   Dim r_int_PerVin     As Integer
   Dim r_int_OtrVin     As Integer
   Dim r_int_VinVin     As Integer
   Dim r_int_RelLab     As Integer
   Dim r_int_RelAcc     As Integer
   Dim r_int_TDoVin     As Integer
   Dim r_str_NDoVin     As String
   Dim r_int_TipBus     As Integer
   
   'Relación Laboral
   '0 - Ninguno
   '1 - Director
   '2 - Funcionario
   '3 - Trabajador
   '4 - Director y Funcionario
   '5 - Cónyuge o Pariente
   
   'Accionista
   '0 - No es Accionista
   '1 - Si es Accionista
   '2 - Cónyuge o Pariente
   
   'Inicializando Variables de Salida
   r_int_RelLab = 0
   r_int_RelAcc = 0
   r_int_TDoVin = 0
   r_str_NDoVin = ""
   r_int_TipBus = 0
   r_int_FlgTra = 0
   r_int_FlgDHb = 0
   r_int_PerVin = 0
   r_int_VinVin = 0
   
   'Buscando en Maestro de Personal
   r_str_Parame = "SELECT * FROM ADM_FICPER WHERE "
   r_str_Parame = r_str_Parame & "FICPER_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "FICPER_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "FICPER_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_int_TDoVin = r_rst_Princi!FICPER_TIPDOC
      r_str_NDoVin = Trim(r_rst_Princi!FICPER_NUMDOC)
      r_int_FlgTra = 1
      r_int_TipBus = 1
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Buscando en Derecho-Habientes
   r_str_Parame = "SELECT * FROM ADM_DERHAB WHERE "
   r_str_Parame = r_str_Parame & "DERHAB_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "DERHAB_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "DERHAB_SITUAC = 1"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
      
      r_int_TDoVin = r_rst_Princi!DERHAB_TDOPER
      r_str_NDoVin = Trim(r_rst_Princi!DERHAB_NDOPER)
      r_int_FlgDHb = 1
      r_int_TipBus = 2
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Buscando en Tabla de Otros Vinculados de Personal
   r_str_Parame = "SELECT * FROM ADM_PERVIN WHERE "
   r_str_Parame = r_str_Parame & "PERVIN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "PERVIN_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "PERVIN_SITUAC = 1"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_int_TDoVin = r_rst_Princi!PERVIN_TDOPER
      r_str_NDoVin = Trim(r_rst_Princi!PERVIN_NDOPER)
      r_int_PerVin = 1
      r_int_TipBus = 3
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Buscando en Otros Funcionarios miCasita
   r_str_Parame = "SELECT * FROM ADM_VINOTR WHERE "
   r_str_Parame = r_str_Parame & "VINOTR_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "VINOTR_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "VINOTR_SITUAC = 1"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_int_TDoVin = r_rst_Princi!VINOTR_TIPDOC
      r_str_NDoVin = Trim(r_rst_Princi!VINOTR_TIPDOC)
      r_int_OtrVin = r_rst_Princi!VINOTR_CODVIN
      r_int_TipBus = 4
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Buscando en Vinculados de Otros Funcionarios miCasita
   r_str_Parame = "SELECT * FROM ADM_VINVIN WHERE "
   r_str_Parame = r_str_Parame & "VINVIN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "VINVIN_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "VINVIN_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_int_TDoVin = r_rst_Princi!VINVIN_TDOVIN
      r_str_NDoVin = Trim(r_rst_Princi!VINVIN_NDOVIN)
      r_int_VinVin = 1
      r_int_TipBus = 5
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   If r_int_FlgTra = 1 Then
      'Si es Trabajador
      r_int_RelLab = 3
   ElseIf r_int_FlgDHb = 1 Or r_int_PerVin = 1 Then
      'Si es Derecho-Habiente o tiene otra Vinculación con Trabajador
      r_int_RelLab = 5
   End If
   
   If r_int_OtrVin > 0 And r_int_OtrVin < 4 Then
      Select Case r_int_OtrVin
         Case 1:  r_int_RelLab = 1     'Si es Director
         Case 2:  r_int_RelLab = 2     'Si es Funcionario
         Case 3:  r_int_RelLab = 4     'Si es Funcionar y Director
      End Select
   End If
   
   If r_int_VinVin = 1 Then
      r_int_RelLab = 5                 'Si es Cónyuge o Pariente de Otro Funcionario
   End If
   
   p_RelLab = r_int_RelLab
   p_TDoVin = r_int_TDoVin
   p_NDoVin = r_str_NDoVin
   
   p_TipVin = r_int_TipBus
End Sub

Public Sub modmip_gs_RelAcc(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByRef p_RelAcc As Integer, ByRef p_TDoVin As Integer, ByRef p_NDoVin As String, ByRef p_TipVin As Integer)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_TipDoc    -  Tipo de DOI de Persona a buscar
'                 p_NumDoc    -  Número de DOI de Persona a buscar
'                 p_RelAcc    -  Flag de Retorno si persona tiene Relación como Accionista
'                 p_TDoVin    -  Tipo de DOI de Persona Vinculada
'                 p_NDoVin    -  Número de DOI de Persona Vinculada
'                 p_TipVin    -  Tipo de Vinculación

'F. Creación   :  20-12-2009
'Objetivo      :  Busca Vinculación de una Persona con Accionistas de miCasita
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_int_RelAcc     As Integer
Dim r_int_TDoVin     As Integer
Dim r_str_NDoVin     As String
Dim r_int_TipBus     As Integer
   
   'Accionista
   '0 - No es Accionista
   '1 - Si es Accionista
   '2 - Cónyuge o Pariente
   
   'Inicializando Variables de Salida
   r_int_RelAcc = 0
   r_int_TDoVin = 0
   r_str_NDoVin = ""
   r_int_TipBus = 0
   
   'Buscando en Maestro de Personal
   r_str_Parame = "SELECT * FROM ADM_ACCMAE WHERE "
   r_str_Parame = r_str_Parame & "ACCMAE_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "ACCMAE_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "ACCMAE_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_int_TDoVin = r_rst_Princi!ACCMAE_TIPDOC
      r_str_NDoVin = Trim(r_rst_Princi!ACCMAE_NUMDOC)
      r_int_RelAcc = 1
      r_int_TipBus = 1
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   'Buscando en Otros Funcionarios miCasita
   r_str_Parame = "SELECT * FROM ADM_ACCVIN WHERE "
   r_str_Parame = r_str_Parame & "ACCVIN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "ACCVIN_NUMDOC = '" & p_NumDoc & "' AND "
   r_str_Parame = r_str_Parame & "ACCVIN_SITUAC = 1"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_int_TDoVin = r_rst_Princi!ACCVIN_TDOACC
      r_str_NDoVin = Trim(r_rst_Princi!ACCVIN_NDOACC)
      r_int_RelAcc = 2
      r_int_TipBus = 2
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   p_RelAcc = r_int_RelAcc
   p_TDoVin = r_int_TDoVin
   p_NDoVin = r_str_NDoVin
   p_TipVin = r_int_TipBus
End Sub

Public Function modmip_gf_Consulta_NomTra(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_TipDoc    -  Tipo de DOI de Trabajador a buscar
'                 p_NumDoc    -  Número de DOI de Trabajador a buscar

'F. Creación   :  22-12-2009
'Objetivo      :  Devuelve el Nombre del Trabajador
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim r_str_Parame        As String
Dim r_rst_Listas        As ADODB.Recordset
   
   modmip_gf_Consulta_NomTra = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM ADM_FICPER WHERE "
   r_str_Parame = r_str_Parame & "FICPER_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "FICPER_NUMDOC = '" & (p_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Listas.BOF And r_rst_Listas.EOF) Then
      r_rst_Listas.MoveFirst
      modmip_gf_Consulta_NomTra = Trim(r_rst_Listas!FICPER_APEPAT) & " " & Trim(r_rst_Listas!FICPER_APEMAT) & " " & Trim(r_rst_Listas!FICPER_NOMBRE)
   End If
   
   r_rst_Listas.Close
   Set r_rst_Listas = Nothing
End Function

Public Function modmip_gf_Consulta_NomOtrFun(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_TipDoc    -  Tipo de DOI de Trabajador a buscar
'                 p_NumDoc    -  Número de DOI de Trabajador a buscar

'F. Creación   :  22-12-2009
'Objetivo      :  Devuelve el Nombre del Funcionario miCasita
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   Dim r_str_Parame        As String
   Dim r_rst_Listas        As ADODB.Recordset
   
   modmip_gf_Consulta_NomOtrFun = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM ADM_VINOTR WHERE "
   r_str_Parame = r_str_Parame & "VINOTR_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "VINOTR_NUMDOC = '" & (p_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Listas.BOF And r_rst_Listas.EOF) Then
      r_rst_Listas.MoveFirst
      modmip_gf_Consulta_NomOtrFun = Trim(r_rst_Listas!VINOTR_APEPAT) & " " & Trim(r_rst_Listas!VINOTR_APEMAT) & " " & Trim(r_rst_Listas!VINOTR_NOMBRE)
   End If
   
   r_rst_Listas.Close
   Set r_rst_Listas = Nothing
End Function

Public Function modmip_gf_Consulta_NomAcc(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_TipDoc    -  Tipo de DOI de Trabajador a buscar
'                 p_NumDoc    -  Número de DOI de Trabajador a buscar

'F. Creación   :  22-12-2009
'Objetivo      :  Devuelve el Nombre del Accionista
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   Dim r_str_Parame        As String
   Dim r_rst_Listas        As ADODB.Recordset
   
   modmip_gf_Consulta_NomAcc = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM ADM_ACCMAE WHERE "
   r_str_Parame = r_str_Parame & "ACCMAE_TIPDOC = " & CStr(p_TipDoc) & " AND "
   r_str_Parame = r_str_Parame & "ACCMAE_NUMDOC = '" & (p_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Listas.BOF And r_rst_Listas.EOF) Then
      r_rst_Listas.MoveFirst
      
      If r_rst_Listas!ACCMAE_TIPDOC = 7 Then
         modmip_gf_Consulta_NomAcc = Trim(r_rst_Listas!ACCMAE_RAZSOC)
      Else
         modmip_gf_Consulta_NomAcc = Trim(r_rst_Listas!ACCMAE_APEPAT) & " " & Trim(r_rst_Listas!ACCMAE_APEMAT) & " " & Trim(r_rst_Listas!ACCMAE_NOMBRE)
      End If
   End If
   
   r_rst_Listas.Close
   Set r_rst_Listas = Nothing
End Function

Public Sub modmip_gs_DatApo(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, p_Grid As MSFlexGrid)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If g_rst_Princi!DATGEN_APOTDO > 0 Then
         p_Grid.Redraw = False
         g_rst_Princi.MoveFirst
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Documento de Identidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DATGEN_APOTDO)) & " - " & Trim(g_rst_Princi!DATGEN_APONDO & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Apellidos y Nombres"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOAPP) & " " & Trim(g_rst_Princi!DATGEN_APOAPM) & " " & Trim(g_rst_Princi!DATGEN_APONOM)
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Sexo"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DATGEN_APOSEX))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Fecha de Nacimiento"
         p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_APOFNC))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nacionalidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_APONAC))
      
         If Trim(g_rst_Princi!DATGEN_APONAC) = "004028" Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Lugar de Nacimiento"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_APOLNC, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_APOLNC, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_APOLNC))
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Estado Civil"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_APOECV))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nivel de Estudios"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DATGEN_APOEST))
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Profesión"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DATGEN_APOPRF))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         If Len(Trim(g_rst_Princi!DATGEN_REFERE & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Referencia"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Departamento / Provincia / Distrito"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
         
         If Len(Trim(g_rst_Princi!DATGEN_APOTEL & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Domicilio"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOTEL & "")
         End If
      
         If Len(Trim(g_rst_Princi!DATGEN_APOCEL & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Celular"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOCEL & "")
         End If
            
         If Len(Trim(g_rst_Princi!DatGen_DirEle & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "E-mail"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_APOCOR & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Autorización Envío"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_APOAEN))
         End If
         
         p_Grid.Redraw = True
         Call gs_UbiIniGrid(p_Grid)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_EvaCre(p_Grid As MSFlexGrid)
   Dim r_rst_Princi     As ADODB.Recordset
   Dim r_str_EmpSeg     As String
   
   'Buscando Empresa de Seguros
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_str_EmpSeg = Trim(r_rst_Princi!SOLMAE_ESGDES & "")
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Sub
   End If
   
   r_rst_Princi.MoveFirst
   If r_rst_Princi!EVACRE_MTOPRE_CAL > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Monto Préstamo Aprobado"
      p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Plazo Aprobado"
      p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:            p_Grid.Text = CStr(r_rst_Princi!EVACRE_PLAANO_CAL) & " Años "
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Período de Gracia Aprobado"
      p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:            p_Grid.Text = CStr(r_rst_Princi!EVACRE_PERGRA_CAL) & IIf(r_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Cuota Extraordinaria Aprobado"
      p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(r_rst_Princi!EVACRE_CUODBL_CAL))
      
      If moddat_g_int_TipMon <> 1 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Cambio de Aprobación"
         p_Grid.Col = 1:                  p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:         p_Grid.Text = "S/. " & gf_FormatoNumero(r_rst_Princi!EVACRE_TIPCAM, 14, 4)
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Tipo de Seguro Aprobado"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_TipSeg(r_str_EmpSeg, r_rst_Princi!EVACRE_TIPSEG_CAL)
      
      p_Grid.Rows = p_Grid.Rows + 2
   Else
      p_Grid.Rows = p_Grid.Rows + 1
   End If
   
   p_Grid.Row = p_Grid.Rows - 1:
   p_Grid.Col = 0:                     p_Grid.Text = "Total Ingreso Líquido Neto S/."
   p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:            p_Grid.Text = "S/. " & gf_FormatoNumero(r_rst_Princi!EVACRE_INGNET, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Cuota Mensual Máxima"
   p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:            p_Grid.Text = "S/. " & gf_FormatoNumero(r_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Cuota Mensual Máxima M. Prest."
   p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_rst_Princi!EVACRE_CUOMPR, 12, 2)
      
   If moddat_g_int_TipMon <> 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Cambio (Cálculo Ingresos)"
      p_Grid.Col = 1:                  p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:         p_Grid.Text = "S/. " & gf_FormatoNumero(r_rst_Princi!EVACRE_TCAING, 14, 4)
   End If

   Call gs_UbiIniGrid(p_Grid)
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_DatInm_OLD(p_Grid As MSFlexGrid)
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Grid.Redraw = False
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         p_Grid.Col = 1
         p_Grid.Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.Text = "Tipo de Inmueble"
      p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.Text = "Dirección"
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      If Len(Trim(g_rst_Princi!SOLINM_REFERE & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                   p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      End If
      
      If Len(Trim(g_rst_Princi!SOLINM_ESTACI & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                   p_Grid.Text = "Estacionamiento"
         p_Grid.Col = 1:                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.Text = "Departamento / Provincia / Distrito"
      p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.Text = "Proyecto miCasita"
      p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                p_Grid.Text = "Proyecto anclado en Otra IFI"
            p_Grid.Col = 1:                p_Grid.Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                p_Grid.Text = "Nombre Proyecto"
            p_Grid.Col = 1:                p_Grid.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                p_Grid.Text = "Nombre Proyecto"
               p_Grid.Col = 1:                p_Grid.Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                      p_Grid.Text = "Propietario / Promotor"
         p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                      p_Grid.Text = "Docum. Identidad Propietario/Promotor"
         p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                      p_Grid.Text = "Nombre o Razón Social"
         p_Grid.Col = 1:                      p_Grid.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                      p_Grid.Text = "Dirección"
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         If Len(Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Referencia"
            p_Grid.Col = 1:                      p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:          p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                         p_Grid.Text = "Departamento / Provincia / Distrito"
         p_Grid.Col = 1:                         p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         If Len(Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Teléfono"
            p_Grid.Col = 1:                      p_Grid.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         End If
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Docum. Identidad Constructor"
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Nombre o Razón Social"
            p_Grid.Col = 1:                      p_Grid.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            If Len(Trim(g_rst_Princi!SOLINM_REFERE_CON & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            End If
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            If Len(Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Teléfono"
               p_Grid.Col = 1:                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
            End If
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Proyecto Vinculado"
            Else
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Entidad Financiera"
               p_Grid.Col = 1:                   p_Grid.Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Proyecto No Vinculado"
            End If
         
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Docum. Identidad Propietario"
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Nombre o Razón Social"
            p_Grid.Col = 1:                      p_Grid.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            If Len(Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            End If
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            If Len(Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.Text = "Teléfono"
               p_Grid.Col = 1:                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
            End If
         Else
            'Promotor
            p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Doc. Ident. Promotor"
            p_Grid.Col = 1:                      p_Grid.Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Razón Social Promotor"
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Doc. Ident. Constructor"
            p_Grid.Col = 1:                      p_Grid.Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.Text = "Razón Social Constructor"
            p_Grid.Col = 1:                      p_Grid.Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      p_Grid.Redraw = True
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_DatInm(p_Grid As MSFlexGrid, p_Color As Boolean)
Dim r_str_Cadena    As String

   Call gs_LimpiaGrid(p_Grid)
   r_str_Cadena = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT SL.*, EL.EVALEG_FEENIN, PY.DATGEN_VENTDO, PY.DATGEN_VENNDO, PY.DATGEN_CONTDO, PY.DATGEN_CONNDO "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLINM SL "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN PY ON PY.DATGEN_CODIGO = SL.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "   LEFT JOIN TRA_EVALEG EL ON EL.EVALEG_NUMSOL = SL.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & "  WHERE SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Grid.Redraw = False
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.Text = "Modalidad"
      
      '*********DATOS DEL INMUEBLE*********
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         p_Grid.Col = 1
         p_Grid.Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = "Tipo de Inmueble"
      p_Grid.Col = 1:                       p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = "Dirección"
      p_Grid.Col = 1:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      If Len(Trim(g_rst_Princi!SOLINM_REFERE & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
         p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
         p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      End If
      
      If Len(Trim(g_rst_Princi!SOLINM_ESTACI & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
         p_Grid.Text = "Estacionamiento"
         p_Grid.Col = 1:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
         p_Grid.Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = "Departamento / Provincia / Distrito"
      p_Grid.Col = 1:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      If Not IsNull(g_rst_Princi!EVALEG_FEENIN) Then
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
         p_Grid.Text = "Fecha Entrega Inmueble"
         p_Grid.Col = 1:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
         p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FEENIN))
      End If
      
      p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = "Proyecto miCasita"
      p_Grid.Col = 1:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
      p_Grid.Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
            
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         '-------------------ANTIGUOS CREDITOS---------------------
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                p_Grid.CellForeColor = modgen_g_con_ColNeg
            p_Grid.Text = "Proyecto anclado en Otra IFI"
            p_Grid.Col = 1:                p_Grid.CellForeColor = modgen_g_con_ColNeg
            p_Grid.Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                p_Grid.CellForeColor = modgen_g_con_ColNeg
            p_Grid.Text = "Nombre Proyecto"
            p_Grid.Col = 1:                p_Grid.CellForeColor = modgen_g_con_ColNeg
            p_Grid.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                p_Grid.CellForeColor = modgen_g_con_ColNeg
               p_Grid.Text = "Nombre Proyecto"
               p_Grid.Col = 1:                p_Grid.CellForeColor = modgen_g_con_ColNeg
               p_Grid.Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
         If (Len(Trim(g_rst_Princi!SOLINM_TIPDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)) > 0) Then
            '*********DATOS DEL PROMOTOR*********
             p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = "Propietario / Promotor"
             p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
             
             p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = "Docum. Identidad Propietario/Promotor"
             p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
                
             p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = "Nombre o Razón Social"
             p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
                
             p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = "Dirección"
             p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
                
             If Len(Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")) > 0 Then
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Referencia"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
             End If
                
             p_Grid.Rows = p_Grid.Rows + 1:          p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                         p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = "Departamento / Provincia / Distrito"
             p_Grid.Col = 1:                         p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
             p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
                
             If Len(Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")) > 0 Then
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Teléfono"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
             End If
         Else
             If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
                'Promotor
                 r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_VENTDO, g_rst_Princi!DATGEN_VENNDO)
                 If (Len(Trim(r_str_Cadena)) > 0) Then
                     p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                     p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                     p_Grid.Text = "Doc. Ident. Promotor"
                     p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                     p_Grid.Text = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO)
                          
                     p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                     p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                     p_Grid.Text = "Razón Social Promotor"
                     p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                     p_Grid.Text = r_str_Cadena
                 End If
             End If
         End If
         
         '*********DATOS DEL CONSTRUCTOR*********
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = "Docum. Identidad Constructor"
            p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = "Nombre o Razón Social"
            p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = "Dirección"
            p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            If Len(Trim(g_rst_Princi!SOLINM_REFERE_CON & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
               p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
               p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            End If
            
            p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
            p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            If Len(Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
               p_Grid.Text = "Teléfono"
               p_Grid.Col = 1:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
               p_Grid.Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
            End If
         Else
            If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
               'Constructor
                r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_CONTDO, g_rst_Princi!DATGEN_CONNDO)
                If (Len(Trim(r_str_Cadena)) > 0) Then
                    p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                    p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = "Doc. Ident. Constructor"
                    p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = CStr(g_rst_Princi!DATGEN_CONTDO) & "-" & Trim(g_rst_Princi!DATGEN_CONNDO)
                    
                    p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                    p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = "Razón Social Constructor"
                    p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = r_str_Cadena
                End If
            End If
         End If
      Else
      '-------------------NUEVO CREDITOS---------------------
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
               p_Grid.Text = "Proyecto Vinculado"
            Else
               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
               p_Grid.Text = "Entidad Financiera"
               p_Grid.Col = 1:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
               p_Grid.Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)

               p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                   p_Grid.CellForeColor = modgen_g_con_ColNeg
               p_Grid.Text = "Proyecto No Vinculado"
            End If

            p_Grid.Col = 1:                      p_Grid.CellForeColor = modgen_g_con_ColNeg
            p_Grid.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Then ' Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            '*********DATOS DEL PROMOTOR*********
            If (Len(Trim(g_rst_Princi!SOLINM_TIPDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)) > 0) Then
                p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Docum. Identidad Propietario"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
                
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Nombre o Razón Social"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
                
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Dirección"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                                  " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                                  IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
                
                If Len(Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")) > 0 Then
                   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
                   p_Grid.Col = 0:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                   p_Grid.Text = "Referencia"
                   p_Grid.Col = 1:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
                End If
                
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Departamento / Provincia / Distrito"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
                
                If Len(Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")) > 0 Then
                   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
                   p_Grid.Col = 0:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                   p_Grid.Text = "Teléfono"
                   p_Grid.Col = 1:                   p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                   p_Grid.Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
                End If
            Else
                If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
                   'Promotor
                    r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_VENTDO, g_rst_Princi!DATGEN_VENNDO)
                    If (Len(Trim(r_str_Cadena)) > 0) Then
                        p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                        p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                        p_Grid.Text = "Doc. Ident. Promotor"
                        p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                        p_Grid.Text = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO)
                        
                        p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                        p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                        p_Grid.Text = "Razón Social Promotor"
                        p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                        p_Grid.Text = r_str_Cadena
                    End If
                End If
            End If
            If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
               'Constructor
                r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_CONTDO, g_rst_Princi!DATGEN_CONNDO)
                If (Len(Trim(r_str_Cadena)) > 0) Then
                    p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                    p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = "Doc. Ident. Constructor"
                    p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = CStr(g_rst_Princi!DATGEN_CONTDO) & "-" & Trim(g_rst_Princi!DATGEN_CONNDO)
                    
                    p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                    p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = "Razón Social Constructor"
                    p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                    p_Grid.Text = r_str_Cadena
                End If
            End If
         Else
         '*********BIEN FUTURO**********
           'Promotor
            r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            If (Len(Trim(r_str_Cadena)) > 0) Then
                p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Doc. Ident. Promotor"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
                
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = "Razón Social Promotor"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColAzu, modgen_g_con_ColNeg)
                p_Grid.Text = r_str_Cadena
            End If
            'Constructor
            r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
            If (Len(Trim(r_str_Cadena)) > 0) Then
                p_Grid.Rows = p_Grid.Rows + 2:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                p_Grid.Text = "Doc. Ident. Constructor"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                p_Grid.Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
                
                p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                p_Grid.Text = "Razón Social Constructor"
                p_Grid.Col = 1:                      p_Grid.CellForeColor = IIf(p_Color = True, modgen_g_con_ColRoj, modgen_g_con_ColNeg)
                p_Grid.Text = r_str_Cadena
            End If
         End If
      End If
      
      p_Grid.Redraw = True
      
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_DatCre(p_Grid As MSFlexGrid, ByRef r_arr_Mtz() As moddat_g_tpo_DatCom)
   modmip_g_int_DiaMor = 0
   Call gs_LimpiaGrid(p_Grid)
   ReDim r_arr_Mtz(0)
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      moddat_g_str_NumOpe = ""
   Else
      moddat_g_str_NumOpe = g_rst_Genera!HIPMAE_NUMOPE
   End If
   
   If IsNull(moddat_g_str_NumOpe) Or moddat_g_str_NumOpe = "" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
      g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   p_Grid.Redraw = False
   g_rst_Princi.MoveFirst

   If IsNull(moddat_g_str_NumOpe) Or moddat_g_str_NumOpe = "" Then

          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Producto"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)

          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Sub-Producto"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)

          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Tipo de Evaluación"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))

          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Primera Vivienda"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))

          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Moneda del Préstamo"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Solicitud"
          p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))

          If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
             If g_rst_Princi!SOLMAE_TIPMON = 2 Then
                p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Valor de Compra Venta"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)

                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Cuota Inicial"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2) & " (" & Format((CDbl(g_rst_Princi!SOLMAE_APOPRO_DOL) - CDbl(g_rst_Princi!SOLMAE_AFPMTO) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)) / g_rst_Princi!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%)"

                p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                     p_Grid.Text = "                Aporte Propio"
                p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(CDbl(g_rst_Princi!SOLMAE_APOPRO_DOL) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) - CDbl(g_rst_Princi!SOLMAE_AFPMTO), 12, 2)

               If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                     p_Grid.Text = "                Importe BBP"
                  p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
               End If

                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Monto Préstamo"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
          Else
                p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Valor de Compra Venta"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)

                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Cuota Inicial"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format((CDbl(g_rst_Princi!SOLMAE_APOPRO_SOL) - CDbl(g_rst_Princi!SOLMAE_AFPMTO) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)) / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%) "  'INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & " "

                p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                     p_Grid.Text = "                Aporte Propio"
                p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(CDbl(g_rst_Princi!SOLMAE_APOPRO_SOL) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO)), 12, 2)

               If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                     p_Grid.Text = "                Importe BBP"
                  p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
               End If

                'If g_rst_Princi!SOLMAE_CODPRD = "021" Or g_rst_Princi!SOLMAE_CODPRD = "022" Or g_rst_Princi!SOLMAE_CODPRD = "023" Then
'                If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
'                   p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%) INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & " "
'                Else
'                   p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%)"
'                End If

                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Monto Préstamo"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
             End If

             p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Tasa de Interés"
             p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Plazo (Años)"
             p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Período de Gracia (Meses)"
             p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Interés Capitalizado"
             p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
             p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_INTGRA, 12, 2)

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Cuotas Extraordinarias"
             p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Compañía de Seguros"
             p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Tipo de Seguro Desgravamen"
             p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Día de Pago"
             p_Grid.Col = 1:                        p_Grid.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
             modmip_g_int_DiaMor = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")

             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Observaciones"
             p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
          End If

          If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
              p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
              p_Grid.Col = 0:                        p_Grid.Text = "Institución Financiera de Ahorro"
              p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)

              p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
              p_Grid.Col = 0:                        p_Grid.Text = "Monto Mínimo de Ahorro Mensual"
              p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
              p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_MONAHO)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)

              p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
              p_Grid.Col = 0:                        p_Grid.Text = "Meses Ahorrados"
              p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
          End If
   Else

         'Cargando en Grid
         p_Grid.Rows = 0
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Producto"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)

         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Sub-Producto"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)

         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Tipo de Evaluación"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))

         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Primera Vivienda"

         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_PRIVIV))

         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Moneda Préstamo"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))

         p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Valor Compra Venta"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_CVTSOL, g_rst_Princi!HIPMAE_CVTDOL), 12, 2) 'g_rst_Princi!HIPMAE_CVTSOL

         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Cuota Inicial" '"Aporte Propio"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_APOSOL, g_rst_Princi!HIPMAE_APODOL), 12, 2) 'g_rst_Princi!HIPMAE_APOSOL

         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "                Aporte Propio"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(CDbl(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_APOSOL, g_rst_Princi!HIPMAE_APODOL)) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO)), 12, 2) 'g_rst_Princi!HIPMAE_APOSOL

      If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "                Importe BBP"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
      End If
'         If moddat_g_int_TipMon = 1 Then
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Valor Compra Venta"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
'
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Aporte Propio"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            If g_rst_Princi!HIPMAE_CODPRD = "021" Or g_rst_Princi!HIPMAE_CODPRD = "022" Or g_rst_Princi!HIPMAE_CODPRD = "023" Then
'               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ") "
'            Else
'               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
'            End If
'         Else
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Valor Compra Venta"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
'
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Aporte Propio"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
'         End If

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Monto Préstamo"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)

         If g_rst_Princi!HIPMAE_FECESC > 0 Then
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha Firma EE.PP"
            p_Grid.Col = 1:                 p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
         End If

         p_Grid.Rows = p_Grid.Rows + 2:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Plazo"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Tasa de Interés"
         p_Grid.Col = 1:                    p_Grid.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Nro. de Cuotas"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Período de Gracia"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Cuotas Extraordinarias"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Día de Pago"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_DIAPAG)

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Interes Capitalizado"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Compañía de Seguros"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")

         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Tipo de Seguro Desg."
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)

'         p_Grid.Rows = p_Grid.Rows + 2
'         p_Grid.Row = p_Grid.Rows - 1
'         p_Grid.Col = 0
'         p_Grid.Text = "Consejero Hipotecario"
'
'         p_Grid.Col = 1
'         p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!HIPMAE_CONHIP)
   End If

    'Datos Evaluación de Créditos
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT TO_DATE(EVACRE_FECCOM,'YYYY/MM/DD') AS EVACRE_FECCOM, EVACRE_NROACT "
   g_str_Parame = g_str_Parame & "FROM TRA_EVACRE "
    g_str_Parame = g_str_Parame & "WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

    If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
        p_Grid.Rows = p_Grid.Rows + 2:            p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                           p_Grid.Text = "Fecha del comité de crédito"

        p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                           p_Grid.Text = "Número de acta"
    Else
        If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha del comité de crédito"

            p_Grid.Rows = p_Grid.Rows + 1:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Número de acta"
        Else
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha del comité de crédito"
            p_Grid.Col = 1:                 p_Grid.Text = IIf(IsNull(g_rst_GenAux!EVACRE_FECCOM), "", (g_rst_GenAux!EVACRE_FECCOM))

            p_Grid.Rows = p_Grid.Rows + 1:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Número de acta"
            p_Grid.Col = 1:                 p_Grid.Text = IIf(IsNull(g_rst_GenAux!EVACRE_NROACT), "", (g_rst_GenAux!EVACRE_NROACT))

            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
        End If
    End If

    p_Grid.Rows = p_Grid.Rows + 2:            p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                           p_Grid.Text = "Consejero Hipotecario"
    p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)

    p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                           p_Grid.Text = "Ejecutivo de Seguimiento"
    p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)

    p_Grid.Redraw = True
    Call gs_UbiIniGrid(p_Grid)

    ReDim r_arr_Mtz(0)

    r_arr_Mtz(0).DatCom_PriViv = g_rst_Princi!SOLMAE_PRIVIV
    r_arr_Mtz(0).DatCom_TipMon = g_rst_Princi!SOLMAE_TIPMON
    If (IsNull(g_rst_Princi!SOLMAE_CODMOD) = False) Then
       r_arr_Mtz(0).DatCom_CodMod = Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "00")
    Else
       r_arr_Mtz(0).DatCom_CodMod = "00"
    End If
    r_arr_Mtz(0).DatCom_TasInt = g_rst_Princi!SOLMAE_TASINT
    r_arr_Mtz(0).DatCom_ComVta_Dol = g_rst_Princi!SOLMAE_COMVTA_DOL
    r_arr_Mtz(0).DatCom_ApoPro_Dol = g_rst_Princi!SOLMAE_APOPRO_DOL
    r_arr_Mtz(0).DatCom_ComVta_Sol = g_rst_Princi!SOLMAE_COMVTA_SOL
    r_arr_Mtz(0).DatCom_ApoPro_Sol = g_rst_Princi!SOLMAE_APOPRO_SOL
    r_arr_Mtz(0).DatCom_MtoPre_Sol = g_rst_Princi!SOLMAE_MTOPRE_SOL
    r_arr_Mtz(0).DatCom_MtoPre_Dol = g_rst_Princi!SOLMAE_MTOPRE_DOL
    r_arr_Mtz(0).DatCom_PerGra = g_rst_Princi!SOLMAE_PERGRA
    r_arr_Mtz(0).DatCom_PlaAno = g_rst_Princi!SOLMAE_PLAANO
    r_arr_Mtz(0).DatCom_CuoExt = g_rst_Princi!SOLMAE_CUOEXT
    r_arr_Mtz(0).DatCom_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
    r_arr_Mtz(0).DatCom_IntGra = g_rst_Princi!SOLMAE_INTGRA
    r_arr_Mtz(0).DatCom_Observ = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
    r_arr_Mtz(0).DatCom_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
    r_arr_Mtz(0).DatCom_CodSub = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
    r_arr_Mtz(0).DatCom_EjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
    r_arr_Mtz(0).DatCom_ConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
    r_arr_Mtz(0).DatCom_FecSol = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
    r_arr_Mtz(0).DatCom_CodIns = g_rst_Princi!SOLMAE_CODINS
    r_arr_Mtz(0).DatCom_TipEva = g_rst_Princi!SOLMAE_TIPEVA
    r_arr_Mtz(0).DatCom_MtoPre_Mpr = g_rst_Princi!SOLMAE_MTOPRE_MPR
    r_arr_Mtz(0).DatCom_EsgDes = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
    r_arr_Mtz(0).DatCom_ComVta_Mon = g_rst_Princi!SOLMAE_COMVTA_MON
    r_arr_Mtz(0).DatCom_MtoPre_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_MTOPRE_CAL) = True, 0, g_rst_Princi!SOLMAE_MTOPRE_CAL)
    r_arr_Mtz(0).DatCom_PlaAno_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_PLAANO_CAL) = True, 0, g_rst_Princi!SOLMAE_PLAANO_CAL)
    r_arr_Mtz(0).DatCom_PerGra_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_PERGRA_CAL) = True, 0, g_rst_Princi!SOLMAE_PERGRA_CAL)
    r_arr_Mtz(0).DatCom_CuoExt_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_CUOEXT_CAL) = True, 0, g_rst_Princi!SOLMAE_CUOEXT_CAL)
    r_arr_Mtz(0).DatCom_TipSeg_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_TIPSEG_CAL) = True, 0, g_rst_Princi!SOLMAE_TIPSEG_CAL)
    r_arr_Mtz(0).DatCom_MesAho = IIf(IsNull(g_rst_Princi!SOLMAE_MESAHO) = True, 0, g_rst_Princi!SOLMAE_MESAHO)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_DatCre_2(p_Grid As MSFlexGrid, ByRef r_arr_Mtz() As moddat_g_tpo_DatCom)
    Call gs_LimpiaGrid(p_Grid)
    ReDim r_arr_Mtz(0)
    
    If IsNull(moddat_g_str_NumOpe) Or moddat_g_str_NumOpe = "" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
      g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
    Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
    End If
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
        Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Exit Sub
    End If
    
    p_Grid.Redraw = False
    g_rst_Princi.MoveFirst
    
    If IsNull(moddat_g_str_NumOpe) Or moddat_g_str_NumOpe = "" Then
    
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Producto"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
          
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Sub-Producto"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
          
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Tipo de Evaluación"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
          
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Primera Vivienda"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))
          
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Moneda del Préstamo"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
          
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Solicitud"
          p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
          If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
             If g_rst_Princi!SOLMAE_TIPMON = 2 Then
                p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Valor de Compra Venta"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
             
                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Cuota Inicial"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2) & " (" & Format((CDbl(g_rst_Princi!SOLMAE_APOPRO_DOL) - CDbl(g_rst_Princi!SOLMAE_AFPMTO) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)) / g_rst_Princi!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%)"
                  
                p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                     p_Grid.Text = "                Aporte Propio"
                p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(CDbl(g_rst_Princi!SOLMAE_APOPRO_DOL) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) - CDbl(g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
         
               If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                     p_Grid.Text = "                Importe BBP"
                  p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
               End If
                  
                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Monto Préstamo"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
          Else
                p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Valor de Compra Venta"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
             
                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Cuota Inicial"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format((CDbl(g_rst_Princi!SOLMAE_APOPRO_SOL) - CDbl(g_rst_Princi!SOLMAE_AFPMTO) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)) / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%) "  'INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & " "
                
                p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                     p_Grid.Text = "                Aporte Propio"
                p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(CDbl(g_rst_Princi!SOLMAE_APOPRO_SOL) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(g_rst_Princi!SOLMAE_BMSMTO) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO)), 12, 2)
         
               If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                     p_Grid.Text = "                Importe BBP"
                  p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
               End If
               If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
                  p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
                  p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
                  p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
                  p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
               End If
                
                'If g_rst_Princi!SOLMAE_CODPRD = "021" Or g_rst_Princi!SOLMAE_CODPRD = "022" Or g_rst_Princi!SOLMAE_CODPRD = "023" Then
'                If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
'                   p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%) INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & " "
'                Else
'                   p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%)"
'                End If
             
                p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
                p_Grid.Col = 0:                        p_Grid.Text = "Monto Préstamo"
                p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
                p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
             End If
          
             p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Tasa de Interés"
             p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
          
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Plazo (Años)"
             p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
          
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Período de Gracia (Meses)"
             p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"
         
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Interés Capitalizado"
             p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
             p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_INTGRA, 12, 2)
         
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Cuotas Extraordinarias"
             p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))
             
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Compañía de Seguros"
             p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
             
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Tipo de Seguro Desgravamen"
             p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
             
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Día de Pago"
             p_Grid.Col = 1:                        p_Grid.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
             
             p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
             p_Grid.Col = 0:                        p_Grid.Text = "Observaciones"
             p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
          End If
         
          If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
              p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
              p_Grid.Col = 0:                        p_Grid.Text = "Institución Financiera de Ahorro"
              p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
              
              p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
              p_Grid.Col = 0:                        p_Grid.Text = "Monto Mínimo de Ahorro Mensual"
              p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
              p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_MONAHO)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
              
              p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
              p_Grid.Col = 0:                        p_Grid.Text = "Meses Ahorrados"
              p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
          End If
   Else
         
         'Cargando en Grid
         p_Grid.Rows = 0
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Producto"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
         
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Sub-Producto"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
         
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Tipo de Evaluación"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))

         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Primera Vivienda"
         
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_PRIVIV))
         
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Moneda Préstamo"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
         
         p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Valor Compra Venta"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_CVTSOL, g_rst_Princi!HIPMAE_CVTDOL), 12, 2) 'g_rst_Princi!HIPMAE_CVTSOL
   
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Cuota Inicial" '"Aporte Propio"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_APOSOL, g_rst_Princi!HIPMAE_APODOL), 12, 2) 'g_rst_Princi!HIPMAE_APOSOL
      
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "                Aporte Propio"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(CDbl(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_APOSOL, g_rst_Princi!HIPMAE_APODOL)) - CDbl(g_rst_Princi!SOLMAE_FMVBBP) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO)), 12, 2) 'g_rst_Princi!HIPMAE_APOSOL

      If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "                Importe BBP"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
      End If
'         If moddat_g_int_TipMon = 1 Then
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Valor Compra Venta"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
'
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Aporte Propio"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            If g_rst_Princi!HIPMAE_CODPRD = "021" Or g_rst_Princi!HIPMAE_CODPRD = "022" Or g_rst_Princi!HIPMAE_CODPRD = "023" Then
'               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ") "
'            Else
'               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
'            End If
'         Else
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Valor Compra Venta"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
'
'            p_Grid.Rows = p_Grid.Rows + 1
'            p_Grid.Row = p_Grid.Rows - 1
'            p_Grid.Col = 0
'            p_Grid.Text = "Aporte Propio"
'
'            p_Grid.Col = 1
'            p_Grid.CellFontName = "Lucida Console"
'            p_Grid.CellFontSize = 8
'            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
'         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Monto Préstamo"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
         
         If g_rst_Princi!HIPMAE_FECESC > 0 Then
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha Firma EE.PP"
            p_Grid.Col = 1:                 p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
         End If
         
         p_Grid.Rows = p_Grid.Rows + 2:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Plazo"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Tasa de Interés"
         p_Grid.Col = 1:                    p_Grid.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Nro. de Cuotas"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Período de Gracia"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Cuotas Extraordinarias"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Día de Pago"
         p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_DIAPAG)
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Interes Capitalizado"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Compañía de Seguros"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:     p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Tipo de Seguro Desg."
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
         
'         p_Grid.Rows = p_Grid.Rows + 2
'         p_Grid.Row = p_Grid.Rows - 1
'         p_Grid.Col = 0
'         p_Grid.Text = "Consejero Hipotecario"
'
'         p_Grid.Col = 1
'         p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!HIPMAE_CONHIP)
   End If
   
    'Datos Evaluación de Créditos
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT TO_DATE(EVACRE_FECCOM,'YYYY/MM/DD') AS EVACRE_FECCOM, EVACRE_NROACT "
    g_str_Parame = g_str_Parame & "FROM TRA_EVACRE "
    g_str_Parame = g_str_Parame & "WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "
 
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
        p_Grid.Rows = p_Grid.Rows + 2:            p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                           p_Grid.Text = "Fecha del comité de crédito"
        
        p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                           p_Grid.Text = "Número de acta"
    Else
        If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha del comité de crédito"
            
            p_Grid.Rows = p_Grid.Rows + 1:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Número de acta"
        Else
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha del comité de crédito"
            p_Grid.Col = 1:                 p_Grid.Text = IIf(IsNull(g_rst_GenAux!EVACRE_FECCOM), "", (g_rst_GenAux!EVACRE_FECCOM))
            
            p_Grid.Rows = p_Grid.Rows + 1:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Número de acta"
            p_Grid.Col = 1:                 p_Grid.Text = IIf(IsNull(g_rst_GenAux!EVACRE_NROACT), "", (g_rst_GenAux!EVACRE_NROACT))
            
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
        End If
    End If
    
    p_Grid.Rows = p_Grid.Rows + 2:            p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                           p_Grid.Text = "Consejero Hipotecario"
    p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
    
    p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                           p_Grid.Text = "Ejecutivo de Seguimiento"
    p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
    p_Grid.Redraw = True
    Call gs_UbiIniGrid(p_Grid)
       
    ReDim r_arr_Mtz(0)
        
    r_arr_Mtz(0).DatCom_PriViv = g_rst_Princi!SOLMAE_PRIVIV
    r_arr_Mtz(0).DatCom_TipMon = g_rst_Princi!SOLMAE_TIPMON
    If (IsNull(g_rst_Princi!SOLMAE_CODMOD) = False) Then
       r_arr_Mtz(0).DatCom_CodMod = Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "00")
    Else
       r_arr_Mtz(0).DatCom_CodMod = "00"
    End If
    r_arr_Mtz(0).DatCom_TasInt = g_rst_Princi!SOLMAE_TASINT
    r_arr_Mtz(0).DatCom_ComVta_Dol = g_rst_Princi!SOLMAE_COMVTA_DOL
    r_arr_Mtz(0).DatCom_ApoPro_Dol = g_rst_Princi!SOLMAE_APOPRO_DOL
    r_arr_Mtz(0).DatCom_ComVta_Sol = g_rst_Princi!SOLMAE_COMVTA_SOL
    r_arr_Mtz(0).DatCom_ApoPro_Sol = g_rst_Princi!SOLMAE_APOPRO_SOL
    r_arr_Mtz(0).DatCom_MtoPre_Sol = g_rst_Princi!SOLMAE_MTOPRE_SOL
    r_arr_Mtz(0).DatCom_MtoPre_Dol = g_rst_Princi!SOLMAE_MTOPRE_DOL
    r_arr_Mtz(0).DatCom_PerGra = g_rst_Princi!SOLMAE_PERGRA
    r_arr_Mtz(0).DatCom_PlaAno = g_rst_Princi!SOLMAE_PLAANO
    r_arr_Mtz(0).DatCom_CuoExt = g_rst_Princi!SOLMAE_CUOEXT
    r_arr_Mtz(0).DatCom_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
    r_arr_Mtz(0).DatCom_IntGra = g_rst_Princi!SOLMAE_INTGRA
    r_arr_Mtz(0).DatCom_Observ = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
    r_arr_Mtz(0).DatCom_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
    r_arr_Mtz(0).DatCom_CodSub = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
    r_arr_Mtz(0).DatCom_EjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
    r_arr_Mtz(0).DatCom_ConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
    r_arr_Mtz(0).DatCom_FecSol = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
    r_arr_Mtz(0).DatCom_CodIns = g_rst_Princi!SOLMAE_CODINS
    r_arr_Mtz(0).DatCom_TipEva = g_rst_Princi!SOLMAE_TIPEVA
    r_arr_Mtz(0).DatCom_MtoPre_Mpr = g_rst_Princi!SOLMAE_MTOPRE_MPR
    r_arr_Mtz(0).DatCom_EsgDes = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
    r_arr_Mtz(0).DatCom_ComVta_Mon = g_rst_Princi!SOLMAE_COMVTA_MON
    r_arr_Mtz(0).DatCom_MtoPre_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_MTOPRE_CAL) = True, 0, g_rst_Princi!SOLMAE_MTOPRE_CAL)
    r_arr_Mtz(0).DatCom_PlaAno_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_PLAANO_CAL) = True, 0, g_rst_Princi!SOLMAE_PLAANO_CAL)
    r_arr_Mtz(0).DatCom_PerGra_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_PERGRA_CAL) = True, 0, g_rst_Princi!SOLMAE_PERGRA_CAL)
    r_arr_Mtz(0).DatCom_CuoExt_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_CUOEXT_CAL) = True, 0, g_rst_Princi!SOLMAE_CUOEXT_CAL)
    r_arr_Mtz(0).DatCom_TipSeg_Cal = IIf(IsNull(g_rst_Princi!SOLMAE_TIPSEG_CAL) = True, 0, g_rst_Princi!SOLMAE_TIPSEG_CAL)
    r_arr_Mtz(0).DatCom_MesAho = IIf(IsNull(g_rst_Princi!SOLMAE_MESAHO) = True, 0, g_rst_Princi!SOLMAE_MESAHO)
                                
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_DatCre_Old(p_Grid As MSFlexGrid, p_TxtObs As TextBox)
    Call gs_LimpiaGrid(p_Grid)
    
    g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
    g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
        Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Exit Sub
    End If
   
    p_Grid.Redraw = False
    g_rst_Princi.MoveFirst
    
    p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                     p_Grid.Text = "Producto"
    p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
    
    p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                     p_Grid.Text = "Sub-Producto"
    p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
    
    p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                     p_Grid.Text = "Tipo de Evaluación"
    p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
    
    p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                     p_Grid.Text = "Primera Vivienda"
    p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))
    
    p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                     p_Grid.Text = "Moneda del Préstamo"
    p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
    
    p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Solicitud"
    p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))

    If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
       If g_rst_Princi!SOLMAE_TIPMON = 2 Then
          p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                        p_Grid.Text = "Valor de Compra Venta"
          p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
          p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
       
          p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                        p_Grid.Text = "Aporte Propio"
          p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
          p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_DOL / g_rst_Princi!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%)"
       
          p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                        p_Grid.Text = "Monto Préstamo"
          p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
          p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
       Else
          p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                        p_Grid.Text = "Valor de Compra Venta"
          p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
          p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
       
          p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                        p_Grid.Text = "Aporte Propio"
          p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
          
          'If g_rst_Princi!SOLMAE_CODPRD = "021" Or g_rst_Princi!SOLMAE_CODPRD = "022" Or g_rst_Princi!SOLMAE_CODPRD = "023" Then
          If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
             p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%) INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & " "
          Else
             p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & " (" & Format(g_rst_Princi!SOLMAE_APOPRO_SOL / g_rst_Princi!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%)"
          End If
       
          p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                        p_Grid.Text = "Monto Préstamo"
          p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
          p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
       End If
    
       p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Tasa de Interés"
       p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
    
       p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Plazo"
       p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
    
       p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Período de Gracia (Meses)"
       p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"
    
       p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Cuotas Extraordinarias"
       p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))
       
       p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Compañía de Seguros"
       p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
       
       p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Tipo de Seguro Desgravamen"
       p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
       
       p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
       p_Grid.Col = 0:                        p_Grid.Text = "Día de Pago"
       p_Grid.Col = 1:                        p_Grid.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
    End If
    
    If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
        p_Grid.Rows = p_Grid.Rows + 2:         p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                        p_Grid.Text = "Institución Financiera de Ahorro"
        p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
        
        p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                        p_Grid.Text = "Monto Mínimo de Ahorro Mensual"
        p_Grid.Col = 1:                        p_Grid.CellFontName = "Lucida Console"
        p_Grid.CellFontSize = 8:               p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_MONAHO)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
        
        p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                        p_Grid.Text = "Meses Ahorrados"
        p_Grid.Col = 1:                        p_Grid.Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
    End If
    
    'Datos Evaluación de Créditos
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT TO_DATE(EVACRE_FECCOM,'YYYY/MM/DD') AS EVACRE_FECCOM, EVACRE_NROACT "
    g_str_Parame = g_str_Parame & "FROM TRA_EVACRE "
    g_str_Parame = g_str_Parame & "WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "
 
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
        p_Grid.Rows = p_Grid.Rows + 2:            p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                           p_Grid.Text = "Fecha del comité de crédito"
        
        p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
        p_Grid.Col = 0:                           p_Grid.Text = "Número de acta"
    Else
        If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha del comité de crédito"
            
            p_Grid.Rows = p_Grid.Rows + 1:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Número de acta"
        Else
            p_Grid.Rows = p_Grid.Rows + 2:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Fecha del comité de crédito"
            p_Grid.Col = 1:                 p_Grid.Text = IIf(IsNull(g_rst_GenAux!EVACRE_FECCOM), "", (g_rst_GenAux!EVACRE_FECCOM))
7
            p_Grid.Rows = p_Grid.Rows + 1:  p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                 p_Grid.Text = "Número de acta"
            p_Grid.Col = 1:                 p_Grid.Text = IIf(IsNull(g_rst_GenAux!EVACRE_NROACT), "", (g_rst_GenAux!EVACRE_NROACT))
            
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
        End If
    End If
    
    '
    p_Grid.Rows = p_Grid.Rows + 2:            p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                           p_Grid.Text = "Consejero Hipotecario"
    p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
    
    p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
    p_Grid.Col = 0:                           p_Grid.Text = "Ejecutivo de Seguimiento"
    p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
    p_Grid.Redraw = True
    Call gs_UbiIniGrid(p_Grid)
    
    p_TxtObs.Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, p_Grid As MSFlexGrid, p_Indice As Integer)
Dim r_str_Cadena As String
Dim r_str_CadAux As String

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Grid.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Documento de Identidad"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Apellidos y Nombres"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If g_rst_Princi!DatGen_FLGDOA = 1 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Documento Adicional de Identidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Sexo"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Nacimiento"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nacionalidad"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Lugar de Nacimiento"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Estado Civil"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DATGEN_REGCYG), "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nivel de Estudios"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Profesión"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))
      
      If (Not IsNull(g_rst_Princi!DATGEN_TIPAFP)) Then
          p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
          p_Grid.Col = 0:                     p_Grid.Text = "Afp"
          p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("517", CStr(g_rst_Princi!DATGEN_TIPAFP))
      End If
      
      If p_Indice = 0 Then 'titular
         If g_rst_Princi!DatGen_DepEco > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Nro. Dependientes Económicos"
            p_Grid.Col = 1:                  p_Grid.Text = CStr(g_rst_Princi!DatGen_DepEco)
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Edades"
            p_Grid.Col = 1:                  p_Grid.Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
         End If
      End If
      
      If p_Indice = 0 Then 'titular
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "País Residencia"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("500", CStr(g_rst_Princi!DATGEN_PAIRES))
         
         modmip_g_int_PaiRes = CLng(g_rst_Princi!DATGEN_PAIRES)
         
         If Trim(g_rst_Princi!DATGEN_PAIRES & "") = "004028" Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                        IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
            
            If Len(Trim(g_rst_Princi!DATGEN_REFERE & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                  p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
            End If
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
                                        
            moddat_g_str_UbiGeo = Left(Format(g_rst_Princi!DatGen_Ubigeo, "000000"), 4)
         Else
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                  p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(g_rst_Princi!DATGEN_PAIRES, "000000"), Trim(g_rst_Princi!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_EXTCPO & "")
         End If
         
         '*********************************************
         'datgen_telefo1, datgen_telefo2, datgen_telefo
         r_str_Cadena = "": r_str_CadAux = ""
         r_str_Cadena = Trim(g_rst_Princi!DatGen_Telefo & "")
         r_str_CadAux = Trim(g_rst_Princi!DatGen_Telefo1 & "")
         If Len(r_str_CadAux) > 0 Then
            r_str_Cadena = IIf(Len(r_str_Cadena) > 0, r_str_Cadena & " - " & r_str_CadAux, r_str_CadAux)
         End If
         r_str_CadAux = Trim(g_rst_Princi!DatGen_Telefo2 & "")
         If Len(r_str_CadAux) > 0 Then
            r_str_Cadena = IIf(Len(r_str_Cadena) > 0, r_str_Cadena & " - " & r_str_CadAux, r_str_CadAux)
         End If
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Teléfono"
         p_Grid.Col = 1:                     p_Grid.Text = r_str_Cadena
         '*********************************************
         'datgen_numcel, datgen_numcel1, datgen_numcel2
         
         r_str_Cadena = "": r_str_CadAux = ""
         r_str_Cadena = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         r_str_CadAux = Trim(g_rst_Princi!DATGEN_NUMCEL1 & "")
         If Len(r_str_CadAux) > 0 Then
            r_str_Cadena = IIf(Len(r_str_Cadena) > 0, r_str_Cadena & " - " & r_str_CadAux, r_str_CadAux)
         End If
         r_str_CadAux = Trim(g_rst_Princi!DATGEN_NUMCEL2 & "")
         If Len(r_str_CadAux) > 0 Then
            r_str_Cadena = IIf(Len(r_str_Cadena) > 0, r_str_Cadena & " - " & r_str_CadAux, r_str_CadAux)
         End If
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Celular"
         p_Grid.Col = 1:                     p_Grid.Text = r_str_Cadena

         If Len(Trim(g_rst_Princi!DatGen_DirEle & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "E-mail"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DatGen_DirEle & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Autorización Envío"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
         End If
         
         If g_rst_Princi!DATGEN_TDOVIN > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Vinculado"
            p_Grid.Col = 1
            
            If g_rst_Princi!DATGEN_TIPVIN = 1 Then
               p_Grid.Text = "TRABAJADOR"
            ElseIf g_rst_Princi!DATGEN_TIPVIN = 2 Or g_rst_Princi!DATGEN_TIPVIN = 3 Then
               p_Grid.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Princi!DATGEN_TDOVIN, Trim(g_rst_Princi!DATGEN_NDOVIN)) & ")"
            ElseIf g_rst_Princi!DATGEN_TIPVIN = 4 Then
               p_Grid.Text = "FUNCIONARIO"
            ElseIf g_rst_Princi!DATGEN_TIPVIN = 5 Then
               p_Grid.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Princi!DATGEN_TDOVIN, Trim(g_rst_Princi!DATGEN_NDOVIN)) & ")"
            End If
         End If
         
         If g_rst_Princi!DATGEN_TDOACC > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Accionista"
            p_Grid.Col = 1
            
            If g_rst_Princi!DATGEN_ACCVIN = 1 Then
               p_Grid.Text = "ACCIONISTA"
            ElseIf g_rst_Princi!DATGEN_ACCVIN = 2 Then
               p_Grid.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Princi!DATGEN_TDOACC, Trim(g_rst_Princi!DATGEN_NDOACC)) & ")"
            End If
         End If
         
         moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
         moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      End If
      'aqui
      p_Grid.Redraw = True
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Económica Principal
   Call modmip_gs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice, p_Grid)
   Call modmip_gs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice, p_Grid)
End Sub

Public Sub modmip_gs_DatCli_OLD(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, p_Grid As MSFlexGrid, p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Grid.Redraw = False
      
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Documento de Identidad"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Apellidos y Nombres"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If g_rst_Princi!DatGen_FLGDOA = 1 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Documento Adicional de Identidad"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Sexo"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Nacimiento"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nacionalidad"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Lugar de Nacimiento"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Estado Civil"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DATGEN_REGCYG), "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nivel de Estudios"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Profesión"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))
      
      If p_Indice = 0 Then
         If g_rst_Princi!DatGen_DepEco > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Nro. Dependientes Económicos"
            p_Grid.Col = 1:                  p_Grid.Text = CStr(g_rst_Princi!DatGen_DepEco)
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Edades"
            p_Grid.Col = 1:                  p_Grid.Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
         End If
      End If
      
      If p_Indice = 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "País Residencia"
         p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("500", CStr(g_rst_Princi!DATGEN_PAIRES))
         
         modmip_g_int_PaiRes = CLng(g_rst_Princi!DATGEN_PAIRES)
         
         If Trim(g_rst_Princi!DATGEN_PAIRES & "") = "004028" Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                        IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
            
            If Len(Trim(g_rst_Princi!DATGEN_REFERE & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                  p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
            End If
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
                                        
            moddat_g_str_UbiGeo = Left(Format(g_rst_Princi!DatGen_Ubigeo, "000000"), 4)
         Else
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                  p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(g_rst_Princi!DATGEN_PAIRES, "000000"), Trim(g_rst_Princi!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DATGEN_EXTCPO & "")
         End If
      
         If Len(Trim(g_rst_Princi!DatGen_Telefo & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Domicilio"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DatGen_Telefo & "")
         End If
      
         If Len(Trim(g_rst_Princi!DATGEN_NUMCEL & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Teléfono Celular"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         End If
         
         If Len(Trim(g_rst_Princi!DatGen_DirEle & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "E-mail"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!DatGen_DirEle & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Autorización Envío"
            p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
         End If
         
         moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
         moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      End If
      
      p_Grid.Redraw = True
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Económica Principal
   Call modmip_gs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice, p_Grid)
   Call modmip_gs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice, p_Grid)
End Sub

Public Sub modmip_gs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer, p_Grid As MSFlexGrid)
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Grid.Redraw = False
   
      p_Grid.Rows = p_Grid.Rows + 2:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Ocupación " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ACTECO_CODACT)
      
      Select Case g_rst_Princi!ACTECO_CODACT
         Case 11: Call modmip_gs_ActEco_Dep(p_Grid)
         Case 21: Call modmip_gs_ActEco_Ind(p_Grid)
         Case 31: Call modmip_gs_ActEco_Com(p_Grid)
         Case 41: Call modmip_gs_ActEco_Acc(p_Grid)
         Case 51: Call modmip_gs_ActEco_Ren(p_Grid)
         Case 61: Call modmip_gs_ActEco_Otr(p_Grid)
      End Select
      
      p_Grid.Redraw = True
      Call gs_UbiIniGrid(p_Grid)
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub modmip_gs_ActEco_Dep(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1
   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0
   p_Grid.Text = "Documento Identidad Empleador"

   p_Grid.Col = 1
   p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   p_Grid.Rows = p_Grid.Rows + 1
   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0
   p_Grid.Text = "Situación como Trabajador"

   p_Grid.Col = 1
   p_Grid.Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
      p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
      
      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                        IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")
   
            If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
            End If
   
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU & ""))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
         End If
      
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                           p_Grid.Text = "Fax"
            p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
         End If
      Else
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                        " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                        IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
      
            If Len(Trim(g_rst_Princi!ActEco_Dep_Refere & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            End If
      
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
      
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_DEP_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_EXTCPO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         If Len(Trim(g_rst_Princi!ActEco_Dep_NumFax & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                           p_Grid.Text = "Fax"
            p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      If Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Teléfono RR.HH"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELERH & "")
      End If
   
      If Len(Trim(g_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Anexo RR.HH"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_Dep_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Fecha de Ingreso"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Cargo"
   p_Grid.Col = 1:                           p_Grid.Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Anexo"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
   End If
   
   If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Teléfono Directo"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
   End If
End Sub

Private Sub modmip_gs_ActEco_Dep_OLD(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1
   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0
   p_Grid.Text = "Documento Identidad Empleador"

   p_Grid.Col = 1
   p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   p_Grid.Rows = p_Grid.Rows + 1
   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0
   p_Grid.Text = "Situación como Trabajador"

   p_Grid.Col = 1
   p_Grid.Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
      p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
      
      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                        IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")
   
            If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
            End If
   
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU & ""))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
         End If
      
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                           p_Grid.Text = "Fax"
            p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
         End If
      Else
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                        " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                        IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
      
            If Len(Trim(g_rst_Princi!ActEco_Dep_Refere & "")) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            End If
      
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
      
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_DEP_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_DEP_EXTCPO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         If Len(Trim(g_rst_Princi!ActEco_Dep_NumFax & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                           p_Grid.Text = "Fax"
            p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      If Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Teléfono RR.HH"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELERH & "")
      End If
   
      If Len(Trim(g_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Anexo RR.HH"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_Dep_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Fecha de Ingreso"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Cargo"
   p_Grid.Col = 1:                           p_Grid.Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Anexo"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
   End If
   
   If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Teléfono Directo"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
   End If
End Sub

Private Sub modmip_gs_ActEco_Ind(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   If modmip_g_int_PaiRes = 4028 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección"
   
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")
   
      If Len(Trim(g_rst_Princi!ActEco_Ind_Refere & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      End If
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Indartamento / Provincia / Distrito"
      p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   Else
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_IND_EXTDIR & "")
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
      p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_IND_EXTCIU))
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_IND_EXTCPO & "")
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   If Len(Trim(g_rst_Princi!ActEco_Ind_NumFax & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                           p_Grid.Text = "Fax"
      p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
   End If

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                        p_Grid.Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_ind_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Fecha de Inicio de Actividades"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Contrato de Locación de Servicios"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad Empleador"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
         p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
         
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                        IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")
   
            If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
            End If
   
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Fax"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      p_Grid.Rows = p_Grid.Rows + 1:               p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                              p_Grid.Text = "Fecha de Ingreso"
      p_Grid.Col = 1:                              p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      p_Grid.Rows = p_Grid.Rows + 1:               p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                              p_Grid.Text = "Cargo"
      p_Grid.Col = 1:                              p_Grid.Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub modmip_gs_ActEco_Ind_OLD(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   If modmip_g_int_PaiRes = 4028 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección"
   
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")
   
      If Len(Trim(g_rst_Princi!ActEco_Ind_Refere & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      End If
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Indartamento / Provincia / Distrito"
      p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   Else
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_IND_EXTDIR & "")
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
      p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_IND_EXTCIU))
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_IND_EXTCPO & "")
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   If Len(Trim(g_rst_Princi!ActEco_Ind_NumFax & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                           p_Grid.Text = "Fax"
      p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
   End If

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                        p_Grid.Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_ind_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Fecha de Inicio de Actividades"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Contrato de Locación de Servicios"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad Empleador"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
   
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
         p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
         
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
         p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
         
         If modmip_g_int_PaiRes = 4028 Then
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
            p_Grid.Col = 1
            p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                        " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                        IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                        IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")
   
            If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
               p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
               p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
               p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
            End If
   
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
            p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
         Else
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
            p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
         
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
         End If
         
         p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
         p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                        p_Grid.Text = "Fax"
            p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      p_Grid.Rows = p_Grid.Rows + 1:               p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                              p_Grid.Text = "Fecha de Ingreso"
      p_Grid.Col = 1:                              p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      p_Grid.Rows = p_Grid.Rows + 1:               p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                              p_Grid.Text = "Cargo"
      p_Grid.Col = 1:                              p_Grid.Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub modmip_gs_ActEco_Com(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Razón Social"
   p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Nombre Comercial"
   p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   If modmip_g_int_PaiRes = 4028 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección"
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")
   
      If Len(Trim(g_rst_Princi!ActEco_Com_Refere & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      End If
   
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Departamento / Provincia / Distrito"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   Else
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Dirección"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_COM_EXTDIR & "")
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
      p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_COM_EXTCIU))
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_COM_EXTCPO & "")
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   If Len(Trim(g_rst_Princi!ActEco_Com_NumFax & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fax"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
   End If

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                  p_Grid.Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Giro Comercial"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_com_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Fecha de Inicio de Operaciones"
   p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Cargo"
   p_Grid.Col = 1:                  p_Grid.Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
End Sub

Private Sub modmip_gs_ActEco_Com_OLD(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Razón Social"
   p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Nombre Comercial"
   p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   If modmip_g_int_PaiRes = 4028 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección"
      p_Grid.Col = 1
      p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")
   
      If Len(Trim(g_rst_Princi!ActEco_Com_Refere & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      End If
   
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Departamento / Provincia / Distrito"
      p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   Else
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Dirección"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_COM_EXTDIR & "")
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
      p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(Format(modmip_g_int_PaiRes, "000000"), Trim(g_rst_Princi!ACTECO_COM_EXTCIU))
   
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
      p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ACTECO_COM_EXTCPO & "")
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   If Len(Trim(g_rst_Princi!ActEco_Com_NumFax & "")) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fax"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
   End If

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                  p_Grid.Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Giro Comercial"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_com_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Fecha de Inicio de Operaciones"
   p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Cargo"
   p_Grid.Col = 1:                  p_Grid.Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
End Sub

Private Sub modmip_gs_ActEco_Acc(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
      p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
      
      If modmip_g_int_PaiRes = 4028 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
         End If

         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
         p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      Else
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
      
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
         p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
      
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
      p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Fax"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_acc_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Antigüedad"
   p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
End Sub

Private Sub modmip_gs_ActEco_Acc_OLD(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                  p_Grid.Text = "Documento Identidad"
   p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Razón Social"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nombre Comercial"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "CIIU"
      p_Grid.Col = 1:                  p_Grid.Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Oficina"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("234", CStr(g_rst_Princi!ActEco_Dep_TipOfi))
      
      If modmip_g_int_PaiRes = 4028 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Dirección"
         p_Grid.Col = 1
         p_Grid.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Genera!DATGEN_INTDPT) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         If Len(Trim(g_rst_Genera!DATGEN_REFERE)) > 0 Then
            p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                     p_Grid.Text = "Referencia"
            p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Genera!DATGEN_REFERE & "")
         End If

         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Departamento / Provincia / Distrito"
         p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      Else
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Domicilio"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTDIR & "")
      
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Ciudad"
         p_Grid.Col = 1:                        p_Grid.Text = modmip_gf_Consulta_CiuExt(g_rst_Genera!DATGEN_PAIRES, Trim(g_rst_Genera!DATGEN_EXTCIU))
      
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Código Postal"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_EXTCPO & "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                           p_Grid.Text = "Teléfono(s)"
      p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      If Len(Trim(g_rst_Genera!DATGEN_NUMFAX & "")) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                        p_Grid.Text = "Fax"
         p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Genera!DATGEN_NUMFAX & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                     p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_acc_MonIng))

   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Antigüedad"
   p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
End Sub

Private Sub modmip_gs_ActEco_Ren(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_ren_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Dirección de Propiedad 01"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Nombre de Arrendatario"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Fecha de Inicio de Alquiler"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Alquiler Mensual"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección de Propiedad 02"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nombre de Arrendatario"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Inicio de Alquiler"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Teléfono(s)"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Alquiler Mensual"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub modmip_gs_ActEco_Ren_OLD(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                        p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_ren_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Dirección de Propiedad 01"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Nombre de Arrendatario"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Fecha de Inicio de Alquiler"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Teléfono(s)"
   p_Grid.Col = 1:                        p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                        p_Grid.Text = "Alquiler Mensual"
   p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Dirección de Propiedad 02"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Nombre de Arrendatario"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Fecha de Inicio de Alquiler"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Teléfono(s)"
      p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                     p_Grid.Text = "Alquiler Mensual"
      p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub modmip_gs_ActEco_Otr(p_Grid As MSFlexGrid)
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Moneda de Ingresos"
   p_Grid.Col = 1:                           p_Grid.Text = moddat_gf_Consulta_ParDes("113", CStr(g_rst_Princi!ActEco_otr_MonIng))
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Ingreso Neto"
   p_Grid.Col = 1:                           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!ActEco_Otr_IngNet, 15, 2)

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Actividad"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Otr_Activi & "")

   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "CIIU"
   p_Grid.Col = 1:                           p_Grid.Text = g_rst_Princi!ActEco_Otr_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Otr_CodCiu))
   
   p_Grid.Rows = p_Grid.Rows + 1:            p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                           p_Grid.Text = "Observaciones"
   p_Grid.Col = 1:                           p_Grid.Text = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
End Sub


Public Sub modmip_gs_EvaTas(p_Grid As MSFlexGrid)
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Empresa Peritaje"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Nombre Perito"
      p_Grid.Col = 1:                       p_Grid.Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Código REPEV SBS"
      p_Grid.Col = 1:                       p_Grid.Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Nro. de Informe"
      p_Grid.Col = 1:                       p_Grid.Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Fecha Evaluación"
      p_Grid.Col = 1:                       p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Año de Construcción"
      p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!EVATAS_ANOCON)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Nro. de Pisos"
      p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!EVATAS_NUMPIS)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Nro. de Sótanos"
      p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!EVATAS_NUMSOT)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Tipo de Inmueble"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!EVATAS_TIPINM))
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Uso de Inmueble"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("222", CStr(g_rst_Princi!EVATAS_USOINM))
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Material de Construcción"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("223", CStr(g_rst_Princi!EVATAS_MATCON))
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Tipo de Moneda"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!EVATAS_TIPMON))
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Tipo de Cambio"
      p_Grid.Col = 1:                       p_Grid.Text = "S/.  " & gf_FormatoNumero(g_rst_Princi!EVATAS_TIPCAM, 10, 4)
      
      'Total
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Area Terreno (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Area Construida (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Suma Asegurada (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Comercial (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Realización (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Terreno (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Edificación (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Areas Comunes (Total)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
      'Inmueble
      p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Area Terreno (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2) & " m2"
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Area Construida (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2) & " m2"
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Suma Asegurada (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Comercial (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Realización (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Terreno (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Edificación (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Areas Comunes (Inmueble)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
      'Estacionamiento 1
      If g_rst_Princi!EVATAS_FLGEST_ES1 = 1 Then
         p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Area Terreno (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2) & " m2"
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Area Construida (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2) & " m2"
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Suma Asegurada (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Comercial (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Realización (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Terreno (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Edificación (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Areas Comunes (Estac. 1)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_ES2 = 1 Then
         p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Area Terreno (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2) & " m2"
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Area Construida (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2) & " m2"
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Suma Asegurada (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Comercial (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Realización (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Terreno (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Edificación (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Areas Comunes (Estac. 2)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_DEP = 1 Then
         p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Area Terreno (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Area Construida (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Suma Asegurada (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Comercial (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Realización (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Terreno (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Edificación (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Valor Areas Comunes (Depósito)"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
      End If
      
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_EvaSeg(p_Grid As MSFlexGrid, p_TxtObs As TextBox)
   p_TxtObs.Text = ""
   
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Empresa de Seguros"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
   
      p_Grid.Rows = p_Grid.Rows + 2:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Seguro Desgravamen"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fecha Evaluación (Seg. Desgravamen)"
      p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVADES))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Valor (Seg. Desgravamen)"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Valor a Aplicar"
      p_Grid.Col = 1:                  p_Grid.Text = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      
      p_Grid.Rows = p_Grid.Rows + 2:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fecha Evaluación (Seg. Inmueble)"
      p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Tipo de Valor (Seg. Inmueble)"
      p_Grid.Col = 1:                  p_Grid.Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Valor a Aplicar"
      p_Grid.Col = 1:                  p_Grid.Text = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
      
      p_TxtObs.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
      
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_Buscar_EvaLeg(p_Grid1 As MSFlexGrid, p_Grid2 As MSFlexGrid, p_TxtObs As TextBox, p_TxtCom As TextBox)
   Call gs_LimpiaGrid(p_Grid1)
   Call gs_LimpiaGrid(p_Grid2)

   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      p_TxtObs.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      
      If g_rst_Princi!EVALEG_FECCOM > 0 Then
         p_TxtCom.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
      End If
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
         p_Grid1.Rows = p_Grid1.Rows + 1:    p_Grid1.Row = p_Grid1.Rows - 1
         p_Grid1.Col = 0:                    p_Grid1.Text = "Fecha Firma Contrato Compra Venta"
         p_Grid1.Col = 1:                    p_Grid1.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               p_Grid1.Rows = p_Grid1.Rows + 1:    p_Grid1.Row = p_Grid1.Rows - 1
               p_Grid1.Col = 0:                    p_Grid1.Text = "Tipo de Cambio SBS"
               p_Grid1.Col = 1:                    p_Grid1.Text = Format(g_rst_Princi!EVALEG_TCASBS, "###,##0.0000")
            End If
         End If
      
         If g_rst_Princi!EVALEG_TCACVT > 0 Then
            p_Grid1.Rows = p_Grid1.Rows + 1:       p_Grid1.Row = p_Grid1.Rows - 1
            p_Grid1.Col = 0:                       p_Grid1.Text = "Tipo de Cambio aplicado"
            p_Grid1.Col = 1:                       p_Grid1.Text = Format(g_rst_Princi!EVALEG_TCACVT, "###,##0.0000")
         End If
      End If
      
      If g_rst_Princi!EVALEG_FIRCON > 0 Then
         p_Grid1.Rows = p_Grid1.Rows + 1:          p_Grid1.Row = p_Grid1.Rows - 1
         p_Grid1.Col = 0:                          p_Grid1.Text = "Fecha Firma Contrato"
         p_Grid1.Col = 1:                          p_Grid1.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      
         p_Grid1.Rows = p_Grid1.Rows + 1:          p_Grid1.Row = p_Grid1.Rows - 1
         p_Grid1.Col = 0:                          p_Grid1.Text = "Notaria"
         p_Grid1.Col = 1:                          p_Grid1.Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT)
      
         p_Grid1.Rows = p_Grid1.Rows + 1:          p_Grid1.Row = p_Grid1.Rows - 1
         p_Grid1.Col = 0:                          p_Grid1.Text = "Representante Legal 1"
         p_Grid1.Col = 1:                          p_Grid1.Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1)
      
         p_Grid1.Rows = p_Grid1.Rows + 1:          p_Grid1.Row = p_Grid1.Rows - 1
         p_Grid1.Col = 0:                          p_Grid1.Text = "Representante Legal 2"
         p_Grid1.Col = 1:                          p_Grid1.Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2)
         
         p_Grid1.Rows = p_Grid1.Rows + 1:          p_Grid1.Row = p_Grid1.Rows - 1
         p_Grid1.Col = 0:                          p_Grid1.Text = "Monto Hipoteca"
         p_Grid1.Col = 1:                          p_Grid1.CellFontName = "Lucida Console"
         p_Grid1.CellFontSize = 8:                 p_Grid1.Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
      End If
      
      If p_Grid1.Rows > 0 Then
         Call gs_UbiIniGrid(p_Grid1)
      End If
      
      If g_rst_Princi!EVALEG_FECBLQ_INM > 0 Then
         p_Grid2.Rows = p_Grid2.Rows + 1:          p_Grid2.Row = p_Grid2.Rows - 1
         p_Grid2.Col = 0:                          p_Grid2.Text = "Sede Registral"
         p_Grid2.Col = 1:                          p_Grid2.Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!EVALEG_SEDREG))
         
         p_Grid2.Rows = p_Grid2.Rows + 1:          p_Grid2.Row = p_Grid2.Rows - 1
         p_Grid2.Col = 0:                          p_Grid2.Text = "Fecha Bloqueo (Inmueble)"
         p_Grid2.Col = 1:                          p_Grid2.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
         
         p_Grid2.Rows = p_Grid2.Rows + 1:          p_Grid2.Row = p_Grid2.Rows - 1
         p_Grid2.Col = 0:                          p_Grid2.Text = "Doc. Registral (Inmueble)"
         p_Grid2.Col = 1:                          p_Grid2.Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_INM)
         
         Select Case g_rst_Princi!EVALEG_TIPDOC_INM
            Case 1
               p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
               
            Case 2
               p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
               
            Case 3
               p_Grid2.Text = p_Grid2.Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
         End Select
         
         If g_rst_Princi!EVALEG_FLGEST_ES1 = 1 Then
            p_Grid2.Rows = p_Grid2.Rows + 2:    p_Grid2.Row = p_Grid2.Rows - 1
            p_Grid2.Col = 0:                    p_Grid2.Text = "Fecha Bloqueo (Estac. 1)"
            p_Grid2.Col = 1:                    p_Grid2.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES1))
            
            p_Grid2.Rows = p_Grid2.Rows + 1:    p_Grid2.Row = p_Grid2.Rows - 1
            p_Grid2.Col = 0:                    p_Grid2.Text = "Doc. Registral (Estac. 1)"
            p_Grid2.Col = 1:                    p_Grid2.Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES1)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES1
               Case 1
                  p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
               
               Case 2
                  p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
               
               Case 3
                  p_Grid2.Text = p_Grid2.Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_ES2 = 1 Then
            p_Grid2.Rows = p_Grid2.Rows + 2:    p_Grid2.Row = p_Grid2.Rows - 1
            p_Grid2.Col = 0:                    p_Grid2.Text = "Fecha Bloqueo (Estac. 2)"
            p_Grid2.Col = 1:                    p_Grid2.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES2))
            
            p_Grid2.Rows = p_Grid2.Rows + 1:    p_Grid2.Row = p_Grid2.Rows - 1
            p_Grid2.Col = 0:                    p_Grid2.Text = "Doc. Registral (Estac. 2)"
            p_Grid2.Col = 1:                    p_Grid2.Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES2)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES2
               Case 1
                  p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
                  
               Case 2
                  p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
                  
               Case 3
                  p_Grid2.Text = p_Grid2.Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
                  
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_DEP = 1 Then
            p_Grid2.Rows = p_Grid2.Rows + 2:    p_Grid2.Row = p_Grid2.Rows - 1
            p_Grid2.Col = 0:                    p_Grid2.Text = "Fecha Bloqueo (Depósito)"
            p_Grid2.Col = 1:                    p_Grid2.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_DEP))
            
            p_Grid2.Rows = p_Grid2.Rows + 1:    p_Grid2.Row = p_Grid2.Rows - 1
            p_Grid2.Col = 0:                    p_Grid2.Text = "Doc. Registral (Depósito)"
            p_Grid2.Col = 1:                    p_Grid2.Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_DEP)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_DEP
               Case 1
                  p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
               
               Case 2
                  p_Grid2.Text = p_Grid2.Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
                  
               Case 3
                  p_Grid2.Text = p_Grid2.Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
      End If
      
      If p_Grid2.Rows > 0 Then
         Call gs_UbiIniGrid(p_Grid2)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_PolSeg(p_Grid As MSFlexGrid)
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM TRA_POLIZA WHERE "
   g_str_Parame = g_str_Parame & "POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fecha Emisión Póliza Seg. Desgravamen"
      p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!POLIZA_FEMDES))
   
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nro. de Póliza (s)"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!POLIZA_NUMDES & "") & IIf(Len(Trim(g_rst_Princi!POLIZA_NUMCYG & "")) = 0, "", " / " & Trim(g_rst_Princi!POLIZA_NUMCYG & ""))
      
      p_Grid.Rows = p_Grid.Rows + 2:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fecha Emisión Póliza Seg. Vivienda"
      p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!POLIZA_FEMVIV))
   
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Nro. de Póliza"
      p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!POLIZA_NUMVIV & "")
      
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_Buscar_TraCof(p_Grid As MSFlexGrid, p_TxtObs As TextBox)
   p_TxtObs.Text = ""
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
   g_str_Parame = g_str_Parame & "EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:         p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                        p_Grid.Text = "Fecha Envío"
      p_Grid.Col = 1:                        p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECENV))
   
      If g_rst_Princi!EVACOF_FECREC > 0 Then
         If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"003"
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Nro. Operación Mivivienda"
            p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
            
            p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
            p_Grid.Col = 0:                  p_Grid.Text = "Fecha Aprobación Mivivienda"
            p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_APRMVI))
         End If
      
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Nro. Carta COFIDE"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!EVACOF_NUMCAR & "")
         
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Fecha Recepción Carta COFIDE"
         p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECREC))
         
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Nro. Operación COFIDE"
         p_Grid.Col = 1:                     p_Grid.Text = Trim(g_rst_Princi!EVACOF_CODMVI & "")
      
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Fecha Desembolso COFIDE"
         p_Grid.Col = 1:                     p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
         
         p_Grid.Rows = p_Grid.Rows + 1:      p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                     p_Grid.Text = "Importe Desembolsado"
         p_Grid.Col = 1:                     p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:            p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACOF_MTODES, 12, 2)
         
         p_TxtObs.Text = Trim(g_rst_Princi!EVACOF_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_TraMVi(p_Grid As MSFlexGrid, p_TxtObs As TextBox)
   p_TxtObs.Text = ""
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM TRA_EVAMVI WHERE "
   g_str_Parame = g_str_Parame & "EVAMVI_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                  p_Grid.Text = "Fecha Envío"
      p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVAMVI_FECENV))
   
      If g_rst_Princi!EVAMVI_FECREC > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Fecha de Recepción"
         p_Grid.Col = 1:                  p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVAMVI_FECREC))
      
         p_Grid.Rows = p_Grid.Rows + 1:   p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                  p_Grid.Text = "Nro. Expediente Mivivienda"
         p_Grid.Col = 1:                  p_Grid.Text = Trim(g_rst_Princi!EVAMVI_CODMVI & "")
      
         p_TxtObs.Text = Trim(g_rst_Princi!EVAMVI_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(p_Grid)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Sub modmip_gs_Carga_PryInm_Combo(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipPry As Integer)
   ReDim p_Arregl(0)
   p_Combo.Clear
      
   g_str_Parame = "SELECT * FROM PRY_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = " & CStr(p_TipPry) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_TITULO ASC "
   
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
      p_Combo.AddItem Trim(g_rst_Listas!DATGEN_TITULO)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!DATGEN_TITULO)
            
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub modmip_gs_Rpt_EvaIns(ByVal p_NomRpt As String, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIns As Integer, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_NomRpt    -  Nombre Reporte
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_CodIns    - Código de Instancia
'                 p_CodIte    - Código de Item o Detalle (Para Tipo de Proyecto Inmobiliario)
'F. Creación   :  02-03-2010
'Objetivo      :  Permite llenar la Base de Datos de Solicitudes en Trámite por Instancia de Evaluación
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   
   Screen.MousePointer = 11
   
   g_str_Parame = "DELETE FROM RPT_EVAHIP WHERE EVAHIP_NOMTER = '" & modgen_g_str_NombPC & "' AND EVAHIP_NOMRPT = '" & p_NomRpt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
         
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   Do While Not r_rst_SolMae.EOF
      r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
   
      'Abriendo Cursor para Proyecto
      r_int_FlgPry = 0
      r_str_CodPry = ""
      r_str_NomPry = ""
      
      r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
         MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
         r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
      
         If Len(Trim(p_CodIte)) > 0 Then
            If r_rst_SolInm!SOLINM_PRYCOD = p_CodIte Then
               r_int_FlgPry = 1
            End If
         Else
            r_int_FlgPry = 1
         End If
      Else
         r_str_CodMod = "SIN INMUEBLE"
         r_int_FlgPry = 1
      End If
      
      r_rst_SolInm.Close
      Set r_rst_SolInm = Nothing
      
      If r_int_FlgPry = 1 Then
         'Para obtener Nombre de Proyecto
         If Len(Trim(r_str_CodPry)) > 0 Then
            r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
         
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
               MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
               r_rst_Proyec.MoveFirst
               
               r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
            End If
         
            r_rst_Proyec.Close
            Set r_rst_Proyec = Nothing
         End If
      
         'Abriendo Cursor para Seguimiento
         r_str_Cadena = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = " & CStr(p_CodIns)
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Seguim, 3) Then
            MsgBox "Error al leer tabla TRA_SEGUIM.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
         'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
         r_str_DetObs = ""
         r_int_TpoObs = 0
         
         r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODOCU = 21"
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
            MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
            r_rst_SegDet.MoveFirst
            
            Do While Not r_rst_SegDet.EOF
               If r_rst_SegDet!SEGFECACT > 0 Then
                  r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               Else
                  r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  
                  r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
               End If
            
               r_rst_SegDet.MoveNext
            Loop
         End If
      
         r_rst_SegDet.Close
         Set r_rst_SegDet = Nothing
      
         'Calculando Tiempo en Instancia
         r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI))))
         
         'Insertando Registros
         g_str_Parame = "INSERT INTO RPT_EVAHIP ("
         g_str_Parame = g_str_Parame & "EVAHIP_NOMTER, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMRPT, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMITE, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRODUC, "
         g_str_Parame = g_str_Parame & "EVAHIP_DOCIDE, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMCLI, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECINI, "
         g_str_Parame = g_str_Parame & "EVAHIP_SITINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_RESDES, "
         g_str_Parame = g_str_Parame & "EVAHIP_TIPEVA, "
         g_str_Parame = g_str_Parame & "EVAHIP_MONPRE, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APOSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRESOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APODOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PREDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOTRA, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_CONHIP, "
         g_str_Parame = g_str_Parame & "EVAHIP_CODMOD, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMPRY) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("002", CStr(p_CodIns)) & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("023", CStr(r_rst_Seguim!SEGUIM_SITUAC)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("004", CStr(r_rst_SolMae!SOLMAE_SITINS)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON)) & "', "
      
         If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         Else
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_DOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_DOL) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
         End If
         
         g_str_Parame = g_str_Parame & CStr(r_int_TpoIns) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_TpoObs) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_DetObs & "', "
         g_str_Parame = g_str_Parame & "'" & r_rst_SolMae!SOLMAE_CONHIP & "', "
         
         g_str_Parame = g_str_Parame & "'" & r_str_CodMod & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NomPry & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            MsgBox "Error al insertar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
         r_rst_Seguim.Close
         Set r_rst_Seguim = Nothing
      End If
      
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
End Sub

Public Sub modmip_gs_Rpt_EvaIns_Dbl(ByVal p_NomRpt As String, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIns As Integer, ByVal p_InsSec As Integer, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_NomRpt    -  Nombre Reporte
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_CodIns    - Código de Instancia
'                 p_InsSec    - Código de Instancia Secundaria
'                 p_CodIte    - Código de Item o Detalle (Para Tipo de Proyecto Inmobiliario)
'F. Creación   :  02-03-2010
'Objetivo      :  Permite llenar la Base de Datos de Solicitudes en Trámite por Instancia de Evaluación
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   Dim r_rst_SegPri     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   
   Screen.MousePointer = 11
   
   g_str_Parame = "DELETE FROM RPT_EVAHIP WHERE EVAHIP_NOMTER = '" & modgen_g_str_NombPC & "' AND EVAHIP_NOMRPT = '" & p_NomRpt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
         
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   Do While Not r_rst_SolMae.EOF
      'Abriendo Cursor para Seguimiento Principal
      r_str_Cadena = "SELECT * FROM TRA_SEGUIM WHERE "
      r_str_Cadena = r_str_Cadena & "SEGUIM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
      r_str_Cadena = r_str_Cadena & "SEGUIM_CODINS = " & CStr(p_InsSec) & " AND "
      r_str_Cadena = r_str_Cadena & "(SEGUIM_SITUAC = 9 OR SEGUIM_SITUAC = 3)"
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegPri, 3) Then
         MsgBox "Error al leer tabla TRA_SEGUIM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegPri.BOF And r_rst_SegPri.EOF) Then
         r_rst_SegPri.MoveFirst
         r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
      
         'Abriendo Cursor para Proyecto
         r_int_FlgPry = 0
         r_str_CodPry = ""
         r_str_NomPry = ""
         r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
            MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
            r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
            If Len(Trim(p_CodIte)) > 0 Then
               If r_rst_SolInm!SOLINM_PRYCOD = p_CodIte Then
                  r_int_FlgPry = 1
               End If
            Else
               r_int_FlgPry = 1
            End If
         Else
            r_str_CodMod = "SIN INMUEBLE"
            r_int_FlgPry = 1
         End If
         
         r_rst_SolInm.Close
         Set r_rst_SolInm = Nothing
         
         If r_int_FlgPry = 1 Then
            'Para obtener Nombre de Proyecto
            If Len(Trim(r_str_CodPry)) > 0 Then
               r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
            
               If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
                  MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
                  Exit Sub
               End If
               If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
                  r_rst_Proyec.MoveFirst
                  r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
               End If
            
               r_rst_Proyec.Close
               Set r_rst_Proyec = Nothing
            End If
         
            'Abriendo Cursor para Seguimiento
            r_str_Cadena = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = " & CStr(p_InsSec)
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Seguim, 3) Then
               MsgBox "Error al leer tabla TRA_SEGUIM.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
            r_str_DetObs = ""
            r_int_TpoObs = 0
            
            r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
            r_str_Cadena = r_str_Cadena & "SEGDET_CODINS = " & CStr(p_InsSec) & " AND "
            r_str_Cadena = r_str_Cadena & "SEGDET_CODOCU = 21"
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
               MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
               r_rst_SegDet.MoveFirst
               
               Do While Not r_rst_SegDet.EOF
                  If r_rst_SegDet!SEGFECACT > 0 Then
                     r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  Else
                     r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                     r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
                  End If
               
                  r_rst_SegDet.MoveNext
               Loop
            End If
         
            r_rst_SegDet.Close
            Set r_rst_SegDet = Nothing
         
            'Calculando Tiempo en Instancia
            r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI))))
            
            'Insertando Registros
            g_str_Parame = "INSERT INTO RPT_EVAHIP ("
            g_str_Parame = g_str_Parame & "EVAHIP_NOMTER, "
            g_str_Parame = g_str_Parame & "EVAHIP_NOMRPT, "
            g_str_Parame = g_str_Parame & "EVAHIP_NUMSOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_DESINS, "
            g_str_Parame = g_str_Parame & "EVAHIP_NUMITE, "
            g_str_Parame = g_str_Parame & "EVAHIP_PRODUC, "
            g_str_Parame = g_str_Parame & "EVAHIP_DOCIDE, "
            g_str_Parame = g_str_Parame & "EVAHIP_NOMCLI, "
            g_str_Parame = g_str_Parame & "EVAHIP_FECSOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_FECINI, "
            g_str_Parame = g_str_Parame & "EVAHIP_SITINS, "
            g_str_Parame = g_str_Parame & "EVAHIP_RESDES, "
            g_str_Parame = g_str_Parame & "EVAHIP_TIPEVA, "
            g_str_Parame = g_str_Parame & "EVAHIP_MONPRE, "
            g_str_Parame = g_str_Parame & "EVAHIP_CVTSOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_APOSOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_PRESOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_CVTDOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_APODOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_PREDOL, "
            g_str_Parame = g_str_Parame & "EVAHIP_TPOTRA, "
            g_str_Parame = g_str_Parame & "EVAHIP_TPOOBS, "
            g_str_Parame = g_str_Parame & "EVAHIP_DESOBS, "
            g_str_Parame = g_str_Parame & "EVAHIP_CONHIP, "
            g_str_Parame = g_str_Parame & "EVAHIP_CODMOD, "
            g_str_Parame = g_str_Parame & "EVAHIP_NOMPRY) "
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
            g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO) & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("002", CStr(p_InsSec)) & "', "
            g_str_Parame = g_str_Parame & "1, "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD) & "', "
            g_str_Parame = g_str_Parame & "'" & CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO) & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO) & "', "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)) & "', "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI)) & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("023", CStr(r_rst_Seguim!SEGUIM_SITUAC)) & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("004", CStr(r_rst_SolMae!SOLMAE_SITINS)) & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA)) & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON)) & "', "
         
            If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
               g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_SOL) & ", "
               g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_SOL) & ", "
               g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
            Else
               g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_DOL) & ", "
               g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_DOL) & ", "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & "0, "
               g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
            End If
            
            g_str_Parame = g_str_Parame & CStr(r_int_TpoIns) & ", "
            g_str_Parame = g_str_Parame & CStr(r_int_TpoObs) & ", "
            g_str_Parame = g_str_Parame & "'" & r_str_DetObs & "', "
            g_str_Parame = g_str_Parame & "'" & r_rst_SolMae!SOLMAE_CONHIP & "', "
            
            g_str_Parame = g_str_Parame & "'" & r_str_CodMod & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_NomPry & "') "
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               MsgBox "Error al insertar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            r_rst_Seguim.Close
            Set r_rst_Seguim = Nothing
         End If
      End If
      
      r_rst_SegPri.Close
      Set r_rst_SegPri = Nothing
      
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
End Sub

Public Sub modmip_gs_Rpt_EvaObs(ByVal p_NomRpt As String, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIns As Integer, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_NomRpt    -  Nombre Reporte
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_CodIns    - Código de Instancia
'                 p_CodIte    - Código de Item o Detalle (Para Tipo de Proyecto Inmobiliario)
'F. Creación   :  02-03-2010
'Objetivo      :  Permite llenar la Base de Datos de Solicitudes Observadas por Instancia
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   
   Screen.MousePointer = 11
   
   g_str_Parame = "DELETE FROM RPT_EVAHIP WHERE EVAHIP_NOMTER = '" & modgen_g_str_NombPC & "' AND EVAHIP_NOMRPT = '" & p_NomRpt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGUIM B "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = SEGUIM_NUMSOL AND "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
         
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_SITUAC = 3"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   Do While Not r_rst_SolMae.EOF
      r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
   
      'Abriendo Cursor para Proyecto
      r_int_FlgPry = 0
      r_str_CodPry = ""
      r_str_NomPry = ""
      
      r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
         MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
         r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
      
         If Len(Trim(p_CodIte)) > 0 Then
            If r_rst_SolInm!SOLINM_PRYCOD = p_CodIte Then
               r_int_FlgPry = 1
            End If
         Else
            r_int_FlgPry = 1
         End If
      Else
         r_str_CodMod = "SIN INMUEBLE"
         r_int_FlgPry = 1
      End If
      
      r_rst_SolInm.Close
      Set r_rst_SolInm = Nothing
      
      If r_int_FlgPry = 1 Then
         'Para obtener Nombre de Proyecto
         If Len(Trim(r_str_CodPry)) > 0 Then
            r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
         
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
               MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
               r_rst_Proyec.MoveFirst
               r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
            End If
         
            r_rst_Proyec.Close
            Set r_rst_Proyec = Nothing
         End If
      
         'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
         r_str_DetObs = ""
         r_int_TpoObs = 0
         
         r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODOCU = 21"
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
            MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
            r_rst_SegDet.MoveFirst
            
            Do While Not r_rst_SegDet.EOF
               If r_rst_SegDet!SEGFECACT > 0 Then
                  r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               Else
                  r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
               End If
            
               r_rst_SegDet.MoveNext
            Loop
         End If
      
         r_rst_SegDet.Close
         Set r_rst_SegDet = Nothing
      
         'Calculando Tiempo en Instancia
         r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI))))
         
         'Insertando Registros
         g_str_Parame = "INSERT INTO RPT_EVAHIP ("
         g_str_Parame = g_str_Parame & "EVAHIP_NOMTER, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMRPT, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMITE, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRODUC, "
         g_str_Parame = g_str_Parame & "EVAHIP_DOCIDE, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMCLI, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECINI, "
         g_str_Parame = g_str_Parame & "EVAHIP_SITINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_RESDES, "
         g_str_Parame = g_str_Parame & "EVAHIP_TIPEVA, "
         g_str_Parame = g_str_Parame & "EVAHIP_MONPRE, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APOSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRESOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APODOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PREDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOTRA, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_CONHIP, "
         g_str_Parame = g_str_Parame & "EVAHIP_CODMOD, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMPRY) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("002", CStr(p_CodIns)) & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("023", CStr(r_rst_SolMae!SEGUIM_SITUAC)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("004", CStr(r_rst_SolMae!SOLMAE_SITINS)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON)) & "', "
      
         If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         Else
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_DOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_DOL) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
         End If
         
         g_str_Parame = g_str_Parame & CStr(r_int_TpoIns) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_TpoObs) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_DetObs & "', "
         g_str_Parame = g_str_Parame & "'" & r_rst_SolMae!SOLMAE_CONHIP & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodMod & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NomPry & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            MsgBox "Error al insertar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
End Sub

Public Sub modmip_gs_Rpt_EvaObs_Dbl(ByVal p_NomRpt As String, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIns As Integer, ByVal p_InsSec As Integer, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_NomRpt    -  Nombre Reporte
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_CodIns    - Código de Instancia
'                 p_InsSec    - Código de Instancia Secundaria
'                 p_CodIte    - Código de Item o Detalle (Para Tipo de Proyecto Inmobiliario)
'F. Creación   :  02-03-2010
'Objetivo      :  Permite llenar la Base de Datos de Solicitudes Observadas por Instancia
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   
   Screen.MousePointer = 11
   
   g_str_Parame = "DELETE FROM RPT_EVAHIP WHERE EVAHIP_NOMTER = '" & modgen_g_str_NombPC & "' AND EVAHIP_NOMRPT = '" & p_NomRpt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGUIM B "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = SEGUIM_NUMSOL AND "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
         
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_InsSec) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_SITUAC = 3"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   Do While Not r_rst_SolMae.EOF
      r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
      r_int_FlgPry = 0
      r_str_CodPry = ""
      r_str_NomPry = ""
      
      'Abriendo Cursor para Proyecto
      r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
         MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
         r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
      
         If Len(Trim(p_CodIte)) > 0 Then
            If r_rst_SolInm!SOLINM_PRYCOD = p_CodIte Then
               r_int_FlgPry = 1
            End If
         Else
            r_int_FlgPry = 1
         End If
      Else
         r_str_CodMod = "SIN INMUEBLE"
         r_int_FlgPry = 1
      End If
      
      r_rst_SolInm.Close
      Set r_rst_SolInm = Nothing
      
      If r_int_FlgPry = 1 Then
         'Para obtener Nombre de Proyecto
         If Len(Trim(r_str_CodPry)) > 0 Then
            r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
         
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
               MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
               r_rst_Proyec.MoveFirst
               r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
            End If
         
            r_rst_Proyec.Close
            Set r_rst_Proyec = Nothing
         End If
      
         'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
         r_str_DetObs = ""
         r_int_TpoObs = 0
         
         r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODINS = " & CStr(p_InsSec) & " AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODOCU = 21"
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
            MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
            r_rst_SegDet.MoveFirst
            
            Do While Not r_rst_SegDet.EOF
               If r_rst_SegDet!SEGFECACT > 0 Then
                  r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               Else
                  r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
               End If
            
               r_rst_SegDet.MoveNext
            Loop
         End If
      
         r_rst_SegDet.Close
         Set r_rst_SegDet = Nothing
      
         'Calculando Tiempo en Instancia
         r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI))))
         
         'Insertando Registros
         g_str_Parame = "INSERT INTO RPT_EVAHIP ("
         g_str_Parame = g_str_Parame & "EVAHIP_NOMTER, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMRPT, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMITE, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRODUC, "
         g_str_Parame = g_str_Parame & "EVAHIP_DOCIDE, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMCLI, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECINI, "
         g_str_Parame = g_str_Parame & "EVAHIP_SITINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_RESDES, "
         g_str_Parame = g_str_Parame & "EVAHIP_TIPEVA, "
         g_str_Parame = g_str_Parame & "EVAHIP_MONPRE, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APOSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRESOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APODOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PREDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOTRA, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_CONHIP, "
         g_str_Parame = g_str_Parame & "EVAHIP_CODMOD, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMPRY) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("002", CStr(p_InsSec)) & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("023", CStr(r_rst_SolMae!SEGUIM_SITUAC)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("004", CStr(r_rst_SolMae!SOLMAE_SITINS)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON)) & "', "
      
         If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         Else
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_DOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_DOL) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
         End If
         
         g_str_Parame = g_str_Parame & CStr(r_int_TpoIns) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_TpoObs) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_DetObs & "', "
         g_str_Parame = g_str_Parame & "'" & r_rst_SolMae!SOLMAE_CONHIP & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodMod & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NomPry & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            MsgBox "Error al insertar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
End Sub

Public Sub modmip_gs_Rpt_AprCon(ByVal p_NomRpt As String, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIns As Integer, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_NomRpt    -  Nombre Reporte
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_CodIns    - Código de Instancia
'                 p_CodIte    - Código de Item o Detalle (Para Tipo de Proyecto Inmobiliario)
'F. Creación   :  02-03-2010
'Objetivo      :  Permite llenar la Base de Datos de Solicitudes Observadas por Instancia
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_SegCon     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_str_FecCon     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   
   Screen.MousePointer = 11
   
   g_str_Parame = "DELETE FROM RPT_EVAHIP WHERE EVAHIP_NOMTER = '" & modgen_g_str_NombPC & "' AND EVAHIP_NOMRPT = '" & p_NomRpt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGCON B "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = SEGCON_NUMSOL AND "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
         
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SEGCON_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGCON_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   Do While Not r_rst_SolMae.EOF
      r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
   
      'Abriendo Cursor para Proyecto
      r_int_FlgPry = 0
      r_str_CodPry = ""
      r_str_NomPry = ""
      
      r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
         MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
         r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
      
         If Len(Trim(p_CodIte)) > 0 Then
            If r_rst_SolInm!SOLINM_PRYCOD = p_CodIte Then
               r_int_FlgPry = 1
            End If
         Else
            r_int_FlgPry = 1
         End If
      Else
         r_str_CodMod = "SIN INMUEBLE"
         r_int_FlgPry = 1
      End If
      
      r_rst_SolInm.Close
      Set r_rst_SolInm = Nothing
      
      If r_int_FlgPry = 1 Then
         'Para obtener Nombre de Proyecto
         If Len(Trim(r_str_CodPry)) > 0 Then
            r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
         
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
               MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
               r_rst_Proyec.MoveFirst
               r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
            End If
         
            r_rst_Proyec.Close
            Set r_rst_Proyec = Nothing
         End If
      
         'Abriendo Cursor para Detalle de Aprobación Condicionada
         r_str_FecCon = ""
         r_str_DetObs = ""
         r_int_TpoObs = 0
         
         r_str_Cadena = "SELECT * FROM TRA_SEGCON WHERE SEGCON_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
         r_str_Cadena = r_str_Cadena & "SEGCON_CODINS = " & CStr(p_CodIns) & " "
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegCon, 3) Then
            MsgBox "Error al leer tabla TRA_SEGCON.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegCon.BOF And r_rst_SegCon.EOF) Then
            r_rst_SegCon.MoveFirst
            r_str_FecCon = gf_FormatoFecha(CStr(r_rst_SegCon!SEGFECCRE))
            r_int_TpoObs = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegCon!SEGFECCRE))))
            r_str_DetObs = Trim(r_rst_SegCon!SEGCON_OBSCON & "")
         End If
      
         r_rst_SegCon.Close
         Set r_rst_SegCon = Nothing
      
         'Insertando Registros
         g_str_Parame = "INSERT INTO RPT_EVAHIP ("
         g_str_Parame = g_str_Parame & "EVAHIP_NOMTER, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMRPT, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMITE, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRODUC, "
         g_str_Parame = g_str_Parame & "EVAHIP_DOCIDE, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMCLI, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECINI, "
         g_str_Parame = g_str_Parame & "EVAHIP_TIPEVA, "
         g_str_Parame = g_str_Parame & "EVAHIP_MONPRE, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APOSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRESOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APODOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PREDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_CONHIP, "
         g_str_Parame = g_str_Parame & "EVAHIP_CODMOD, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMPRY) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("002", CStr(p_CodIns)) & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)) & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_FecCon & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON)) & "', "
      
         If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         Else
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_DOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_DOL) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
         End If
         
         g_str_Parame = g_str_Parame & CStr(r_int_TpoObs) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_DetObs & "', "
         g_str_Parame = g_str_Parame & "'" & r_rst_SolMae!SOLMAE_CONHIP & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodMod & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NomPry & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            MsgBox "Error al insertar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
End Sub

Public Sub modmip_gs_Rpt_SolEva(ByVal p_NomRpt As String, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIns As Integer, ByVal p_CodIte As String, ByVal p_Situac As Integer, ByVal p_FecIni As String, ByVal p_FecFin As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_NomRpt    -  Nombre Reporte
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                                4 - Por Tipo de Evaluación
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_CodIns    - Código de Instancia
'                 p_CodIte    - Código de Item o Detalle (Para Tipo de Proyecto Inmobiliario)
'                 p_Situac    - Código de Resultado en Instancia
'                               0 - Todos los Resultados
'                               1 - Aprobadas
'                               2 - Rechazadas
'F. Creación   :  02-03-2010
'Objetivo      :  Permite llenar la Base de Datos de Solicitudes Observadas por Instancia
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_str_FecCon     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   
   Screen.MousePointer = 11
   
   g_str_Parame = "DELETE FROM RPT_EVAHIP WHERE EVAHIP_NOMTER = '" & modgen_g_str_NombPC & "' AND EVAHIP_NOMRPT = '" & p_NomRpt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al borrar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGUIM B WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SEGUIM_NUMSOL AND "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
         
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & p_CodFil & "' AND "
         End If
   
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_TIPEVA = " & p_CodFil & " AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_FECFIN >= " & Format(CDate(p_FecIni), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_FECFIN <= " & Format(CDate(p_FecFin), "yyyymmdd") & " AND "
   
   If p_Situac > 0 Then
      g_str_Parame = g_str_Parame & "SEGUIM_SITUAC = " & CStr(p_Situac)
   Else
      g_str_Parame = g_str_Parame & "(SEGUIM_SITUAC = 1 OR SEGUIM_SITUAC = 2)"
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   Do While Not r_rst_SolMae.EOF
      r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
   
      'Abriendo Cursor para Proyecto
      r_int_FlgPry = 0
      r_str_CodPry = ""
      r_str_NomPry = ""
      
      r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
         MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
         r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
      
         If Len(Trim(p_CodIte)) > 0 Then
            If r_rst_SolInm!SOLINM_PRYCOD = p_CodIte Then
               r_int_FlgPry = 1
            End If
         Else
            r_int_FlgPry = 1
         End If
      Else
         r_str_CodMod = "SIN INMUEBLE"
         r_int_FlgPry = 1
      End If
      
      r_rst_SolInm.Close
      Set r_rst_SolInm = Nothing
      
      If r_int_FlgPry = 1 Then
         'Para obtener Nombre de Proyecto
         If Len(Trim(r_str_CodPry)) > 0 Then
            r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
         
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
               MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
               r_rst_Proyec.MoveFirst
               
               r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
            End If
         
            r_rst_Proyec.Close
            Set r_rst_Proyec = Nothing
         End If
      
         'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
         r_str_DetObs = ""
         r_int_TpoObs = 0
         
         r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
         r_str_Cadena = r_str_Cadena & "SEGDET_CODOCU = 21"
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
            MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
            r_rst_SegDet.MoveFirst
            
            Do While Not r_rst_SegDet.EOF
               If r_rst_SegDet!SEGFECACT > 0 Then
                  r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               Else
                  r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
               End If
            
               r_rst_SegDet.MoveNext
            Loop
         End If
      
         r_rst_SegDet.Close
         Set r_rst_SegDet = Nothing
      
         'Abriendo Cursor en Seguimiento para obtener Motivo de Rechazo
         r_str_DetObs = ""
         
         If r_rst_SolMae!SEGUIM_SITUAC = 2 Then
            r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
            r_str_Cadena = r_str_Cadena & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
            r_str_Cadena = r_str_Cadena & "SEGDET_CODOCU = 13 "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
               MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
               r_rst_SegDet.MoveFirst
               r_str_DetObs = moddat_gf_Consulta_ParDes("003", CStr(r_rst_SegDet!SEGDET_MOTREC)) & Chr(10) & Chr(13) & Trim(r_rst_SegDet!SEGDET_OBSERV & "")
            End If
         
            r_rst_SegDet.Close
            Set r_rst_SegDet = Nothing
         End If
      
         'Insertando Registros
         g_str_Parame = "INSERT INTO RPT_EVAHIP ("
         g_str_Parame = g_str_Parame & "EVAHIP_NOMTER, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMRPT, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESINS, "
         g_str_Parame = g_str_Parame & "EVAHIP_NUMITE, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRODUC, "
         g_str_Parame = g_str_Parame & "EVAHIP_DOCIDE, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMCLI, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECINI, "
         g_str_Parame = g_str_Parame & "EVAHIP_FECFIN, "
         g_str_Parame = g_str_Parame & "EVAHIP_TIPEVA, "
         g_str_Parame = g_str_Parame & "EVAHIP_MONPRE, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APOSOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PRESOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_CVTDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_APODOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_PREDOL, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOTRA, "
         g_str_Parame = g_str_Parame & "EVAHIP_TPOOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_DESOBS, "
         g_str_Parame = g_str_Parame & "EVAHIP_CONHIP, "
         g_str_Parame = g_str_Parame & "EVAHIP_CODMOD, "
         g_str_Parame = g_str_Parame & "EVAHIP_NOMPRY, "
         g_str_Parame = g_str_Parame & "EVAHIP_SITINS) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & p_NomRpt & "', "
         g_str_Parame = g_str_Parame & "'" & gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("002", CStr(p_CodIns)) & "', "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)) & "', "
         g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA)) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON)) & "', "
      
         If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_SOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         Else
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_COMVTA_DOL) & ", "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_APOPRO_DOL) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & CStr(r_rst_SolMae!SOLMAE_MTOPRE_MPR) & ", "
         End If
         
         g_str_Parame = g_str_Parame & CStr(CInt(CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI))))) & ", "
         g_str_Parame = g_str_Parame & CStr(r_int_TpoObs) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_DetObs & "', "
         g_str_Parame = g_str_Parame & "'" & r_rst_SolMae!SOLMAE_CONHIP & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_CodMod & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_NomPry & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("023", CStr(r_rst_SolMae!SEGUIM_SITUAC)) & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            MsgBox "Error al insertar registro en tabla RPT_EVAHIP.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
End Sub

Public Sub modmip_gs_Exc_Tramit(ByVal p_CodIns As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodIns    -  Código de Instancia
'F. Creación   :  02-03-2010
'Objetivo      :  Permite exportar la información de una Instancia de Evaluación ordenado por Consejero Hipotecario
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   Dim r_str_PrySel     As String

   r_str_PrySel = ""
   
   If p_TipFil = 3 Then
      If Len(Trim(p_CodIte)) > 0 Then
         r_str_PrySel = p_CodIte
      End If
   End If

   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
      
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
   
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & Format(p_CodFil, "00") & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
      .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "DOC. IDENTIDAD":                .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NOMBRE CLIENTE":                .Columns("E").ColumnWidth = 40
      .Cells(1, 6) = "F. SOLICITUD":                  .Columns("F").ColumnWidth = 15:        .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "F. INGRESO INSTANCIA":          .Columns("G").ColumnWidth = 15:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "TIPO EVALUACION":               .Columns("H").ColumnWidth = 40
      .Cells(1, 9) = "CONSEJERO HIPOTECARIO":         .Columns("I").ColumnWidth = 30:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "INSTANCIA":                    .Columns("J").ColumnWidth = 30:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "TIEMPO INSTANCIA":             .Columns("K").ColumnWidth = 20
      .Cells(1, 12) = "TIEMPO OBSERVADO":             .Columns("L").ColumnWidth = 20
      .Cells(1, 13) = "TIEMPO EVALUACION":            .Columns("M").ColumnWidth = 20
      .Cells(1, 14) = "MONEDA":                       .Columns("N").ColumnWidth = 15:        .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 15) = "V. INMUEBLE":                  .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "CUOTA INICIAL":                .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "PORC. INICIAL":                .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "MTO. CREDITO S/.":             .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "MTO. CREDITO US$":             .Columns("S").ColumnWidth = 20
      .Cells(1, 20) = "SITUACION INSTANCIA":          .Columns("T").ColumnWidth = 30
      .Cells(1, 21) = "MODALIDAD":                    .Columns("U").ColumnWidth = 60
      .Cells(1, 22) = "PROYECTO INMOBILIARIO":        .Columns("V").ColumnWidth = 120
      .Cells(1, 23) = "OBSERVACION":                  .Columns("W").ColumnWidth = 200
      
      .Range(.Cells(1, 1), .Cells(1, 23)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 23)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   
   Do While Not r_rst_SolMae.EOF
      r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
   
      'Abriendo Cursor para Proyecto
      r_int_FlgPry = 0
      r_str_CodPry = ""
      r_str_NomPry = ""
      
      r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
         MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
         r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
      
         If Len(Trim(r_str_PrySel)) > 0 Then
            If r_str_CodPry = r_str_PrySel Then
               r_int_FlgPry = 1
            End If
         Else
            r_int_FlgPry = 1
         End If
      Else
         r_str_CodMod = "SIN INMUEBLE"
         r_int_FlgPry = 1
      End If
      
      r_rst_SolInm.Close
      Set r_rst_SolInm = Nothing
      
      If r_int_FlgPry = 1 Then
         'Para obtener Nombre de Proyecto
         If Len(Trim(r_str_CodPry)) > 0 Then
            r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
         
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
               MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
               r_rst_Proyec.MoveFirst
               r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
            End If
         
            r_rst_Proyec.Close
            Set r_rst_Proyec = Nothing
         End If
      
         'Abriendo Cursor para Seguimiento
         r_str_Cadena = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = " & CStr(p_CodIns)
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Seguim, 3) Then
            MsgBox "Error al leer tabla TRA_SEGUIM.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      
         'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
         r_str_DetObs = ""
         r_int_TpoObs = 0
         
         r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGDET_CODINS = " & CStr(p_CodIns) & " AND SEGDET_CODOCU = 21"
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
            MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
            r_rst_SegDet.MoveFirst
            
            Do While Not r_rst_SegDet.EOF
               If r_rst_SegDet!SEGFECACT > 0 Then
                  r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               Else
                  r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
               End If
            
               r_rst_SegDet.MoveNext
            Loop
         End If
      
         r_rst_SegDet.Close
         Set r_rst_SegDet = Nothing
      
         'Calculando Tiempo en Instancia
         r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI))))
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = moddat_gf_Consulta_ParDes("002", CStr(p_CodIns))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_int_TpoIns
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_int_TpoObs
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = r_int_TpoIns - r_int_TpoObs
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON))
         
         If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = moddat_gf_Consulta_ParDes("023", CStr(r_rst_Seguim!SEGUIM_SITUAC))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = r_str_CodMod
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = r_str_NomPry
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = r_str_DetObs
      
         r_rst_Seguim.Close
         Set r_rst_Seguim = Nothing
      
         r_int_ConVer = r_int_ConVer + 1
      End If
   
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   Select Case p_TipFil
      Case 1
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 23)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      
      Case 2
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 23)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      
      Case 3
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 23)).Sort r_obj_Excel.Range("V1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Exc_Tramit_Dbl(ByVal p_CodIns As Integer, ByVal p_InsSec As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodIns    -  Código de Instancia
'F. Creación   :  02-03-2010
'Objetivo      :  Permite exportar la información de una Instancia de Evaluación ordenado por Consejero Hipotecario
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_rst_SolInm     As ADODB.Recordset
   Dim r_rst_Proyec     As ADODB.Recordset
   Dim r_rst_SegPri     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   Dim r_int_FlgPry     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodMod     As String
   Dim r_str_PrySel     As String

   r_str_PrySel = ""
   
   If p_TipFil = 3 Then
      If Len(Trim(p_CodIte)) > 0 Then
         r_str_PrySel = p_CodIte
      End If
   End If

   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
      
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
   
      Case 3
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODMOD = '" & Format(p_CodFil, "00") & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
      .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "DOC. IDENTIDAD":                .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NOMBRE CLIENTE":                .Columns("E").ColumnWidth = 40
      .Cells(1, 6) = "F. SOLICITUD":                  .Columns("F").ColumnWidth = 15:        .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "F. INGRESO INSTANCIA":          .Columns("G").ColumnWidth = 15:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "TIPO EVALUACION":               .Columns("H").ColumnWidth = 40
      .Cells(1, 9) = "CONSEJERO HIPOTECARIO":         .Columns("I").ColumnWidth = 30:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "INSTANCIA":                    .Columns("J").ColumnWidth = 30:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "TIEMPO INSTANCIA":             .Columns("K").ColumnWidth = 20
      .Cells(1, 12) = "TIEMPO OBSERVADO":             .Columns("L").ColumnWidth = 20
      .Cells(1, 13) = "TIEMPO EVALUACION":            .Columns("M").ColumnWidth = 20
      .Cells(1, 14) = "MONEDA":                       .Columns("N").ColumnWidth = 15:        .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 15) = "V. INMUEBLE":                  .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "CUOTA INICIAL":                .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "PORC. INICIAL":                .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "MTO. CREDITO S/.":             .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "MTO. CREDITO US$":             .Columns("S").ColumnWidth = 20
      .Cells(1, 20) = "SITUACION INSTANCIA":          .Columns("T").ColumnWidth = 30
      .Cells(1, 21) = "MODALIDAD":                    .Columns("U").ColumnWidth = 60
      .Cells(1, 22) = "PROYECTO INMOBILIARIO":        .Columns("V").ColumnWidth = 120
      .Cells(1, 23) = "OBSERVACION":                  .Columns("W").ColumnWidth = 200
      
      .Range(.Cells(1, 1), .Cells(1, 23)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 23)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   
   Do While Not r_rst_SolMae.EOF
      'Abriendo Cursor para Seguimiento Principal
      r_str_Cadena = "SELECT * FROM TRA_SEGUIM WHERE "
      r_str_Cadena = r_str_Cadena & "SEGUIM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
      r_str_Cadena = r_str_Cadena & "SEGUIM_CODINS = " & CStr(p_InsSec) & " AND "
      r_str_Cadena = r_str_Cadena & "(SEGUIM_SITUAC = 9 OR SEGUIM_SITUAC = 3)"
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegPri, 3) Then
         MsgBox "Error al leer tabla TRA_SEGUIM.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegPri.BOF And r_rst_SegPri.EOF) Then
         r_rst_SegPri.MoveFirst
         r_str_CodMod = moddat_gf_Consulta_ParDes("272", r_rst_SolMae!SOLMAE_CODMOD & "")
      
         'Abriendo Cursor para Proyecto
         r_int_FlgPry = 0
         r_str_CodPry = ""
         r_str_NomPry = ""
         
         r_str_Cadena = "SELECT * FROM CRE_SOLINM WHERE SOLINM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SolInm, 3) Then
            MsgBox "Error al leer tabla CRE_SOLINM.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SolInm.BOF And r_rst_SolInm.EOF) Then
            r_str_CodPry = Trim(r_rst_SolInm!SOLINM_PRYCOD & "")
         
            If Len(Trim(r_str_PrySel)) > 0 Then
               If r_str_CodPry = r_str_PrySel Then
                  r_int_FlgPry = 1
               End If
            Else
               r_int_FlgPry = 1
            End If
         Else
            r_str_CodMod = "SIN INMUEBLE"
            r_int_FlgPry = 1
         End If
         
         r_rst_SolInm.Close
         Set r_rst_SolInm = Nothing
         
         If r_int_FlgPry = 1 Then
            'Para obtener Nombre de Proyecto
            If Len(Trim(r_str_CodPry)) > 0 Then
               r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & r_str_CodPry & "' "
            
               If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
                  MsgBox "Error al leer tabla PRY_DATGEN.", vbExclamation, modgen_g_str_NomPlt
                  Exit Sub
               End If
            
               If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
                  r_rst_Proyec.MoveFirst
                  r_str_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
               End If
            
               r_rst_Proyec.Close
               Set r_rst_Proyec = Nothing
            End If
         
            'Abriendo Cursor para Seguimiento
            r_str_Cadena = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGUIM_CODINS = " & CStr(p_InsSec)
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Seguim, 3) Then
               MsgBox "Error al leer tabla TRA_SEGUIM.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
         
            'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
            r_str_DetObs = ""
            r_int_TpoObs = 0
            
            r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGDET_CODINS = " & CStr(p_InsSec) & " AND SEGDET_CODOCU = 21"
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
               MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
               r_rst_SegDet.MoveFirst
               
               Do While Not r_rst_SegDet.EOF
                  If r_rst_SegDet!SEGFECACT > 0 Then
                     r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                  Else
                     r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
                     r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
                  End If
               
                  r_rst_SegDet.MoveNext
               Loop
            End If
         
            r_rst_SegDet.Close
            Set r_rst_SegDet = Nothing
         
            'Calculando Tiempo en Instancia
            r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI))))
            
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD)
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO)
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_Seguim!SEGUIM_FECINI)))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = moddat_gf_Consulta_ParDes("002", CStr(p_CodIns))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_int_TpoIns
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_int_TpoObs
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = r_int_TpoIns - r_int_TpoObs
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON))
            
            If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
            End If
            
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = moddat_gf_Consulta_ParDes("023", CStr(r_rst_Seguim!SEGUIM_SITUAC))
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = r_str_CodMod
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = r_str_NomPry
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = r_str_DetObs
         
            r_rst_Seguim.Close
            Set r_rst_Seguim = Nothing
            r_int_ConVer = r_int_ConVer + 1
         End If
      End If
   
      r_rst_SegPri.Close
      Set r_rst_SegPri = Nothing
   
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   Select Case p_TipFil
      Case 1
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 23)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      
      Case 2
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 23)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      
      Case 3
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 23)).Sort r_obj_Excel.Range("V1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Exc_Observ(ByVal p_CodIns As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodIns    -  Código de Instancia
'                 p_CodPrd    -  Código de Producto
'F. Creación   :  02-03-2010
'Objetivo      :  Permite exportar la información de Solicitudes Observadas en una Instancia de Evaluación ordenado por Producto
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer

   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGUIM B WHERE "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SEGUIM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_SITUAC = 3 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
      .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "DOC. IDENTIDAD":                .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NOMBRE CLIENTE":                .Columns("E").ColumnWidth = 40
      .Cells(1, 6) = "F. SOLICITUD":                  .Columns("F").ColumnWidth = 15:        .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "F. INGRESO INSTANCIA.":         .Columns("G").ColumnWidth = 15:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "TIPO EVALUACION":               .Columns("H").ColumnWidth = 40
      .Cells(1, 9) = "CONSEJERO HIPOTECARIO":         .Columns("I").ColumnWidth = 30:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "INSTANCIA":                    .Columns("J").ColumnWidth = 30:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "TIEMPO INSTANCIA":             .Columns("K").ColumnWidth = 20
      .Cells(1, 12) = "TIEMPO OBSERVADO":             .Columns("L").ColumnWidth = 20
      .Cells(1, 13) = "TIEMPO EVALUACION":            .Columns("M").ColumnWidth = 20
      .Cells(1, 14) = "MONEDA":                       .Columns("N").ColumnWidth = 15:        .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 15) = "V. INMUEBLE":                  .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "CUOTA INICIAL":                .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "PORC. INICIAL":                .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "MTO. CREDITO S/.":             .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "MTO. CREDITO US$":             .Columns("S").ColumnWidth = 20
      .Cells(1, 20) = "SITUACION INSTANCIA":          .Columns("T").ColumnWidth = 30
      .Cells(1, 21) = "OBSERVACION":                  .Columns("U").ColumnWidth = 200
      
      .Range(.Cells(1, 1), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 21)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   
   Do While Not r_rst_SolMae.EOF
      'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
      r_str_DetObs = ""
      r_int_TpoObs = 0
      
      r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGDET_CODINS = " & CStr(p_CodIns) & " AND SEGDET_CODOCU = 21"
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
         MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
         r_rst_SegDet.MoveFirst
         
         Do While Not r_rst_SegDet.EOF
            If r_rst_SegDet!SEGFECACT > 0 Then
               r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
            Else
               r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
            End If
         
            r_rst_SegDet.MoveNext
         Loop
      End If
   
      r_rst_SegDet.Close
      Set r_rst_SegDet = Nothing
   
      'Calculando Tiempo en Instancia
      r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI))))
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = moddat_gf_Consulta_ParDes("002", CStr(p_CodIns))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_int_TpoIns
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = r_int_TpoIns - r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON))
   
      If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = moddat_gf_Consulta_ParDes("023", CStr(r_rst_SolMae!SEGUIM_SITUAC))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = r_str_DetObs
   
      r_int_ConVer = r_int_ConVer + 1
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   Select Case p_TipFil
      Case 1
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 21)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
         
      Case 2
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 21)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Exc_Observ_Dbl(ByVal p_CodIns As Integer, ByVal p_InsSec As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodIns    -  Código de Instancia
'                 p_InsSec    -  Código de Instancia Secundaria
'                 p_CodPrd    -  Código de Producto
'F. Creación   :  02-03-2010
'Objetivo      :  Permite exportar la información de Solicitudes Observadas en una Instancia de Evaluación ordenado por Producto
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer

   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGUIM B WHERE "
   
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
   End Select
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SEGUIM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_InsSec) & " AND "
   g_str_Parame = g_str_Parame & "SEGUIM_SITUAC = 3 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
      .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "DOC. IDENTIDAD":                .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NOMBRE CLIENTE":                .Columns("E").ColumnWidth = 40
      .Cells(1, 6) = "F. SOLICITUD":                  .Columns("F").ColumnWidth = 15:        .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "F. INGRESO INSTANCIA.":         .Columns("G").ColumnWidth = 15:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "TIPO EVALUACION":               .Columns("H").ColumnWidth = 40
      .Cells(1, 9) = "CONSEJERO HIPOTECARIO":         .Columns("I").ColumnWidth = 30:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "INSTANCIA":                    .Columns("J").ColumnWidth = 30:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "TIEMPO INSTANCIA":             .Columns("K").ColumnWidth = 20
      .Cells(1, 12) = "TIEMPO OBSERVADO":             .Columns("L").ColumnWidth = 20
      .Cells(1, 13) = "TIEMPO EVALUACION":            .Columns("M").ColumnWidth = 20
      .Cells(1, 14) = "MONEDA":                       .Columns("N").ColumnWidth = 15:        .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 15) = "V. INMUEBLE":                  .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "CUOTA INICIAL":                .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "PORC. INICIAL":                .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "MTO. CREDITO S/.":             .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "MTO. CREDITO US$":             .Columns("S").ColumnWidth = 20
      .Cells(1, 20) = "SITUACION INSTANCIA":          .Columns("T").ColumnWidth = 30
      .Cells(1, 21) = "OBSERVACION":                  .Columns("U").ColumnWidth = 200
      
      .Range(.Cells(1, 1), .Cells(1, 21)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 21)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   
   Do While Not r_rst_SolMae.EOF
      'Abriendo Cursos para Seguimiento detalle (Sólo Observaciones)
      r_str_DetObs = ""
      r_int_TpoObs = 0
      
      r_str_Cadena = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND SEGDET_CODINS = " & CStr(p_InsSec) & " AND SEGDET_CODOCU = 21"
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
         MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
         r_rst_SegDet.MoveFirst
         
         Do While Not r_rst_SegDet.EOF
            If r_rst_SegDet!SEGFECACT > 0 Then
               r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
            Else
               r_int_TpoObs = r_int_TpoObs + CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
            End If
         
            r_rst_SegDet.MoveNext
         Loop
      End If
   
      r_rst_SegDet.Close
      Set r_rst_SegDet = Nothing
   
      'Calculando Tiempo en Instancia
      r_int_TpoIns = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI))))
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = moddat_gf_Consulta_ParDes("002", CStr(p_InsSec))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_int_TpoIns
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = r_int_TpoIns - r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON))
   
      If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = moddat_gf_Consulta_ParDes("023", CStr(r_rst_SolMae!SEGUIM_SITUAC))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = r_str_DetObs
   
      r_int_ConVer = r_int_ConVer + 1
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   Select Case p_TipFil
      Case 1
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 21)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
         
      Case 2
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 21)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Exc_AprCon(ByVal p_CodIns As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodIns    -  Código de Instancia
'                 p_CodPrd    -  Código de Producto
'F. Creación   :  02-03-2010
'Objetivo      :  Permite exportar la información de Solicitudes con Aprobación Condicionada una Instancia de Evaluación ordenado por Producto
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_rst_SolMae     As ADODB.Recordset
Dim r_rst_SegCon     As ADODB.Recordset
Dim r_str_Cadena     As String
Dim r_str_DetObs     As String
Dim r_str_FecCon     As String
Dim r_int_TpoObs     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGCON B ON B.SEGCON_NUMSOL = A.SOLMAE_NUMERO AND B.SEGCON_CODINS = " & CStr(p_CodIns) & " AND B.SEGCON_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SUBPRD C ON C.SUBPRD_CODPRD = A.SOLMAE_CODPRD AND C.SUBPRD_CODSUB = A.SOLMAE_CODSUB "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_SITUAC = 1 "
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "   AND SOLMAE_CODPRD = '" & p_CodFil & "' "
         End If
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "   AND SOLMAE_CONHIP = '" & p_CodFil & "' "
         End If
   End Select
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
      .Cells(1, 3) = "SUB-PRODUCTO":                  .Columns("C").ColumnWidth = 70
      .Cells(1, 4) = "SOLICITUD":                     .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "DOC. IDENTIDAD":                .Columns("E").ColumnWidth = 15:        .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "NOMBRE CLIENTE":                .Columns("F").ColumnWidth = 40
      .Cells(1, 7) = "F. SOLICITUD":                  .Columns("G").ColumnWidth = 20:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "F. APROB. CONDIC.":             .Columns("H").ColumnWidth = 20:        .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 9) = "TIPO EVALUACION":               .Columns("I").ColumnWidth = 40
      .Cells(1, 10) = "CONSEJERO HIPOTECARIO":        .Columns("J").ColumnWidth = 30:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "INSTANCIA":                    .Columns("K").ColumnWidth = 30:        .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 12) = "TIEMPO APROBAC. CONDIC.":      .Columns("L").ColumnWidth = 25
      .Cells(1, 13) = "MONEDA":                       .Columns("M").ColumnWidth = 15:        .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 14) = "V. INMUEBLE":                  .Columns("N").ColumnWidth = 20
      .Cells(1, 15) = "CUOTA INICIAL":                .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "PORC. INICIAL":                .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "MTO. CREDITO S/.":             .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "MTO. CREDITO US$":             .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "CONDICIONES":                  .Columns("S").ColumnWidth = 200
      
      .Range(.Cells(1, 1), .Cells(1, 19)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 19)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   Do While Not r_rst_SolMae.EOF
      'Abriendo Cursor para Detalle de Aprobación Condicionada
      r_str_FecCon = ""
      r_str_DetObs = ""
      r_int_TpoObs = 0
      
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "SELECT * FROM TRA_SEGCON WHERE SEGCON_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      r_str_Cadena = r_str_Cadena & "   AND SEGCON_CODINS = " & CStr(p_CodIns) & " "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegCon, 3) Then
         MsgBox "Error al leer tabla TRA_SEGCON.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegCon.BOF And r_rst_SegCon.EOF) Then
         r_rst_SegCon.MoveFirst
         r_str_FecCon = gf_FormatoFecha(CStr(r_rst_SegCon!SEGFECCRE))
         r_int_TpoObs = CInt(date - CDate(gf_FormatoFecha(CStr(r_rst_SegCon!SEGFECCRE))))
         r_str_DetObs = Trim(r_rst_SegCon!SEGCON_OBSCON & "")
      End If
   
      r_rst_SegCon.Close
      Set r_rst_SegCon = Nothing
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(r_rst_SolMae!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(r_rst_SolMae!SUBPRD_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(r_str_FecCon)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = moddat_gf_Consulta_ParDes("002", CStr(p_CodIns))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON))
      If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = r_str_DetObs
   
      r_int_ConVer = r_int_ConVer + 1
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   Select Case p_TipFil
      Case 1
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 19)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      Case 2
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 19)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Exc_ExcEva(ByVal p_CodIns As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_FecIni As String, ByVal p_FecFin As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros     :
'                  p_CodIns    -  Código de Instancia
'                  p_ConHip    -  Código de Consejero Hipotecario
'F. Creación    :  02-03-2010
'Objetivo       :  Permite exportar la información de Solicitudes con Excepción en una Instancia de Evaluación ordenado por Consejero Hipotecario
'Realizado por  :  Miguel Ikehara Punk
'Actualizado por:  Luana Martinez de la Flor
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_SegExc     As ADODB.Recordset
   Dim r_rst_HipMae     As ADODB.Recordset
   Dim r_str_Cadena     As String
   Dim r_str_DetExc     As String
   Dim r_str_FecExc     As String
   Dim r_str_NumOpe     As String
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, TRA_SEGEXC B WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SEGEXC_NUMSOL AND "
   Select Case p_TipFil
      Case 1
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & p_CodFil & "' AND "
         End If
         
      Case 2
         If Len(Trim(p_CodFil)) > 0 Then
            g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & p_CodFil & "' AND "
         End If
   End Select
   g_str_Parame = g_str_Parame & "SEGEXC_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "B.SEGFECCRE >= " & p_FecIni & " AND "
   g_str_Parame = g_str_Parame & "B.SEGFECCRE <= " & p_FecFin & " "
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_CODPRD, SOLMAE_NUMERO "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      If p_TipFil = 1 Then  'LMD21102011
         MsgBox "No existe información para el consejero hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Else
         MsgBox "No existe información para el producto.", vbExclamation, modgen_g_str_NomPlt
      End If
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
      .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "OPERACION":                     .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "DOC. IDENTIDAD":                .Columns("E").ColumnWidth = 15:        .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "NOMBRE CLIENTE":                .Columns("F").ColumnWidth = 40
      .Cells(1, 7) = "F. SOLICITUD":                  .Columns("G").ColumnWidth = 20:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "F. EXCEPCION":                  .Columns("H").ColumnWidth = 20:        .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 9) = "TIPO EVALUACION":               .Columns("I").ColumnWidth = 40
      .Cells(1, 10) = "CONSEJERO HIPOTECARIO":        .Columns("J").ColumnWidth = 30:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "INSTANCIA":                    .Columns("K").ColumnWidth = 30:        .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 12) = "MONEDA":                       .Columns("L").ColumnWidth = 15:        .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 13) = "V. INMUEBLE":                  .Columns("M").ColumnWidth = 20
      .Cells(1, 14) = "CUOTA INICIAL":                .Columns("N").ColumnWidth = 20
      .Cells(1, 15) = "PORC. INICIAL":                .Columns("O").ColumnWidth = 20
      .Cells(1, 16) = "MTO. CREDITO S/.":             .Columns("P").ColumnWidth = 20
      .Cells(1, 17) = "MTO. CREDITO US$":             .Columns("Q").ColumnWidth = 20
      .Cells(1, 18) = "SITUAC. SOLIC.":               .Columns("R").ColumnWidth = 20
      .Cells(1, 19) = "AUTORIZACION":                 .Columns("S").ColumnWidth = 40
      .Cells(1, 20) = "MOTIVO EXCEPCION":             .Columns("T").ColumnWidth = 50
      .Cells(1, 21) = "DESCRIPCION EXCEPCION":        .Columns("U").ColumnWidth = 200
      
      .Range(.Cells(1, 1), .Cells(1, 20)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 20)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   
   Do While Not r_rst_SolMae.EOF
      'Abriendo Cursor para Detalle de Aprobación Condicionada
      r_str_FecExc = ""
      r_str_DetExc = ""
      
      r_str_Cadena = "SELECT * FROM TRA_SEGEXC WHERE SEGEXC_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' AND "
      r_str_Cadena = r_str_Cadena & "SEGEXC_CODINS = " & CStr(p_CodIns) & " AND "
      r_str_Cadena = r_str_Cadena & "SEGEXC_NUMEXC = " & CStr(r_rst_SolMae!SEGEXC_NUMEXC) & " "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegExc, 3) Then
         MsgBox "Error al leer tabla TRA_SEGEXC.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegExc.BOF And r_rst_SegExc.EOF) Then
         r_rst_SegExc.MoveFirst
         
         r_str_FecExc = gf_FormatoFecha(CStr(r_rst_SegExc!SEGFECCRE))
         r_str_DetExc = Trim(r_rst_SegExc!SEGEXC_DESCRI & "")
      End If
   
      r_rst_SegExc.Close
      Set r_rst_SegExc = Nothing
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = moddat_gf_Consulta_Produc(r_rst_SolMae!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
      
      'Busca numero de operacion
      r_str_Cadena = "SELECT * FROM CRE_HIPMAE WHERE HIPMAE_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "'"
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_HipMae, 3) Then
         MsgBox "Error al leer tabla CRE_HIPMAE.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
            
      r_str_NumOpe = ""
      If Not (r_rst_HipMae.BOF And r_rst_HipMae.EOF) Then
         r_rst_HipMae.MoveFirst
         r_str_NumOpe = r_rst_HipMae!HIPMAE_NUMOPE
      End If
   
      r_rst_HipMae.Close
      Set r_rst_HipMae = Nothing
      
      If r_str_NumOpe = "" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumOpe(r_str_NumOpe)
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(r_rst_SolMae!SOLMAE_TITTDO, r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(r_str_FecExc)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Consulta_ParDes("038", CStr(r_rst_SolMae!SOLMAE_TIPEVA))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = moddat_gf_Consulta_ParDes("002", CStr(p_CodIns))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = moddat_gf_Consulta_ParDes("229", CStr(r_rst_SolMae!SOLMAE_TIPMON))
      
      If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = moddat_gf_Consulta_ParDes("020", CStr(r_rst_SolMae!SOLMAE_SITUAC))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = moddat_gf_Consulta_ParDes("243", CStr(r_rst_SolMae!SEGEXC_TIPAUT))
      If IsNull(r_rst_SolMae!SEGEXC_MOTEXC) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = IIf(CStr(r_rst_SolMae!SEGEXC_MOTEXC) = 0, "", moddat_gf_Consulta_ParDes("042", CStr(r_rst_SolMae!SEGEXC_MOTEXC)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = r_str_DetExc
         
      r_int_ConVer = r_int_ConVer + 1
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   'SE COMENTA PORQUE SE ADICIONO UN ORDER BY EN LA CONSULTA INICIAL
   'Select Case p_TipFil
   '   Case 1
   '      r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 20)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   '
   '    Case 2
   '      r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 20)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   'End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Exc_EvaSol(ByVal p_CodIns As Integer, ByVal p_TipFil As Integer, ByVal p_CodFil As String, ByVal p_Situac As Integer, ByVal p_FecIni As String, ByVal p_FecFin As String, ByVal p_CodIte As String)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Parámetros    :
'                 p_CodIns    -  Código de Instancia
'                 p_TipFil    -  Tipo de Filtro
'                                1 - Producto
'                                2 - Consejero Hipotecario
'                                3 - Modalidad
'                                4 - Por Tipo de Evaluación
'                 p_CodFil    - Código Detalle de Tipo de Filtro
'                 p_Situac    -  Tipo de Resultado
'                                0 - Todas
'                                1 - Aprobada
'                                2 - Rechazada
'                 p_FecIni    -  Fecha de Inicio
'                 p_FecFin    -  Fecha de Fin
'F. Creación   :  02-03-2010
'Objetivo      :  Permite exportar la información de una Instancia de Evaluación (Resultado Final) ordenado por Producto
'Realizado por :  Miguel Ikehara Punk
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_Cadena     As String
Dim r_str_DetObs     As String
Dim r_str_FecCon     As String
Dim r_int_TpoObs     As Integer
Dim r_int_FlgPry     As Integer
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodMod     As String
Dim r_rst_SolMae     As ADODB.Recordset
Dim r_rst_SegDet     As ADODB.Recordset
Dim r_rst_SegCon     As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       SOLMAE_FECSOL, TRIM(E.PARDES_DESCRI) AS TIPO_EVALUAC, SOLMAE_CONHIP, TRIM(F.PARDES_DESCRI) AS INSTANCIA, TRIM(G.PARDES_DESCRI) AS SITUACION, "
   g_str_Parame = g_str_Parame & "       SEGUIM_FECINI, SEGUIM_FECFIN, SOLMAE_TIPMON, TRIM(H.PARDES_DESCRI) AS TIPO_MONEDA, SOLMAE_MTOPRE_MPR, SOLMAE_TASINT, SOLMAE_PLAANO, "
   g_str_Parame = g_str_Parame & "       SOLMAE_CODMOD, SOLMAE_CODPRD, SOLMAE_TIPEVA, SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_DOL, SEGUIM_SITUAC, SOLMAE_COMVTA_SOL, SOLMAE_APOPRO_SOL, "
   g_str_Parame = g_str_Parame & "       PRODUC_DESCRI, NVL(I.EVACRE_NROACT, '') AS NRO_ACTA, M.SEGEXC_MOTEXC, TRIM(X.SUBPRD_DESCRI) AS SUB_PRODUCTO, O.HIPMAE_FECDES, "
   g_str_Parame = g_str_Parame & "       K.DATGEN_TITULO, NVL(L.PARDES_DESCRI, 'SIN INMUEBLE') AS MODALIDAD, TRIM(N.PARDES_DESCRI) AS EXCEPCION, TRIM(M.SEGEXC_DESCRI) AS DESCRI_EXCEP, "
   g_str_Parame = g_str_Parame & "       TRIM(P.PARDES_DESCRI) AS SITUACION_ACTUAL, SOLMAE_SITUAC, TRIM(Q.PARDES_DESCRI) AS INSTANCIA_ACTUAL, TRIM(R.PARDES_DESCRI) AS MOT_RECHAZO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_SEGUIM B ON B.SEGUIM_NUMSOL = A.SOLMAE_NUMERO AND B.SEGUIM_CODINS = " & CStr(p_CodIns) & " AND B.SEGUIM_FECFIN >= " & p_FecIni & " AND B.SEGUIM_FECFIN <= " & p_FecFin & " "
   If p_Situac > 0 Then
      g_str_Parame = g_str_Parame & "        AND SEGUIM_SITUAC = " & CStr(p_Situac)
   Else
      g_str_Parame = g_str_Parame & "        AND (SEGUIM_SITUAC = 1 OR SEGUIM_SITUAC = 2)"
   End If
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.SOLMAE_TITTDO AND TRIM(D.DATGEN_NUMDOC) = TRIM(A.SOLMAE_TITNDO) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 038 AND E.PARDES_CODITE = A.SOLMAE_TIPEVA "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 002 AND F.PARDES_CODITE = " & CStr(p_CodIns) & " "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 023 AND G.PARDES_CODITE = SEGUIM_SITUAC "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 229 AND H.PARDES_CODITE = SOLMAE_TIPMON "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES P ON P.PARDES_CODGRP = 020 AND P.PARDES_CODITE = SOLMAE_SITUAC "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 002 AND Q.PARDES_CODITE = SOLMAE_CODINS "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SUBPRD X ON X.SUBPRD_CODPRD = A.SOLMAE_CODPRD AND X.SUBPRD_CODSUB = A.SOLMAE_CODSUB "
   g_str_Parame = g_str_Parame & "  LEFT JOIN TRA_EVACRE I ON I.EVACRE_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM J ON J.SOLINM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN K ON K.DATGEN_CODIGO = J.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES L ON L.PARDES_CODGRP = 272 AND L.PARDES_CODITE = '0000'||A.SOLMAE_CODMOD "
   g_str_Parame = g_str_Parame & "  LEFT JOIN TRA_SEGEXC M ON M.SEGEXC_NUMSOL = A.SOLMAE_NUMERO AND SEGEXC_CODINS = 21 "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES N ON N.PARDES_CODGRP = 042 AND N.PARDES_CODITE = M.SEGEXC_MOTEXC "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPMAE O ON O.HIPMAE_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 003 AND R.PARDES_CODITE = SOLMAE_MOTREC "
   g_str_Parame = g_str_Parame & " "
   Select Case p_TipFil
      Case 1: If Len(Trim(p_CodFil)) > 0 Then g_str_Parame = g_str_Parame & " WHERE  SOLMAE_CODPRD = '" & p_CodFil & "' "
      Case 2: If Len(Trim(p_CodFil)) > 0 Then g_str_Parame = g_str_Parame & " WHERE  SOLMAE_CONHIP = '" & p_CodFil & "' "
      Case 3: If Len(Trim(p_CodFil)) > 0 Then g_str_Parame = g_str_Parame & " WHERE  SOLMAE_CODMOD = '" & p_CodFil & "' "
      Case 4: If Len(Trim(p_CodFil)) > 0 Then g_str_Parame = g_str_Parame & " WHERE  SOLMAE_TIPEVA = '" & p_CodFil & "' "
   End Select
   If (Len(Trim(p_CodIte)) > 0) Then
      g_str_Parame = g_str_Parame & " AND SOLINM_PRYCOD = " & Trim(p_CodIte)
   End If
   g_str_Parame = g_str_Parame & " ORDER BY SOLMAE_FECSOL"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_SolMae, 3) Then
      MsgBox "Error al leer tabla CRE_SOLMAE.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_rst_SolMae.BOF And r_rst_SolMae.EOF Then
      MsgBox "No se encontraron registros para los criterios de busqueda seleccionados.", vbExclamation, modgen_g_str_NomPlt
      r_rst_SolMae.Close
      Set r_rst_SolMae = Nothing
      Exit Sub
   End If
   
   r_rst_SolMae.MoveFirst
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 6:         .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40:        .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 3) = "SUB-PRODUCTO":                  .Columns("C").ColumnWidth = 65:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "SOLICITUD":                     .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "DOC. IDENTIDAD":                .Columns("E").ColumnWidth = 15:        .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "NOMBRE CLIENTE":                .Columns("F").ColumnWidth = 44
      .Cells(1, 7) = "F. SOLICITUD":                  .Columns("G").ColumnWidth = 12:        .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "TASA":                          .Columns("H").ColumnWidth = 10:
      .Cells(1, 9) = "PLAZO":                         .Columns("I").ColumnWidth = 10:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "NRO. DE ACTA":                 .Columns("J").ColumnWidth = 13:        .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "TIPO EVALUACION":              .Columns("K").ColumnWidth = 22:        .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 12) = "CONSEJERO HIPOTECARIO":        .Columns("L").ColumnWidth = 24:        .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 13) = "INSTANCIA":                    .Columns("M").ColumnWidth = 24:        .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 14) = "RESULTADO EVALUACION":         .Columns("N").ColumnWidth = 24:        .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 15) = "INGRESO INSTANCIA":            .Columns("O").ColumnWidth = 18:        .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 16) = "SALIDA INSTANCIA":             .Columns("P").ColumnWidth = 18:        .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 17) = "TIEMPO INSTANCIA":             .Columns("Q").ColumnWidth = 18:        .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 18) = "TIEMPO OBSERVADO":             .Columns("R").ColumnWidth = 18:        .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 19) = "TIEMPO EVALUACION":            .Columns("S").ColumnWidth = 20:        .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 20) = "MONEDA":                       .Columns("T").ColumnWidth = 10:        .Columns("T").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 21) = "V. INMUEBLE":                  .Columns("U").ColumnWidth = 12
      .Cells(1, 22) = "CUOTA INICIAL":                .Columns("V").ColumnWidth = 14
      .Cells(1, 23) = "PORC. INICIAL":                .Columns("W").ColumnWidth = 13
      .Cells(1, 24) = "MTO. CREDITO S/.":             .Columns("X").ColumnWidth = 18
      .Cells(1, 25) = "MTO. CREDITO US$":             .Columns("Y").ColumnWidth = 18
      .Cells(1, 26) = "MODALIDAD":                    .Columns("Z").ColumnWidth = 50
      .Cells(1, 27) = "PROYECTO INMOBILIARIO":        .Columns("AA").ColumnWidth = 50
      .Cells(1, 28) = "MOTIVO RECHAZO":               .Columns("AB").ColumnWidth = 140
      .Cells(1, 29) = "COD. EXCEP.":                  .Columns("AC").ColumnWidth = 15:       .Columns("AC").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 30) = "EXCEPCION":                    .Columns("AD").ColumnWidth = 80
      .Cells(1, 31) = "OBSERVACION EXCEPCION":        .Columns("AE").ColumnWidth = 100
      .Cells(1, 32) = "CONDICIONES":                  .Columns("AF").ColumnWidth = 100
      .Cells(1, 33) = "SITUACION ACTUAL":             .Columns("AG").ColumnWidth = 50:       .Columns("AG").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(1, 33)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 33)).HorizontalAlignment = xlHAlignCenter
   End With
     
   r_int_ConVer = 2
   Do While Not r_rst_SolMae.EOF
      r_str_DetObs = ""
      r_int_TpoObs = 0
      
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      r_str_Cadena = r_str_Cadena & "   AND SEGDET_CODINS = " & CStr(p_CodIns) & " "
      r_str_Cadena = r_str_Cadena & "   AND SEGDET_CODOCU = 21"
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
         MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
         r_rst_SegDet.MoveFirst
         Do While Not r_rst_SegDet.EOF
            If r_rst_SegDet!SEGFECACT > 0 Then
               r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
            Else
               r_int_TpoObs = r_int_TpoObs + CInt(CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_SegDet!SEGFECCRE))))
               r_str_DetObs = Trim(r_rst_SegDet!SEGDET_OBSERV & "")
            End If
            r_rst_SegDet.MoveNext
         Loop
      End If
      
      r_rst_SegDet.Close
      Set r_rst_SegDet = Nothing
      
      'Abriendo Cursor para obtener Motivo de Rechazo
      r_str_DetObs = ""
      If r_rst_SolMae!SEGUIM_SITUAC = 2 Then
         r_str_Cadena = ""
         r_str_Cadena = r_str_Cadena & "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
         r_str_Cadena = r_str_Cadena & "   AND SEGDET_CODINS = " & CStr(p_CodIns) & " "
         r_str_Cadena = r_str_Cadena & "   AND SEGDET_CODOCU = 13 "
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegDet, 3) Then
            MsgBox "Error al leer tabla TRA_SEGDET.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
         
         If Not (r_rst_SegDet.BOF And r_rst_SegDet.EOF) Then
            r_rst_SegDet.MoveFirst
            r_str_DetObs = moddat_gf_Consulta_ParDes("003", CStr(r_rst_SegDet!SEGDET_MOTREC)) & Chr(10) & Chr(13) & Trim(r_rst_SegDet!SEGDET_OBSERV & "")
         End If
         
         r_rst_SegDet.Close
         Set r_rst_SegDet = Nothing
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(r_rst_SolMae!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(r_rst_SolMae!SUB_PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(r_rst_SolMae!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(r_rst_SolMae!SOLMAE_TITTDO) & "-" & Trim(r_rst_SolMae!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(r_rst_SolMae!NOM_CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = r_rst_SolMae!SOLMAE_TASINT
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = r_rst_SolMae!SOLMAE_PLAANO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(r_rst_SolMae!NRO_ACTA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(r_rst_SolMae!TIPO_EVALUAC)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(r_rst_SolMae!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(r_rst_SolMae!INSTANCIA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(r_rst_SolMae!SITUACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = CInt(CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI))))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CInt(CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECFIN))) - CDate(gf_FormatoFecha(CStr(r_rst_SolMae!SEGUIM_FECINI)))) - r_int_TpoObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(r_rst_SolMae!TIPO_MONEDA)
      
      If r_rst_SolMae!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(r_rst_SolMae!SOLMAE_COMVTA_SOL, "###,###0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(r_rst_SolMae!SOLMAE_APOPRO_SOL / r_rst_SolMae!SOLMAE_COMVTA_SOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = 0
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(r_rst_SolMae!SOLMAE_COMVTA_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(r_rst_SolMae!SOLMAE_APOPRO_DOL / r_rst_SolMae!SOLMAE_COMVTA_DOL * 100, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(r_rst_SolMae!SOLMAE_MTOPRE_MPR, "###,##0.00")
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Trim(r_rst_SolMae!MODALIDAD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = Trim(r_rst_SolMae!DATGEN_TITULO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = r_str_DetObs
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Trim(r_rst_SolMae!SEGEXC_MOTEXC)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Trim(r_rst_SolMae!EXCEPCION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Trim(r_rst_SolMae!DESCRI_EXCEP)
      
      'Abriendo Cursor para obtener Condiciones
      r_str_DetObs = ""
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "SELECT * FROM TRA_SEGCON WHERE SEGCON_NUMSOL = '" & r_rst_SolMae!SOLMAE_NUMERO & "' "
      r_str_Cadena = r_str_Cadena & "   AND SEGCON_CODINS = " & CStr(p_CodIns) & " "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_SegCon, 3) Then
         MsgBox "Error al leer tabla TRA_SEGCON.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (r_rst_SegCon.BOF And r_rst_SegCon.EOF) Then
         r_rst_SegCon.MoveFirst
         r_str_DetObs = Trim(r_rst_SegCon!SEGCON_OBSCON & "")
      End If
      
      r_rst_SegCon.Close
      Set r_rst_SegCon = Nothing
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Trim(r_str_DetObs)
      
      If r_rst_SolMae!SOLMAE_SITUAC = 1 Then
         'En Tramite
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Trim(r_rst_SolMae!SITUACION_ACTUAL) & " - " & Trim(r_rst_SolMae!INSTANCIA_ACTUAL)
      ElseIf r_rst_SolMae!SOLMAE_SITUAC = 2 Then
         'Desembolsada
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Trim(r_rst_SolMae!SITUACION_ACTUAL) & " - " & CDate(gf_FormatoFecha(CStr(r_rst_SolMae!HIPMAE_FECDES)))
         
      ElseIf r_rst_SolMae!SOLMAE_SITUAC = 3 Then
         'Rechazada
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Trim(r_rst_SolMae!SITUACION_ACTUAL) & " - " & Trim(r_rst_SolMae!MOT_RECHAZO)
      Else
         'Otro
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = "NO SE PUDO DETERMINAR SITUACION"
      End If
      
      r_int_ConVer = r_int_ConVer + 1
      r_rst_SolMae.MoveNext
   Loop
   
   r_rst_SolMae.Close
   Set r_rst_SolMae = Nothing
   
   'Ordenando por Producto y Cliente
   Select Case p_TipFil
      Case 1   'Producto
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 33)).Sort r_obj_Excel.Range("B1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      Case 2   'Consejero Hipotecario
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 33)).Sort r_obj_Excel.Range("I1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      Case 3   'Modalidad - Proyecto
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 33)).Sort r_obj_Excel.Range("X1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
      Case 4   'Tipo de Evaluación
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 33)).Sort r_obj_Excel.Range("H1"), xlAscending, r_obj_Excel.Range("E1"), , xlAscending, , , xlYes
   End Select
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Public Sub modmip_gs_Cbr_RegAcc(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_NumOpe As String, ByVal p_TipAcc As Integer, ByVal p_FecAcc As String, _
                                ByVal p_HorAcc As String, ByVal p_UsuAcc As String, ByVal p_Telefo As String, ByVal p_CorEle As String, ByVal p_Direcc As String, _
                                ByVal p_CodFor As String, ByVal p_ResAcc As Integer, ByVal p_RptCli As Integer, ByVal p_MotAtr As Integer, ByVal p_FecCom As String, _
                                ByVal p_SitCom As Integer, ByVal p_Observ As String, ByVal p_CodNot As String, ByVal p_CuoCom As Integer)
                                
   Dim r_str_Parame     As String
   Dim r_rst_Grabar     As ADODB.Recordset

   r_str_Parame = "USP_CBR_REGACC ("
   r_str_Parame = r_str_Parame & CStr(p_TipDoc) & ", "
   r_str_Parame = r_str_Parame & "'" & p_NumDoc & "', "
   r_str_Parame = r_str_Parame & "'" & p_NumOpe & "', "
   r_str_Parame = r_str_Parame & p_FecAcc & ", "
   r_str_Parame = r_str_Parame & p_HorAcc & ", "
   r_str_Parame = r_str_Parame & CStr(p_TipAcc) & ", "
   r_str_Parame = r_str_Parame & "'" & p_UsuAcc & "', "
   r_str_Parame = r_str_Parame & "'" & p_Telefo & "', "
   r_str_Parame = r_str_Parame & "'" & p_CorEle & "', "
   r_str_Parame = r_str_Parame & "'" & p_Direcc & "', "
   r_str_Parame = r_str_Parame & "'" & p_CodFor & "', "
   r_str_Parame = r_str_Parame & CStr(p_ResAcc) & ", "
   r_str_Parame = r_str_Parame & CStr(p_RptCli) & ", "
   r_str_Parame = r_str_Parame & CStr(p_MotAtr) & ", "
   r_str_Parame = r_str_Parame & p_FecCom & ", "
   r_str_Parame = r_str_Parame & CStr(p_SitCom) & ", "
   r_str_Parame = r_str_Parame & "'" & p_Observ & "', "
   r_str_Parame = r_str_Parame & "'" & p_CodNot & "', "
   r_str_Parame = r_str_Parame & CStr(p_CuoCom) & ", "
   
   'Datos de Auditoria
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
   r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Grabar, 2) Then
      MsgBox "No se pudo ejecutar el procedimiento USP_CBR_REGACC.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
End Sub

Public Sub modmip_gs_Cbr_TelCli(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   Dim r_int_CygTDo     As Integer
   Dim r_str_CygNDo     As String
   Dim r_rst_Genera     As ADODB.Recordset

   p_Combo.Clear
   ReDim p_Arregl(0)

   'Obteniendo Información del Maestro de Clientes
   r_int_CygTDo = 0
   r_str_CygNDo = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CLI_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND DATGEN_NUMDOC = '" & p_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      'Agregando Teléfono Fijo
      If Len(Trim(r_rst_Genera!DatGen_Telefo & "")) > 0 Then
         p_Combo.AddItem "TELEFONO FIJO CLIENTE - " & Trim(r_rst_Genera!DatGen_Telefo & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DatGen_Telefo)
      End If
      
      'Agregando Teléfono Fijo Adicional 1
      If Len(Trim(r_rst_Genera!DatGen_Telefo1 & "")) > 0 Then
         p_Combo.AddItem "TELEFONO FIJO CLIENTE ADICIONAL - " & Trim(r_rst_Genera!DatGen_Telefo1 & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DatGen_Telefo1)
      End If
      
      'Agregando Teléfono Fijo Adicional 2
      If Len(Trim(r_rst_Genera!DatGen_Telefo2 & "")) > 0 Then
         p_Combo.AddItem "TELEFONO FIJO CLIENTE ADICIONAL - " & Trim(r_rst_Genera!DatGen_Telefo2 & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DatGen_Telefo2)
      End If
      
      'Agregando Teléfono Celular Personal
      If Len(Trim(r_rst_Genera!DATGEN_NUMCEL & "")) > 0 Then
         p_Combo.AddItem "TELEFONO CELULAR CLIENTE - " & Trim(r_rst_Genera!DATGEN_NUMCEL & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DATGEN_NUMCEL)
      End If
      
      'Agregando Teléfono Celular Personal Adicional 1
      If Len(Trim(r_rst_Genera!DATGEN_NUMCEL1 & "")) > 0 Then
         p_Combo.AddItem "TELEFONO CELULAR CLIENTE ADICIONAL - " & Trim(r_rst_Genera!DATGEN_NUMCEL1 & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DATGEN_NUMCEL1)
      End If
      
      'Agregando Teléfono Celular Personal Adicional 2
      If Len(Trim(r_rst_Genera!DATGEN_NUMCEL2 & "")) > 0 Then
         p_Combo.AddItem "TELEFONO CELULAR CLIENTE ADICIONAL - " & Trim(r_rst_Genera!DATGEN_NUMCEL2 & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DATGEN_NUMCEL2)
      End If
      
      If (r_rst_Genera!DATGEN_ESTCIV = 2 And r_rst_Genera!DATGEN_REGCYG = 1) Or r_rst_Genera!DATGEN_ESTCIV = 5 Then
         r_int_CygTDo = r_rst_Genera!DATGEN_CYGTDO
         r_str_CygNDo = Trim(r_rst_Genera!DATGEN_CYGNDO)
      End If
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing

   'Obteniendo Información del Maestro de Clientes (Cónyuge)
   If r_int_CygTDo > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CLI_DATGEN "
      g_str_Parame = g_str_Parame & " WHERE DATGEN_TIPDOC = " & CStr(r_int_CygTDo) & " "
      g_str_Parame = g_str_Parame & "   AND DATGEN_NUMDOC = '" & r_str_CygNDo & "'"
   
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
         r_rst_Genera.MoveFirst
         
         'Agregando Teléfono Celular Personal
         If Len(Trim(r_rst_Genera!DATGEN_NUMCEL & "")) > 0 Then
            p_Combo.AddItem "TELEFONO CELULAR CONYUGE - " & Trim(r_rst_Genera!DATGEN_NUMCEL & "")
            ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
            p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DATGEN_NUMCEL)
         End If
      End If
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If

   'Cargando Teléfonos de Referencias Adicionales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CBR_REFPER "
   g_str_Parame = g_str_Parame & " WHERE REFPER_TDOCLI = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND REFPER_NDOCLI = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND REFPER_TIPREF = 1 "
   g_str_Parame = g_str_Parame & "   AND REFPER_SITUAC = 1 ORDER BY REFPER_NUMREF"
         
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         p_Combo.AddItem Trim(r_rst_Genera!REFPER_NOMREF) & " - " & Trim(r_rst_Genera!REFPER_NUMTEL & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!REFPER_NUMTEL)
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   'Cargando Teléfonos de CLI_ADCTEL
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CLI_ADCTEL "
   g_str_Parame = g_str_Parame & " WHERE ADCTEL_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND ADCTEL_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "   AND ADCTEL_ESTADO = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY ADCTEL_TIPTEL"
         
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         If r_rst_Genera!ADCTEL_TIPTEL = 1 Then
            p_Combo.AddItem "TELEFONO FIJO CLIENTE ADICIONAL - " & Trim(r_rst_Genera!ADCTEL_NUMERO & "")
            ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
            p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!ADCTEL_NUMERO)
         Else
            p_Combo.AddItem "TELEFONO CELULAR CLIENTE ADICIONAL - " & Trim(r_rst_Genera!ADCTEL_NUMERO & "")
            ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
            p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!ADCTEL_NUMERO)
         End If
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modmip_gs_Cbr_CorCli(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   Dim r_int_CygTDo     As Integer
   Dim r_str_CygNDo     As String
   Dim r_rst_Genera     As ADODB.Recordset

   p_Combo.Clear
   ReDim p_Arregl(0)

   'Obteniendo Información del Maestro de Clientes
   r_int_CygTDo = 0
   r_str_CygNDo = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      'Agregando Correo Electrónico
      If Len(Trim(r_rst_Genera!DatGen_DirEle & "")) > 0 Then
         p_Combo.AddItem "PERSONAL CLIENTE  - " & Trim(r_rst_Genera!DatGen_DirEle & "")
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DatGen_DirEle)
      End If
      
      If (r_rst_Genera!DATGEN_ESTCIV = 2 And r_rst_Genera!DATGEN_REGCYG = 1) Or r_rst_Genera!DATGEN_ESTCIV = 5 Then
         r_int_CygTDo = r_rst_Genera!DATGEN_CYGTDO
         r_str_CygNDo = Trim(r_rst_Genera!DATGEN_CYGNDO)
      End If
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing

   'Obteniendo Información del Maestro de Clientes (Cónyuge)
   If r_int_CygTDo > 0 Then
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(r_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & r_str_CygNDo & "'"
   
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
         r_rst_Genera.MoveFirst
         
         'Agregando Correo Personal
         If Len(Trim(r_rst_Genera!DatGen_DirEle & "")) > 0 Then
            p_Combo.AddItem "PERSONAL CONYUGE - " & Trim(r_rst_Genera!DatGen_DirEle & "")
         
            ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
            p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!DatGen_DirEle)
         End If
      End If
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If

   'Cargando Correos de Referencias Adicionales
   g_str_Parame = "SELECT * FROM CBR_REFPER WHERE "
   g_str_Parame = g_str_Parame & "REFPER_TDOCLI = " & CStr(p_TipDoc) & " AND  "
   g_str_Parame = g_str_Parame & "REFPER_NDOCLI = '" & p_NumDoc & "' AND  "
   g_str_Parame = g_str_Parame & "REFPER_TIPREF = 3 AND "
   g_str_Parame = g_str_Parame & "REFPER_SITUAC = 1 ORDER BY REFPER_NUMREF"
         
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         p_Combo.AddItem Trim(r_rst_Genera!REFPER_NOMREF) & " - " & Trim(r_rst_Genera!REFPER_DIRELE & "")
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!REFPER_DIRELE)
         
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing

   'Cargando Teléfonos de CLI_ADCEML
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CLI_ADCEML WHERE "
   g_str_Parame = g_str_Parame & " ADCEML_TIPDOC = " & CStr(p_TipDoc) & " AND  "
   g_str_Parame = g_str_Parame & " ADCEML_NUMDOC = '" & p_NumDoc & "' AND  "
   g_str_Parame = g_str_Parame & " ADCEML_ESTADO = 1 "
         
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         p_Combo.AddItem "PERSONAL CLIENTE ADICIONAL - " & Trim(r_rst_Genera!ADCEML_NOMEML & "")
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(r_rst_Genera!ADCEML_NOMEML)
         
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modmip_gs_Cbr_DirCli(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   Dim r_rst_Genera     As ADODB.Recordset
   Dim r_str_Direcc     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String

   p_Combo.Clear
   ReDim p_Arregl(0)

   
   'Dirección de Correspondencia
   r_str_Direcc = ""
   
   g_str_Parame = "SELECT * FROM PVT_DIRCOR WHERE "
   g_str_Parame = g_str_Parame & "DIRCOR_TDOCLI = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DIRCOR_NDOCLI = '" & p_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      r_str_Direcc = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Genera!DIRCOR_TIPVIA)) & _
                     " " & Trim(r_rst_Genera!DIRCOR_NOMVIA) & " " & Trim(r_rst_Genera!DIRCOR_NUMVIA) & _
                     IIf(Len(Trim(r_rst_Genera!DIRCOR_INTDPT)) > 0, " - DPTO / INT.: " & Trim(r_rst_Genera!DIRCOR_INTDPT), "") & " " & _
                     IIf(Len(Trim(r_rst_Genera!DIRCOR_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(r_rst_Genera!DIRCOR_TIPZON)) & " " & Trim(r_rst_Genera!DIRCOR_NOMZON), "") & Chr(10) & Chr(13) & _
                     moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DIRCOR_UBIGEO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DIRCOR_UBIGEO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(r_rst_Genera!DIRCOR_UBIGEO))
      
      p_Combo.AddItem "CORRESPONDENCIA - " & r_str_Direcc
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = r_str_Direcc
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Genera!DIRCOR_TIPVIA)) & _
                                                 " " & Trim(r_rst_Genera!DIRCOR_NOMVIA) & " " & Trim(r_rst_Genera!DIRCOR_NUMVIA) & _
                                                 IIf(Len(Trim(r_rst_Genera!DIRCOR_INTDPT)) > 0, " - DPTO / INT.: " & Trim(r_rst_Genera!DIRCOR_INTDPT), "") & " " & _
                                                 IIf(Len(Trim(r_rst_Genera!DIRCOR_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(r_rst_Genera!DIRCOR_TIPZON)) & " " & Trim(r_rst_Genera!DIRCOR_NOMZON), "")
                     
      p_Arregl(UBound(p_Arregl)).Genera_Prefij = moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DIRCOR_UBIGEO, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DIRCOR_UBIGEO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(r_rst_Genera!DIRCOR_UBIGEO))
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
   
   
   'Dirección de Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, r_str_Direcc, r_str_Distri)

   p_Combo.AddItem "INMUEBLE - " & r_str_Direcc & Chr(10) & Chr(13) & r_str_Distri
   
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = r_str_Direcc & Chr(10) & Chr(13) & r_str_Distri
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = r_str_Direcc
   p_Arregl(UBound(p_Arregl)).Genera_Prefij = r_str_Distri

   'Obteniendo Dirección de Solicitud
   r_str_Direcc = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      r_str_Direcc = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Genera!DatGen_TipVia)) & _
                     " " & Trim(r_rst_Genera!DatGen_NomVia) & " " & Trim(r_rst_Genera!DatGen_Numero) & _
                     IIf(Len(Trim(r_rst_Genera!DATGEN_INTDPT)) > 0, " - DPTO / INT.: " & Trim(r_rst_Genera!DATGEN_INTDPT), "") & " " & _
                     IIf(Len(Trim(r_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(r_rst_Genera!DatGen_TipZon)) & " " & Trim(r_rst_Genera!DatGen_NomZon), "") & Chr(10) & Chr(13) & _
                     moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(r_rst_Genera!DatGen_Ubigeo))
      
      
      p_Combo.AddItem "SOLICITUD - " & r_str_Direcc
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = r_str_Direcc
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Genera!DatGen_TipVia)) & _
                                                 " " & Trim(r_rst_Genera!DatGen_NomVia) & " " & Trim(r_rst_Genera!DatGen_Numero) & _
                                                 IIf(Len(Trim(r_rst_Genera!DATGEN_INTDPT)) > 0, " - DPTO / INT.: " & Trim(r_rst_Genera!DATGEN_INTDPT), "") & " " & _
                                                 IIf(Len(Trim(r_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(r_rst_Genera!DatGen_TipZon)) & " " & Trim(r_rst_Genera!DatGen_NomZon), "")
      p_Arregl(UBound(p_Arregl)).Genera_Prefij = moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DatGen_Ubigeo, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!DatGen_Ubigeo, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(r_rst_Genera!DatGen_Ubigeo))
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing


   'Cargando Teléfonos de Referencias Adicionales
   g_str_Parame = "SELECT * FROM CBR_REFPER WHERE "
   g_str_Parame = g_str_Parame & "REFPER_TDOCLI = " & CStr(p_TipDoc) & " AND  "
   g_str_Parame = g_str_Parame & "REFPER_NDOCLI = '" & p_NumDoc & "' AND  "
   g_str_Parame = g_str_Parame & "REFPER_TIPREF = 2 AND "
   g_str_Parame = g_str_Parame & "REFPER_SITUAC = 1 ORDER BY REFPER_NUMREF"
         
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Genera!REFPER_TIPVIA))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(r_rst_Genera!REFPER_TIPZON))
         
         'Departamento
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!REFPER_UBIGEO, 2) & "0000")
         
         'Provincia
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(r_rst_Genera!REFPER_UBIGEO, 4) & "00")
         
         'Distrito
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(r_rst_Genera!REFPER_UBIGEO))
         
         r_str_Direcc = r_str_TipVia & " " & Trim(r_rst_Genera!REFPER_NOMVIA) & " " & Trim(r_rst_Genera!REFPER_NUMVIA)
      
         If Len(Trim(Trim(r_rst_Genera!REFPER_INTDPT))) > 0 Then
            r_str_Direcc = r_str_Direcc & " DPTO / INT.: " & Trim(r_rst_Genera!REFPER_INTDPT)
         End If
      
         If Len(Trim(Trim(r_rst_Genera!REFPER_NOMZON))) > 0 Then
            r_str_Direcc = r_str_Direcc & " - " & r_str_TipZon & " " & Trim(r_rst_Genera!REFPER_NOMZON)
         End If
         
         p_Combo.AddItem Trim(r_rst_Genera!REFPER_NOMREF) & " - " & r_str_Direcc & " " & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = r_str_Direcc & " " & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = r_str_Direcc
         p_Arregl(UBound(p_Arregl)).Genera_Prefij = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modmip_gs_Cbr_ResAcc(p_Combo As ComboBox, ByVal p_ValIni As Integer, ByVal p_ValFin As Integer)
   Dim r_rst_Genera     As ADODB.Recordset
   
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '303' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "(PARDES_CODITE = '000101' OR (PARDES_CODITE >= " & Format(p_ValIni, "000000") & " AND PARDES_CODITE <= " & Format(p_ValFin, "000000") & ")) "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         p_Combo.AddItem Trim$(r_rst_Genera!PARDES_DESCRI)
         p_Combo.ItemData(p_Combo.NewIndex) = CInt(r_rst_Genera!PARDES_CODITE)
         r_rst_Genera.MoveNext
      Loop
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub modmip_gs_CbrCom(ByVal p_NumOpe As String)
   Dim r_rst_Grabar     As ADODB.Recordset

   g_str_Parame = "UPDATE CBR_REGACC SET "
   g_str_Parame = g_str_Parame & "REGACC_SITCOM = 5 "
   g_str_Parame = g_str_Parame & "WHERE "
   g_str_Parame = g_str_Parame & "REGACC_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "(REGACC_SITCOM = 2 OR REGACC_SITCOM = 1)"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Grabar, 2) Then
       Exit Sub
   End If
End Sub

Public Function modmip_gf_TotalPBP(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_TipFlg As Integer) As Integer
   modmip_gf_TotalPBP = 0

   g_str_Parame = "SELECT COUNT(*) AS TOTAL FROM CRE_DETPBP WHERE "
   g_str_Parame = g_str_Parame & "DETPBP_PERMES = " & CStr(p_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "DETPBP_PERANO = " & CStr(p_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "DETPBP_FLGPBP = " & CStr(p_TipFlg)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      modmip_gf_TotalPBP = g_rst_Princi!Total
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Public Sub modmip_gs_Consulta_NomPry(ByVal p_CodPry As String, ByRef p_NomPry As String, ByRef p_NomPro As String, ByRef p_TipPry As String)
   Dim r_str_Cadena     As String
   Dim r_rst_Proyec     As ADODB.Recordset
   Dim r_rst_Empres     As ADODB.Recordset

   p_NomPry = ""
   p_NomPro = ""
   p_TipPry = ""

   r_str_Cadena = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & p_CodPry & "' "

   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Proyec, 3) Then
      Exit Sub
   End If

   If Not (r_rst_Proyec.BOF And r_rst_Proyec.EOF) Then
      r_rst_Proyec.MoveFirst
      
      p_NomPry = Trim(r_rst_Proyec!DATGEN_TITULO)
      
      If r_rst_Proyec!DATGEN_PRYMCS = 1 Then
         p_TipPry = "PROYECTO VINCULADO"
      Else
         p_TipPry = "PROYECTO NO VINCULADO"
      End If
      
      'Buscando Información datos de Empresas
      r_str_Cadena = "SELECT * FROM EMP_DATGEN WHERE DATGEN_EMPTDO = " & CStr(r_rst_Proyec!DATGEN_VENTDO) & " AND DATGEN_EMPNDO = '" & Trim(r_rst_Proyec!DATGEN_VENNDO) & "' "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Empres, 3) Then
         Exit Sub
      End If
      
      If Not (r_rst_Empres.BOF And r_rst_Empres.EOF) Then
         r_rst_Empres.MoveFirst
         p_NomPro = Trim(r_rst_Empres!DATGEN_RAZSOC)
      End If
      
      r_rst_Empres.Close
      Set r_rst_Empres = Nothing
   End If

   r_rst_Proyec.Close
   Set r_rst_Proyec = Nothing
End Sub

Public Function modmip_gf_Consulta_NomEmp(ByVal p_TipEmp As Integer, ByVal p_NomEmp As String) As String
   Dim r_rst_Empres     As ADODB.Recordset
   Dim r_str_Cadena     As String
      
   modmip_gf_Consulta_NomEmp = ""
   
   'Buscando Información datos de Empresas
   r_str_Cadena = "SELECT * FROM EMP_DATGEN WHERE DATGEN_EMPTDO = " & CStr(p_TipEmp) & " AND DATGEN_EMPNDO = '" & Trim(p_NomEmp) & "' "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Empres, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Empres.BOF And r_rst_Empres.EOF) Then
      r_rst_Empres.MoveFirst
      modmip_gf_Consulta_NomEmp = Trim(r_rst_Empres!DATGEN_RAZSOC)
   End If
   
   r_rst_Empres.Close
   Set r_rst_Empres = Nothing
End Function

Public Function modmip_gf_Consulta_SecEco(ByVal p_CodSec As Integer) As String
   Dim r_rst_Genera     As ADODB.Recordset

   modmip_gf_Consulta_SecEco = ""

   g_str_Parame = "SELECT * FROM MNT_SECECO WHERE "
   g_str_Parame = g_str_Parame & "SECECO_CODIGO = " & CStr(p_CodSec) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      modmip_gf_Consulta_SecEco = Trim(r_rst_Genera!SECECO_DESCRI & "")
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Sub modmip_gs_DatNumOpe(ByVal p_NumOpe As String, ByVal p_Grid As MSFlexGrid, Optional ByRef p_TipGar As Integer)
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)

   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                  'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                  'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                  'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                  'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   p_TipGar = g_rst_Princi!HIPMAE_TIPGAR
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   Call gs_LimpiaGrid(p_Grid)
   
   'Cargando en Grid
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.CellFontBold = True:           p_Grid.Text = "Número de Operación"
   p_Grid.Col = 1:                       p_Grid.CellFontBold = True:           p_Grid.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.CellFontBold = True:           p_Grid.Text = "Situación"
   p_Grid.Col = 1:                       p_Grid.CellFontBold = True:           'p_Grid.Text = moddat_g_str_Situac
   If moddat_g_int_Situac = 6 Then
      p_Grid.Text = moddat_g_str_Situac & "    -    FECHA : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECCAN))
   Else
      p_Grid.Text = moddat_g_str_Situac
   End If
   p_Grid.CellForeColor = modgen_g_con_ColNar
   
   p_Grid.Rows = p_Grid.Rows + 1:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.CellFontBold = True:           p_Grid.Text = "Cliente"
   p_Grid.Col = 1:                       p_Grid.CellFontBold = True:           p_Grid.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "Cónyuge"
      p_Grid.Col = 1:                    p_Grid.Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Producto"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   p_Grid.Rows = p_Grid.Rows + 1:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Moneda Préstamo"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_Moneda
   
   p_Grid.Rows = p_Grid.Rows + 2:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Primera Vivienda"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!HIPMAE_PRIVIV)
   
   p_Grid.Rows = p_Grid.Rows + 1:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Modalidad"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_DesMod
   
   p_Grid.Rows = p_Grid.Rows + 1:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Dirección Inmueble"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_Direcc
   
   p_Grid.Rows = p_Grid.Rows + 1:        p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Distrito"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "Proyecto Inmobiliario"
      p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         p_Grid.Text = p_Grid.Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
'   If moddat_g_int_TipMon = 1 Then
      p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "Valor Compra Venta"
      p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_CVTSOL, g_rst_Princi!HIPMAE_CVTDOL), 12, 2) 'g_rst_Princi!HIPMAE_CVTSOL
   
      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "Cuota Inicial"
      p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_APOSOL, g_rst_Princi!HIPMAE_APODOL), 12, 2) & " (" & Format((IIf(moddat_g_int_TipMon = 1, CDbl(g_rst_Princi!SOLMAE_APOPRO_SOL), CDbl(g_rst_Princi!SOLMAE_APOPRO_DOL)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO))) / IIf(moddat_g_int_TipMon = 1, CDbl(g_rst_Princi!HIPMAE_CVTSOL), CDbl(g_rst_Princi!HIPMAE_CVTDOL)) * 100, "##0.00") & "%)"
            
      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "                Aporte Propio"
      p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(CDbl(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!HIPMAE_APOSOL, g_rst_Princi!HIPMAE_APODOL)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO)), 12, 2)

      If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe BBP"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
      End If
'   Else
'      p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
'      p_Grid.Col = 0:                    p_Grid.Text = "Valor Compra Venta"
'      p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
'      p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
'
'      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
'      p_Grid.Col = 0:                    p_Grid.Text = "Aporte Propio"
'      p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
'      p_Grid.CellFontSize = 8:           p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
'   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Monto Desembolsado"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Monto Préstamo"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Interés Capitalizado"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Total Préstamo"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Fecha Activación"
   p_Grid.Col = 1:                       p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Fecha Desembolso"
   p_Grid.Col = 1:                       p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "Fecha Firma EE.PP"
      p_Grid.Col = 1:                    p_Grid.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) = 0 Then ' moddat_g_str_CodPrd <> "002" And moddat_g_str_CodPrd <> "011" Then
      p_Grid.Rows = p_Grid.Rows + 2: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd):  p_Grid.Text = "Nro. Operación Mivivienda" '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd):  p_Grid.Text = "Nro. Operación COFIDE"     '"003"
         Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd): p_Grid.Text = "Nro. Operación COFIDE"     '"004"
         Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd): p_Grid.Text = "Nro. Operación COFIDE"     '"007", "009", "010", "012" "013", "014", "015", "016", "017", "018", "019", "021", "022", "023"
      End Select
      
      p_Grid.Col = 1:     p_Grid.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"003"
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                       p_Grid.Text = "Nro. Operación Mivivienda"
         p_Grid.Col = 1:                       p_Grid.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                          p_Grid.Text = "Monto Préstamo (Tramo No Conces.)"
      p_Grid.Col = 1:                          p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:                 p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                          p_Grid.Text = "Monto Préstamo (Tramo Conces.)"
      p_Grid.Col = 1:                          p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:                 p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"001" "003"
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                       p_Grid.Text = "Tasa de Interés Mivivienda"
         p_Grid.Col = 1:                       p_Grid.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then  '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                       p_Grid.Text = "Tasa de Interés COFIDE"
         p_Grid.Col = 1:                       p_Grid.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                       p_Grid.Text = "Tasa de Comisión COFIDE"
         p_Grid.Col = 1:                       p_Grid.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Plazo"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Tasa de Interés"
   p_Grid.Col = 1:                       p_Grid.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Nro. de Cuotas"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Período de Gracia"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Cuotas Extraordinarias"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Compañía de Seguros"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Tipo de Seguro Desg."
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Tipo Garantía"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Monto Garantía"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Saldo Capital"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) = 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Saldo Capital (Tramo Conces.)"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   End If
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Total Saldo Deudor"
   p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
   p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Cuotas Pendientes de Pago"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Días de Atraso"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Consejero Hipotecario"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_NomConHip
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Ejecutivo de Seguimiento"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(p_Grid)
End Sub

Public Sub modmip_gs_DatSolCre(ByVal p_NumSol As String, ByVal p_Grid As MSFlexGrid)
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
   
   Call gs_LimpiaGrid(p_Grid)
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Cliente
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO & "")
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Cónyuge
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!SOLMAE_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!SOLMAE_CYGNDO & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Producto
   moddat_g_str_CodPrd = g_rst_Princi!SOLMAE_CODPRD
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))
   moddat_g_str_CodSub = g_rst_Princi!SOLMAE_CODSUB
   
   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   'Modalidad
   moddat_g_str_CodMod = ""
   moddat_g_str_DesMod = ""
   
   If Len(Trim(g_rst_Princi!SOLMAE_CODMOD & "")) > 0 Then
      moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
      moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)
   End If
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_EJESEG))
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))
   
   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   'Situación
   moddat_g_int_Situac = g_rst_Princi!SOLMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
   
   'Inmueble Identificado
   moddat_g_int_InmIde = g_rst_Princi!SOLMAE_INMIDE
   
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS
   
   'Según Situación
   moddat_g_str_FecDes = ""
   moddat_g_str_FecAnu = ""
   moddat_g_str_FecRec = ""
   moddat_g_int_TipRec = 0
   moddat_g_int_MotRec = 0
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)

   'Cargando en Grid
   p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                          p_Grid.Text = "Número de Solicitud"
   p_Grid.Col = 1:                          p_Grid.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
   
   p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                          p_Grid.Text = "Cliente"
   p_Grid.Col = 1:                          p_Grid.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & " - " & Trim(g_rst_Princi!SOLMAE_TITNDO) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!SOLMAE_CYGTDO > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Cónyuge"
      p_Grid.Col = 1:                       p_Grid.Text = CStr(moddat_g_int_CygTDo) & " - " & Trim(moddat_g_str_CygNDo) & " / " & moddat_g_str_CygNom
   End If
   
   'Apoderado
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera!DATGEN_APOTDO > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Apoderado"
      p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Genera!DATGEN_APOTDO) & " - " & Trim(g_rst_Genera!DATGEN_APONDO) & " / " & Trim(g_rst_Genera!DATGEN_APOAPP) & " " & Trim(g_rst_Genera!DATGEN_APOAPM) & " " & Trim(g_rst_Genera!DATGEN_APONOM)
   End If
   
   p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                          p_Grid.Text = "Producto"
   p_Grid.Col = 1:                          p_Grid.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                          p_Grid.Text = "Primera Vivienda"
   p_Grid.Col = 1:                          p_Grid.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))
   
   p_Grid.Rows = p_Grid.Rows + 1:       p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                          p_Grid.Text = "Moneda Préstamo"
   p_Grid.Col = 1:                          p_Grid.Text = moddat_g_str_Moneda
   
   If Len(Trim(moddat_g_str_Direcc)) > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Modalidad"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_DesMod
   
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Dirección Inmueble"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_Direcc
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Distrito"
      p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_Distri
      
      If Len(Trim(r_str_CodPry)) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Proyecto Inmobiliario"
         p_Grid.Col = 1:                    p_Grid.Text = moddat_gf_Consulta_NomPry(r_str_CodPry)
      ElseIf Len(Trim(r_str_NomPry)) > 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "Proyecto Inmobiliario"
         p_Grid.Col = 1:                    p_Grid.Text = r_str_NomPry & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Then
      p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Valor Compra Venta"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2))
      
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Cuota Inicial"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)) & " (" & Format((IIf(moddat_g_int_TipMon = 1, CDbl(g_rst_Princi!SOLMAE_APOPRO_SOL), CDbl(g_rst_Princi!SOLMAE_APOPRO_DOL)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO))) / IIf(moddat_g_int_TipMon = 1, CDbl(g_rst_Princi!SOLMAE_COMVTA_SOL), CDbl(g_rst_Princi!SOLMAE_COMVTA_DOL)) * 100, "##0.00") & "%)"

'      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
'         p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & IIf(moddat_g_int_TipMon = 1, gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2), gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)) & " (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "###,###,##0.00") & ") "
'      Else
'      End If
      
      p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                    p_Grid.Text = "                Aporte Propio"
      p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(CDbl(IIf(moddat_g_int_TipMon = 1, g_rst_Princi!SOLMAE_APOPRO_SOL, g_rst_Princi!SOLMAE_APOPRO_DOL)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO)) - CDbl(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO)), 12, 2)

      If Not IsNull(g_rst_Princi!SOLMAE_FMVBBP) And g_rst_Princi!SOLMAE_FMVBBP <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe BBP"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_FMVBBP), 0, g_rst_Princi!SOLMAE_FMVBBP), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_AFPMTO) And g_rst_Princi!SOLMAE_AFPMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe AFP"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_AFPMTO), 0, g_rst_Princi!SOLMAE_AFPMTO), 12, 2)
      End If
      If Not IsNull(g_rst_Princi!SOLMAE_BMSMTO) And g_rst_Princi!SOLMAE_BMSMTO <> 0 Then
         p_Grid.Rows = p_Grid.Rows + 1: p_Grid.Row = p_Grid.Rows - 1
         p_Grid.Col = 0:                    p_Grid.Text = "                Importe BMS"
         p_Grid.Col = 1:                    p_Grid.CellFontName = "Lucida Console"
         p_Grid.CellFontSize = 8:           p_Grid.Text = "                " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(IIf(IsNull(g_rst_Princi!SOLMAE_BMSMTO), 0, g_rst_Princi!SOLMAE_BMSMTO), 12, 2)
      End If
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Monto Préstamo"
      p_Grid.Col = 1:                       p_Grid.CellFontName = "Lucida Console"
      p_Grid.CellFontSize = 8:              p_Grid.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_MPR, 12, 2)
   End If
      
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Tasa de Interés"
   p_Grid.Col = 1:                       p_Grid.Text = Format(g_rst_Princi!SOLMAE_TASINT, "##0.00") & "%"
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Plazo"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Número de Cuotas"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PLAANO * 12)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Período de Gracia"
   p_Grid.Col = 1:                       p_Grid.Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"

   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Cuotas Extraordinarias"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Compañía de Seguros"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Tipo de Seguro Desgravamen"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Día de Pago"
   p_Grid.Col = 1:                       p_Grid.Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Situación"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_Situac
   p_Grid.CellFontBold = True:           p_Grid.CellForeColor = modgen_g_con_ColNar
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Fecha de Ingreso"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_FecIng
   
   p_Grid.Rows = p_Grid.Rows + 2:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Consejero Hipotecario"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_NomConHip
   
   p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
   p_Grid.Col = 0:                       p_Grid.Text = "Ejecutivo Seguimiento"
   p_Grid.Col = 1:                       p_Grid.Text = moddat_g_str_NomEjeSeg
   
   p_Grid.Rows = p_Grid.Rows + 1
   
   If g_rst_Genera!DATGEN_TDOVIN > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Vinculado"
      
      p_Grid.Col = 1
      
      If g_rst_Genera!DATGEN_TIPVIN = 1 Then
         p_Grid.Text = "TRABAJADOR"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 2 Or g_rst_Genera!DATGEN_TIPVIN = 3 Then
         p_Grid.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 4 Then
         p_Grid.Text = "FUNCIONARIO"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 5 Then
         p_Grid.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      Else
         p_Grid.Text = ""
      End If
   End If
   
   If g_rst_Genera!DATGEN_TDOACC > 0 Then
      p_Grid.Rows = p_Grid.Rows + 1:    p_Grid.Row = p_Grid.Rows - 1
      p_Grid.Col = 0:                       p_Grid.Text = "Accionista"
      p_Grid.Col = 1
      
      If g_rst_Genera!DATGEN_ACCVIN = 1 Then
         p_Grid.Text = "ACCIONISTA"
      ElseIf g_rst_Genera!DATGEN_ACCVIN = 2 Then
         p_Grid.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Genera!DATGEN_TDOACC, Trim(g_rst_Genera!DATGEN_NDOACC)) & ")"
      End If
   End If
   
   modmip_g_int_PaiRes = g_rst_Genera!DATGEN_PAIRES
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(p_Grid)
End Sub

Public Function modmip_gf_Genera_CodGen(ByVal p_TipFol As Integer, ByVal p_CodPro As String) As String
   Dim r_lng_CodOpe     As Long
   Dim r_str_CodOpe     As String
   
   modmip_gf_Genera_CodGen = ""
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   g_str_Parame = "SELECT * FROM CRE_FOLIOS WHERE "
   g_str_Parame = g_str_Parame & "FOLIOS_TIPFOL = " & p_TipFol & " AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODPRD = '" & p_CodPro & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODSUC = '001' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_PERANO = " & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_CodOpe = 1
   Else
      r_lng_CodOpe = g_rst_Genera!FOLIOS_NUMERO + 1
   End If

   r_str_CodOpe = p_CodPro & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & Format(r_lng_CodOpe, "000000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_CRE_FOLIOS ("
      g_str_Parame = g_str_Parame & p_TipFol & ", "
      g_str_Parame = g_str_Parame & "'" & p_CodPro & "', "
      g_str_Parame = g_str_Parame & "'001', "
      g_str_Parame = g_str_Parame & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_CodOpe) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1, "
      
      If r_lng_CodOpe = 1 Then
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
         If MsgBox("No se pudo completar el procedimiento USP_CRE_FOLIOS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   modmip_gf_Genera_CodGen = r_str_CodOpe
End Function

