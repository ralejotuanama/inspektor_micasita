Attribute VB_Name = "modsec"
Option Explicit

Public modsec_g_str_Period             As String
Public modsec_g_str_TipCam             As String
Public modsec_g_bln_Period             As Boolean
Public modsec_g_str_NumSol             As String
Public modsec_g_int_DiaMor             As Integer
Public modsec_g_str_CadCon             As String
Public modsec_g_str_RutIni             As String
Public modsec_g_str_KeyDes             As String
Public modsec_g_str_CodUsu             As String
Public modsec_g_str_Codigo             As String

Public modsec_g_dbl_MtoSol             As Double
Public modsec_g_dbl_ITFSol             As Double

Public modsec_g_str_CliNac             As String   'Nacionalidad Cliente x
Public modsec_g_str_ClEsCi             As String   'Estado Civil Cliente x
Public modsec_g_str_CygNom             As String   'Nombre Conyuge x
Public modsec_g_str_CygNac             As String   'Nacionalidad Conyuge x
Public modsec_g_str_CyEsCi             As String   'Estado Civil Conyuge x
Public modsec_g_str_DomCom             As String   'Domicilio Común x
Public modsec_g_str_RepNom             As String   'Nombre Representante miCasita
Public modsec_g_str_RepDni             As String   'Dni Representante miCasita
Public modsec_g_str_RepCar             As String   'Cargo Representante miCasita
Public modsec_g_str_VeDeSo             As String   'Denominación Social Vendedora x
Public modsec_g_str_VeReUc             As String   'Registro Único Contribuyenye Vendedora x
Public modsec_g_str_VenDom             As String   'Domicilio Vendedora x
Public modsec_g_str_VeRDni             As String   'Dni Representante Vendedora
Public modsec_g_str_VeRNom             As String   'Nombre Representante Vendedora
Public modsec_g_str_VeRCar             As String   'Cargo Representante Vendedora
Public modsec_g_str_VeRdIp             As String   'Datos Inscripción Poder Representante Vendedora
Public modsec_g_str_InmNum             As String   'Número Inmueble x
Public modsec_g_str_ImInRe             As String   'Información Registral Inmueble
Public modsec_g_str_InArTe             As String   'Área Terreno Inmueble x
Public modsec_g_str_InArTc             As String   'Área Total Construida Inmueble x
Public modsec_g_str_InEsTe             As String   'Especificaciones Técnicas Inmueble
Public modsec_g_str_InVaVi             As String   'Valor Vivienda Inmueble
Public modsec_g_str_InVaTv             As String   'Valor Total Vivienda Inmueble
Public modsec_g_str_InVaTa             As String   'Valor Tasación Inmueble
Public modsec_g_str_InMoGr             As String   'Monto Gravamen Inmueble
Public modsec_g_str_PryNom             As String   'Nombre Proyecto x
Public modsec_g_str_PryDir             As String   'Dirección Proyecto x
Public modsec_g_str_PryNpE             As String   'Número Partida Electrónica Proyecto
Public modsec_g_str_PrInFi             As String   'Inmueble Financiar Proyecto
Public modsec_g_str_PreMon             As String   'Monto Prestamo x
Public modsec_g_str_PrePla             As String   'Plazo Préstamo x
Public modsec_g_str_PrNuCu             As String   'Número Cuotas Préstamo
Public modsec_g_str_PrTeAf             As String   'Tasa Efectiva Anual Fija Préstamo
Public modsec_g_str_PrPeGr             As String   'Periodo Grácia Préstamo x
Public modsec_g_str_PrCoDm             As String   'Comisión Desembolso MiVivienda Préstamo
Public modsec_g_str_PrCoDc             As String   'Comisión Desembolso CME Préstamo
Public modsec_g_str_PrPeDc             As String   'Penalidad Resolución CME Préstamo
Public modsec_g_str_PrFoDe             As String   'Forma Desembolso Préstamo
Public modsec_g_str_InmMan             As String   'Manzana Inmueble
Public modsec_g_str_InmEdi             As String   'Edificio Inmueble
Public modsec_g_str_InmDis             As String   'Distrito Inmueble
Public modsec_g_str_InEsNu             As String   'Estacionamiento Número Inmueble
Public modsec_g_str_InEsDi             As String   'Estacionamiento Distrito Inmueble
'Public modsec_g_str_InPaEl             As String   'Partida Electrónica Inmueble
Public modsec_g_str_InEpEl             As String   'Estacionamiento Partida Electrónica Inmueble


'INICIO DE CODIGO PARA ZIPEAR ARCHIVOS *****************************************************************************************************************

Public Type ZIPUSERFUNCTIONS
DLLPrnt As Long
DLLPassword As Long
DLLComment As Long
DLLService As Long
End Type


Public Type ZPOPT
fSuffix As Long
fEncrypt As Long
fSystem As Long
fVolume As Long
fExtra As Long
fNoDirEntries As Long
fExcludeDate As Long
fIncludeDate As Long
fVerbose As Long
fQuiet As Long
fCRLF_LF As Long
fLF_CRLF As Long
fJunkDir As Long
fRecurse As Long
fGrow As Long
fForce As Long
fMove As Long
fDeleteEntries As Long
fUpdate As Long
fFreshen As Long
fJunkSFX As Long
fLatestTime As Long
fComment As Long
fOffsets As Long
fPrivilege As Long
fEncryption As Long
fRepair As Long
flevel As Byte
date As String
szRootDir As String
End Type

Public Type ZIPnames
    s(0 To 99) As String
End Type

Public Type CBChar
    ch(4096) As Byte
End Type

Public Declare Function ZpInit Lib "zip32.dll" (ByRef Zipfun As ZIPUSERFUNCTIONS) As Long
Public Declare Function ZpSetOptions Lib "zip32.dll" (ByRef Opts As ZPOPT) As Long
Public Declare Function ZpArchive Lib "zip32.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long
Function FuncionParaProcesarPassword(ByRef B1 As Byte, L As Long, ByRef B2 As Byte, ByRef B3 As Byte) As Long
    FuncionParaProcesarPassword = 0
End Function
Function FuncionParaProcesarServicios(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarServicios = 0
End Function
Function FuncionParaProcesarMensajes(ByRef fname As CBChar, ByVal x As Long) As Long
    FuncionParaProcesarMensajes = 0
End Function
Function FuncionParaProcesarComentarios(Comentario As CBChar) As CBChar
    Comentario.ch(0) = vbNullString
    FuncionParaProcesarComentarios = Comentario
End Function
Public Function DevolverDireccionMemoria(Direccion As Long) As Long
    DevolverDireccionMemoria = Direccion
End Function

'FIN DE CODIGO PARA ZIPEAR ARCHIVOS *****************************************************************************************************************

   

Public Sub modsec_gs_Carga_TodAno(p_Combo As ComboBox, p_AnoIni As Integer, ByVal p_AnoAct As Integer)
   Dim AnoIni As Integer
   For AnoIni = p_AnoIni To p_AnoAct
      p_Combo.AddItem Trim$(AnoIni)
      p_Combo.ItemData(p_Combo.NewIndex) = AnoIni
   Next AnoIni
End Sub

Public Sub modsec_gs_Carga_AnoVal(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = "SELECT MIN(HIPCIE_PERANO) AS MINANO, MAX(HIPCIE_PERANO) AS MAXANO FROM CRE_HIPCIE"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      Call modsec_gs_Carga_TodAno(p_Combo, g_rst_Genera!MINANO, g_rst_Genera!MAXANO)
      Exit Sub
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function modsec_gf_Buscar_NumSol(ByVal p_NumOpe As String) As String
   modsec_gf_Buscar_NumSol = ""
   
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      modsec_gf_Buscar_NumSol = Trim(g_rst_Listas!HIPMAE_NUMSOL)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_Buscar_NomCyg(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, Optional ByRef moddat_g_str_CygNDo) As String
   modsec_gf_Buscar_NomCyg = "0"
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & Trim(p_NumDoc) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      modsec_gf_Buscar_NomCyg = "" & Trim(g_rst_Listas!DATGEN_CYGTDO)
      If Not IsNull(g_rst_Listas!DATGEN_CYGNDO) Then
         moddat_g_str_CygNDo = "" & Trim(g_rst_Listas!DATGEN_CYGNDO)
      End If

   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_Buscar_TipPrv(ByVal p_TipPrv As Integer) As String
   modsec_gf_Buscar_TipPrv = ""
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = 352 AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE = " & p_TipPrv & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_Buscar_TipPrv = Trim(g_rst_Listas!PARDES_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub modsec_gf_Decryp()
    ' decrypt file sub
    Dim csCrypt     As New clsCrypto
    Dim strFile     As String
    Dim lFileLength As String
    
    ' get length of file
    lFileLength = FileLen(modgen_g_str_RutWin & modgen_g_con_ArcIni)
    ' allocate string to hold file
    strFile = String(lFileLength, vbNullChar)
    ' open file in binary mode
    Open modgen_g_str_RutWin & modgen_g_con_ArcIni For Binary Access Read As #1
    
    Get 1, , strFile
    Close #1
    ' set password
    csCrypt.Password = modsec_g_str_KeyDes
    csCrypt.InBuffer = strFile
    ' generate password
    If Not csCrypt.GeneratePasswordKey Then Exit Sub
    ' decrypt message
    If Not csCrypt.DecryptFileData Then Exit Sub
    csCrypt.DestroySessionKey
    modsec_g_str_CadCon = csCrypt.OutBuffer
    ' check for valid data
    If csCrypt.OutBuffer <> "" Then
        ' delete current file
        'Kill txtFileName.Text
        ' creat new file
        'Open modsec_g_str_RutIni & modgen_g_con_ArcAux For Binary Access Write As #2
        Open modgen_g_str_RutWin & modgen_g_con_ArcAux For Binary Access Write As #2
        Put 2, , csCrypt.OutBuffer
        Close #2
    End If
End Sub

Public Function modsec_gf_GetPrivateProfileString(ByVal p_Seccion As String, ByVal p_Llave As String, ByVal p_Null As String, p_Respuesta As String, ByVal p_TamRpt As Integer, ByVal p_Archivo As String) As Integer

   Dim r_str_Cadena     As String
   Dim r_int_Contad     As Integer
   
   Do While Len(modsec_g_str_CadCon) > 0
   
      If UCase$(Trim$(VBA.Left(modsec_g_str_CadCon, Len(p_Llave)))) = UCase$(Trim$(p_Llave)) Then
         For r_int_Contad = 1 To Len(modsec_g_str_CadCon$)
            If Mid$(modsec_g_str_CadCon$, r_int_Contad, 1) = "=" Then
               If p_Seccion = "CONEXION" Then
                  p_Respuesta = Mid$(Trim$(modsec_g_str_CadCon), r_int_Contad + 1, Len(modsec_g_str_CadCon) - InStr(modsec_g_str_CadCon, "RUTA_SRV"))
               Else
                  p_Respuesta = Mid$(Trim$(modsec_g_str_CadCon), r_int_Contad + 1, Len(modsec_g_str_CadCon))
               End If
               p_Respuesta = Trim(p_Respuesta)
               p_TamRpt = Len(p_Respuesta)
               modsec_gf_GetPrivateProfileString = p_TamRpt
               Exit Function
            End If
         Next
      End If
      
      modsec_g_str_CadCon = Right(modsec_g_str_CadCon, Len(modsec_g_str_CadCon) - 1)
            
   Loop
   
   p_Respuesta = p_Null
   p_TamRpt = Len(p_Respuesta)
   modsec_gf_GetPrivateProfileString = p_TamRpt

End Function

Public Function modsec_gf_Consulta_SecEco(ByVal p_CodPrd As String) As String
   modsec_gf_Consulta_SecEco = ""
   
   p_CodPrd = Format(p_CodPrd, "000")

   g_str_Parame = "SELECT * FROM MNT_SECECO WHERE "
   g_str_Parame = g_str_Parame & "SECECO_CODIGO = '" & p_CodPrd & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   modsec_gf_Consulta_SecEco = Trim$(g_rst_Listas!SECECO_DESCRI)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_Fin_Del_Mes(fecha As Variant) As Date
  
   If IsDate(fecha) Then
      modsec_gf_Fin_Del_Mes = DateAdd("m", 1, fecha)
      modsec_gf_Fin_Del_Mes = DateSerial(Year(modsec_gf_Fin_Del_Mes), Month(modsec_gf_Fin_Del_Mes), 1)
      modsec_gf_Fin_Del_Mes = DateAdd("d", -1, modsec_gf_Fin_Del_Mes)
   End If
   
End Function

'Calcula la cantidad de días de un año
Public Function modsec_gf_Dias_del_Año(ByVal valor As Variant) As Integer
    If IsDate(valor) Or _
       IsNumeric(valor) Then
            modsec_gf_Dias_del_Año = IIf(modsec_gf_Saltar_Year(valor), 366, 365)
    End If
End Function
  
Public Function modsec_gf_Saltar_Year(ByVal valor As Variant) As Boolean
  
    On Error GoTo LocalError
  
    Dim iYear As Integer
       
    If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)
  
    If TypeName(iYear) = "Integer" Then
        modsec_gf_Saltar_Year = Day(DateSerial(iYear, 3, 0)) = 29
    End If
Exit Function
  
LocalError:
End Function

Public Function modsec_gf_FormatoHora(ByVal p_Hora As String) As String
   
   p_Hora = Format(p_Hora, "000000")
   
   If CLng(p_Hora) = 0 Then
      modsec_gf_FormatoHora = ""
   Else
      If Left(p_Hora, 2) = 12 Then
         modsec_gf_FormatoHora = CStr(CInt(Left(p_Hora, 2))) & ":" & Mid(p_Hora, 3, 2) & ":" & Right(p_Hora, 2) & " p.m"
      ElseIf Left(p_Hora, 2) >= 12 Then
         modsec_gf_FormatoHora = CStr(CInt(Left(p_Hora, 2)) - 12) & ":" & Mid(p_Hora, 3, 2) & ":" & Right(p_Hora, 2) & " p.m"
      Else
         modsec_gf_FormatoHora = CStr(CInt(Left(p_Hora, 2))) & ":" & Mid(p_Hora, 3, 2) & ":" & Right(p_Hora, 2) & " a.m"
      End If
      
   End If
End Function

Public Function gs_modsec_Genera(ByVal p_parametro As String, ByVal p_Opcion As Integer, ByVal p_caracter As String, ByVal p_longitud As Integer, Optional ByVal p_Separador As String) As String
   Dim l_cadena As String
   Dim l_cadaux As String
   Dim l_contad As Integer
   Dim l_longit As Integer
   
   If p_Separador = "" Then
      p_Separador = "."
   End If
   
   If p_Opcion = 1 Then
      l_longit = Len(p_parametro)
      For l_contad = 1 To l_longit Step 1
         If Mid(p_parametro, l_contad, 1) <> p_Separador Then
            l_cadena = l_cadena & Mid(p_parametro, l_contad, 1)
         End If
      Next
      
      If p_caracter = " " Then
         For l_contad = Len(l_cadena) To p_longitud - 1 Step 1
            l_cadena = p_caracter & l_cadena
         Next
         gs_modsec_Genera = l_cadena
      
      Else
         If Left(Trim(l_cadena), 1) = "-" Then
            p_longitud = p_longitud - 1
         End If
         For l_contad = 1 To p_longitud Step 1
            l_cadaux = l_cadaux & p_caracter
         Next
         gs_modsec_Genera = Format(l_cadena, l_cadaux)
         
      End If
   
   ElseIf p_Opcion = 2 Then
   
      l_longit = Len(p_parametro)
   
      For l_contad = l_longit To p_longitud - 1 Step 1
         p_parametro = p_parametro & p_caracter
      Next
      
      gs_modsec_Genera = p_parametro

   End If

End Function

Public Function gs_modsec_SumArr(ByVal p_param1 As String, ByVal p_param2 As String, Optional p_param3 As String) As String
   
   If p_param3 = "" Then
      If p_param1 <> "" And p_param2 <> "" Then
         gs_modsec_SumArr = CDbl(p_param1) + CDbl(p_param2)
      ElseIf p_param1 = "" And p_param2 = "" Then
         gs_modsec_SumArr = 0
      Else
         If p_param1 = "" Then
            gs_modsec_SumArr = 0 + CDbl(p_param2)
         Else
            gs_modsec_SumArr = CDbl(p_param1) + 0
         End If
      End If
   Else
      If p_param1 <> "" And p_param2 <> "" Then
         gs_modsec_SumArr = CDbl(p_param1) - CDbl(p_param2)
      ElseIf p_param1 = "" And p_param2 = "" Then
         gs_modsec_SumArr = 0
      Else
         If p_param1 = "" Then
            gs_modsec_SumArr = 0 - CDbl(p_param2)
         Else
            gs_modsec_SumArr = CDbl(p_param1) - 0
         End If
      End If
   End If

End Function

Public Function gs_modsec_FluDes(ByRef p_MonNac As Double, ByRef p_MonExt As Double, ByVal p_CodAno As Integer, ByVal p_CodMes As Integer, ByVal p_UltDia As Integer, ByVal p_CodCiu As String, ByVal p_Opcion As Integer) As Double

   'ByRef p_CanDes As Integer,
   gs_modsec_FluDes = 0
   
   If p_Opcion = 1 Then
   
      g_str_Parame = "SELECT COUNT(*) AS CANTIDAD, SUM (DECODE(COMCIE_TIPMON,1,COMCIE_MTOPRE)) AS NACIONAL, SUM(DECODE(COMCIE_TIPMON,2,COMCIE_MTOPRE*COMCIE_TIPCAM)) AS EXTRANJERO FROM CRE_COMCIE WHERE "
      g_str_Parame = g_str_Parame & "COMCIE_FECDES >= " & p_CodAno & Format(p_CodMes, "00") & "01" & " AND "
      g_str_Parame = g_str_Parame & "COMCIE_FECDES <= " & p_CodAno & Format(p_CodMes, "00") & Format(p_UltDia, "00")
     
   Else
   
      g_str_Parame = "SELECT COUNT(*) AS CANTIDAD, SUM (DECODE(HIPCIE_TIPMON,1,HIPCIE_MTOPRE)) AS NACIONAL, SUM(DECODE(HIPCIE_TIPMON,2,HIPCIE_MTOPRE*HIPCIE_TIPCAM)) AS EXTRANJERO FROM CRE_HIPCIE WHERE "
      g_str_Parame = g_str_Parame & "HIPCIE_FECDES >= " & p_CodAno & Format(p_CodMes, "00") & "01" & " AND "
      g_str_Parame = g_str_Parame & "HIPCIE_FECDES <= " & p_CodAno & Format(p_CodMes, "00") & Format(p_UltDia, "00")
      
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   'p_CanDes = g_rst_Princi!CANTIDAD
   gs_modsec_FluDes = g_rst_GenAux!CANTIDAD
   
   If IsNull(g_rst_GenAux!NACIONAL) Then
      p_MonNac = 0
   Else
      p_MonNac = g_rst_GenAux!NACIONAL
   End If
   
   If IsNull(g_rst_GenAux!EXTRANJERO) Then
      p_MonExt = 0
   Else
      p_MonExt = Format(g_rst_GenAux!EXTRANJERO, "###,###,##0.00")
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing

End Function

Public Function gs_modsec_ObtNum(ByVal p_parametro As String) As String
   Dim l_cadena As String
   Dim l_contad As Integer
   Dim l_longit As Integer
      
   l_longit = Len(p_parametro)
      
   For l_contad = 1 To l_longit Step 1
      If Mid(p_parametro, l_contad, 1) <> "." Then
         If IsNumeric(Mid(p_parametro, l_contad, 1)) Then
            l_cadena = l_cadena & Mid(p_parametro, l_contad, 1)
         End If
      End If
   Next
      
   gs_modsec_ObtNum = l_cadena

End Function

Public Sub modsec_gs_Carga_ClaPrv(p_Combo As ComboBox)
   p_Combo.Clear
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCLA  "
   g_str_Parame = g_str_Parame & "WHERE TIPCLA_TIPCRE = 13 "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCLA_CODIGO ASC "

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
      p_Combo.AddItem Trim(g_rst_Listas!TIPCLA_DESCRI)
                  
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function modsec_gf_Buscar_TipCla(ByVal p_TipCre As Integer, ByVal p_Codigo As String) As String
   modsec_gf_Buscar_TipCla = ""
   
   g_str_Parame = "SELECT * FROM CTB_TIPCLA WHERE "
   g_str_Parame = g_str_Parame & "TIPCLA_TIPCRE = '" & CStr(p_TipCre) & "' AND "
   g_str_Parame = g_str_Parame & "TIPCLA_CODIGO = '" & CStr(p_Codigo) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_Buscar_TipCla = Trim(g_rst_Listas!TIPCLA_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_Consulta_Promot(ByVal p_CodPro As String, ByVal p_NumPro As String) As String
   modsec_gf_Consulta_Promot = ""
      
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = '" & p_CodPro & "' AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & p_NumPro & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   modsec_gf_Consulta_Promot = Trim$(g_rst_Listas!DATGEN_RAZSOC)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_Buscar_NomCta(ByVal p_CtaCtb As String) As String
   modsec_gf_Buscar_NomCta = ""
   
   g_str_Parame = "SELECT * FROM CNTBL_CNTA WHERE CNTA_CTBL= '" & Trim(p_CtaCtb) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_Buscar_NomCta = UCase(Trim(g_rst_Listas!DESC_CNTA))
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_Buscar_NtaCta(ByVal p_PerAno As String, ByVal p_PerMes As String, ByVal p_NroLib As String, ByVal p_NroAsi As String) As String
   modsec_gf_Buscar_NtaCta = ""
   
   g_str_Parame = "SELECT * FROM CNTBL_ASIENTO WHERE ANO = " & Trim(p_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "MES = " & Trim(p_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO = " & Trim(p_NroLib) & " AND "
   g_str_Parame = g_str_Parame & "NRO_ASIENTO = " & Trim(p_NroAsi) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_Buscar_NtaCta = UCase(Trim(g_rst_Listas!TIPO_NOTA))
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_CanReg(ByVal p_NomTab As String, p_PerAno As Integer, p_PerMes As Integer) As Integer
   modsec_gf_CanReg = 0
   
   g_str_Parame = "SELECT COUNT(*) AS CANTIDAD FROM " & p_NomTab
   g_str_Parame = g_str_Parame & " WHERE " & Right(p_NomTab, 6) & "_PERMES = " & p_PerMes
   g_str_Parame = g_str_Parame & " AND " & Right(p_NomTab, 6) & "_PERANO = " & p_PerAno
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_CanReg = Trim(g_rst_Listas!CANTIDAD)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_BusPai(ByVal p_TipDoc As Integer, p_NumDoc As String) As String
   modsec_gf_BusPai = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_TIPDOC = " & p_TipDoc
   g_str_Parame = g_str_Parame & " AND DATGEN_NUMDOC = " & p_NumDoc
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
   
      g_str_Parame = "SELECT * FROM MNT_PAISUN "
      g_str_Parame = g_str_Parame & "WHERE TRIM(PAISUN_CODPAI) = '" & Trim(g_rst_Listas!DATGEN_NACPAI) & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         modsec_gf_BusPai = Int(Trim(g_rst_Genera!PAISUN_CODSUN))
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
End Function

Public Function modsec_gf_BusMov(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer) As Integer
   modsec_gf_BusMov = 1
   
   g_str_Parame = "SELECT MAX(DETITF_NROCOM) AS NROCOM FROM CTB_DETITF "
   g_str_Parame = g_str_Parame & "WHERE DETITF_TIPDOC = 6 AND "
   g_str_Parame = g_str_Parame & "DETITF_PERMES = " & p_PerMes & " AND "
   g_str_Parame = g_str_Parame & "DETITF_PERANO = " & p_PerAno
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If IsNull(Int(g_rst_Listas!NROCOM)) Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      modsec_gf_BusMov = Int(g_rst_Listas!NROCOM) + 1
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function modsec_gf_GenFec(ByVal p_Fecha As String) As String
   
   modsec_gf_GenFec = Right(p_Fecha, 4) & Mid(p_Fecha, 4, 2) & Left(p_Fecha, 2)

End Function

