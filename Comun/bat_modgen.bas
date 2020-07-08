Attribute VB_Name = "bat_modgen"
Option Explicit

Global Const modgen_g_con_ArcIni = "HIPOTECA.INI"

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CB_SHOWDROPDOWN = &H14F
Public Const WM_USER = &H400
Public Const EM_LIMITTEXT = WM_USER + 21

Public modgen_g_str_RutWin          As String        'Ruta de Windows
Public modgen_g_str_NombPC          As String        'Nombre PC
Public modgen_g_str_CodSuc          As String        'Código de Sucursal
Public modgen_g_str_CodUsu          As String        'Código de Usuario
Public modgen_g_str_NomUsu          As String        'Nombre de Usuario
Public modgen_g_int_TipUsu          As Integer       'Nombre de Usuario
Public modgen_g_int_TpoCad          As Integer
Public modgen_g_str_Mail_Asunto     As String
Public modgen_g_str_Mail_Mensaj     As String
Public modgen_g_int_FlgBat          As Integer
Public modgen_g_str_NomPlt          As String
Public modgen_g_str_NumRev          As String

Public Sub gs_ObtieneRuta()
   Dim r_int_NumFil     As Integer
   Dim r_str_Cadena     As String
   
   'Obtiene Ruta de Windows
   modgen_g_str_RutWin = gf_RutaWindows()
   
   'Verifica existencia de archivo INI
   If Not gf_Existe_Archivo(modgen_g_str_RutWin, modgen_g_con_ArcIni) Then
      MsgBox "No se ha encontrado el archivo de Configuración. Comuníquese con el Dpto. de Sistemas.", vbCritical, modgen_g_str_NomPlt
      End
   End If
End Sub

Public Function gf_RutaWindows() As String
   Dim r_lng_Numero     As Long
   Dim r_str_Cadena     As String

   r_str_Cadena = String$(145, 0)              ' Size Buffer
   r_lng_Numero = GetWindowsDirectory(r_str_Cadena, 145)  ' Make API Call
   r_str_Cadena = Left$(r_str_Cadena, r_lng_Numero)             ' Trim Buffer
    
   If Right$(r_str_Cadena, 1) <> "\" Then      ' Add \ if necessary
      gf_RutaWindows = r_str_Cadena + "\"
   Else
      gf_RutaWindows = r_str_Cadena
   End If
End Function

Public Sub gs_SetFocus(p_Control As Control)
   If p_Control.Enabled And p_Control.Visible Then
      p_Control.SetFocus
   End If
End Sub

Public Function gf_Existe_Archivo(ByVal p_Direct As String, ByVal p_Archiv As String) As Integer
   If Dir$(Trim$(p_Direct) & p_Archiv) <> "" Then
      gf_Existe_Archivo = True
      Exit Function
   End If
   
   gf_Existe_Archivo = False
End Function

Public Function gf_LeeInis(ByVal p_NomIni As String, ByVal p_NomSec As String, ByVal p_Llave As String) As String
   'p_NomIni   -  Ruta y Nombre de Archivo a Leer
   'p_NomSec   -  Nombre de Sección
   'p_Llave    -  Nombre de Llave
   
   Dim r_int_FlgRpt     As Integer
   Dim r_str_RptObt     As String
   
   r_str_RptObt = String(255, " ")
   r_int_FlgRpt = GetPrivateProfileString(p_NomSec, p_Llave, "", r_str_RptObt, Len(r_str_RptObt), p_NomIni)
   
   If r_int_FlgRpt <> 0 Then
      gf_LeeInis = Left$(Trim$(r_str_RptObt), r_int_FlgRpt)
   Else
      gf_LeeInis = ""
   End If
End Function

Public Function GetPrivateProfileString(ByVal p_Seccion As String, ByVal p_Llave As String, ByVal p_Null As String, p_Respuesta As String, ByVal p_TamRpt As Integer, ByVal p_Archivo As String) As Integer
   Dim r_int_NumFil     As Integer
   Dim r_str_Cadena     As String
   Dim r_int_Contad     As Integer

   r_int_NumFil = FreeFile
   Open p_Archivo For Input Shared As r_int_NumFil
   
   Do While Not EOF(r_int_NumFil)
      Line Input #r_int_NumFil, r_str_Cadena
      
      If UCase$(Trim$(r_str_Cadena)) = "[" & UCase$(Trim(p_Seccion)) & "]" Then
         Line Input #r_int_NumFil, r_str_Cadena
         
         Do While InStr(r_str_Cadena, "[") = 0
            If Not r_str_Cadena = "" Then
               If UCase$(Trim$(VBA.Left(r_str_Cadena, Len(p_Llave)))) = UCase$(Trim$(p_Llave)) Then
                  For r_int_Contad = 1 To Len(r_str_Cadena$)
                     If Mid$(r_str_Cadena$, r_int_Contad, 1) = "=" Then
                        p_Respuesta = Mid$(Trim$(r_str_Cadena), r_int_Contad + 1, Len(r_str_Cadena))
                        p_Respuesta = Trim(p_Respuesta)
                        p_TamRpt = Len(p_Respuesta)
                        GetPrivateProfileString = p_TamRpt
                        Close #r_int_NumFil
                        Exit Function
                     End If
                  Next
               End If
            End If
            
            If Not EOF(r_int_NumFil) Then
               Line Input #r_int_NumFil, r_str_Cadena
            Else
               Exit Do
            End If
         Loop
         
         p_Respuesta = p_Null
         p_TamRpt = Len(p_Respuesta)
         GetPrivateProfileString = p_TamRpt
         Close #r_int_NumFil
         Exit Function
      End If
   Loop
   Close #r_int_NumFil
   
   p_Respuesta = p_Null
   p_TamRpt = Len(p_Respuesta)
   GetPrivateProfileString = p_TamRpt
End Function

Public Function gf_NombrePC() As String
   
   Dim r_lng_Numero     As Long
   Dim r_str_Cadena     As String

   r_str_Cadena = String$(145, 32)                        'Size Buffer
   r_lng_Numero = GetComputerName(r_str_Cadena, 145)     ' Make API Call
   
   r_str_Cadena = Trim(r_str_Cadena)
   
   gf_NombrePC = Left(r_str_Cadena, Len(r_str_Cadena) - 1)
End Function

Public Function gf_FormatoFecha(ByVal p_Fecha As String) As String
   If CLng(p_Fecha) = 0 Then
      gf_FormatoFecha = ""
   Else
      gf_FormatoFecha = Right(p_Fecha, 2) & "/" & Mid(p_Fecha, 5, 2) & "/" & Left(p_Fecha, 4)
   End If
End Function

Public Function gf_FormatoHora(ByVal p_Hora As String) As String
   p_Hora = Format(p_Hora, "000000")
   gf_FormatoHora = Left(p_Hora, 2) & ":" & Mid(p_Hora, 3, 2) & ":" & Mid(p_Hora, 5, 2)
End Function

Public Function gf_FormatoNumero(ByVal p_Numero As Double, p_LarTot As Integer, Optional ByVal p_NumDec As Integer) As String
   Dim r_str_Numero     As String
   
   If p_NumDec = 4 Then
      r_str_Numero = Format(p_Numero, "###,###,###,##0.0000")
   Else
      r_str_Numero = Format(p_Numero, "###,###,###,##0.00")
   End If
   r_str_Numero = Space(p_LarTot - Len(r_str_Numero)) & r_str_Numero
   
   gf_FormatoNumero = r_str_Numero
End Function

