Attribute VB_Name = "modgen"
Option Explicit

Global Const modgen_g_con_CadOri = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789"
Global Const modgen_g_con_CadEnc = "@q#$%&k()y?u¡![w*+-<>ñ\_.,:;]=os¬ia/¿"
Global Const modgen_g_con_CadOri_A = "A6rClEgYB7sUDkFvIa8tebd9ZcjuGwTmHJiKxnLXpM1ySoNhÑ2OzP3QR45V0 =[];._-fqW\"
Global Const modgen_g_con_CadEnc_A = "#13t45670\º!|$%&/()¿¡^[]*8+¨ç{}.:,;-_acuhe9i=kr@pzbwf2ñJG?RSMAIOPWVÑKLH<"
Global Const modgen_g_con_NUMERO = "0123456789"
Global Const modgen_g_con_LETRAS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑ"
Global Const modgen_g_con_ArcIni = "HIPOTECA.INI"
Global Const modgen_g_con_ArcAux = "AUXILIAR.INI"
Global Const modgen_g_con_PltPar = "miCasita - Plataforma de Parámetros"
Global Const modgen_g_con_OpeCaj = "miCasita - Plataforma de Caja"
Global Const modgen_g_con_OpeTra = "miCasita - Plataforma de Operaciones"
Global Const modgen_g_con_GesCre = "miCasita - Plataforma de Gestión de Créditos"
Global Const modgen_g_con_GesPry = "miCasita - Plataforma de Gestión de Proyectos Inmobiliarios"
Global Const modgen_g_con_ConRep = "miCasita - Plataforma de Consulta"
Global Const modgen_g_con_EvaCre = "miCasita - Plataforma de Evaluación Crediticia"
Global Const modgen_g_con_EvaLeg = "miCasita - Plataforma de Legal"
Global Const modgen_g_con_AteCli = "miCasita - Plataforma de Atención Comercial"
Global Const modgen_g_con_SimCre = "miCasita - Plataforma de Simulación de Créditos Hipotecarios"
Global Const modgen_g_con_GesAud = "miCasita - Plataforma de Gestión de Auditoria"
Global Const modgen_g_con_GesOcu = "miCasita - Plataforma de Oficial de Cumplimiento"
Global Const modgen_g_con_GesAsi = "miCasita - Plataforma de Gestión de Asistencia"
Global Const modgen_g_con_GesRie = "miCasita - Plataforma de Gestión de Riesgos"
Global Const moddat_g_str_RutLoc = "C:\SBSMIC"
Global Const moddat_g_str_RutAnx = "\\10.10.10.158\Aplica\Micro"
Global Const moddat_g_str_RutFac = "\\10.10.10.158\Aplica\Fact"
Global Const moddat_g_str_RutCFi = "\\10.10.10.158\Aplica\Plant"
Global Const moddat_g_str_RutCor = "\\10.10.10.158\Aplica\Email"
Global Const moddat_g_str_NomCF1 = "plantilla.doc"
Global Const moddat_g_int_PerLim = 18
Global Const moddat_g_str_Msje01 = "La cuota calculada es mayor que la cuota aprobada."
Global Const moddat_g_str_Msje02 = "Ajuste de cuota."

Global Const modgen_g_con_ColVer = &H8000&
Global Const modgen_g_con_ColNeg = &H0&
Global Const modgen_g_con_ColRoj = &HFF&
Global Const modgen_g_con_ColAzu = &HFF0000
Global Const modgen_g_con_ColCya = &H808000
Global Const modgen_g_con_ColBla = &H80000005
Global Const modgen_g_con_ColNar = &H80FF&
Global Const modgen_g_con_ColMag = &HFF00FF
Global Const modgen_g_con_ColAma = &H80FFFF

Public modgen_g_str_NomPlt    As String
Public modgen_g_str_NumRev    As String
Public moddat_g_str_NomEsq    As String
Public moddat_g_str_EntDat    As String
Public moddat_g_str_ClaDat    As String

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public Const CB_SHOWDROPDOWN = &H14F
Public Const WM_USER = &H400
Public Const EM_LIMITTEXT = WM_USER + 21

Public Type tpo_Imprim
   Imprim_ConLen  As String
End Type

Public g_arr_Imprim()      As tpo_Imprim
Public modgen_g_str_Linea  As String
Public modgen_g_int_NumLin As Integer
Public modgen_g_int_NumPag As Integer

'Variables para Actualización de Exe
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

'Inspektor
Public modgen_g_str_rptwebservice As String
'Inspektor

Public modgen_g_int_FlgExc As Integer  'Flag Motivo de Excepcion LMD13102011
Public modgen_g_str_Nombre As String
Public modgen_g_int_NroObs As Integer
Public modgen_g_int_CmbMes As Integer
Public modgen_g_int_FrmPwd As Integer  'LMD 1=MrcAsi 2=DetAsi 3=SegAsi
Public modgen_g_str_HorSel As String   'LMD
Public modgen_g_str_DetObs As String   'LMD
Public modgen_g_int_ObsDia As Integer  'LMD
Public modgen_g_str_ObsMes As String   'LMD
Public modgen_g_int_ObsAno As Integer  'LMD
Public modgen_g_str_TitAsi As String   'LMD

Public Type TIPONOTIFICARICONO
    cbSize              As Long
    hwnd                As Long
    uId                 As Long
    uFlags              As Long
    ucallbackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As TIPONOTIFICARICONO) As Boolean
Public Declare Function WinExec& Lib "kernel32" (ByVal IpCmdLine As String, ByVal nCmdShow As Long)

Global g_Aplica                  As TIPONOTIFICARICONO

Public admusu_g_str_CodUsu    As String
Public admusu_g_str_NomUsu    As String
Public admusu_g_str_CodPlt    As String
Public admusu_g_str_NomPlt    As String
Public admusu_g_str_CodTUs    As String
Public admusu_g_str_NomTUs    As String

Public modvar_g_int_TipPan    As Integer
Public modvar_g_int_TipCom    As Integer
Public modvar_g_int_TipMon    As Integer
Public modvar_g_int_PlaIni    As Integer
Public modvar_g_int_PlaFin    As Integer
Public modgen_g_int_DiaFer    As String

Public Sub gs_CentraForm(p_Formulario As Form)
   p_Formulario.Top = (Screen.Height - p_Formulario.Height) / 2
   p_Formulario.Left = (Screen.Width - p_Formulario.Width) / 2
End Sub

Public Sub gs_SelecTodo(p_Control As Control)
   p_Control.SelStart = 0
   p_Control.SelLength = Len(p_Control.Text)
End Sub

Public Function gf_ValidaCaracter(ByVal p_TeclaPres As Integer, ByVal p_CarAdmite As String) As Integer
   Dim int_l_CarTempo     As Integer
   Dim int_l_FlgEnter     As Integer

   If p_TeclaPres = 8 Then
      gf_ValidaCaracter = p_TeclaPres
      Exit Function
   End If

   If p_TeclaPres = 13 Then
      int_l_FlgEnter = True
   Else
      int_l_FlgEnter = False
   End If

   'Este módulo de la funcion transforma de minusculas a mayusculas.
   If InStr("abcdefghijklmnñopqrstuvwxyz", Chr$(p_TeclaPres)) <> 0 Then
      int_l_CarTempo = Asc(UCase$(Chr$(p_TeclaPres)))
   Else
      int_l_CarTempo = p_TeclaPres
   End If

   'Este modulo valida que caracteres [si seran] o [no seran] aceptados. Según p_FlgExcluye
   If p_CarAdmite <> "" Then
      p_CarAdmite = UCase$(p_CarAdmite)

      If (InStr(p_CarAdmite, Chr$(int_l_CarTempo)) = 0) Then
         int_l_CarTempo = 0
      End If
   End If

   gf_ValidaCaracter = int_l_CarTempo   '***--- Retorno Final.
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

Public Sub gs_LimpiaGrid(p_Grid As MSFlexGrid)
   p_Grid.Rows = 0
End Sub

Public Sub gs_UbiIniGrid(p_Grid As MSFlexGrid)
   p_Grid.Row = 0
   p_Grid.Col = 0
   
   p_Grid.RowSel = 0
   p_Grid.ColSel = p_Grid.Cols - 1
End Sub

Public Sub gs_UbicaGrid(p_Grid As MSFlexGrid, ByVal p_Fila As Integer)
   p_Grid.Row = p_Fila
   p_Grid.Col = 0
   
   p_Grid.RowSel = p_Fila
   p_Grid.ColSel = p_Grid.Cols - 1
End Sub

Public Sub gs_RefrescaGrid(p_Grid As MSFlexGrid)
   p_Grid.Col = 0
   p_Grid.ColSel = p_Grid.Cols - 1
   p_Grid.RowSel = p_Grid.Row
End Sub

Function gf_Seg_Encrip(p_CadEnc As String) As String
   Dim r_str_CadEnc     As String
   Dim r_int_Posici     As Integer
   Dim r_int_Contad     As Integer

   gf_Seg_Encrip = ""
   r_str_CadEnc = ""
   
   For r_int_Contad = 1 To Len(p_CadEnc$)
      r_int_Posici = InStr(modgen_g_con_CadOri, Mid$(p_CadEnc, r_int_Contad, 1))
      
      If r_int_Posici > 0 Then
         r_str_CadEnc = r_str_CadEnc + Mid(modgen_g_con_CadEnc, r_int_Posici, 1)
      End If
   Next r_int_Contad
   
   gf_Seg_Encrip = Trim$(r_str_CadEnc)
End Function

Function gf_Seg_Desenc(p_CadEnc As String) As String
   Dim r_str_CadEnc     As String
   Dim r_int_Posici     As Integer
   Dim r_int_Contad     As Integer

   gf_Seg_Desenc = ""
   r_str_CadEnc = ""
   
   For r_int_Contad = 1 To Len(p_CadEnc)
      r_int_Posici = InStr(modgen_g_con_CadEnc, Mid$(p_CadEnc, r_int_Contad, 1))
      If r_int_Posici > 0 Then
         r_str_CadEnc = r_str_CadEnc + Mid$(modgen_g_con_CadOri, r_int_Posici, 1)
      End If
   Next r_int_Contad
   
   gf_Seg_Desenc = Trim$(r_str_CadEnc)
End Function

Public Sub gs_BuscarCombo(p_Combo As ComboBox, p_Cadena As String)
   Dim r_int_Contad  As Integer
   Dim r_int_Ubicad  As Integer
   
   If Len(Trim$(p_Cadena)) = 0 Then
      Exit Sub
   End If
   
   r_int_Ubicad = False
   
   For r_int_Contad = 0 To p_Combo.ListCount - 1
      p_Combo.ListIndex = r_int_Contad
      
      If Left(p_Combo.Text, Len(p_Cadena)) = p_Cadena Then
         r_int_Ubicad = True
         
         Exit For
      End If
   Next r_int_Contad
   
   If Not r_int_Ubicad Then
      p_Combo.ListIndex = -1
   End If
End Sub

Public Sub gs_ObtieneRuta()
   Dim r_int_NumFil     As Integer
   Dim r_str_Cadena     As String
   
   'Obtiene Ruta de Windows
   modgen_g_str_RutWin = gf_RutaWindows()
   modsec_g_str_RutIni = modgen_g_str_RutWin
   
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

Public Sub gs_TamForm(p_Form As Form, p_TamAnc As Integer, p_TamAlt As Integer)
   If p_Form.WindowState <> vbMinimized And p_Form.WindowState <> vbMaximized Then
      p_Form.Width = p_TamAnc
      p_Form.Height = p_TamAlt
      
      Call gs_CentraForm(p_Form)
   End If
End Sub
   
Public Sub gs_AutoCopia_Exe()
   Dim r_lng_FlagOk     As Long         ' Indicador de Proceso Ok
   Dim r_str_Tempor     As String       ' Temporal para datos
   Dim r_str_Result     As String       ' Archivo de Respuesta
   Dim r_str_Respue     As String       ' Contiene Respuesta si se Termina Programa o No
   Dim r_int_Buffer     As Integer      ' Buffer Libre para Apertura de Archivo
   Dim r_str_RutExe     As String         'Ruta de Versión Actualizada de Ejecutable

   On Error GoTo gs_Error_Autocopia_Exe
   
   'Obtiene Ruta de Ejecutable para Actualización
   r_str_RutExe = gf_LeeInis(modsec_g_str_RutIni & modgen_g_con_ArcAux, "RUTA_SRV", "PATH_EXE")
         
   If r_str_RutExe <> "" Then
      r_str_Tempor = r_str_RutExe + "\UPD_EXE.EXE " + App.Path + " " + App.EXEName
      
      'Para cuando este en Producción debe buscar el UPD_EXE en la misma ruta del ejecutable local
      r_str_Tempor = App.Path + "\UPD_EXE.EXE " + App.Path + " " + App.EXEName
      
      r_lng_FlagOk = Shell(r_str_Tempor, 1)
   End If
   
   If Dir$(CurDir + "\" + App.EXEName + ".PRM") = "" Then
      Exit Sub
   End If

   r_str_Respue = ""
   Do
      r_str_Result = Dir$(CurDir + "\" + App.EXEName + ".PRM")
      DoEvents
      
      If r_str_Result <> "" Then
         r_str_Result = CurDir + "\" + App.EXEName + ".PRM"
         
         If FileLen(r_str_Result) > 0 Then
            r_int_Buffer = FreeFile
            Open r_str_Result For Input Shared As r_int_Buffer
            Input #r_int_Buffer, r_str_Respue
            Close (r_int_Buffer)
            Kill r_str_Result
         Else
            r_str_Respue = "NO RESPUESTA"
         End If
      End If
   Loop While r_str_Respue = ""
    
   Select Case r_str_Respue
      Case "SALIR"
         End
   End Select
   
   Exit Sub

gs_Error_Autocopia_Exe:
    MsgBox "Se ha producido el siguiente error al realizar la Autocopia local." & Chr(13) & Error$, vbCritical, modgen_g_str_NomPlt
    DoEvents
    Exit Sub
End Sub

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

Public Sub gs_BuscarCombo_Item(p_Combo As ComboBox, p_Item As Integer)
   Dim r_int_Contad  As Integer
   Dim r_int_Ubicad  As Integer
   
   r_int_Ubicad = -1
   
   For r_int_Contad = 0 To p_Combo.ListCount - 1
      If p_Item = p_Combo.ItemData(r_int_Contad) Then
         r_int_Ubicad = r_int_Contad
         
         Exit For
      End If
   Next r_int_Contad
   
   p_Combo.ListIndex = r_int_Ubicad
End Sub

Public Sub gs_BuscarCombo_Item_Long(p_Combo As ComboBox, p_Item As Long)
   Dim r_int_Contad  As Integer
   Dim r_int_Ubicad  As Integer
   
   r_int_Ubicad = -1
   
   For r_int_Contad = 0 To p_Combo.ListCount - 1
      If p_Item = p_Combo.ItemData(r_int_Contad) Then
         r_int_Ubicad = r_int_Contad
         
         Exit For
      End If
   Next r_int_Contad
   
   p_Combo.ListIndex = r_int_Ubicad
End Sub

Public Sub gs_SorteaGrid(p_Grid As MSFlexGrid, p_NumCol As Integer, p_TipDat As String)
   Dim r_int_TipDat  As Integer
   
   If p_Grid.Rows = 0 Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Select Case p_TipDat
      Case "C":   r_int_TipDat = flexSortStringAscending
      Case "N":   r_int_TipDat = flexSortNumericAscending
      Case "C-":   r_int_TipDat = flexSortStringDescending
      Case "N-":   r_int_TipDat = flexSortNumericDescending
   End Select
   
   p_Grid.Col = p_NumCol
   p_Grid.Sort = r_int_TipDat
   
   Call gs_UbiIniGrid(p_Grid)
   
   Screen.MousePointer = 0
End Sub

Public Function gf_Valida_RUC(ByVal p_Numero, p_DigVer As String) As Integer
   Dim r_int_Contad As Integer
   Dim r_int_Sumato As Integer
   Dim r_int_RestoN As Integer
   Dim r_lng_LarNum As Long
   Dim r_str_DigVer As String
   Dim r_str_Digito As String
   Dim r_str_Residu As Integer

   p_Numero = Trim(p_Numero)
   r_int_Sumato = 0
   r_str_Digito = "5432765432"
   r_lng_LarNum = Len(Trim$(p_Numero))
   r_str_Digito = Right$(r_str_Digito, r_lng_LarNum)
   
   If Not IsNumeric(p_Numero) Then
      gf_Valida_RUC = False
      Exit Function
   End If
   
   If Len(Trim(p_Numero)) <> 11 Then
      gf_Valida_RUC = False
      Exit Function
   End If
   
   For r_int_Contad = 1 To r_lng_LarNum - 1
       r_int_Sumato = r_int_Sumato + CInt(Mid$(r_str_Digito, r_int_Contad, 1)) * CInt(Mid$(p_Numero, r_int_Contad, 1))
   Next r_int_Contad
   
   r_str_Residu = r_int_Sumato Mod 11
   r_int_RestoN = 11 - r_str_Residu  '(r_int_Sumato + Int(r_int_Sumato / 11) * 11)

   Select Case r_int_RestoN
      Case 10:    r_str_DigVer = "0"
      Case 11:    r_str_DigVer = "1"
      Case Else:  r_str_DigVer = CStr(r_int_RestoN)
   End Select

   If UCase$(p_DigVer) = r_str_DigVer Then
      gf_Valida_RUC = True
   Else
      gf_Valida_RUC = False
   End If
End Function

Public Function gf_Busca_Arregl(p_Arregl() As moddat_tpo_Genera, p_Codigo As String) As Integer
   Dim r_int_Contad  As Integer
   Dim r_int_FlgUbi  As Integer
   
   gf_Busca_Arregl = 0
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      If p_Arregl(r_int_Contad).Genera_Codigo = p_Codigo Then
         gf_Busca_Arregl = r_int_Contad
         Exit For
      End If
   Next r_int_Contad
End Function

Public Function gf_Truncar_Numero(ByVal p_Numero As Double, ByVal p_NumDec As Integer) As String
   Dim r_str_Numero  As String
   Dim r_int_PtoDec  As Integer
   Dim r_str_Entero  As String
   Dim r_str_Decima  As String
   
   r_str_Numero = CStr(p_Numero)
   
   r_int_PtoDec = InStr(r_str_Numero, ".")
   
   If r_int_PtoDec > 0 Then
      r_str_Entero = Left(r_str_Numero, r_int_PtoDec - 1)
      r_str_Decima = Mid(r_str_Numero, r_int_PtoDec + 1, p_NumDec)
   Else
      r_str_Entero = r_str_Numero
      r_str_Decima = "00"
   End If
   
   gf_Truncar_Numero = r_str_Entero & "." & r_str_Decima
End Function

Public Function gf_NueImp_Numero(ByVal p_Numero As Double) As String
     
   Dim r_str_Entero  As String
   Dim r_str_Decima  As String
   Dim r_int_lonNum  As Integer
   Dim r_str_VarAux  As String
   Dim r_str_PriDig  As String
   Dim r_str_UltDig  As String
   Dim r_int_Digito  As Integer
   
   r_str_Entero = 0
            
   r_str_VarAux = Right(Format(p_Numero, "0.00"), 2)
   r_int_lonNum = Len(CStr(Format(p_Numero, "0.00")))
   r_str_UltDig = Right(r_str_VarAux, 1)
   r_str_PriDig = Left(r_str_VarAux, 1)
   
   If r_str_UltDig < 5 Then
      r_str_UltDig = 0
   Else
      r_str_UltDig = 5
   End If
   
   r_str_Entero = Left(CStr(Format(p_Numero, "0.00")), (r_int_lonNum) - 3)
   r_str_Decima = r_str_PriDig & r_str_UltDig

   
   gf_NueImp_Numero = r_str_Entero & "." & r_str_Decima
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

Public Function gf_FormatoTipCam(ByVal p_Numero As Double, p_LarTot As Integer, Optional ByVal p_NumDec As Integer) As String
   Dim r_str_Numero     As String
   
   r_str_Numero = Format(p_Numero, "###,###,##0.0000")
   r_str_Numero = Space(p_LarTot - Len(r_str_Numero)) & r_str_Numero
   
   gf_FormatoTipCam = r_str_Numero
End Function

Public Sub gs_BuscarCombo_Text(p_Combo As ComboBox, ByVal p_Codigo As String, ByVal p_NumPos As Integer)
Dim r_int_Contad     As Integer
   
   For r_int_Contad = 0 To p_Combo.ListCount - 1
      p_Combo.ListIndex = r_int_Contad
      
'      If UCase(Trim(p_Combo.Text)) = UCase(Trim(p_Codigo)) Then
'         Exit Sub
'      End If
         
      If p_NumPos = -1 Then
         If UCase(Trim(p_Combo.Text)) = UCase(Trim(p_Codigo)) Then
            Exit Sub
         End If
      Else
         If Left(p_Combo.Text, p_NumPos) = p_Codigo Then
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   If p_NumPos = -1 Then
      p_Combo.ListIndex = -1
   End If
End Sub

Public Function gs_CalcularEdad(p_Fecha1 As Date, p_Fecha2 As Date) As String
    Dim r_dbl_NumDia    As Double
    Dim r_int_NumAno    As Integer
    Dim r_int_NumMes    As Integer
    
    r_dbl_NumDia = p_Fecha2 - p_Fecha1
    r_int_NumAno = Int(r_dbl_NumDia / 365.25)
    
    r_dbl_NumDia = r_dbl_NumDia - (365.25 * r_int_NumAno)
    
    r_int_NumMes = Int(r_dbl_NumDia / 30.4375)
    
    gs_CalcularEdad = Format(r_int_NumAno, "00") & Format(r_int_NumMes, "00")
End Function

Public Sub gs_Imprim_ComPag()
   Dim r_int_Contad     As Integer
   Dim r_int_ConErr     As Integer
   Dim r_int_NumFil     As Integer
   
   On Error GoTo gs_Imprim_ComPag_Error
   
   r_int_ConErr = 0
   GoTo gs_Imprim_ComPag_Impresion
   
gs_Imprim_ComPag_Impresion:
   
   'Printer.Font.Name = "Courier New"
   'Printer.Font.Size = 8
   'Printer.Orientation = vbPRORPortrait
   'Printer.Copies = 1
   
   r_int_NumFil = FreeFile
   
   Open "LPT1" For Output As r_int_NumFil
   
   Print #r_int_NumFil, Chr(27); "@"
   
   For r_int_Contad = 1 To UBound(g_arr_Imprim)
      If g_arr_Imprim(r_int_Contad).Imprim_ConLen = "SP" Then
         Print #r_int_NumFil, Chr(12)
      Else
         Print #r_int_NumFil, Chr(27); "!"; Chr(4); g_arr_Imprim(r_int_Contad).Imprim_ConLen
      End If
   Next r_int_Contad
   
   Print #r_int_NumFil, Chr(12)
   
   Close #r_int_NumFil
   
   Exit Sub
   
gs_Imprim_ComPag_Error:
   r_int_ConErr = r_int_ConErr + 1
   MsgBox "Se ha producido un error en la Impresora. Por favor verfique las conexiones.", vbExclamation, modgen_g_str_NomPlt
   
   If r_int_ConErr = 4 Then
      If MsgBox("¿Desea seguir reintentando la Impresión del Voucher?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      Else
         r_int_ConErr = 0
      End If
   End If
   GoTo gs_Imprim_ComPag_Impresion

End Sub

Public Sub gs_Imprim_ComPag_1()
   Dim r_int_Contad     As Integer
   Dim r_int_ConErr     As Integer
   
   On Error GoTo gs_Imprim_ComPag_Error
   
   r_int_ConErr = 0
   GoTo gs_Imprim_ComPag_Impresion
   
gs_Imprim_ComPag_Impresion:
   
   'Printer.Font.Name = "Courier New"
   'Printer.Font.Size = 8
   'Printer.Orientation = vbPRORPortrait
   'Printer.Copies = 1
   
   For r_int_Contad = 1 To UBound(g_arr_Imprim)
      If g_arr_Imprim(r_int_Contad).Imprim_ConLen = "SP" Then
         Printer.NewPage
      Else
         Printer.Print g_arr_Imprim(r_int_Contad).Imprim_ConLen
      End If
   Next r_int_Contad
   
   Printer.EndDoc
   
   Exit Sub
   
gs_Imprim_ComPag_Error:
   r_int_ConErr = r_int_ConErr + 1
   MsgBox "Se ha producido un error en la Impresora. Por favor verfique las conexiones.", vbExclamation, modgen_g_str_NomPlt
   
   If r_int_ConErr = 4 Then
      If MsgBox("¿Desea seguir reintentando la Impresión del Voucher?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      Else
         r_int_ConErr = 0
      End If
   End If
   GoTo gs_Imprim_ComPag_Impresion

End Sub

Public Sub gs_LinImp(p_Conten As String)
   ReDim Preserve g_arr_Imprim(UBound(g_arr_Imprim) + 1)
   g_arr_Imprim(UBound(g_arr_Imprim)).Imprim_ConLen = p_Conten
End Sub

Public Sub gs_Imprim(ByVal p_TipLet As Integer, ByVal p_TamLet As Double, p_Orienta As Integer)
   Dim r_int_Contad     As Integer
   Dim r_int_ConErr     As Integer
   
   On Error GoTo gs_Imprim_Error
   
   r_int_ConErr = 0
   GoTo gs_Imprim_Impresion
   
gs_Imprim_Impresion:
   
   For r_int_Contad = 1 To UBound(g_arr_Imprim)
      If g_arr_Imprim(r_int_Contad).Imprim_ConLen = "SP" Then
         Printer.NewPage
      Else
         Printer.Print g_arr_Imprim(r_int_Contad).Imprim_ConLen
      End If
   Next r_int_Contad
   
   Printer.EndDoc
   Exit Sub
   
gs_Imprim_Error:
   r_int_ConErr = r_int_ConErr + 1
   MsgBox "Se ha producido un error en la Impresora. Por favor verfique las conexiones.", vbExclamation, modgen_g_str_NomPlt
   
   If r_int_ConErr = 4 Then
      If MsgBox("¿Desea seguir reintentando la Impresión del Voucher?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      Else
         r_int_ConErr = 0
      End If
   End If
   GoTo gs_Imprim_Impresion

End Sub

Public Function gf_FormatoNumEnt(ByVal p_Numero As Double, ByVal p_Cantid As Integer) As String
   Dim r_str_Numero     As String
   
   gf_FormatoNumEnt = ""
   
   r_str_Numero = Format(p_Numero, "###,###,##0")
   
   gf_FormatoNumEnt = Space(p_Cantid - Len(r_str_Numero)) & r_str_Numero
End Function

Public Sub admusu_gs_Carga_Plataf(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_Situac As Integer)
   ReDim p_Arregl(0)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM SEG_PLTMAE "
   
   If p_Situac > 0 Then
      g_str_Parame = g_str_Parame & "WHERE PLTMAE_SITUAC = " & CStr(p_Situac) & " "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY PLTMAE_CODIGO ASC"

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
      p_Combo.AddItem Trim(g_rst_Listas!PLTMAE_CODIGO) & " - " & Trim$(g_rst_Listas!PLTMAE_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!PLTMAE_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!PLTMAE_DESCRI & "")
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub admusu_gs_Carga_TipUsu(ByVal p_CodPlt As String, p_Combo As ComboBox)
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE PARDES_CODGRP = '351' AND PARDES_CODITE <> '000000' AND PARDES_SITUAC = 1 AND SUBSTR(PARDES_DESCRI,1,6) = '" & p_CodPlt & "' "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   'Si no hay datos registrados
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(Mid(g_rst_Listas!PARDES_DESCRI, 10))
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!PARDES_CODITE)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub admusu_gs_Carga_TipUsu_Plataf(ByVal p_CodPlt As String, p_Combo As ComboBox)
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM SEG_MAETIP WHERE MAETIP_SITUAC = 1 AND MAETIP_CODPLT = '" & p_CodPlt & "' "
   g_str_Parame = g_str_Parame & "ORDER BY MAETIP_CODIGO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   'Si no hay datos registrados
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!MAETIP_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!MAETIP_CODIGO)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function admusu_gf_Consulta_NomPlt(ByVal p_CodPlt As String) As String
   admusu_gf_Consulta_NomPlt = ""
   
   g_str_Parame = "SELECT * FROM SEG_PLTMAE WHERE PLTMAE_CODIGO = '" & p_CodPlt & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   
   admusu_gf_Consulta_NomPlt = Trim$(g_rst_Listas!PLTMAE_DESCRI)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub admusu_gf_Verifica_Caducidad()
   Dim r_str_FecCam     As String
   Dim r_int_TpoCad     As Integer
   Dim r_int_FlgPri     As Integer
   
   
   r_int_FlgPri = 0
   GoTo admusu_gf_Verifica_Caducidad_Inicia
   
admusu_gf_Verifica_Caducidad_Inicia:
   r_str_FecCam = ""
   r_int_TpoCad = 0
   
   g_str_Parame = "SELECT * FROM SEG_USUMAE_LOG WHERE USUMAELOG_CODIGO = '" & modgen_g_str_CodUsu & "' AND USUMAELOG_ACCION = 5 ORDER BY LOGFECCRE DESC, LOGHORCRE DESC"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_str_FecCam = CStr(g_rst_Princi!LOGFECCRE)
      r_int_TpoCad = CInt(date - CDate(gf_FormatoFecha(r_str_FecCam)))
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   GoTo admusu_gf_Verifica_Caducidad_Verifica

admusu_gf_Verifica_Caducidad_Verifica:
   If modgen_g_int_TpoCad > 0 Then
      If r_int_TpoCad >= modgen_g_int_TpoCad Or Len(Trim(r_str_FecCam)) = 0 Then
         If r_int_FlgPri = 0 Then
            MsgBox "Debe cambiar la Contraseña de Acceso.", vbInformation, modgen_g_str_NomPlt
            
            frm_IdeUsu_02.Show 1
            r_int_FlgPri = 1
            
            GoTo admusu_gf_Verifica_Caducidad_Inicia
         Else
            Call gs_Desconecta_Servidor
            End
         End If
      End If
   End If
End Sub

Public Function gf_ComaDecimal(ByVal p_Numero As String, ByVal p_Decima As Integer) As String
   Dim r_int_Posici     As Integer
   
   r_int_Posici = InStr(1, p_Numero, ".")
   
   If r_int_Posici = 0 Then
      gf_ComaDecimal = p_Numero & "," & String(p_Decima, "0")
   Else
      gf_ComaDecimal = Mid(p_Numero, 1, r_int_Posici - 1) & "," & Mid(p_Numero, r_int_Posici + 1)
   End If
End Function

Private Function gf_Unidades(ByVal p_Numero As Long, Optional ByVal bMl As Boolean = False) As String
   Select Case p_Numero
      Case 1
         If bMl Then
            gf_Unidades = "un"
         Else
            gf_Unidades = "uno"
         End If
         
      Case 2:  gf_Unidades = "dos"
      Case 3:  gf_Unidades = "tres"
      Case 4:  gf_Unidades = "cuatro"
      Case 5:  gf_Unidades = "cinco"
      Case 6:  gf_Unidades = "seis"
      Case 7:  gf_Unidades = "siete"
      Case 8:  gf_Unidades = "ocho"
      Case 9:  gf_Unidades = "nueve"
   End Select
End Function

Private Function gf_Decenas(ByVal p_Numero As Long, Optional ByVal bMl As Boolean = False) As String
    
    If p_Numero < 10 Then
        gf_Decenas = gf_Unidades(p_Numero)
    ElseIf p_Numero = 10 Then
        gf_Decenas = "diez"
    ElseIf p_Numero = 11 Then
        gf_Decenas = "once"
    ElseIf p_Numero = 12 Then
        gf_Decenas = "doce"
    ElseIf p_Numero = 13 Then
        gf_Decenas = "trece"
    ElseIf p_Numero = 14 Then
        gf_Decenas = "catorce"
    ElseIf p_Numero = 15 Then
        gf_Decenas = "quince"
    ElseIf p_Numero = 16 Then
        gf_Decenas = "dieciséis"
    ElseIf p_Numero = 17 Then
        gf_Decenas = "diecisiete"
    ElseIf p_Numero = 18 Then
        gf_Decenas = "dieciocho"
    ElseIf p_Numero = 19 Then
        gf_Decenas = "diecinueve"
    'ElseIf p_Numero >= 16 And p_Numero < 20 Then
    '    gf_Decenas = "diez y " & gf_Unidades(p_Numero - 10)
    ElseIf p_Numero >= 20 And p_Numero < 30 Then
        gf_Decenas = "veint" & IIf((p_Numero - 20) = 0, "e", "i" & gf_Unidades(p_Numero - 20, bMl))
    ElseIf p_Numero >= 30 And p_Numero < 40 Then
        gf_Decenas = "treinta" & IIf((p_Numero - 30) = 0, "", " y " & gf_Unidades(p_Numero - 30, bMl))
    ElseIf p_Numero >= 40 And p_Numero < 50 Then
        gf_Decenas = "cuarenta" & IIf((p_Numero - 40) = 0, "", " y " & gf_Unidades(p_Numero - 40, bMl))
    ElseIf p_Numero >= 50 And p_Numero < 60 Then
        gf_Decenas = "cincuenta" & IIf((p_Numero - 50) = 0, "", " y " & gf_Unidades(p_Numero - 50, bMl))
    ElseIf p_Numero >= 60 And p_Numero < 70 Then
        gf_Decenas = "sesenta" & IIf((p_Numero - 60) = 0, "", " y " & gf_Unidades(p_Numero - 60, bMl))
    ElseIf p_Numero >= 70 And p_Numero < 80 Then
        gf_Decenas = "setenta" & IIf((p_Numero - 70) = 0, "", " y " & gf_Unidades(p_Numero - 70, bMl))
    ElseIf p_Numero >= 80 And p_Numero < 90 Then
        gf_Decenas = "ochenta" & IIf((p_Numero - 80) = 0, "", " y " & gf_Unidades(p_Numero - 80, bMl))
    ElseIf p_Numero >= 90 And p_Numero < 100 Then
        gf_Decenas = "noventa" & IIf((p_Numero - 90) = 0, "", " y " & gf_Unidades(p_Numero - 90, bMl))
    End If

End Function

Private Function gf_Centenas(ByVal p_Numero As Long, Optional ByVal bMl As Boolean = False) As String
    
    If p_Numero < 100 Then
        gf_Centenas = gf_Decenas(p_Numero, bMl)
    ElseIf p_Numero >= 100 And p_Numero < 200 Then
        gf_Centenas = "cien" & IIf((p_Numero - 100) = 0, "", "to " & gf_Decenas(p_Numero - 100, bMl))
    ElseIf p_Numero >= 200 And p_Numero < 300 Then
        gf_Centenas = "doscientos" & IIf((p_Numero - 200) = 0, "", " " & gf_Decenas(p_Numero - 200, bMl))
    ElseIf p_Numero >= 300 And p_Numero < 400 Then
        gf_Centenas = "trecientos" & IIf((p_Numero - 300) = 0, "", " " & gf_Decenas(p_Numero - 300, bMl))
    ElseIf p_Numero >= 400 And p_Numero < 500 Then
        gf_Centenas = "cuatrocientos" & IIf((p_Numero - 400) = 0, "", " " & gf_Decenas(p_Numero - 400, bMl))
    ElseIf p_Numero >= 500 And p_Numero < 600 Then
        gf_Centenas = "quinientos" & IIf((p_Numero - 500) = 0, "", " " & gf_Decenas(p_Numero - 500, bMl))
    ElseIf p_Numero >= 600 And p_Numero < 700 Then
        gf_Centenas = "seiscientos" & IIf((p_Numero - 600) = 0, "", " " & gf_Decenas(p_Numero - 600, bMl))
    ElseIf p_Numero >= 700 And p_Numero < 800 Then
        gf_Centenas = "setecientos" & IIf((p_Numero - 700) = 0, "", " " & gf_Decenas(p_Numero - 700, bMl))
    ElseIf p_Numero >= 800 And p_Numero < 900 Then
        gf_Centenas = "ochocientos" & IIf((p_Numero - 800) = 0, "", " " & gf_Decenas(p_Numero - 800, bMl))
    ElseIf p_Numero >= 900 And p_Numero < 1000 Then
        gf_Centenas = "novecientos" & IIf((p_Numero - 900) = 0, "", " " & gf_Decenas(p_Numero - 900, bMl))
    End If

End Function

Private Function gf_Millares(ByVal p_Numero As Long) As String
    If p_Numero < 1000 Then
        gf_Millares = gf_Centenas(p_Numero)
    ElseIf p_Numero >= 1000 And p_Numero < 2000 Then
        gf_Millares = "mil " & gf_Centenas(p_Numero - Int(p_Numero / 1000) * 1000)
    Else
        gf_Millares = gf_Centenas(Int(p_Numero / 1000), True) & " mil " & gf_Centenas(p_Numero - Int(p_Numero / 1000) * 1000)
    End If
End Function

Public Function gf_Convertir_NumLet(ByVal p_Numero As Long) As String
   Dim r_str_Litera As String
        
   'Conseguimos el literal
   r_str_Litera = gf_Millares(p_Numero)
    
   'Cambiamos la primera letra a mayúscula
   gf_Convertir_NumLet = Trim(UCase(Left(r_str_Litera, 1)) & Mid(r_str_Litera, 2))
End Function

Public Function admusu_gf_Consulta_TipUsu(ByVal p_TipUsu As String) As String
   admusu_gf_Consulta_TipUsu = ""

   g_str_Parame = "SELECT * FROM SEG_MAETIP WHERE MAETIP_CODIGO = '" & p_TipUsu & "' "
   g_str_Parame = g_str_Parame & "ORDER BY MAETIP_CODIGO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   'Si no hay datos registrados
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   
   admusu_gf_Consulta_TipUsu = Trim(g_rst_Listas!MAETIP_DESCRI)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function GetBoiler(ByVal sFile As String) As String
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function

Public Function gf_ValidarEmail(ByVal Email As String) As Boolean
Dim i As Integer, iLen As Integer, caracter As String
Dim pos As Integer, bp As Boolean, iPos As Integer, iPos2 As Integer

   On Local Error GoTo Err_Sub
   
   gf_ValidarEmail = False
   Email = Trim$(Email)
   If Email = vbNullString Then
      Exit Function
   End If
   Email = LCase$(Email)
   iLen = Len(Email)
    
   For i = 1 To iLen
      caracter = Mid(Email, i, 1)
      If (Not (caracter Like "[a-z]")) And (Not (caracter Like "[0-9]")) Then
         If InStr(1, "_-" & "." & "@", caracter) > 0 Then
            If bp = True Then
               Exit Function
            Else
               bp = True
               If i = 1 Or i = iLen Then
                  Exit Function
               End If
               If caracter = "@" Then
                  If iPos = 0 Then
                      iPos = i
                  Else
                      Exit Function
                  End If
               End If
               If caracter = "." Then
                  iPos2 = i
               End If
            End If
         Else
            Exit Function
         End If
      Else
         bp = False
      End If
   Next i
   If iPos = 0 Or iPos2 = 0 Then
      Exit Function
   End If
    
   If iPos2 < iPos Then
      Exit Function
   End If
   
   gf_ValidarEmail = True
   Exit Function
   
Err_Sub:
    On Local Error Resume Next
    gf_ValidarEmail = False
End Function


'-------------------------------------------------------------------------------------------------------------------
'Realizado por :  Luana Martinez de la Flor
'F. Creación   :  13-10-2011
'Objetivo      :  Cadena que muestra campos de la tabla de Excepciones (TRA_SEGEXC)
'Parámetros    :  -
'-------------------------------------------------------------------------------------------------------------------
Public Function modgen_gf_Buscar_Excepc() As String
   g_str_Parame = "SELECT SEGEXC_CODINS,SEGEXC_DESCRI,SEGEXC_TIPAUT, T.SEGFECCRE,T.SEGHORCRE, TRIM(NVL(PARDES_DESCRI,0)) AS PARDES_DESCRI "
   g_str_Parame = g_str_Parame & "FROM TRA_SEGEXC T "
   g_str_Parame = g_str_Parame & "LEFT JOIN MNT_PARDES ON (SEGEXC_MOTEXC=PARDES_CODITE AND PARDES_CODGRP=42) "
   g_str_Parame = g_str_Parame & "WHERE SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   modgen_gf_Buscar_Excepc = g_str_Parame
End Function

'-------------------------------------------------------------------------------------------------------------------
'Realizado por :  Rafael Durand
'F. Creación   :  18-04-2013
'Objetivo      :  Calcula el valor de la cuota segun cronograma
'Parámetros    :
'-------------------------------------------------------------------------------------------------------------------
Public Sub modgen_gf_Buscar_CuotaMensual(ByRef p_Arr_TNC() As String, ByRef p_Arr_TC() As String, ByVal p_int_CuoDbl As Integer, ByVal p_int_PerGra As Integer, ByVal p_int_TipPrd As Integer, ByRef p_dbl_CuoMen As Double, ByRef p_int_CuoIncPBP As Double, ByRef p_dbl_IngReq As Double, ByVal p_str_CodPrd As String, ByVal p_str_CodSub As String)
'******* Parametros de entrada
'01- p_Arr_TNC()  : Variable Tipo Matriz que recibira el cronograma TNC Cliente
'02- p_Arr_TC()   : Variable Tipo Matriz que recibira el cronograma TC Cliente
'03- p_int_CuoDbl : Variable que recibira el Periodo de Gracia
'04- p_int_PerGra : Variable que recibira el Periodo de Gracia
'05- p_int_TipPrd : Variable que recibira el Tipo de Producto

   p_dbl_CuoMen = 0
   p_int_CuoIncPBP = 0
   p_dbl_IngReq = 0
   
   'Determina cuota mensual TNC
   Select Case p_int_CuoDbl
      Case 1:        'NO
         p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 2, 9), "###,###,##0.00") & " "
      Case 2:        'JULIO
         If Month(p_Arr_TNC(p_int_PerGra + 2, 2)) <> 7 Then
            p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 2, 9), "###,###,##0.00") & " "
         Else
            p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 3, 9), "###,###,##0.00") & " "
         End If
      Case 3:        'DICIEMBRE
         If Month(p_Arr_TNC(p_int_PerGra + 2, 2)) <> 12 Then
            p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 2, 9), "###,###,##0.00") & " "
         Else
            p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 3, 9), "###,###,##0.00") & " "
         End If
      Case 4:        'JULIO Y DICIEMBRE
         If (Month(p_Arr_TNC(p_int_PerGra + 2, 2)) <> 7) And (Month(p_Arr_TNC(p_int_PerGra + 2, 2)) <> 12) Then
            p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 2, 9), "###,###,##0.00") & " "
         Else
            p_dbl_CuoMen = Format(p_Arr_TNC(p_int_PerGra + 3, 9), "###,###,##0.00") & " "
         End If
   End Select
   
   If p_int_TipPrd = 1 Then
      'MIVIVIENDA
      If UBound(p_Arr_TC) > 0 Then
         p_int_CuoIncPBP = Format(p_dbl_CuoMen + (p_Arr_TC(1, 7) / 6), "###,###,##0.00")
      Else
         p_int_CuoIncPBP = Format(p_dbl_CuoMen, "###,###,##0.00")
      End If
      
      
      If p_str_CodPrd = "004" Then
         p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
      End If
      If p_str_CodPrd = "007" Then
         If p_str_CodSub = "024" Then
            p_dbl_IngReq = Format(p_int_CuoIncPBP / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_int_CuoIncPBP / 40 * 100, "###,##0.00")
         End If
      End If
      If (p_str_CodPrd = "009") Or (p_str_CodPrd = "010") Or (p_str_CodPrd = "012") Then
         p_dbl_IngReq = Format(p_int_CuoIncPBP / 30 * 100, "###,##0.00")
      End If
      If p_str_CodPrd = "013" Or p_str_CodPrd = "014" Or p_str_CodPrd = "015" Or p_str_CodPrd = "017" Or p_str_CodPrd = "018" Then
         If p_str_CodSub = "005" Then
            p_dbl_IngReq = Format(p_int_CuoIncPBP / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_int_CuoIncPBP / 40 * 100, "###,##0.00")
         End If
      End If
      If p_str_CodPrd = "016" Then
         If p_str_CodSub = "007" Then
            p_dbl_IngReq = Format(p_int_CuoIncPBP / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_int_CuoIncPBP / 40 * 100, "###,##0.00")
         End If
      End If
      
   ElseIf p_int_TipPrd = 2 Then
      'MICASITA
      If (p_str_CodPrd = "002") Then
         If p_str_CodSub = "004" Then
            p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
         Else
            If p_dbl_CuoMen <= 600 Then
               p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00") & " "
            ElseIf p_dbl_CuoMen <= 1400 Then
               p_dbl_IngReq = Format(p_dbl_CuoMen / 35 * 100, "###,##0.00") & " "
            Else
               p_dbl_IngReq = Format(p_dbl_CuoMen / 40 * 100, "###,##0.00") & " "
            End If
         End If
      End If
      If p_str_CodPrd = "011" Then
         If p_str_CodSub = "005" Or p_str_CodSub = "014" Or p_str_CodSub = "015" Or p_str_CodSub = "016" Then
            p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
         Else
            If p_dbl_CuoMen <= 600 Then
               p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00") & " "
            ElseIf p_dbl_CuoMen <= 1400 Then
               p_dbl_IngReq = Format(p_dbl_CuoMen / 35 * 100, "###,##0.00") & " "
            Else
               p_dbl_IngReq = Format(p_dbl_CuoMen / 40 * 100, "###,##0.00") & " "
            End If
         End If
      End If
      
   ElseIf p_int_TipPrd = 3 Then
      'MIVIVIENDA 1 CALENDARIO
      If p_str_CodPrd = "019" Or p_str_CodPrd = "021" Or p_str_CodPrd = "023" Then
         If p_str_CodSub = "005" Then
            p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_dbl_CuoMen / 40 * 100, "###,##0.00")
         End If
      ElseIf p_str_CodPrd = "022" Then
         If p_str_CodSub = "005" Or p_str_CodSub = "006" Or p_str_CodSub = "007" Or p_str_CodSub = "008" Or p_str_CodSub = "029" Or p_str_CodSub = "029" Then
            p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_dbl_CuoMen / 40 * 100, "###,##0.00")
         End If
      ElseIf p_str_CodPrd = "024" Then
         If p_str_CodSub = "005" Or p_str_CodSub = "006" Then
            p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_dbl_CuoMen / 40 * 100, "###,##0.00")
         End If
      ElseIf p_str_CodPrd = "025" Then
         If p_str_CodSub = "005" Or p_str_CodSub = "010" Or p_str_CodSub = "011" Or p_str_CodSub = "020" Or p_str_CodSub = "021" Or p_str_CodSub = "022" Then
            p_dbl_IngReq = Format(p_dbl_CuoMen / 30 * 100, "###,##0.00")
         Else
            p_dbl_IngReq = Format(p_dbl_CuoMen / 40 * 100, "###,##0.00")
         End If
      End If
   End If
   
End Sub

Public Function gf_Seg_ValLogin(p_TxtUsu As TextBox, p_TxtPwd As TextBox) As Boolean
Dim r_int_ConErr   As String
   
   r_int_ConErr = 0
   gf_Seg_ValLogin = False

   If Len(Trim(p_TxtUsu.Text)) = 0 Then
      MsgBox "Debe ingresar el Código del Usuario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtUsu)
      gf_Seg_ValLogin = False
      Exit Function
   End If

   If Len(Trim(p_TxtPwd.Text)) = 0 Then
      MsgBox "Debe ingresar la Contraseña del Usuario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd)
      gf_Seg_ValLogin = False
      Exit Function
   End If

   g_str_Parame = "SELECT * FROM SEG_USUMAE WHERE USUMAE_CODIGO = '" & p_TxtUsu.Text & "' AND USUMAE_SITUAC = 1"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   'Verificación de Usuario
   'Si no hay datos registrados
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing

      MsgBox "El Usuario no está registrado en la base de datos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtUsu)

      r_int_ConErr = r_int_ConErr + 1

      If r_int_ConErr = 3 Then
         Call gs_Desconecta_Servidor
         End
      End If

      Exit Function
   End If

   'Verificación de Contraseña
   g_rst_Princi.MoveFirst

   If gf_Seg_Desenc(g_rst_Princi!USUMAE_CONTRA) <> p_TxtPwd.Text Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing

      MsgBox "La Contraseña es incorrecta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd)

      r_int_ConErr = r_int_ConErr + 1

      If r_int_ConErr = 3 Then
         Call gs_Desconecta_Servidor
         End
      End If
      Exit Function
   End If

   modgen_g_str_CodUsu = p_TxtUsu
   modgen_g_str_NomUsu = Trim(g_rst_Princi!USUMAE_NOMBRE)
   modgen_g_int_TpoCad = g_rst_Princi!USUMAE_TPOCAD

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Verificación de Acceso a la Plataforma
   g_str_Parame = "SELECT * FROM SEG_USUTIP WHERE USUTIP_CODUSU = '" & p_TxtUsu.Text & "' AND USUTIP_CODPLT = '" & UCase(App.EXEName) & "' AND USUTIP_SITUAC = 1"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing

      MsgBox "El Usuario no tiene acceso a esta Plataforma.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtUsu)

      r_int_ConErr = r_int_ConErr + 1

      If r_int_ConErr = 3 Then
         Call gs_Desconecta_Servidor
         End
      End If

      Exit Function
   End If

   modgen_g_int_TipUsu = CInt(g_rst_Princi!USUTIP_TIPUSU)
   gf_Seg_ValLogin = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call admusu_gf_Verifica_Caducidad
End Function

Public Function gf_Seg_CamClave(ByVal p_TxtPwd_Act As TextBox, ByVal p_TxtPwd_Nue As TextBox, ByVal p_TxtPwd_Rep As TextBox) As Boolean
Dim r_str_MsjFmt  As String

   gf_Seg_CamClave = False
   r_str_MsjFmt = "La contraseña debe tener mínimo 6 caracteres, considerando al menos un número y una letra."
   
   If Len(Trim(p_TxtPwd_Act.Text)) = 0 Then
      MsgBox "La contraseña actual esta vacía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd_Act)
      gf_Seg_CamClave = False
      Exit Function
   End If

   If gf_Seg_FormatoPwd(p_TxtPwd_Nue.Text) = False Then
      MsgBox r_str_MsjFmt, vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd_Nue)
      gf_Seg_CamClave = False
      Exit Function
   End If
   
   If gf_Seg_FormatoPwd(p_TxtPwd_Rep.Text) = False Then
      MsgBox r_str_MsjFmt, vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd_Rep)
      gf_Seg_CamClave = False
      Exit Function
   End If

   If p_TxtPwd_Act.Text = p_TxtPwd_Nue.Text Then
      MsgBox "La contraseña nueva no puede ser igual a la actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd_Nue)
      gf_Seg_CamClave = False
      Exit Function
   End If

   If p_TxtPwd_Nue.Text <> p_TxtPwd_Rep.Text Then
      MsgBox "La confirmación de la contraseña no coincide con la contraseña nueva.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd_Rep)
      gf_Seg_CamClave = False
      Exit Function
   End If

   g_str_Parame = "SELECT USUMAE_CONTRA FROM SEG_USUMAE WHERE USUMAE_CODIGO = '" & modgen_g_str_CodUsu & "' AND USUMAE_SITUAC = 1"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      gf_Seg_CamClave = False
      Exit Function
   End If
   
   g_rst_Princi.MoveFirst
      
   If gf_Seg_Desenc(Trim(g_rst_Princi!USUMAE_CONTRA)) <> p_TxtPwd_Act.Text Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
         
      MsgBox "La Contraseña Actual es incorrecta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(p_TxtPwd_Act)
      gf_Seg_CamClave = False
      Exit Function
   End If
     
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      gf_Seg_CamClave = False
      Exit Function
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_SEG_USUMAE_CAMBIOCLAVE ("
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & gf_Seg_Encrip(p_TxtPwd_Nue.Text) & "', "
         
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
            gf_Seg_CamClave = False
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   MsgBox "El cambio de contraseña se realizó con éxito.", vbInformation, modgen_g_str_NomPlt
   gf_Seg_CamClave = True
End Function

Function gf_Seg_FormatoPwd(p_Clave As String) As Boolean
Dim r_int_NumFil  As Integer
Dim r_int_TotCad  As Integer
Dim r_int_TotNum  As Integer

   gf_Seg_FormatoPwd = False
   
   If Len(Trim(p_Clave)) < 6 Then
      gf_Seg_FormatoPwd = False
      Exit Function
   End If
      
   r_int_NumFil = 0
   r_int_TotCad = 0
   r_int_TotNum = 0
   
   For r_int_NumFil = 1 To Len(Trim(p_Clave))
       If IsNumeric(Mid(Trim(p_Clave), r_int_NumFil, 1)) Then
          r_int_TotNum = r_int_TotNum + 1
       End If
       If InStr(modgen_g_con_LETRAS, Mid(UCase(Trim(p_Clave)), r_int_NumFil, 1)) Then
          r_int_TotCad = r_int_TotCad + 1
       End If
   Next
   
   If r_int_TotNum = 0 And r_int_TotCad > 0 Then
      gf_Seg_FormatoPwd = False
      Exit Function
   End If
   
   If r_int_TotNum > 0 And r_int_TotCad = 0 Then
      gf_Seg_FormatoPwd = False
      Exit Function
   End If
      
   gf_Seg_FormatoPwd = True
End Function

Public Function gf_Genera_Gastos_Cierre(ByVal p_CodPrd As String, ByVal p_CodPry As String, p_CodMod As String, p_ComVta As Double, p_ValEst As Double, p_MtoPre As Double, _
                                        ByRef p_GasTas As Double, ByRef p_GasNot As Double, ByRef p_BloReg As Double, ByRef p_RegMin As Double, ByRef p_RegHip As Double, ByRef p_ImpITF As Double) As Double
Dim r_dbl_GasTas        As Double
Dim r_dbl_GasNot        As Double
Dim r_dbl_BloReg        As Double
Dim r_dbl_GasMin        As Double
Dim r_dbl_GasMin_Inm    As Double
Dim r_dbl_GasMin_Est    As Double
Dim r_dbl_GasMin_Gar    As Double
Dim r_dbl_GasHip        As Double
Dim r_dbl_ValITF        As Double
Dim r_int_TipBie        As Integer

Dim r_dbl_TasInm        As Double
Dim r_dbl_FacInm        As Double
Dim r_dbl_FicInm        As Double
Dim r_dbl_TasEst        As Double
Dim r_dbl_FacEst        As Double
Dim r_dbl_FicEst        As Double
Dim r_dbl_TasGar        As Double
Dim r_dbl_FacGar        As Double
Dim r_dbl_FicGar        As Double

   gf_Genera_Gastos_Cierre = 0
         
   r_dbl_GasTas = 0
   r_dbl_GasNot = 0
   r_dbl_GasMin_Est = 0
   r_dbl_GasMin_Inm = 0
   r_dbl_GasMin = 0
   
   'Obtener Gastos de Tasación, Notariales y demás Parámetros
   Call gf_Obtener_Parametro(p_CodPrd, p_CodPry, r_dbl_GasTas, r_dbl_GasNot, r_dbl_TasInm, r_dbl_FacInm, r_dbl_FicInm, r_dbl_TasEst, r_dbl_FacEst, r_dbl_FicEst, r_dbl_TasGar, r_dbl_FacGar, r_dbl_FicGar)
   
   p_GasTas = Format(CDbl(r_dbl_GasTas), "###,###,##0.00")
   p_GasNot = Format(CDbl(r_dbl_GasNot), "###,###,##0.00")
   
   If r_dbl_GasNot = 0 Then
      MsgBox "No se asignaron valores para los Gastos Notariales.", vbExclamation, modgen_g_str_NomPlt
   End If
   
   
   'Determinar Tipo de Bien según Datos del Inmueble ingresados
   If p_CodMod = "001" Then
      r_int_TipBie = 2                                       'Bien Terminado
   Else
      r_int_TipBie = 1                                       'Bien Futuro
   End If
   
   'Bloqueo Registral
   If r_int_TipBie > 0 Then
      If r_int_TipBie = 2 Then
         If CDbl(p_ValEst) = 0 Then
            r_dbl_BloReg = r_dbl_FicInm
         Else
            r_dbl_BloReg = r_dbl_FicInm * 2
         End If
      Else
         r_dbl_BloReg = 0
      End If
   Else
      r_dbl_BloReg = 0
   End If
   
   p_BloReg = Format(CDbl(r_dbl_BloReg), "###,###,##0.00")
   
   'Gastos por registro de minuta
   If CDbl(p_ComVta) > 0 Then
      r_dbl_GasMin_Inm = (p_ComVta * r_dbl_FacInm) + r_dbl_FicInm
   End If
   
   If CDbl(p_ValEst) > 0 Then
      r_dbl_GasMin_Est = (p_ValEst * r_dbl_FacEst) + r_dbl_FicEst
   End If
   
   If CDbl(p_ComVta) > 0 And CDbl(p_ValEst) = 0 Then
      r_dbl_GasMin = r_dbl_GasMin_Inm + r_dbl_GasMin_Est + CDbl(r_dbl_FicInm)
   ElseIf CDbl(p_ComVta) = 0 And CDbl(p_ValEst) > 0 Then
      r_dbl_GasMin = r_dbl_GasMin_Inm + r_dbl_GasMin_Est + CDbl(r_dbl_FicEst)
   ElseIf CDbl(p_ComVta) > 0 And CDbl(p_ValEst) > 0 Then
      r_dbl_GasMin = r_dbl_GasMin_Inm + r_dbl_GasMin_Est + CDbl(r_dbl_FicInm) + CDbl(r_dbl_FicEst)
   End If
   
   p_RegMin = Format(CDbl(r_dbl_GasMin), "###,###,##0.00")
   
   'Gastos por registro de Hipoteca
   r_dbl_GasMin_Gar = CDbl(CDbl(CDbl(CDbl(p_ComVta) + CDbl(p_ValEst)) * r_dbl_FacGar) + r_dbl_FicGar) * 1.1
   
   r_dbl_GasHip = CDbl(r_dbl_GasMin_Gar * 1.15) + r_dbl_FicGar
   
   p_RegHip = Format(CDbl(r_dbl_GasHip), "###,###,##0.00")
   
   'Valor ITF
   r_dbl_ValITF = CDbl(p_MtoPre) * (0.005 / 100) * 2
   
   p_ImpITF = Format(CDbl(r_dbl_ValITF), "###,###,##0.00")
   
   'Totaliza los gastos de cierre
   gf_Genera_Gastos_Cierre = r_dbl_GasNot + r_dbl_BloReg + r_dbl_GasMin + r_dbl_GasHip + r_dbl_ValITF
   
   '  Adiciona el costo del POS 4.2%
   '  gf_Genera_Gastos_Cierre = gf_Genera_Gastos_Cierre * 1.042
End Function

Public Function gf_Obtener_Parametro(ByVal p_CodPrd As String, ByVal p_CodPry As String, ByRef p_ImpTas As Double, ByRef p_ImpNot As Double, ByRef p_TasInm As Double, ByRef p_FacInm As Double, ByRef p_FicInm As Double, _
                                     ByRef p_TasEst As Double, ByRef p_FacEst As Double, ByRef p_FicEst As Double, ByRef p_TasGar As Double, ByRef p_FacGar As Double, ByRef p_FicGar As Double)
Dim r_str_Princi     As String

   p_ImpTas = 0
   p_ImpNot = 0
   p_TasInm = 0
   p_FacInm = 0
   p_FicInm = 0
   p_TasEst = 0
   p_FacEst = 0
   p_FicEst = 0
   p_TasGar = 0
   p_FacGar = 0
   p_FicGar = 0
   
   r_str_Princi = ""
   r_str_Princi = r_str_Princi & " SELECT "
   r_str_Princi = r_str_Princi & "       ( SELECT GASPAR_GASTAS_MTO "
   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODPRT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 1 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_GASTAS_MTO > 0 ) AS MONTO_PERITO, "
   
   r_str_Princi = r_str_Princi & "       ( SELECT GASPAR_GASNOT_MTO "
   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 2 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_GASNOT_MTO > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_TASINM > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FACINM > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FICINM > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_TASEST > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FACEST > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FICEST > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS000 > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS001 > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS002 > 0 ) AS MONTO_NOTARIA "
   r_str_Princi = r_str_Princi & "   FROM DUAL "
   
   If Not gf_EjecutaSQL(r_str_Princi, g_rst_Genera, 3) Then
       Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      If Not IsNull(g_rst_Genera!MONTO_PERITO) Then
         p_ImpTas = CDbl(g_rst_Genera!MONTO_PERITO)
      End If
      If Not IsNull(g_rst_Genera!MONTO_NOTARIA) Then
         p_ImpNot = CDbl(g_rst_Genera!MONTO_NOTARIA)
      End If
   End If
   
   g_rst_Genera.Close
   
   r_str_Princi = ""
   r_str_Princi = r_str_Princi & "        SELECT GASPAR_REGMIN_TASINM AS TASINM, GASPAR_REGMIN_FACINM AS FACINM, GASPAR_REGMIN_FICINM AS FICINM, "
   r_str_Princi = r_str_Princi & "               GASPAR_REGMIN_TASEST AS TASEST, GASPAR_REGMIN_FACEST AS FACEST, GASPAR_REGMIN_FICEST AS FICEST, "
   r_str_Princi = r_str_Princi & "               GASPAR_REGGAR_TAS000 AS TASGAR, GASPAR_REGGAR_TAS001 AS FACGAR, GASPAR_REGGAR_TAS002 AS FICGAR "
   r_str_Princi = r_str_Princi & "          FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "               INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "         WHERE GASPAR_TIPTAB = 2 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_GASNOT_MTO > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGMIN_TASINM > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGMIN_FACINM > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGMIN_FICINM > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGMIN_TASEST > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGMIN_FACEST > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGMIN_FICEST > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGGAR_TAS000 > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGGAR_TAS001 > 0 "
   r_str_Princi = r_str_Princi & "           AND GASPAR_REGGAR_TAS002 > 0 "
   
   If Not gf_EjecutaSQL(r_str_Princi, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      p_TasInm = CDbl(g_rst_Genera!TASINM)
      p_FacInm = CDbl(g_rst_Genera!FACINM)
      p_FicInm = CDbl(g_rst_Genera!FICINM)
      
      p_TasEst = CDbl(g_rst_Genera!TASEST)
      p_FacEst = CDbl(g_rst_Genera!FACEST)
      p_FicEst = CDbl(g_rst_Genera!FICEST)
      
      p_TasGar = CDbl(g_rst_Genera!TASGAR)
      p_FacGar = CDbl(g_rst_Genera!FACGAR)
      p_FicGar = CDbl(g_rst_Genera!FICGAR)
   End If
End Function

Public Function gf_Valida_GastoCierre(ByVal p_CodPrd As String, ByVal p_CodPry As String) As Integer
Dim r_str_Princi  As String
Dim r_rst_Genera  As ADODB.Recordset

   gf_Valida_GastoCierre = 0
   r_str_Princi = ""
   
   r_str_Princi = r_str_Princi & " SELECT "
   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODPRT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 1 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 ) AS PERITO, "
   
   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODPRT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 1 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_GASTAS_MTO > 0 ) AS MONTO_PERITO, "
   
   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 2 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 ) AS NOTARIA, "
   
   r_str_Princi = r_str_Princi & "       ( SELECT COUNT(*) "
   r_str_Princi = r_str_Princi & "           FROM TRA_GASPAR A "
   r_str_Princi = r_str_Princi & "                INNER JOIN PRY_DATGEN ON TRIM(TO_CHAR(TRIM(DATGEN_CODNOT), '000000')) = TRIM(GASPAR_CODEMP) AND GASPAR_CODPRY = DATGEN_CODIGO "
   r_str_Princi = r_str_Princi & "          WHERE GASPAR_TIPTAB = 2 AND GASPAR_CODPRD = '" & p_CodPrd & "' AND GASPAR_CODPRY = '" & p_CodPry & "' AND GASPAR_SITUAC = 1 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_GASNOT_MTO > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_TASINM > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FACINM > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FICINM > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_TASEST > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FACEST > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGMIN_FICEST > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS000 > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS001 > 0 "
   r_str_Princi = r_str_Princi & "            AND GASPAR_REGGAR_TAS002 > 0 ) AS TASAS_NOTARIA "
      
   r_str_Princi = r_str_Princi & "   FROM DUAL "
   
   If Not gf_EjecutaSQL(r_str_Princi, r_rst_Genera, 3) Then
      Exit Function
   End If
    
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      If CInt(r_rst_Genera!PERITO) = 0 Then
         gf_Valida_GastoCierre = 1
      ElseIf CInt(r_rst_Genera!PERITO) > 0 And CInt(r_rst_Genera!MONTO_PERITO) = 0 Then
         gf_Valida_GastoCierre = 4
      ElseIf CInt(r_rst_Genera!NOTARIA) = 0 Then
         gf_Valida_GastoCierre = 2
      ElseIf CInt(r_rst_Genera!NOTARIA) > 0 And CInt(r_rst_Genera!TASAS_NOTARIA) = 0 Then
         gf_Valida_GastoCierre = 3
      End If
   End If
    
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Function gf_Obtener_Proyec(p_NumSol As String) As String
Dim r_str_Parame     As String

   gf_Obtener_Proyec = ""

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT SOLINM_PRYCOD "
   r_str_Parame = r_str_Parame & "   FROM CRE_SOLMAE  "
   r_str_Parame = r_str_Parame & "        INNER JOIN CRE_SOLINM ON SOLINM_NUMSOL = SOLMAE_NUMERO"
   r_str_Parame = r_str_Parame & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   If Not g_rst_Genera.EOF Then
      If Not IsNull(g_rst_Genera!SOLINM_PRYCOD) Then
         gf_Obtener_Proyec = g_rst_Genera!SOLINM_PRYCOD
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

