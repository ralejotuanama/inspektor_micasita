VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   1455
   ClientTop       =   2385
   ClientWidth     =   7665
   Icon            =   "AteCli_frm_519.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3555
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7695
      _Version        =   65536
      _ExtentX        =   13573
      _ExtentY        =   6271
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7605
         _Version        =   65536
         _ExtentX        =   13414
         _ExtentY        =   1191
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   660
            TabIndex        =   14
            Top             =   30
            Width           =   4965
            _Version        =   65536
            _ExtentX        =   8758
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes Desembolsadas x Proyecto"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   660
            TabIndex        =   15
            Top             =   330
            Width           =   2715
            _Version        =   65536
            _ExtentX        =   4789
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Por Consejero Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_519.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   750
         Width           =   7605
         _Version        =   65536
         _ExtentX        =   13414
         _ExtentY        =   1138
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_519.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6990
            Picture         =   "AteCli_frm_519.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_519.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1440
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2055
         Left            =   30
         TabIndex        =   12
         Top             =   1440
         Width           =   7605
         _Version        =   65536
         _ExtentX        =   13414
         _ExtentY        =   3625
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.ComboBox cmb_Proyec 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   5775
         End
         Begin VB.ComboBox cmb_TipPry 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5775
         End
         Begin VB.CheckBox chk_Proyec 
            Caption         =   "Todos los Proyectos"
            Height          =   285
            Left            =   1770
            TabIndex        =   3
            Top             =   1050
            Width           =   1845
         End
         Begin VB.CheckBox chk_TipPry 
            Caption         =   "Todos los Tipos de Proyectos"
            Height          =   285
            Left            =   1770
            TabIndex        =   1
            Top             =   390
            Width           =   2475
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1380
            Width           =   5775
         End
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejeros"
            Height          =   315
            Left            =   1770
            TabIndex        =   5
            Top             =   1710
            Width           =   2685
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto:"
            Height          =   255
            Left            =   90
            TabIndex        =   17
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Proyecto:"
            Height          =   255
            Left            =   90
            TabIndex        =   16
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Consejero Hipotecario:"
            Height          =   255
            Left            =   90
            TabIndex        =   13
            Top             =   1410
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ConHip()   As moddat_tpo_Genera
Dim l_arr_Proyec()   As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_ConHip_Click()
   If chk_ConHip.Value = 1 Then
      cmb_ConHip.ListIndex = -1
      cmb_ConHip.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_ConHip.Value = 0 Then
      cmb_ConHip.Enabled = True
      Call gs_SetFocus(cmb_ConHip)
   End If
End Sub

Private Sub chk_Proyec_Click()
   If chk_Proyec.Value = 1 Then
      cmb_Proyec.ListIndex = -1
      cmb_Proyec.Enabled = False
      Call gs_SetFocus(cmb_ConHip)
   ElseIf chk_Proyec.Value = 0 Then
      cmb_Proyec.Enabled = True
      Call gs_SetFocus(cmb_Proyec)
   End If
End Sub

Private Sub chk_TipPry_Click()
   If chk_TipPry.Value = 1 Then
      cmb_TipPry.ListIndex = -1
      cmb_TipPry.Enabled = False
      cmb_Proyec.Enabled = False
      chk_Proyec.Value = 1
      chk_Proyec.Enabled = False
      Call gs_SetFocus(cmb_ConHip)
   ElseIf chk_TipPry.Value = 0 Then
      chk_Proyec.Enabled = True
      cmb_TipPry.Enabled = True
      cmb_Proyec.Enabled = True
      chk_Proyec.Value = 0
      
      Call gs_SetFocus(cmb_TipPry)
   End If
End Sub

'Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call cmb_ConHip_Click
'   End If
'End Sub

'Private Sub cmb_Proyec_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call cmb_Proyec_Click
'   End If
'End Sub

Private Sub cmb_TipPry_Click()
   
   If cmb_TipPry.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      Call Carga_PryInm_Combo(cmb_Proyec, l_arr_Proyec, cmb_TipPry.ItemData(cmb_TipPry.ListIndex))
      
      Screen.MousePointer = 0
   End If
   
End Sub

Private Sub cmb_TipPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPry_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
    
   'Validación
   If chk_ConHip = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   If chk_Proyec = 0 Then
      If cmb_Proyec.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Proyec)
         Exit Sub
      End If
   End If
   
   If chk_TipPry = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_GenExc
   
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_PryMcs     As String
   
   'Validación
   If chk_ConHip = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar a un Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   If chk_Proyec = 0 Then
      If cmb_Proyec.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Proyec)
         Exit Sub
      End If
   End If
   
   If chk_TipPry = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el Reporte de Solicitudes Desembolsadas por Proyecto?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(Date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_DESPRY "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_SOLMAE, PRY_DATGEN, CRE_SOLINM, CRE_HIPMAE WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = HIPMAE_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "DATGEN_CODIGO = SOLINM_PRYCOD AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
            
         'Para obtener SI es un proyecto vinculado (Mi Casita)
         r_str_PryMcs = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_PRYMCS))
         'r_str_PryMcs = cmb_Proyec
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_DESPRY("
         g_str_Parame = g_str_Parame & "DESPRY_NOMRPT, "
         g_str_Parame = g_str_Parame & "DESPRY_FECCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_HORCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_TERCRE, "
         g_str_Parame = g_str_Parame & "DESPRY_NUMSOL, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYTIT, "
         g_str_Parame = g_str_Parame & "DESPRY_PRYMCS, "
         g_str_Parame = g_str_Parame & "DESPRY_FECDES) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_08.RPT" & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_TITULO & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_PryMcs & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_FECDES & "') "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
  
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = UCase(moddat_g_str_EntDat) & ".CRE_PRODUC"
   crp_Imprim.DataFiles(2) = UCase(moddat_g_str_EntDat) & ".CLI_DATGEN"
   crp_Imprim.DataFiles(3) = UCase(moddat_g_str_EntDat) & ".RPT_DESPRY"
   
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_08.RPT"
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
  ' Call gs_SetFocus(cmb_ConHip)
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   
   cmb_TipPry.AddItem ("VINCULADO")
   cmb_TipPry.ItemData(cmb_TipPry.NewIndex) = 1
      
   cmb_TipPry.AddItem ("NO VINCULADO")
   cmb_TipPry.ItemData(cmb_TipPry.NewIndex) = 2

End Sub

Private Sub Carga_PryInm_Combo(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipPry As Integer)
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

Private Sub fs_GenExc()
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CLI_DATGEN B, PRY_DATGEN C, CRE_HIPMAE D, CRE_SOLINM E WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = HIPMAE_NUMSOL AND "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = SOLINM_NUMSOL AND "
   g_str_Parame = g_str_Parame & "DATGEN_CODIGO = SOLINM_PRYCOD AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = DATGEN_NUMDOC AND "
   
   If chk_TipPry.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = '" & (cmb_TipPry.ListIndex + 1) & "' AND "
   End If
   
   If chk_Proyec.Value = 0 Then
      g_str_Parame = g_str_Parame & "DATGEN_CODIGO = '" & l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo & "' AND "
   End If
      
   If chk_ConHip.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "' AND "
   End If
    
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_CONHIP ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
          
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   'Si no encuentra ninguna Solicitud
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "CONSEJ. HIPOT."
      .Cells(1, 3) = "PRODUCTO"
      .Cells(1, 4) = "SOLICITUD"
      .Cells(1, 5) = "DOC. IDENTIDAD"
      .Cells(1, 6) = "NOMBRE CLIENTE"
      .Cells(1, 7) = "F. SOLICITUD"
      .Cells(1, 8) = "F. DESEMBOLSO"
      .Cells(1, 9) = "TIP. EVALUACION"
      .Cells(1, 10) = "TIP. MONEDA"
      .Cells(1, 11) = "V. INMUEBLE S/."
      .Cells(1, 12) = "V. INMUEBLE US$."
      .Cells(1, 13) = "PORC. INICIAL"
      .Cells(1, 14) = "MTO. CREDITO S/."
      .Cells(1, 15) = "MTO. CREDITO US$."
      
      .Range(.Cells(1, 1), .Cells(1, 15)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      
      .Columns("B").ColumnWidth = 32
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 32
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 43
            
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 16
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 22
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 22
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 17
      .Columns("L").ColumnWidth = 16
      .Columns("M").ColumnWidth = 17
      .Columns("N").ColumnWidth = 18
      .Columns("O").ColumnWidth = 18
                 
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
               
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
      
         
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "NUEVOS SOLES"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "DOLARES AMERICANOS"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = 0
      End If
              
      If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
         End If
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(0) + "%"
      End If
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = 0
      End If
         
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
      Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = 0
      End If
         
      r_int_ConVer = r_int_ConVer + 1
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub
