VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_43 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8565
   Icon            =   "AteCli_frm_569.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3015
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8565
      _Version        =   65536
      _ExtentX        =   15108
      _ExtentY        =   5318
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
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
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
            TabIndex        =   8
            Top             =   60
            Width           =   7725
            _Version        =   65536
            _ExtentX        =   13626
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Solicitudes en Trámite con Aprobación Crediticia y Pago de Gastos de Cierre Cancelados"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin Threed.SSPanel ssp_TipCon 
            Height          =   285
            Left            =   660
            TabIndex        =   9
            Top             =   330
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   503
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
            Left            =   90
            Picture         =   "AteCli_frm_569.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   10
         Top             =   780
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7830
            Picture         =   "AteCli_frm_569.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_569.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_569.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1320
            Top             =   120
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
         Height          =   765
         Left            =   60
         TabIndex        =   11
         Top             =   2190
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
         _ExtentY        =   1349
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
         Begin VB.CheckBox chk_TipCon 
            Caption         =   "Todos"
            Height          =   315
            Left            =   1470
            TabIndex        =   2
            Top             =   420
            Width           =   4425
         End
         Begin VB.ComboBox cmb_TipCon 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   6495
         End
         Begin VB.Label lbl_TipCon 
            Caption         =   "Tipo de Consulta:"
            Height          =   465
            Left            =   60
            TabIndex        =   12
            Top             =   150
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   675
         Left            =   60
         TabIndex        =   13
         Top             =   1470
         Width           =   8445
         _Version        =   65536
         _ExtentX        =   14896
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   6495
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Reporte:"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   210
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_ConHip()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_TipCon_Click()
   If chk_TipCon.Value = 1 Then
      cmb_TipCon.ListIndex = -1
      cmb_TipCon.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_TipCon.Value = 0 Then
      cmb_TipCon.Enabled = True
      Call gs_SetFocus(cmb_TipCon)
   End If
End Sub

Private Sub cmb_TipCon_Click()
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_TipCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCon_Click
   End If
End Sub

Private Sub cmb_TipRep_Click()
   Call fs_Limpia
   If cmb_TipRep.ListIndex <> -1 Then
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
         Call moddat_gs_Carga_Produc(cmb_TipCon, l_arr_Produc, 4)
         chk_TipCon.Caption = "Todos los Productos"
         lbl_TipCon.Caption = "Producto:"
         ssp_TipCon.Caption = "Por Producto"
      Else
         Call moddat_gs_Carga_EjecMC(cmb_TipCon, l_arr_ConHip, 121)
         chk_TipCon.Caption = "Todos los Consejeros Hipotecarios"
         lbl_TipCon.Caption = "Consejero Hipotecario:"
         ssp_TipCon.Caption = "Por Consejero Hipotecario"
      End If
   End If
   Call gs_SetFocus(cmb_TipCon)
End Sub

Private Sub fs_Limpia()
   cmb_TipCon.Clear
   chk_TipCon.Value = 0
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRep_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If chk_TipCon.Value = 0 Then
      If cmb_TipCon.ListIndex = -1 Then
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_TipCon)
         Exit Sub
      End If
   End If
   
     'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenExc_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenExc_ConHip
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   'Validación
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If chk_TipCon = 0 Then
      If cmb_TipCon.ListIndex = -1 Then
         If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
            MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
            MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_TipCon)
         Exit Sub
      End If
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
      Call fs_GenImp_TipPro
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      Call fs_GenImp_ConHip
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   cmb_TipRep.AddItem "POR PRODUCTO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   If modgen_g_int_TipUsu <> 20121 Then
      cmb_TipRep.AddItem "POR CONSEJERO HIPOTECARIO"
      cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   End If
   cmb_TipRep.ListIndex = -1
   
   Call fs_Limpia
   Call gs_CentraForm(Me)
End Sub

Private Function ff_GasAdm(ByVal p_NumSol As String, Optional ByRef p_FecPag As Double) As Double
   ff_GasAdm = 0
   p_FecPag = 0
   
   g_str_Parame = "SELECT GASADM_PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_GasAdm = ff_GasAdm + g_rst_Listas!GASADM_PAGIMP
         p_FecPag = g_rst_Listas!GASADM_PAGFEC
         
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_GenImp_ConHip()
Dim r_dbl_GasAdm     As Double
Dim r_dbl_GasFec     As Double

   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_04.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, NVL(B.PAGIMP,0) AS PAGIMP, B.GASADM_PAGFEC"
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        LEFT JOIN (SELECT GASADM_NUMSOL, SUM(GASADM_PAGIMP) AS PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "                          GASADM_SITUAC = 1  GROUP BY GASADM_NUMSOL, GASADM_PAGFEC) B ON B.GASADM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Consejeros Hipotecarios
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS > 21 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
         'Para obtener Total de Gastos de Cierre (Pagados)
         'r_dbl_GasAdm = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO, r_dbl_GasFec)
         r_dbl_GasAdm = g_rst_Princi!PAGIMP
         r_dbl_GasFec = IIf(IsNull(g_rst_Princi!GASADM_PAGFEC), 0, g_rst_Princi!GASADM_PAGFEC)
                          
         If r_dbl_GasFec <> 0 Then
            'Insertando Registro
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
            g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
            g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
            g_str_Parame = g_str_Parame & "SOLTRA_TOTGAS, "
            g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC) "
            
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_04.RPT" & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
            g_str_Parame = g_str_Parame & CStr(r_dbl_GasAdm) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_GasFec) & ") "
                     
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
 
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_04.RPT"
        
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_04.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' "
        
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenImp_TipPro()
  Dim r_dbl_GasAdm     As Double
  Dim r_dbl_GasFec     As Double

   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA WHERE "
   g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT = 'ATE_RPTSOL_03.RPT' AND "
   g_str_Parame = g_str_Parame & "SOLTRA_TERCRE = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   'g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, NVL(B.PAGIMP,0) AS PAGIMP, B.GASADM_PAGFEC"
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "        LEFT JOIN (SELECT GASADM_NUMSOL, SUM(GASADM_PAGIMP) AS PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "                          GASADM_SITUAC = 1  GROUP BY GASADM_NUMSOL, GASADM_PAGFEC) B ON B.GASADM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  WHERE "
   
   'Si no escogio todos los Productos
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   'Restricción por Tipo de Usuario
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS > 21 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
         'Para obtener Total de Gastos de Cierre (Pagados)
         'r_dbl_GasAdm = ff_GasAdm(g_rst_Princi!SOLMAE_NUMERO, r_dbl_GasFec)
         r_dbl_GasAdm = g_rst_Princi!PAGIMP
         r_dbl_GasFec = IIf(IsNull(g_rst_Princi!GASADM_PAGFEC), 0, g_rst_Princi!GASADM_PAGFEC)
           
         If r_dbl_GasFec <> 0 And r_dbl_GasAdm <> 0 Then
            'Insertando Registro
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RPT_SOLTRA("
            g_str_Parame = g_str_Parame & "SOLTRA_NOMRPT, "
            g_str_Parame = g_str_Parame & "SOLTRA_FECCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_HORCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_TERCRE, "
            g_str_Parame = g_str_Parame & "SOLTRA_NUMSOL, "
            g_str_Parame = g_str_Parame & "SOLTRA_TOTGAS, "
            g_str_Parame = g_str_Parame & "SOLTRA_PAGFEC) "
            
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & "ATE_RPTSOL_03.RPT" & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
            g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SOLMAE_NUMERO & "', "
            g_str_Parame = g_str_Parame & CStr(r_dbl_GasAdm) & ", "
            g_str_Parame = g_str_Parame & CStr(r_dbl_GasFec) & ") "
                     
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
         End If
      
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
        
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "RPT_SOLTRA"
          
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ATE_RPTSOL_03.RPT"
        
   crp_Imprim.SelectionFormula = "{RPT_SOLTRA.SOLTRA_NOMRPT} = 'ATE_RPTSOL_03.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SOLTRA.SOLTRA_TERCRE} = '" & modgen_g_str_NombPC & "' "
        
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc_ConHip()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_dbl_GasAdm     As Double
Dim r_dbl_GasFec     As Double
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_NUMERO, A.SOLMAE_CONHIP, C.PRODUC_DESCRI, A.SOLMAE_TITTDO, A.SOLMAE_TITNDO, A.SOLMAE_TIPMON, A.SOLMAE_FECSOL, "
   g_str_Parame = g_str_Parame & "        TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS CLIENTE, D.PARDES_DESCRI AS INSTANCIA, "
   g_str_Parame = g_str_Parame & "        A.SOLMAE_COMVTA_SOL, A.SOLMAE_COMVTA_DOL, A.SOLMAE_APOPRO_SOL, A.SOLMAE_APOPRO_DOL, A.SOLMAE_MTOPRE_MPR, NVL(E.PAGIMP,0) AS PAGIMP, E.GASADM_PAGFEC "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A "
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON SOLMAE_TITTDO = DATGEN_TIPDOC AND SOLMAE_TITNDO = DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_CODINS AND D.PARDES_CODGRP = '002' "
   g_str_Parame = g_str_Parame & "   LEFT JOIN (SELECT GASADM_NUMSOL, SUM(GASADM_PAGIMP) AS PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE GASADM_SITUAC = 1 AND GASADM_CODGAS = 11 GROUP BY GASADM_NUMSOL, GASADM_PAGFEC) E ON E.GASADM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  WHERE "
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "        SOLMAE_CONHIP = '" & l_arr_ConHip(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   g_str_Parame = g_str_Parame & "        SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "        SOLMAE_CODINS > 21 "
   g_str_Parame = g_str_Parame & "  ORDER BY SOLMAE_CONHIP ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC "
      
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
      .Cells(1, 8) = "INSTANCIA ACTUAL"
      .Cells(1, 9) = "FECHA DE PAGO"
      .Cells(1, 10) = "TIP. MONEDA"
      .Cells(1, 11) = "V. INMUEBLE S/."
      .Cells(1, 12) = "V. INMUEBLE US$."
      .Cells(1, 13) = "PORC. INICIAL"
      .Cells(1, 14) = "MTO. CREDITO S/."
      .Cells(1, 15) = "MTO. CREDITO US$."
      .Cells(1, 16) = "PAGO G.C. S/."
      .Cells(1, 17) = "PAGO G.C. US$."
      
      .Range(.Cells(1, 1), .Cells(1, 17)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 17)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 32
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 40
      .Columns("G").ColumnWidth = 13
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 30
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 22
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 21
      .Columns("K").ColumnWidth = 14
      .Columns("L").ColumnWidth = 18
      .Columns("M").ColumnWidth = 14
      .Columns("N").ColumnWidth = 16
      .Columns("O").ColumnWidth = 18
      .Columns("P").ColumnWidth = 14
      .Columns("Q").ColumnWidth = 15
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not g_rst_Princi.EOF
      r_dbl_GasAdm = g_rst_Princi!PAGIMP
      r_dbl_GasFec = IIf(IsNull(g_rst_Princi!GASADM_PAGFEC), 0, g_rst_Princi!GASADM_PAGFEC)
      
      If r_dbl_GasFec > 0 And r_dbl_GasAdm > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!SOLMAE_CONHIP)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!PRODUC_DESCRI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(r_dbl_GasFec)))
              
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "SOLES"
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
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_dbl_GasAdm, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = 0
         End If
         If g_rst_Princi!SOLMAE_TIPMON = 2 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_dbl_GasAdm, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = 0
         End If
                                 
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_TipPro()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_dbl_GasAdm     As Double
   Dim r_dbl_GasFec     As Double
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_SOLTRA "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SOLMAE_CODPRD, PRODUC_DESCRI, SOLMAE_NUMERO, SOLMAE_TITTDO, SOLMAE_TITNDO, SOLMAE_TIPMON, SOLMAE_CONHIP, SOLMAE_FECSOL, "
   g_str_Parame = g_str_Parame & "        TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, D.PARDES_DESCRI AS INSTANCIA, E.GASADM_PAGFEC, "
   g_str_Parame = g_str_Parame & "        SOLMAE_COMVTA_SOL, SOLMAE_COMVTA_DOL, SOLMAE_APOPRO_SOL, SOLMAE_APOPRO_DOL, SOLMAE_MTOPRE_MPR , NVL(E.PAGIMP,0) AS PAGIMP "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE A"
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN B ON A.SOLMAE_TITTDO = B.DATGEN_TIPDOC AND A.SOLMAE_TITNDO = B.DATGEN_NUMDOC "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON D.PARDES_CODITE = A.SOLMAE_CODINS AND D.PARDES_CODGRP = '002' "
   g_str_Parame = g_str_Parame & "   LEFT JOIN (SELECT GASADM_NUMSOL, SUM(GASADM_PAGIMP) AS PAGIMP, GASADM_PAGFEC FROM TRA_GASADM WHERE GASADM_SITUAC = 1 AND GASADM_CODGAS = 11 GROUP BY GASADM_NUMSOL, GASADM_PAGFEC) E ON E.GASADM_NUMSOL = A.SOLMAE_NUMERO "
   g_str_Parame = g_str_Parame & "  WHERE "
   If chk_TipCon.Value = 0 Then
      g_str_Parame = g_str_Parame & "        SOLMAE_CODPRD = '" & l_arr_Produc(cmb_TipCon.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   If modgen_g_int_TipUsu = 20121 Then
      g_str_Parame = g_str_Parame & "        SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   g_str_Parame = g_str_Parame & "        SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "        SOLMAE_CODINS > 21 "
   g_str_Parame = g_str_Parame & "  ORDER BY SOLMAE_CODPRD ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   'Si no se encuentra ninguna solicitud
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "SOLICITUD"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "F. SOLICITUD"
      .Cells(1, 7) = "INSTANCIA ACTUAL"
      .Cells(1, 8) = "FECHA DE PAGO"
      .Cells(1, 9) = "CONSEJ. HIPOT."
      .Cells(1, 10) = "TIPO DE MONEDA"
      .Cells(1, 11) = "V. INMUEBLE"
      .Cells(1, 12) = "PORC. INICIAL"
      .Cells(1, 13) = "MTO. CREDITO"
      .Cells(1, 14) = "MONTO DE PAGO"
   
      .Range(.Cells(1, 1), .Cells(1, 14)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 14)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 32
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 12
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 30
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 17
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 21
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 15
      .Columns("L").ColumnWidth = 15
      .Columns("M").ColumnWidth = 13
      .Columns("N").ColumnWidth = 17
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_dbl_GasAdm = g_rst_Princi!PAGIMP
      r_dbl_GasFec = IIf(IsNull(g_rst_Princi!GASADM_PAGFEC), 0, g_rst_Princi!GASADM_PAGFEC)
       
      If r_dbl_GasFec > 0 And r_dbl_GasAdm > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(r_dbl_GasFec)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!SOLMAE_CONHIP)
         
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "SOLES"
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "DOLARES AMERICANOS"
         End If
         If g_rst_Princi!SOLMAE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00")
         End If
         If g_rst_Princi!SOLMAE_COMVTA_SOL > 0 Or g_rst_Princi!SOLMAE_COMVTA_DOL > 0 Then
            If g_rst_Princi!SOLMAE_TIPMON = 1 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_SOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_SOL, "###,###,##0.00") * 100) + "%"
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(Format(g_rst_Princi!SOLMAE_APOPRO_DOL, "###,###,##0.00") / Format(g_rst_Princi!SOLMAE_COMVTA_DOL, "###,###,##0.00") * 100) + "%"
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CStr(0) + "%"
         End If
               
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!SOLMAE_MTOPRE_MPR, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(r_dbl_GasAdm, "###,###,##0.00")
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
