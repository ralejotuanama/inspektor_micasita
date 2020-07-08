VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Seg_SolHip_51 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   4110
   ClientTop       =   1950
   ClientWidth     =   15045
   Icon            =   "OpeTra_frm_180.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7965
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15090
      _Version        =   65536
      _ExtentX        =   26617
      _ExtentY        =   14049
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_180.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SegSol 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_180.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Detalle de Seguimiento de Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_180.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Solicitudes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_180.frx":0EEA
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_180.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos de Búsqueda de Solicitudes"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14370
            Picture         =   "OpeTra_frm_180.frx":14FE
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   13515
         End
         Begin VB.CheckBox chk_Produc 
            Caption         =   "Todos los Productos"
            Height          =   315
            Left            =   1410
            TabIndex        =   1
            Top             =   390
            Width           =   2685
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            Height          =   495
            Left            =   630
            TabIndex        =   13
            Top             =   60
            Width           =   8835
            _Version        =   65536
            _ExtentX        =   15584
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitudes de Créditos Hipotecarios"
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
            Left            =   60
            Picture         =   "OpeTra_frm_180.frx":1940
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5655
         Left            =   30
         TabIndex        =   14
         Top             =   2250
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
         _ExtentY        =   9975
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5235
            Left            =   60
            TabIndex        =   2
            Top             =   360
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   9234
            _Version        =   393216
            Rows            =   45
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_FecIns 
            Height          =   285
            Left            =   8265
            TabIndex        =   15
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Instancia"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   2340
            TabIndex        =   16
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   3720
            TabIndex        =   17
            Top             =   60
            Width           =   3495
            _Version        =   65536
            _ExtentX        =   6165
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_InsAct 
            Height          =   285
            Left            =   9300
            TabIndex        =   18
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia Actual"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   7200
            TabIndex        =   19
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   2250
            _Version        =   65536
            _ExtentX        =   3969
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_ConHip 
            Height          =   285
            Left            =   13200
            TabIndex        =   21
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cons. Hipotecario"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_SitIns 
            Height          =   285
            Left            =   11520
            TabIndex        =   22
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación Instancia"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frm_Seg_SolHip_51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmd_Buscar_Click()
   If chk_Produc.Value = 0 Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar_Creditos
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BusCli_Click()
   frm_Seg_SolHip_52.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Produc.ListIndex = -1
   chk_Produc.Value = 0
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub cmd_SegSol_Click()
   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Seg_SolHip_53.Show 1
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion - XXX
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(True)
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Produc)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc_Comerc(cmb_Produc, l_arr_Produc, 4)
   grd_Listad.Cols = 26
   grd_Listad.ColWidth(0) = 2250
   grd_Listad.ColWidth(1) = 1400
   grd_Listad.ColWidth(2) = 3460
   grd_Listad.ColWidth(3) = 1050
   grd_Listad.ColWidth(4) = 1050
   grd_Listad.ColWidth(5) = 2215
   grd_Listad.ColWidth(6) = 1685
   grd_Listad.ColWidth(7) = 1400
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColWidth(15) = 0
   grd_Listad.ColWidth(16) = 0
   grd_Listad.ColWidth(17) = 0
   grd_Listad.ColWidth(18) = 0
   grd_Listad.ColWidth(19) = 0
   grd_Listad.ColWidth(20) = 0
   grd_Listad.ColWidth(21) = 0
   grd_Listad.ColWidth(22) = 0
   grd_Listad.ColWidth(23) = 0
   grd_Listad.ColWidth(24) = 0
   grd_Listad.ColWidth(25) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
   chk_Produc.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_SegSol.Enabled = Not p_Activa
End Sub

''********************************************
''********************************************
'''''MODIFICACION
''********************************************
''********************************************
Private Sub fs_Buscar_Creditos()
Dim r_int_FlgIn1     As Integer
Dim r_int_FlgIn2     As Integer
   
   g_str_Parame = "  "
   g_str_Parame = g_str_Parame & " SELECT TRIM(PRODUC_DESCRI) AS PRODUCTO, SOL.SOLMAE_NUMERO AS SOLICITUD "
   g_str_Parame = g_str_Parame & "        ,TRIM(SUBSTRC(SOLMAE_NUMERO, 1,3)||'-'||SUBSTRC(SOLMAE_NUMERO, 4,3)||'-'||SUBSTRC(SOLMAE_NUMERO, 7,2)||'-'||SUBSTRC(SOLMAE_NUMERO, 9,4) ) AS NROSOL"
   g_str_Parame = g_str_Parame & "        ,TRIM(SOLMAE_TITTDO||'-'||TRIM(SOLMAE_TITNDO)) AS DNI"
   g_str_Parame = g_str_Parame & "        ,CASE WHEN DATGEN_APECAS IS NULL THEN TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "              WHEN LENGTH(TRIM(DATGEN_APECAS)) = 0 THEN TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "              ELSE TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' DE '||TRIM(DATGEN_APECAS)||' '||TRIM(DATGEN_NOMBRE)"
   g_str_Parame = g_str_Parame & "         END  AS CLIENTE "
   g_str_Parame = g_str_Parame & "        ,TO_DATE(SOLMAE_FECSOL,'YYYY/MM/DD') AS FECHASOL"
   g_str_Parame = g_str_Parame & "        ,TRIM(MO.PARDES_DESCRI) AS MONEDA"
   g_str_Parame = g_str_Parame & "        ,SOLMAE_MTOPRE_MPR   AS MONTO "
   g_str_Parame = g_str_Parame & "        ,NVL(S.PARPRD_DESCRI, ' ')  AS MODALIDAD_PRESTAMO "
   g_str_Parame = g_str_Parame & "        ,TRIM(IT.PARDES_DESCRI) AS INSTANCIA"
   g_str_Parame = g_str_Parame & "        ,TO_DATE(SEGUIM_FECINI,'YYYY/MM/DD') AS FECHAINS"
   g_str_Parame = g_str_Parame & "        ,CASE WHEN SOLMAE_CODINS = 41 THEN TRIM("
   g_str_Parame = g_str_Parame & "                   DECODE((SELECT SEGUIM_SITUAC FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = SOLMAE_NUMERO AND SEGUIM_CODINS = 42) ,3"
   g_str_Parame = g_str_Parame & "                   ,(SELECT PARDES_DESCRI FROM TRA_SEGUIM "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES TI ON (PARDES_CODGRP=023 AND PARDES_CODITE=SEGUIM_SITUAC)"
   g_str_Parame = g_str_Parame & "                     WHERE SEGUIM_NUMSOL = SOLMAE_NUMERO AND SEGUIM_CODINS = 42) "
   g_str_Parame = g_str_Parame & "                    ,TRIM(TI.PARDES_DESCRI)) )"
   g_str_Parame = g_str_Parame & "              WHEN SOLMAE_CODINS = 61 THEN  TRIM("
   g_str_Parame = g_str_Parame & "                   DECODE((SELECT SEGUIM_SITUAC FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL =solmae_numero  AND SEGUIM_CODINS = 62 ) ,3"
   g_str_Parame = g_str_Parame & "                   ,(SELECT PARDES_DESCRI FROM TRA_SEGUIM "
   g_str_Parame = g_str_Parame & "                       LEFT JOIN MNT_PARDES TI ON (PARDES_CODGRP=023 AND PARDES_CODITE=SEGUIM_SITUAC)"
   g_str_Parame = g_str_Parame & "                      WHERE SEGUIM_NUMSOL = solmae_numero AND SEGUIM_CODINS = 62) "
   g_str_Parame = g_str_Parame & "                     ,TRIM(TI.PARDES_DESCRI))  )"
   g_str_Parame = g_str_Parame & "         ELSE TRIM(TI.PARDES_DESCRI)  END  AS SITUAINSTANCIA"
   g_str_Parame = g_str_Parame & "        ,CASE WHEN SOLMAE_CODINS = 41 THEN   TRIM("
   g_str_Parame = g_str_Parame & "                   (SELECT PARDES_DESCRI FROM TRA_SEGUIM "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES TI ON (PARDES_CODGRP=023 AND PARDES_CODITE=SEGUIM_SITUAC)"
   g_str_Parame = g_str_Parame & "                     WHERE SEGUIM_NUMSOL = SOLMAE_NUMERO AND SEGUIM_CODINS = 42) )"
   g_str_Parame = g_str_Parame & "              WHEN SOLMAE_CODINS = 61 THEN  TRIM("
   g_str_Parame = g_str_Parame & "                   (SELECT PARDES_DESCRI FROM TRA_SEGUIM "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES TI ON (PARDES_CODGRP=023 AND PARDES_CODITE=SEGUIM_SITUAC)"
   g_str_Parame = g_str_Parame & "                     WHERE SEGUIM_NUMSOL = SOLMAE_NUMERO AND SEGUIM_CODINS = 62) )"
   g_str_Parame = g_str_Parame & "         ELSE TRIM(TI.PARDES_DESCRI)  END PRUEBA"
   g_str_Parame = g_str_Parame & "        ,SOLMAE_CONHIP AS CONSEJERO "
   g_str_Parame = g_str_Parame & "        ,NVL(TRIM(DR.PARDES_DESCRI),'-') AS PRYMICASITA"
   g_str_Parame = g_str_Parame & "        ,TRIM(NVL(DECODE(SOLINM_PRYCOD, 1,SOLINM_PRYNOM, DECODE(SOLINM_PRYCOD, NULL,SOLINM_PRYNOM, DATGEN_TITULO)),'-') ) AS PROYECTO"
   g_str_Parame = g_str_Parame & "        ,TRIM(SOLINM_TIPDOC_PRO ||'-'|| SOLINM_NUMDOC_PRO) AS IdPromotor"
   g_str_Parame = g_str_Parame & "        ,NVL(CASE WHEN SOLINM_TIPDOC_PRO = 7 THEN TRIM(PR.DATGEN_RAZSOC)"
   g_str_Parame = g_str_Parame & "             ELSE TRIM(SOLINM_RAZSOC_PRO) END, '-') AS PROMOTOR"
   g_str_Parame = g_str_Parame & "        ,NVL(CASE WHEN SOLINM_TIPDOC_CON= 0 THEN '-' "
   g_str_Parame = g_str_Parame & "             ELSE TRIM(SOLINM_TIPDOC_CON ||'-'|| SOLINM_NUMDOC_CON) END, '-') AS IdConstructor"
   g_str_Parame = g_str_Parame & "        ,NVL(CASE WHEN SOLINM_TIPDOC_CON = 0 THEN '-' "
   g_str_Parame = g_str_Parame & "                  WHEN SOLINM_TIPDOC_CON = 1 THEN TRIM(SOLINM_RAZSOC_CON) "
   g_str_Parame = g_str_Parame & "             ELSE TRIM(CN.DATGEN_RAZSOC) END, '-') AS CONSTRUCT"
   g_str_Parame = g_str_Parame & "        ,SOLMAE_CODPRD, SOLMAE_FECSOL, SP.SUBPRD_DESCRI, SOL.SOLMAE_COMVTA_SOL, EC.EVACRE_CUOMPR, SOL.SOLMAE_CODINS, "
   g_str_Parame = g_str_Parame & "       NVL((SELECT COUNT(*) "
   g_str_Parame = g_str_Parame & "         FROM TRA_SEGEXC "
   g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES ON (SEGEXC_MOTEXC=PARDES_CODITE AND PARDES_CODGRP=42) "
   g_str_Parame = g_str_Parame & "        WHERE SEGEXC_NUMSOL = SOL.SOLMAE_NUMERO),0) AS NRO_CRED_EXCEP "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE SOL"
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PRODUC    ON (SOLMAE_CODPRD=PRODUC_CODIGO)"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES MO ON (PARDES_CODGRP=204 AND SOLMAE_TIPMON=PARDES_CODITE)"
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_SUBPRD SP ON SP.SUBPRD_CODPRD = SOLMAE_CODPRD AND SP.SUBPRD_CODSUB=SOL.SOLMAE_CODSUB"
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_PARPRD S  ON S.PARPRD_CODPRD = SOLMAE_CODPRD AND S.PARPRD_CODSUB = SOLMAE_CODSUB AND S.PARPRD_CODGRP = '003' AND S.PARPRD_CODITE = '0'||SOLMAE_CODMOD "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN    ON (SOLMAE_TITTDO=DATGEN_TIPDOC AND SOLMAE_TITNDO=DATGEN_NUMDOC)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN CRE_SOLINM    ON (SOLMAE_NUMERO=SOLINM_NUMSOL)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN    ON (DATGEN_CODIGO=SOLINM_PRYCOD)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES DR ON (DR.PARDES_CODGRP=214 AND SOLINM_PRYMCS=DR.PARDES_CODITE)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN PR ON (SOLINM_TIPDOC_PRO=PR.DATGEN_EMPTDO AND SOLINM_NUMDOC_PRO=PR.DATGEN_EMPNDO)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN EMP_DATGEN CN ON (SOLINM_TIPDOC_CON=CN.DATGEN_EMPTDO AND SOLINM_NUMDOC_CON=CN.DATGEN_EMPNDO)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES IT ON (IT.PARDES_CODGRP=002 AND IT.PARDES_CODITE=SOLMAE_CODINS)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN TRA_SEGUIM    ON (SEGUIM_NUMSOL=SOLMAE_NUMERO AND SEGUIM_CODINS=SOLMAE_CODINS)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES TI ON (TI.PARDES_CODGRP=023 AND TI.PARDES_CODITE=SEGUIM_SITUAC)"
   g_str_Parame = g_str_Parame & "   LEFT JOIN TRA_EVACRE EC ON EC.EVACRE_NUMSOL = SOL.SOLMAE_NUMERO"
   g_str_Parame = g_str_Parame & "  WHERE  "
   If chk_Produc.Value = 0 Then
       g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   If modgen_g_int_TipUsu = 20121 Then          'Si Tipo de Usuario es Consejero Hipotecario
      g_str_Parame = g_str_Parame & "SOLMAE_CONHIP = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   ElseIf modgen_g_int_TipUsu = 20111 Then      'Si Tipo de Usuario es Ejecutivo de Seguimiento
      g_str_Parame = g_str_Parame & "SOLMAE_EJESEG = '" & moddat_gf_Buscar_CodEje_UsuSis(modgen_g_str_CodUsu) & "' AND "
   End If
   g_str_Parame = g_str_Parame & " SOLMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY CLIENTE "
       
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
    
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!NROSOL)
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!CLIENTE)
      
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!FECHASOL)
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!FECHAINS)
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!INSTANCIA)
      
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!SITUAINSTANCIA)
      
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!CONSEJERO)
      
      grd_Listad.Col = 8
      grd_Listad.Text = g_rst_Princi!SOLMAE_CODPRD
      
      grd_Listad.Col = 9
      grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_FECSOL)
      
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(g_rst_Princi!MONEDA)
      
      grd_Listad.Col = 11
      grd_Listad.Text = CStr(g_rst_Princi!MONTO)
      
      grd_Listad.Col = 12
      grd_Listad.Text = CStr(g_rst_Princi!PRYMICASITA)
      
      grd_Listad.Col = 13
      grd_Listad.Text = CStr(g_rst_Princi!PROYECTO)
      
      grd_Listad.Col = 14
      grd_Listad.Text = CStr(g_rst_Princi!IdPromotor)
      
      grd_Listad.Col = 15
      grd_Listad.Text = CStr(g_rst_Princi!PROMOTOR)
      
      grd_Listad.Col = 16
      grd_Listad.Text = CStr(g_rst_Princi!IdConstructor)
      
      grd_Listad.Col = 17
      grd_Listad.Text = CStr(g_rst_Princi!CONSTRUCT)
      
      grd_Listad.Col = 18
      grd_Listad.Text = CStr(g_rst_Princi!DNI)
      
      grd_Listad.Col = 19
      grd_Listad.Text = Trim(CStr(g_rst_Princi!MODALIDAD_PRESTAMO))
      
      grd_Listad.Col = 20
      grd_Listad.Text = Trim(CStr(g_rst_Princi!SUBPRD_DESCRI))
      
      grd_Listad.Col = 21
      grd_Listad.Text = Trim(CStr(g_rst_Princi!SOLMAE_COMVTA_SOL))
      
      grd_Listad.Col = 22
      If IsNull(g_rst_Princi!EVACRE_CUOMPR) Then
         grd_Listad.Text = "0.00 "
      Else
         grd_Listad.Text = Trim(CStr(g_rst_Princi!EVACRE_CUOMPR))
      End If
      
      grd_Listad.Col = 23
      grd_Listad.Text = ""
      
      grd_Listad.Col = 24
      grd_Listad.Text = ""
      
      If CInt(g_rst_Princi!SOLMAE_CODINS) > 21 Then
         If ff_GasAdm(g_rst_Princi!SOLICITUD) Then
            grd_Listad.Col = 23
            grd_Listad.Text = "X"
         Else
            grd_Listad.Col = 24
            grd_Listad.Text = "X"
         End If
      End If
      
      grd_Listad.Col = 25
      grd_Listad.Text = ""
      If g_rst_Princi!NRO_CRED_EXCEP > 0 Then
         grd_Listad.Text = "X"
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   'Ordenando por Nombre de Clientes
   pnl_Tit_NomCli.Tag = "A"
   Call gs_SorteaGrid(grd_Listad, 3, "C")
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Call gs_UbiIniGrid(grd_Listad)
   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Function ff_GasAdm(ByVal p_NumSol As String) As Boolean
Dim r_rst_PagGto  As Recordset
   
   ff_GasAdm = False
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND GASADM_CODGAS = 11 "
   g_str_Parame = g_str_Parame & "   AND GASADM_SITUAC = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PagGto, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_PagGto.BOF And r_rst_PagGto.EOF) Then
      r_rst_PagGto.MoveFirst
      Do While Not r_rst_PagGto.EOF
         ff_GasAdm = True
         r_rst_PagGto.MoveNext
      Loop
   End If
   
   r_rst_PagGto.Close
   Set r_rst_PagGto = Nothing
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_NroFil, 1) = "PRODUCTO":                 .Columns("A").ColumnWidth = 50:     .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 2) = "SUB-PRODUCTO":             .Columns("B").ColumnWidth = 70:     .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 3) = "NRO SOLICITUD":            .Columns("C").ColumnWidth = 16:     .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4) = "DNI TITULAR":              .Columns("D").ColumnWidth = 12:     .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5) = "APELLIDOS Y NOMBRES":      .Columns("E").ColumnWidth = 42:     .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 6) = "MONEDA":                   .Columns("F").ColumnWidth = 21:     .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 7) = "VAL. INMUEBLE":            .Columns("G").ColumnWidth = 15:     .Columns("G").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 8) = "MONTO PREST.":             .Columns("H").ColumnWidth = 15:     .Columns("H").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 9) = "MAX. CUO. APR.":           .Columns("I").ColumnWidth = 15:     .Columns("I").HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 10) = "F. SOLICITUD":            .Columns("J").ColumnWidth = 12:     .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 11) = "F. INSTANCIA":            .Columns("K").ColumnWidth = 12:     .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 12) = "INSTANCIA ACTUAL":        .Columns("L").ColumnWidth = 30:     .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 13) = "SITUACIÓN INSTANCIA":     .Columns("M").ColumnWidth = 43:     .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 14) = "CONSEJ. HIPOTECARIO":     .Columns("N").ColumnWidth = 21:     .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 15) = "PRY MICASITA":            .Columns("O").ColumnWidth = 15:     .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 16) = "PROYECTO":                .Columns("P").ColumnWidth = 50:     .Columns("P").HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 17) = "DOI PROMOTOR":            .Columns("Q").ColumnWidth = 17:     .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 18) = "PROMOTOR":                .Columns("R").ColumnWidth = 60:     .Columns("R").HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 19) = "DOI CONSTRUCTOR":         .Columns("S").ColumnWidth = 17:     .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 20) = "CONSTRUCTOR":             .Columns("T").ColumnWidth = 50:     .Columns("T").HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 21) = "MODALIDAD":               .Columns("U").ColumnWidth = 40:     .Columns("U").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 22) = "APROBADO C/GTO.CIERRE":   .Columns("V").ColumnWidth = 25:     .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 23) = "APROBADO S/GTO.CIERRE":   .Columns("W").ColumnWidth = 25:     .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 24) = "APROB. CRED. EXCEP.":     .Columns("X").ColumnWidth = 20:     .Columns("X").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 25)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 25)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 20)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 18)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 10)
         .Cells(r_int_NroFil, 7) = Format(grd_Listad.TextMatrix(r_int_nroaux, 21), "###,##0.00")
         .Cells(r_int_NroFil, 8) = Format(grd_Listad.TextMatrix(r_int_nroaux, 11), "###,##0.00")
         .Cells(r_int_NroFil, 9) = Format(grd_Listad.TextMatrix(r_int_nroaux, 22), "###,##0.00")
         .Cells(r_int_NroFil, 10) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 11) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 13) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_NroFil, 14) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_nroaux, 12)
         .Cells(r_int_NroFil, 16) = grd_Listad.TextMatrix(r_int_nroaux, 13)
         .Cells(r_int_NroFil, 17) = grd_Listad.TextMatrix(r_int_nroaux, 14)
         .Cells(r_int_NroFil, 18) = grd_Listad.TextMatrix(r_int_nroaux, 15)
         .Cells(r_int_NroFil, 19) = grd_Listad.TextMatrix(r_int_nroaux, 16)
         .Cells(r_int_NroFil, 20) = grd_Listad.TextMatrix(r_int_nroaux, 17)
         .Cells(r_int_NroFil, 21) = grd_Listad.TextMatrix(r_int_nroaux, 19)
         .Cells(r_int_NroFil, 22) = grd_Listad.TextMatrix(r_int_nroaux, 23)
         .Cells(r_int_NroFil, 23) = grd_Listad.TextMatrix(r_int_nroaux, 24)
         .Cells(r_int_NroFil, 24) = grd_Listad.TextMatrix(r_int_nroaux, 25)
         r_int_NroFil = r_int_NroFil + 1
      Next
      
      .Columns("G").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("H").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("I").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub chk_Produc_Click()
   If chk_Produc.Value = 1 Then
      cmb_Produc.ListIndex = -1
      cmb_Produc.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   ElseIf chk_Produc.Value = 0 Then
      cmb_Produc.Enabled = True
      Call gs_SetFocus(cmb_Produc)
   End If
End Sub

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_FlgActEnv = 0
   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_ConHip_Click()
   If Len(Trim(pnl_Tit_ConHip.Tag)) = 0 Or pnl_Tit_ConHip.Tag = "D" Then
      pnl_Tit_ConHip.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "C")
   Else
      pnl_Tit_ConHip.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecIns_Click()
   If Len(Trim(pnl_Tit_FecIns.Tag)) = 0 Or pnl_Tit_FecIns.Tag = "D" Then
      pnl_Tit_FecIns.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_FecIns.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecSol_Click()
   If Len(Trim(pnl_Tit_FecSol.Tag)) = 0 Or pnl_Tit_FecSol.Tag = "D" Then
      pnl_Tit_FecSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 9, "N")
   Else
      pnl_Tit_FecSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 9, "N-")
   End If
End Sub

Private Sub pnl_Tit_InsAct_Click()
   If Len(Trim(pnl_Tit_InsAct.Tag)) = 0 Or pnl_Tit_InsAct.Tag = "D" Then
      pnl_Tit_InsAct.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_InsAct.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_Produc_Click()
   If Len(Trim(pnl_Tit_Produc.Tag)) = 0 Or pnl_Tit_Produc.Tag = "D" Then
      pnl_Tit_Produc.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Produc.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_SitIns_Click()
   If Len(Trim(pnl_Tit_SitIns.Tag)) = 0 Or pnl_Tit_SitIns.Tag = "D" Then
      pnl_Tit_SitIns.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_SitIns.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

