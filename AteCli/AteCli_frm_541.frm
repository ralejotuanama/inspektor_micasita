VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Con_CreHip_52 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   3540
   ClientTop       =   930
   ClientWidth     =   10425
   Icon            =   "AteCli_frm_541.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10485
      _Version        =   65536
      _ExtentX        =   18494
      _ExtentY        =   17806
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
         Height          =   3615
         Left            =   30
         TabIndex        =   1
         Top             =   6420
         Width           =   10395
         _Version        =   65536
         _ExtentX        =   18336
         _ExtentY        =   6376
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
         Begin Threed.SSPanel pnl_Cuo_TotSal 
            Height          =   315
            Left            =   8610
            TabIndex        =   2
            Top             =   3240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cuo_TotPag 
            Height          =   315
            Left            =   7320
            TabIndex        =   3
            Top             =   3240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Cuo_TotDeu 
            Height          =   315
            Left            =   6030
            TabIndex        =   4
            Top             =   3240
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
            ForeColor       =   16777215
            BackColor       =   192
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
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Cuotas 
            Height          =   2565
            Left            =   60
            TabIndex        =   5
            Top             =   630
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   4524
            _Version        =   393216
            Rows            =   11
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   330
            Width           =   765
            _Version        =   65536
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuota"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencim."
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   6030
            TabIndex        =   8
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Cuota"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   2130
            TabIndex        =   9
            Top             =   330
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "D. Atraso"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   7320
            TabIndex        =   10
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Pagado"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   4740
            TabIndex        =   11
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ult. Pago"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   8610
            TabIndex        =   12
            Top             =   330
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Deudor"
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   3120
            TabIndex        =   13
            Top             =   330
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Begin VB.Label Label12 
            Caption         =   "Lista de Cuotas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales ==> US$ "
            Height          =   255
            Left            =   4350
            TabIndex        =   14
            Top             =   3270
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
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
            Height          =   585
            Left            =   660
            TabIndex        =   17
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Consulta de Crédito Hipotecario"
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
            Picture         =   "AteCli_frm_541.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4935
         Left            =   30
         TabIndex        =   18
         Top             =   1440
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
         _ExtentY        =   8705
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
            Height          =   4575
            Left            =   60
            TabIndex        =   19
            Top             =   330
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   8070
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Datos del Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   645
         Left            =   30
         TabIndex        =   21
         Top             =   750
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
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
         Begin VB.CommandButton cmd_EstCta 
            Height          =   585
            Left            =   3630
            Picture         =   "AteCli_frm_541.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Imprimir Estado de Cuenta"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCli 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_541.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Información del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_SolCre 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_541.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Datos de la Solicitud de Crédito"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9750
            Picture         =   "AteCli_frm_541.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatHip 
            Height          =   585
            Left            =   3030
            Picture         =   "AteCli_frm_541.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Datos de la Hipoteca"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Height          =   585
            Left            =   2430
            Picture         =   "AteCli_frm_541.frx":1940
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Ver Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_541.frx":1C4A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Cronogramas de Pago"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_541.frx":1F54
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   1740
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_Con_CreHip_52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_DatCli_Click()
   frm_Con_CreHip_53.Show 1
End Sub

Private Sub cmd_DatHip_Click()
   frm_Con_CreHip_58.Show 1
End Sub

Private Sub cmd_DatInm_Click()
   frm_Con_CreHip_54.Show 1
End Sub

Private Sub cmd_EstCta_Click()
   frm_Con_CreHip_59.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   frm_Con_CreHip_55.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SolCre_Click()
   frm_Con_SolHip_52.Show 1
End Sub

Private Sub cmd_VerPag_Click()
   frm_Con_CreHip_56.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_Listad.ColWidth(0) = 2850
   grd_Listad.ColWidth(1) = 9680
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Cuotas
   grd_Cuotas.ColWidth(0) = 750
   grd_Cuotas.ColWidth(1) = 1295
   grd_Cuotas.ColWidth(2) = 1005
   grd_Cuotas.ColWidth(3) = 1625
   grd_Cuotas.ColWidth(4) = 1295
   grd_Cuotas.ColWidth(5) = 1295
   grd_Cuotas.ColWidth(6) = 1295
   grd_Cuotas.ColWidth(7) = 1295
   grd_Cuotas.ColAlignment(0) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(1) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(2) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(3) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(4) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(5) = flexAlignRightCenter
   grd_Cuotas.ColAlignment(6) = flexAlignRightCenter
   grd_Cuotas.ColAlignment(7) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_Cuotas)
   pnl_Cuo_TotDeu.Caption = "0.00 "
   pnl_Cuo_TotPag.Caption = "0.00 "
   pnl_Cuo_TotSal.Caption = "0.00 "
End Sub

Private Sub fs_Buscar()

   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   
   'Buscando en Maestro de Clientes
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera!DATGEN_TDOVIN > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                       grd_Listad.Text = "Vinculado"
      
      grd_Listad.Col = 1
      If g_rst_Genera!DATGEN_TIPVIN = 1 Then
         grd_Listad.Text = "TRABAJADOR"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 2 Or g_rst_Genera!DATGEN_TIPVIN = 3 Then
         grd_Listad.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 4 Then
         grd_Listad.Text = "FUNCIONARIO"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 5 Then
         grd_Listad.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      Else
         grd_Listad.Text = ""
      End If
   End If
   
   If g_rst_Genera!DATGEN_TDOACC > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                       grd_Listad.Text = "Accionista"
      grd_Listad.Col = 1
      
      If g_rst_Genera!DATGEN_ACCVIN = 1 Then
         grd_Listad.Text = "ACCIONISTA"
      ElseIf g_rst_Genera!DATGEN_ACCVIN = 2 Then
         grd_Listad.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Genera!DATGEN_TDOACC, Trim(g_rst_Genera!DATGEN_NDOACC)) & ")"
      End If
   End If
   
   modmip_g_int_PaiRes = g_rst_Genera!DATGEN_PAIRES
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
   
   lbl_Totale.Caption = "Totales ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
   
   'Buscando Cuotas
   Call fs_Buscar_Cuotas
   Call gs_SetFocus(grd_Cuotas)
End Sub

Private Sub fs_Buscar_ant()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 3 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
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

   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.CellFontBold = True:           grd_Listad.Text = "Número de Operación"
   grd_Listad.Col = 1:                       grd_Listad.CellFontBold = True:           grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.CellFontBold = True:           grd_Listad.Text = "Situación"
   grd_Listad.Col = 1:                       grd_Listad.CellFontBold = True:           grd_Listad.Text = moddat_g_str_Situac
   grd_Listad.CellForeColor = modgen_g_con_ColNar
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.CellFontBold = True:           grd_Listad.Text = "Cliente"
   grd_Listad.Col = 1:                       grd_Listad.CellFontBold = True:           grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Cónyuge"
      grd_Listad.Col = 1:                    grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Producto"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Moneda Préstamo"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_Moneda
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Primera Vivienda"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!HIPMAE_PRIVIV)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Modalidad"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_DesMod
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Dirección Inmueble"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_Direcc
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Distrito"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Proyecto Inmobiliario"
      grd_Listad.Col = 1:                    grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         grd_Listad.Text = grd_Listad.Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 2: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Valor Compra Venta"
      grd_Listad.Col = 1:                    grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8:           grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Aporte Propio"
      grd_Listad.Col = 1:                    grd_Listad.CellFontName = "Lucida Console"
      
      'If moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
         grd_Listad.CellFontSize = 8:           grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ")"
      Else
         grd_Listad.CellFontSize = 8:           grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
      End If
   Else
      grd_Listad.Rows = grd_Listad.Rows + 2: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Valor Compra Venta"
      grd_Listad.Col = 1:                    grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8:           grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Aporte Propio"
      grd_Listad.Col = 1:                    grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8:           grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Monto Desembolsado"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Monto Préstamo"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Interés Capitalizado"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Total Préstamo"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Fecha Activación"
   grd_Listad.Col = 1:                       grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Fecha Desembolso"
   grd_Listad.Col = 1:                       grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                    grd_Listad.Text = "Fecha Firma EE.PP"
      grd_Listad.Col = 1:                    grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) = 0 Then ' moddat_g_str_CodPrd <> "002" And moddat_g_str_CodPrd <> "011" Then
      grd_Listad.Rows = grd_Listad.Rows + 2: grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación Mivivienda" '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd):  grd_Listad.Text = "Nro. Operación COFIDE"     '"003"
         Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"     '"004"
         Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"     '"007", "009", "010", "012" "013", "014", "015", "016", "017", "018", "019", "021", "022", "023"
      End Select
      
      grd_Listad.Col = 1:     grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"003"
         grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                       grd_Listad.Text = "Nro. Operación Mivivienda"
         grd_Listad.Col = 1:                       grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                          grd_Listad.Text = "Monto Préstamo (Tramo No Conces.)"
      grd_Listad.Col = 1:                          grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8:                 grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                          grd_Listad.Text = "Monto Préstamo (Tramo Conces.)"
      grd_Listad.Col = 1:                          grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8:                 grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then  '"001" "003"
         grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                       grd_Listad.Text = "Tasa de Interés Mivivienda"
         grd_Listad.Col = 1:                       grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then  '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
         grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                       grd_Listad.Text = "Tasa de Interés COFIDE"
         grd_Listad.Col = 1:                       grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                       grd_Listad.Text = "Tasa de Comisión COFIDE"
         grd_Listad.Col = 1:                       grd_Listad.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Plazo"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Tasa de Interés"
   grd_Listad.Col = 1:                       grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Nro. de Cuotas"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Período de Gracia"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Cuotas Extraordinarias"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Compañía de Seguros"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Tipo de Seguro Desg."
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Tipo Garantía"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Monto Garantía"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Saldo Capital"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Saldo Capital (Tramo Conces.)"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Total Saldo Deudor"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Cuotas Pendientes de Pago"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Días de Atraso"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Consejero Hipotecario"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_NomConHip
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Ejecutivo de Seguimiento"
   grd_Listad.Col = 1:                       grd_Listad.Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscando en Maestro de Clientes
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND DATGEN_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera!DATGEN_TDOVIN > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                       grd_Listad.Text = "Vinculado"
      
      grd_Listad.Col = 1
      If g_rst_Genera!DATGEN_TIPVIN = 1 Then
         grd_Listad.Text = "TRABAJADOR"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 2 Or g_rst_Genera!DATGEN_TIPVIN = 3 Then
         grd_Listad.Text = "VINCULADO A TRABAJADOR (" & modmip_gf_Consulta_NomTra(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 4 Then
         grd_Listad.Text = "FUNCIONARIO"
      ElseIf g_rst_Genera!DATGEN_TIPVIN = 5 Then
         grd_Listad.Text = "VINCULADO A FUNCIONARIO (" & modmip_gf_Consulta_NomOtrFun(g_rst_Genera!DATGEN_TDOVIN, Trim(g_rst_Genera!DATGEN_NDOVIN)) & ")"
      Else
         grd_Listad.Text = ""
      End If
   End If
   
   If g_rst_Genera!DATGEN_TDOACC > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0:                       grd_Listad.Text = "Accionista"
      grd_Listad.Col = 1
      
      If g_rst_Genera!DATGEN_ACCVIN = 1 Then
         grd_Listad.Text = "ACCIONISTA"
      ElseIf g_rst_Genera!DATGEN_ACCVIN = 2 Then
         grd_Listad.Text = "VINCULADO A ACCIONISTA (" & modmip_gf_Consulta_NomAcc(g_rst_Genera!DATGEN_TDOACC, Trim(g_rst_Genera!DATGEN_NDOACC)) & ")"
      End If
   End If
   
   modmip_g_int_PaiRes = g_rst_Genera!DATGEN_PAIRES
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
   
   lbl_Totale.Caption = "Totales ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
   
   'Buscando Cuotas
   Call fs_Buscar_Cuotas
   Call gs_SetFocus(grd_Cuotas)
End Sub

Private Sub fs_Buscar_Cuotas()
Dim r_dbl_Pag_TotCuo    As Double
Dim r_dbl_Pag_Capita    As Double
Dim r_dbl_Pag_Intere    As Double
Dim r_dbl_Pag_SegDes    As Double
Dim r_dbl_Pag_SegViv    As Double
Dim r_dbl_Pag_OtrCar    As Double
Dim r_dbl_Pag_IntMor    As Double
Dim r_dbl_Pag_IntCom    As Double
Dim r_dbl_Pag_GasCob    As Double
Dim r_dbl_Pag_OtrGas    As Double
Dim r_dbl_Deu_TotCuo    As Double
Dim r_dbl_Deu_Capita    As Double
Dim r_dbl_Deu_Intere    As Double
Dim r_dbl_Deu_SegDes    As Double
Dim r_dbl_Deu_SegViv    As Double
Dim r_dbl_Deu_OtrCar    As Double
Dim r_dbl_Deu_IntMor    As Double
Dim r_dbl_Deu_IntCom    As Double
Dim r_dbl_Deu_GasCob    As Double
Dim r_dbl_Deu_OtrGas    As Double
Dim r_dbl_Sal_TotCuo    As Double
Dim r_dbl_Sal_Capita    As Double
Dim r_dbl_Sal_Intere    As Double
Dim r_dbl_Sal_SegDes    As Double
Dim r_dbl_Sal_SegViv    As Double
Dim r_dbl_Sal_OtrCar    As Double
Dim r_dbl_Sal_IntMor    As Double
Dim r_dbl_Sal_IntCom    As Double
Dim r_dbl_Sal_GasCob    As Double
Dim r_dbl_Sal_OtrGas    As Double
Dim r_dbl_Gen_TotDeu    As Double
Dim r_dbl_Gen_TotPag    As Double
Dim r_dbl_Gen_TotSal    As Double

   r_dbl_Gen_TotDeu = 0
   r_dbl_Gen_TotPag = 0
   r_dbl_Gen_TotSal = 0
   
   'Cuotas Vencidas
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Cuotas.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         'A Pagar
         r_dbl_Deu_Capita = CDbl(Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_CAPBBP, "###,###,##0.00"))
         r_dbl_Deu_Intere = CDbl(Format(g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_INTBBP, "###,###,##0.00"))
         r_dbl_Deu_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_Deu_SegViv = CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_Deu_OtrCar = CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_Deu_IntMor = CDbl(Format(g_rst_Princi!HIPCUO_INTMOR, "###,###,##0.00"))
         r_dbl_Deu_IntCom = CDbl(Format(g_rst_Princi!HIPCUO_INTCOM, "###,###,##0.00"))
         r_dbl_Deu_GasCob = CDbl(Format(g_rst_Princi!HIPCUO_GASCOB, "###,###,##0.00"))
         r_dbl_Deu_OtrGas = CDbl(Format(g_rst_Princi!HIPCUO_OTRGAS, "###,###,##0.00"))
         
         r_dbl_Deu_TotCuo = 0
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_Capita
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_Intere
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_SegDes
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_SegViv
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_OtrCar
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_IntMor
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_IntCom
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_GasCob
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_OtrGas
         
         'Pagado
         r_dbl_Pag_Capita = CDbl(Format(g_rst_Princi!HIPCUO_CAPPAG + g_rst_Princi!HIPCUO_CBPPAG, "###,###,##0.00"))
         r_dbl_Pag_Intere = CDbl(Format(g_rst_Princi!HIPCUO_INTPAG + g_rst_Princi!HIPCUO_IBPPAG, "###,###,##0.00"))
         r_dbl_Pag_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00"))
         r_dbl_Pag_SegViv = CDbl(Format(g_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00"))
         r_dbl_Pag_OtrCar = CDbl(Format(g_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00"))
         r_dbl_Pag_IntCom = CDbl(Format(g_rst_Princi!HIPCUO_ICOPAG, "###,###,##0.00"))
         r_dbl_Pag_IntMor = CDbl(Format(g_rst_Princi!HIPCUO_IMOPAG, "###,###,##0.00"))
         r_dbl_Pag_GasCob = CDbl(Format(g_rst_Princi!HIPCUO_GCOPAG, "###,###,##0.00"))
         r_dbl_Pag_OtrGas = CDbl(Format(g_rst_Princi!HIPCUO_OTGPAG, "###,###,##0.00"))
         
         r_dbl_Pag_TotCuo = 0
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Capita
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Intere
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegDes
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegViv
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrCar
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntCom
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntMor
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_GasCob
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrGas
         
         'Saldo Pago
         r_dbl_Sal_Capita = r_dbl_Deu_Capita - r_dbl_Pag_Capita
         r_dbl_Sal_Intere = r_dbl_Deu_Intere - r_dbl_Pag_Intere
         r_dbl_Sal_IntCom = r_dbl_Deu_IntCom - r_dbl_Pag_IntCom
         r_dbl_Sal_IntMor = r_dbl_Deu_IntMor - r_dbl_Pag_IntMor
         r_dbl_Sal_GasCob = r_dbl_Deu_GasCob - r_dbl_Pag_GasCob
         r_dbl_Sal_OtrGas = r_dbl_Deu_OtrGas - r_dbl_Pag_OtrGas
         r_dbl_Sal_SegDes = r_dbl_Deu_SegDes - r_dbl_Pag_SegDes
         r_dbl_Sal_SegViv = r_dbl_Deu_SegViv - r_dbl_Pag_SegViv
         r_dbl_Sal_OtrCar = r_dbl_Deu_OtrCar - r_dbl_Pag_OtrCar
         
         'Total Cuota
         r_dbl_Sal_TotCuo = 0
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_Capita
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_Intere
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_SegDes
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_SegViv
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_OtrCar
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_IntCom
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_IntMor
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_GasCob
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_OtrGas
         
         grd_Cuotas.Rows = grd_Cuotas.Rows + 1
         grd_Cuotas.Row = grd_Cuotas.Rows - 1
         
         grd_Cuotas.Col = 0
         grd_Cuotas.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         
         grd_Cuotas.Col = 1
         grd_Cuotas.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         'Si Situación es No-Pagado
         If g_rst_Princi!HIPCUO_SITUAC = 2 Then
            If moddat_g_int_Situac = 2 Then
               If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) < CDate(moddat_g_str_FecSis) Then
                  grd_Cuotas.Col = 2
                  grd_Cuotas.Text = CStr(CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "VENCIDA"
               Else
                  grd_Cuotas.Col = 2
                  grd_Cuotas.Text = "-"
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "POR VENCER"
               End If
            End If
         Else
            If CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))) > 0 Then
               grd_Cuotas.Col = 2
               grd_Cuotas.Text = CStr(CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
            Else
               grd_Cuotas.Col = 2
               grd_Cuotas.Text = "-"
            End If
            
            grd_Cuotas.Col = 3
            grd_Cuotas.Text = "PAGADA"
         End If
         
         If g_rst_Princi!HIPCUO_FECPAG > 0 Then
            grd_Cuotas.Col = 4
            grd_Cuotas.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))
         End If
      
         'Valor Cuota
         grd_Cuotas.Col = 5
         grd_Cuotas.Text = Format(r_dbl_Deu_TotCuo, "###,###,##0.00")
                           
         'Importe Pagado
         grd_Cuotas.Col = 6
         grd_Cuotas.Text = Format(r_dbl_Pag_TotCuo, "###,###,##0.00")
      
         'Saldo
         grd_Cuotas.Col = 7
         grd_Cuotas.Text = Format(r_dbl_Sal_TotCuo, "###,###,##0.00")
      
         'Sumando Totales
         r_dbl_Gen_TotDeu = r_dbl_Gen_TotDeu + r_dbl_Deu_TotCuo
         r_dbl_Gen_TotPag = r_dbl_Gen_TotPag + r_dbl_Pag_TotCuo
         r_dbl_Gen_TotSal = r_dbl_Gen_TotSal + r_dbl_Sal_TotCuo
      
         g_rst_Princi.MoveNext
      Loop
      
      pnl_Cuo_TotDeu.Caption = Format(r_dbl_Gen_TotDeu, "###,###,##0.00") & " "
      pnl_Cuo_TotPag.Caption = Format(r_dbl_Gen_TotPag, "###,###,##0.00") & " "
      pnl_Cuo_TotSal.Caption = Format(r_dbl_Gen_TotSal, "###,###,##0.00") & " "
      
      grd_Cuotas.Redraw = True
      Call gs_UbiIniGrid(grd_Cuotas)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Cuotas_DblClick()
   If grd_Cuotas.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Cuotas.Col = 0
   moddat_g_int_NumCuo = CInt(grd_Cuotas)
   
   Call gs_RefrescaGrid(grd_Cuotas)
   
   frm_Con_CreHip_60.Show 1
End Sub

Private Sub grd_Cuotas_SelChange()
   If grd_Cuotas.Rows > 2 Then
      grd_Cuotas.RowSel = grd_Cuotas.Row
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

