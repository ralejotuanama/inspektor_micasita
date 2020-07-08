VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SegSol_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   1710
   ClientTop       =   1110
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_013.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8205
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   14473
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   23
         Top             =   7380
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_013.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatSol 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_013.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Consulta de Solicitud de Crédito"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_013.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_VerIns 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_013.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Detalle de Instancia"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1200
            Top             =   60
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   3645
         Left            =   30
         TabIndex        =   9
         Top             =   3690
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6429
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
            Height          =   2955
            Left            =   30
            TabIndex        =   0
            Top             =   330
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5212
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   6270
            TabIndex        =   13
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Fin Evaluac."
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   9030
            TabIndex        =   11
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
            ForeColor       =   16777215
            BackColor       =   32768
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
            Left            =   4800
            TabIndex        =   10
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Inicio Evaluac."
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia"
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   7650
            TabIndex        =   22
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Días Transc."
            ForeColor       =   16777215
            BackColor       =   32768
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
         Begin Threed.SSPanel pnl_TotDia 
            Height          =   285
            Left            =   7650
            TabIndex        =   40
            Top             =   3300
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   16777215
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2085
         Left            =   30
         TabIndex        =   6
         Top             =   1560
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3678
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
         Begin Threed.SSPanel pnl_EjeSeg 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Top             =   720
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   8160
            TabIndex        =   15
            Top             =   1380
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_FecDes 
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   1380
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecRec 
            Height          =   315
            Left            =   1440
            TabIndex        =   19
            Top             =   1710
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_ConHip 
            Height          =   315
            Left            =   1440
            TabIndex        =   29
            Top             =   1050
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   32
            Top             =   60
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecAnu 
            Height          =   315
            Left            =   8160
            TabIndex        =   38
            Top             =   1710
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label6 
            Caption         =   "F. Anulación:"
            Height          =   315
            Left            =   6780
            TabIndex        =   39
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1125
         End
         Begin VB.Label Label7 
            Caption         =   "Consejero:"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Image img_Pend 
            Height          =   480
            Left            =   6750
            Picture         =   "AteCli_frm_013.frx":0EA4
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image img_Observ 
            Height          =   480
            Left            =   6300
            Picture         =   "AteCli_frm_013.frx":12E6
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image img_Aprueb 
            Height          =   480
            Left            =   5370
            Picture         =   "AteCli_frm_013.frx":15F0
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image img_Rechaz 
            Height          =   480
            Left            =   5850
            Picture         =   "AteCli_frm_013.frx":18FA
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "F. Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "F. Desembolso:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   1380
            Width           =   1185
         End
         Begin VB.Label Label4 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   6780
            TabIndex        =   14
            Top             =   1380
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Seguim:"
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   720
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            TabIndex        =   5
            Top             =   60
            Width           =   3135
            _Version        =   65536
            _ExtentX        =   5530
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "AteCli_frm_013.frx":1C04
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   24
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   25
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   27
            Top             =   390
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   8040
            TabIndex        =   34
            Top             =   30
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   8040
            TabIndex        =   36
            Top             =   360
            Width           =   3435
            _Version        =   65536
            _ExtentX        =   6059
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SOLICITUD EN TRAMITE"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label8 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   6840
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Ingreso:"
            Height          =   315
            Left            =   6840
            TabIndex        =   35
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_DatSol_Click()
   If moddat_g_int_Situac <> 9 Then
      frm_SegSol_04.Show 1
   Else
      frm_ConSol_01.Show 1
   End If
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call fs_Imp_SolGen
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerIns_Click()
   Dim r_int_CodIns  As Integer
   Dim r_int_Situac  As Integer
   Dim r_dbl_TipCam  As Double
   
   grd_Listad.Col = 5
   r_int_CodIns = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 6
   r_int_Situac = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgAct = 1
   
   Select Case r_int_CodIns
      Case 11     'Atención Comercial
         frm_SegSol_11.Show 1
   
      Case 21     'Aprobación Crediticia Inicial
         frm_SegSol_20.Show 1
      
      Case 31  'Aceptación de Cliente
         If r_int_Situac = 9 Then
            If modgen_g_int_TipUsu = 20900 Or modgen_g_int_TipUsu = 20120 Or modgen_g_int_TipUsu = 20121 Then
               MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            r_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
         
            If r_dbl_TipCam = 0 Then
               MsgBox "Debe solicitar el ingreso del Tipo de Cambio de " & moddat_gf_Consulta_ParDes("204", "2") & ".", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            If moddat_g_int_TipMon <> 1 Then
               r_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
            
               If r_dbl_TipCam = 0 Then
                  MsgBox "Debe solicitar el ingreso del Tipo de Cambio de " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)) & ".", vbExclamation, modgen_g_str_NomPlt
                  Exit Sub
               End If
            End If
            
            frm_SegSol_21.Show 1
         Else
            frm_SegSol_22.Show 1
         End If
      
      Case 32  'Trámites del Cliente
         If r_int_Situac = 9 Then
            If modgen_g_int_TipUsu = 20900 Or modgen_g_int_TipUsu = 20120 Or modgen_g_int_TipUsu = 20121 Then
               MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
               Exit Sub
            End If
            
            frm_SegSol_23.Show 1
         Else
            frm_SegSol_24.Show 1
         End If
         
      Case 41  'Tasación del Inmueble
         frm_SegSol_05.Show 1
         
      Case 42  'Evaluación de Seguros
         frm_SegSol_06.Show 1
         
      Case 51  'Evaluación Legal
         frm_SegSol_07.Show 1
         
      Case 61  'Pólizas de Seguro
         frm_SegSol_08.Show 1
         
      Case 62  'Trámites Mivivienda
         Select Case moddat_g_str_CodPrd
            Case "001": frm_SegSol_13.Show 1
            Case "004": frm_SegSol_14.Show 1
         End Select
         
      Case 71  'Verificación Crediticia
         frm_SegSol_09.Show 1
         
      Case 72  'Autorización de Desembolso
         frm_SegSol_10.Show 1
         
      Case 81  'Desembolso
         frm_SegSol_12.Show 1
         
      Case 91  'Rechazo Administrativo
         frm_SegSol_18.Show 1
         
   End Select
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      
      Call fs_Buscar_DatGen
      Call fs_Buscar_Seguim
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli

   Call fs_Inicia
   Call fs_Buscar_DatGen
   Call fs_Buscar_Seguim
   
   If moddat_g_int_Situac = 9 Then
      cmd_VerIns.Enabled = False
      cmd_Imprim.Enabled = False
   End If
   
   Call gs_SetFocus(grd_Listad)
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 4725
   grd_Listad.ColWidth(1) = 1475
   grd_Listad.ColWidth(2) = 1375
   grd_Listad.ColWidth(3) = 1375
   grd_Listad.ColWidth(4) = 2095
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter

End Sub

Private Sub fs_Buscar_DatGen()
   'Consulta de Datos de la Solicitud
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Obteniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)

   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG)
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_EJESEG))
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP)
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(Trim(g_rst_Princi!SOLMAE_CONHIP))

   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Sub-Producto
   moddat_g_str_CodSub = Trim(g_rst_Princi!SOLMAE_CODSUB & "")

   'Fecha de Ingreso
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   'Fecha de Rechazo
   If g_rst_Princi!SOLMAE_FECREC > 0 Then
      moddat_g_str_FecRec = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
   Else
      moddat_g_str_FecRec = ""
   End If
   
   'Situación
   moddat_g_int_Situac = g_rst_Princi!SOLMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
   
   'Inmueble Identificado
   moddat_g_int_InmIde = g_rst_Princi!SOLMAE_INMIDE
   
   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS
   
   moddat_g_str_NumOpe = ""
   moddat_g_str_FecDes = ""
   
   If g_rst_Princi!SOLMAE_SITUAC = 2 Then
      'Obteniendo Información del Desembolso
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMSOL = '" & moddat_g_str_NumSol & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      'Número de Operación
      moddat_g_str_NumOpe = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
      
      'Fecha de Desembolso
      moddat_g_str_FecDes = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   'Fecha de Anulación
   moddat_g_str_FecAnu = ""
   If g_rst_Princi!SOLMAE_SITUAC = 9 Then
      moddat_g_str_FecAnu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
   End If
   
   'Datos del Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!SOLMAE_CYGTDO
   moddat_g_str_CygNDo = Trim(g_rst_Princi!SOLMAE_CYGNDO & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing


   'Asignando Valores
   pnl_EjeSeg.Caption = moddat_g_str_NomEjeSeg
   pnl_ConHip.Caption = moddat_g_str_NomConHip
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_FecDes.Caption = moddat_g_str_FecDes
   pnl_FecRec.Caption = moddat_g_str_FecRec
   pnl_FecAnu.Caption = moddat_g_str_FecAnu
   pnl_NumOpe.Caption = moddat_g_str_NumOpe
   
   pnl_Situac.Caption = moddat_g_str_Situac
   Select Case moddat_g_int_Situac
      Case 1: pnl_Situac.ForeColor = modgen_g_con_ColAzu
      Case 2: pnl_Situac.ForeColor = modgen_g_con_ColVer
      Case 3: pnl_Situac.ForeColor = modgen_g_con_ColRoj
   End Select
End Sub

Private Sub fs_Buscar_Seguim()
   Dim r_int_DiaTra     As Integer
   Dim r_int_DiaTas     As Integer
   Dim r_int_DiaSeg     As Integer
   Dim r_int_DiaPol     As Integer
   Dim r_int_DiaMVi     As Integer
   
   Call gs_LimpiaGrid(grd_Listad)
   
   r_int_DiaTra = 0
   r_int_DiaTas = 0
   r_int_DiaSeg = 0
   r_int_DiaPol = 0
   r_int_DiaMVi = 0
      
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "'"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'Instancia
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGUIM_CODINS, "000000"))
      
      grd_Listad.Col = 5
      grd_Listad.Text = g_rst_Princi!SEGUIM_CODINS
      
      'Fecha de Inicio
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
      'Fecha de Fin
      grd_Listad.Col = 2
      If g_rst_Princi!SEGUIM_FECFIN > 0 Then
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
         
         'Días Transcurridos
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!SEGUIM_DIATRA)
         
         If g_rst_Princi!SEGUIM_CODINS = 41 Or g_rst_Princi!SEGUIM_CODINS = 42 Then
            If g_rst_Princi!SEGUIM_CODINS = 41 Then
               r_int_DiaTas = g_rst_Princi!SEGUIM_DIATRA
            Else
               r_int_DiaSeg = g_rst_Princi!SEGUIM_DIATRA
            End If
            
            If g_rst_Princi!SEGUIM_CODINS = 42 Then
               If r_int_DiaTas > r_int_DiaSeg Then
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaTas
               Else
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaSeg
               End If
            End If
         ElseIf g_rst_Princi!SEGUIM_CODINS = 61 Or g_rst_Princi!SEGUIM_CODINS = 62 Then
            If g_rst_Princi!SEGUIM_CODINS = 61 Then
               r_int_DiaPol = g_rst_Princi!SEGUIM_DIATRA
            Else
               r_int_DiaMVi = g_rst_Princi!SEGUIM_DIATRA
            End If
            
            If g_rst_Princi!SEGUIM_CODINS = 62 Or (g_rst_Princi!SEGUIM_CODINS = 61 And moddat_g_str_CodPrd = "002") Then
               If r_int_DiaPol > r_int_DiaMVi Then
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaPol
               Else
                  r_int_DiaTra = r_int_DiaTra + r_int_DiaMVi
               End If
            End If
         Else
            r_int_DiaTra = r_int_DiaTra + g_rst_Princi!SEGUIM_DIATRA
         End If
      Else
         If moddat_g_int_Situac = 1 Then
            r_int_DiaTra = r_int_DiaTra + CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
         Else
            r_int_DiaTra = r_int_DiaTra + CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
         End If
      End If
      
      'Situación
      grd_Listad.Col = 4
      Select Case g_rst_Princi!SEGUIM_SITUAC
         Case 1, 7
            Set grd_Listad.CellPicture = img_Aprueb
            grd_Listad.CellPictureAlignment = flexAlignCenterTop
            
         Case 2
            Set grd_Listad.CellPicture = img_Rechaz
            grd_Listad.CellPictureAlignment = flexAlignCenterTop
         
         Case 3
            Set grd_Listad.CellPicture = img_Observ
            grd_Listad.CellPictureAlignment = flexAlignCenterTop
            
         Case 8, 9
            Set grd_Listad.CellPicture = img_Pend
            grd_Listad.CellPictureAlignment = flexAlignCenterTop
      End Select
      
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!SEGUIM_SITUAC)
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If moddat_g_int_Situac <> 3 And moddat_g_int_Situac <> 9 Then
      pnl_TotDia.Caption = CStr(r_int_DiaTra)
   Else
      If moddat_g_int_Situac = 3 Then
         pnl_TotDia.Caption = CStr(CInt(CDate(moddat_g_str_FecRec) - CDate(moddat_g_str_FecIng)))
      Else
         pnl_TotDia.Caption = CStr(CInt(CDate(moddat_g_str_FecAnu) - CDate(moddat_g_str_FecIng)))
      End If
   End If

   Call gs_UbiIniGrid(grd_Listad)

   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_VerIns_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Imp_SolGen()
   Dim r_str_NumSol     As String
   Dim r_str_FecFin     As String
   Dim r_int_Situac     As Integer
   
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC1"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC4"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_SOLIC5"
   DoEvents
   
   'Cabecera de Reporte
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_str_NumSol = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      r_str_FecFin = ""
         
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC1 WHERE SOLIC1_NUMSOL = '" & r_str_NumSol & "'"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
                           
      moddat_g_rst_RecDAO("SOLIC1_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
      moddat_g_rst_RecDAO("SOLIC1_NUMSOL") = r_str_NumSol
      moddat_g_rst_RecDAO("SOLIC1_DOCIDE") = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      moddat_g_rst_RecDAO("SOLIC1_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      moddat_g_rst_RecDAO("SOLIC1_FECING") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
      If g_rst_Princi!SOLMAE_SITUAC = 9 Then
         r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
      Else
         moddat_g_rst_RecDAO("SOLIC1_FECANU") = ""
      End If
      
      r_int_Situac = g_rst_Princi!SOLMAE_SITUAC
      
      If g_rst_Princi!SOLMAE_SITUAC = 1 Or g_rst_Princi!SOLMAE_SITUAC = 3 Then
         If g_rst_Princi!SOLMAE_SITUAC = 1 Or (g_rst_Princi!SOLMAE_SITUAC = 3 And Trim(g_rst_Princi!SOLMAE_TIPREC) = 1) Then
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
         ElseIf g_rst_Princi!SOLMAE_SITUAC = 3 Then
            moddat_g_rst_RecDAO("SOLIC1_CODINS") = 91
            moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", CStr(91))
         End If
      Else
         moddat_g_rst_RecDAO("SOLIC1_CODINS") = 0
         moddat_g_rst_RecDAO("SOLIC1_NOMINS") = ""
      End If
      
      If g_rst_Princi!SOLMAE_SITUAC = 1 Then
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = moddat_gf_Consulta_ParDes("004", Trim(g_rst_Princi!SOLMAE_SITINS))
      Else
         moddat_g_rst_RecDAO("SOLIC1_SITINS") = ""
      End If
      
      If g_rst_Princi!SOLMAE_SITUAC = 3 Then
         r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
      
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECREC))
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SOLMAE_MOTREC))
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = ff_ObsRec(g_rst_Princi!SOLMAE_NUMERO, g_rst_Princi!SOLMAE_TIPREC) & " "
      Else
         moddat_g_rst_RecDAO("SOLIC1_FECREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_TIPREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_MOTREC") = ""
         moddat_g_rst_RecDAO("SOLIC1_OBSERV") = " "
      End If
      
      If g_rst_Princi!SOLMAE_SITUAC = 2 Then
         g_str_Parame = "SELECT * FROM CRE_HIPMAE B WHERE "
         g_str_Parame = g_str_Parame & "HIPMAE_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            If g_rst_Genera!HIPMAE_FECESC > 0 Then
               r_str_FecFin = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECESC))
            Else
               moddat_g_rst_RecDAO("SOLIC1_CODINS") = g_rst_Princi!SOLMAE_CODINS
               moddat_g_rst_RecDAO("SOLIC1_NOMINS") = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
            End If
            
            moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = Left(g_rst_Genera!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Genera!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Genera!HIPMAE_NUMOPE, 5)
            moddat_g_rst_RecDAO("SOLIC1_FECDES") = gf_FormatoFecha(CStr(g_rst_Genera!HIPMAE_FECDES))
            moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = g_rst_Genera!HIPMAE_MTOPRE
            moddat_g_rst_RecDAO("SOLIC1_MONEDA") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Genera!HIPMAE_MONEDA))
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      Else
         moddat_g_rst_RecDAO("SOLIC1_NUMOPE") = ""
         moddat_g_rst_RecDAO("SOLIC1_FECDES") = ""
         moddat_g_rst_RecDAO("SOLIC1_MTOPRE") = 0
         moddat_g_rst_RecDAO("SOLIC1_MONEDA") = ""
      End If
      
      If Len(Trim(r_str_FecFin)) = 0 Then
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
      Else
         moddat_g_rst_RecDAO("SOLIC1_TPOTRA") = CInt(CDate(r_str_FecFin) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))))
      End If
      
      moddat_g_rst_RecDAO("SOLIC1_SITUAC") = moddat_gf_Consulta_ParDes("020", CStr(g_rst_Princi!SOLMAE_SITUAC))
      moddat_g_rst_RecDAO("SOLIC1_CONHIP") = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
      
      
      moddat_g_rst_RecDAO("SOLIC1_NUMOBS") = 0
      moddat_g_rst_RecDAO("SOLIC1_INIINS") = ""
      moddat_g_rst_RecDAO("SOLIC1_FININS") = ""
      moddat_g_rst_RecDAO("SOLIC1_TPOINS") = 0
      moddat_g_rst_RecDAO("SOLIC1_INIOBS") = ""
      moddat_g_rst_RecDAO("SOLIC1_FINOBS") = ""
      moddat_g_rst_RecDAO("SOLIC1_TPOOBS") = 0
      
      moddat_g_rst_RecDAO("SOLIC1_MODALI") = ""
      If g_rst_Princi!SOLMAE_SITUAC = 2 And Len(Trim(g_rst_Princi!SOLMAE_CODMOD)) > 0 Then
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "000")) Then
            moddat_g_rst_RecDAO("SOLIC1_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
   End If
   
   DoEvents
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Detalle por Instancias)
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGUIM_CODINS ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC4 WHERE SOLIC4_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("SOLIC4_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC4_CODINS") = g_rst_Princi!SEGUIM_CODINS
         moddat_g_rst_RecDAO("SOLIC4_NOMINS") = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGUIM_CODINS, "000000"))
      
         moddat_g_rst_RecDAO("SOLIC4_FECINI") = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
         If g_rst_Princi!SEGUIM_FECFIN > 0 Then
            moddat_g_rst_RecDAO("SOLIC4_FECFIN") = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
            moddat_g_rst_RecDAO("SOLIC4_TPOTRA") = g_rst_Princi!SEGUIM_DIATRA
         Else
            If r_int_Situac = 1 Then
               moddat_g_rst_RecDAO("SOLIC4_TPOTRA") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))))
            End If
         End If
      
         moddat_g_rst_RecDAO("SOLIC4_SITUAC") = moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
      
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   DoEvents
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando en DAO (Detalle por Ocurrencias)
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGDET_CODINS ASC, SEGDET_FECOCU ASC, SEGHORCRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_SOLIC5 WHERE SOLIC5_NUMSOL = '" & r_str_NumSol & "'"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("SOLIC5_NUMSOL") = r_str_NumSol
         moddat_g_rst_RecDAO("SOLIC5_CODINS") = g_rst_Princi!SEGDET_CODINS
         moddat_g_rst_RecDAO("SOLIC5_NOMINS") = moddat_gf_Consulta_ParDes("002", Format(g_rst_Princi!SEGDET_CODINS, "000000"))
      
         moddat_g_rst_RecDAO("SOLIC5_FECOCU") = gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))
         moddat_g_rst_RecDAO("SOLIC5_HOROCU") = gf_FormatoHora(CStr(g_rst_Princi!SEGHORCRE))
      
         moddat_g_rst_RecDAO("SOLIC5_CODOCU") = g_rst_Princi!SEGDET_CODOCU
         moddat_g_rst_RecDAO("SOLIC5_DESOCU") = moddat_gf_Consulta_ParDes("004", CStr(g_rst_Princi!SEGDET_CODOCU))
         
         If g_rst_Princi!SEGDET_NUMOBS > 0 Then
            moddat_g_rst_RecDAO("SOLIC5_NUMOBS") = g_rst_Princi!SEGDET_NUMOBS
            moddat_g_rst_RecDAO("SOLIC5_OBSERV") = Trim(g_rst_Princi!SEGDET_OBSERV & "")
            
            If g_rst_Princi!SEGFECACT > 0 Then
               moddat_g_rst_RecDAO("SOLIC5_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
               moddat_g_rst_RecDAO("SOLIC5_HORDES") = gf_FormatoHora(CStr(g_rst_Princi!SEGHORACT))
               moddat_g_rst_RecDAO("SOLIC5_DESCAR") = Trim(g_rst_Princi!SEGDET_OBSDES & "")
               moddat_g_rst_RecDAO("SOLIC5_DIAOBS") = CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))))
            Else
               moddat_g_rst_RecDAO("SOLIC5_FECDES") = ""
               moddat_g_rst_RecDAO("SOLIC5_HORDES") = ""
               moddat_g_rst_RecDAO("SOLIC5_DESCAR") = " "
               moddat_g_rst_RecDAO("SOLIC5_DIAOBS") = CInt(Date - CDate(gf_FormatoFecha(CStr(g_rst_Princi!SEGDET_FECOCU))))
            End If
         Else
            moddat_g_rst_RecDAO("SOLIC5_NUMOBS") = 0
            moddat_g_rst_RecDAO("SOLIC5_OBSERV") = " "
            moddat_g_rst_RecDAO("SOLIC5_FECDES") = ""
            moddat_g_rst_RecDAO("SOLIC5_HORDES") = ""
            moddat_g_rst_RecDAO("SOLIC5_DESCAR") = " "
         End If
      
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   DoEvents
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SOLHIP_05.RPT"
   crp_Imprim.Action = 1
End Sub

Private Function ff_ObsRec(ByVal p_NumSol As String, ByVal p_TipRec As Integer) As String
   ff_ObsRec = " "
   
   If p_TipRec = 1 Then
      g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
      g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 13 "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      
      ff_ObsRec = Trim(g_rst_Genera!SEGDET_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   ElseIf p_TipRec = 3 Then
      g_str_Parame = "SELECT * FROM TRA_RECADM WHERE "
      g_str_Parame = g_str_Parame & "RECADM_NUMSOL = '" & p_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Function
      End If
   
      DoEvents
      g_rst_Genera.MoveFirst
      
      ff_ObsRec = Trim(g_rst_Genera!RECADM_OBSERV & "")
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
End Function


