VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SegSol_23 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   10395
   ClientLeft      =   2460
   ClientTop       =   1065
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_023.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18336
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
      Begin Threed.SSPanel SSPanel34 
         Height          =   1635
         Left            =   30
         TabIndex        =   44
         Top             =   7890
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2884
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisDoc 
            Height          =   1245
            Left            =   30
            TabIndex        =   45
            Top             =   360
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2196
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel35 
            Height          =   285
            Left            =   9870
            TabIndex        =   46
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Recepcionado"
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
         Begin Threed.SSPanel SSPanel36 
            Height          =   285
            Left            =   60
            TabIndex        =   47
            Top             =   60
            Width           =   9825
            _Version        =   65536
            _ExtentX        =   17330
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1755
         Left            =   30
         TabIndex        =   1
         Top             =   1890
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3096
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
         Begin VB.TextBox txt_ObsRec 
            Height          =   645
            Left            =   1440
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1050
            Width           =   10035
         End
         Begin Threed.SSPanel pnl_SalIns 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   390
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
         Begin Threed.SSPanel pnl_SitIns 
            Height          =   315
            Left            =   7650
            TabIndex        =   4
            Top             =   60
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
         Begin Threed.SSPanel pnl_MotRec 
            Height          =   315
            Left            =   1440
            TabIndex        =   5
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
         Begin Threed.SSPanel pnl_IngIns 
            Height          =   315
            Left            =   1440
            TabIndex        =   6
            Top             =   60
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
         Begin VB.Label Label9 
            Caption         =   "F. Salida Inst.:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label Label24 
            Caption         =   "Situación Inst.:"
            Height          =   315
            Left            =   6420
            TabIndex        =   10
            Top             =   30
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Motivo Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label19 
            Caption         =   "Observaciones de Rechazo:"
            Height          =   555
            Left            =   60
            TabIndex        =   8
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label Label7 
            Caption         =   "F. Ingreso Inst.:"
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   12
         Top             =   9570
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
            Left            =   30
            Picture         =   "AteCli_frm_023.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Nueva Observación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_023.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_023.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_023.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   2160
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   1590
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   3240
            Top             =   210
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   14
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
            Height          =   615
            Left            =   630
            TabIndex        =   15
            Top             =   30
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud - Trámites del Cliente"
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
            Picture         =   "AteCli_frm_023.frx":0FDC
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1785
         Left            =   30
         TabIndex        =   16
         Top             =   3690
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3149
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
            Height          =   1125
            Left            =   30
            TabIndex        =   17
            Top             =   630
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1984
            _Version        =   393216
            Rows            =   21
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   330
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ocurrencia"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   2400
            TabIndex        =   19
            Top             =   330
            Width           =   8805
            _Version        =   65536
            _ExtentX        =   15531
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Ocurrencia"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1230
            TabIndex        =   20
            Top             =   330
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "H. Ocurrencia"
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
         Begin VB.Label Label3 
            Caption         =   "Seguimiento de Ocurrencias"
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
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   2805
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   22
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1931
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
            TabIndex        =   23
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
            TabIndex        =   24
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
            Left            =   7650
            TabIndex        =   25
            Top             =   60
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
            Left            =   7650
            TabIndex        =   26
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   27
            Top             =   720
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
         Begin Threed.SSPanel pnl_FecRec 
            Height          =   315
            Left            =   10320
            TabIndex        =   28
            Top             =   60
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
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Ingreso:"
            Height          =   315
            Left            =   6420
            TabIndex        =   32
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   6420
            TabIndex        =   31
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label12 
            Caption         =   "F. Rechazo:"
            Height          =   315
            Left            =   9270
            TabIndex        =   29
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2325
         Left            =   30
         TabIndex        =   35
         Top             =   5520
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4101
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
         Begin MSFlexGridLib.MSFlexGrid grd_GasAdm 
            Height          =   1575
            Left            =   30
            TabIndex        =   37
            Top             =   360
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_TotGas 
            Height          =   315
            Left            =   10020
            TabIndex        =   36
            Top             =   1950
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4020
            TabIndex        =   39
            Top             =   60
            Width           =   2385
            _Version        =   65536
            _ExtentX        =   4207
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Moneda"
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
            Left            =   6390
            TabIndex        =   40
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
            Left            =   7590
            TabIndex        =   41
            Top             =   60
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
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
            Left            =   9540
            TabIndex        =   42
            Top             =   60
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Pago"
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
         Begin VB.Label Label4 
            Caption         =   "Total de Gastos:"
            Height          =   315
            Left            =   8640
            TabIndex        =   43
            Top             =   1950
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgPag  As Integer

Private Sub cmd_Aprueb_Click()
   Dim r_int_TipDoc     As Integer
   Dim r_int_CodAct     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_FlgDoc     As Integer
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodGrp     As String
   Dim r_str_CodIte     As String

   If l_int_FlgPag = 0 Then
      MsgBox "Los Gastos Administrativos no han sido cancelados.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'Validando Documentos a Recibir
   grd_LisDoc.Redraw = False
   
   r_int_FlgDoc = 1
   
   For r_int_Contad = 0 To grd_LisDoc.Rows - 1
      grd_LisDoc.Row = r_int_Contad
      
      grd_LisDoc.Col = 1
      If Trim(grd_LisDoc.Text) = "X" Then
         r_int_FlgDoc = 2
         Exit For
      End If
   Next r_int_Contad
   
   grd_LisDoc.Redraw = True
   
   Call gs_UbiIniGrid(grd_LisDoc)
   
   If r_int_FlgDoc = 1 Then
      MsgBox "Debe seleccionar los Documentos que han sido recibidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_LisDoc)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(pnl_IngIns.Caption))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 32, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 32, 12, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Grabando Documentos Recibidos
   grd_LisDoc.Redraw = False
   For r_int_Contad = 0 To grd_LisDoc.Rows - 1
      grd_LisDoc.Row = r_int_Contad
      
      grd_LisDoc.Col = 1
      If grd_LisDoc.Text = "X" Then
         grd_LisDoc.Col = 2:  r_int_TipDoc = CInt(grd_LisDoc.Text)
         grd_LisDoc.Col = 3:  r_str_CodGrp = grd_LisDoc.Text
         grd_LisDoc.Col = 4:  r_int_CodAct = CInt(grd_LisDoc.Text)
         grd_LisDoc.Col = 5:  r_str_CodIte = grd_LisDoc.Text
      
         If Not moddat_gf_Inserta_SolDoc(moddat_g_str_NumSol, r_int_TipDoc, moddat_g_str_CodPrd, moddat_g_str_CodSub, r_int_CodAct, r_str_CodGrp, r_str_CodIte, Format(CDate(moddat_g_str_FecSis), "yyyymmdd")) Then
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   grd_LisDoc.Redraw = True
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 41) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, 42) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, 41) Then
      Exit Sub
   End If
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "APROBACION EN TRAMITES DEL CLIENTE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_int_Contad     As Integer
   Dim r_str_Moneda     As String
   Dim r_str_NomGas     As String
   Dim r_dbl_Import     As Double

   If modgen_g_int_TipUsu = 20900 Or modgen_g_int_TipUsu = 20120 Or modgen_g_int_TipUsu = 20121 Then
      MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If grd_GasAdm.Rows = 0 Then
      MsgBox "No se han asignado los Gastos Administrativos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If l_int_FlgPag = 1 Then
      MsgBox "Los Gastos Administrativos ya fueron cancelados.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir Hoja Resumen de Gastos de Cierre?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_GADMCB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_GADMDT"
                        
   'Grabando en DAO
   moddat_g_str_CadDAO = "SELECT * FROM RPT_GADMCB WHERE GADMCB_NUMSOL = '" & pnl_NumSol.Caption & "'"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)

   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("GADMCB_NUMSOL") = pnl_NumSol.Caption
   moddat_g_rst_RecDAO("GADMCB_NOMCLI") = moddat_g_str_NomCli
   moddat_g_rst_RecDAO("GADMCB_DOCIDE") = moddat_gf_Consulta_ParDes("203", CStr(moddat_g_int_TipDoc)) & " NRO. " & moddat_g_str_NumDoc
   moddat_g_rst_RecDAO("GADMCB_OBSERV") = " "
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   'Grabando en DAO (Detalle)
   grd_GasAdm.Redraw = False
   For r_int_Contad = 0 To grd_GasAdm.Rows - 1
      grd_GasAdm.Row = r_int_Contad
      
      grd_GasAdm.Col = 0:  r_str_NomGas = grd_GasAdm.Text
      grd_GasAdm.Col = 1:  r_str_Moneda = grd_GasAdm.Text
      grd_GasAdm.Col = 2:  r_dbl_Import = CDbl(grd_GasAdm.Text)
   
      moddat_g_str_CadDAO = "SELECT * FROM RPT_GADMDT WHERE GADMDT_NUMSOL = '" & pnl_NumSol.Caption & "'"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
      moddat_g_rst_RecDAO.AddNew
                           
      moddat_g_rst_RecDAO("GADMDT_NUMSOL") = pnl_NumSol.Caption
      moddat_g_rst_RecDAO("GADMDT_NOMGAS") = r_str_NomGas
      moddat_g_rst_RecDAO("GADMDT_MONEDA") = r_str_Moneda
      moddat_g_rst_RecDAO("GADMDT_IMPORT") = r_dbl_Import
                           
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
   Next r_int_Contad
   
   grd_GasAdm.Redraw = True
   Call gs_UbiIniGrid(grd_GasAdm)
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_GASADM_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = 32
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(pnl_IngIns.Caption))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 32, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 32, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      modgen_g_str_Mail_Asunto = "RECHAZO EN TRAMITES DEL CLIENTE (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_gf_Consulta_ParDes("003", moddat_g_int_MotRec)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_Observ
      
      Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_con_AteCli
      
      moddat_g_int_FlgAct = 2
      
      Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   
   Call fs_Carga_DatGen
   
   Call fs_Buscar_Seguim
   Call fs_Buscar_LisOcu
   Call fs_Buscar_GasAdm
   
   Call fs_Carga_Docume
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_LisOcu.ColWidth(0) = 1155
   grd_LisOcu.ColWidth(1) = 1185
   grd_LisOcu.ColWidth(2) = 8805
   
   grd_LisOcu.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(2) = flexAlignLeftCenter

   grd_GasAdm.ColWidth(0) = 3960
   grd_GasAdm.ColWidth(1) = 2370
   grd_GasAdm.ColWidth(2) = 1190
   grd_GasAdm.ColWidth(3) = 1955
   grd_GasAdm.ColWidth(4) = 1650
   
   grd_GasAdm.ColAlignment(0) = flexAlignLeftCenter
   grd_GasAdm.ColAlignment(1) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(2) = flexAlignRightCenter
   grd_GasAdm.ColAlignment(3) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(4) = flexAlignCenterCenter

   'Inicializando Rejilla
   grd_LisDoc.ColWidth(0) = 9805
   grd_LisDoc.ColWidth(1) = 1290
   grd_LisDoc.ColWidth(2) = 0
   grd_LisDoc.ColWidth(3) = 0
   grd_LisDoc.ColWidth(4) = 0
   grd_LisDoc.ColWidth(5) = 0
   grd_LisDoc.ColWidth(6) = 0
   
   grd_LisDoc.ColAlignment(0) = flexAlignLeftCenter
   grd_LisDoc.ColAlignment(1) = flexAlignCenterCenter
End Sub

Private Sub fs_Carga_DatGen()
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Situac.Caption = moddat_g_str_Situac
   Select Case moddat_g_int_Situac
      Case 1: pnl_Situac.ForeColor = modgen_g_con_ColAzu
      Case 2: pnl_Situac.ForeColor = modgen_g_con_ColVer
      Case 3: pnl_Situac.ForeColor = modgen_g_con_ColRoj
   End Select
   
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecRec.Caption = moddat_g_str_FecRec
End Sub

Private Sub fs_Buscar_Seguim()
   pnl_IngIns.Caption = ""
   pnl_SalIns.Caption = ""
   pnl_SitIns.Caption = ""

   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 32"
   
   If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      g_rst_Princi.MoveFirst
      
      pnl_IngIns.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
      If g_rst_Princi!SEGUIM_FECFIN > 0 Then
         pnl_SalIns.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
      End If
      
      pnl_SitIns.Caption = moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_LisOcu()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   pnl_MotRec.Caption = ""
   txt_ObsRec.Text = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 32 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisOcu.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisOcu.Rows = grd_LisOcu.Rows + 1
      grd_LisOcu.Row = grd_LisOcu.Rows - 1
      
      'Fecha de Ocurrencia
      grd_LisOcu.Col = 0
      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu.Col = 1
      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripción Ocurrencia
      grd_LisOcu.Col = 2
      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGFECACT > 0 Then
         r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         
         grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & r_str_FecOcu
         grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
      End If
      
      If g_rst_Princi!SEGDET_CODOCU = 13 Then
         'Si la Solicitud está rechazada en la Instancia
         pnl_MotRec.Caption = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SEGDET_MOTREC))
         txt_ObsRec.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
End Sub

Private Sub fs_Buscar_GasAdm()
   Dim r_dbl_Import  As Double
   
   l_int_FlgPag = 0
   
   r_dbl_Import = 0
   
   Call gs_LimpiaGrid(grd_GasAdm)
   pnl_TotGas.Caption = "0.00 "
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_GasAdm.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_GasAdm.Rows = grd_GasAdm.Rows + 1
      
      grd_GasAdm.Row = grd_GasAdm.Rows - 1
      
      'Buscando Descripción de Gastos Administrativos
      grd_GasAdm.Col = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")) Then
         grd_GasAdm.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
      End If
      
      'Tipo de Moneda
      grd_GasAdm.Col = 1
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!GASADM_TIPMON))
      
      'Importe
      grd_GasAdm.Col = 2
      grd_GasAdm.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
      
      r_dbl_Import = r_dbl_Import + g_rst_Princi!GASADM_IMPORT
      
      'Situación
      grd_GasAdm.Col = 3
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("001", CStr(g_rst_Princi!GASADM_SITUAC))
      
      'Fecha de Pago
      grd_GasAdm.Col = 4
      If g_rst_Princi!GASADM_PAGFEC > 0 Then
         l_int_FlgPag = 1
         grd_GasAdm.Text = gf_FormatoFecha(CStr(g_rst_Princi!GASADM_PAGFEC))
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   pnl_TotGas.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   
   grd_GasAdm.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_GasAdm)
End Sub

Private Sub fs_Carga_Docume()
   '0 - Descripción
   '1 - Selección
   '2 - Tipo de Origen de Documento
   '3 - Código de Grupo
   '4 - Código de Actividad Económica
   '5 - Código de Item
   
   Call gs_LimpiaGrid(grd_LisDoc)
   
   'Documentos del Vendedor
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '012' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   grd_LisDoc.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_LisDoc.Rows = grd_LisDoc.Rows + 1
      
      grd_LisDoc.Row = grd_LisDoc.Rows - 1
      
      grd_LisDoc.Col = 0:     grd_LisDoc.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
      grd_LisDoc.Col = 1:     grd_LisDoc.Text = ""
      grd_LisDoc.Col = 2:     grd_LisDoc.Text = "3"
      grd_LisDoc.Col = 3:     grd_LisDoc.Text = "012"
      grd_LisDoc.Col = 4:     grd_LisDoc.Text = "0"
      grd_LisDoc.Col = 5:     grd_LisDoc.Text = g_rst_Genera!PARPRD_CODITE
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_LisDoc.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   
   'Documentos del Inmueble
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '013' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   grd_LisDoc.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_LisDoc.Rows = grd_LisDoc.Rows + 1
      
      grd_LisDoc.Row = grd_LisDoc.Rows - 1
      
      grd_LisDoc.Col = 0:     grd_LisDoc.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
      grd_LisDoc.Col = 1:     grd_LisDoc.Text = ""
      grd_LisDoc.Col = 2:     grd_LisDoc.Text = "3"
      grd_LisDoc.Col = 3:     grd_LisDoc.Text = "013"
      grd_LisDoc.Col = 4:     grd_LisDoc.Text = "0"
      grd_LisDoc.Col = 5:     grd_LisDoc.Text = g_rst_Genera!PARPRD_CODITE
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_LisDoc.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   
   Call gs_UbiIniGrid(grd_LisDoc)
End Sub

Private Sub grd_GasAdm_SelChange()
   If grd_GasAdm.Rows > 2 Then
      grd_GasAdm.RowSel = grd_GasAdm.Row
   End If
End Sub

Private Sub grd_LisDoc_DblClick()
   If grd_LisDoc.Rows > 0 Then
      grd_LisDoc.Col = 1
      
      If grd_LisDoc.Text = "X" Then
         grd_LisDoc.Text = ""
      Else
         grd_LisDoc.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_LisDoc)

   End If
End Sub

Private Sub grd_LisDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_LisDoc_DblClick
   End If
End Sub

Private Sub grd_LisDoc_SelChange()
   If grd_LisDoc.Rows > 2 Then
      grd_LisDoc.RowSel = grd_LisDoc.Row
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodConHip)
   
   'Jefe de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(130)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Ventas
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(120)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director Comercial
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(100)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Operaciones
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(220)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director de Producción
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(200)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub


