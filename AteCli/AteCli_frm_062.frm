VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SegSol_20 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10590
   ClientLeft      =   2910
   ClientTop       =   960
   ClientWidth     =   11625
   Icon            =   "AteCli_frm_062.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18653
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
         Height          =   1755
         Left            =   30
         TabIndex        =   10
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
            TabIndex        =   24
            Top             =   1050
            Width           =   10035
         End
         Begin Threed.SSPanel pnl_SalIns 
            Height          =   315
            Left            =   1440
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   25
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
            TabIndex        =   48
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
         Begin VB.Label Label7 
            Caption         =   "F. Ingreso Inst.:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Observaciones de Rechazo:"
            Height          =   555
            Left            =   60
            TabIndex        =   27
            Top             =   1050
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "Motivo Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label24 
            Caption         =   "Situación Inst.:"
            Height          =   315
            Left            =   6420
            TabIndex        =   14
            Top             =   30
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "F. Salida Inst.:"
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   390
            Width           =   1005
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1935
         Left            =   30
         TabIndex        =   41
         Top             =   7770
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3413
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
         Begin VB.TextBox txt_ObsEva 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   675
            Left            =   60
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Text            =   "AteCli_frm_062.frx":000C
            Top             =   1200
            Width           =   11445
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   885
            Left            =   30
            TabIndex        =   45
            Top             =   300
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1561
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label6 
            Caption         =   "Resumen de Evaluación"
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
            TabIndex        =   44
            Top             =   60
            Width           =   2805
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   3
         Top             =   9750
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1440
            Top             =   270
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
         Begin VB.CommandButton cmd_Imprim 
            Enabled         =   0   'False
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_062.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Nueva Observación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_062.frx":0452
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   2850
            Top             =   150
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
            Left            =   2280
            Top             =   150
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
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
            Height          =   615
            Left            =   630
            TabIndex        =   5
            Top             =   30
            Width           =   6945
            _Version        =   65536
            _ExtentX        =   12250
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud - Evaluación Crediticia"
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
            Picture         =   "AteCli_frm_062.frx":0894
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1785
         Left            =   30
         TabIndex        =   6
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
            TabIndex        =   0
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
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   42
            Top             =   60
            Width           =   2805
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   46
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
            TabIndex        =   50
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
         Begin VB.Label Label12 
            Caption         =   "F. Rechazo:"
            Height          =   315
            Left            =   9270
            TabIndex        =   51
            Top             =   60
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label8 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   6420
            TabIndex        =   23
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Ingreso:"
            Height          =   315
            Left            =   6420
            TabIndex        =   22
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   2205
         Left            =   30
         TabIndex        =   28
         Top             =   5520
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3889
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
         Begin VB.TextBox txt_Observ 
            Height          =   705
            Left            =   5160
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Text            =   "AteCli_frm_062.frx":0B9E
            Top             =   750
            Width           =   6315
         End
         Begin VB.TextBox txt_Descar 
            Height          =   675
            Left            =   5160
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Text            =   "AteCli_frm_062.frx":0BA2
            Top             =   1470
            Width           =   6315
         End
         Begin VB.CommandButton cmd_NueObs 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_062.frx":0BA6
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Nueva Observación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_CanObs 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_062.frx":0FE8
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_GraObs 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_062.frx":12F2
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   390
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro."
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
            Left            =   960
            TabIndex        =   35
            Top             =   390
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Emisión"
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
            Left            =   2310
            TabIndex        =   36
            Top             =   390
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Descargo"
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisObs 
            Height          =   1485
            Left            =   60
            TabIndex        =   37
            Top             =   690
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   2619
            _Version        =   393216
            Rows            =   30
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label4 
            Caption         =   "Seguimiento de Observaciones"
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
            TabIndex        =   40
            Top             =   60
            Width           =   2805
         End
         Begin VB.Label Label1 
            Caption         =   "Observación:"
            Height          =   315
            Left            =   4080
            TabIndex        =   39
            Top             =   750
            Width           =   1035
         End
         Begin VB.Label Label11 
            Caption         =   "Descargo:"
            Height          =   315
            Left            =   4080
            TabIndex        =   38
            Top             =   1470
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgEdi        As Integer
Dim l_dbl_TasInt        As Double
Dim l_int_PlaAno_Cal    As Integer
Dim l_dbl_MtoPre_Cal    As Double
Dim l_int_FlgCon        As Integer
Dim l_str_EmpSeg        As String
Dim l_int_TipSeg_Cal    As Integer
Dim l_dbl_ComVta        As Double
Dim l_int_PerGra_Cal    As Integer

Private Sub cmd_CanObs_Click()
   l_int_FlgEdi = 1
   
   txt_Observ.Text = ""
   txt_Descar.Text = ""
   
   If grd_LisObs.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisObs)
      Call grd_LisObs_Click
   End If
   
   Call fs_Activa_Obs(True)
   Call gs_SetFocus(grd_LisObs)
End Sub

Private Sub cmd_GraObs_Click()
   Dim r_str_NumObs     As String
   Dim r_str_Descar     As String
   
   If Len(Trim(txt_Descar.Text)) = 0 Then
      MsgBox "Debe ingresar la Observación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descar)
      Exit Sub
   End If

   r_str_Descar = txt_Descar.Text

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_LisObs.Col = 0
   r_str_NumObs = grd_LisObs.Text
   
   Call gs_RefrescaGrid(grd_LisObs)
   
   If Not moddat_gf_Modifica_SegDet_Observ(moddat_g_str_NumSol, 21, 21, CStr(CInt(r_str_NumObs)), txt_Descar.Text, 2) Then
      Exit Sub
   End If
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, 21, 0, 9, 2) Then
      Exit Sub
   End If
   
   Call fs_Activa_Obs(True)
   Call fs_Buscar_LisObs
   
   Call fs_Buscar_LisOcu
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "DESCARGO DE OBSERVACION EN EVALUACION CREDITICIA (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & r_str_Descar
   
   Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
   MsgBox "Se genero el Descargo de la Observación a la Solicitud.", vbInformation, modgen_g_con_AteCli
   
   l_int_FlgEdi = 1
   
   moddat_g_int_FlgAct = 2
End Sub

Private Sub cmd_Imprim_Click()
   If modgen_g_int_TipUsu = 20900 Then
      MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir la Carta de Aprobación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CARAPR"
                        
   'Grabando en DAO
   moddat_g_str_CadDAO = "SELECT * FROM RPT_CARAPR WHERE CARAPR_NOMCLI = '" & moddat_g_str_NomCli & "'"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)

   moddat_g_rst_RecDAO.AddNew
                        
   moddat_g_rst_RecDAO("CARAPR_FECEMI") = "San Isidro, " & Format(Day(CDate(pnl_SalIns.Caption))) & " " & Left(moddat_gf_Consulta_ParDes("033", CStr(Month(CDate(pnl_SalIns.Caption)))), 1) & LCase(Mid(moddat_gf_Consulta_ParDes("033", CStr(Month(CDate(pnl_SalIns.Caption)))), 2)) & " del " & Format(Year(pnl_SalIns.Caption), "0000")
   moddat_g_rst_RecDAO("CARAPR_NOMCLI") = moddat_g_str_NomCli
   moddat_g_rst_RecDAO("CARAPR_DOCIDE") = moddat_gf_Consulta_ParDes("203", CStr(moddat_g_int_TipDoc)) & " Nro. " & moddat_g_str_NumDoc
   moddat_g_rst_RecDAO("CARAPR_TASINT") = CStr(l_dbl_TasInt) & " %"
   
   moddat_g_rst_RecDAO("CARAPR_IMPORT") = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(l_dbl_MtoPre_Cal, "###,###,##0.00") & " " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   moddat_g_rst_RecDAO("CARAPR_PLAZOC") = Format(l_int_PlaAno_Cal * 12, "000") & " meses"
   moddat_g_rst_RecDAO("CARAPR_PERGRA") = CStr(l_int_PerGra_Cal) & IIf(l_int_PerGra_Cal = 1, " mes", " meses")
   
   moddat_g_rst_RecDAO("CARAPR_PRODUC") = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   
   If l_int_FlgCon = 1 Then
      moddat_g_rst_RecDAO("CARAPR_OBSERV") = txt_ObsEva.Text
   Else
      moddat_g_rst_RecDAO("CARAPR_OBSERV") = " "
   End If
                        
   moddat_g_rst_RecDAO.Update
   moddat_g_rst_RecDAO.Close
   
   'If l_int_FlgCon = 2 Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_01.RPT"
   'Else
   '   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_CARAPR_02.RPT"
   'End If
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_NueObs_Click()
   Dim r_str_Observ  As String

   If modgen_g_int_TipUsu = 20900 Then
      MsgBox "No tiene acceso a esta opción.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If

   r_str_Observ = ""
   
   If grd_LisObs.Rows > 0 Then
      grd_LisObs.Row = 0
      
      grd_LisObs.Col = 3
      r_str_Observ = grd_LisObs.Text
      
      grd_LisObs.Col = 2
      
      If Len(Trim(grd_LisObs.Text)) > 0 Then
         Call gs_RefrescaGrid(grd_LisObs)
         
         MsgBox "No tiene observaciones pendientes de descargo.", vbExclamation, modgen_g_str_NomPlt
         
         Exit Sub
      End If
      Call gs_RefrescaGrid(grd_LisObs)
   Else
      MsgBox "No tiene observaciones registradas.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   l_int_FlgEdi = 2
   
   txt_Observ.Text = r_str_Observ
   txt_Descar.Text = ""
   Call fs_Activa_Obs(False)
   
   Call gs_SetFocus(txt_Descar)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   
   Call fs_Carga_DatGen
   
   Call fs_Activa_Obs(True)
   
   Call fs_Buscar_Seguim
   Call fs_Buscar_LisOcu
   Call fs_Buscar_LisObs
   Call fs_Buscar_EvaCre
   
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

   grd_LisObs.ColWidth(0) = 885
   grd_LisObs.ColWidth(1) = 1355
   grd_LisObs.ColWidth(2) = 1365
   grd_LisObs.ColWidth(3) = 0
   grd_LisObs.ColWidth(4) = 0
   
   grd_LisObs.ColAlignment(0) = flexAlignCenterCenter
   grd_LisObs.ColAlignment(1) = flexAlignCenterCenter
   grd_LisObs.ColAlignment(2) = flexAlignCenterCenter

   grd_LisEva.ColWidth(0) = 3300
   grd_LisEva.ColWidth(1) = 7940

   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Activa_Obs(ByVal p_Activa As Integer)
   cmd_NueObs.Enabled = p_Activa
   cmd_Imprim.Enabled = p_Activa
   grd_LisObs.Enabled = p_Activa
   
   cmd_GraObs.Enabled = Not p_Activa
   cmd_CanObs.Enabled = Not p_Activa
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
End Sub

Private Sub grd_LisObs_Click()
   If grd_LisObs.Rows > 0 Then
      grd_LisObs.Col = 3
      txt_Observ.Text = grd_LisObs.Text
      
      grd_LisObs.Col = 4
      txt_Descar.Text = grd_LisObs.Text
      
      Call gs_RefrescaGrid(grd_LisObs)
   End If
End Sub

Private Sub grd_LisObs_SelChange()
   If grd_LisObs.Rows > 2 Then
      grd_LisObs.RowSel = grd_LisObs.Row
   End If
   
   Call grd_LisObs_Click
End Sub


Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
End Sub

Private Sub txt_Descar_GotFocus()
   If l_int_FlgEdi = 2 Then
      Call gs_SelecTodo(txt_Descar)
   End If
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   If l_int_FlgEdi = 2 Then
      If KeyAscii = 13 Then
         Call gs_SetFocus(cmd_GraObs)
      Else
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsEva_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsRec_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Carga_DatGen()
   Dim r_int_CygTDo     As Integer
   Dim r_str_CygNDo     As String
   Dim r_int_EdaCli     As Integer

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
   
   l_dbl_TasInt = 0
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      g_rst_Princi.MoveFirst
      
      l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
      l_str_EmpSeg = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
      
      If moddat_g_int_TipMon = 1 Then
         l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
      Else
         l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
      End If
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Seguim()
   pnl_IngIns.Caption = ""
   pnl_SalIns.Caption = ""
   pnl_SitIns.Caption = ""

   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 21"
   
   If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      g_rst_Princi.MoveFirst
      
      pnl_IngIns.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
      If g_rst_Princi!SEGUIM_FECFIN > 0 Then
         pnl_SalIns.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
      End If
      
      moddat_g_int_SitIns = g_rst_Princi!SEGUIM_SITUAC
      pnl_SitIns.Caption = moddat_gf_Consulta_ParDes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_LisOcu()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   cmd_Imprim.Enabled = False
   
   pnl_MotRec.Caption = ""
   txt_ObsRec.Text = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21 "
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
         cmd_NueObs.Enabled = False
      ElseIf g_rst_Princi!SEGDET_CODOCU = 12 Then
         'Si la Solicitud está aprobada en la Instancia
         cmd_Imprim.Enabled = True
         cmd_NueObs.Enabled = False
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
   Call gs_SetFocus(grd_LisOcu)
End Sub

Private Sub fs_Buscar_LisObs()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisObs)
   
   txt_Observ.Text = ""
   txt_Descar.Text = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 21 AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisObs.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisObs.Rows = grd_LisObs.Rows + 1
      grd_LisObs.Row = grd_LisObs.Rows - 1
      
      'Número de Observación
      grd_LisObs.Col = 0
      grd_LisObs.Text = Format(g_rst_Princi!SEGDET_NUMOBS, "000")
      
      'Fecha de Observación
      r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      grd_LisObs.Col = 1
      grd_LisObs.Text = r_str_FecOcu
      
      'Fecha de Descargo
      If g_rst_Princi!SEGFECACT > 0 Then
         r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         
         grd_LisObs.Col = 2
         grd_LisObs.Text = r_str_FecOcu
      End If
      
      grd_LisObs.Col = 3
      grd_LisObs.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisObs.Col = 4
      grd_LisObs.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisObs.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   l_int_FlgEdi = 1
   
   Call gs_UbiIniGrid(grd_LisObs)
   Call grd_LisObs_Click
End Sub

Private Sub fs_Buscar_EvaCre()
   Dim r_dbl_Portes        As Double
   Dim r_int_TipVal_Viv    As Integer
   Dim r_dbl_Import_Viv    As Double
   Dim r_int_TipVal_Des    As Integer
   Dim r_dbl_Import_Des    As Double
   Dim r_dbl_SegViv        As Double
   Dim r_dbl_CuoMen        As Double
   Dim r_dbl_PorCon        As Double
   Dim r_dbl_TopCon        As Double
   Dim r_int_TipSeg        As Integer
   Dim r_int_EdaAct        As Integer
   Dim r_int_EdaCli        As Integer
   Dim r_int_EdaCyg        As Integer

   Call gs_LimpiaGrid(grd_LisEva)
   txt_ObsEva.Text = ""
   
   g_str_Parame = "SELECT * FROM TRA_EVACRE WHERE "
   g_str_Parame = g_str_Parame & "EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Aprobación Condicionada"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!EVACRE_FLGCON))
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.Text = "Monto Préstamo Aprobado"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_MTOPRE_CAL, 12, 2)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.Text = "Plazo Aprobado"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PLAANO_CAL) & " Años "
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.Text = "Período de Gracia Aprobado"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = CStr(g_rst_Princi!EVACRE_PERGRA_CAL) & IIf(g_rst_Princi!EVACRE_PERGRA_CAL = 1, " Mes", " Meses")
   
   If moddat_g_int_TipMon <> 1 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Tipo de Cambio de Aprobación"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_TIPCAM, 12, 4)
   End If
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.Text = "Tipo de Seguro Aprobado"
   
   grd_LisEva.Col = 1
   grd_LisEva.Text = moddat_gf_Consulta_TipSeg(l_str_EmpSeg, g_rst_Princi!EVACRE_TIPSEG_CAL)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 2
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Total Ingreso Líquido Neto S/."
   
   grd_LisEva.Col = 1
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Cuota Mensual Máxima S/."
   
   grd_LisEva.Col = 1
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_CUOSOL, 12, 2)
   
   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Cuota Mensual Máxima M. Prest. (" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ")"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_CUOMPR, 12, 2)
      
   If moddat_g_int_TipMon <> 1 Then
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.CellForeColor = modgen_g_con_ColRoj
      grd_LisEva.Text = "Tipo de Cambio (Cálculo Ingresos)"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.CellForeColor = modgen_g_con_ColRoj
      grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_TCAING, 12, 4)
   End If

   grd_LisEva.Rows = grd_LisEva.Rows + 1
   grd_LisEva.Row = grd_LisEva.Rows - 1
   grd_LisEva.Col = 0
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.Text = "Total Deuda (S/.)"
   
   grd_LisEva.Col = 1
   grd_LisEva.CellForeColor = modgen_g_con_ColRoj
   grd_LisEva.CellFontName = "Lucida Console"
   grd_LisEva.CellFontSize = 8
   grd_LisEva.Text = gf_FormatoNumero(g_rst_Princi!EVACRE_MTODEU, 12, 2)

   txt_ObsEva.Text = Trim(g_rst_Princi!EVACRE_OBSEVA & "")
   
   l_int_FlgCon = g_rst_Princi!EVACRE_FLGCON
   l_int_PlaAno_Cal = g_rst_Princi!EVACRE_PLAANO_CAL
   l_dbl_MtoPre_Cal = g_rst_Princi!EVACRE_MTOPRE_CAL
   l_int_PerGra_Cal = g_rst_Princi!EVACRE_PERGRA_CAL

   Call gs_UbiIniGrid(grd_LisEva)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodConHip)
   
   'Evaluador de Creditos
   r_str_Cadena = moddat_gf_UsuObs(moddat_g_str_NumSol, 21)
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_UsuSis(r_str_Cadena)
   
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
   
   'Jefe de Créditos
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(210)
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

