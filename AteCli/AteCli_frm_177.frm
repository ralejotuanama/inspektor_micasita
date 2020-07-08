VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_ConCre_54 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   1260
   ClientTop       =   1230
   ClientWidth     =   15075
   Icon            =   "AteCli_frm_177.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   15425
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
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   14985
         _Version        =   65536
         _ExtentX        =   26432
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
            Left            =   14370
            Picture         =   "AteCli_frm_177.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   14985
         _Version        =   65536
         _ExtentX        =   26432
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
            Height          =   315
            Left            =   690
            TabIndex        =   28
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   690
            TabIndex        =   29
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Detalle de Pago"
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
            Picture         =   "AteCli_frm_177.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   4
         Top             =   1440
         Width           =   14985
         _Version        =   65536
         _ExtentX        =   26432
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   390
            Width           =   13395
            _Version        =   65536
            _ExtentX        =   23627
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   4125
         Left            =   30
         TabIndex        =   9
         Top             =   2250
         Width           =   14985
         _Version        =   65536
         _ExtentX        =   26432
         _ExtentY        =   7276
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
         Begin MSFlexGridLib.MSFlexGrid grd_InfPag 
            Height          =   3735
            Left            =   60
            TabIndex        =   10
            Top             =   330
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   6588
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
            Caption         =   "Información del Pago"
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
            TabIndex        =   11
            Top             =   60
            Width           =   2025
         End
      End
      Begin Threed.SSPanel SSPanel31 
         Height          =   2265
         Left            =   30
         TabIndex        =   12
         Top             =   6420
         Width           =   14985
         _Version        =   65536
         _ExtentX        =   26432
         _ExtentY        =   3995
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
            Height          =   1245
            Left            =   60
            TabIndex        =   13
            Top             =   630
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   2196
            _Version        =   393216
            Rows            =   21
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel32 
            Height          =   285
            Left            =   90
            TabIndex        =   14
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
         Begin Threed.SSPanel SSPanel34 
            Height          =   285
            Left            =   840
            TabIndex        =   15
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital"
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
         Begin Threed.SSPanel SSPanel35 
            Height          =   285
            Left            =   1980
            TabIndex        =   16
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés"
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
         Begin Threed.SSPanel SSPanel36 
            Height          =   285
            Left            =   3120
            TabIndex        =   17
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Desg."
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
         Begin Threed.SSPanel SSPanel37 
            Height          =   285
            Left            =   4260
            TabIndex        =   18
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Inmueb."
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
         Begin Threed.SSPanel SSPanel38 
            Height          =   285
            Left            =   5400
            TabIndex        =   19
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Portes"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   6540
            TabIndex        =   20
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. Comp."
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
         Begin Threed.SSPanel SSPanel40 
            Height          =   285
            Left            =   7680
            TabIndex        =   21
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. Morat."
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
         Begin Threed.SSPanel pnl_TotPag 
            Height          =   315
            Left            =   13380
            TabIndex        =   22
            Top             =   1890
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   8820
            TabIndex        =   23
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Gastos Cobr."
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
            Left            =   9960
            TabIndex        =   24
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Otros Gastos"
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   285
            Left            =   13380
            TabIndex        =   25
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total"
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
            Left            =   11100
            TabIndex        =   30
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital PBP"
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
            Left            =   12240
            TabIndex        =   31
            Top             =   330
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Interés PBP"
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
         Begin VB.Label Label1 
            Caption         =   "Desglose del Pago"
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
            TabIndex        =   27
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Pagado ==> US$"
            Height          =   315
            Left            =   11580
            TabIndex        =   26
            Top             =   1890
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_ConCre_54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Call fs_Buscar_InfMov
   Call fs_Buscar_DesPag
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_InfPag.ColWidth(0) = 2650
   grd_InfPag.ColWidth(1) = 10000
   
   grd_InfPag.ColAlignment(0) = flexAlignLeftCenter
   grd_InfPag.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad.ColWidth(0) = 755
   grd_Listad.ColWidth(1) = 1125
   grd_Listad.ColWidth(2) = 1125
   grd_Listad.ColWidth(3) = 1125
   grd_Listad.ColWidth(4) = 1125
   grd_Listad.ColWidth(5) = 1125
   grd_Listad.ColWidth(6) = 1125
   grd_Listad.ColWidth(7) = 1125
   grd_Listad.ColWidth(8) = 1125
   grd_Listad.ColWidth(9) = 1125
   grd_Listad.ColWidth(10) = 1125
   grd_Listad.ColWidth(11) = 1125
   grd_Listad.ColWidth(12) = 1125
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignRightCenter
   grd_Listad.ColAlignment(12) = flexAlignRightCenter
End Sub

Private Sub fs_Buscar_DesPag()
   Dim r_dbl_TotCuo     As Double
   Dim r_dbl_TotPag     As Double

   Call gs_LimpiaGrid(grd_Listad)
   lbl_Totale.Caption = "Total ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "

   'Obteniendo Información de Pagos
   g_str_Parame = "SELECT * FROM CRE_HIPPAG WHERE "
   g_str_Parame = g_str_Parame & "HIPPAG_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPPAG_FECMOV = " & opecaj_g_str_FecMov & " AND "
   g_str_Parame = g_str_Parame & "HIPPAG_NUMMOV = " & opecaj_g_str_NumMov & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPPAG_NUMCUO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_TotPag = 0
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         r_dbl_TotCuo = 0
         
         grd_Listad.Col = 0
         grd_Listad.Text = CStr(g_rst_Princi!HIPPAG_NUMCUO)
         
         grd_Listad.Col = 1
         If g_rst_Princi!HIPPAG_CAPITA > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_CAPITA, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 2
         If g_rst_Princi!HIPPAG_INTERE > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_INTERE, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 3
         If g_rst_Princi!HIPPAG_DESORG > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_DESORG, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 4
         If g_rst_Princi!HIPPAG_VIVORG > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_VIVORG, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 5
         If g_rst_Princi!HIPPAG_OTRORG > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_OTRORG, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 6
         If g_rst_Princi!HIPPAG_INTCOM > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_INTCOM, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 7
         If g_rst_Princi!HIPPAG_INTMOR > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_INTMOR, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 8
         If g_rst_Princi!HIPPAG_GASCOB > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_GASCOB, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 9
         If g_rst_Princi!HIPPAG_OTRGAS > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_OTRGAS, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 10
         If g_rst_Princi!HIPPAG_CAPBBP > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_CAPBBP, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 11
         If g_rst_Princi!HIPPAG_INTBBP > 0 Then
            grd_Listad.Text = Format(g_rst_Princi!HIPPAG_INTBBP, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_Listad.Text)
         Else
            grd_Listad.Text = "-"
         End If
         
         grd_Listad.Col = 12
         grd_Listad.Text = Format(r_dbl_TotCuo, "###,###,##0.00")
         
         r_dbl_TotPag = r_dbl_TotPag + r_dbl_TotCuo
      
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      
      pnl_TotPag.Caption = Format(r_dbl_TotPag, "###,###,##0.00") & " "
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_InfPag_SelChange()
   If grd_InfPag.Rows > 2 Then
      grd_InfPag.RowSel = grd_InfPag.Row
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar_InfMov()
   Call gs_LimpiaGrid(grd_InfPag)
   
   'Obteniendo Información del Movimiento de Pago
   g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & opecaj_g_str_FecMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & opecaj_g_str_NumMov & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_InfPag.Rows = grd_InfPag.Rows + 1
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "Fecha de Pago"
   
   grd_InfPag.Col = 1
   If g_rst_Princi!CAJMOV_FECDEP > 0 Then
      grd_InfPag.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECDEP))
   Else
      grd_InfPag.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
   End If
   
   grd_InfPag.Rows = grd_InfPag.Rows + 1
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "Fecha de Movimiento (Registro)"
   
   grd_InfPag.Col = 1
   grd_InfPag.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
   
   grd_InfPag.Rows = grd_InfPag.Rows + 1
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "Número de Movimiento"
   
   grd_InfPag.Col = 1
   grd_InfPag.Text = Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
   
   grd_InfPag.Rows = grd_InfPag.Rows + 1
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "Forma de Pago"
   
   grd_InfPag.Col = 1
   If g_rst_Princi!CAJMOV_CODBAN = "000000" Then
      grd_InfPag.Text = "EFECTIVO"
   Else
      grd_InfPag.Text = "ABONO EN BANCO"
      
      grd_InfPag.Rows = grd_InfPag.Rows + 2
      grd_InfPag.Row = grd_InfPag.Rows - 1
      grd_InfPag.Col = 0
      grd_InfPag.Text = "Banco"
      
      grd_InfPag.Col = 1
      grd_InfPag.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!CAJMOV_CODBAN)
   
      grd_InfPag.Rows = grd_InfPag.Rows + 1
      grd_InfPag.Row = grd_InfPag.Rows - 1
      grd_InfPag.Col = 0
      grd_InfPag.Text = "Número de Cuenta"
      
      grd_InfPag.Col = 1
      grd_InfPag.Text = Trim(g_rst_Princi!CAJMOV_NUMCTA)
   
      grd_InfPag.Rows = grd_InfPag.Rows + 1
      grd_InfPag.Row = grd_InfPag.Rows - 1
      grd_InfPag.Col = 0
      grd_InfPag.Text = "Número de Comprobante"
      
      grd_InfPag.Col = 1
      grd_InfPag.Text = Trim(g_rst_Princi!CAJMOV_NUMCOM)
   
      grd_InfPag.Rows = grd_InfPag.Rows + 1
      grd_InfPag.Row = grd_InfPag.Rows - 1
      grd_InfPag.Col = 0
      grd_InfPag.Text = "Tipo de Registro"
   
      grd_InfPag.Col = 1
      grd_InfPag.Text = moddat_gf_Consulta_ParDes("239", CStr(g_rst_Princi!CAJMOV_TIPREG))
      
      If g_rst_Princi!CAJMOV_TIPREG = 2 Then
         grd_InfPag.Rows = grd_InfPag.Rows + 1
         grd_InfPag.Row = grd_InfPag.Rows - 1
         grd_InfPag.Col = 0
         grd_InfPag.Text = "Fecha de Recaudo"
         
         grd_InfPag.Col = 1
         grd_InfPag.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECREC))
         
         grd_InfPag.Rows = grd_InfPag.Rows + 1
         grd_InfPag.Row = grd_InfPag.Rows - 1
         grd_InfPag.Col = 0
         grd_InfPag.Text = "Oficina del Banco"
         
         grd_InfPag.Col = 1
         grd_InfPag.Text = Trim(g_rst_Princi!CAJMOV_OFIPAG & "")
         
         grd_InfPag.Rows = grd_InfPag.Rows + 1
         grd_InfPag.Row = grd_InfPag.Rows - 1
         grd_InfPag.Col = 0
         grd_InfPag.Text = "Forma de Pago en Banco"
         
         grd_InfPag.Col = 1
         grd_InfPag.Text = Trim(g_rst_Princi!CAJMOV_FORPAG & "")
         
         grd_InfPag.Rows = grd_InfPag.Rows + 1
         grd_InfPag.Row = grd_InfPag.Rows - 1
         grd_InfPag.Col = 0
         grd_InfPag.Text = "Canal de Pago en Banco"
         
         grd_InfPag.Col = 1
         grd_InfPag.Text = Trim(g_rst_Princi!CAJMOV_CANPAG & "")
      End If
   End If
   
   grd_InfPag.Rows = grd_InfPag.Rows + 2
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "Importe Pagado"
   
   grd_InfPag.Col = 1
   grd_InfPag.CellFontName = "Lucida Console"
   grd_InfPag.CellFontSize = 8
   grd_InfPag.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!CAJMOV_IMPTOT, 12, 2)
   
   grd_InfPag.Rows = grd_InfPag.Rows + 1
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "ITF"
   
   grd_InfPag.Col = 1
   grd_InfPag.CellFontName = "Lucida Console"
   grd_InfPag.CellFontSize = 8
   grd_InfPag.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!CAJMOV_ITFIMP, 12, 2)
   
   grd_InfPag.Rows = grd_InfPag.Rows + 1
   grd_InfPag.Row = grd_InfPag.Rows - 1
   grd_InfPag.Col = 0
   grd_InfPag.Text = "Importe Neto de Pago"
   
   grd_InfPag.Col = 1
   grd_InfPag.CellFontName = "Lucida Console"
   grd_InfPag.CellFontSize = 8
   grd_InfPag.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!CAJMOV_IMPPAG, 12, 2)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_InfPag)
End Sub




