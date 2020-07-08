VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_ConCre_53 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10140
   ClientLeft      =   2715
   ClientTop       =   795
   ClientWidth     =   13260
   Icon            =   "AteCli_frm_176.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13245
      _Version        =   65536
      _ExtentX        =   23363
      _ExtentY        =   17912
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
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
            Left            =   12540
            Picture         =   "AteCli_frm_176.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerDet 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_176.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
            Left            =   720
            TabIndex        =   33
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
            Left            =   720
            TabIndex        =   34
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Detalle de Cuota"
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
            Picture         =   "AteCli_frm_176.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
            TabIndex        =   6
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
            TabIndex        =   7
            Top             =   390
            Width           =   11535
            _Version        =   65536
            _ExtentX        =   20346
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
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   1515
         Left            =   30
         TabIndex        =   10
         Top             =   2250
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
         _ExtentY        =   2672
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
         Begin MSFlexGridLib.MSFlexGrid grd_InfCuo 
            Height          =   1125
            Left            =   60
            TabIndex        =   11
            Top             =   330
            Width           =   13035
            _ExtentX        =   22992
            _ExtentY        =   1984
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
            Caption         =   "Información de la Cuota"
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
            TabIndex        =   12
            Top             =   60
            Width           =   4725
         End
      End
      Begin Threed.SSPanel SSPanel31 
         Height          =   1875
         Left            =   30
         TabIndex        =   13
         Top             =   8220
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
         _ExtentY        =   3307
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
            Height          =   825
            Left            =   60
            TabIndex        =   14
            Top             =   660
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   1455
            _Version        =   393216
            Rows            =   21
            Cols            =   8
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
            TabIndex        =   15
            Top             =   360
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro."
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
         Begin Threed.SSPanel SSPanel33 
            Height          =   285
            Left            =   570
            TabIndex        =   16
            Top             =   360
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Pago"
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
            Left            =   1740
            TabIndex        =   17
            Top             =   360
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Pago"
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
            Left            =   3600
            TabIndex        =   18
            Top             =   360
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Movim."
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
            Left            =   4770
            TabIndex        =   19
            Top             =   360
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Movim."
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
            Left            =   5940
            TabIndex        =   20
            Top             =   360
            Width           =   3285
            _Version        =   65536
            _ExtentX        =   5794
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Banco"
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
            Left            =   9210
            TabIndex        =   21
            Top             =   360
            Width           =   2475
            _Version        =   65536
            _ExtentX        =   4366
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Cuenta"
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
            Left            =   11640
            TabIndex        =   22
            Top             =   360
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
            Left            =   11640
            TabIndex        =   23
            Top             =   1500
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
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
         Begin VB.Label Label1 
            Caption         =   "Amortizaciones de la Cuota"
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
            TabIndex        =   25
            Top             =   60
            Width           =   4725
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Pagado ==> US$ "
            Height          =   315
            Left            =   9750
            TabIndex        =   24
            Top             =   1530
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4365
         Left            =   30
         TabIndex        =   26
         Top             =   3810
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
         _ExtentY        =   7699
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
         Begin MSFlexGridLib.MSFlexGrid grd_DesCuo 
            Height          =   3645
            Left            =   60
            TabIndex        =   27
            Top             =   660
            Width           =   13035
            _ExtentX        =   22992
            _ExtentY        =   6429
            _Version        =   393216
            Rows            =   10
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   360
            Width           =   6465
            _Version        =   65536
            _ExtentX        =   11404
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción"
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
            Left            =   6540
            TabIndex        =   29
            Top             =   360
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe Total"
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
            Left            =   8610
            TabIndex        =   30
            Top             =   360
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe Amortizado"
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
            Left            =   10680
            TabIndex        =   31
            Top             =   360
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe Pend. Pago"
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
         Begin VB.Label Label3 
            Caption         =   "Desglose de la Cuota"
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
            TabIndex        =   32
            Top             =   60
            Width           =   4725
         End
      End
   End
End
Attribute VB_Name = "frm_ConCre_53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerDet_Click()
   Call grd_Listad_DblClick
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   Call fs_Buscar_DatCuo
   Call fs_Buscar_PagCuo
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_InfCuo.ColWidth(0) = 2650
   grd_InfCuo.ColWidth(1) = 10000
   
   grd_InfCuo.ColAlignment(0) = flexAlignLeftCenter
   grd_InfCuo.ColAlignment(1) = flexAlignLeftCenter
   
   grd_DesCuo.ColWidth(0) = 6455
   grd_DesCuo.ColWidth(1) = 2075
   grd_DesCuo.ColWidth(2) = 2075
   grd_DesCuo.ColWidth(3) = 2075
   
   grd_DesCuo.ColAlignment(0) = flexAlignLeftCenter
   grd_DesCuo.ColAlignment(1) = flexAlignRightCenter
   grd_DesCuo.ColAlignment(2) = flexAlignRightCenter
   grd_DesCuo.ColAlignment(3) = flexAlignRightCenter
   
   grd_Listad.ColWidth(0) = 485
   grd_Listad.ColWidth(1) = 1175
   grd_Listad.ColWidth(2) = 1865
   grd_Listad.ColWidth(3) = 1175
   grd_Listad.ColWidth(4) = 1175
   grd_Listad.ColWidth(5) = 3275
   grd_Listad.ColWidth(6) = 2465
   grd_Listad.ColWidth(7) = 1145
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
End Sub

Private Sub fs_Buscar_PagCuo()
   Dim r_dbl_TCaDes     As Double
   Dim r_dbl_TCaViv     As Double
   Dim r_dbl_TCaOtr     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_TotCuo     As Double
   Dim r_dbl_TotPag     As Double
   Dim r_rst_GenAux     As ADODB.Recordset

   lbl_Totale.Caption = "Total ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
   Call gs_LimpiaGrid(grd_Listad)

   'Obteniendo Información de Pagos
   g_str_Parame = "SELECT * FROM CRE_HIPPAG WHERE "
   g_str_Parame = g_str_Parame & "HIPPAG_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPPAG_NUMCUO = " & CStr(moddat_g_int_NumCuo) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPPAG_NUMPAG DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_TotPag = 0
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Obteniendo Información del Movimiento de Pago
         g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
         g_str_Parame = g_str_Parame & "CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
         g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & CStr(g_rst_Princi!HIPPAG_FECMOV) & " AND "
         g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & CStr(g_rst_Princi!HIPPAG_NUMMOV) & " "
         g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_FECMOV DESC, CAJMOV_NUMMOV DESC"
   
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_GenAux, 3) Then
            Exit Sub
         End If
      
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = CStr(g_rst_Princi!HIPPAG_NUMPAG)
         
         grd_Listad.Col = 1
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPPAG_FECPAG))
         
         grd_Listad.Col = 2
         If r_rst_GenAux!CAJMOV_CODBAN = "000000" Then
            grd_Listad.Text = "EFECTIVO"
         Else
            grd_Listad.Text = "ABONO EN BANCO"
         End If
         
         grd_Listad.Col = 3
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPPAG_FECMOV))
      
         grd_Listad.Col = 4
         grd_Listad.Text = Format(g_rst_Princi!HIPPAG_NUMMOV, "00000")
      
         If r_rst_GenAux!CAJMOV_CODBAN <> "000000" Then
            grd_Listad.Col = 5
            grd_Listad.Text = moddat_gf_Consulta_ParDes("505", r_rst_GenAux!CAJMOV_CODBAN)
         
            grd_Listad.Col = 6
            grd_Listad.Text = Trim(r_rst_GenAux!CAJMOV_NUMCTA)
         End If
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(g_rst_Princi!HIPPAG_PAGMPR, "###,###,##0.00")
         
         r_dbl_TotPag = r_dbl_TotPag + g_rst_Princi!HIPPAG_PAGMPR
         
         r_rst_GenAux.Close
         Set r_rst_GenAux = Nothing
      
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      
      
      pnl_TotPag.Caption = Format(r_dbl_TotPag, "###,###,##0.00") & " "
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_DesCuo_SelChange()
   If grd_DesCuo.Rows > 2 Then
      grd_DesCuo.RowSel = grd_DesCuo.Row
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
  
   grd_Listad.Col = 3
   opecaj_g_str_FecMov = Format(CDate(grd_Listad.Text), "yyyymmdd")
  
   grd_Listad.Col = 4
   opecaj_g_str_NumMov = CStr(CLng(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_ConCre_54.Show 1
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar_DatCuo()
   Dim r_dbl_TotCuo     As Double
   Dim r_dbl_PenPag     As Double
   Dim r_dbl_PagCuo     As Double
   
   Call gs_LimpiaGrid(grd_InfCuo)
   Call gs_LimpiaGrid(grd_DesCuo)
   
   'Obteniendo Información de Cuota
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & CStr(moddat_g_int_NumCuo) & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_InfCuo.Rows = grd_InfCuo.Rows + 1
      grd_InfCuo.Row = grd_InfCuo.Rows - 1
      grd_InfCuo.Col = 0
      grd_InfCuo.Text = "Número de Cuota"
      
      grd_InfCuo.Col = 1
      grd_InfCuo.Text = CStr(moddat_g_int_NumCuo)
      
      grd_InfCuo.Rows = grd_InfCuo.Rows + 1
      grd_InfCuo.Row = grd_InfCuo.Rows - 1
      grd_InfCuo.Col = 0
      grd_InfCuo.Text = "Fecha de Vencimiento"
      
      grd_InfCuo.Col = 1
      grd_InfCuo.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
      
      
      'Si Situación es No-Pagado
      If g_rst_Princi!HIPCUO_SITUAC = 2 Then
         If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) < CDate(moddat_g_str_FecSis) Then
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Situación"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "VENCIDA"
            
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Días de Atraso"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = CStr(CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
         Else
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Situación"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "POR VENCER"
            
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Días de Atraso"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "0"
         End If
      Else
         grd_InfCuo.Rows = grd_InfCuo.Rows + 1
         grd_InfCuo.Row = grd_InfCuo.Rows - 1
         grd_InfCuo.Col = 0
         grd_InfCuo.Text = "Situación"
         
         grd_InfCuo.Col = 1
         grd_InfCuo.Text = "PAGADA"
      
         If CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))) > 0 Then
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Días de Atraso"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = CStr(CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
         Else
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Días de Atraso"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "0"
         End If
      End If
      
      'Capital
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = "Capital"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA - g_rst_Princi!HIPCUO_CAPPAG, 12, 2)
      
      'Interes
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = "Interés"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE - g_rst_Princi!HIPCUO_INTPAG, 12, 2)
      
      'Seguro Desgravamen
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = "Seguro Desgravamen"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_DESPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG - g_rst_Princi!HIPCUO_DESPAG, 12, 2)
      
      'Seguro Inmueble
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = "Seguro Inmueble"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG - g_rst_Princi!HIPCUO_VIVPAG, 12, 2)
      
      'Portes
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = "Portes"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColAzu
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG - g_rst_Princi!HIPCUO_OTRPAG, 12, 2)
      
      'Interes Moratorio
      grd_DesCuo.Rows = grd_DesCuo.Rows + 2
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = "Interés Moratorio"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTMOR, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_IMOPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTMOR - g_rst_Princi!HIPCUO_IMOPAG, 12, 2)
      
      'Interes Compensatorio
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = "Interés Compensatorio"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTCOM, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_ICOPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTCOM - g_rst_Princi!HIPCUO_ICOPAG, 12, 2)
         
      'Gastos de Cobranza
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = "Gastos de Cobranza"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_GASCOB, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_GCOPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_GASCOB - g_rst_Princi!HIPCUO_GCOPAG, 12, 2)
         
      'Otros Gastos
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = "Otros Gastos"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRGAS, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTGPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRGAS - g_rst_Princi!HIPCUO_OTGPAG, 12, 2)
         
      'Capital PBP
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = "Capital PBP"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPBBP, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_CBPPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPBBP - g_rst_Princi!HIPCUO_CBPPAG, 12, 2)
         
      'Interés PBP
      grd_DesCuo.Rows = grd_DesCuo.Rows + 1
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = "Interés PBP"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTBBP, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_IBPPAG, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColVer
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTBBP - g_rst_Princi!HIPCUO_IBPPAG, 12, 2)
         
      'Total
      r_dbl_TotCuo = 0
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_CAPITA
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_INTERE
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_DESORG
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_VIVORG
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_OTRORG
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_INTMOR
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_INTCOM
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_GASCOB
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_OTRGAS
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_CAPBBP
      r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Princi!HIPCUO_INTBBP
         
      
      'Pagado
      r_dbl_PagCuo = 0
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_CAPPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_INTPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_DESPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_VIVPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_OTRPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_ICOPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_IMOPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_GCOPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_OTGPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_CBPPAG
      r_dbl_PagCuo = r_dbl_PagCuo + g_rst_Princi!HIPCUO_IBPPAG
      
      'Pendiente de Pago
      r_dbl_PenPag = 0
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_CAPITA - g_rst_Princi!HIPCUO_CAPPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_INTERE - g_rst_Princi!HIPCUO_INTPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_DESORG - g_rst_Princi!HIPCUO_DESPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_VIVORG - g_rst_Princi!HIPCUO_VIVPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_OTRORG - g_rst_Princi!HIPCUO_OTRPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_INTMOR - g_rst_Princi!HIPCUO_IMOPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_INTCOM - g_rst_Princi!HIPCUO_ICOPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_GASCOB - g_rst_Princi!HIPCUO_GCOPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_OTRGAS - g_rst_Princi!HIPCUO_OTGPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_CAPBBP - g_rst_Princi!HIPCUO_CBPPAG
      r_dbl_PenPag = r_dbl_PenPag + g_rst_Princi!HIPCUO_INTBBP - g_rst_Princi!HIPCUO_IBPPAG
   
      'Totales
      grd_DesCuo.Rows = grd_DesCuo.Rows + 2
      grd_DesCuo.Row = grd_DesCuo.Rows - 1
      grd_DesCuo.Col = 0
      grd_DesCuo.CellForeColor = modgen_g_con_ColRoj
      grd_DesCuo.CellFontBold = True
      grd_DesCuo.Text = "Totales"
   
      grd_DesCuo.Col = 1
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColRoj
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_dbl_TotCuo, 12, 2)
      
      grd_DesCuo.Col = 2
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColRoj
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_dbl_PagCuo, 12, 2)
      
      grd_DesCuo.Col = 3
      grd_DesCuo.CellFontName = "Lucida Console"
      grd_DesCuo.CellFontSize = 8
      grd_DesCuo.CellForeColor = modgen_g_con_ColRoj
      grd_DesCuo.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_dbl_PenPag, 12, 2)
      
      Call gs_UbiIniGrid(grd_DesCuo)
      Call gs_UbiIniGrid(grd_InfCuo)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub




