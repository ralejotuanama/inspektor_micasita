VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_SegSol_13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   1425
   ClientTop       =   1920
   ClientWidth     =   12810
   Icon            =   "AteCli_frm_038.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7065
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   12462
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   1425
         Left            =   30
         TabIndex        =   7
         Top             =   4770
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   2514
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
         Begin Threed.SSPanel pnl_NumCar 
            Height          =   315
            Left            =   1620
            TabIndex        =   9
            Top             =   60
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   1050
            Width           =   2445
            _Version        =   65536
            _ExtentX        =   4313
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_FecEmi 
            Height          =   315
            Left            =   1620
            TabIndex        =   12
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
         Begin Threed.SSPanel pnl_FecVal 
            Height          =   315
            Left            =   1620
            TabIndex        =   14
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
         Begin Threed.SSPanel pnl_Import 
            Height          =   315
            Left            =   8820
            TabIndex        =   16
            Top             =   1020
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "20,000.00 "
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipMon 
            Height          =   315
            Left            =   8820
            TabIndex        =   18
            Top             =   690
            Width           =   3855
            _Version        =   65536
            _ExtentX        =   6800
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NUEVOS SOLES"
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
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   7470
            TabIndex        =   19
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label Label11 
            Caption         =   "Importe:"
            Height          =   315
            Left            =   7470
            TabIndex        =   17
            Top             =   1050
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "F. Emisión Valor:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label Label9 
            Caption         =   "F. Emisión Carta:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label8 
            Caption         =   "Número Operación:"
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Número Carta:"
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Width           =   4365
            _Version        =   65536
            _ExtentX        =   7699
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Seguimiento de Trámites COFIDE"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   5190
            TabIndex        =   6
            Top             =   120
            Width           =   7485
            _Version        =   65536
            _ExtentX        =   13203
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_038.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1425
         Left            =   30
         TabIndex        =   20
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   2514
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   21
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   8820
            TabIndex        =   22
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   1620
            TabIndex        =   23
            Top             =   1050
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   24
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   25
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_IniEva 
            Height          =   315
            Left            =   8820
            TabIndex        =   26
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_DiaTra 
            Height          =   315
            Left            =   11220
            TabIndex        =   27
            Top             =   720
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   8820
            TabIndex        =   28
            Top             =   60
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   3360
            TabIndex        =   29
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
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
         Begin Threed.SSPanel pnl_FinEva 
            Height          =   315
            Left            =   10020
            TabIndex        =   30
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_SitIns 
            Height          =   315
            Left            =   8820
            TabIndex        =   31
            Top             =   1050
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   40
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "días"
            Height          =   315
            Left            =   11700
            TabIndex        =   39
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "Período Eval.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   38
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   37
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   35
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   34
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Situac. Instanc.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   32
            Top             =   1050
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2505
         Left            =   30
         TabIndex        =   41
         Top             =   2220
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   4419
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
            Left            =   1620
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   1800
            Width           =   11025
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1095
            Left            =   30
            TabIndex        =   0
            Top             =   360
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   21
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
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
            Left            =   3120
            TabIndex        =   43
            Top             =   60
            Width           =   9285
            _Version        =   65536
            _ExtentX        =   16378
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
            Left            =   1590
            TabIndex        =   44
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
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
         Begin Threed.SSPanel pnl_MotRec 
            Height          =   315
            Left            =   1620
            TabIndex        =   45
            Top             =   1470
            Width           =   11025
            _Version        =   65536
            _ExtentX        =   19447
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin VB.Label Label18 
            Caption         =   "Motivo de Rechazo:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   1500
            Width           =   1485
         End
         Begin VB.Label Label19 
            Caption         =   "Observaciones de Rechazo:"
            Height          =   555
            Left            =   60
            TabIndex        =   46
            Top             =   1800
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   48
         Top             =   6240
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12030
            Picture         =   "AteCli_frm_038.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_13"
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

   Call fs_Inicia
   Call fs_Carga_DatGen
   
   'Cargar Seguimiento de Instancia (Información General)
   Call fs_Buscar_Seguim
   
   'Cargar Detalle de Seguimiento de la Instancia
   Call fs_Buscar_LisOcu
   
   'Buscar Información de la Evaluación
   Call fs_Buscar_InfCof
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Carga_DatGen()
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecIng.Caption = moddat_g_str_FecIng
   
   pnl_Situac.Caption = moddat_g_str_Situac
   
   Select Case moddat_g_int_Situac
      Case 1: pnl_Situac.ForeColor = modgen_g_con_ColAzu
      Case 2: pnl_Situac.ForeColor = modgen_g_con_ColVer
      Case 3: pnl_Situac.ForeColor = modgen_g_con_ColRoj
   End Select

End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1535
   grd_Listad.ColWidth(1) = 1535
   grd_Listad.ColWidth(2) = 9280
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar_LisOcu()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_Listad)
   pnl_MotRec.Caption = ""
   txt_ObsRec.Text = ""
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_TraCof) & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han registrado detalles del Seguimiento.", vbExclamation, modgen_g_con_PltPar
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Fecha de Ocurrencia
      grd_Listad.Col = 0
      grd_Listad.Text = r_str_FecOcu
      
      'Hora de Ocurrencia
      r_str_FecOcu = Format(g_rst_Princi!SEGHORCRE, "000000")
      r_str_FecOcu = Mid(r_str_FecOcu, 1, 2) & ":" & Mid(r_str_FecOcu, 3, 2) & ":" & Mid(r_str_FecOcu, 5, 2)
      
      grd_Listad.Col = 1
      grd_Listad.Text = r_str_FecOcu
      
      'Descripción Ocurrencia
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_Pardes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 13
            pnl_MotRec.Caption = moddat_gf_Consulta_Pardes("003", CStr(g_rst_Princi!SEGDET_MOTREC))
            txt_ObsRec.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_Buscar_Seguim()
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(modatecli_g_con_TraCof)
   
   If gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      g_rst_Princi.MoveFirst
      
      pnl_IniEva.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECINI))
      
      If g_rst_Princi!SEGUIM_FECFIN > 0 Then
         pnl_FinEva.Caption = gf_FormatoFecha(CStr(g_rst_Princi!SEGUIM_FECFIN))
         pnl_DiaTra.Caption = CStr(g_rst_Princi!SEGUIM_DIATRA) & " "
      End If
      
      pnl_SitIns.Caption = moddat_gf_Consulta_Pardes("023", CStr(g_rst_Princi!SEGUIM_SITUAC))
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Buscar_InfCof()
   Dim r_str_FecOcu  As String
   Dim r_str_NumCar  As String
   
   pnl_NumCar.Caption = ""
   pnl_FecEmi.Caption = ""
   pnl_FecVal.Caption = ""
   pnl_NumOpe.Caption = ""
   pnl_TipMon.Caption = ""
   pnl_Import.Caption = ""
   
   g_str_Parame = "SELECT * FROM TRA_DETCOF WHERE "
   g_str_Parame = g_str_Parame & "DETCOF_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If Len(Trim(g_rst_Princi!DETCOF_NUMCAR)) > 0 Then
      pnl_NumCar.Caption = Trim(g_rst_Princi!DETCOF_NUMCAR)
      pnl_NumOpe.Caption = Trim(g_rst_Princi!DETCOF_NUMOPE)
      pnl_Import.Caption = Format(g_rst_Princi!DETCOF_IMPORT, "###,###,#0.00") & " "
      
      g_str_Parame = "SELECT * FROM TRA_CARCOF WHERE "
      g_str_Parame = g_str_Parame & "CARCOF_NUMCAR = '" & pnl_NumCar.Caption & "'"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      pnl_FecEmi.Caption = Right(CStr(g_rst_Genera!CARCOF_FECEMI), 2) & "/" & Mid(CStr(g_rst_Genera!CARCOF_FECEMI), 5, 2) & "/" & Left(CStr(g_rst_Genera!CARCOF_FECEMI), 4)
      pnl_FecVal.Caption = Right(CStr(g_rst_Genera!CARCOF_FECVAL), 2) & "/" & Mid(CStr(g_rst_Genera!CARCOF_FECVAL), 5, 2) & "/" & Left(CStr(g_rst_Genera!CARCOF_FECVAL), 4)
      
      pnl_TipMon.Caption = moddat_gf_Consulta_Pardes("204", CStr(g_rst_Genera!CARCOF_TIPMON))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

