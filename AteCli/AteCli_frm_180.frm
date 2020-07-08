VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_ConCre_57 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   1650
   ClientTop       =   675
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_180.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10035
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   17701
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
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Left            =   10920
            Picture         =   "AteCli_frm_180.frx":000C
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
            Height          =   315
            Left            =   690
            TabIndex        =   103
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
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   690
            TabIndex        =   104
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cronogramas de Pago"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "AteCli_frm_180.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   4
         Top             =   1440
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   390
            Width           =   9915
            _Version        =   65536
            _ExtentX        =   17489
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
      Begin Threed.SSPanel SSPanel22 
         Height          =   7725
         Left            =   30
         TabIndex        =   9
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   13626
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
         Begin TabDlg.SSTab tab_Cronog 
            Height          =   7605
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   13414
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Cliente - Tramo No Concesional"
            TabPicture(0)   =   "AteCli_frm_180.frx":0758
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_Totale(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel63"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel60"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel58"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel43"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel42"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel41"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel38"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel5"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel4"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "grd_CliNCo_Listad"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_CliNCo_Intere"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_CliNCo_SegPre"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_CliNCo_SegViv"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_CliNCo_Capita"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).ControlCount=   17
            TabCaption(1)   =   "Cliente - Tramo Concesional"
            TabPicture(1)   =   "AteCli_frm_180.frx":0774
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_Totale(1)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "pnl_CliCon_Capita"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "pnl_CliCon_Intere"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "SSPanel21"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "SSPanel13"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "SSPanel12"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "SSPanel11"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "SSPanel10"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "grd_CliCon_Listad"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "SSPanel9"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).Control(10)=   "pnl_CliCon_TotCuo"
            Tab(1).Control(10).Enabled=   0   'False
            Tab(1).ControlCount=   11
            TabCaption(2)   =   "Mivivienda - Tramo No Concesional"
            TabPicture(2)   =   "AteCli_frm_180.frx":0790
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lbl_Totale(2)"
            Tab(2).Control(1)=   "pnl_MViNCo_Comisi"
            Tab(2).Control(2)=   "SSPanel3"
            Tab(2).Control(3)=   "SSPanel62"
            Tab(2).Control(4)=   "SSPanel61"
            Tab(2).Control(5)=   "SSPanel59"
            Tab(2).Control(6)=   "SSPanel36"
            Tab(2).Control(7)=   "SSPanel35"
            Tab(2).Control(8)=   "SSPanel34"
            Tab(2).Control(9)=   "SSPanel33"
            Tab(2).Control(10)=   "SSPanel2"
            Tab(2).Control(11)=   "SSPanel30"
            Tab(2).Control(12)=   "pnl_MViNCo_TotCuo"
            Tab(2).Control(13)=   "pnl_MViNCo_OtrCar"
            Tab(2).Control(14)=   "grd_MViNCo_Listad"
            Tab(2).Control(15)=   "pnl_MViNCo_Intere"
            Tab(2).Control(16)=   "pnl_MViNCo_SegPre"
            Tab(2).Control(17)=   "pnl_MViNCo_SegViv"
            Tab(2).Control(18)=   "pnl_MViNCo_Capita"
            Tab(2).ControlCount=   19
            TabCaption(3)   =   "Mivivienda - Tramo Concesional"
            TabPicture(3)   =   "AteCli_frm_180.frx":07AC
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lbl_Totale(3)"
            Tab(3).Control(1)=   "pnl_MViCon_Comisi"
            Tab(3).Control(2)=   "pnl_MViCon_Capita"
            Tab(3).Control(3)=   "pnl_MViCon_Intere"
            Tab(3).Control(4)=   "SSPanel20"
            Tab(3).Control(5)=   "SSPanel19"
            Tab(3).Control(6)=   "SSPanel18"
            Tab(3).Control(7)=   "SSPanel17"
            Tab(3).Control(8)=   "SSPanel16"
            Tab(3).Control(9)=   "SSPanel15"
            Tab(3).Control(10)=   "grd_MviCon_Listad"
            Tab(3).Control(11)=   "SSPanel14"
            Tab(3).Control(12)=   "pnl_MViCon_TotCuo"
            Tab(3).ControlCount=   13
            TabCaption(4)   =   "Cofide"
            TabPicture(4)   =   "AteCli_frm_180.frx":07C8
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "lbl_Totale(4)"
            Tab(4).Control(1)=   "pnl_CofNCo_Comisi"
            Tab(4).Control(2)=   "pnl_CofNCo_Capita"
            Tab(4).Control(3)=   "pnl_CofNCo_Intere"
            Tab(4).Control(4)=   "SSPanel55"
            Tab(4).Control(5)=   "SSPanel54"
            Tab(4).Control(6)=   "SSPanel49"
            Tab(4).Control(7)=   "SSPanel47"
            Tab(4).Control(8)=   "SSPanel46"
            Tab(4).Control(9)=   "SSPanel45"
            Tab(4).Control(10)=   "grd_CofNCo_Listad"
            Tab(4).Control(11)=   "SSPanel44"
            Tab(4).Control(12)=   "pnl_CofNCo_TotCuo"
            Tab(4).ControlCount=   13
            Begin Threed.SSPanel pnl_MViCon_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   11
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2280
               TabIndex        =   12
               Top             =   7230
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_SegViv 
               Height          =   285
               Left            =   5790
               TabIndex        =   13
               Top             =   7230
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_SegPre 
               Height          =   285
               Left            =   4620
               TabIndex        =   14
               Top             =   7230
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_Intere 
               Height          =   285
               Left            =   3450
               TabIndex        =   15
               Top             =   7230
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
               Height          =   6525
               Left            =   30
               TabIndex        =   16
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11509
               _Version        =   393216
               Rows            =   21
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -67530
               TabIndex        =   17
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -65190
               TabIndex        =   18
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   19
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel27 
               Height          =   285
               Left            =   -73770
               TabIndex        =   20
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel28 
               Height          =   285
               Left            =   -71970
               TabIndex        =   21
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel29 
               Height          =   285
               Left            =   -70140
               TabIndex        =   22
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel31 
               Height          =   285
               Left            =   -66480
               TabIndex        =   23
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel32 
               Height          =   285
               Left            =   -64650
               TabIndex        =   24
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   25
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
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
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   26
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel48 
               Height          =   285
               Left            =   -73770
               TabIndex        =   27
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel50 
               Height          =   285
               Left            =   -71970
               TabIndex        =   28
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel51 
               Height          =   285
               Left            =   -70140
               TabIndex        =   29
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel52 
               Height          =   285
               Left            =   -66480
               TabIndex        =   30
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel53 
               Height          =   285
               Left            =   -64650
               TabIndex        =   31
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   32
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   6960
               TabIndex        =   33
               Top             =   7230
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_TotCuo 
               Height          =   285
               Left            =   8130
               TabIndex        =   34
               Top             =   7230
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -70950
               TabIndex        =   35
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
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
            Begin MSFlexGridLib.MSFlexGrid grd_MviCon_Listad 
               Height          =   6525
               Left            =   -74970
               TabIndex        =   36
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11509
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -74940
               TabIndex        =   37
               Top             =   360
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -74190
               TabIndex        =   38
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -72690
               TabIndex        =   39
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -67470
               TabIndex        =   40
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
               Left            =   -65730
               TabIndex        =   41
               Top             =   360
               Width           =   1710
               _Version        =   65536
               _ExtentX        =   3016
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
               Left            =   -69210
               TabIndex        =   42
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
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
            Begin Threed.SSPanel pnl_MViCon_Intere 
               Height          =   285
               Left            =   -70950
               TabIndex        =   43
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   44
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   45
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -68370
               TabIndex        =   46
               Top             =   7230
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -70530
               TabIndex        =   47
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interes"
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   6525
               Left            =   -74970
               TabIndex        =   48
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11509
               _Version        =   393216
               Rows            =   21
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -74940
               TabIndex        =   49
               Top             =   360
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74190
               TabIndex        =   50
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -72690
               TabIndex        =   51
               Top             =   360
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   -68370
               TabIndex        =   52
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -66210
               TabIndex        =   53
               Top             =   360
               Width           =   2205
               _Version        =   65536
               _ExtentX        =   3889
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -70530
               TabIndex        =   54
               Top             =   7230
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   55
               Top             =   7230
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_Capita 
               Height          =   285
               Left            =   -72840
               TabIndex        =   56
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_SegViv 
               Height          =   285
               Left            =   -69630
               TabIndex        =   57
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_SegPre 
               Height          =   285
               Left            =   -70710
               TabIndex        =   58
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_Intere 
               Height          =   285
               Left            =   -71790
               TabIndex        =   59
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin MSFlexGridLib.MSFlexGrid grd_MViNCo_Listad 
               Height          =   6525
               Left            =   -74970
               TabIndex        =   60
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11509
               _Version        =   393216
               Rows            =   21
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_MViNCo_OtrCar 
               Height          =   285
               Left            =   -68550
               TabIndex        =   61
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_TotCuo 
               Height          =   285
               Left            =   -66390
               TabIndex        =   62
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   3450
               TabIndex        =   63
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   60
               TabIndex        =   64
               Top             =   360
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   840
               TabIndex        =   65
               Top             =   360
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   2280
               TabIndex        =   66
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   8130
               TabIndex        =   67
               Top             =   360
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   9420
               TabIndex        =   68
               Top             =   360
               Width           =   1560
               _Version        =   65536
               _ExtentX        =   2752
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   4620
               TabIndex        =   69
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   5790
               TabIndex        =   70
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   6960
               TabIndex        =   71
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   -71790
               TabIndex        =   72
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
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
            Begin Threed.SSPanel SSPanel2 
               Height          =   285
               Left            =   -74940
               TabIndex        =   73
               Top             =   360
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
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
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   -74250
               TabIndex        =   74
               Top             =   360
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
               Left            =   -72840
               TabIndex        =   75
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
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
               Left            =   -66390
               TabIndex        =   76
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
               Left            =   -65310
               TabIndex        =   77
               Top             =   360
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   -70710
               TabIndex        =   78
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   -69630
               TabIndex        =   79
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   -68550
               TabIndex        =   80
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   285
               Left            =   -67470
               TabIndex        =   81
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "C. COFIDE"
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
            Begin Threed.SSPanel pnl_MViNCo_Comisi 
               Height          =   285
               Left            =   -67470
               TabIndex        =   82
               Top             =   7230
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   83
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -70950
               TabIndex        =   84
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
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
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   6525
               Left            =   -74970
               TabIndex        =   85
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   11509
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -74940
               TabIndex        =   86
               Top             =   360
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
            Begin Threed.SSPanel SSPanel46 
               Height          =   285
               Left            =   -74190
               TabIndex        =   87
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel47 
               Height          =   285
               Left            =   -72690
               TabIndex        =   88
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
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
            Begin Threed.SSPanel SSPanel49 
               Height          =   285
               Left            =   -67470
               TabIndex        =   89
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel54 
               Height          =   285
               Left            =   -65730
               TabIndex        =   90
               Top             =   360
               Width           =   1710
               _Version        =   65536
               _ExtentX        =   3016
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel55 
               Height          =   285
               Left            =   -69210
               TabIndex        =   91
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
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
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70950
               TabIndex        =   92
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   93
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   94
               Top             =   7230
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   3
               Left            =   -74610
               TabIndex        =   102
               Top             =   7230
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   2
               Left            =   -74790
               TabIndex        =   101
               Top             =   7230
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   1
               Left            =   -74610
               TabIndex        =   100
               Top             =   7230
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   0
               Left            =   390
               TabIndex        =   99
               Top             =   7230
               Width           =   1845
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   98
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   97
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label4 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   96
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   4
               Left            =   -74610
               TabIndex        =   95
               Top             =   7230
               Width           =   1845
            End
         End
      End
   End
End
Attribute VB_Name = "frm_ConCre_57"
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
   
   Call fs_Carga_Cro_CliNCo
   
   If moddat_g_str_CodPrd = "002" Then
      tab_Cronog.TabCaption(0) = "Cliente"
      
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = False
      tab_Cronog.TabVisible(3) = False
      tab_Cronog.TabVisible(4) = False
   Else
      tab_Cronog.TabCaption(0) = "Cliente No Concesional"
      tab_Cronog.TabCaption(1) = "Cliente Concesional"
      
      If moddat_g_str_CodPrd = "001" Then
         tab_Cronog.TabCaption(2) = "Mivivienda No Concesional"
         tab_Cronog.TabCaption(3) = "Mivivienda Concesional"
         
         tab_Cronog.TabVisible(4) = False
      ElseIf moddat_g_str_CodPrd = "003" Then
         tab_Cronog.TabCaption(2) = "Mivivienda No Concesional"
         tab_Cronog.TabCaption(3) = "Mivivienda Concesional"
         tab_Cronog.TabCaption(4) = "Cofide"
         
         tab_Cronog.TabVisible(4) = True
      ElseIf moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Then
         tab_Cronog.TabCaption(2) = "Cofide No Concesional"
         tab_Cronog.TabCaption(3) = "Cofide Concesional"
         
         tab_Cronog.TabVisible(4) = False
      End If
      
      Call fs_Carga_Cro_CliCon
      Call fs_Carga_Cro_MViNCo
      Call fs_Carga_Cro_MViCon
      
      If moddat_g_str_CodPrd = "003" Then
         Call fs_Carga_Cro_CofNCo
      End If
   End If
   
   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 795
   grd_CliNCo_Listad.ColWidth(1) = 1425
   grd_CliNCo_Listad.ColWidth(2) = 1180
   grd_CliNCo_Listad.ColWidth(3) = 1170
   grd_CliNCo_Listad.ColWidth(4) = 1160
   grd_CliNCo_Listad.ColWidth(5) = 1160
   grd_CliNCo_Listad.ColWidth(6) = 1160
   grd_CliNCo_Listad.ColWidth(7) = 1320
   grd_CliNCo_Listad.ColWidth(8) = 1560
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter

   'Mivivienda No Concesional
   grd_MViNCo_Listad.ColWidth(0) = 695
   grd_MViNCo_Listad.ColWidth(1) = 1415
   grd_MViNCo_Listad.ColWidth(2) = 1070
   grd_MViNCo_Listad.ColWidth(3) = 1070
   grd_MViNCo_Listad.ColWidth(4) = 1080
   grd_MViNCo_Listad.ColWidth(5) = 1080
   grd_MViNCo_Listad.ColWidth(6) = 1080
   grd_MViNCo_Listad.ColWidth(7) = 1080
   grd_MViNCo_Listad.ColWidth(8) = 1080
   grd_MViNCo_Listad.ColWidth(9) = 1290
   
   grd_MViNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(9) = flexAlignRightCenter

   'Mivivienda Concesional
   grd_MviCon_Listad.ColWidth(0) = 770
   grd_MviCon_Listad.ColWidth(1) = 1485
   grd_MviCon_Listad.ColWidth(2) = 1730
   grd_MviCon_Listad.ColWidth(3) = 1740
   grd_MviCon_Listad.ColWidth(4) = 1740
   grd_MviCon_Listad.ColWidth(5) = 1740
   grd_MviCon_Listad.ColWidth(6) = 1730
   
   grd_MviCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(6) = flexAlignRightCenter
   

   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 770
   grd_CliCon_Listad.ColWidth(1) = 1485
   grd_CliCon_Listad.ColWidth(2) = 2170
   grd_CliCon_Listad.ColWidth(3) = 2160
   grd_CliCon_Listad.ColWidth(4) = 2170
   grd_CliCon_Listad.ColWidth(5) = 2160
   
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter
   
   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 770
   grd_CofNCo_Listad.ColWidth(1) = 1485
   grd_CofNCo_Listad.ColWidth(2) = 1730
   grd_CofNCo_Listad.ColWidth(3) = 1740
   grd_CofNCo_Listad.ColWidth(4) = 1740
   grd_CofNCo_Listad.ColWidth(5) = 1740
   grd_CofNCo_Listad.ColWidth(6) = 1730
   
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Carga_Cro_CliNCo()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_CliNCo_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CliNCo_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
         grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
         
         r_dbl_ImpCuo = 0
         
         grd_CliNCo_Listad.Col = 0
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CliNCo_Listad.Col = 1
         grd_CliNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CliNCo_Listad.Col = 2
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 3
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 4
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 5
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 6
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 7
         grd_CliNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CliNCo_Listad.Col = 8
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_SegDes = r_dbl_SegDes + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_SegViv = r_dbl_SegViv + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_CliNCo_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_CliNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CliNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_SegDes, "###,###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_CliCon()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_CliCon_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CliCon_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
         grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
         
         r_dbl_ImpCuo = 0
         
         grd_CliCon_Listad.Col = 0
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CliCon_Listad.Col = 1
         grd_CliCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CliCon_Listad.Col = 2
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliCon_Listad.Text)
         
         grd_CliCon_Listad.Col = 3
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliCon_Listad.Text)
         
         grd_CliCon_Listad.Col = 4
         grd_CliCon_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CliCon_Listad.Col = 5
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_CliCon_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_CliCon_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CliCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub grd_CliCon_Listad_SelChange()
   If grd_CliCon_Listad.Rows > 2 Then
      grd_CliCon_Listad.RowSel = grd_CliCon_Listad.Row
   End If
End Sub

Private Sub grd_CliNCo_Listad_SelChange()
   If grd_CliNCo_Listad.Rows > 2 Then
      grd_CliNCo_Listad.RowSel = grd_CliNCo_Listad.Row
   End If
End Sub

Private Sub grd_MViNCo_Listad_SelChange()
   If grd_MViNCo_Listad.Rows > 2 Then
      grd_MViNCo_Listad.RowSel = grd_MViNCo_Listad.Row
   End If
End Sub

Private Sub grd_MViCon_Listad_SelChange()
   If grd_MviCon_Listad.Rows > 2 Then
      grd_MviCon_Listad.RowSel = grd_MviCon_Listad.Row
   End If
End Sub

Private Sub fs_Carga_Cro_MViNCo()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_Comisi     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_MViNCo_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_MViNCo_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         grd_MViNCo_Listad.Row = grd_MViNCo_Listad.Rows - 1
         
         r_dbl_ImpCuo = 0
         
         grd_MViNCo_Listad.Col = 0
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_MViNCo_Listad.Col = 1
         grd_MViNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_MViNCo_Listad.Col = 2
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 3
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 4
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 5
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 6
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 7
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 8
         grd_MViNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_MViNCo_Listad.Col = 9
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_SegDes = r_dbl_SegDes + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_SegViv = r_dbl_SegViv + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_MViNCo_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_MViNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_MViNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_MViNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_MViNCo_SegPre.Caption = Format(r_dbl_SegDes, "###,###,##0.00") & " "
   pnl_MViNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_MViNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_MViNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_MViNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_MViCon()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_Comisi     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_MviCon_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_MviCon_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_MviCon_Listad.Rows = grd_MviCon_Listad.Rows + 1
         grd_MviCon_Listad.Row = grd_MviCon_Listad.Rows - 1
         
         r_dbl_ImpCuo = 0
         
         grd_MviCon_Listad.Col = 0
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_MviCon_Listad.Col = 1
         grd_MviCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_MviCon_Listad.Col = 2
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MviCon_Listad.Text)
         
         grd_MviCon_Listad.Col = 3
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MviCon_Listad.Text)
         
         grd_MviCon_Listad.Col = 4
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MviCon_Listad.Text)
         
         grd_MviCon_Listad.Col = 5
         grd_MviCon_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_MviCon_Listad.Col = 6
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_MviCon_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_MviCon_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_MViCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_MViCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_MViCon_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_MViCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_CofNCo()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_Comisi     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_CofNCo_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 5 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CofNCo_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
         grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
         
         r_dbl_ImpCuo = 0
         
         grd_CofNCo_Listad.Col = 0
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CofNCo_Listad.Col = 1
         grd_CofNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CofNCo_Listad.Col = 2
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 3
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 4
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 5
         grd_CofNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CofNCo_Listad.Col = 6
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_CofNCo_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_CofNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub



