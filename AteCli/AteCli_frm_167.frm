VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_ConSol_59 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9540
   ClientLeft      =   2790
   ClientTop       =   1800
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_167.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   16801
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
         Left            =   30
         TabIndex        =   1
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
         Begin Threed.SSPanel pnl_AprCon 
            Height          =   555
            Left            =   8460
            TabIndex        =   2
            Top             =   60
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "CLIENTE CON APROBACION CONDICIONADA PENDIENTE"
            ForeColor       =   16777215
            BackColor       =   128
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
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   630
            TabIndex        =   3
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Solicitud de Crédito Hipotecario"
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
            Left            =   660
            TabIndex        =   4
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Trámites COFIDE"
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
            Picture         =   "AteCli_frm_167.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   5
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
            Picture         =   "AteCli_frm_167.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel25 
         Height          =   1785
         Left            =   30
         TabIndex        =   7
         Top             =   2250
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   1665
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   2937
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Datos Cliente"
            TabPicture(0)   =   "AteCli_frm_167.frx":0758
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos Inmueble"
            TabPicture(1)   =   "AteCli_frm_167.frx":0774
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Crédito"
            TabPicture(2)   =   "AteCli_frm_167.frx":0790
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   0
               Left            =   60
               TabIndex        =   9
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   2
               Left            =   -74940
               TabIndex        =   10
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1245
               Index           =   1
               Left            =   -74940
               TabIndex        =   11
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2196
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3465
         Left            =   30
         TabIndex        =   12
         Top             =   4080
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6112
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
         Begin TabDlg.SSTab tab_Seguim 
            Height          =   3345
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5900
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento en Instancia"
            TabPicture(0)   =   "AteCli_frm_167.frx":07AC
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label7"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_DesOcu"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel8"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel13"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "grd_LisOcu"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel10"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txt_Descar"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txt_Observ"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "AteCli_frm_167.frx":07C8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label4"
            Tab(1).Control(1)=   "Label3"
            Tab(1).Control(2)=   "Label2"
            Tab(1).Control(3)=   "pnl_TipAut"
            Tab(1).Control(4)=   "pnl_DesExc"
            Tab(1).Control(5)=   "SSPanel12"
            Tab(1).Control(6)=   "SSPanel11"
            Tab(1).Control(7)=   "SSPanel9"
            Tab(1).Control(8)=   "SSPanel5"
            Tab(1).Control(9)=   "SSPanel4"
            Tab(1).Control(10)=   "grd_LisExc"
            Tab(1).Control(11)=   "txt_ObsExc"
            Tab(1).ControlCount=   12
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "AteCli_frm_167.frx":07E4
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txt_ObsCon"
            Tab(2).Control(1)=   "txt_LevCon"
            Tab(2).Control(2)=   "SSPanel17"
            Tab(2).Control(3)=   "grd_LisCon"
            Tab(2).Control(4)=   "SSPanel18"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel20"
            Tab(2).Control(7)=   "pnl_InsCon"
            Tab(2).Control(8)=   "Label15"
            Tab(2).Control(9)=   "Label14"
            Tab(2).Control(10)=   "Label12"
            Tab(2).ControlCount=   11
            TabCaption(3)   =   "Rechazo de Solicitud"
            TabPicture(3)   =   "AteCli_frm_167.frx":0800
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txt_ObsRec"
            Tab(3).Control(1)=   "pnl_MotRec"
            Tab(3).Control(2)=   "Label19"
            Tab(3).Control(3)=   "Label10"
            Tab(3).ControlCount=   4
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   19
               Text            =   "AteCli_frm_167.frx":081C
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Text            =   "AteCli_frm_167.frx":0820
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Text            =   "AteCli_frm_167.frx":0824
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Text            =   "AteCli_frm_167.frx":0828
               Top             =   1980
               Width           =   10005
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Text            =   "AteCli_frm_167.frx":082C
               Top             =   2640
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsRec 
               Height          =   2595
               Left            =   -73560
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   690
               Width           =   9915
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   30
               TabIndex        =   20
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   855
               Left            =   30
               TabIndex        =   21
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
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
               TabIndex        =   22
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Ocurrencia"
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   2400
               TabIndex        =   23
               Top             =   360
               Width           =   8595
               _Version        =   65536
               _ExtentX        =   15161
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Ocurrencia"
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
               Left            =   1230
               TabIndex        =   24
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Ocurrencia"
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
            Begin Threed.SSPanel pnl_DesOcu 
               Height          =   315
               Left            =   1320
               TabIndex        =   25
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
               Height          =   855
               Left            =   -74970
               TabIndex        =   26
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   -74940
               TabIndex        =   27
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Excepción"
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
               Left            =   -69330
               TabIndex        =   28
               Top             =   360
               Width           =   5325
               _Version        =   65536
               _ExtentX        =   9393
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Descripción Excepción"
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
               Left            =   -73770
               TabIndex        =   29
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "H. Excepción"
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
               Left            =   -72600
               TabIndex        =   30
               Top             =   360
               Width           =   3285
               _Version        =   65536
               _ExtentX        =   5794
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
               Height          =   45
               Left            =   -74970
               TabIndex        =   31
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin Threed.SSPanel pnl_DesExc 
               Height          =   315
               Left            =   -73680
               TabIndex        =   32
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_TipAut 
               Height          =   315
               Left            =   -73650
               TabIndex        =   33
               Top             =   2970
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   45
               Left            =   -74970
               TabIndex        =   34
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
               Height          =   855
               Left            =   -74970
               TabIndex        =   35
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1508
               _Version        =   393216
               Rows            =   21
               Cols            =   4
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -74940
               TabIndex        =   36
               Top             =   360
               Width           =   2745
               _Version        =   65536
               _ExtentX        =   4842
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Instancia"
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
               Left            =   -65610
               TabIndex        =   37
               Top             =   360
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -72210
               TabIndex        =   38
               Top             =   360
               Width           =   6615
               _Version        =   65536
               _ExtentX        =   11668
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Condiciones de Aprobación"
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
            Begin Threed.SSPanel pnl_InsCon 
               Height          =   315
               Left            =   -73680
               TabIndex        =   39
               Top             =   1650
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin Threed.SSPanel pnl_MotRec 
               Height          =   315
               Left            =   -73560
               TabIndex        =   40
               Top             =   360
               Width           =   9915
               _Version        =   65536
               _ExtentX        =   17489
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
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   51
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   50
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   49
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label2 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   48
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   47
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   46
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   45
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   44
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   43
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Observaciones de Rechazo:"
               Height          =   555
               Left            =   -74940
               TabIndex        =   42
               Top             =   690
               Width           =   1155
            End
            Begin VB.Label Label10 
               Caption         =   "Motivo Rechazo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   41
               Top             =   360
               Width           =   1275
            End
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   52
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   53
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   54
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   7650
            TabIndex        =   55
            Top             =   30
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SOLICITUD EN TRAMITE"
            ForeColor       =   16711680
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
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   255
            Left            =   60
            TabIndex        =   58
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   57
            Top             =   390
            Width           =   645
         End
         Begin VB.Label Label6 
            Caption         =   "Situación:"
            Height          =   315
            Left            =   6240
            TabIndex        =   56
            Top             =   30
            Width           =   1005
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1875
         Left            =   30
         TabIndex        =   59
         Top             =   7590
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.TextBox txt_ObsEva 
            Height          =   705
            Left            =   30
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            Text            =   "AteCli_frm_167.frx":0830
            Top             =   1140
            Width           =   11445
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1065
            Left            =   30
            TabIndex        =   61
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   1879
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_ConSol_59"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgRec     As Integer
Dim l_int_AprCon     As Integer

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Situac.Caption = moddat_g_str_Situac
   
   Call fs_Inicia
   
   'Buscar Información de la Solicitud
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   
   Call fs_DatInm    'Datos del Inmueble
   Call fs_DatCre    'Datos del Crédito
   
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   Call fs_Buscar_LisCon      'Buscando Aprobaciones Condicionadas
   
   'Si Situación no es Rechazado
   If moddat_g_int_Situac = 1 Then
      pnl_Situac.ForeColor = modgen_g_con_ColAzu
   ElseIf moddat_g_int_Situac = 2 Then
      pnl_Situac.ForeColor = modgen_g_con_ColVer
   ElseIf moddat_g_int_Situac = 3 Then
      pnl_Situac.ForeColor = modgen_g_con_ColRoj
   End If
   
   'Si existe Rechazo en la Instancia
   If l_int_FlgRec = 0 Then
      tab_Seguim.TabVisible(3) = False
   End If
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If

   'Si no hay Aprobaciones Condicionadas
   If grd_LisCon.Rows = 0 Then
      tab_Seguim.TabVisible(2) = False
   End If
   
   'Si no hay Aprobaciones Condicionadas Pendiente
   If l_int_AprCon = 0 Then
      pnl_AprCon.Visible = False
   End If
   
   Call fs_Buscar_DatEva
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar_LisOcu()
   l_int_FlgRec = 0
   
   Call gs_LimpiaGrid(grd_LisOcu)
   
   moddat_g_int_NumObs = 0
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 62 "
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
      
      If g_rst_Princi!SEGDET_CODOCU = 21 Then
         If g_rst_Princi!SEGFECACT > 0 Then
            grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
            grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
         Else
            moddat_g_int_NumObs = g_rst_Princi!SEGDET_NUMOBS
         End If
      ElseIf g_rst_Princi!SEGDET_CODOCU = 13 Then
         'Si la Solicitud está rechazada en la Instancia
         pnl_MotRec.Caption = moddat_gf_Consulta_ParDes("003", CStr(g_rst_Princi!SEGDET_MOTREC))
         txt_ObsRec.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
         
         l_int_FlgRec = 1
      End If
      
      grd_LisOcu.Col = 3
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu.Col = 4
      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisOcu.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu)
   Call grd_LisOcu_Click
End Sub

Private Sub fs_Buscar_LisExc()
   Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisExc)
   
   g_str_Parame = "SELECT * FROM TRA_SEGEXC WHERE "
   g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisExc.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisExc.Rows = grd_LisExc.Rows + 1
      grd_LisExc.Row = grd_LisExc.Rows - 1
      
      'Fecha de Excepción
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepción
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripción Excepción
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorización
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisExc.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisExc)
   Call grd_LisExc_Click
End Sub

Private Sub fs_Buscar_LisCon()
   l_int_AprCon = 0
   
   Call gs_LimpiaGrid(grd_LisCon)
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     Exit Sub
   End If
   
   grd_LisCon.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisCon.Rows = grd_LisCon.Rows + 1
      grd_LisCon.Row = grd_LisCon.Rows - 1
      
      'Instancia
      grd_LisCon.Col = 0
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      
      'Descripción Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situación
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      If g_rst_Princi!SEGCON_SITUAC = 1 Then
         l_int_AprCon = 1
      End If
      
      'Descripción Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSLEV & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisCon.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisCon)
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisCon_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisCon.Rows > 0 Then
      grd_LisCon.Col = 0
      pnl_InsCon.Caption = grd_LisCon.Text
   
      grd_LisCon.Col = 1
      txt_ObsCon.Text = grd_LisCon.Text
      
      grd_LisCon.Col = 3
      txt_LevCon.Text = grd_LisCon.Text
      
      Call gs_RefrescaGrid(grd_LisCon)
   End If
End Sub

Private Sub grd_LisCon_SelChange()
   If grd_LisCon.Rows > 2 Then
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
   
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
End Sub

Private Sub grd_LisExc_Click()
   Dim r_str_FecExc     As String
   Dim r_str_HorExc     As String
   Dim r_str_InsExc     As String

   If grd_LisExc.Rows > 0 Then
      grd_LisExc.Col = 0
      r_str_FecExc = grd_LisExc.Text
      
      grd_LisExc.Col = 1
      r_str_HorExc = grd_LisExc.Text
      
      grd_LisExc.Col = 2
      r_str_InsExc = grd_LisExc.Text
      
      pnl_DesExc.Caption = "Día: " & r_str_FecExc & " - " & r_str_HorExc & " hrs. - " & r_str_InsExc
   
      grd_LisExc.Col = 3
      txt_ObsExc.Text = grd_LisExc.Text
      
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
      
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
   End If
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   
   Call grd_LisExc_Click
End Sub

Private Sub grd_LisOcu_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisOcu.Rows > 0 Then
      grd_LisOcu.Col = 0
      r_str_FecOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 1
      r_str_HorOcu = grd_LisOcu.Text
      
      grd_LisOcu.Col = 2
      r_str_DesOcu = grd_LisOcu.Text
      
      pnl_DesOcu.Caption = "Día: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu.Col = 3
      txt_Observ.Text = grd_LisOcu.Text
      
      grd_LisOcu.Col = 4
      txt_Descar.Text = grd_LisOcu.Text
      
      Call gs_RefrescaGrid(grd_LisOcu)
   End If
End Sub

Private Sub grd_LisOcu_SelChange()
   If grd_LisOcu.Rows > 2 Then
      grd_LisOcu.RowSel = grd_LisOcu.Row
   End If
   
   Call grd_LisOcu_Click
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsRec_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   Dim r_str_TipCli     As String
   
   r_str_TipCli = ""

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(0).Redraw = False
      
      If p_Indice = 1 Then
         r_str_TipCli = " (Cónyuge)"
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      End If
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Documento de Identidad" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Apellidos y Nombres" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Estado Civil"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If

      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Celular" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Domicilio"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Referencia"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Teléfono Domicilio"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(0).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(0))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(1).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Tipo de Inmueble"
         
      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Dirección"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Referencia"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Estacionamiento"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Departamento / Provincia / Distrito"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Proyecto miCasita"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = "Nombre Proyecto"
   
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Nombre Proyecto"
   
               grd_Listad(1).Col = 1
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(1).Rows = grd_Listad(1).Rows + 2
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Propietario / Promotor"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Nombre o Razón Social"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Dirección"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Referencia"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Teléfono"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Docum. Identidad Constructor"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Nombre o Razón Social"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Dirección"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Referencia"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Teléfono"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Proyecto Vinculado"
            Else
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Entidad Financiera"
         
               grd_Listad(1).Col = 1
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Docum. Identidad Propietario"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Nombre o Razón Social"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Dirección"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Referencia"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Teléfono"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Doc. Ident. Promotor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Razón Social Promotor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Doc. Ident. Constructor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Razón Social Constructor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(1).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(1))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(2))
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad(2).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Sub-Producto"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Tipo de Evaluación"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Moneda del Préstamo"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Fecha de Solicitud"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Tasa de Interés"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   
   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Valor de Compra Venta"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Aporte Propio"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Monto Préstamo"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
      Else
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Valor de Compra Venta"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Aporte Propio"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).Text = "Monto Préstamo"
      
         grd_Listad(2).Col = 1
         grd_Listad(2).CellFontName = "Lucida Console"
         grd_Listad(2).CellFontSize = 8
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
      End If
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Plazo (Años)"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Período de Gracia (Meses)"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_PERGRA)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Cuotas Extraordinarias"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Compañía de Seguros"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Día de Pago"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Institución Financiera de Ahorro"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Monto Mínimo de Ahorro Mensual"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!SOLMAE_MONAHO) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Meses Ahorrados"
   
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
   End If
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Observaciones"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Consejero Hipotecario"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Ejecutivo de Seguimiento"

   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   grd_Listad(2).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(2))
   
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 2
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
   
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad

   'Lista de Ocurrencias
   grd_LisOcu.ColWidth(0) = 1155
   grd_LisOcu.ColWidth(1) = 1185
   grd_LisOcu.ColWidth(2) = 8595
   grd_LisOcu.ColWidth(3) = 0
   grd_LisOcu.ColWidth(4) = 0
   
   grd_LisOcu.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOcu.ColAlignment(2) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_LisOcu)

   pnl_DesOcu.Caption = ""
   txt_Observ.Text = ""
   txt_Descar.Text = ""

   'Lista de Excepciones
   grd_LisExc.ColWidth(0) = 1175
   grd_LisExc.ColWidth(1) = 1175
   grd_LisExc.ColWidth(2) = 3275
   grd_LisExc.ColWidth(3) = 5325
   grd_LisExc.ColWidth(4) = 0
   
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_LisExc)

   pnl_DesExc.Caption = ""
   txt_ObsExc.Text = ""
   pnl_TipAut.Caption = ""

   'Lista de Aprobaciones Condicionadas
   grd_LisCon.ColWidth(0) = 2735
   grd_LisCon.ColWidth(1) = 6605
   grd_LisCon.ColWidth(2) = 1625
   grd_LisCon.ColWidth(3) = 0
   
   grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(2) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_LisCon)

   pnl_InsCon.Caption = ""
   txt_ObsCon.Text = ""
   txt_LevCon.Text = ""

   'Lista de Datos de Evaluación
   grd_LisEva.ColWidth(0) = 3300
   grd_LisEva.ColWidth(1) = 7940

   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar_DatEva()
   txt_ObsEva.Text = ""
   Call gs_LimpiaGrid(grd_LisEva)
   
   g_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
   g_str_Parame = g_str_Parame & "EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Fecha Envío"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECENV))
   
      If g_rst_Princi!EVACOF_FECREC > 0 Then
         If moddat_g_str_CodPrd = "003" Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Operación Mivivienda"
      
            grd_LisEva.Col = 1
            grd_LisEva.Text = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Fecha Aprobación Mivivienda"
      
            grd_LisEva.Col = 1
            grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_APRMVI))
         End If
      
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Nro. Carta COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = Trim(g_rst_Princi!EVACOF_NUMCAR & "")
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Fecha Recepción Carta COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECREC))
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Nro. Operación COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = Trim(g_rst_Princi!EVACOF_CODMVI & "")
      
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Fecha Desembolso COFIDE"
   
         grd_LisEva.Col = 1
         grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
         
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Importe Desembolsado"
   
         grd_LisEva.Col = 1
         grd_LisEva.CellFontName = "Lucida Console"
         grd_LisEva.CellFontSize = 8
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACOF_MTODES, 12, 2)
         
         txt_ObsEva.Text = Trim(g_rst_Princi!EVACOF_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(grd_LisEva)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub



