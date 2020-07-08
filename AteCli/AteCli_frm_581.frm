VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Tra_CarAFP_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12345
   Icon            =   "AteCli_frm_581.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12345
      _Version        =   65536
      _ExtentX        =   21775
      _ExtentY        =   17754
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   3465
         Left            =   30
         TabIndex        =   1
         Top             =   4680
         Width           =   12255
         _Version        =   65536
         _ExtentX        =   21616
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
            TabIndex        =   2
            Top             =   60
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   5900
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento de Instancias"
            TabPicture(0)   =   "AteCli_frm_581.frx":000C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "txt_Observ"
            Tab(0).Control(1)=   "txt_Descar"
            Tab(0).Control(2)=   "SSPanel10"
            Tab(0).Control(3)=   "grd_LisOcu"
            Tab(0).Control(4)=   "SSPanel13"
            Tab(0).Control(5)=   "SSPanel14"
            Tab(0).Control(6)=   "SSPanel5"
            Tab(0).Control(7)=   "pnl_DesOcu"
            Tab(0).Control(8)=   "Label7"
            Tab(0).Control(9)=   "Label8"
            Tab(0).Control(10)=   "Label11"
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "Excepciones Aplicadas"
            TabPicture(1)   =   "AteCli_frm_581.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_motivo"
            Tab(1).Control(1)=   "Label4"
            Tab(1).Control(2)=   "Label3"
            Tab(1).Control(3)=   "Label6"
            Tab(1).Control(4)=   "pnl_motivo"
            Tab(1).Control(5)=   "pnl_TipAut"
            Tab(1).Control(6)=   "pnl_DesExc"
            Tab(1).Control(7)=   "SSPanel16"
            Tab(1).Control(8)=   "SSPanel15"
            Tab(1).Control(9)=   "SSPanel12"
            Tab(1).Control(10)=   "SSPanel11"
            Tab(1).Control(11)=   "SSPanel9"
            Tab(1).Control(12)=   "grd_LisExc"
            Tab(1).Control(13)=   "txt_ObsExc"
            Tab(1).ControlCount=   14
            TabCaption(2)   =   "Aprobación Condicionada"
            TabPicture(2)   =   "AteCli_frm_581.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label12"
            Tab(2).Control(1)=   "Label14"
            Tab(2).Control(2)=   "Label15"
            Tab(2).Control(3)=   "pnl_InsCon"
            Tab(2).Control(4)=   "SSPanel20"
            Tab(2).Control(5)=   "SSPanel19"
            Tab(2).Control(6)=   "SSPanel18"
            Tab(2).Control(7)=   "grd_LisCon"
            Tab(2).Control(8)=   "SSPanel17"
            Tab(2).Control(9)=   "txt_LevCon"
            Tab(2).Control(10)=   "txt_ObsCon"
            Tab(2).ControlCount=   11
            TabCaption(3)   =   "Seguimiento AFP"
            TabPicture(3)   =   "AteCli_frm_581.frx":0060
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "Label16"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "Label13"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "Label10"
            Tab(3).Control(2).Enabled=   0   'False
            Tab(3).Control(3)=   "pnl_DesOcu_AFP"
            Tab(3).Control(3).Enabled=   0   'False
            Tab(3).Control(4)=   "SSPanel25"
            Tab(3).Control(4).Enabled=   0   'False
            Tab(3).Control(5)=   "SSPanel23"
            Tab(3).Control(5).Enabled=   0   'False
            Tab(3).Control(6)=   "SSPanel22"
            Tab(3).Control(6).Enabled=   0   'False
            Tab(3).Control(7)=   "SSPanel8"
            Tab(3).Control(7).Enabled=   0   'False
            Tab(3).Control(8)=   "grd_LisOcu_AFP"
            Tab(3).Control(8).Enabled=   0   'False
            Tab(3).Control(9)=   "txt_Descar_AFP"
            Tab(3).Control(9).Enabled=   0   'False
            Tab(3).Control(10)=   "txt_Observ_AFP"
            Tab(3).Control(10).Enabled=   0   'False
            Tab(3).ControlCount=   11
            Begin VB.TextBox txt_Observ_AFP 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   70
               Text            =   "AteCli_frm_581.frx":007C
               Top             =   1980
               Width           =   10725
            End
            Begin VB.TextBox txt_Descar_AFP 
               Height          =   645
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   69
               Text            =   "AteCli_frm_581.frx":0080
               Top             =   2640
               Width           =   10725
            End
            Begin VB.TextBox txt_ObsCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "AteCli_frm_581.frx":0084
               Top             =   1980
               Width           =   10755
            End
            Begin VB.TextBox txt_LevCon 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Text            =   "AteCli_frm_581.frx":0088
               Top             =   2640
               Width           =   10755
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   975
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   1980
               Width           =   10755
            End
            Begin VB.TextBox txt_Observ 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "AteCli_frm_581.frx":008C
               Top             =   1980
               Width           =   10755
            End
            Begin VB.TextBox txt_Descar 
               Height          =   645
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "AteCli_frm_581.frx":0090
               Top             =   2640
               Width           =   10755
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   -74970
               TabIndex        =   8
               Top             =   1560
               Width           =   12045
               _Version        =   65536
               _ExtentX        =   21246
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
               Left            =   -74970
               TabIndex        =   9
               Top             =   660
               Width           =   12045
               _ExtentX        =   21246
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
               Left            =   -74940
               TabIndex        =   10
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
               Left            =   -72600
               TabIndex        =   11
               Top             =   360
               Width           =   9375
               _Version        =   65536
               _ExtentX        =   16536
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -73770
               TabIndex        =   12
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
               Left            =   -73680
               TabIndex        =   13
               Top             =   1650
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
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
               TabIndex        =   14
               Top             =   660
               Width           =   12045
               _ExtentX        =   21246
               _ExtentY        =   1508
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -74940
               TabIndex        =   15
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -69330
               TabIndex        =   16
               Top             =   360
               Width           =   6075
               _Version        =   65536
               _ExtentX        =   10716
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -73770
               TabIndex        =   17
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -72600
               TabIndex        =   18
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   45
               Left            =   -74970
               TabIndex        =   19
               Top             =   1560
               Width           =   12045
               _Version        =   65536
               _ExtentX        =   21246
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
               Left            =   -73710
               TabIndex        =   20
               Top             =   1650
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
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
               Left            =   -73710
               TabIndex        =   21
               Top             =   2970
               Width           =   4155
               _Version        =   65536
               _ExtentX        =   7329
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "INGRESO A INSTANCIA"
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
               TabIndex        =   22
               Top             =   1560
               Width           =   12045
               _Version        =   65536
               _ExtentX        =   21246
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
               TabIndex        =   23
               Top             =   660
               Width           =   12045
               _ExtentX        =   21246
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
               TabIndex        =   24
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
               TabIndex        =   25
               Top             =   360
               Width           =   2355
               _Version        =   65536
               _ExtentX        =   4154
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
               TabIndex        =   26
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
               TabIndex        =   27
               Top             =   1650
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
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
            Begin Threed.SSPanel pnl_motivo 
               Height          =   315
               Left            =   -68580
               TabIndex        =   28
               Top             =   2970
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "MOTIVO DE EXCEPCION"
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
               Begin Threed.SSPanel SSPanel21 
                  Height          =   315
                  Left            =   6090
                  TabIndex        =   29
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   5205
                  _Version        =   65536
                  _ExtentX        =   9181
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "INGRESOS 4A CATEG. NO SE PUEDEN CONFIRMAR "
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu_AFP 
               Height          =   855
               Left            =   30
               TabIndex        =   71
               Top             =   660
               Width           =   12045
               _ExtentX        =   21246
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
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   60
               TabIndex        =   72
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   1230
               TabIndex        =   73
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
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   2400
               TabIndex        =   74
               Top             =   360
               Width           =   9315
               _Version        =   65536
               _ExtentX        =   16431
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   45
               Left            =   0
               TabIndex        =   75
               Top             =   1560
               Width           =   11325
               _Version        =   65536
               _ExtentX        =   19976
               _ExtentY        =   79
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.21
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
            End
            Begin Threed.SSPanel pnl_DesOcu_AFP 
               Height          =   315
               Left            =   1320
               TabIndex        =   76
               Top             =   1650
               Width           =   10725
               _Version        =   65536
               _ExtentX        =   18918
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Día: 10/05/2008 - 17:00 hrs - INGRESO A INSTANCIA"
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
            Begin VB.Label Label10 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   79
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label13 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   78
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label16 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   77
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label Label15 
               Caption         =   "Condiciones de Aprobación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   39
               Top             =   2010
               Width           =   1155
            End
            Begin VB.Label Label14 
               Caption         =   "Instancia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   38
               Top             =   1680
               Width           =   1155
            End
            Begin VB.Label Label12 
               Caption         =   "Levantamiento de Condiciones:"
               Height          =   615
               Left            =   -74940
               TabIndex        =   37
               Top             =   2670
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   36
               Top             =   2970
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   35
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   34
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   33
               Top             =   1980
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   32
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   31
               Top             =   2640
               Width           =   1035
            End
            Begin VB.Label lbl_motivo 
               Caption         =   "Motivo:"
               Height          =   255
               Left            =   -69300
               TabIndex        =   30
               Top             =   3030
               Width           =   645
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   40
         Top             =   30
         Width           =   12255
         _Version        =   65536
         _ExtentX        =   21616
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
            TabIndex        =   41
            Top             =   60
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Cartas de Pre Conformidad AFP"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   11760
            Top             =   120
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   11130
            Top             =   0
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
            Left            =   10560
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_581.frx":0094
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   645
         Left            =   30
         TabIndex        =   42
         Top             =   720
         Width           =   12255
         _Version        =   65536
         _ExtentX        =   21616
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
         Begin VB.CommandButton cmd_NueObs 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_581.frx":039E
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Descargo de Observación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_581.frx":07E0
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Grabar datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_581.frx":0C22
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Impresión formato Pre Conformidad AFP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11640
            Picture         =   "AteCli_frm_581.frx":1064
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   45
         Top             =   1380
         Width           =   12255
         _Version        =   65536
         _ExtentX        =   21616
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   46
            Top             =   390
            Width           =   10755
            _Version        =   65536
            _ExtentX        =   18971
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   47
            Top             =   60
            Width           =   2235
            _Version        =   65536
            _ExtentX        =   3942
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
         Begin Threed.SSPanel pnl_FecSol 
            Height          =   315
            Left            =   10170
            TabIndex        =   48
            Top             =   60
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/01/9999"
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
         Begin VB.Label Label2 
            Caption         =   "Fecha Solicitud:"
            Height          =   315
            Left            =   8790
            TabIndex        =   51
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   50
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   420
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2505
         Left            =   30
         TabIndex        =   52
         Top             =   2160
         Width           =   12255
         _Version        =   65536
         _ExtentX        =   21616
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   2385
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   4207
            _Version        =   393216
            Style           =   1
            Tabs            =   8
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Datos Cliente"
            TabPicture(0)   =   "AteCli_frm_581.frx":14A6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos Cónyuge"
            TabPicture(1)   =   "AteCli_frm_581.frx":14C2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Apoderado"
            TabPicture(2)   =   "AteCli_frm_581.frx":14DE
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "AteCli_frm_581.frx":14FA
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos Crédito"
            TabPicture(4)   =   "AteCli_frm_581.frx":1516
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(4)"
            Tab(4).Control(1)=   "txt_ObsSol"
            Tab(4).ControlCount=   2
            TabCaption(5)   =   "Ev. Crediticia"
            TabPicture(5)   =   "AteCli_frm_581.frx":1532
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(5)"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Tasación"
            TabPicture(6)   =   "AteCli_frm_581.frx":154E
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "grd_Listad(6)"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "Ev. Seguros"
            TabPicture(7)   =   "AteCli_frm_581.frx":156A
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(7)"
            Tab(7).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   405
               Left            =   -73710
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   54
               Text            =   "AteCli_frm_581.frx":1586
               Top             =   1920
               Width           =   10785
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1965
               Index           =   0
               Left            =   60
               TabIndex        =   55
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
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
               Height          =   1965
               Index           =   1
               Left            =   -74940
               TabIndex        =   56
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
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
               Height          =   1965
               Index           =   2
               Left            =   -74940
               TabIndex        =   57
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
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
               Height          =   1965
               Index           =   3
               Left            =   -74940
               TabIndex        =   58
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1545
               Index           =   4
               Left            =   -74940
               TabIndex        =   59
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   2725
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
               Height          =   1965
               Index           =   5
               Left            =   -74940
               TabIndex        =   60
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
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
               Height          =   1965
               Index           =   6
               Left            =   -74940
               TabIndex        =   61
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
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
               Height          =   1965
               Index           =   7
               Left            =   -74940
               TabIndex        =   62
               Top             =   360
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   3466
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   1965
         Left            =   30
         TabIndex        =   63
         Top             =   8160
         Width           =   12255
         _Version        =   65536
         _ExtentX        =   21616
         _ExtentY        =   3466
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
         Begin TabDlg.SSTab tab_Princi 
            Height          =   1875
            Left            =   30
            TabIndex        =   64
            Top             =   30
            Width           =   12165
            _ExtentX        =   21458
            _ExtentY        =   3307
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Comentario Comercial"
            TabPicture(0)   =   "AteCli_frm_581.frx":158A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "txt_ComCom"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Comentario Legal"
            TabPicture(1)   =   "AteCli_frm_581.frx":15A6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txt_ComLeg"
            Tab(1).ControlCount=   1
            Begin VB.TextBox txt_ComLeg 
               Height          =   1425
               Left            =   -74970
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   67
               Text            =   "AteCli_frm_581.frx":15C2
               Top             =   360
               Width           =   12075
            End
            Begin VB.TextBox txt_ComCom 
               Height          =   1425
               Left            =   30
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   65
               Text            =   "AteCli_frm_581.frx":15C9
               Top             =   360
               Width           =   12075
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_CarAFP_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_AprCon     As Integer

Private Sub cmd_Grabar_Click()
   If txt_ComCom.Text = "" Then
      MsgBox "Debe ingresar Comentario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ComCom)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de guardar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Actualizando Comentario
   If Not fs_InsAct_TraAfp(moddat_g_str_NumSol, moddat_g_int_CodIns, 3, Trim(txt_ComCom.Text)) Then
      Exit Sub
   End If
End Sub

Private Sub cmd_Imprim_Click()
   moddat_g_int_FlgAct_1 = 1
   Call fs_Imprimir
      
   If moddat_g_int_FlgAct_1 = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_LisOcu            'Buscando Ocurrencias de Instancia
      Call fs_Buscar_LisOcu_AFP        'Buscando Ocurrencias de AFP
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_NueObs_Click()
   If moddat_g_int_NumObs = 0 Then
      MsgBox "No hay observación pendiente de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_DesObs = ""
   moddat_g_int_FlgAct_1 = 1
   frm_Tra_CarAFP_03.Show 1

   If moddat_g_int_FlgAct_1 = 2 Then
'      If moddat_g_int_CodIns = 32 Then l_int_CodIns = 31
'      If moddat_g_int_CodIns = 42 Then l_int_CodIns = 41
      
      If Not moddat_gf_Modifica_SegDet_Observ(moddat_g_str_NumSol, moddat_g_int_CodIns, 99, CStr(moddat_g_int_NumObs), moddat_g_str_DesObs, 2) Then 'l_int_CodIns,
         Exit Sub
      End If
   
'      'Actualizando en Instancia
'      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, moddat_g_int_CodIns, 0, 9, 2) Then
'         Exit Sub
'      End If
   
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = moddat_gf_Consulta_ParDes("002", CStr(moddat_g_int_CodIns)) & " - DESCARGO DE OBSERVACION " & "(Cliente: " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " - " & moddat_g_str_NomCli & ")"
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & gf_Formato_NumSol(moddat_g_str_NumSol) & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "FECHA               : " & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "HORA                : " & Format(Time, "hh:mm:ss") & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_DesObs
   
      Call fs_Envia_CorreoEle(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodConHip, moddat_g_str_CodEjeSeg, moddat_g_str_NumSol, 1, True, False, False)
   
      Screen.MousePointer = 11
      Call fs_Buscar_LisOcu            'Buscando Ocurrencias de Instancia
      Call fs_Buscar_LisOcu_AFP        'Buscando Ocurrencias de Instancia AFP
      Screen.MousePointer = 0
      
      moddat_g_int_FlgAct = 2
   End If
End Sub

Private Sub cmd_Salida_Click()
   If moddat_g_int_TipRep = 1 Then
      frm_Tra_CarAFP_01.fs_Buscar
   End If
   Unload Me
End Sub

Private Sub fs_Imprimir()
Dim r_int_CodIns     As Integer

   If MsgBox("¿Está seguro de imprimir el Formato?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
     
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT MAX(B.SEGDET_CODINS) AS CODINS FROM TRA_SEGDET B WHERE B.SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY B.SEGFECCRE DESC, B.SEGHORCRE DESC "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
'      Exit Sub
'   End If
'
'   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
'      g_rst_Genera.MoveFirst
'      r_int_CodIns = g_rst_Genera!CODINS
'   End If
'
'   g_rst_Genera.Close
'   Set g_rst_Genera = Nothing
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, r_int_CodIns, 100, 0, "", 0, 0) Then '62
      Exit Sub
   End If
   
   'Actualizando Fecha de Impresión
   If Not fs_InsAct_TraAfp(moddat_g_str_NumSol, r_int_CodIns, 2, "") Then
      Exit Sub
   End If

   crp_Imprim.Reset
   crp_Imprim.WindowTitle = "Presentacion Preliminar"
   crp_Imprim.WindowHeight = 730
   crp_Imprim.WindowWidth = 1400
   crp_Imprim.WindowLeft = 0
   crp_Imprim.WindowTop = 0
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
   crp_Imprim.DataFiles(1) = "CLI_DATGEN"
   crp_Imprim.DataFiles(2) = "MNT_PARDES"

   crp_Imprim.SelectionFormula = "{CRE_SOLMAE.SOLMAE_NUMERO} = '" & Trim(moddat_g_str_NumSol) & "' "
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_TRAAFP_01.RPT"
   crp_Imprim.WindowShowPrintBtn = True
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Function fs_InsAct_TraAfp(ByVal p_NumSol As String, ByVal p_CodIns As Integer, ByVal p_TipOpe As Integer, ByVal p_Coment As String) As Integer
   fs_InsAct_TraAfp = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_TRAAFP ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitu
      g_str_Parame = g_str_Parame & CStr(0) & ", "                                                 'Tipo de Operación: 1- Inicial, 2- PPG
      g_str_Parame = g_str_Parame & "'', "                                                         'Comentario Legal
      g_str_Parame = g_str_Parame & CStr(0) & ", "                                                 'Fecha de Evaluación Legal
      g_str_Parame = g_str_Parame & CStr(Format(moddat_g_str_FecSis, "yyyymmdd")) & ", "           'Fecha de Impresión
      g_str_Parame = g_str_Parame & "'" & p_Coment & "', "                                         'Comentario Operaciones
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                                          'Instancia
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(p_TipOpe) & " )"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_TRAAFP. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   fs_InsAct_TraAfp = True
End Function

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   moddat_g_int_CodIns = 33
  
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_FecSol.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   Call fs_Activa(True)
   
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatApo(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(2))         'Buscar Información del Apoderado
   Call modmip_gs_DatInm(grd_Listad(3), False)                                            'Buscar Información del Inmueble
   
   'Buscar Información del Crédito
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   txt_ObsSol.Text = r_arr_Mtz(0).DatCom_Observ
   
   Call fs_EvaCre                                                                         'Datos de Evaluación Crediticia
   Call modmip_gs_EvaTas(grd_Listad(6))                                                   'Datos de Tasación
   Call modmip_gs_EvaSeg(grd_Listad(7))                                                   'Datos de Seguros
   Call fs_Buscar_LisOcu_AFP                                                              'Buscando Ocurrencias de Carta de Pre-conformidad AFP
   Call fs_Buscar_Coment                                                                  'Buscar comentarios de Legal y Operaciones
      
   'No se muestra Ocurrencias de Instancia
   tab_Seguim.TabVisible(0) = False
      
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If

   'Si no hay Aprobaciones Condicionadas
   If grd_LisCon.Rows = 0 Then
      tab_Seguim.TabVisible(2) = False
   End If
   
   'Si no hay Ocurrencias para carta de pre-conformidad AFP
   If grd_LisOcu_AFP.Rows = 0 Then
      tab_Seguim.TabVisible(3) = False
   End If
      
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 7
      grd_Listad(r_int_Contad).ColWidth(0) = 2900
      grd_Listad(r_int_Contad).ColWidth(1) = 8800
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad

   'Lista de Ocurrencias
   grd_LisOcu.ColWidth(0) = 1155
   grd_LisOcu.ColWidth(1) = 1185
   grd_LisOcu.ColWidth(2) = 9500
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
   grd_LisExc.ColWidth(3) = 6500
   grd_LisExc.ColWidth(4) = 0
   grd_LisExc.ColWidth(5) = 0
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisExc)

   pnl_DesExc.Caption = ""
   txt_ObsExc.Text = ""
   pnl_TipAut.Caption = ""
   pnl_motivo.Caption = ""

   'Lista de Aprobaciones Condicionadas
   grd_LisCon.ColWidth(0) = 2735
   grd_LisCon.ColWidth(1) = 6605
   grd_LisCon.ColWidth(2) = 2250
   grd_LisCon.ColWidth(3) = 0
   grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(2) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisCon)

   pnl_InsCon.Caption = ""
   txt_ObsCon.Text = ""
   txt_LevCon.Text = ""
   txt_ComCom.Text = ""
   txt_ComLeg.Text = ""
   
   'Lista de Ocurrencias Carta de pre-conformidad afp
   grd_LisOcu_AFP.ColWidth(0) = 1175 '1155
   grd_LisOcu_AFP.ColWidth(1) = 1175 '1185
   grd_LisOcu_AFP.ColWidth(2) = 9500
   grd_LisOcu_AFP.ColWidth(3) = 0
   grd_LisOcu_AFP.ColWidth(4) = 0
   grd_LisOcu_AFP.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOcu_AFP.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOcu_AFP.ColAlignment(2) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_LisOcu_AFP)

   pnl_DesOcu_AFP.Caption = ""
   txt_Observ_AFP.Text = ""
   txt_Descar_AFP.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Activa As Boolean)
   If moddat_g_int_FlgCre = 1 Then
      cmd_NueObs.Enabled = p_Activa
      cmd_Grabar.Enabled = p_Activa
      cmd_Imprim.Enabled = p_Activa
   ElseIf moddat_g_int_FlgCre = 3 Then
      cmd_NueObs.Enabled = p_Activa
      cmd_Grabar.Enabled = Not p_Activa
      cmd_Imprim.Enabled = Not p_Activa
   Else
      cmd_NueObs.Enabled = Not p_Activa
      cmd_Grabar.Enabled = Not p_Activa
      cmd_Imprim.Enabled = Not p_Activa
   End If
End Sub

Private Sub fs_Buscar_LisOcu()
Dim r_str_FecOcu  As String
Dim r_int_CodIns  As Integer
   
   Call gs_LimpiaGrid(grd_LisOcu)
   moddat_g_int_NumObs = 0
   
   'Obtiene la instancia actual
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAX(B.SEGDET_CODINS) AS CODINS FROM TRA_SEGDET B WHERE B.SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY B.SEGFECCRE DESC, B.SEGHORCRE DESC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      r_int_CodIns = g_rst_Genera!CODINS
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT SEGFECCRE, SEGHORCRE, SEGDET_CODOCU, SEGFECACT, SEGHORACT, SEGDET_OBSERV, SEGDET_OBSDES "
   g_str_Parame = g_str_Parame & "    FROM TRA_SEGDET "
   g_str_Parame = g_str_Parame & "   WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "     AND SEGDET_CODINS = " & r_int_CodIns & " "
   g_str_Parame = g_str_Parame & "     AND SEGDET_CODOCU NOT IN (94,95,96,97,98,99,100,101,102,103) "
   g_str_Parame = g_str_Parame & "   ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "
   
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

Private Sub fs_Buscar_LisOcu_AFP()
Dim r_str_FecOcu  As String
      
   Call gs_LimpiaGrid(grd_LisOcu_AFP)
'   If moddat_g_int_CodIns = 32 Then l_int_CodIns = 31
'   If moddat_g_int_CodIns = 42 Then l_int_CodIns = 41
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.SEGFECCRE, A.SEGHORCRE, A.SEGDET_CODOCU, A.SEGFECACT, A.SEGHORACT, A.SEGDET_OBSERV, A.SEGDET_OBSDES, A.SEGDET_NUMOBS "
   g_str_Parame = g_str_Parame & "   FROM TRA_SEGDET A "
   g_str_Parame = g_str_Parame & "  WHERE A.SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "    AND A.SEGDET_CODINS = " & moddat_g_int_CodIns & " " 'l_int_CodIns
   g_str_Parame = g_str_Parame & "    AND A.SEGDET_CODOCU IN (94,95,96,97,98,99,100,101,102,103) "
   g_str_Parame = g_str_Parame & "  ORDER BY A.SEGFECCRE DESC, A.SEGHORCRE DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_LisOcu_AFP.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_LisOcu_AFP.Rows = grd_LisOcu_AFP.Rows + 1
      grd_LisOcu_AFP.Row = grd_LisOcu_AFP.Rows - 1
      
      'Fecha de Ocurrencia
      grd_LisOcu_AFP.Col = 0
      grd_LisOcu_AFP.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Ocurrencia
      grd_LisOcu_AFP.Col = 1
      grd_LisOcu_AFP.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Descripción Ocurrencia
      grd_LisOcu_AFP.Col = 2
      grd_LisOcu_AFP.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
      
      If g_rst_Princi!SEGFECACT > 0 Then
         r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
         
         grd_LisOcu_AFP.Text = grd_LisOcu_AFP.Text & " (DESCARGO EFECTUADO - " & r_str_FecOcu
         grd_LisOcu_AFP.Text = grd_LisOcu_AFP.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
      End If
      
      If g_rst_Princi!SEGDET_CODOCU = 99 Then
         If g_rst_Princi!SEGFECACT = 0 Then
            moddat_g_int_NumObs = g_rst_Princi!SEGDET_NUMOBS
         End If
      End If
      
      grd_LisOcu_AFP.Col = 3
      grd_LisOcu_AFP.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
      
      grd_LisOcu_AFP.Col = 4
      grd_LisOcu_AFP.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
      
      g_rst_Princi.MoveNext
   Loop
   grd_LisOcu_AFP.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisOcu_AFP)
   Call grd_LisOcu_AFP_Click
End Sub

Private Sub fs_Buscar_LisExc()
Dim r_str_FecOcu  As String
   
   Call gs_LimpiaGrid(grd_LisExc)
   g_str_Parame = modgen_gf_Buscar_Excepc
   
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
      
      'Motivo de Excepción
      grd_LisExc.Col = 5
      grd_LisExc.Text = Trim(g_rst_Princi!PARDES_DESCRI)
      
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
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_SEGCON "
   g_str_Parame = g_str_Parame & " WHERE SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & " ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
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

Private Sub fs_Buscar_Coment()
   txt_ComCom.Text = ""
   txt_ComLeg.Text = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRAAFP_COMLEG, TRAAFP_COMOPE FROM CRE_TRAAFP "
   g_str_Parame = g_str_Parame & " WHERE TRAAFP_NUMSOL = '" & moddat_g_str_NumSol & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If Not g_rst_Princi.EOF Then
      If Not IsNull(g_rst_Princi!TRAAFP_COMOPE) Then
         txt_ComCom.Text = g_rst_Princi!TRAAFP_COMOPE
      End If
      If Not IsNull(g_rst_Princi!TRAAFP_COMLEG) Then
         txt_ComLeg.Text = g_rst_Princi!TRAAFP_COMLEG
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
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
   
'   Call grd_LisCon_Click
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
      
      grd_LisExc.Col = 5
      If LCase(Trim(r_str_InsExc)) = LCase("EVALUACION CREDITICIA") Then
         pnl_motivo.Visible = True
         lbl_motivo.Visible = True
         pnl_motivo.Caption = IIf(grd_LisExc.Text = "0", " ", grd_LisExc.Text)
      Else
         pnl_motivo.Visible = False
         lbl_motivo.Visible = False
         pnl_motivo.Caption = ""
      End If
      
      Call gs_SetFocus(grd_LisExc)
      Call gs_RefrescaGrid(grd_LisExc)
   Else
      pnl_DesExc.Caption = ""
      txt_ObsExc.Text = ""
      pnl_TipAut.Caption = ""
      pnl_motivo.Caption = ""
   End If
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   
   Call grd_LisExc_Click
End Sub

Private Sub grd_LisOcu_AFP_Click()
   Dim r_str_FecOcu     As String
   Dim r_str_HorOcu     As String
   Dim r_str_DesOcu     As String

   If grd_LisOcu_AFP.Rows > 0 Then
      grd_LisOcu_AFP.Col = 0
      r_str_FecOcu = grd_LisOcu_AFP.Text
      
      grd_LisOcu_AFP.Col = 1
      r_str_HorOcu = grd_LisOcu_AFP.Text
      
      grd_LisOcu_AFP.Col = 2
      r_str_DesOcu = grd_LisOcu_AFP.Text
      
      pnl_DesOcu_AFP.Caption = "Día: " & r_str_FecOcu & " - " & r_str_HorOcu & " hrs. - " & r_str_DesOcu
   
      grd_LisOcu_AFP.Col = 3
      txt_Observ_AFP.Text = grd_LisOcu_AFP.Text
      
      grd_LisOcu_AFP.Col = 4
      txt_Descar_AFP.Text = grd_LisOcu_AFP.Text
      
      Call gs_RefrescaGrid(grd_LisOcu_AFP)
   End If
End Sub

Private Sub grd_LisOcu_AFP_SelChange()
   If grd_LisOcu_AFP.Rows > 2 Then
      grd_LisOcu_AFP.RowSel = grd_LisOcu_AFP.Row
   End If
   
'   Call grd_LisOcu_AFP_Click
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
   
'   Call grd_LisOcu_Click
End Sub

Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub txt_ComCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      If moddat_g_int_FlgCre = 2 Or moddat_g_int_FlgCre = 3 Then
         KeyAscii = 0
'      ElseIf moddat_g_int_TipRep = 2 Then
'         KeyAscii = 0
      Else
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
      End If
   End If
End Sub

Private Sub txt_ComLeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
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

Private Sub txt_Observ_AFP_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Descar_AFP_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_DatTas()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   grd_Listad(6).Redraw = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Empresa Peritaje"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Nombre Perito"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Código REPEV SBS"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Nro. de Informe"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      
      grd_Listad(6).Rows = grd_Listad(5).Rows + 1
      grd_Listad(6).Row = grd_Listad(5).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Fecha Evaluación"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Año de Construcción"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = CStr(g_rst_Princi!EVATAS_ANOCON)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Nro. de Pisos"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = CStr(g_rst_Princi!EVATAS_NUMPIS)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Nro. de Sótanos"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = CStr(g_rst_Princi!EVATAS_NUMSOT)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Tipo de Inmueble"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!EVATAS_TIPINM))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Uso de Inmueble"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("222", CStr(g_rst_Princi!EVATAS_USOINM))
      
      grd_Listad(6).Rows = grd_Listad(5).Rows + 1
      grd_Listad(6).Row = grd_Listad(5).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Material de Construcción"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("223", CStr(g_rst_Princi!EVATAS_MATCON))
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Tipo de Moneda"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!EVATAS_TIPMON))
      
      'Total
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Area Terreno (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Area Construida (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Suma Asegurada (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Valor Comercial (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Valor Realización (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Valor Terreno (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Valor Edificación (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).Text = "Valor Areas Comunes (Total)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
      'Inmueble
      grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Area Terreno (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2) & " m2"
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Area Construida (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2) & " m2"
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Suma Asegurada (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Valor Comercial (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Valor Realización (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Valor Terreno (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Valor Edificación (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).Text = "Valor Areas Comunes (Inmueble)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
      'Estacionamiento 1
      If g_rst_Princi!EVATAS_FLGEST_ES1 = 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Area Terreno (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2) & " m2"
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Area Construida (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2) & " m2"
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Suma Asegurada (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Comercial (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Realización (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Terreno (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Edificación (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Areas Comunes (Estac. 1)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_ES2 = 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Area Terreno (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2) & " m2"
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Area Construida (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2) & " m2"
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Suma Asegurada (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Valor Comercial (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Valor Realización (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Valor Terreno (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Valor Edificación (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).Text = "Valor Areas Comunes (Estac. 2)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_DEP = 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Area Terreno (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Area Construida (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Suma Asegurada (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Comercial (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Realización (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Terreno (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Edificación (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).Text = "Valor Areas Comunes (Depósito)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   grd_Listad(6).Redraw = True

   Call gs_UbiIniGrid(grd_Listad(6))
End Sub

Private Sub fs_DatSeg()
   Call gs_LimpiaGrid(grd_Listad(7))
   
   grd_Listad(7).Redraw = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_EVASEG "
   g_str_Parame = g_str_Parame & " WHERE EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Empresa de Seguros"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
   
      grd_Listad(7).Rows = grd_Listad(7).Rows + 2
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Tipo de Seguro Desgravamen"

      grd_Listad(7).Col = 1
      grd_Listad(7).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Fecha Evaluación (Seg. Desgravamen)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVADES))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Tipo de Valor (Seg. Desgravamen)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Valor a Aplicar"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 2
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Fecha Evaluación (Seg. Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Tipo de Valor (Seg. Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Valor a Aplicar"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad(7).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(7))
End Sub

Private Sub fs_EvaCre()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   'Obteniendo Ingreso Neto
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_EVACRE "
   g_str_Parame = g_str_Parame & " WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).CellForeColor = modgen_g_con_ColRoj
   grd_Listad(5).Text = "Total Ingreso Líquido Neto S/."
   
   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).CellForeColor = modgen_g_con_ColRoj
   grd_Listad(5).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVACRE_INGNET, 12, 2)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Cuota Aceptada
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Cuota (S/.)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_SOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Cuota (Moneda Prest.)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_MPR, 12, 2)

   If g_rst_Princi!SOLMAE_TIPMON <> 1 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Cambio"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_TCAMPR_APR, 14, 4)
   End If

   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")

   Call gs_UbiIniGrid(grd_Listad(5))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
