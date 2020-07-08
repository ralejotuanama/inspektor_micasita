VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_SegSol_51 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   1875
   ClientTop       =   1140
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_148.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   15372
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
         Height          =   3795
         Left            =   30
         TabIndex        =   1
         Top             =   4860
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   6694
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
            Height          =   3675
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   6482
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Seguimiento en Instancia"
            TabPicture(0)   =   "AteCli_frm_148.frx":000C
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
            TabCaption(1)   =   "Excepciones"
            TabPicture(1)   =   "AteCli_frm_148.frx":0028
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
            Begin VB.TextBox txt_Observ 
               Height          =   675
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "AteCli_frm_148.frx":0044
               Top             =   2220
               Width           =   10005
            End
            Begin VB.TextBox txt_Descar 
               Height          =   675
               Left            =   1320
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "AteCli_frm_148.frx":0048
               Top             =   2910
               Width           =   10005
            End
            Begin VB.TextBox txt_ObsExc 
               Height          =   1035
               Left            =   -73680
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "AteCli_frm_148.frx":004C
               Top             =   2220
               Width           =   10005
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   45
               Left            =   30
               TabIndex        =   6
               Top             =   1800
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisOcu 
               Height          =   1095
               Left            =   30
               TabIndex        =   7
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1931
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
               TabIndex        =   8
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
               TabIndex        =   9
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
               TabIndex        =   10
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
               TabIndex        =   11
               Top             =   1890
               Width           =   10005
               _Version        =   65536
               _ExtentX        =   17648
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
            Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
               Height          =   1095
               Left            =   -74970
               TabIndex        =   12
               Top             =   660
               Width           =   11325
               _ExtentX        =   19976
               _ExtentY        =   1931
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
               TabIndex        =   13
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
               TabIndex        =   14
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               TabIndex        =   17
               Top             =   1800
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
               TabIndex        =   18
               Top             =   1890
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
               Left            =   -73680
               TabIndex        =   19
               Top             =   3270
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
            Begin VB.Label Label7 
               Caption         =   "Comentario u Observación:"
               Height          =   495
               Left            =   60
               TabIndex        =   25
               Top             =   2220
               Width           =   1155
            End
            Begin VB.Label Label8 
               Caption         =   "Ocurrencia:"
               Height          =   315
               Left            =   60
               TabIndex        =   24
               Top             =   1890
               Width           =   1155
            End
            Begin VB.Label Label11 
               Caption         =   "Descargo:"
               Height          =   315
               Left            =   60
               TabIndex        =   23
               Top             =   2910
               Width           =   1035
            End
            Begin VB.Label Label2 
               Caption         =   "Autorizado por:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   22
               Top             =   3270
               Width           =   1095
            End
            Begin VB.Label Label3 
               Caption         =   "Excepción:"
               Height          =   315
               Left            =   -74940
               TabIndex        =   21
               Top             =   1890
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Descripción:"
               Height          =   495
               Left            =   -74940
               TabIndex        =   20
               Top             =   2220
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2895
         Left            =   30
         TabIndex        =   26
         Top             =   1920
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5106
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
            Height          =   2775
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4895
            _Version        =   393216
            Style           =   1
            Tabs            =   7
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "AteCli_frm_148.frx":0050
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "AteCli_frm_148.frx":006C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Patrimonio"
            TabPicture(2)   =   "AteCli_frm_148.frx":0088
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Referencias Personales"
            TabPicture(3)   =   "AteCli_frm_148.frx":00A4
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Inmueble"
            TabPicture(4)   =   "AteCli_frm_148.frx":00C0
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(2)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Datos del Crédito"
            TabPicture(5)   =   "AteCli_frm_148.frx":00DC
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "Label5"
            Tab(5).Control(1)=   "grd_Listad(5)"
            Tab(5).Control(2)=   "txt_ObsSol"
            Tab(5).ControlCount=   3
            TabCaption(6)   =   "Docum. Recibidos"
            TabPicture(6)   =   "AteCli_frm_148.frx":00F8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "grd_Listad(6)"
            Tab(6).ControlCount=   1
            Begin VB.TextBox txt_ObsSol 
               Height          =   675
               Left            =   -73650
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   28
               Text            =   "AteCli_frm_148.frx":0114
               Top             =   1980
               Width           =   10005
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2385
               Index           =   0
               Left            =   60
               TabIndex        =   29
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4207
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
               Height          =   2385
               Index           =   1
               Left            =   -74940
               TabIndex        =   30
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4207
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
               Height          =   2385
               Index           =   2
               Left            =   -74940
               TabIndex        =   31
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4207
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
               Height          =   2385
               Index           =   3
               Left            =   -74940
               TabIndex        =   32
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4207
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
               Height          =   2385
               Index           =   4
               Left            =   -74940
               TabIndex        =   33
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4207
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
               Height          =   1605
               Index           =   5
               Left            =   -74940
               TabIndex        =   34
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   2831
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
               Height          =   2385
               Index           =   6
               Left            =   -74940
               TabIndex        =   35
               Top             =   360
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4207
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label5 
               Caption         =   "Observaciones:"
               Height          =   495
               Left            =   -74910
               TabIndex        =   36
               Top             =   1980
               Width           =   1155
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   37
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
            Left            =   630
            TabIndex        =   38
            Top             =   30
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   660
            TabIndex        =   42
            Top             =   330
            Width           =   6405
            _Version        =   65536
            _ExtentX        =   11298
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Registro de Solicitudes"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10920
            Top             =   60
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
            Left            =   10290
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_148.frx":0118
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   39
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
            Picture         =   "AteCli_frm_148.frx":0422
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueObs 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_148.frx":0864
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Registro de Observación"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   43
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   767
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
            TabIndex        =   44
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   4440
            TabIndex        =   45
            Top             =   60
            Width           =   7035
            _Version        =   65536
            _ExtentX        =   12409
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   3690
            TabIndex        =   46
            Top             =   60
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frm_SegSol_51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_NueObs_Click()
   If moddat_g_int_NumObs = 0 Then
      MsgBox "No hay observación pendiente de descargo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_CodIns = 11
   moddat_g_str_DesObs = ""
   
   moddat_g_int_FlgAct_1 = 1
   
   frm_SegSol_52.Show 1
   
   If moddat_g_int_FlgAct_1 = 2 Then
      If Not moddat_gf_Modifica_SegDet_Observ(moddat_g_str_NumSol, moddat_g_int_CodIns, 21, CStr(moddat_g_int_NumObs), moddat_g_str_DesObs, 2) Then
         Exit Sub
      End If
   
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, moddat_g_int_CodIns, 0, 9, 2) Then
         Exit Sub
      End If
   
      'Enviando Correo Electrónico
      modgen_g_str_Mail_Asunto = moddat_gf_Consulta_ParDes("002", CStr(moddat_g_int_CodIns)) & " - DESCARGO DE OBSERVACION " & "(" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
      modgen_g_str_Mail_Mensaj = ""
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & gf_Formato_NumSol(moddat_g_str_NumSol) & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & moddat_g_str_DesObs
   
      Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
      Screen.MousePointer = 11
      Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
      Screen.MousePointer = 0
      
      moddat_g_int_FlgAct = 2
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   
   'Buscar Información de Solicitud de Crédito
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""

   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)       'Buscar Información del Cliente
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)       'Buscar Información del Cónyuge
   
   Call fs_DatPat          'Datos del Patrimonio
   Call fs_DatRef          'Referencias Personales
   Call fs_DatInm          'Datos del Inmueble
   Call fs_DatCre          'Datos del Crédito
   Call fs_SolDoc          'Documentos Recibidos
   
   Call fs_Buscar_LisOcu      'Buscando Ocurrencias de Instancia
   Call fs_Buscar_LisExc      'Buscando Excepciones
   
   'Si no hay Excepciones aplicadas
   If grd_LisExc.Rows = 0 Then
      tab_Seguim.TabVisible(1) = False
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer

   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 5
      grd_Listad(r_int_Contad).ColWidth(0) = 2900
      grd_Listad(r_int_Contad).ColWidth(1) = 7950
   
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
   
   grd_Listad(6).ColWidth(0) = 10850
   grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter

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
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento de Identidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Apellidos y Nombres"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If g_rst_Princi!DatGen_FLGDOA = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Documento Adicional de Identidad"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      End If
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Sexo"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Fecha de Nacimiento"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nacionalidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Lugar de Nacimiento"
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      Else
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = "<< NO REGISTRADO >>"
      End If
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Estado Civil"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nivel de Estudios"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Profesión"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Celular"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Nro. Dependientes Económicos"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = CStr(g_rst_Princi!DatGen_DepEco)
      
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Edades"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "E-mail"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Autorización Envío"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Domicilio"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Teléfono Domicilio"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Económica Principal
   Call fs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice)
   Call fs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice)
End Sub

Private Sub fs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer)
   Dim r_var_ColTxt
   
   If p_OrdAct = 1 Then
      r_var_ColTxt = modgen_g_con_ColAzu
   Else
      r_var_ColTxt = modgen_g_con_ColRoj
   End If

   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 2
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = "Ocupación " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ActEco_CodAct)
      
      Select Case g_rst_Princi!ActEco_CodAct
         Case 11: Call fs_ActEco_Dep(p_Indice, r_var_ColTxt)
         Case 21: Call fs_ActEco_Ind(p_Indice, r_var_ColTxt)
         Case 31: Call fs_ActEco_Com(p_Indice, r_var_ColTxt)
         Case 41: Call fs_ActEco_Acc(p_Indice, r_var_ColTxt)
         Case 51: Call fs_ActEco_Ren(p_Indice, r_var_ColTxt)
         Case 61: Call fs_ActEco_Otr(p_Indice, r_var_ColTxt)
      End Select
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_ActEco_Dep(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad Empleador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Situación como Trabajador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Dep_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Dep_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")) > 0, " ANEXO: " & Trim(g_rst_Princi!ActEco_Dep_AnexRH & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELERH & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0, " ANEXO: " & Trim(g_rst_Genera!DATGEN_ANEXRH & ""), "")

      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
      Else
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                     " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Frecuencia de Haberes"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("210", CStr(g_rst_Princi!ActEco_Dep_FreHab))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Ingreso"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Area"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Anexo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono Directo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Celular"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Celula & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "E-mail"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")

   If g_rst_Princi!ActEco_Dep_TraAnt = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador Anterior"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc_Ant & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc_Ant & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s) (Empleador Anterior)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Cese (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
   End If
End Sub

Private Sub fs_ActEco_Ind(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Indartamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Actividades"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Contrato de Locación de Servicios"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Cargo"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub fs_ActEco_Com(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Razón Social"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Giro Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ventas Mensuales"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_VtaMen, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Operaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Régimen Tributario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("215", CStr(g_rst_Princi!ActEco_Com_RegTri))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_PorPar, 7, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tipo de Local Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("208", CStr(g_rst_Princi!ActEco_Com_TipLoc))
   
   If g_rst_Princi!ActEco_Com_TipLoc = 2 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_AlqMen, 15, 2)
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
   End If
End Sub

Private Sub fs_ActEco_Acc(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Acc_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Acc_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Acc_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Acc_NomVia) & " " & Trim(g_rst_Princi!ActEco_Acc_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Acc_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Acc_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Acc_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Acc_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                  " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Antigüedad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_PorPar, 7, 2)
End Sub

Private Sub fs_ActEco_Ren(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección de Propiedad 01"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Alquiler Mensual"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección de Propiedad 02"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub fs_ActEco_Otr(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Otr_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Actividad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Activi & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Otr_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Otr_CodCiu))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Observaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
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

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(2).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(2).Col = 1
         grd_Listad(2).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Tipo de Inmueble"
         
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Dirección"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Referencia"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Estacionamiento"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Departamento / Provincia / Distrito"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 2
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Proyecto miCasita"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Nombre Proyecto"
   
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Nombre Proyecto"
   
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 2
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Propietario / Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Nombre o Razón Social"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Dirección"
         
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Referencia"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Teléfono"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Docum. Identidad Constructor"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Nombre o Razón Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Dirección"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Teléfono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto Vinculado"
            Else
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Entidad Financiera"
         
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Docum. Identidad Propietario"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Nombre o Razón Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Dirección"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Teléfono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Doc. Ident. Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Razón Social Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Doc. Ident. Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Razón Social Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(2).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(2))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatRef()
   Dim r_var_ColTxt

   r_var_ColTxt = modgen_g_con_ColNeg

   Call gs_LimpiaGrid(grd_Listad(3))

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SOLREF_TIPREF ASC, SOLREF_NUMREF ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(3).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If grd_Listad(3).Rows = 0 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
         End If
         
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Referencia"
            
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("010", CStr(g_rst_Princi!SOLREF_TIPREF))
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Parentesco"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         
         If g_rst_Princi!SOLREF_TIPREF = 1 Then
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
         Else
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
         End If
         
         If g_rst_Princi!SOLREF_TIPPAR <> 8 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).CellForeColor = r_var_ColTxt
            grd_Listad(3).Text = "Apellidos y Nombres"
      
            grd_Listad(3).Col = 1
            grd_Listad(3).CellForeColor = r_var_ColTxt
            grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
         
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).CellForeColor = r_var_ColTxt
            grd_Listad(3).Text = "Teléfono"
   
            grd_Listad(3).Col = 1
            grd_Listad(3).CellForeColor = r_var_ColTxt
            grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
         
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).CellForeColor = r_var_ColTxt
            grd_Listad(3).Text = "Celular"
      
            grd_Listad(3).Col = 1
            grd_Listad(3).CellForeColor = r_var_ColTxt
            grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
         End If
   
         g_rst_Princi.MoveNext
         
         If r_var_ColTxt = modgen_g_con_ColNeg Then
            r_var_ColTxt = modgen_g_con_ColAzu
         Else
            r_var_ColTxt = modgen_g_con_ColNeg
         End If
      Loop
      
      grd_Listad(3).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(3))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
   'Mostrar Todos los Documentos Recibidos
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
   
   grd_Listad(4).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Fecha de Adquisición (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Dirección (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Número de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Línea Crédito (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Pago Mínimo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Número de Operación (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Monto del Préstamo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = CStr(g_rst_Genera!SOLDEU_PLAMEN)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
   End If
   
   grd_Listad(4).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   'Mostrar Todos los Documentos Recibidos
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
   
   grd_Listad(5).Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Producto"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_g_str_NomPrd
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Sub-Producto"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tipo de Evaluación"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Moneda del Préstamo"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Fecha de Solicitud"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))

   If g_rst_Princi!SOLMAE_COMVTA_MON > 0 Then
      If g_rst_Princi!SOLMAE_TIPMON = 2 Then
         grd_Listad(5).Rows = grd_Listad(5).Rows + 2
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Valor de Compra Venta"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Aporte Propio"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Monto Préstamo"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
      Else
         grd_Listad(5).Rows = grd_Listad(5).Rows + 2
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Valor de Compra Venta"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Aporte Propio"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
      
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Monto Préstamo"
      
         grd_Listad(5).Col = 1
         grd_Listad(5).CellFontName = "Lucida Console"
         grd_Listad(5).CellFontSize = 8
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
      End If
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 2
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tasa de Interés"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Plazo"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PLAANO) & " Años"
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Período de Gracia (Meses)"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PERGRA) & " Meses"
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Cuotas Extraordinarias"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Compañía de Seguros"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Seguro Desgravamen"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Día de Pago"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   End If
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 2
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Institución Financiera de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Mínimo de Ahorro Mensual"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_MONAHO)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Meses Ahorrados"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
   End If
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 2
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Consejero Hipotecario"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Ejecutivo de Seguimiento"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   grd_Listad(5).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(5))
   
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   moddat_g_str_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   txt_ObsSol.Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(6).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
   
      grd_Listad(6).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(6).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(6).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(6))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub txt_Descar_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsEva_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Buscar_LisOcu()
   Call gs_LimpiaGrid(grd_LisOcu)
   
   moddat_g_int_NumObs = 0
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = 11    "
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
      
      'Número de Observación
      'grd_LisOcu.Col = 0
      'grd_LisOcu.Text = Format(g_rst_Princi!SEGDET_NUMOBS, "000")
      
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

Private Sub txt_ObsExc_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsSol_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodConHip)
   
   'Evaluador de Creditos
   r_str_Cadena = moddat_gf_UsuObs(moddat_g_str_NumSol, 11)
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
   
   'Director General
   'r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(10)
   'If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
   '   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   '   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   'End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub


