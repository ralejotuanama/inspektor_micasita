VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_MntCli_67 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   Icon            =   "AteCli_frm_579.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel111 
      Height          =   8625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10665
      _Version        =   65536
      _ExtentX        =   18812
      _ExtentY        =   15214
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
         Height          =   2595
         Left            =   30
         TabIndex        =   32
         Top             =   1380
         Width           =   10605
         _Version        =   65536
         _ExtentX        =   18706
         _ExtentY        =   4577
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
         Begin TabDlg.SSTab Tab_Deuda 
            Height          =   2415
            Left            =   60
            TabIndex        =   66
            Top             =   90
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   4260
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Datos Cliente"
            TabPicture(0)   =   "AteCli_frm_579.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Operaciones"
            TabPicture(1)   =   "AteCli_frm_579.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_LisOpe"
            Tab(1).Control(1)=   "pnl_Tit_NumOpe"
            Tab(1).Control(2)=   "pnl_Tit_Produc"
            Tab(1).Control(3)=   "SSPanel3"
            Tab(1).ControlCount=   4
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Left            =   120
               TabIndex        =   67
               Top             =   480
               Width           =   10245
               _ExtentX        =   18071
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisOpe 
               Height          =   1515
               Left            =   -74880
               TabIndex        =   68
               Top             =   780
               Width           =   10215
               _ExtentX        =   18018
               _ExtentY        =   2672
               _Version        =   393216
               Rows            =   6
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Tit_NumOpe 
               Height          =   285
               Left            =   -74850
               TabIndex        =   69
               Top             =   480
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Nro. Operación"
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
            Begin Threed.SSPanel pnl_Tit_Produc 
               Height          =   285
               Left            =   -72900
               TabIndex        =   70
               Top             =   480
               Width           =   5700
               _Version        =   65536
               _ExtentX        =   10054
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Producto"
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
               Left            =   -67320
               TabIndex        =   71
               Top             =   480
               Width           =   1995
               _Version        =   65536
               _ExtentX        =   3528
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Estado"
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
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   10605
         _Version        =   65536
         _ExtentX        =   18706
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
            Height          =   555
            Left            =   690
            TabIndex        =   34
            Top             =   30
            Width           =   5565
            _Version        =   65536
            _ExtentX        =   9816
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Actualización de Datos del Cliente"
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
            Picture         =   "AteCli_frm_579.frx":0044
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   645
         Left            =   30
         TabIndex        =   35
         Top             =   720
         Width           =   10605
         _Version        =   65536
         _ExtentX        =   18706
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
            Left            =   9990
            Picture         =   "AteCli_frm_579.frx":034E
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4575
         Left            =   30
         TabIndex        =   37
         Top             =   4020
         Width           =   10605
         _Version        =   65536
         _ExtentX        =   18706
         _ExtentY        =   8070
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
         Begin TabDlg.SSTab Tab_DatCli 
            Height          =   4395
            Index           =   1
            Left            =   60
            TabIndex        =   38
            Top             =   90
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   7752
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Direcciones"
            TabPicture(0)   =   "AteCli_frm_579.frx":0790
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel5"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel9"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "grd_LisDir"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Teléfonos"
            TabPicture(1)   =   "AteCli_frm_579.frx":07AC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_LisTel"
            Tab(1).Control(1)=   "SSPanel8"
            Tab(1).Control(2)=   "SSPanel11"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Emails"
            TabPicture(2)   =   "AteCli_frm_579.frx":07C8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_LisEml"
            Tab(2).Control(1)=   "SSPanel10"
            Tab(2).Control(2)=   "SSPanel12"
            Tab(2).ControlCount=   3
            Begin MSFlexGridLib.MSFlexGrid grd_LisDir 
               Height          =   1305
               Left            =   120
               TabIndex        =   39
               Top             =   480
               Width           =   10230
               _ExtentX        =   18045
               _ExtentY        =   2302
               _Version        =   393216
               Rows            =   10
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisTel 
               Height          =   1305
               Left            =   -74880
               TabIndex        =   40
               Top             =   480
               Width           =   10230
               _ExtentX        =   18045
               _ExtentY        =   2302
               _Version        =   393216
               Rows            =   10
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisEml 
               Height          =   1305
               Left            =   -74880
               TabIndex        =   41
               Top             =   480
               Width           =   10230
               _ExtentX        =   18045
               _ExtentY        =   2302
               _Version        =   393216
               Rows            =   10
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   1755
               Left            =   120
               TabIndex        =   42
               Top             =   2520
               Width           =   10245
               _Version        =   65536
               _ExtentX        =   18071
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
               Begin VB.ComboBox cmb_TipVia 
                  Height          =   315
                  Left            =   960
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   210
                  Width           =   1905
               End
               Begin VB.TextBox txt_NomVia 
                  Height          =   315
                  Left            =   4110
                  MaxLength       =   120
                  TabIndex        =   3
                  Text            =   "txt_NomVia"
                  Top             =   210
                  Width           =   2085
               End
               Begin VB.TextBox txt_NumVia 
                  Height          =   315
                  Left            =   7650
                  MaxLength       =   30
                  TabIndex        =   4
                  Text            =   "txt_NumVia"
                  Top             =   210
                  Width           =   1185
               End
               Begin VB.TextBox txt_IntDpt 
                  Height          =   315
                  Left            =   8850
                  MaxLength       =   30
                  TabIndex        =   5
                  Text            =   "txt_IntDpt"
                  Top             =   210
                  Width           =   1215
               End
               Begin VB.ComboBox cmb_TipZon 
                  Height          =   315
                  Left            =   960
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   540
                  Width           =   1905
               End
               Begin VB.TextBox txt_NomZon 
                  Height          =   315
                  Left            =   4110
                  MaxLength       =   120
                  TabIndex        =   7
                  Text            =   "txt_NomZon"
                  Top             =   540
                  Width           =   2085
               End
               Begin VB.ComboBox cmb_DptDir 
                  Height          =   315
                  Left            =   7650
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   540
                  Width           =   2415
               End
               Begin VB.ComboBox cmb_PrvDir 
                  Height          =   315
                  Left            =   960
                  TabIndex        =   9
                  Text            =   "cmb_PrvDir"
                  Top             =   870
                  Width           =   1905
               End
               Begin VB.ComboBox cmb_DstDir 
                  Height          =   315
                  Left            =   4110
                  TabIndex        =   10
                  Text            =   "cmb_DstDir"
                  Top             =   870
                  Width           =   2085
               End
               Begin VB.TextBox txt_Refere 
                  Height          =   315
                  Left            =   7650
                  MaxLength       =   250
                  TabIndex        =   11
                  Text            =   "txt_Refere"
                  Top             =   870
                  Width           =   2415
               End
               Begin VB.TextBox txt_Coment 
                  Height          =   345
                  Index           =   0
                  Left            =   4110
                  MaxLength       =   2000
                  TabIndex        =   13
                  Top             =   1200
                  Width           =   5955
               End
               Begin VB.ComboBox cmb_Estado 
                  Height          =   315
                  Index           =   0
                  Left            =   960
                  Style           =   2  'Dropdown List
                  TabIndex        =   12
                  Top             =   1200
                  Width           =   1905
               End
               Begin VB.Label Label19 
                  Caption         =   "Tipo Vía:"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   53
                  Top             =   255
                  Width           =   885
               End
               Begin VB.Label Label20 
                  Caption         =   "Nombre Vía:"
                  Height          =   285
                  Left            =   3030
                  TabIndex        =   52
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.Label Label21 
                  Caption         =   "Nro - Int/Dpt/Mz:"
                  Height          =   285
                  Left            =   6360
                  TabIndex        =   51
                  Top             =   255
                  Width           =   1320
               End
               Begin VB.Label Label22 
                  Caption         =   "Tipo Zona:"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   50
                  Top             =   570
                  Width           =   885
               End
               Begin VB.Label Label23 
                  Caption         =   "Nombre Zona:"
                  Height          =   285
                  Left            =   3030
                  TabIndex        =   49
                  Top             =   570
                  Width           =   1095
               End
               Begin VB.Label Label24 
                  Caption         =   "Departamento:"
                  Height          =   315
                  Left            =   6360
                  TabIndex        =   48
                  Top             =   570
                  Width           =   1320
               End
               Begin VB.Label Label25 
                  Caption         =   "Provincia:"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   47
                  Top             =   915
                  Width           =   885
               End
               Begin VB.Label Label26 
                  Caption         =   "Distrito:"
                  Height          =   315
                  Left            =   3030
                  TabIndex        =   46
                  Top             =   915
                  Width           =   1095
               End
               Begin VB.Label Label28 
                  Caption         =   "Referencia:"
                  Height          =   285
                  Left            =   6360
                  TabIndex        =   45
                  Top             =   915
                  Width           =   1605
               End
               Begin VB.Label Label5 
                  Caption         =   "Comentarios:"
                  Height          =   315
                  Left            =   3030
                  TabIndex        =   44
                  Top             =   1245
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Estado:"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   43
                  Top             =   1245
                  Width           =   885
               End
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   675
               Left            =   120
               TabIndex        =   54
               Top             =   1800
               Width           =   10245
               _Version        =   65536
               _ExtentX        =   18071
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
               Begin VB.CommandButton cmd_Grabar 
                  Height          =   585
                  Index           =   0
                  Left            =   9030
                  Picture         =   "AteCli_frm_579.frx":07E4
                  Style           =   1  'Graphical
                  TabIndex        =   14
                  ToolTipText     =   "Grabar Datos"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Editar 
                  Height          =   585
                  Index           =   0
                  Left            =   8430
                  Picture         =   "AteCli_frm_579.frx":0C26
                  Style           =   1  'Graphical
                  TabIndex        =   15
                  ToolTipText     =   "Modificar Registro"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Cancel 
                  Height          =   585
                  Index           =   0
                  Left            =   9660
                  Picture         =   "AteCli_frm_579.frx":0F30
                  Style           =   1  'Graphical
                  TabIndex        =   16
                  ToolTipText     =   "Cancelar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Agrega 
                  Height          =   585
                  Index           =   0
                  Left            =   7830
                  Picture         =   "AteCli_frm_579.frx":123A
                  Style           =   1  'Graphical
                  TabIndex        =   1
                  ToolTipText     =   "Nuevo Registro"
                  Top             =   30
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   675
               Left            =   -74880
               TabIndex        =   55
               Top             =   1800
               Width           =   10245
               _Version        =   65536
               _ExtentX        =   18071
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
               Begin VB.CommandButton cmd_Agrega 
                  Height          =   585
                  Index           =   1
                  Left            =   7830
                  Picture         =   "AteCli_frm_579.frx":1544
                  Style           =   1  'Graphical
                  TabIndex        =   17
                  ToolTipText     =   "Nuevo Registro"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Cancel 
                  Height          =   585
                  Index           =   1
                  Left            =   9630
                  Picture         =   "AteCli_frm_579.frx":184E
                  Style           =   1  'Graphical
                  TabIndex        =   24
                  ToolTipText     =   "Cancelar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Editar 
                  Height          =   585
                  Index           =   1
                  Left            =   8430
                  Picture         =   "AteCli_frm_579.frx":1B58
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  ToolTipText     =   "Modificar Registro"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Grabar 
                  Height          =   585
                  Index           =   1
                  Left            =   9030
                  Picture         =   "AteCli_frm_579.frx":1E62
                  Style           =   1  'Graphical
                  TabIndex        =   22
                  ToolTipText     =   "Grabar Datos"
                  Top             =   30
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   675
               Left            =   -74880
               TabIndex        =   56
               Top             =   1800
               Width           =   10275
               _Version        =   65536
               _ExtentX        =   18124
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
               Begin VB.CommandButton cmd_Grabar 
                  Height          =   585
                  Index           =   2
                  Left            =   9060
                  Picture         =   "AteCli_frm_579.frx":22A4
                  Style           =   1  'Graphical
                  TabIndex        =   29
                  ToolTipText     =   "Grabar Datos"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Editar 
                  Height          =   585
                  Index           =   2
                  Left            =   8430
                  Picture         =   "AteCli_frm_579.frx":26E6
                  Style           =   1  'Graphical
                  TabIndex        =   30
                  ToolTipText     =   "Modificar Registro"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Cancel 
                  Height          =   585
                  Index           =   2
                  Left            =   9630
                  Picture         =   "AteCli_frm_579.frx":29F0
                  Style           =   1  'Graphical
                  TabIndex        =   31
                  ToolTipText     =   "Cancelar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Agrega 
                  Height          =   585
                  Index           =   2
                  Left            =   7830
                  Picture         =   "AteCli_frm_579.frx":2CFA
                  Style           =   1  'Graphical
                  TabIndex        =   25
                  ToolTipText     =   "Nuevo Registro"
                  Top             =   30
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   1755
               Left            =   -74880
               TabIndex        =   57
               Top             =   2520
               Width           =   10245
               _Version        =   65536
               _ExtentX        =   18071
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
               Begin VB.TextBox txt_NumTel 
                  Height          =   315
                  Left            =   1410
                  MaxLength       =   25
                  TabIndex        =   19
                  Text            =   "txt_NumTel"
                  Top             =   540
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_Estado 
                  Height          =   315
                  Index           =   1
                  Left            =   1410
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   870
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_TipTel 
                  Height          =   315
                  Left            =   1410
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   210
                  Width           =   2715
               End
               Begin VB.TextBox txt_Coment 
                  Height          =   345
                  Index           =   1
                  Left            =   1410
                  MaxLength       =   2000
                  TabIndex        =   21
                  Top             =   1200
                  Width           =   8535
               End
               Begin VB.Label Label16 
                  Caption         =   "Teléfono:"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   61
                  Top             =   540
                  Width           =   1245
               End
               Begin VB.Label Label8 
                  Caption         =   "Estado:"
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  TabIndex        =   60
                  Top             =   870
                  Width           =   1245
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo:"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   59
                  Top             =   210
                  Width           =   1245
               End
               Begin VB.Label Label3 
                  Caption         =   "Comentarios:"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   58
                  Top             =   1200
                  Width           =   1245
               End
            End
            Begin Threed.SSPanel SSPanel12 
               Height          =   1755
               Left            =   -74880
               TabIndex        =   62
               Top             =   2520
               Width           =   10275
               _Version        =   65536
               _ExtentX        =   18124
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
               Begin VB.TextBox txt_Coment 
                  Height          =   345
                  Index           =   2
                  Left            =   1380
                  MaxLength       =   2000
                  TabIndex        =   28
                  Top             =   870
                  Width           =   8685
               End
               Begin VB.ComboBox cmb_Estado 
                  Height          =   315
                  Index           =   2
                  Left            =   1380
                  Style           =   2  'Dropdown List
                  TabIndex        =   27
                  Top             =   540
                  Width           =   2355
               End
               Begin VB.TextBox txt_DirEle 
                  Height          =   315
                  Left            =   1380
                  MaxLength       =   120
                  TabIndex        =   26
                  Text            =   "txt_DirEle"
                  Top             =   210
                  Width           =   5505
               End
               Begin VB.Label Label4 
                  Caption         =   "Comentarios:"
                  Height          =   315
                  Left            =   150
                  TabIndex        =   65
                  Top             =   900
                  Width           =   1365
               End
               Begin VB.Label Label7 
                  Caption         =   "Estado:"
                  Height          =   315
                  Left            =   150
                  TabIndex        =   64
                  Top             =   570
                  Width           =   1365
               End
               Begin VB.Label Label17 
                  Caption         =   "E-mail:"
                  Height          =   285
                  Left            =   150
                  TabIndex        =   63
                  Top             =   240
                  Width           =   1365
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_67"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_Contad  As Integer
Dim l_int_Correl  As Integer
Dim l_str_PrvDir  As String
Dim l_str_DptDir  As String
Dim l_int_FlgCmb  As Integer
Dim l_str_DstDir  As String

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_Estado_Click(Index As Integer)
   Call gs_SetFocus(txt_Coment(Index))
End Sub

Private Sub cmb_Estado_KeyPress(Index As Integer, KeyAscii As Integer)
   Call cmb_Estado_Click(Index)
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_TipTel_Click()
   Call gs_SetFocus(txt_NumTel)
End Sub

Private Sub cmb_TipTel_KeyPress(KeyAscii As Integer)
   Call cmb_TipTel_Click
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
End Sub

Private Sub cmd_Agrega_Click(Index As Integer)
   Call fs_Activa(True, Index)
   Select Case Index
      Case 0:   moddat_g_int_FlgGrb = 1: Call gs_SetFocus(cmb_TipVia)
      Case 1:   moddat_g_int_FlgGrb_1 = 1: Call gs_SetFocus(cmb_TipTel)
      Case 2:   moddat_g_int_FlgGrb_2 = 1: Call gs_SetFocus(txt_DirEle)
   End Select
End Sub

Private Sub cmd_Cancel_Click(Index As Integer)
   Call fs_Activa(False, Index)
   Call fs_Limpia(Index)
   
   Select Case Index
      Case 0:
         Call gs_SetFocus(grd_LisDir)
      
         If grd_LisDir.Rows = 0 Then
            cmd_Agrega(Index).Enabled = True
            cmd_Editar(Index).Enabled = False
            grd_LisDir.Enabled = False
         End If
      Case 1:
         Call gs_SetFocus(grd_LisTel)
      
         If grd_LisTel.Rows = 0 Then
            cmd_Agrega(Index).Enabled = True
            cmd_Editar(Index).Enabled = False
            grd_LisTel.Enabled = False
         End If
      Case 2:
         Call gs_SetFocus(grd_LisEml)
      
         If grd_LisEml.Rows = 0 Then
            cmd_Agrega(Index).Enabled = True
            cmd_Editar(Index).Enabled = False
            grd_LisEml.Enabled = False
         End If
   End Select
End Sub

Private Sub cmd_Editar_Click(Index As Integer)
   Call fs_Activa(True, Index)
   Select Case Index
      Case 0:   moddat_g_int_FlgGrb = 2: grd_LisDir_DblClick: Call gs_SetFocus(cmb_TipVia)
      Case 1:   moddat_g_int_FlgGrb_1 = 2: grd_LisTel_DblClick: Call gs_SetFocus(cmb_TipTel)
      Case 2:   moddat_g_int_FlgGrb_2 = 2: grd_LisEml_DblClick: Call gs_SetFocus(txt_DirEle)
   End Select
End Sub

Private Sub cmd_Grabar_Click(Index As Integer)
   'Call moddat_gs_FecSis
   Select Case Index
      Case 0:
         If cmb_TipVia.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipVia)
            Exit Sub
         End If
         If Len(Trim(txt_NomVia.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomVia)
            Exit Sub
         End If
         If Len(Trim(txt_NumVia.Text)) = 0 Then
            MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumVia)
            Exit Sub
         End If
         If cmb_TipZon.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipZon)
            Exit Sub
         End If
         If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
            If Len(Trim(txt_NomZon.Text)) = 0 Then
               MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_NomZon)
               Exit Sub
            End If
         End If
         If cmb_DptDir.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_DptDir)
            Exit Sub
         End If
         If cmb_PrvDir.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_PrvDir)
            Exit Sub
         End If
         If cmb_DstDir.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_DstDir)
            Exit Sub
         End If
         Tab_DatCli(1).Tab = 0
         
      Case 1:
         If cmb_TipTel.ListIndex = -1 Then
            MsgBox "Debe ingresar el Tipo de Teléfono.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipTel)
            Exit Sub
         End If
         If Len(Trim(txt_NumTel.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Teléfono.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumTel)
            Exit Sub
         End If
         If cmb_Estado(Index).ListIndex = -1 Then
            MsgBox "Debe ingresar el estado del Teléfono.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Estado(Index))
            Exit Sub
         End If
         Tab_DatCli(1).Tab = 1
         
      Case 2:
         If Len(Trim(txt_DirEle.Text)) = 0 Then
            MsgBox "Debe ingresar el E-mail del cliente.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_DirEle)
            Exit Sub
         End If
         If Not gf_ValidarEmail(txt_DirEle.Text) Then
            MsgBox "El E-mail del cliente no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_DirEle)
            Exit Sub
         End If
         If cmb_Estado(Index).ListIndex = -1 Then
            MsgBox "Debe ingresar el estado del Email.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Estado(Index))
            Exit Sub
         End If
         Tab_DatCli(1).Tab = 2
   End Select
      
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Obtiene Correlativo
   g_str_Parame = ""
   Select Case Index
      Case 0:  g_str_Parame = g_str_Parame & " SELECT MAX(ADCDIR_CORREL) AS CORRELATIVO "
               g_str_Parame = g_str_Parame & "   FROM CLI_ADCDIR  "
               g_str_Parame = g_str_Parame & "  WHERE ADCDIR_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND ADCDIR_NUMDOC = '" & moddat_g_str_NumDoc & "' "
               
      Case 1:  g_str_Parame = g_str_Parame & " SELECT MAX(ADCTEL_CORREL) AS CORRELATIVO "
               g_str_Parame = g_str_Parame & "   FROM CLI_ADCTEL  "
               g_str_Parame = g_str_Parame & "  WHERE ADCTEL_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND ADCTEL_NUMDOC = '" & moddat_g_str_NumDoc & "' "

      Case 2:  g_str_Parame = g_str_Parame & " SELECT MAX(ADCEML_CORREL) AS CORRELATIVO "
               g_str_Parame = g_str_Parame & "   FROM CLI_ADCEML  "
               g_str_Parame = g_str_Parame & "  WHERE ADCEML_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND ADCEML_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   End Select

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not g_rst_Genera.BOF And Not g_rst_Genera.EOF Then
      l_int_Correl = IIf(IsNull(g_rst_Genera!CORRELATIVO), 0, g_rst_Genera!CORRELATIVO) + 1
   End If
   
   'Grabando Información del Cliente
   g_str_Parame = "USP_CLI_ACTUALIZA_DATOS ("
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgCre) & ", "

   Select Case Index
      Case 0:
            If moddat_g_int_FlgGrb = 1 Then
               g_str_Parame = g_str_Parame & CStr(l_int_Correl) & ", "
            Else
               g_str_Parame = g_str_Parame & CStr(moddat_g_str_Codigo) & ", "
            End If
            g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
            g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
            g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
            g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
            g_str_Parame = g_str_Parame & CStr(cmb_Estado(Index).ItemData(cmb_Estado(Index).ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_Coment(Index).Text) & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & moddat_g_int_FlgGrb & ", "
      Case 1:
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            If moddat_g_int_FlgGrb_1 = 1 Then
               g_str_Parame = g_str_Parame & CStr(l_int_Correl) & ", "
            Else
               g_str_Parame = g_str_Parame & CStr(moddat_g_str_Codigo) & ", "
            End If
            g_str_Parame = g_str_Parame & CStr(cmb_TipTel.ItemData(cmb_TipTel.ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NumTel.Text & "', "
            g_str_Parame = g_str_Parame & CStr(cmb_Estado(Index).ItemData(cmb_Estado(Index).ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_Coment(Index).Text) & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & moddat_g_int_FlgGrb_1 & ", "
      Case 2:
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            If moddat_g_int_FlgGrb_2 = 1 Then
               g_str_Parame = g_str_Parame & CStr(l_int_Correl) & ", "
            Else
               g_str_Parame = g_str_Parame & CStr(moddat_g_str_Codigo) & ", "
            End If
            g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
            g_str_Parame = g_str_Parame & CStr(cmb_Estado(Index).ItemData(cmb_Estado(Index).ListIndex)) & ", "
            g_str_Parame = g_str_Parame & "'" & Trim(txt_Coment(Index).Text) & "', "
            g_str_Parame = g_str_Parame & moddat_g_int_FlgGrb_2 & ", "
   End Select
   g_str_Parame = g_str_Parame & Index & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTUALIZA_DATOS.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct_1 = 2
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Select Case Index
      Case 0: Call gs_LimpiaGrid(grd_LisDir): Call fs_DirCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      Case 1: Call gs_LimpiaGrid(grd_LisTel): Call fs_TelCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
      Case 2: Call gs_LimpiaGrid(grd_LisEml): Call fs_EmlCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   End Select
   Call fs_Limpia(Index)
   Call fs_Activa(False, Index)
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicia
   For l_int_Contad = 0 To 2
      Call fs_Activa(False, l_int_Contad)
      Call fs_Limpia(l_int_Contad)
   Next l_int_Contad
   
   If Me.grd_LisDir.Row = 0 Then
      cmd_Editar(0).Enabled = True
   Else
      cmd_Editar(0).Enabled = False
   End If
   If Me.grd_LisTel.Row = 0 Then
      cmd_Editar(1).Enabled = True
   Else
      cmd_Editar(1).Enabled = False
   End If
   If Me.grd_LisEml.Row = 0 Then
      cmd_Editar(2).Enabled = True
   Else
      cmd_Editar(2).Enabled = False
   End If
   
   'Buscar Información del Cliente
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   'Datos del Cliente
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad, 0)      'Buscar Información del Cliente
   
   'Operaciones del Cliente
   Call fs_DatOpe(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Direcciones del Cliente
   Call fs_DirCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Teléfonos del Cliente
   Call fs_TelCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Emails del Cliente
   Call fs_EmlCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
'   If moddat_g_int_FlgGrb = 1 Then
'      Call fs_Activa(True)
'      cmd_Cancel.Enabled = False
'      Call fs_Limpia
'   Else
'      Call fs_Limpia
'      Call fs_Cargar_Datos
'      Call fs_Activa(False)
'   End If

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Datos del Cliente
   grd_Listad.ColWidth(0) = 2900
   grd_Listad.ColWidth(1) = 7940
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   'Operaciones del Cliente
   grd_LisOpe.ColWidth(0) = 2000
   grd_LisOpe.ColWidth(1) = 5650
   grd_LisOpe.ColWidth(2) = 1850
   grd_LisOpe.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(1) = flexAlignLeftCenter
   grd_LisOpe.ColAlignment(2) = flexAlignCenterCenter
   
   'Direcciones
   grd_LisDir.ColWidth(0) = 2000
   grd_LisDir.ColWidth(1) = 4800
   grd_LisDir.ColWidth(2) = 2000
   grd_LisDir.ColWidth(3) = 2000
   grd_LisDir.ColWidth(4) = 2000
   grd_LisDir.ColWidth(5) = 2000
   grd_LisDir.ColWidth(6) = 0
   grd_LisDir.ColAlignment(0) = flexAlignCenterCenter
   grd_LisDir.ColAlignment(1) = flexAlignLeftCenter
   grd_LisDir.ColAlignment(2) = flexAlignLeftCenter
   grd_LisDir.ColAlignment(3) = flexAlignLeftCenter
   grd_LisDir.ColAlignment(4) = flexAlignLeftCenter
   grd_LisDir.ColAlignment(5) = flexAlignLeftCenter
   grd_LisDir.ColAlignment(6) = flexAlignLeftCenter
   
   'Teléfonos
   grd_LisTel.Cols = 6
   grd_LisTel.ColWidth(0) = 1500
   grd_LisTel.ColWidth(1) = 1800
   grd_LisTel.ColWidth(2) = 2100
   grd_LisTel.ColWidth(3) = 1500
   grd_LisTel.ColWidth(4) = 0
   grd_LisTel.ColWidth(5) = 4200
   grd_LisTel.ColAlignment(0) = flexAlignCenterCenter
   grd_LisTel.ColAlignment(1) = flexAlignCenterCenter
   grd_LisTel.ColAlignment(2) = flexAlignCenterCenter
   grd_LisTel.ColAlignment(3) = flexAlignCenterCenter
   
   'Emails
   grd_LisEml.ColWidth(0) = 1500
   grd_LisEml.ColWidth(1) = 4000
   grd_LisEml.ColWidth(2) = 1500
   grd_LisEml.ColWidth(3) = 0
   grd_LisEml.ColWidth(4) = 4200
   grd_LisEml.ColAlignment(0) = flexAlignCenterCenter
   grd_LisEml.ColAlignment(1) = flexAlignLeftCenter
   grd_LisEml.ColAlignment(2) = flexAlignCenterCenter
   
   For l_int_Contad = 0 To 2
      cmb_Estado(l_int_Contad).AddItem "HABILITADO"
      cmb_Estado(l_int_Contad).ItemData(cmb_Estado(l_int_Contad).NewIndex) = CInt(1)
      cmb_Estado(l_int_Contad).AddItem "DESHABILITADO"
      cmb_Estado(l_int_Contad).ItemData(cmb_Estado(l_int_Contad).NewIndex) = CInt(2)
   Next l_int_Contad
   
   cmb_TipTel.AddItem "FIJO"
   cmb_TipTel.ItemData(cmb_TipTel.NewIndex) = CInt(1)
   cmb_TipTel.AddItem "MOVIL"
   cmb_TipTel.ItemData(cmb_TipTel.NewIndex) = CInt(2)
      
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_LisOpe)
   Call gs_LimpiaGrid(grd_LisDir)
   Call gs_LimpiaGrid(grd_LisTel)
   Call gs_LimpiaGrid(grd_LisEml)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub

Private Sub fs_DatOpe(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'Buscando Operaciones como Cliente Titular
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = " & p_TipDoc & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      grd_LisOpe.Redraw = False
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         grd_LisOpe.Text = Mid(g_rst_Princi!HIPMAE_NUMOPE, 1, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 4, 2) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 6, 5)
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         
         grd_LisOpe.Col = 2
         grd_LisOpe.Text = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
         
         g_rst_Princi.MoveNext
      Loop
      grd_LisOpe.Redraw = True
      Call gs_UbiIniGrid(grd_LisOpe)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DirCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   grd_LisDir.Redraw = False
   Call gs_LimpiaGrid(grd_LisDir)
   
   'CABECERA
   grd_LisDir.Rows = grd_LisDir.Rows + 2
   grd_LisDir.Row = grd_LisDir.Rows - 1
   grd_LisDir.FixedRows = 1

   grd_LisDir.Row = 0
   grd_LisDir.Col = 0:   grd_LisDir.Text = "INGRESO":       grd_LisDir.CellAlignment = flexAlignCenterCenter
   grd_LisDir.Col = 1:   grd_LisDir.Text = "DOMICILIO":     grd_LisDir.CellAlignment = flexAlignCenterCenter
   grd_LisDir.Col = 2:   grd_LisDir.Text = "DEPARTAMENTO":  grd_LisDir.CellAlignment = flexAlignCenterCenter
   grd_LisDir.Col = 3:   grd_LisDir.Text = "PROVINCIA":     grd_LisDir.CellAlignment = flexAlignCenterCenter
   grd_LisDir.Col = 4:   grd_LisDir.Text = "DISTRITO":      grd_LisDir.CellAlignment = flexAlignCenterCenter
   grd_LisDir.Col = 5:   grd_LisDir.Text = "ESTADO":        grd_LisDir.CellAlignment = flexAlignCenterCenter
   grd_LisDir.Rows = grd_LisDir.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ADCDIR_CORREL, ADCDIR_INGDAT, ADCDIR_TIPVIA, ADCDIR_NOMVIA, ADCDIR_NUMERO, ADCDIR_INTDPT, "
   g_str_Parame = g_str_Parame & "        ADCDIR_TIPZON , ADCDIR_NOMZON, ADCDIR_REFERE, ADCDIR_UBIGEO, ADCDIR_ESTADO, ADCDIR_COMENT "
   g_str_Parame = g_str_Parame & "   FROM CLI_ADCDIR  "
   g_str_Parame = g_str_Parame & "  WHERE ADCDIR_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND ADCDIR_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_LisDir.Rows = grd_LisDir.Rows + 1
         grd_LisDir.Row = grd_LisDir.Rows - 1
         
         grd_LisDir.Col = 0
         grd_LisDir.Text = moddat_gf_Consulta_ParDes("524", CStr(g_rst_Princi!ADCDIR_INGDAT))
      
         grd_LisDir.Col = 1
         grd_LisDir.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ADCDIR_TIPVIA)) & _
                           " " & Trim(g_rst_Princi!ADCDIR_NOMVIA) & " " & Trim(g_rst_Princi!ADCDIR_NUMERO) & _
                           IIf(Len(Trim(g_rst_Princi!ADCDIR_INTDPT)) > 0, " (" & Trim(g_rst_Princi!ADCDIR_INTDPT) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!ADCDIR_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ADCDIR_TIPZON)) & " " & Trim(g_rst_Princi!ADCDIR_NOMZON), "")
    
         grd_LisDir.Col = 2
         grd_LisDir.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ADCDIR_UBIGEO, 2) & "0000")
                           
         grd_LisDir.Col = 3
         grd_LisDir.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ADCDIR_UBIGEO, 4) & "00")

         grd_LisDir.Col = 4
         grd_LisDir.Text = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ADCDIR_UBIGEO))
         
         grd_LisDir.Col = 5
         grd_LisDir.Text = gf_NomEstado(cmb_Estado(0), g_rst_Princi!ADCDIR_ESTADO)
        
         grd_LisDir.Col = 6
         grd_LisDir.Text = CStr(g_rst_Princi!ADCDIR_CORREL)
                 
         g_rst_Princi.MoveNext
      Loop
      
      cmd_Editar(0).Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_LisDir)
   grd_LisDir.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   fs_Limpia (0)
End Sub

Public Function gf_NomEstado(p_Combo As ComboBox, p_Item As Integer) As String
Dim r_int_Contad  As Integer
Dim r_int_Ubicad  As Integer
   
   r_int_Ubicad = -1
   
   For r_int_Contad = 0 To p_Combo.ListCount - 1
      If p_Item = p_Combo.ItemData(r_int_Contad) Then
         r_int_Ubicad = r_int_Contad
         Exit For
      End If
   Next r_int_Contad
   p_Combo.ListIndex = r_int_Ubicad
   gf_NomEstado = p_Combo.Text
End Function

Private Sub fs_TelCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   grd_LisTel.Redraw = False
   Call gs_LimpiaGrid(grd_LisTel)
   
   'CABECERA
   grd_LisTel.Rows = grd_LisTel.Rows + 2
   grd_LisTel.Row = grd_LisTel.Rows - 1
   grd_LisTel.FixedRows = 1

   grd_LisTel.Row = 0
   grd_LisTel.Col = 0:   grd_LisTel.Text = "INGRESO":          grd_LisTel.CellAlignment = flexAlignCenterCenter
   grd_LisTel.Col = 1:   grd_LisTel.Text = "TIPO TELÉFONO":    grd_LisTel.CellAlignment = flexAlignCenterCenter
   grd_LisTel.Col = 2:   grd_LisTel.Text = "NÚMERO":           grd_LisTel.CellAlignment = flexAlignCenterCenter
   grd_LisTel.Col = 3:   grd_LisTel.Text = "ESTADO":           grd_LisTel.CellAlignment = flexAlignCenterCenter
   grd_LisTel.Col = 5:   grd_LisTel.Text = "COMENTARIO":       grd_LisTel.CellAlignment = flexAlignCenterCenter
   
   grd_LisTel.Rows = grd_LisTel.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ADCTEL_INGDAT, ADCTEL_TIPTEL, ADCTEL_NUMERO, ADCTEL_ESTADO, ADCTEL_CORREL, ADCTEL_COMENT "
   g_str_Parame = g_str_Parame & "   FROM CLI_ADCTEL  "
   g_str_Parame = g_str_Parame & "  WHERE ADCTEL_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND ADCTEL_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_LisTel.Rows = grd_LisTel.Rows + 1
         grd_LisTel.Row = grd_LisTel.Rows - 1
         
         grd_LisTel.Col = 0
         grd_LisTel.Text = moddat_gf_Consulta_ParDes("524", CStr(g_rst_Princi!ADCTEL_INGDAT))
      
         grd_LisTel.Col = 1
         grd_LisTel.Text = gf_NomEstado(cmb_TipTel, Trim(g_rst_Princi!ADCTEL_TIPTEL))
         
         grd_LisTel.Col = 2
         grd_LisTel.Text = Trim(g_rst_Princi!ADCTEL_NUMERO)
         
         grd_LisTel.Col = 3
         grd_LisTel.Text = gf_NomEstado(cmb_Estado(1), Trim(g_rst_Princi!ADCTEL_ESTADO))
         
         grd_LisTel.Col = 4
         grd_LisTel.Text = CStr(g_rst_Princi!ADCTEL_CORREL)
         
         grd_LisTel.Col = 5
         If IsNull(g_rst_Princi!ADCTEL_COMENT) Then
            grd_LisTel.Text = ""
         Else
            grd_LisTel.Text = CStr(g_rst_Princi!ADCTEL_COMENT)
         End If
            
                                                  
         g_rst_Princi.MoveNext
      Loop
      
      cmd_Editar(1).Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_LisTel)
   grd_LisTel.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Limpia(1)
End Sub

Private Sub fs_EmlCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   grd_LisEml.Redraw = False
   Call gs_LimpiaGrid(grd_LisEml)
   
   'CABECERA
   grd_LisEml.Rows = grd_LisEml.Rows + 2
   grd_LisEml.Row = grd_LisEml.Rows - 1
   grd_LisEml.FixedRows = 1

   grd_LisEml.Row = 0
   grd_LisEml.Col = 0:   grd_LisEml.Text = "INGRESO":    grd_LisEml.CellAlignment = flexAlignCenterCenter
   grd_LisEml.Col = 1:   grd_LisEml.Text = "EMAIL":      grd_LisEml.CellAlignment = flexAlignCenterCenter
   grd_LisEml.Col = 2:   grd_LisEml.Text = "ESTADO":     grd_LisEml.CellAlignment = flexAlignCenterCenter
   grd_LisEml.Col = 4:   grd_LisEml.Text = "COMENTARIO": grd_LisEml.CellAlignment = flexAlignCenterCenter
   grd_LisEml.Rows = grd_LisEml.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ADCEML_CORREL, ADCEML_INGDAT, ADCEML_NOMEML, ADCEML_ESTADO, ADCEML_COMENT "
   g_str_Parame = g_str_Parame & "   FROM CLI_ADCEML "
   g_str_Parame = g_str_Parame & "  WHERE ADCEML_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND ADCEML_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      'grd_LisEml.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisEml.Rows = grd_LisEml.Rows + 1
         grd_LisEml.Row = grd_LisEml.Rows - 1
         
         grd_LisEml.Col = 0
         grd_LisEml.Text = moddat_gf_Consulta_ParDes("524", CStr(g_rst_Princi!ADCEML_INGDAT))
      
         grd_LisEml.Col = 1
         grd_LisEml.Text = Trim(g_rst_Princi!ADCEML_NOMEML)
                    
         grd_LisEml.Col = 2
         grd_LisEml.Text = gf_NomEstado(cmb_Estado(2), Trim(g_rst_Princi!ADCEML_ESTADO))
         
         grd_LisEml.Col = 3
         grd_LisEml.Text = CStr(g_rst_Princi!ADCEML_CORREL)
         
         grd_LisEml.Col = 4
         If IsNull(g_rst_Princi!ADCEML_COMENT) Then
            grd_LisEml.Text = ""
         Else
            grd_LisEml.Text = CStr(g_rst_Princi!ADCEML_COMENT)
         End If
      
         g_rst_Princi.MoveNext
      Loop
      
      cmd_Editar(2).Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_LisEml)
   grd_LisEml.Redraw = True
         
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Limpia(2)
End Sub

Private Sub fs_Limpia(ByVal p_Indice As Integer)
   Select Case p_Indice
      Case 0:
            cmb_TipVia.ListIndex = -1
            txt_NomVia.Text = ""
            txt_NumVia.Text = ""
            txt_IntDpt.Text = ""
            cmb_TipZon.ListIndex = -1
            txt_NomZon.Text = ""
            cmb_DptDir.ListIndex = -1
            cmb_PrvDir.Clear
            cmb_DstDir.Clear
            txt_Refere.Text = ""
            
      Case 1:
            cmb_TipTel.ListIndex = -1
            txt_NumTel.Text = ""
      Case 2:
            txt_DirEle.Text = ""
   End Select
   cmb_Estado(p_Indice).ListIndex = -1
   txt_Coment(p_Indice).Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer, ByVal p_Indice As Integer)
   Select Case p_Indice
      Case 0:
            cmb_TipVia.Enabled = p_Habilita
            txt_NomVia.Enabled = p_Habilita
            txt_NumVia.Enabled = p_Habilita
            txt_IntDpt.Enabled = p_Habilita
            cmb_TipZon.Enabled = p_Habilita
            txt_NomZon.Enabled = p_Habilita
            cmb_DptDir.Enabled = p_Habilita
            cmb_PrvDir.Enabled = p_Habilita
            cmb_DstDir.Enabled = p_Habilita
            txt_Refere.Enabled = p_Habilita
      Case 1:
            cmb_TipTel.Enabled = p_Habilita
            txt_NumTel.Enabled = p_Habilita
      Case 2:
            txt_DirEle.Enabled = p_Habilita
   End Select
   cmd_Agrega(p_Indice).Enabled = Not p_Habilita
   cmd_Editar(p_Indice).Enabled = Not p_Habilita
   cmb_Estado(p_Indice).Enabled = p_Habilita
   txt_Coment(p_Indice).Enabled = p_Habilita
   cmd_Grabar(p_Indice).Enabled = p_Habilita
   cmd_Cancel(p_Indice).Enabled = p_Habilita
End Sub

Private Sub grd_LisDir_DblClick()
   grd_LisDir.Col = 6
   moddat_g_str_Codigo = grd_LisDir.Text
   'moddat_g_int_CodIns = grd_LisDir.Text
   
   If Not IsNull(CStr(moddat_g_str_Codigo)) And CStr(moddat_g_str_Codigo) <> "" Then
      Call gs_RefrescaGrid(grd_LisDir)
      moddat_g_int_FlgGrb = 2
      
      'Obteniendo Información del Registro
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT * FROM CLI_ADCDIR "
      g_str_Parame = g_str_Parame & "  WHERE ADCDIR_CORREL = " & CStr(moddat_g_str_Codigo) & " "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
          Exit Sub
      End If
      
      g_rst_GenAux.MoveFirst
      
      Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_GenAux!ADCDIR_TIPVIA)
      txt_NomVia.Text = Trim(g_rst_GenAux!ADCDIR_NOMVIA)
      txt_NumVia.Text = Trim(g_rst_GenAux!ADCDIR_NUMERO)
      txt_IntDpt.Text = IIf(IsNull(Trim(g_rst_GenAux!ADCDIR_INTDPT)), "", Trim(g_rst_GenAux!ADCDIR_INTDPT))
      Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_GenAux!ADCDIR_TIPZON)
      txt_NomZon.Text = Trim(g_rst_GenAux!ADCDIR_NOMZON)
     
      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_GenAux!ADCDIR_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_GenAux!ADCDIR_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_GenAux!ADCDIR_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_GenAux!ADCDIR_UBIGEO, 2), Mid(g_rst_GenAux!ADCDIR_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_GenAux!ADCDIR_UBIGEO, 2)))

      txt_Refere.Text = IIf(IsNull(Trim(g_rst_GenAux!ADCDIR_REFERE)), "", Trim(g_rst_GenAux!ADCDIR_REFERE))
      Call gs_BuscarCombo_Item(cmb_Estado(0), g_rst_GenAux!ADCDIR_ESTADO)
      txt_Coment(0).Text = IIf(IsNull(g_rst_GenAux!ADCDIR_COMENT), "", Trim(g_rst_GenAux!ADCDIR_COMENT))
      
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      
      Call fs_Activa(True, 0)
      Call gs_SetFocus(cmb_TipVia)
   Else
      MsgBox "Debe seleccionar Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Activa(False, 0)
      Call gs_SetFocus(grd_LisDir)
   End If
End Sub

Private Sub grd_LisEml_DblClick()
   grd_LisEml.Col = 3
   moddat_g_str_Codigo = grd_LisEml.Text
         
   If Not IsNull(CStr(moddat_g_str_Codigo)) And CStr(moddat_g_str_Codigo) <> "" Then
      Call gs_RefrescaGrid(grd_LisEml)
      moddat_g_int_FlgGrb_2 = 2
      
      'Obteniendo Información del Registro
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT * FROM CLI_ADCEML "
      g_str_Parame = g_str_Parame & "  WHERE ADCEML_CORREL = " & CStr(moddat_g_str_Codigo) & " "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      g_rst_Genera.MoveFirst
      txt_DirEle.Text = Trim(g_rst_Genera!ADCEML_NOMEML)
      Call gs_BuscarCombo_Item(cmb_Estado(2), g_rst_Genera!ADCEML_ESTADO)
      txt_Coment(2).Text = IIf(IsNull(Trim(g_rst_Genera!ADCEML_COMENT)), "", Trim(g_rst_Genera!ADCEML_COMENT))
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      Call fs_Activa(True, 2)
      Call gs_SetFocus(txt_DirEle)
   Else
      MsgBox "Debe seleccionar Email.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Activa(False, 2)
      Call gs_SetFocus(grd_LisEml)
   End If
End Sub

Private Sub grd_LisTel_DblClick()
   grd_LisTel.Col = 4
   moddat_g_str_Codigo = grd_LisTel.Text
         
   If Not IsNull(CStr(moddat_g_str_Codigo)) And CStr(moddat_g_str_Codigo) <> "" Then
      Call gs_RefrescaGrid(grd_LisTel)
      moddat_g_int_FlgGrb_1 = 2
      
      'Obteniendo Información del Registro
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT * FROM CLI_ADCTEL "
      g_str_Parame = g_str_Parame & "  WHERE ADCTEL_CORREL = " & CStr(moddat_g_str_Codigo) & " "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      g_rst_Genera.MoveFirst
      Call gs_BuscarCombo_Item(cmb_TipTel, g_rst_Genera!ADCTEL_TIPTEL)
      txt_NumTel.Text = Trim(g_rst_Genera!ADCTEL_NUMERO)
      Call gs_BuscarCombo_Item(cmb_Estado(1), g_rst_Genera!ADCTEL_ESTADO)
      txt_Coment(1).Text = IIf(IsNull(Trim(g_rst_Genera!ADCTEL_COMENT)), "", Trim(g_rst_Genera!ADCTEL_COMENT))
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      Call fs_Activa(True, 1)
      Call gs_SetFocus(cmb_TipTel)
   Else
      MsgBox "Debe seleccionar Teléfono.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Activa(False, 1)
      Call gs_SetFocus(grd_LisTel)
   End If
End Sub

Private Sub txt_Coment_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar(Index))
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-;:()_.")
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Estado(2))
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_numTel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Estado(1))
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Estado(0))
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub
