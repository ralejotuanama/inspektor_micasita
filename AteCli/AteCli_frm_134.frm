VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_PryNvi_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9570
   ClientLeft      =   2265
   ClientTop       =   585
   ClientWidth     =   11595
   Icon            =   "AteCli_frm_134.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9570
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   11595
      _Version        =   65536
      _ExtentX        =   20452
      _ExtentY        =   16880
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
         Left            =   60
         TabIndex        =   52
         Top             =   60
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin Threed.SSPanel pnl_TituloForm 
            Height          =   495
            Left            =   630
            TabIndex        =   53
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión de Proyectos No Vinculados"
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
            Picture         =   "AteCli_frm_134.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   54
         Top             =   750
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_134.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10860
            Picture         =   "AteCli_frm_134.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1980
         Left            =   60
         TabIndex        =   55
         Top             =   5490
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   3492
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
         Begin VB.CommandButton cmd_Editar_Pro 
            Height          =   585
            Left            =   10800
            Picture         =   "AteCli_frm_134.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia_Pro 
            Height          =   585
            Left            =   10200
            Picture         =   "AteCli_frm_134.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar_Pro 
            Height          =   585
            Left            =   9600
            Picture         =   "AteCli_frm_134.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.TextBox txt_NumDoc_Pro 
            Height          =   315
            Left            =   6030
            MaxLength       =   12
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   270
            Width           =   2025
         End
         Begin VB.ComboBox cmb_TipDoc_Pro 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   270
            Width           =   2775
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Pro 
            Height          =   1305
            Left            =   60
            TabIndex        =   42
            Top             =   630
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label6 
            Caption         =   "Promotor"
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
            TabIndex        =   58
            Top             =   30
            Width           =   1845
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   57
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   56
            Top             =   330
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1980
         Left            =   60
         TabIndex        =   59
         Top             =   7530
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   3492
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
         Begin VB.ComboBox cmb_TipDoc_Con 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   270
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc_Con 
            Height          =   315
            Left            =   6030
            MaxLength       =   12
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   270
            Width           =   2025
         End
         Begin VB.CommandButton cmd_Buscar_Con 
            Height          =   585
            Left            =   9600
            Picture         =   "AteCli_frm_134.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia_Con 
            Height          =   585
            Left            =   10200
            Picture         =   "AteCli_frm_134.frx":17C2
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar_Con 
            Height          =   585
            Left            =   10800
            Picture         =   "AteCli_frm_134.frx":1ACC
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   30
            Width           =   585
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Con 
            Height          =   1305
            Left            =   60
            TabIndex        =   48
            Top             =   630
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   62
            Top             =   330
            Width           =   1845
         End
         Begin VB.Label Label8 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   4800
            TabIndex        =   61
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label7 
            Caption         =   "Constructor"
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
            TabIndex        =   60
            Top             =   30
            Width           =   1845
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3975
         Left            =   60
         TabIndex        =   0
         Top             =   1440
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   7011
         _Version        =   393216
         Tabs            =   5
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Datos Principales"
         TabPicture(0)   =   "AteCli_frm_134.frx":1DD6
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "pnl_SolEva"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos Adicionales I"
         TabPicture(1)   =   "AteCli_frm_134.frx":1DF2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSPanel5"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Datos Adicionales II"
         TabPicture(2)   =   "AteCli_frm_134.frx":1E0E
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "SSPanel7"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Relacion Consejeros"
         TabPicture(3)   =   "AteCli_frm_134.frx":1E2A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSPanel9"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Competencias"
         TabPicture(4)   =   "AteCli_frm_134.frx":1E46
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "SSPanel3"
         Tab(4).ControlCount=   1
         Begin Threed.SSPanel pnl_SolEva 
            Height          =   3435
            Left            =   -74880
            TabIndex        =   63
            Top             =   420
            Width           =   11235
            _Version        =   65536
            _ExtentX        =   19817
            _ExtentY        =   6059
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
            Begin VB.ComboBox cmb_TipoBien 
               Height          =   315
               ItemData        =   "AteCli_frm_134.frx":1E62
               Left            =   7650
               List            =   "AteCli_frm_134.frx":1E64
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   1350
               Width           =   2325
            End
            Begin VB.ComboBox cmb_EntFin 
               Height          =   315
               ItemData        =   "AteCli_frm_134.frx":1E66
               Left            =   1920
               List            =   "AteCli_frm_134.frx":1E68
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   150
               Width           =   9165
            End
            Begin VB.TextBox txt_NomPry 
               Height          =   315
               Left            =   1920
               MaxLength       =   250
               TabIndex        =   2
               Text            =   "Text1"
               Top             =   480
               Width           =   9135
            End
            Begin VB.TextBox txt_DesPry 
               Height          =   525
               Left            =   1920
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "AteCli_frm_134.frx":1E6A
               Top             =   810
               Width           =   9135
            End
            Begin VB.ComboBox cmb_Situac 
               Height          =   315
               ItemData        =   "AteCli_frm_134.frx":1E72
               Left            =   1920
               List            =   "AteCli_frm_134.frx":1E74
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1350
               Width           =   1695
            End
            Begin VB.ComboBox cmb_ConAso 
               Height          =   315
               Left            =   5340
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1350
               Width           =   765
            End
            Begin VB.TextBox txt_Refere 
               Height          =   315
               Left            =   7650
               MaxLength       =   250
               TabIndex        =   16
               Text            =   "Text1"
               Top             =   3000
               Width           =   3405
            End
            Begin VB.ComboBox cmb_DstDir 
               Height          =   315
               Left            =   1920
               TabIndex        =   15
               Text            =   "cmb_DstDir"
               Top             =   3000
               Width           =   4185
            End
            Begin VB.ComboBox cmb_PrvDir 
               Height          =   315
               Left            =   7650
               TabIndex        =   14
               Text            =   "cmb_PrvDir"
               Top             =   2670
               Width           =   3405
            End
            Begin VB.ComboBox cmb_DptDir 
               Height          =   315
               Left            =   1920
               TabIndex        =   13
               Text            =   "cmb_DptDir"
               Top             =   2670
               Width           =   4185
            End
            Begin VB.TextBox txt_NomZon 
               Height          =   315
               Left            =   7650
               MaxLength       =   120
               TabIndex        =   12
               Text            =   "Text1"
               Top             =   2340
               Width           =   3405
            End
            Begin VB.ComboBox cmb_TipZon 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   2340
               Width           =   4185
            End
            Begin VB.TextBox txt_NumVia 
               Height          =   315
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   9
               Text            =   "Text1"
               Top             =   2010
               Width           =   1695
            End
            Begin VB.TextBox txt_NomVia 
               Height          =   315
               Left            =   7650
               MaxLength       =   120
               TabIndex        =   8
               Text            =   "Text1"
               Top             =   1680
               Width           =   3405
            End
            Begin VB.ComboBox cmb_TipVia 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   1680
               Width           =   4185
            End
            Begin VB.TextBox txt_Interi 
               Height          =   315
               Left            =   7650
               MaxLength       =   15
               TabIndex        =   10
               Text            =   "Text1"
               Top             =   2010
               Width           =   2325
            End
            Begin VB.Label Label80 
               Caption         =   "Tipo de Bien:"
               Height          =   285
               Left            =   6360
               TabIndex        =   170
               Top             =   1380
               Width           =   1485
            End
            Begin VB.Label Label18 
               Caption         =   "Entidad Financiera:"
               Height          =   225
               Left            =   90
               TabIndex        =   97
               Top             =   180
               Width           =   1725
            End
            Begin VB.Label Label32 
               Caption         =   "Nombre Proyecto:"
               Height          =   285
               Left            =   90
               TabIndex        =   77
               Top             =   510
               Width           =   1785
            End
            Begin VB.Label Label2 
               Caption         =   "Descripción Proyecto:"
               Height          =   285
               Left            =   90
               TabIndex        =   76
               Top             =   810
               Width           =   1785
            End
            Begin VB.Label Label4 
               Caption         =   "Situación:"
               Height          =   285
               Left            =   90
               TabIndex        =   75
               Top             =   1380
               Width           =   1785
            End
            Begin VB.Label Label10 
               Caption         =   "Convenio Asociación:"
               Height          =   285
               Left            =   3720
               TabIndex        =   74
               Top             =   1380
               Width           =   1665
            End
            Begin VB.Label Label28 
               Caption         =   "Referencia:"
               Height          =   285
               Left            =   6360
               TabIndex        =   73
               Top             =   3030
               Width           =   1485
            End
            Begin VB.Label Label26 
               Caption         =   "Distrito:"
               Height          =   315
               Left            =   90
               TabIndex        =   72
               Top             =   3030
               Width           =   1785
            End
            Begin VB.Label Label25 
               Caption         =   "Provincia:"
               Height          =   315
               Left            =   6360
               TabIndex        =   71
               Top             =   2700
               Width           =   1485
            End
            Begin VB.Label Label24 
               Caption         =   "Departamento:"
               Height          =   315
               Left            =   90
               TabIndex        =   70
               Top             =   2700
               Width           =   1785
            End
            Begin VB.Label Label23 
               Caption         =   "Nombre Zona:"
               Height          =   285
               Left            =   6360
               TabIndex        =   69
               Top             =   2370
               Width           =   1485
            End
            Begin VB.Label Label22 
               Caption         =   "Tipo de Zona:"
               Height          =   315
               Left            =   90
               TabIndex        =   68
               Top             =   2370
               Width           =   1785
            End
            Begin VB.Label Label11 
               Caption         =   "Nro. / Mza. / Lot.:"
               Height          =   285
               Left            =   90
               TabIndex        =   67
               Top             =   2040
               Width           =   1785
            End
            Begin VB.Label Label12 
               Caption         =   "Nombre Vía:"
               Height          =   285
               Left            =   6360
               TabIndex        =   66
               Top             =   1710
               Width           =   1485
            End
            Begin VB.Label Label19 
               Caption         =   "Tipo de Vía:"
               Height          =   315
               Left            =   90
               TabIndex        =   65
               Top             =   1710
               Width           =   1785
            End
            Begin VB.Label Label13 
               Caption         =   "Int. / Dpto.:"
               Height          =   285
               Left            =   6360
               TabIndex        =   64
               Top             =   2040
               Width           =   1485
            End
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   3435
            Left            =   -74880
            TabIndex        =   78
            Top             =   420
            Width           =   11235
            _Version        =   65536
            _ExtentX        =   19817
            _ExtentY        =   6059
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
            Begin VB.ComboBox cmb_TipoGar 
               Height          =   315
               ItemData        =   "AteCli_frm_134.frx":1E76
               Left            =   7230
               List            =   "AteCli_frm_134.frx":1E83
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   1920
               Width           =   3795
            End
            Begin VB.TextBox txt_Contacto 
               Height          =   330
               Left            =   1410
               MaxLength       =   50
               TabIndex        =   17
               Text            =   "Text1"
               Top             =   135
               Width           =   4065
            End
            Begin EditLib.fpDoubleSingle ipp_PreMin 
               Height          =   330
               Left            =   1410
               TabIndex        =   21
               Top             =   1575
               Width           =   1275
               _Version        =   196608
               _ExtentX        =   2249
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.TextBox txt_Cargo 
               Height          =   330
               Left            =   1410
               MaxLength       =   25
               TabIndex        =   18
               Text            =   "Text1"
               Top             =   495
               Width           =   4065
            End
            Begin VB.TextBox txt_Telefono 
               Height          =   330
               Left            =   1410
               MaxLength       =   30
               TabIndex        =   20
               Text            =   "Text1"
               Top             =   1215
               Width           =   4065
            End
            Begin VB.TextBox txt_Email 
               Height          =   330
               Left            =   1410
               MaxLength       =   40
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   855
               Width           =   4065
            End
            Begin EditLib.fpDoubleSingle ipp_PreMax 
               Height          =   330
               Left            =   4200
               TabIndex        =   22
               Top             =   1575
               Width           =   1275
               _Version        =   196608
               _ExtentX        =   2249
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_AreaMin 
               Height          =   315
               Left            =   1410
               TabIndex        =   23
               Top             =   1935
               Width           =   1275
               _Version        =   196608
               _ExtentX        =   2249
               _ExtentY        =   556
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_AreaMax 
               Height          =   315
               Left            =   4200
               TabIndex        =   24
               Top             =   1935
               Width           =   1275
               _Version        =   196608
               _ExtentX        =   2249
               _ExtentY        =   556
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_TotVen 
               Height          =   330
               Left            =   10020
               TabIndex        =   28
               Top             =   135
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   ""
               FixedPoint      =   0   'False
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_TotUni 
               Height          =   330
               Left            =   7230
               TabIndex        =   27
               Top             =   135
               Width           =   1005
               _Version        =   196608
               _ExtentX        =   1773
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   ""
               FixedPoint      =   0   'False
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_TotDisp 
               Height          =   330
               Left            =   7230
               TabIndex        =   29
               Top             =   495
               Width           =   1005
               _Version        =   196608
               _ExtentX        =   1773
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   1
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   ""
               FixedPoint      =   0   'False
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   -1  'True
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Avance 
               Height          =   330
               Left            =   7230
               TabIndex        =   31
               Top             =   855
               Width           =   1005
               _Version        =   196608
               _ExtentX        =   1773
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   1
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Dispon 
               Height          =   330
               Left            =   10020
               TabIndex        =   30
               Top             =   495
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   1
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Partic 
               Height          =   330
               Left            =   10020
               TabIndex        =   33
               Top             =   1215
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   1
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "-9000000000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Coloca 
               Height          =   330
               Left            =   7230
               TabIndex        =   32
               Top             =   1215
               Width           =   1005
               _Version        =   196608
               _ExtentX        =   1773
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   ""
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_NumEtapa 
               Height          =   330
               Left            =   7230
               TabIndex        =   34
               Top             =   1575
               Width           =   1005
               _Version        =   196608
               _ExtentX        =   1773
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0"
               DecimalPlaces   =   0
               DecimalPoint    =   ""
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "-9000000000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDateTime ipp_FecFin 
               Height          =   315
               Left            =   4200
               TabIndex        =   26
               Top             =   2280
               Width           =   1305
               _Version        =   196608
               _ExtentX        =   2302
               _ExtentY        =   556
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483637
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   3
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   0
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   ""
               DateCalcMethod  =   0
               DateTimeFormat  =   0
               UserDefinedFormat=   ""
               DateMax         =   "00000000"
               DateMin         =   "00000000"
               TimeMax         =   "000000"
               TimeMin         =   "000000"
               TimeString1159  =   ""
               TimeString2359  =   ""
               DateDefault     =   "00000000"
               TimeDefault     =   "000000"
               TimeStyle       =   0
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               PopUpType       =   0
               DateCalcY2KSplit=   60
               CaretPosition   =   0
               IncYear         =   1
               IncMonth        =   1
               IncDay          =   1
               IncHour         =   1
               IncMinute       =   1
               IncSecond       =   1
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               StartMonth      =   4
               ButtonAlign     =   0
               BoundDataType   =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDateTime ipp_FecIni 
               Height          =   315
               Left            =   1410
               TabIndex        =   25
               Top             =   2280
               Width           =   1305
               _Version        =   196608
               _ExtentX        =   2302
               _ExtentY        =   556
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483637
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   3
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   0
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   ""
               DateCalcMethod  =   0
               DateTimeFormat  =   0
               UserDefinedFormat=   ""
               DateMax         =   "00000000"
               DateMin         =   "00000000"
               TimeMax         =   "000000"
               TimeMin         =   "000000"
               TimeString1159  =   ""
               TimeString2359  =   ""
               DateDefault     =   "00000000"
               TimeDefault     =   "000000"
               TimeStyle       =   0
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               PopUpType       =   0
               DateCalcY2KSplit=   60
               CaretPosition   =   0
               IncYear         =   1
               IncMonth        =   1
               IncDay          =   1
               IncHour         =   1
               IncMinute       =   1
               IncSecond       =   1
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               StartMonth      =   4
               ButtonAlign     =   0
               BoundDataType   =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle ipp_Tasa 
               Height          =   330
               Left            =   10020
               TabIndex        =   35
               Top             =   1560
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
               _ExtentY        =   582
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   1
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   1
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   0
               BorderColor     =   -2147483642
               BorderWidth     =   1
               ButtonDisable   =   0   'False
               ButtonHide      =   0   'False
               ButtonIncrement =   1
               ButtonMin       =   0
               ButtonMax       =   100
               ButtonStyle     =   0
               ButtonWidth     =   0
               ButtonWrap      =   -1  'True
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483633
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               CaretInsert     =   0
               CaretOverWrite  =   3
               UserEntry       =   0
               HideSelection   =   -1  'True
               InvalidColor    =   -2147483637
               InvalidOption   =   0
               MarginLeft      =   3
               MarginTop       =   3
               MarginRight     =   3
               MarginBottom    =   3
               NullColor       =   -2147483637
               OnFocusAlignH   =   0
               OnFocusAlignV   =   0
               OnFocusNoSelect =   0   'False
               OnFocusPosition =   0
               ControlType     =   0
               Text            =   "0.00"
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000"
               MinValue        =   "0"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483633
               Appearance      =   2
               BorderDropShadow=   0
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483633
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label16 
               Caption         =   "Tasa del Proyecto:"
               Height          =   195
               Index           =   0
               Left            =   8460
               TabIndex        =   172
               Top             =   1620
               Width           =   1455
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Garantia:"
               Height          =   195
               Index           =   0
               Left            =   5790
               TabIndex        =   171
               Top             =   1965
               Width           =   1455
            End
            Begin VB.Label Label36 
               Caption         =   "Total Unidades:"
               Height          =   195
               Left            =   5790
               TabIndex        =   96
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label Label38 
               Caption         =   "Total Disponible:"
               Height          =   195
               Left            =   5790
               TabIndex        =   95
               Top             =   540
               Width           =   1215
            End
            Begin VB.Label Label39 
               Caption         =   "% Avance:"
               Height          =   195
               Left            =   5790
               TabIndex        =   94
               Top             =   900
               Width           =   1215
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Colocac. miCasita:"
               Height          =   195
               Left            =   5790
               TabIndex        =   93
               Top             =   1260
               Width           =   1305
            End
            Begin VB.Label Label41 
               Caption         =   "% Participacion:"
               Height          =   195
               Left            =   8460
               TabIndex        =   92
               Top             =   1260
               Width           =   1425
            End
            Begin VB.Label Label40 
               Caption         =   "% Disponible:"
               Height          =   195
               Left            =   8460
               TabIndex        =   91
               Top             =   540
               Width           =   1425
            End
            Begin VB.Label Label37 
               Caption         =   "Total Vendido:"
               Height          =   195
               Left            =   8460
               TabIndex        =   90
               Top             =   180
               Width           =   1425
            End
            Begin VB.Label Label35 
               Caption         =   "Fecha Fin Obra:"
               Height          =   195
               Left            =   2880
               TabIndex        =   89
               Top             =   2325
               Width           =   1275
            End
            Begin VB.Label Label34 
               Caption         =   "Fecha Inicio: Obra:"
               Height          =   195
               Left            =   120
               TabIndex        =   88
               Top             =   2325
               Width           =   1095
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               Caption         =   "Area Max.(m2):"
               Height          =   195
               Left            =   2880
               TabIndex        =   87
               Top             =   1980
               Width           =   1065
            End
            Begin VB.Label Label31 
               Caption         =   "Precio Max.(S/.):"
               Height          =   195
               Left            =   2880
               TabIndex        =   86
               Top             =   1620
               Width           =   1245
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Nº Etapas:"
               Height          =   195
               Left            =   5790
               TabIndex        =   85
               Top             =   1620
               Width           =   885
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Area Min.(m2):"
               Height          =   195
               Left            =   120
               TabIndex        =   84
               Top             =   1980
               Width           =   1020
            End
            Begin VB.Label Label27 
               Caption         =   "E-Mail:"
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   900
               Width           =   1095
            End
            Begin VB.Label Label21 
               Caption         =   "Cargo:"
               Height          =   195
               Left            =   120
               TabIndex        =   82
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "Telefono(s):"
               Height          =   195
               Left            =   120
               TabIndex        =   81
               Top             =   1260
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "Precio Min.(S/.):"
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   1620
               Width           =   1185
            End
            Begin VB.Label Label14 
               Caption         =   "Contacto:"
               Height          =   195
               Left            =   120
               TabIndex        =   79
               Top             =   180
               Width           =   1095
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   3435
            Left            =   -74880
            TabIndex        =   98
            Top             =   420
            Width           =   11235
            _Version        =   65536
            _ExtentX        =   19817
            _ExtentY        =   6059
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
            Begin TabDlg.SSTab SSTab2 
               Height          =   3285
               Left            =   90
               TabIndex        =   99
               Top             =   90
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   5794
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Banco 1"
               TabPicture(0)   =   "AteCli_frm_134.frx":1EAB
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Label52"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Label46"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Label45(1)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Label44(1)"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "Label1"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "Label15(1)"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "Label16(1)"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "ipp_TasaBco1"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "cmb_EntFinBco1"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "Frame1"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "txt_ModEvaBco1"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "txt_ComeBco1"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "cmb_PlazoBco1"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).Control(13)=   "cmb_TipoGarBco1"
               Tab(0).Control(13).Enabled=   0   'False
               Tab(0).Control(14)=   "cmb_CostoBco1"
               Tab(0).Control(14).Enabled=   0   'False
               Tab(0).ControlCount=   15
               TabCaption(1)   =   "Banco 2"
               TabPicture(1)   =   "AteCli_frm_134.frx":1EC7
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Frame2"
               Tab(1).Control(1)=   "txt_ModEvaBco2"
               Tab(1).Control(2)=   "txt_ComeBco2"
               Tab(1).Control(3)=   "cmb_PlazoBco2"
               Tab(1).Control(4)=   "cmb_TipoGarBco2"
               Tab(1).Control(5)=   "cmb_CostoBco2"
               Tab(1).Control(6)=   "cmb_EntFinBco2"
               Tab(1).Control(7)=   "ipp_TasaBco2"
               Tab(1).Control(8)=   "Label68"
               Tab(1).Control(9)=   "Label69"
               Tab(1).Control(10)=   "Label70"
               Tab(1).Control(11)=   "Label71"
               Tab(1).Control(12)=   "Label72"
               Tab(1).Control(13)=   "Label73"
               Tab(1).Control(14)=   "Label74"
               Tab(1).ControlCount=   15
               Begin VB.ComboBox cmb_CostoBco1 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1EE3
                  Left            =   8100
                  List            =   "AteCli_frm_134.frx":1EE5
                  Style           =   2  'Dropdown List
                  TabIndex        =   150
                  Top             =   840
                  Width           =   795
               End
               Begin VB.ComboBox cmb_TipoGarBco1 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1EE7
                  Left            =   3690
                  List            =   "AteCli_frm_134.frx":1EE9
                  Style           =   2  'Dropdown List
                  TabIndex        =   148
                  Top             =   840
                  Width           =   3495
               End
               Begin VB.ComboBox cmb_PlazoBco1 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1EEB
                  Left            =   10200
                  List            =   "AteCli_frm_134.frx":1EED
                  Style           =   2  'Dropdown List
                  TabIndex        =   147
                  Top             =   840
                  Width           =   705
               End
               Begin VB.TextBox txt_ComeBco1 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   146
                  Text            =   "Text1"
                  Top             =   1170
                  Width           =   9195
               End
               Begin VB.TextBox txt_ModEvaBco1 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   100
                  TabIndex        =   145
                  Text            =   "Text1"
                  Top             =   1500
                  Width           =   9195
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Modalidad de Ahorro"
                  Height          =   1155
                  Left            =   90
                  TabIndex        =   126
                  Top             =   2010
                  Width           =   10905
                  Begin VB.ComboBox cmb_CtaIni1Bco1 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1EEF
                     Left            =   1590
                     List            =   "AteCli_frm_134.frx":1EF1
                     Style           =   2  'Dropdown List
                     TabIndex        =   134
                     Top             =   300
                     Width           =   735
                  End
                  Begin VB.ComboBox cmb_Plazo1Bco1 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1EF3
                     Left            =   3480
                     List            =   "AteCli_frm_134.frx":1EF5
                     Style           =   2  'Dropdown List
                     TabIndex        =   133
                     Top             =   300
                     Width           =   615
                  End
                  Begin VB.ComboBox cmb_ValIng1Bco1 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1EF7
                     Left            =   5460
                     List            =   "AteCli_frm_134.frx":1EF9
                     Style           =   2  'Dropdown List
                     TabIndex        =   132
                     Top             =   300
                     Width           =   735
                  End
                  Begin VB.TextBox txt_Com1Bco1 
                     Height          =   285
                     Left            =   7170
                     MaxLength       =   100
                     TabIndex        =   131
                     Text            =   "Text1"
                     Top             =   300
                     Width           =   3615
                  End
                  Begin VB.ComboBox cmb_CtaIni2Bco1 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1EFB
                     Left            =   1590
                     List            =   "AteCli_frm_134.frx":1EFD
                     Style           =   2  'Dropdown List
                     TabIndex        =   130
                     Top             =   660
                     Width           =   735
                  End
                  Begin VB.ComboBox cmb_Plazo2Bco1 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1EFF
                     Left            =   3480
                     List            =   "AteCli_frm_134.frx":1F01
                     Style           =   2  'Dropdown List
                     TabIndex        =   129
                     Top             =   660
                     Width           =   615
                  End
                  Begin VB.ComboBox cmb_ValIng2Bco1 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F03
                     Left            =   5460
                     List            =   "AteCli_frm_134.frx":1F05
                     Style           =   2  'Dropdown List
                     TabIndex        =   128
                     Top             =   660
                     Width           =   735
                  End
                  Begin VB.TextBox txt_Com2Bco1 
                     Height          =   285
                     Left            =   7170
                     MaxLength       =   100
                     TabIndex        =   127
                     Text            =   "Text1"
                     Top             =   660
                     Width           =   3615
                  End
                  Begin VB.Label Label47 
                     AutoSize        =   -1  'True
                     Caption         =   "Cuota Inicial (%):"
                     Height          =   195
                     Index           =   0
                     Left            =   390
                     TabIndex        =   144
                     Top             =   330
                     Width           =   1170
                  End
                  Begin VB.Label Label49 
                     AutoSize        =   -1  'True
                     Caption         =   "Valid. Ingreso (%):"
                     Height          =   195
                     Left            =   4200
                     TabIndex        =   143
                     Top             =   330
                     Width           =   1260
                  End
                  Begin VB.Label Label50 
                     AutoSize        =   -1  'True
                     Caption         =   "Comentario:"
                     Height          =   195
                     Left            =   6270
                     TabIndex        =   142
                     Top             =   330
                     Width           =   840
                  End
                  Begin VB.Label Label51 
                     AutoSize        =   -1  'True
                     Caption         =   "Plazo (meses):"
                     Height          =   195
                     Left            =   2430
                     TabIndex        =   141
                     Top             =   330
                     Width           =   1020
                  End
                  Begin VB.Label Label48 
                     AutoSize        =   -1  'True
                     Caption         =   "Valid. Ingreso (%):"
                     Height          =   195
                     Left            =   4200
                     TabIndex        =   140
                     Top             =   690
                     Width           =   1260
                  End
                  Begin VB.Label Label53 
                     AutoSize        =   -1  'True
                     Caption         =   "Comentario:"
                     Height          =   195
                     Left            =   6270
                     TabIndex        =   139
                     Top             =   690
                     Width           =   840
                  End
                  Begin VB.Label Label54 
                     AutoSize        =   -1  'True
                     Caption         =   "Plazo (meses):"
                     Height          =   195
                     Left            =   2430
                     TabIndex        =   138
                     Top             =   690
                     Width           =   1020
                  End
                  Begin VB.Label Label55 
                     AutoSize        =   -1  'True
                     Caption         =   "Cuota Inicial (%):"
                     Height          =   195
                     Left            =   390
                     TabIndex        =   137
                     Top             =   690
                     Width           =   1170
                  End
                  Begin VB.Label Label56 
                     AutoSize        =   -1  'True
                     Caption         =   "1."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   150
                     TabIndex        =   136
                     Top             =   330
                     Width           =   180
                  End
                  Begin VB.Label Label57 
                     AutoSize        =   -1  'True
                     Caption         =   "2."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   150
                     TabIndex        =   135
                     Top             =   690
                     Width           =   180
                  End
               End
               Begin VB.Frame Frame2 
                  Caption         =   "Modalidad de Ahorro"
                  Height          =   1155
                  Left            =   -74910
                  TabIndex        =   107
                  Top             =   2010
                  Width           =   10905
                  Begin VB.TextBox txt_Com2Bco2 
                     Height          =   285
                     Left            =   7170
                     MaxLength       =   100
                     TabIndex        =   115
                     Text            =   "Text1"
                     Top             =   660
                     Width           =   3615
                  End
                  Begin VB.ComboBox cmb_ValIng2Bco2 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F07
                     Left            =   5460
                     List            =   "AteCli_frm_134.frx":1F09
                     Style           =   2  'Dropdown List
                     TabIndex        =   114
                     Top             =   660
                     Width           =   735
                  End
                  Begin VB.ComboBox cmb_Plazo2Bco2 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F0B
                     Left            =   3480
                     List            =   "AteCli_frm_134.frx":1F0D
                     Style           =   2  'Dropdown List
                     TabIndex        =   113
                     Top             =   660
                     Width           =   615
                  End
                  Begin VB.ComboBox cmb_CtaIni2Bco2 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F0F
                     Left            =   1590
                     List            =   "AteCli_frm_134.frx":1F11
                     Style           =   2  'Dropdown List
                     TabIndex        =   112
                     Top             =   660
                     Width           =   735
                  End
                  Begin VB.TextBox txt_Com1Bco2 
                     Height          =   285
                     Left            =   7170
                     MaxLength       =   100
                     TabIndex        =   111
                     Text            =   "Text1"
                     Top             =   300
                     Width           =   3615
                  End
                  Begin VB.ComboBox cmb_ValIng1Bco2 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F13
                     Left            =   5460
                     List            =   "AteCli_frm_134.frx":1F15
                     Style           =   2  'Dropdown List
                     TabIndex        =   110
                     Top             =   300
                     Width           =   735
                  End
                  Begin VB.ComboBox cmb_Plazo1Bco2 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F17
                     Left            =   3480
                     List            =   "AteCli_frm_134.frx":1F19
                     Style           =   2  'Dropdown List
                     TabIndex        =   109
                     Top             =   300
                     Width           =   615
                  End
                  Begin VB.ComboBox cmb_CtaIni1Bco2 
                     Height          =   315
                     ItemData        =   "AteCli_frm_134.frx":1F1B
                     Left            =   1590
                     List            =   "AteCli_frm_134.frx":1F1D
                     Style           =   2  'Dropdown List
                     TabIndex        =   108
                     Top             =   300
                     Width           =   735
                  End
                  Begin VB.Label Label58 
                     AutoSize        =   -1  'True
                     Caption         =   "2."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   150
                     TabIndex        =   125
                     Top             =   690
                     Width           =   180
                  End
                  Begin VB.Label Label59 
                     AutoSize        =   -1  'True
                     Caption         =   "1."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   150
                     TabIndex        =   124
                     Top             =   330
                     Width           =   180
                  End
                  Begin VB.Label Label60 
                     AutoSize        =   -1  'True
                     Caption         =   "Cuota Inicial (%):"
                     Height          =   195
                     Left            =   390
                     TabIndex        =   123
                     Top             =   690
                     Width           =   1170
                  End
                  Begin VB.Label Label61 
                     AutoSize        =   -1  'True
                     Caption         =   "Plazo (meses):"
                     Height          =   195
                     Left            =   2430
                     TabIndex        =   122
                     Top             =   690
                     Width           =   1020
                  End
                  Begin VB.Label Label62 
                     AutoSize        =   -1  'True
                     Caption         =   "Comentario:"
                     Height          =   195
                     Left            =   6270
                     TabIndex        =   121
                     Top             =   690
                     Width           =   840
                  End
                  Begin VB.Label Label63 
                     AutoSize        =   -1  'True
                     Caption         =   "Valid. Ingreso (%):"
                     Height          =   195
                     Left            =   4200
                     TabIndex        =   120
                     Top             =   690
                     Width           =   1260
                  End
                  Begin VB.Label Label64 
                     AutoSize        =   -1  'True
                     Caption         =   "Plazo (meses):"
                     Height          =   195
                     Left            =   2430
                     TabIndex        =   119
                     Top             =   330
                     Width           =   1020
                  End
                  Begin VB.Label Label65 
                     AutoSize        =   -1  'True
                     Caption         =   "Comentario:"
                     Height          =   195
                     Left            =   6270
                     TabIndex        =   118
                     Top             =   330
                     Width           =   840
                  End
                  Begin VB.Label Label66 
                     AutoSize        =   -1  'True
                     Caption         =   "Valid. Ingreso (%):"
                     Height          =   195
                     Left            =   4200
                     TabIndex        =   117
                     Top             =   330
                     Width           =   1260
                  End
                  Begin VB.Label Label67 
                     AutoSize        =   -1  'True
                     Caption         =   "Cuota Inicial (%):"
                     Height          =   195
                     Left            =   390
                     TabIndex        =   116
                     Top             =   330
                     Width           =   1170
                  End
               End
               Begin VB.TextBox txt_ModEvaBco2 
                  Height          =   315
                  Left            =   -73320
                  MaxLength       =   100
                  TabIndex        =   106
                  Text            =   "Text1"
                  Top             =   1500
                  Width           =   9195
               End
               Begin VB.TextBox txt_ComeBco2 
                  Height          =   315
                  Left            =   -73320
                  MaxLength       =   100
                  TabIndex        =   105
                  Text            =   "Text1"
                  Top             =   1170
                  Width           =   9195
               End
               Begin VB.ComboBox cmb_PlazoBco2 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1F1F
                  Left            =   -64800
                  List            =   "AteCli_frm_134.frx":1F21
                  Style           =   2  'Dropdown List
                  TabIndex        =   104
                  Top             =   840
                  Width           =   705
               End
               Begin VB.ComboBox cmb_TipoGarBco2 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1F23
                  Left            =   -71310
                  List            =   "AteCli_frm_134.frx":1F25
                  Style           =   2  'Dropdown List
                  TabIndex        =   103
                  Top             =   840
                  Width           =   3495
               End
               Begin VB.ComboBox cmb_CostoBco2 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1F27
                  Left            =   -66900
                  List            =   "AteCli_frm_134.frx":1F29
                  Style           =   2  'Dropdown List
                  TabIndex        =   102
                  Top             =   840
                  Width           =   795
               End
               Begin VB.ComboBox cmb_EntFinBco2 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1F2B
                  Left            =   -73320
                  List            =   "AteCli_frm_134.frx":1F2D
                  Style           =   2  'Dropdown List
                  TabIndex        =   101
                  Top             =   510
                  Width           =   9225
               End
               Begin VB.ComboBox cmb_EntFinBco1 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1F2F
                  Left            =   1680
                  List            =   "AteCli_frm_134.frx":1F31
                  Style           =   2  'Dropdown List
                  TabIndex        =   100
                  Top             =   510
                  Width           =   9225
               End
               Begin EditLib.fpDoubleSingle ipp_TasaBco1 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   149
                  Top             =   840
                  Width           =   1005
                  _Version        =   196608
                  _ExtentX        =   1773
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483633
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   0
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000000"
                  MinValue        =   "-9000000000"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  IncDec          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDoubleSingle ipp_TasaBco2 
                  Height          =   315
                  Left            =   -73320
                  TabIndex        =   151
                  Top             =   840
                  Width           =   1005
                  _Version        =   196608
                  _ExtentX        =   1773
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483633
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   0
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000000"
                  MinValue        =   "-9000000000"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  IncDec          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin VB.Label Label16 
                  Caption         =   "Tasa de Proyecto:"
                  Height          =   195
                  Index           =   1
                  Left            =   210
                  TabIndex        =   165
                  Top             =   870
                  Width           =   1425
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Garantia:"
                  Height          =   195
                  Index           =   1
                  Left            =   2970
                  TabIndex        =   164
                  Top             =   870
                  Width           =   645
               End
               Begin VB.Label Label1 
                  Caption         =   "Entidad Financiera:"
                  Height          =   225
                  Left            =   210
                  TabIndex        =   163
                  Top             =   540
                  Width           =   1425
               End
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  Caption         =   "Costo (%):"
                  Height          =   195
                  Index           =   1
                  Left            =   7350
                  TabIndex        =   162
                  Top             =   870
                  Width           =   705
               End
               Begin VB.Label Label45 
                  Caption         =   "Plazo (meses):"
                  Height          =   195
                  Index           =   1
                  Left            =   9120
                  TabIndex        =   161
                  Top             =   870
                  Width           =   1065
               End
               Begin VB.Label Label46 
                  Caption         =   "Comentario:"
                  Height          =   225
                  Left            =   210
                  TabIndex        =   160
                  Top             =   1200
                  Width           =   1425
               End
               Begin VB.Label Label52 
                  Caption         =   "Modalidad Evalua.:"
                  Height          =   225
                  Left            =   210
                  TabIndex        =   159
                  Top             =   1530
                  Width           =   1455
               End
               Begin VB.Label Label68 
                  Caption         =   "Modalidad Evalua.:"
                  Height          =   225
                  Left            =   -74790
                  TabIndex        =   158
                  Top             =   1530
                  Width           =   1455
               End
               Begin VB.Label Label69 
                  Caption         =   "Comentario:"
                  Height          =   225
                  Left            =   -74790
                  TabIndex        =   157
                  Top             =   1200
                  Width           =   1425
               End
               Begin VB.Label Label70 
                  AutoSize        =   -1  'True
                  Caption         =   "Plazo (meses):"
                  Height          =   195
                  Left            =   -65880
                  TabIndex        =   156
                  Top             =   870
                  Width           =   1020
               End
               Begin VB.Label Label71 
                  AutoSize        =   -1  'True
                  Caption         =   "Costo (%):"
                  Height          =   195
                  Left            =   -67650
                  TabIndex        =   155
                  Top             =   870
                  Width           =   705
               End
               Begin VB.Label Label72 
                  Caption         =   "Entidad Financiera:"
                  Height          =   225
                  Left            =   -74790
                  TabIndex        =   154
                  Top             =   540
                  Width           =   1425
               End
               Begin VB.Label Label73 
                  AutoSize        =   -1  'True
                  Caption         =   "Garantia:"
                  Height          =   195
                  Left            =   -72030
                  TabIndex        =   153
                  Top             =   870
                  Width           =   645
               End
               Begin VB.Label Label74 
                  Caption         =   "Tasa de Proyecto:"
                  Height          =   195
                  Left            =   -74790
                  TabIndex        =   152
                  Top             =   870
                  Width           =   1425
               End
            End
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   3435
            Left            =   -74880
            TabIndex        =   166
            Top             =   420
            Width           =   11235
            _Version        =   65536
            _ExtentX        =   19817
            _ExtentY        =   6059
            _StockProps     =   15
            Caption         =   "SSPanel9"
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
            Begin MSFlexGridLib.MSFlexGrid grdLstAsignados 
               Height          =   2775
               Left            =   2310
               TabIndex        =   167
               Top             =   450
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   4895
               _Version        =   393216
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin VB.Label Label43 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Lista de Consejero(s) Hipotecario(s) asignados al Proyecto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2310
               TabIndex        =   168
               Top             =   180
               Width           =   6360
            End
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   3465
            Left            =   120
            TabIndex        =   169
            Top             =   420
            Width           =   11235
            _Version        =   65536
            _ExtentX        =   19817
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
            Begin VB.Frame Frame3 
               Caption         =   "Datos - Bono Verde"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3135
               Left            =   5820
               TabIndex        =   188
               Top             =   220
               Width           =   5295
               Begin VB.TextBox txt_ValAfe 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   2400
                  MaxLength       =   250
                  TabIndex        =   182
                  Text            =   "0.00"
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.ComboBox Cmb_AfeBVe 
                  Height          =   315
                  Left            =   2400
                  Style           =   2  'Dropdown List
                  TabIndex        =   180
                  Top             =   480
                  Width           =   1245
               End
               Begin VB.ComboBox cmb_TipAfe 
                  Height          =   315
                  Left            =   2400
                  Style           =   2  'Dropdown List
                  TabIndex        =   181
                  Top             =   960
                  Width           =   2325
               End
               Begin VB.Label Label47 
                  AutoSize        =   -1  'True
                  Caption         =   "Afecto:"
                  Height          =   195
                  Index           =   2
                  Left            =   480
                  TabIndex        =   191
                  Top             =   525
                  Width           =   510
               End
               Begin VB.Label Label47 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo de Afectación:"
                  Height          =   195
                  Index           =   3
                  Left            =   480
                  TabIndex        =   190
                  Top             =   1005
                  Width           =   1395
               End
               Begin VB.Label Label47 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Afecto:"
                  Height          =   195
                  Index           =   4
                  Left            =   480
                  TabIndex        =   189
                  Top             =   1485
                  Width           =   915
               End
            End
            Begin VB.Frame Frame4 
               Height          =   3135
               Left            =   200
               TabIndex        =   173
               Top             =   220
               Width           =   5535
               Begin VB.ComboBox cmb_AproFile 
                  Height          =   315
                  ItemData        =   "AteCli_frm_134.frx":1F33
                  Left            =   3600
                  List            =   "AteCli_frm_134.frx":1F35
                  Style           =   2  'Dropdown List
                  TabIndex        =   179
                  Top             =   2400
                  Width           =   1425
               End
               Begin VB.CommandButton cmd_DetFec 
                  Caption         =   "...."
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   5070
                  TabIndex        =   174
                  Top             =   480
                  Width           =   345
               End
               Begin EditLib.fpDateTime ipp_FecLim 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   175
                  Top             =   480
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   -1  'True
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   ""
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecInfInm 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   176
                  Top             =   990
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   -1  'True
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   ""
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecInfLeg 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   177
                  Top             =   1440
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   -1  'True
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   ""
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecRevApr 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   178
                  Top             =   1920
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
                  AlignTextV      =   0
                  AllowNull       =   -1  'True
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   ""
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin VB.Label Label75 
                  AutoSize        =   -1  'True
                  Caption         =   "Aprobacion del File Master:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   187
                  Top             =   2460
                  Width           =   1920
               End
               Begin VB.Label Label76 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha de Revision y Aprobacion por Comite:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   186
                  Top             =   1980
                  Width           =   3150
               End
               Begin VB.Label Label77 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha de Informe Legal:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   185
                  Top             =   1500
                  Width           =   1725
               End
               Begin VB.Label Label78 
                  Caption         =   "Fecha de Informe Inmobiliario:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   184
                  Top             =   1020
                  Width           =   2115
               End
               Begin VB.Label Label79 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Limite de Registro Proyecto en RRPP:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   183
                  Top             =   540
                  Width           =   3195
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frm_PryNvi_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_FlgCmb        As Integer
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_arr_Bancos()      As moddat_tpo_Genera
Dim l_arr_BcoComp1()    As moddat_tpo_Genera
Dim l_arr_BcoComp2()    As moddat_tpo_Genera

Private Sub CargarConsejerosAsignados()
Dim r_int_Cont As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EJECMC_CODEJE, (TRIM(EJECMC_APEPAT)||' '||TRIM(EJECMC_APEMAT)||' '||TRIM(EJECMC_NOMBRE)) CONSEJERO "
   g_str_Parame = g_str_Parame & "   FROM CRE_EJECMC "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_ASGCON ON TRIM(CRE_EJECMC.EJECMC_CODEJE) = TRIM(PRY_ASGCON.ASGCON_CONHIP)"
   g_str_Parame = g_str_Parame & "  WHERE ASGCON_CODPRY = '" & moddat_g_str_Codigo & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_Cont = 1
   grdLstAsignados.Rows = 1
   
   If g_rst_Princi.EOF And g_rst_Princi.BOF Then grdLstAsignados.Rows = 2
   Do While Not g_rst_Princi.EOF
      grdLstAsignados.Rows = grdLstAsignados.Rows + 1
      grdLstAsignados.TextMatrix(r_int_Cont, 0) = g_rst_Princi!EJECMC_CODEJE
      grdLstAsignados.TextMatrix(r_int_Cont, 1) = g_rst_Princi!CONSEJERO
      
      r_int_Cont = r_int_Cont + 1
      g_rst_Princi.MoveNext
   Loop
End Sub

Private Function CargarBanco(CodigoBanco As String) As String
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARDES WHERE PARDES_CODGRP = 513 AND PARDES_CODITE = '" & CodigoBanco & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not g_rst_Princi.EOF And Not g_rst_Princi.BOF Then
      CargarBanco = Trim(g_rst_Princi!PARDES_DESCRI)
   End If
End Function

Private Sub Cmb_AfeBVe_Click()
    If Cmb_AfeBVe.ListIndex = 1 Then
        cmb_TipAfe.Enabled = False
        txt_ValAfe.Enabled = False
        cmb_TipAfe.ListIndex = -1
        txt_ValAfe.Text = 0#
    Else
        cmb_TipAfe.Enabled = True
        txt_ValAfe.Enabled = True
        Call gs_SetFocus(cmb_TipAfe)
    End If
End Sub

Private Sub Cmb_AfeBVe_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipAfe)
  End If
End Sub

Private Sub cmb_AproFile_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SSTab1.Tab = 3
      Call gs_SetFocus(grdLstAsignados)
   End If
End Sub

Private Sub cmb_ConAso_Click()
   Call gs_SetFocus(cmb_TipoBien)
End Sub

Private Sub cmb_ConAso_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConAso_Click
   End If
End Sub

Private Sub cmb_TipAfe_Click()
   Call gs_SetFocus(txt_ValAfe)
End Sub

Private Sub cmb_TipAfe_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ValAfe)
  End If
End Sub

Private Sub cmb_TipoBien_Click()
   Call gs_SetFocus(cmb_TipVia)
End Sub

Private Sub cmb_TipoBien_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipoBien_Click
   End If
End Sub

Private Sub cmb_CostoBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PlazoBco1)
   End If
End Sub

Private Sub cmb_CostoBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PlazoBco2)
   End If
End Sub

Private Sub cmb_CtaIni1Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Plazo1Bco1)
   End If
End Sub

Private Sub cmb_CtaIni1Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Plazo1Bco2)
   End If
End Sub

Private Sub cmb_CtaIni2Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Plazo2Bco1)
   End If
End Sub

Private Sub cmb_CtaIni2Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Plazo2Bco2)
   End If
End Sub

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
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
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

Private Sub cmb_DptDir_LostFocus()
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_EntFin_Click()
   cmb_EntFinBco1.Text = cmb_EntFin.Text
End Sub

Private Sub cmb_EntFinBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TasaBco2)
   ElseIf KeyAscii = 8 Then
      cmb_EntFinBco2.ListIndex = -1
   End If
End Sub

Private Sub cmb_Plazo1Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ValIng1Bco1)
   End If
End Sub

Private Sub cmb_Plazo1Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ValIng1Bco2)
   End If
End Sub

Private Sub cmb_Plazo2Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ValIng2Bco1)
   End If
End Sub

Private Sub cmb_Plazo2Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ValIng2Bco2)
   End If
End Sub

Private Sub cmb_PlazoBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ComeBco1)
   End If
End Sub

Private Sub cmb_PlazoBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ComeBco2)
   End If
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
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
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

Private Sub cmb_PrvDir_LostFocus()
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
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

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmb_ConAso)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_TipoGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SSTab1.Tab = 2
      Call gs_SetFocus(ipp_FecLim)
   End If
End Sub

Private Sub cmb_TipoGarBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CostoBco1)
   ElseIf KeyAscii = 8 Then
      cmb_TipoGarBco1.ListIndex = -1
   End If
End Sub

Private Sub cmb_TipoGarBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CostoBco2)
   ElseIf KeyAscii = 8 Then
      cmb_TipoGarBco2.ListIndex = -1
   End If
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

Private Sub cmb_ValIng1Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com1Bco1)
   End If
End Sub

Private Sub cmb_ValIng1Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com1Bco2)
   End If
End Sub

Private Sub cmb_ValIng2Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com2Bco1)
   End If
End Sub

Private Sub cmb_ValIng2Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Com2Bco2)
   End If
End Sub

Private Sub cmd_Buscar_Con_Click()
   If cmb_TipDoc_Con.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc_Con)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc_Con.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc_Con)
      Exit Sub
   End If
   If cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc_Con.Text)) <> 11 Then
         MsgBox "El Número de Documento de Identidad no tiene 11 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Con)
         Exit Sub
      End If
      If Not gf_Valida_RUC(Trim(txt_NumDoc_Con.Text), Mid(Trim(txt_NumDoc_Con.Text), Len(Trim(txt_NumDoc_Con.Text)), 1)) Then
         MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Con)
         Exit Sub
      End If
   End If
   
   'Buscar Empresa en Maestro de Empresas
   moddat_g_int_TipDoc = CStr(cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex))
   moddat_g_str_NumDoc = Trim(txt_NumDoc_Con.Text)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      cmd_Editar_Con.Enabled = False
      moddat_g_int_FlgGrb = 1
      moddat_g_int_FlgAct = 1
      
      frm_PryNVi_03.Show 1
      
      If moddat_g_int_FlgAct = 1 Then
         Exit Sub
      End If
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call fs_Buscar_EmpCon
End Sub

Private Sub cmd_Buscar_Pro_Click()
   If cmb_TipDoc_Pro.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc_Pro)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc_Pro.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc_Pro)
      Exit Sub
   End If
   If cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc_Pro.Text)) <> 11 Then
         MsgBox "El Número de Documento de Identidad no tiene 11 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Pro)
         Exit Sub
      End If
      If Not gf_Valida_RUC(Trim(txt_NumDoc_Pro.Text), Mid(Trim(txt_NumDoc_Pro.Text), Len(Trim(txt_NumDoc_Pro.Text)), 1)) Then
         MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Pro)
         Exit Sub
      End If
   End If
   
   'Buscar Empresa en Maestro de Empresas
   moddat_g_int_TipDoc = CStr(cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex))
   moddat_g_str_NumDoc = Trim(txt_NumDoc_Pro.Text)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      cmd_Editar_Pro.Enabled = False
      moddat_g_int_FlgGrb = 1
      moddat_g_int_FlgAct = 1
      
      frm_PryNVi_03.Show 1
      
      If moddat_g_int_FlgAct = 1 Then
         Exit Sub
      End If
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call fs_Buscar_EmpPro
End Sub

Private Sub cmd_DetFec_Click()
   Dim r_rst_Fecha   As ADODB.Recordset
   
   frm_PryNVi_04.Show 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DATGENRP_FECREG "
   g_str_Parame = g_str_Parame & "  FROM PRY_DATGENRP "
   g_str_Parame = g_str_Parame & " WHERE DATGENRP_CODPRY = '" & moddat_g_str_Codigo & "'"
   g_str_Parame = g_str_Parame & " ORDER BY DATGENRP_FECREG DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Fecha, 3) Then
      Exit Sub
   End If
   
   If Not r_rst_Fecha.EOF And Not r_rst_Fecha.BOF Then
      ipp_FecLim.Text = gf_FormatoFecha(CStr(r_rst_Fecha!DATGENRP_FECREG))
   Else
      ipp_FecLim.Text = ""
   End If
End Sub

Private Sub cmd_Editar_Con_Click()
   'Buscar Empresa en Maestro de Empresas
   moddat_g_int_TipDoc = CStr(cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex))
   moddat_g_str_NumDoc = Trim(txt_NumDoc_Con.Text)
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_PryNVi_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_EmpCon
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Editar_Pro_Click()
   'Buscar Empresa en Maestro de Empresas
   moddat_g_int_TipDoc = CStr(cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex))
   moddat_g_str_NumDoc = Trim(txt_NumDoc_Pro.Text)
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_PryNVi_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar_EmpPro
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_CodPry     As String
Dim r_str_Item       As String
Dim r_str_Item2      As String
Dim r_rst_Genera     As ADODB.Recordset

   If moddat_g_int_TipCli = 2 Then
      If cmb_EntFin.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Entidad Financiera.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_EntFin)
         Exit Sub
      End If
   End If
   If Len(Trim(txt_NomPry.Text)) = 0 Then
      MsgBox "Debe ingresar el nombre del Proyecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomPry)
      Exit Sub
   End If
   If Len(Trim(txt_DesPry.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción del Proyecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DesPry)
      Exit Sub
   End If
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación del Proyecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   If cmb_TipoBien.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Bien del Proyecto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipoBien)
      Exit Sub
   End If
   If cmb_ConAso.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Proyecto tiene Convenio de Asociación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConAso)
      Exit Sub
   End If
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVia)
      Exit Sub
   End If
   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de la Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomVia)
      Exit Sub
   End If
   If Len(Trim(txt_NumVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumVia)
      Exit Sub
   End If
   If cmb_TipZon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipZon)
      Exit Sub
   End If
   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) = 12 Then
      If Len(Trim(txt_NomZon.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomZon)
         Exit Sub
      End If
   End If
   If cmb_DptDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DptDir)
      Exit Sub
   End If
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvDir)
      Exit Sub
   End If
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DstDir)
      Exit Sub
   End If
   If grd_Listad_Pro.Rows = 0 Then
      MsgBox "Debe ingresar el Promotor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc_Pro)
      Exit Sub
   End If
   If grd_Listad_Con.Rows = 0 Then
      MsgBox "Debe ingresar el Constructor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc_Con)
      Exit Sub
   End If
   '------------
   If Cmb_AfeBVe.ListIndex = -1 Then
      MsgBox "Debe seleccionar si afecta Bono Verde.", vbExclamation, modgen_g_str_NomPlt
      SSTab1.Tab = 2
      Call gs_SetFocus(Cmb_AfeBVe)
      Exit Sub
   End If
   If Cmb_AfeBVe.ListIndex = 0 And cmb_TipAfe.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Afectación correspondiente al Bono Verde .", vbExclamation, modgen_g_str_NomPlt
      SSTab1.Tab = 2
      Call gs_SetFocus(cmb_TipAfe)
      Exit Sub
   End If
   If cmb_TipAfe.ListIndex <> -1 And txt_ValAfe.Text = "" Then
      MsgBox "Debe ingresar el valor de afectación del Bono Verde.", vbExclamation, modgen_g_str_NomPlt
      SSTab1.Tab = 2
      Call gs_SetFocus(txt_ValAfe)
      Exit Sub
   End If
   
   If cmb_TipoGar.Text <> "" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT PARDES_CODITE "
      g_str_Parame = g_str_Parame & "  FROM MNT_PARDES "
      g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP = '241' "
      g_str_Parame = g_str_Parame & "   AND TRIM(PARDES_DESCRI) = '" & cmb_TipoGar.Text & "'"

      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
         Exit Sub
      End If

      If Not r_rst_Genera.EOF And Not r_rst_Genera.BOF Then
         r_str_Item = r_rst_Genera!PARDES_CODITE
      Else
         r_str_Item = ""
      End If
   End If
   
   If cmb_TipoGarBco1.Text <> "" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT PARDES_CODITE FROM MNT_PARDES "
      g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP= '241' AND TRIM(PARDES_DESCRI) = '" & cmb_TipoGarBco1.Text & "'"

      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
          Exit Sub
      End If

      If Not r_rst_Genera.EOF And Not r_rst_Genera.BOF Then
         r_str_Item = r_rst_Genera!PARDES_CODITE
      Else
         r_str_Item = ""
      End If
   End If
   
   If cmb_TipoGarBco2.Text <> "" Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT PARDES_CODITE FROM MNT_PARDES "
      g_str_Parame = g_str_Parame & " WHERE PARDES_CODGRP= '241' AND TRIM(PARDES_DESCRI) = '" & cmb_TipoGarBco2.Text & "'"

      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
          Exit Sub
      End If

      If Not r_rst_Genera.EOF And Not r_rst_Genera.BOF Then
         r_str_Item2 = r_rst_Genera!PARDES_CODITE
      Else
         r_str_Item2 = ""
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb_1 = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT MAX(DATGEN_CODIGO) AS ULTIMO "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN "
      g_str_Parame = g_str_Parame & " WHERE SUBSTR(DATGEN_CODIGO,1,3) = '" & Right(Format(Year(date), "0000"), 3) & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Len(Trim(g_rst_Princi!ULTIMO & "")) = 0 Then
         r_str_CodPry = Right(Format(Year(date), "0000"), 3) & "501"
      Else
         r_str_CodPry = Right(Format(Year(date), "0000"), 3) & Format(CInt(Right(g_rst_Princi!ULTIMO, 3)) + 1, "000")
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      r_str_CodPry = moddat_g_str_Codigo
   End If

   'Grabando Información del Cliente
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      g_str_Parame = ""
      g_str_Parame = "USP_PRY_DATGEN_NVI ("
      g_str_Parame = g_str_Parame & "'" & r_str_CodPry & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NomPry.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_DesPry.Text & "', "
      g_str_Parame = g_str_Parame & " " & moddat_g_int_TipCli & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Pro.ItemData(cmb_TipDoc_Pro.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc_Pro.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Con.ItemData(cmb_TipDoc_Con.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc_Con.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      If moddat_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & "'999999', "
         g_str_Parame = g_str_Parame & "2, "
      Else
         g_str_Parame = g_str_Parame & "'" & CStr(l_arr_Bancos(cmb_EntFin.ListIndex + 1).Genera_Codigo) & "' , "
         g_str_Parame = g_str_Parame & CStr(cmb_ConAso.ItemData(cmb_ConAso.ListIndex)) & ", "
      End If
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Interi.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      If Len(Trim(ipp_FecIni.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "Null, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_FecIni.Text, "YYYYMMDD") & ", "
      End If
      If Len(Trim(ipp_FecFin.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "Null, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_FecFin.Text, "YYYYMMDD") & ", "
      End If
      If Len(Trim(ipp_FecLim.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "Null, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_FecLim.Text, "YYYYMMDD") & ", "
      End If
      If Len(Trim(ipp_FecInfInm.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "Null, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_FecInfInm.Text, "YYYYMMDD") & ", "
      End If
      If Len(Trim(ipp_FecInfLeg.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "Null, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_FecInfLeg.Text, "YYYYMMDD") & ", "
      End If
      If Len(Trim(ipp_FecRevApr.Text)) = 0 Then
         g_str_Parame = g_str_Parame & "Null, "
      Else
         g_str_Parame = g_str_Parame & Format(ipp_FecRevApr.Text, "YYYYMMDD") & ", "
      End If
      If cmb_AproFile.ListIndex = -1 Then
         g_str_Parame = g_str_Parame & "0, "
      Else
         g_str_Parame = g_str_Parame & CStr(cmb_AproFile.ItemData(cmb_AproFile.ListIndex)) & ", "
      End If
      
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & txt_Cargo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Email.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telefono.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Contacto.Text & "', "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_PreMin.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_PreMax.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_AreaMin.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_AreaMax.Text) & ", "
      g_str_Parame = g_str_Parame & "" & ipp_Tasa.Text & ", "
      g_str_Parame = g_str_Parame & "'" & ipp_NumEtapa.Text & "', "
      g_str_Parame = g_str_Parame & "" & CInt(ipp_TotUni.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CInt(ipp_TotVen.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CInt(ipp_TotDisp.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_Avance.Text) & ", "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_Dispon.Text) & ", "
      g_str_Parame = g_str_Parame & "" & ipp_Coloca.Text & ", "
      g_str_Parame = g_str_Parame & "" & CDbl(ipp_Partic.Text) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Item & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipoBien.ItemData(cmb_TipoBien.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(Cmb_AfeBVe.ItemData(Cmb_AfeBVe.ListIndex)) & ", "
      If Cmb_AfeBVe.ListIndex = 1 Then
        g_str_Parame = g_str_Parame & "Null, "
        g_str_Parame = g_str_Parame & "Null, "
      Else
        g_str_Parame = g_str_Parame & CStr(cmb_TipAfe.ItemData(cmb_TipAfe.ListIndex)) & ", "
        g_str_Parame = g_str_Parame & "" & CDbl(txt_ValAfe.Text) & ", "
      End If
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_FlgGrb_1) & "')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      'Grabar Competencias
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT DATCOM_CODIGO "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATCOM "
      g_str_Parame = g_str_Parame & " WHERE DATCOM_CODIGO = '" & r_str_CodPry & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If (g_rst_Princi.EOF And g_rst_Princi.BOF) Then
         moddat_g_int_FlgGrb_1 = 1 'Nuevo
      Else
         moddat_g_int_FlgGrb_1 = 2 'Modificar
      End If
      
      g_str_Parame = ""
      g_str_Parame = "USP_PRY_DATCOM ("
      g_str_Parame = g_str_Parame & "'" & r_str_CodPry & "', "
      If moddat_g_int_TipCli = 2 Then
         g_str_Parame = g_str_Parame & "'" & CStr(l_arr_Bancos(cmb_EntFin.ListIndex + 1).Genera_Codigo) & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & CStr(l_arr_BcoComp1(cmb_EntFinBco1.ListIndex + 1).Genera_Codigo) & "', "
      End If
      g_str_Parame = g_str_Parame & " " & CDbl(ipp_TasaBco1.Text) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Item & "', "
      If Len(Trim(cmb_CostoBco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "" & cmb_CostoBco1.Text & ", "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_PlazoBco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "" & cmb_PlazoBco1.Text & ", "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ComeBco1.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ModEvaBco1.Text) & "' , "
      If Len(Trim(cmb_CtaIni1Bco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_CtaIni1Bco1.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_Plazo1Bco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_Plazo1Bco1.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_ValIng1Bco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_ValIng1Bco1.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_Com1Bco1.Text) & "', "
      If Len(Trim(cmb_CtaIni2Bco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_CtaIni2Bco1.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_Plazo2Bco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_Plazo2Bco1.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_ValIng1Bco1.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_ValIng2Bco1.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_Com2Bco1.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(l_arr_BcoComp2(cmb_EntFinBco2.ListIndex + 1).Genera_Codigo) & "', "
      g_str_Parame = g_str_Parame & " " & CDbl(ipp_TasaBco2.Text) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Item2 & "', "
      If Len(Trim(cmb_CostoBco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "" & cmb_CostoBco2.Text & ", "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_PlazoBco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "" & cmb_PlazoBco2.Text & ", "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ComeBco2.Text) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(txt_ModEvaBco2.Text) & "', "
      If Len(Trim(cmb_CtaIni1Bco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_CtaIni1Bco2.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_Plazo1Bco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_Plazo1Bco2.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_ValIng1Bco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_ValIng1Bco2.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_Com1Bco2.Text) & "', "
      If Len(Trim(cmb_CtaIni2Bco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_CtaIni2Bco2.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_Plazo2Bco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_Plazo2Bco2.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      If Len(Trim(cmb_ValIng2Bco2.Text)) > 0 Then
         g_str_Parame = g_str_Parame & "'" & cmb_ValIng2Bco2.Text & "', "
      Else
         g_str_Parame = g_str_Parame & " 0, "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(txt_Com2Bco2.Text) & "', "
      g_str_Parame = g_str_Parame & " " & moddat_g_int_FlgGrb_1 & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
'      If moddat_g_int_CntErr = 6 Then
'         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'            Exit Sub
'         Else
'            moddat_g_int_CntErr = 0
'         End If
'      End If
      
      Screen.MousePointer = 0
   Loop

   moddat_g_int_FlgAct_1 = 2
   MsgBox "Los datos se registraron correctamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Limpia_Con_Click()
   Call fs_Limpia_Con
   Call fs_Activa_Con(True)
   Call gs_SetFocus(cmb_TipDoc_Con)
End Sub

Private Sub cmd_Limpia_Pro_Click()
   Call fs_Limpia_Pro
   Call fs_Activa_Pro(True)
   
   Call gs_SetFocus(cmb_TipDoc_Pro)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_rst_Princi  As ADODB.Recordset
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_TipCli = 1 Then
      pnl_TituloForm.Caption = "Gestión de Proyectos Vinculados"
   Else
      pnl_TituloForm.Caption = "Gestión de Proyectos No Vinculados"
   End If
   
   If moddat_g_int_FlgGrb_1 = 2 Then
      '***************
      'DATOS GENERALES
      '***************
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN "
      g_str_Parame = g_str_Parame & " WHERE DATGEN_CODIGO = '" & moddat_g_str_Codigo & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
         Exit Sub
      End If
         
      r_rst_Princi.MoveFirst
      txt_NomPry.Text = Trim(r_rst_Princi!DATGEN_TITULO & "")
      txt_DesPry.Text = Trim(r_rst_Princi!DATGEN_DESCRI & "")
      If moddat_g_int_TipCli = 2 Then
         cmb_EntFin.Text = CargarBanco(r_rst_Princi!DATGEN_CODBCO)
      End If
      Call gs_BuscarCombo_Item(cmb_Situac, r_rst_Princi!DATGEN_SITUAC)
      If Not IsNull(r_rst_Princi!DATGEN_TIPBIE) Then
         Call gs_BuscarCombo_Item(cmb_TipoBien, r_rst_Princi!DATGEN_TIPBIE)
      End If
      Call gs_BuscarCombo_Item(cmb_ConAso, r_rst_Princi!DATGEN_FLGCON)
      Call gs_BuscarCombo_Item(cmb_TipVia, r_rst_Princi!DatGen_TipVia)
      txt_NomVia.Text = Trim(r_rst_Princi!DatGen_NomVia & "")
      txt_NumVia.Text = Trim(r_rst_Princi!DatGen_numVia & "")
      txt_Interi.Text = Trim(r_rst_Princi!DATGEN_INTDPT & "")
      Call gs_BuscarCombo_Item(cmb_TipZon, r_rst_Princi!DatGen_TipZon)
      txt_NomZon.Text = Trim(r_rst_Princi!DatGen_NomZon & "")
      txt_Refere.Text = Trim(r_rst_Princi!DATGEN_REFERE & "")
      
      'datos para el Tab 2
      txt_Cargo.Text = Trim(r_rst_Princi!DATGEN_NOMCAR & "")
      txt_Email.Text = Trim(r_rst_Princi!DATGEN_CORREO & "")
      txt_Telefono.Text = Trim(r_rst_Princi!DatGen_Telefo & "")
      txt_Contacto.Text = Trim(r_rst_Princi!DATGEN_CONTAC & "")
      ipp_PreMin.Text = Trim(r_rst_Princi!DATGEN_PREMIN & "")
      ipp_PreMax.Text = Trim(r_rst_Princi!DATGEN_PREMAX & "")
      ipp_AreaMin.Text = Trim(r_rst_Princi!DATGEN_AREMIN & "")
      ipp_AreaMax.Text = Trim(r_rst_Princi!DATGEN_AREMAX & "")
      
      If Not IsNull(r_rst_Princi!DATGEN_INIOBR) Then ipp_FecIni.Text = gf_FormatoFecha(CStr(r_rst_Princi!DATGEN_INIOBR))
      If Not IsNull(r_rst_Princi!DATGEN_FINOBR) Then ipp_FecFin.Text = gf_FormatoFecha(CStr(r_rst_Princi!DATGEN_FINOBR))
      If Not IsNull(r_rst_Princi!DATGEN_FECLIM) Then ipp_FecLim.Text = gf_FormatoFecha(CStr(r_rst_Princi!DATGEN_FECLIM))
      If Not IsNull(r_rst_Princi!DATGEN_FECINFINM) Then ipp_FecInfInm.Text = gf_FormatoFecha(CStr(r_rst_Princi!DATGEN_FECINFINM))
      If Not IsNull(r_rst_Princi!DATGEN_FECINFLEG) Then ipp_FecInfLeg.Text = gf_FormatoFecha(CStr(r_rst_Princi!DATGEN_FECINFLEG))
      If Not IsNull(r_rst_Princi!DATGEN_FECAPRCOM) Then ipp_FecRevApr.Text = gf_FormatoFecha(CStr(r_rst_Princi!DATGEN_FECAPRCOM))
      Call gs_BuscarCombo_Item(cmb_AproFile, CInt(r_rst_Princi!DATGEN_FLGFILMAS))

      ipp_TotUni.Text = Trim(r_rst_Princi!DATGEN_TOTUNI & "")
      ipp_TotVen.Text = Trim(r_rst_Princi!DATGEN_TOTVEN & "")
      ipp_TotDisp.Text = Trim(r_rst_Princi!DATGEN_TOTDIS & "")
      ipp_Dispon.Text = Trim(r_rst_Princi!DATGEN_DISPON & "")
      ipp_Avance.Text = Trim(r_rst_Princi!DATGEN_AVANCE & "")
      
      'ipp_Coloca.Text = fs_ObtieneOperaciones_Proyecto(moddat_g_str_Codigo)
      'If ipp_TotUni.Text > 0 Then
      '   If ipp_Coloca.Text > 0 Then
      '      ipp_Partic.Text = Format((ipp_Coloca.Text / ipp_TotUni.Text) * 100, "##0.00")
      '   Else
      '      ipp_Partic.Text = 0
      '   End If
      'Else
      '   ipp_Partic.Text = 0
      'End If
      
      ipp_NumEtapa.Text = Trim(r_rst_Princi!DATGEN_ETAPAS & "")
      ipp_Tasa.Text = Trim(r_rst_Princi!DATGEN_TASA & "")
      ipp_Partic.Text = Trim(r_rst_Princi!DATGEN_PARTIC & "")
      ipp_Coloca.Text = Trim(r_rst_Princi!DATGEN_COLOCA & "")
      
      If Not IsNull(r_rst_Princi!DATGEN_GARANT) And (r_rst_Princi!DATGEN_GARANT > 0) Then
         cmb_TipoGar.Text = moddat_gf_Consulta_ParDes("241", CStr(r_rst_Princi!DATGEN_GARANT))
      End If
           
      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(r_rst_Princi!DatGen_Ubigeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(r_rst_Princi!DatGen_Ubigeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(r_rst_Princi!DatGen_Ubigeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(r_rst_Princi!DatGen_Ubigeo, 2), Mid(r_rst_Princi!DatGen_Ubigeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(r_rst_Princi!DatGen_Ubigeo, 2)))
      
      Call gs_BuscarCombo_Item(cmb_TipDoc_Pro, r_rst_Princi!DATGEN_VENTDO)
      txt_NumDoc_Pro.Text = Trim(r_rst_Princi!DATGEN_VENNDO & "")
      Call cmd_Buscar_Pro_Click
      
      Call gs_BuscarCombo_Item(cmb_TipDoc_Con, r_rst_Princi!DATGEN_CONTDO)
      txt_NumDoc_Con.Text = Trim(r_rst_Princi!DATGEN_CONNDO & "")
      Call cmd_Buscar_Con_Click
      
      Call CargarConsejerosAsignados
                  
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      
      '***************
      'DATOS BONO VERDE
      '***************
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT PY.*, "
      g_str_Parame = g_str_Parame & "       (SELECT BC.PARDES_DESCRI FROM MNT_PARDES BC "
      g_str_Parame = g_str_Parame & "         WHERE BC.PARDES_CODGRP = 214 AND BC.PARDES_CODITE = PY.DATGEN_FLGAFEBV) AS BONO_VERDE, "
      g_str_Parame = g_str_Parame & "       (SELECT BC.PARDES_DESCRI FROM MNT_PARDES BC "
      g_str_Parame = g_str_Parame & "         WHERE BC.PARDES_CODGRP = 278 AND BC.PARDES_CODITE = PY.DATGEN_FLGTIPAFE) AS TIPO_AFEC "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATGEN PY "
      g_str_Parame = g_str_Parame & " WHERE PY.DATGEN_CODIGO = '" & moddat_g_str_Codigo & "' "

      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
         Exit Sub
      End If
      If Not (r_rst_Princi.EOF And r_rst_Princi.BOF) Then
        r_rst_Princi.MoveFirst
        If Not IsNull(r_rst_Princi!DATGEN_FLGAFEBV) Then Cmb_AfeBVe.Text = Trim(r_rst_Princi!BONO_VERDE)
        If Not IsNull(r_rst_Princi!DATGEN_FLGTIPAFE) Then cmb_TipAfe.Text = Trim(r_rst_Princi!TIPO_AFEC)
        If Not IsNull(r_rst_Princi!DATGEN_VALAFEBV) Then txt_ValAfe.Text = Trim(r_rst_Princi!DATGEN_VALAFEBV)
      End If
      
      'Verifica su es proyecto vinculado
      If moddat_g_int_TipCli = 1 Then
         Label18.Visible = False
         cmb_EntFin.Visible = False
         cmb_EntFinBco1.Enabled = True
      'Verifica su es proyecto no vinculado
      ElseIf moddat_g_int_TipCli = 2 Then
         Label18.Visible = True
         cmb_EntFin.Visible = True
         cmb_EntFinBco1.Enabled = False
      End If
      
      '******************
      'DATOS COMPETENCIAS
      '******************
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM PRY_DATCOM "
      g_str_Parame = g_str_Parame & " WHERE DATCOM_CODIGO = '" & moddat_g_str_Codigo & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
         Exit Sub
      End If
         
      If Not (r_rst_Princi.EOF And r_rst_Princi.BOF) Then
         If Not IsNull(r_rst_Princi!DATCOM_ENTFIN1) Then cmb_EntFinBco1.Text = CargarBanco(r_rst_Princi!DATCOM_ENTFIN1)
         ipp_TasaBco1.Text = r_rst_Princi!DATCOM_TASPRY1
         If Not IsNull(r_rst_Princi!DATCOM_TIPGAR1) Then cmb_TipoGarBco1.Text = moddat_gf_Consulta_ParDes("241", CStr(r_rst_Princi!DATCOM_TIPGAR1))
         If Not IsNull(r_rst_Princi!DATCOM_COSTO1) And (r_rst_Princi!DATCOM_COSTO1 > 0) Then cmb_CostoBco1.Text = r_rst_Princi!DATCOM_COSTO1
         If Not IsNull(r_rst_Princi!DATCOM_PLZMES1) And (r_rst_Princi!DATCOM_PLZMES1 > 0) Then cmb_PlazoBco1.Text = r_rst_Princi!DATCOM_PLZMES1
         txt_ComeBco1.Text = r_rst_Princi!DATCOM_COMENT1 & ""
         txt_ModEvaBco1.Text = r_rst_Princi!DATCOM_MODEVA1 & ""
         If (Not IsNull(r_rst_Princi!DATCOM_AHCI011)) And (r_rst_Princi!DATCOM_AHCI011 > 0) Then cmb_CtaIni1Bco1.Text = r_rst_Princi!DATCOM_AHCI011
         If Not IsNull(r_rst_Princi!DATCOM_AHPL011) And (r_rst_Princi!DATCOM_AHPL011 > 0) Then cmb_Plazo1Bco1.Text = r_rst_Princi!DATCOM_AHPL011
         If Not IsNull(r_rst_Princi!DATCOM_AHVI011) And (r_rst_Princi!DATCOM_AHVI011 > 0) Then cmb_ValIng1Bco1.Text = r_rst_Princi!DATCOM_AHVI011
         txt_Com1Bco1.Text = r_rst_Princi!DATCOM_COME011 & ""
         If Not IsNull(r_rst_Princi!DATCOM_AHCI021) And (r_rst_Princi!DATCOM_AHCI021 > 0) Then cmb_CtaIni2Bco1.Text = r_rst_Princi!DATCOM_AHCI021
         If Not IsNull(r_rst_Princi!DATCOM_AHPL021) And (r_rst_Princi!DATCOM_AHPL021 > 0) Then cmb_Plazo2Bco1.Text = r_rst_Princi!DATCOM_AHPL021
         If Not IsNull(r_rst_Princi!DATCOM_AHVI021) And (r_rst_Princi!DATCOM_AHVI021 > 0) Then cmb_ValIng2Bco1.Text = r_rst_Princi!DATCOM_AHVI021
         txt_Com2Bco1.Text = r_rst_Princi!DATCOM_COME021 & ""
         
         If Not IsNull(r_rst_Princi!DATCOM_ENTFIN2) Then cmb_EntFinBco2.Text = CargarBanco(r_rst_Princi!DATCOM_ENTFIN2)
         ipp_TasaBco2.Text = r_rst_Princi!DATCOM_TASPRY2
         If Not IsNull(r_rst_Princi!DATCOM_TIPGAR2) Then cmb_TipoGarBco2.Text = moddat_gf_Consulta_ParDes("241", CStr(r_rst_Princi!DATCOM_TIPGAR2))
         If Not IsNull(r_rst_Princi!DATCOM_COSTO2) And (r_rst_Princi!DATCOM_COSTO2 > 0) Then cmb_CostoBco2.Text = r_rst_Princi!DATCOM_COSTO2
         If Not IsNull(r_rst_Princi!DATCOM_PLZMES2) And (r_rst_Princi!DATCOM_PLZMES2 > 0) Then cmb_PlazoBco2.Text = r_rst_Princi!DATCOM_PLZMES2
         txt_ComeBco2.Text = r_rst_Princi!DATCOM_COMENT2 & ""
         txt_ModEvaBco2.Text = r_rst_Princi!DATCOM_MODEVA2 & ""
         If Not IsNull(r_rst_Princi!DATCOM_AHCI012) And (r_rst_Princi!DATCOM_AHCI012 > 0) Then cmb_CtaIni1Bco2.Text = r_rst_Princi!DATCOM_AHCI012
         If Not IsNull(r_rst_Princi!DATCOM_AHPL012) And (r_rst_Princi!DATCOM_AHPL012 > 0) Then cmb_Plazo1Bco2.Text = r_rst_Princi!DATCOM_AHPL012
         If Not IsNull(r_rst_Princi!DATCOM_AHVI012) And (r_rst_Princi!DATCOM_AHVI012 > 0) Then cmb_ValIng1Bco2.Text = r_rst_Princi!DATCOM_AHVI012
         txt_Com1Bco2.Text = r_rst_Princi!DATCOM_COME012 & ""
         If Not IsNull(r_rst_Princi!DATCOM_AHCI022) And (r_rst_Princi!DATCOM_AHCI022 > 0) Then cmb_CtaIni2Bco2.Text = r_rst_Princi!DATCOM_AHCI022
         If Not IsNull(r_rst_Princi!DATCOM_AHPL022) And (r_rst_Princi!DATCOM_AHPL022 > 0) Then cmb_Plazo2Bco2.Text = r_rst_Princi!DATCOM_AHPL022
         If Not IsNull(r_rst_Princi!DATCOM_AHVI022) And (r_rst_Princi!DATCOM_AHVI022 > 0) Then cmb_ValIng2Bco2.Text = r_rst_Princi!DATCOM_AHVI022
         txt_Com2Bco2.Text = r_rst_Princi!DATCOM_COME022 & ""
      End If
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad_Pro.ColWidth(0) = 3200
   grd_Listad_Pro.ColWidth(1) = 7700
   grd_Listad_Pro.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Pro.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Con.ColWidth(0) = 3200
   grd_Listad_Con.ColWidth(1) = 7700
   grd_Listad_Con.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Con.ColAlignment(1) = flexAlignLeftCenter
   
   With grdLstAsignados
      .TextMatrix(0, 0) = "ID"
      .TextMatrix(0, 1) = "Consejero Hipotecario"
      .FixedAlignment(0) = flexAlignCenterCenter
      .FixedAlignment(1) = flexAlignCenterCenter
      .ColWidth(0) = 1500
      .ColWidth(1) = 4500
   End With

   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipoBien, 1, "377")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ConAso, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(Cmb_AfeBVe, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipAfe, 1, "278")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Pro, 1, "232")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Con, 1, "232")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipoGar, 1, "241")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipoGarBco1, 1, "241")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipoGarBco2, 1, "241")
   Call moddat_gs_Carga_LisIte(cmb_EntFin, l_arr_Bancos, 1, "513")
   Call moddat_gs_Carga_LisIte(cmb_EntFinBco1, l_arr_BcoComp1, 1, "513")
   Call moddat_gs_Carga_LisIte(cmb_EntFinBco2, l_arr_BcoComp2, 1, "513")
   Call moddat_gs_Carga_LisIte_Combo(cmb_AproFile, 1, "214")
   
   SSTab1.Tab = 0
   SSTab2.Tab = 0
   
   'Para pestaña Competencias
   cmb_CostoBco1.AddItem ""
   cmb_CostoBco1.AddItem "1"
   cmb_CostoBco1.AddItem "1.5"
   cmb_CostoBco1.AddItem "2"
   cmb_CostoBco1.AddItem "2.5"
   cmb_CostoBco1.AddItem "3"
   cmb_CostoBco1.AddItem "3.5"
   cmb_CostoBco1.AddItem "4"
   cmb_CostoBco1.AddItem "4.5"
   cmb_CostoBco1.AddItem "5"
   cmb_CostoBco2.AddItem ""
   cmb_CostoBco2.AddItem "1"
   cmb_CostoBco2.AddItem "1.5"
   cmb_CostoBco2.AddItem "2"
   cmb_CostoBco2.AddItem "2.5"
   cmb_CostoBco2.AddItem "3"
   cmb_CostoBco2.AddItem "3.5"
   cmb_CostoBco2.AddItem "4"
   cmb_CostoBco2.AddItem "4.5"
   cmb_CostoBco2.AddItem "5"
   
   cmb_PlazoBco1.AddItem ""
   cmb_PlazoBco1.AddItem "6"
   cmb_PlazoBco1.AddItem "12"
   cmb_PlazoBco1.AddItem "24"
   cmb_PlazoBco2.AddItem ""
   cmb_PlazoBco2.AddItem "6"
   cmb_PlazoBco2.AddItem "12"
   cmb_PlazoBco2.AddItem "24"
   
   cmb_CtaIni1Bco1.AddItem ""
   cmb_CtaIni1Bco1.AddItem "10"
   cmb_CtaIni1Bco1.AddItem "15"
   cmb_CtaIni1Bco1.AddItem "20"
   cmb_CtaIni1Bco1.AddItem "25"
   cmb_CtaIni1Bco1.AddItem "30"
   cmb_CtaIni1Bco1.AddItem "35"
   cmb_CtaIni1Bco1.AddItem "40"
   cmb_CtaIni1Bco2.AddItem ""
   cmb_CtaIni1Bco2.AddItem "10"
   cmb_CtaIni1Bco2.AddItem "15"
   cmb_CtaIni1Bco2.AddItem "20"
   cmb_CtaIni1Bco2.AddItem "25"
   cmb_CtaIni1Bco2.AddItem "30"
   cmb_CtaIni1Bco2.AddItem "35"
   cmb_CtaIni1Bco2.AddItem "40"
   
   cmb_CtaIni2Bco1.AddItem ""
   cmb_CtaIni2Bco1.AddItem "10"
   cmb_CtaIni2Bco1.AddItem "15"
   cmb_CtaIni2Bco1.AddItem "20"
   cmb_CtaIni2Bco1.AddItem "25"
   cmb_CtaIni2Bco1.AddItem "30"
   cmb_CtaIni2Bco1.AddItem "35"
   cmb_CtaIni2Bco1.AddItem "40"
   cmb_CtaIni2Bco2.AddItem ""
   cmb_CtaIni2Bco2.AddItem "10"
   cmb_CtaIni2Bco2.AddItem "15"
   cmb_CtaIni2Bco2.AddItem "20"
   cmb_CtaIni2Bco2.AddItem "25"
   cmb_CtaIni2Bco2.AddItem "30"
   cmb_CtaIni2Bco2.AddItem "35"
   cmb_CtaIni2Bco2.AddItem "40"

   cmb_Plazo1Bco1.AddItem ""
   cmb_Plazo1Bco1.AddItem "1"
   cmb_Plazo1Bco1.AddItem "2"
   cmb_Plazo1Bco1.AddItem "3"
   cmb_Plazo1Bco1.AddItem "4"
   cmb_Plazo1Bco1.AddItem "5"
   cmb_Plazo1Bco1.AddItem "6"
   cmb_Plazo1Bco1.AddItem "7"
   cmb_Plazo1Bco1.AddItem "8"
   cmb_Plazo1Bco1.AddItem "9"
   cmb_Plazo1Bco2.AddItem ""
   cmb_Plazo1Bco2.AddItem "1"
   cmb_Plazo1Bco2.AddItem "2"
   cmb_Plazo1Bco2.AddItem "3"
   cmb_Plazo1Bco2.AddItem "4"
   cmb_Plazo1Bco2.AddItem "5"
   cmb_Plazo1Bco2.AddItem "6"
   cmb_Plazo1Bco2.AddItem "7"
   cmb_Plazo1Bco2.AddItem "8"
   cmb_Plazo1Bco2.AddItem "9"
   
   cmb_Plazo2Bco1.AddItem ""
   cmb_Plazo2Bco1.AddItem "1"
   cmb_Plazo2Bco1.AddItem "2"
   cmb_Plazo2Bco1.AddItem "3"
   cmb_Plazo2Bco1.AddItem "4"
   cmb_Plazo2Bco1.AddItem "5"
   cmb_Plazo2Bco1.AddItem "6"
   cmb_Plazo2Bco1.AddItem "7"
   cmb_Plazo2Bco1.AddItem "8"
   cmb_Plazo2Bco1.AddItem "9"
   cmb_Plazo2Bco2.AddItem ""
   cmb_Plazo2Bco2.AddItem "1"
   cmb_Plazo2Bco2.AddItem "2"
   cmb_Plazo2Bco2.AddItem "3"
   cmb_Plazo2Bco2.AddItem "4"
   cmb_Plazo2Bco2.AddItem "5"
   cmb_Plazo2Bco2.AddItem "6"
   cmb_Plazo2Bco2.AddItem "7"
   cmb_Plazo2Bco2.AddItem "8"
   cmb_Plazo2Bco2.AddItem "9"
   
   cmb_ValIng1Bco1.AddItem ""
   cmb_ValIng1Bco1.AddItem "10"
   cmb_ValIng1Bco1.AddItem "15"
   cmb_ValIng1Bco1.AddItem "20"
   cmb_ValIng1Bco1.AddItem "25"
   cmb_ValIng1Bco1.AddItem "30"
   cmb_ValIng1Bco1.AddItem "35"
   cmb_ValIng1Bco1.AddItem "40"
   cmb_ValIng1Bco1.AddItem "45"
   cmb_ValIng1Bco1.AddItem "50"
   cmb_ValIng1Bco1.AddItem "55"
   cmb_ValIng1Bco1.AddItem "60"
   cmb_ValIng1Bco1.AddItem "65"
   cmb_ValIng1Bco1.AddItem "70"
   cmb_ValIng1Bco1.AddItem "75"
   cmb_ValIng1Bco1.AddItem "80"
   cmb_ValIng1Bco2.AddItem ""
   cmb_ValIng1Bco2.AddItem "10"
   cmb_ValIng1Bco2.AddItem "15"
   cmb_ValIng1Bco2.AddItem "20"
   cmb_ValIng1Bco2.AddItem "25"
   cmb_ValIng1Bco2.AddItem "30"
   cmb_ValIng1Bco2.AddItem "35"
   cmb_ValIng1Bco2.AddItem "40"
   cmb_ValIng1Bco2.AddItem "45"
   cmb_ValIng1Bco2.AddItem "50"
   cmb_ValIng1Bco2.AddItem "55"
   cmb_ValIng1Bco2.AddItem "60"
   cmb_ValIng1Bco2.AddItem "65"
   cmb_ValIng1Bco2.AddItem "70"
   cmb_ValIng1Bco2.AddItem "75"
   cmb_ValIng1Bco2.AddItem "80"
   
   cmb_ValIng2Bco1.AddItem ""
   cmb_ValIng2Bco1.AddItem "10"
   cmb_ValIng2Bco1.AddItem "15"
   cmb_ValIng2Bco1.AddItem "20"
   cmb_ValIng2Bco1.AddItem "25"
   cmb_ValIng2Bco1.AddItem "30"
   cmb_ValIng2Bco1.AddItem "35"
   cmb_ValIng2Bco1.AddItem "40"
   cmb_ValIng2Bco1.AddItem "45"
   cmb_ValIng2Bco1.AddItem "50"
   cmb_ValIng2Bco1.AddItem "55"
   cmb_ValIng2Bco1.AddItem "60"
   cmb_ValIng2Bco1.AddItem "65"
   cmb_ValIng2Bco1.AddItem "70"
   cmb_ValIng2Bco1.AddItem "75"
   cmb_ValIng2Bco1.AddItem "80"
   cmb_ValIng2Bco2.AddItem ""
   cmb_ValIng2Bco2.AddItem "10"
   cmb_ValIng2Bco2.AddItem "15"
   cmb_ValIng2Bco2.AddItem "20"
   cmb_ValIng2Bco2.AddItem "25"
   cmb_ValIng2Bco2.AddItem "30"
   cmb_ValIng2Bco2.AddItem "35"
   cmb_ValIng2Bco2.AddItem "40"
   cmb_ValIng2Bco2.AddItem "45"
   cmb_ValIng2Bco2.AddItem "50"
   cmb_ValIng2Bco2.AddItem "55"
   cmb_ValIng2Bco2.AddItem "60"
   cmb_ValIng2Bco2.AddItem "65"
   cmb_ValIng2Bco2.AddItem "70"
   cmb_ValIng2Bco2.AddItem "75"
   cmb_ValIng2Bco2.AddItem "80"
End Sub

Private Sub fs_Limpia()
   txt_NomPry.Text = ""
   txt_DesPry.Text = ""
   cmb_Situac.ListIndex = -1
   cmb_ConAso.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_Interi.Text = ""
   txt_Interi.Enabled = False
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Cargo.Text = ""
   txt_Email.Text = ""
   txt_Telefono.Text = ""
   txt_Contacto.Text = ""
   
   Call fs_Limpia_Pro
   Call fs_Activa_Pro(True)
   Call fs_Limpia_Con
   Call fs_Activa_Con(True)
   Call fs_Limpia_Competencia
End Sub

Private Sub fs_Limpia_Pro()
   cmb_TipDoc_Pro.ListIndex = -1
   txt_NumDoc_Pro.Text = ""
   Call gs_LimpiaGrid(grd_Listad_Pro)
End Sub

Private Sub fs_Activa_Pro(ByVal p_Activa As Integer)
   cmb_TipDoc_Pro.Enabled = p_Activa
   txt_NumDoc_Pro.Enabled = p_Activa
   cmd_Buscar_Pro.Enabled = p_Activa
   grd_Listad_Pro.Enabled = Not p_Activa
   cmd_Editar_Pro.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia_Con()
   cmb_TipDoc_Con.ListIndex = -1
   txt_NumDoc_Con.Text = ""
   Call gs_LimpiaGrid(grd_Listad_Con)
End Sub

Private Sub fs_Limpia_Competencia()
   txt_ComeBco1.Text = ""
   txt_ModEvaBco1.Text = ""
   txt_Com1Bco1.Text = ""
   txt_Com2Bco1.Text = ""
   txt_ComeBco2.Text = ""
   txt_ModEvaBco2.Text = ""
   txt_Com1Bco2.Text = ""
   txt_Com2Bco2.Text = ""
End Sub

Private Sub fs_Activa_Con(ByVal p_Activa As Integer)
   cmb_TipDoc_Con.Enabled = p_Activa
   txt_NumDoc_Con.Enabled = p_Activa
   cmd_Buscar_Con.Enabled = p_Activa
   grd_Listad_Con.Enabled = Not p_Activa
   cmd_Editar_Con.Enabled = Not p_Activa
End Sub

Private Sub cmb_TipDoc_Pro_Click()
   Call gs_SetFocus(txt_NumDoc_Pro)
End Sub

Private Sub cmb_TipDoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Pro_Click
   End If
End Sub

Private Sub grd_Listad_Con_DblClick()
   If grd_Listad_Con.Rows = 0 Then
      Exit Sub
   End If
   Call cmd_Editar_Con_Click
End Sub

Private Sub grd_Listad_Pro_DblClick()
   If grd_Listad_Pro.Rows = 0 Then
      Exit Sub
   End If
   Call cmd_Editar_Pro_Click
End Sub

Private Sub grdLstAsignados_SelChange()
   If grdLstAsignados.Rows > 2 Then
      grdLstAsignados.RowSel = grdLstAsignados.Row
   End If
End Sub

Private Sub ipp_AreaMax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub ipp_AreaMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreaMax)
   End If
End Sub

Private Sub ipp_Avance_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Coloca)
   End If
End Sub

Private Sub ipp_Coloca_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Partic)
   End If
End Sub

Private Sub ipp_Coloca_LostFocus()
   If CDbl(ipp_TotUni.Text) > 0 Then
      ipp_Partic.Text = Round((ipp_Coloca.Text / ipp_TotUni.Text) * 100, 2)
   Else
      ipp_Partic.Text = 0
   End If
End Sub

Private Sub ipp_Dispon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Avance)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TotUni)
   End If
End Sub

Private Sub ipp_FecInfInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecInfLeg)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
   End If
End Sub

Private Sub ipp_FecInfLeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecRevApr)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecLim_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecInfInm)
   Else
      'KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_FecRevApr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_AproFile)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "/")
   End If
End Sub

Private Sub ipp_NumEtapa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Tasa)
   End If
End Sub

Private Sub ipp_Partic_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumEtapa)
   End If
End Sub

Private Sub ipp_PreMax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreaMin)
   End If
End Sub

Private Sub ipp_PreMin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PreMax)
   End If
End Sub

Private Sub cmb_EntFinBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TasaBco1)
   ElseIf KeyAscii = 8 Then
      cmb_EntFinBco1.ListIndex = -1
   End If
End Sub

Private Sub ipp_Tasa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipoGar)
   End If
End Sub

Private Sub ipp_TasaBco1_GotFocus()
   Call gs_SelecTodo(ipp_TasaBco1)
End Sub

Private Sub ipp_TasaBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipoGarBco1)
   End If
End Sub

Private Sub ipp_TasaBco2_GotFocus()
   Call gs_SelecTodo(ipp_TasaBco2)
End Sub

Private Sub ipp_TasaBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipoGarBco2)
   End If
End Sub

Private Sub ipp_TotDisp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Dispon)
   End If
End Sub

Private Sub ipp_TotUni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TotVen)
   End If
End Sub

Private Sub ipp_TotVen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      ipp_TotVen_LostFocus
   End If
End Sub

Private Sub ipp_TotVen_LostFocus()
   If CDbl(ipp_TotVen.Text) > CDbl(ipp_TotUni.Text) Then
      MsgBox "El valor ingresado no puede ser mayor al total de Unidades", vbExclamation, modgen_g_str_NomPlt
      ipp_TotDisp.Text = 0
      Exit Sub
   End If
   
   ipp_TotDisp.Text = ipp_TotUni.Text - ipp_TotVen.Text
   If CDbl(ipp_TotUni.Text) > 0 Then
      ipp_Dispon.Text = (ipp_TotDisp.Text / ipp_TotUni.Text) * 100
      ipp_Avance.Text = (ipp_TotVen.Text / ipp_TotUni.Text) * 100
   Else
      ipp_Dispon.Text = 0
      ipp_Avance.Text = 0
   End If
   Call gs_SetFocus(ipp_TotDisp)
End Sub

Private Sub txt_Cargo_GotFocus()
   Call gs_SelecTodo(txt_Cargo)
End Sub

Private Sub txt_Cargo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Email)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Com1Bco1_GotFocus()
   Call gs_SelecTodo(txt_Com1Bco1)
End Sub

Private Sub txt_Com1Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CtaIni2Bco1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Com1Bco2_GotFocus()
   Call gs_SelecTodo(txt_Com1Bco2)
End Sub

Private Sub txt_Com1Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CtaIni2Bco2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Com2Bco1_GotFocus()
   Call gs_SelecTodo(txt_Com2Bco1)
End Sub

Private Sub txt_Com2Bco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SSTab2.Tab = 1
      Call gs_SetFocus(cmb_EntFinBco2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Com2Bco2_GotFocus()
   Call gs_SelecTodo(txt_Com2Bco2)
End Sub

Private Sub txt_Com2Bco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_ComeBco1_GotFocus()
   Call gs_SelecTodo(txt_ComeBco1)
End Sub

Private Sub txt_ComeBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ModEvaBco1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_ComeBco2_GotFocus()
   Call gs_SelecTodo(txt_ComeBco2)
End Sub

Private Sub txt_ComeBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ModEvaBco2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Contacto_GotFocus()
   Call gs_SelecTodo(txt_Contacto)
End Sub

Private Sub txt_Contacto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Cargo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & " " & Chr(10))
   End If
End Sub

Private Sub txt_DesPry_GotFocus()
   Call gs_SelecTodo(txt_DesPry)
End Sub

Private Sub txt_DesPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Email_GotFocus()
   Call gs_SelecTodo(txt_Email)
End Sub

Private Sub txt_Email_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefono)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "_@." & Chr(10))
   End If
End Sub

Private Sub txt_ModEvaBco1_GotFocus()
   Call gs_SelecTodo(txt_ModEvaBco1)
End Sub

Private Sub txt_ModEvaBco1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CtaIni1Bco1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_ModEvaBco2_GotFocus()
   Call gs_SelecTodo(txt_ModEvaBco2)
End Sub

Private Sub txt_ModEvaBco2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CtaIni1Bco2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomPry_GotFocus()
   Call gs_SelecTodo(txt_NomPry)
End Sub

Private Sub txt_NomPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DesPry)
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

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Interi_GotFocus()
   Call gs_SelecTodo(txt_Interi)
End Sub

Private Sub txt_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumDoc_Pro_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Pro)
End Sub

Private Sub txt_NumDoc_Pro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar_Pro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub grd_Listad_Pro_SelChange()
   If grd_Listad_Pro.Rows > 2 Then
      grd_Listad_Pro.RowSel = grd_Listad_Pro.Row
   End If
End Sub

Private Sub cmb_TipDoc_Con_Click()
   Call gs_SetFocus(txt_NumDoc_Con)
End Sub

Private Sub cmb_TipDoc_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Con_Click
   End If
End Sub

Private Sub txt_NumDoc_Con_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Con)
End Sub

Private Sub txt_NumDoc_Con_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar_Con)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub grd_Listad_Con_SelChange()
   If grd_Listad_Con.Rows > 2 Then
      grd_Listad_Con.RowSel = grd_Listad_Con.Row
   End If
End Sub

Private Sub fs_Buscar_EmpPro()
   Call gs_LimpiaGrid(grd_Listad_Pro)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   Call fs_Activa_Pro(False)
   cmd_Editar_Pro.Enabled = True
   grd_Listad_Pro.Redraw = False
   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Documento de Identidad"
   
   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = moddat_gf_Consulta_ParDes("203", CStr(moddat_g_int_TipDoc)) & " - " & Trim(moddat_g_str_NumDoc & "")

   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Razón Social"
   
   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")

   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Nombre Comercial"

   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
         
   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Dirección"
         
   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                               " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_numVia) & _
                               IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Referencia"
   
   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         
   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Dpto. / Provin. / Dist."

   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
         
   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Teléfonos"
   
   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "") & IIf(Len(Trim(Trim(g_rst_Princi!DATGEN_TELEF2 & ""))) > 0, " / " & Trim(g_rst_Princi!DATGEN_TELEF2 & ""), "")
         
   grd_Listad_Pro.Rows = grd_Listad_Pro.Rows + 1
   grd_Listad_Pro.Row = grd_Listad_Pro.Rows - 1
   grd_Listad_Pro.Col = 0
   grd_Listad_Pro.Text = "Fax"
   
   grd_Listad_Pro.Col = 1
   grd_Listad_Pro.Text = Trim(g_rst_Princi!DATGEN_NUMFAX & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad_Pro.Redraw = True
   Call gs_UbiIniGrid(grd_Listad_Pro)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SSTab1.Tab = 1
      Call gs_SetFocus(txt_Contacto)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub fs_Buscar_EmpCon()
   Call gs_LimpiaGrid(grd_Listad_Con)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NumDoc & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   Call fs_Activa_Con(False)
   cmd_Editar_Con.Enabled = True
   grd_Listad_Con.Redraw = False
   
   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Documento de Identidad"
   
   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = moddat_gf_Consulta_ParDes("203", CStr(moddat_g_int_TipDoc)) & " - " & Trim(moddat_g_str_NumDoc & "")

   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Razón Social"
   
   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = Trim(g_rst_Princi!DATGEN_RAZSOC & "")

   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Nombre Comercial"

   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = Trim(g_rst_Princi!DATGEN_NOMCOM & "")
         
   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Dirección"
         
   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                               " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_numVia) & _
                               IIf(Len(Trim(g_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(g_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Referencia"
   
   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
         
   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Dpto. / Provin. / Dist."

   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
         
   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Teléfonos"
   
   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "") & IIf(Len(Trim(Trim(g_rst_Princi!DATGEN_TELEF2 & ""))) > 0, " / " & Trim(g_rst_Princi!DATGEN_TELEF2 & ""), "")
         
   grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
   grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
   grd_Listad_Con.Col = 0
   grd_Listad_Con.Text = "Fax"
   
   grd_Listad_Con.Col = 1
   grd_Listad_Con.Text = Trim(g_rst_Princi!DATGEN_NUMFAX & "")
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad_Con.Redraw = True
   Call gs_UbiIniGrid(grd_Listad_Con)
End Sub

Private Sub txt_Telefono_GotFocus()
   Call gs_SelecTodo(txt_Telefono)
End Sub

Private Sub txt_Telefono_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PreMin)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-_. ,;:()/º" & Chr(10))
   End If
End Sub

Private Function fs_ObtieneOperaciones_Proyecto(ByVal p_CodPry As String) As Integer
Dim r_rst_RstPry        As ADODB.Recordset

   fs_ObtieneOperaciones_Proyecto = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS CONTADOR "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC IN (2,9) "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_PRYINM = '" & p_CodPry & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_RstPry, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_RstPry.BOF And r_rst_RstPry.EOF) Then
      r_rst_RstPry.MoveFirst
      fs_ObtieneOperaciones_Proyecto = r_rst_RstPry!CONTADOR
   End If
   
   r_rst_RstPry.Close
   Set r_rst_RstPry = Nothing
End Function

