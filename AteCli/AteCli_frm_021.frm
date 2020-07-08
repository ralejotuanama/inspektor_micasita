VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_GasAdm_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8100
   ClientLeft      =   450
   ClientTop       =   1020
   ClientWidth     =   11565
   Icon            =   "AteCli_frm_021.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8085
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11565
      _Version        =   65536
      _ExtentX        =   20399
      _ExtentY        =   14261
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   765
         Left            =   30
         TabIndex        =   46
         Top             =   7260
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10020
            Picture         =   "AteCli_frm_021.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   10710
            Picture         =   "AteCli_frm_021.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   36
         Top             =   6450
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
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
         Begin VB.ComboBox cmb_GasAdm 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   3855
         End
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1590
            TabIndex        =   5
            Top             =   390
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
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
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_MonGas 
            Height          =   315
            Left            =   2610
            TabIndex        =   45
            Top             =   390
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin VB.Label Label5 
            Caption         =   "Concepto de Gasto:"
            Height          =   285
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   60
            TabIndex        =   37
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   3345
         Left            =   30
         TabIndex        =   32
         Top             =   3060
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   5900
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
         Begin VB.CommandButton cmd_EdiIte 
            Height          =   675
            Left            =   9990
            Picture         =   "AteCli_frm_021.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Edición de Datos"
            Top             =   2610
            Width           =   675
         End
         Begin VB.CommandButton cmd_NueIte 
            Height          =   675
            Left            =   9300
            Picture         =   "AteCli_frm_021.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   2610
            Width           =   675
         End
         Begin VB.CommandButton cmd_BorIte 
            Height          =   675
            Left            =   10710
            Picture         =   "AteCli_frm_021.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   2610
            Width           =   675
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2265
            Left            =   30
            TabIndex        =   0
            Top             =   330
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   3995
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto"
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   6210
            TabIndex        =   34
            Top             =   60
            Width           =   2625
            _Version        =   65536
            _ExtentX        =   4630
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Moneda"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   8820
            TabIndex        =   35
            Top             =   60
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   30
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   17
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Generación Gastos Administrativos"
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
            Left            =   3600
            TabIndex        =   18
            Top             =   150
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
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
            Picture         =   "AteCli_frm_021.frx":1076
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   19
         Top             =   1590
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
            TabIndex        =   20
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
            Left            =   7950
            TabIndex        =   21
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
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   1620
            TabIndex        =   22
            Top             =   1050
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   23
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   24
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO - MIVIVIENDA"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   7950
            TabIndex        =   25
            Top             =   60
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   6600
            TabIndex        =   31
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   6600
            TabIndex        =   29
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   28
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   39
         Top             =   750
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   9300
            Picture         =   "AteCli_frm_021.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   10020
            Picture         =   "AteCli_frm_021.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10740
            Picture         =   "AteCli_frm_021.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   60
            Width           =   2775
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   42
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   41
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   40
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_GasAdm_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_GasAdm()   As moddat_tpo_Genera
Dim l_dbl_Dol_TipCam As Double
Dim l_dbl_Pre_TipCam As Double

Private Sub cmb_GasAdm_Click()
   Call gs_SetFocus(ipp_Import)
   
   If cmb_GasAdm.ListIndex > -1 Then
      ipp_Import.Value = l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Cantid
      pnl_MonGas.Caption = moddat_gf_Consulta_ParDes("204", Right(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 1))
   End If
End Sub

Private Sub cmb_GasAdm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_GasAdm_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_BorIte_Click()
   Dim r_str_CodGas  As String
   
   grd_Listad.Col = 3
   r_str_CodGas = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro de eliminar el Gasto Administrativo asignado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = "USP_BORRAR_TRA_GASADM ('" & moddat_g_str_NumSol & "', " & Left(r_str_CodGas, 2) & "), 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If

   Call fs_Buscar_GasAdm
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumSol.Text)) < 12 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      moddat_g_str_NumSol = msk_NumSol.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   Else
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_DatGen

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_Moneda.Caption = moddat_g_str_Moneda
   pnl_FecIng.Caption = moddat_g_str_FecIng

   If moddat_g_int_InsAct < modatecli_g_con_EvaTas Then
      MsgBox "El Cliente debe aprobar la instancia ACEPTACION CREDITICIA INICIAL para poder asignar Gastos Administrativos. ", vbExclamation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
     
      Exit Sub
   End If

   'Verificando que los Gastos Administrativos no hayan sido pagados
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "Los Gastos Administrativos ya han sido cancelados.", vbExclamation, modgen_g_con_PltPar
     
     Call cmd_Limpia_Click
     Exit Sub
   End If

   'Cargando Gastos Administrativos
   Call moddat_gs_Carga_ParPrd_Combo(cmb_GasAdm, l_arr_GasAdm(), moddat_g_str_CodPrd, "201")

   Call fs_ActivaItem(False)
   Call fs_Activa(False)
   
   'Buscar Gastos Administrativos ya registrados
   Call fs_Buscar_GasAdm
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_EdiIte_Click()
   Dim r_str_CodGas  As Integer
   Dim r_int_Situac  As String
   
   grd_Listad.Col = 3
   r_str_CodGas = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   
   moddat_g_int_FlgGrb = 2
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_CODGAS = " & Left(r_str_CodGas, 2) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   g_rst_Princi.MoveFirst
   
   cmb_GasAdm.ListIndex = gf_Busca_Arregl(l_arr_GasAdm, Format(r_str_CodGas, "000")) - 1
   
   ipp_Import.Value = g_rst_Princi!GASADM_IMPORT
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_ActivaItem(True)
   cmb_GasAdm.Enabled = False
   
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_GasAdm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Gasto Administrativo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_GasAdm)
      Exit Sub
   End If

   If ipp_Import.Value = 0 Then
      MsgBox "Debe ingresar el Importe del Gasto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If

   'Validar que el Gasto no este ingresado si es Agregar
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
      g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "GASADM_CODGAS = " & Left(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 2)
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         MsgBox "El Gasto Administrativo ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_GasAdm)
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_GASADM_ASIGNA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Left(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 2) & ", "
      g_str_Parame = g_str_Parame & Right(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 1) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Import.Text)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   Call fs_Buscar_GasAdm
   
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_LimpiaItem
      Call fs_ActivaItem(True)
      Call gs_SetFocus(cmb_GasAdm)
   Else
      Call cmd_Cancel_Click
   End If
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_NueIte_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_ActivaItem(True)
   Call gs_SetFocus(cmb_GasAdm)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumSol.Enabled = False
         
         msk_NumSol.Mask = ""
         msk_NumSol.Text = ""
         msk_NumSol.Mask = "###-###-##-####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumSol.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumSol)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumSol.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumSol.Mask = ""
      msk_NumSol.Text = ""
      msk_NumSol.Mask = "###-###-##-####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub fs_Inicia()
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   
   grd_Listad.ColWidth(0) = 6150
   grd_Listad.ColWidth(1) = 2610
   grd_Listad.ColWidth(2) = 2280
   grd_Listad.ColWidth(3) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_EdiIte_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub msk_NumSol_GotFocus()
   Call gs_SelecTodo(msk_NumSol)
End Sub

Private Sub msk_NumSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Limpia()
   Call fs_ActivaItem(False)
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumSol.Enabled = False

   msk_NumSol.Mask = ""
   msk_NumSol.Text = ""
   msk_NumSol.Mask = "###-###-##-####"
   
   txt_NumDoc.Text = ""
   
   pnl_Client.Caption = ""
   pnl_NumSol.Caption = ""
   pnl_Produc.Caption = ""
   pnl_Modali.Caption = ""
   pnl_EjeVta.Caption = ""
   pnl_Moneda.Caption = ""
   pnl_FecIng.Caption = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   cmb_GasAdm.ListIndex = -1
   ipp_Import.Value = 0
   
   pnl_MonGas.Caption = ""
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_NueIte.Enabled = Not p_Habilita
   cmd_BorIte.Enabled = Not p_Habilita
   cmd_EdiIte.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   cmb_GasAdm.Enabled = p_Habilita
   ipp_Import.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_NueIte.Enabled = Not p_Habilita
   cmd_BorIte.Enabled = Not p_Habilita
   cmd_EdiIte.Enabled = Not p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   moddat_g_str_NomCli = ""
   moddat_g_str_NomPrd = ""
   moddat_g_str_DesMod = ""
   moddat_g_str_InsAct = ""
   moddat_g_str_EjeVta = ""
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)

   'Ejecutivo de Ventas
   moddat_g_str_CodEje = Trim(g_rst_Princi!SOLMAE_EJEVTA)
   moddat_g_str_EjeVta = moddat_gf_Buscar_NomEje(moddat_g_str_CodEje)

   'Instancia Actual
   moddat_g_int_InsAct = Trim(g_rst_Princi!SOLMAE_CODINS)
   moddat_g_str_InsAct = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SOLMAE_CODINS))
   
   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
End Sub
 
Private Sub fs_Buscar_GasAdm()
   cmd_NueIte.Enabled = True
   cmd_EdiIte.Enabled = False
   cmd_BorIte.Enabled = False
   
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_con_PltPar
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'Buscando Descripción de Gastos Administrativos
      grd_Listad.Col = 0
      If moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, "201", Format(g_rst_Princi!GASADM_CODGAS, "00") & CStr(g_rst_Princi!GASADM_TIPMON)) Then
         grd_Listad.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
      End If
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!GASADM_TIPMON))
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_EdiIte.Enabled = True
      cmd_BorIte.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

