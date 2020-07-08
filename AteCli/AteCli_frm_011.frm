VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_SegSol_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   1425
   ClientTop       =   1665
   ClientWidth     =   13635
   Icon            =   "AteCli_frm_011.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   13635
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7785
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13635
      _Version        =   65536
      _ExtentX        =   24051
      _ExtentY        =   13732
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   735
         Left            =   30
         TabIndex        =   41
         Top             =   6990
         Width           =   13545
         _Version        =   65536
         _ExtentX        =   23892
         _ExtentY        =   1296
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
         Begin VB.CommandButton cmd_SegSol 
            Height          =   675
            Left            =   12810
            Picture         =   "AteCli_frm_011.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Seguimiento de Solicitud"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   825
         Left            =   30
         TabIndex        =   19
         Top             =   750
         Width           =   13545
         _Version        =   65536
         _ExtentX        =   23892
         _ExtentY        =   1455
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
         Begin VB.ComboBox cmb_TipFec 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   4035
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12840
            Picture         =   "AteCli_frm_011.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   12150
            Picture         =   "AteCli_frm_011.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   675
            Left            =   11460
            Picture         =   "AteCli_frm_011.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Buscar Cliente"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10770
            Picture         =   "AteCli_frm_011.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   7380
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   3015
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4035
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   8910
            TabIndex        =   4
            Top             =   390
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
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
            Text            =   "28/09/2004"
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
            Left            =   7380
            TabIndex        =   3
            Top             =   390
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
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
            Text            =   "28/09/2004"
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
         Begin VB.Label Label3 
            Caption         =   "Rango de Fechas:"
            Height          =   315
            Left            =   5910
            TabIndex        =   23
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label Label2 
            Caption         =   "Situación Solicitud:"
            Height          =   315
            Left            =   5910
            TabIndex        =   22
            Top             =   60
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo de Fecha:"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   13545
         _Version        =   65536
         _ExtentX        =   23892
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
            TabIndex        =   14
            Top             =   60
            Width           =   6465
            _Version        =   65536
            _ExtentX        =   11404
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Solicitudes"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_011.frx":1076
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5325
         Left            =   30
         TabIndex        =   16
         Top             =   1620
         Width           =   13545
         _Version        =   65536
         _ExtentX        =   23892
         _ExtentY        =   9393
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Eva 
            Height          =   4965
            Left            =   30
            TabIndex        =   9
            Top             =   330
            Width           =   13425
            _ExtentX        =   23680
            _ExtentY        =   8758
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1590
            TabIndex        =   17
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   3060
            TabIndex        =   24
            Top             =   60
            Width           =   4155
            _Version        =   65536
            _ExtentX        =   7329
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   8580
            TabIndex        =   25
            Top             =   60
            Width           =   4620
            _Version        =   65536
            _ExtentX        =   8149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Instancia Actual"
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
            Left            =   7200
            TabIndex        =   26
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
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
      Begin Threed.SSPanel pnl_SolDes 
         Height          =   5325
         Left            =   30
         TabIndex        =   27
         Top             =   1620
         Width           =   13545
         _Version        =   65536
         _ExtentX        =   23892
         _ExtentY        =   9393
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Des 
            Height          =   4965
            Left            =   30
            TabIndex        =   10
            Top             =   330
            Width           =   13425
            _ExtentX        =   23680
            _ExtentY        =   8758
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   1590
            TabIndex        =   28
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   60
            TabIndex        =   29
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   285
            Left            =   3060
            TabIndex        =   30
            Top             =   60
            Width           =   4155
            _Version        =   65536
            _ExtentX        =   7329
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel SSPanel21 
            Height          =   285
            Left            =   8580
            TabIndex        =   31
            Top             =   60
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   285
            Left            =   7200
            TabIndex        =   32
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
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
         Begin Threed.SSPanel SSPanel23 
            Height          =   285
            Left            =   11580
            TabIndex        =   33
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Desembolso"
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
      Begin Threed.SSPanel pnl_SolRec 
         Height          =   5325
         Left            =   30
         TabIndex        =   34
         Top             =   1620
         Width           =   13545
         _Version        =   65536
         _ExtentX        =   23892
         _ExtentY        =   9393
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Rec 
            Height          =   4965
            Left            =   30
            TabIndex        =   11
            Top             =   330
            Width           =   13425
            _ExtentX        =   23680
            _ExtentY        =   8758
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   1590
            TabIndex        =   35
            Top             =   60
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
            Left            =   3060
            TabIndex        =   37
            Top             =   60
            Width           =   4155
            _Version        =   65536
            _ExtentX        =   7329
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   8580
            TabIndex        =   38
            Top             =   60
            Width           =   3165
            _Version        =   65536
            _ExtentX        =   5583
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Rechazo"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   7200
            TabIndex        =   39
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Solicitud"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   11730
            TabIndex        =   40
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Rechazo"
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
   End
End
Attribute VB_Name = "frm_SegSol_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmb_Situac)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_Situac_Click()
   cmb_TipFec.Clear
   
   If cmb_Situac.ListIndex > -1 Then
      Select Case cmb_Situac.ItemData(cmb_Situac.ListIndex)
         Case 1
            cmb_TipFec.AddItem "POR FECHA DE INGRESO"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 1
            
            pnl_SolEva.Visible = True
            pnl_SolRec.Visible = False
            pnl_SolDes.Visible = False
            
         Case 2
            cmb_TipFec.AddItem "POR FECHA DE INGRESO"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 1
         
            cmb_TipFec.AddItem "POR FECHA DE DESEMBOLSO"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 2
         
            pnl_SolEva.Visible = False
            pnl_SolRec.Visible = False
            pnl_SolDes.Visible = True
            
         Case 3
            cmb_TipFec.AddItem "POR FECHA DE INGRESO"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 1
         
            cmb_TipFec.AddItem "POR FECHA DE RECHAZO"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 2
         
            pnl_SolEva.Visible = False
            pnl_SolRec.Visible = True
            pnl_SolDes.Visible = False
            
         Case 9
            cmb_TipFec.AddItem "POR FECHA DE INGRESO"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 1
            
            cmb_TipFec.AddItem "POR FECHA DE ANULACION"
            cmb_TipFec.ItemData(cmb_TipFec.NewIndex) = 1
            
            pnl_SolEva.Visible = True
            pnl_SolRec.Visible = False
            pnl_SolDes.Visible = False
            
      End Select
      
      cmb_TipFec.ListIndex = -1
      
      Call moddat_gs_FecSis
      ipp_FecIni.Text = Format(CDate(moddat_g_str_FecSis) - CDate(60), "dd/mm/yyyy")
      ipp_FecFin.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   End If
   
   Call gs_SetFocus(cmb_TipFec)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_TipFec_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipFec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipFec_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación de la Solicitud.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If

   If cmb_TipFec.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Fecha.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipFec)
      Exit Sub
   End If

   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   moddat_g_str_CodPrd = Right(l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 3)
   moddat_g_str_NomPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Nombre

   Select Case cmb_Situac.ItemData(cmb_Situac.ListIndex)
      Case 1:  Call fs_Buscar_SolTra
      Case 2
         Select Case cmb_TipFec.ItemData(cmb_TipFec.ListIndex)
            Case 1:  Call fs_Buscar_SolDes(1)
            Case 2:  Call fs_Buscar_SolDes(2)
         End Select
      Case 3
         Select Case cmb_TipFec.ItemData(cmb_TipFec.ListIndex)
            Case 1:  Call fs_Buscar_SolRec(1)
            Case 2:  Call fs_Buscar_SolRec(2)
         End Select
      Case 9
         Select Case cmb_TipFec.ItemData(cmb_TipFec.ListIndex)
            Case 1:  Call fs_Buscar_SolAnu(1)
            Case 2:  Call fs_Buscar_SolAnu(2)
         End Select
   
   End Select
End Sub

Private Sub cmd_BusCli_Click()
   frm_SegSol_02.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Produc.ListIndex = -1
   cmb_Situac.ListIndex = -1
   
   Call cmb_Situac_Click
   
   pnl_SolEva.Visible = True
   pnl_SolRec.Visible = False
   pnl_SolDes.Visible = False
   
   Call gs_LimpiaGrid(grd_Listad_Eva)
   Call gs_LimpiaGrid(grd_Listad_Rec)
   Call gs_LimpiaGrid(grd_Listad_Des)
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SegSol_Click()
   Select Case cmb_Situac.ItemData(cmb_Situac.ListIndex)
      Case 1, 9
         grd_Listad_Eva.Col = 0
         moddat_g_str_NumSol = Mid(grd_Listad_Eva.Text, 1, 3) & Mid(grd_Listad_Eva.Text, 5, 3) & Mid(grd_Listad_Eva.Text, 9, 2) & Mid(grd_Listad_Eva.Text, 12, 4)
         
         grd_Listad_Eva.Col = 1
         moddat_g_int_TipDoc = CInt(Left(grd_Listad_Eva.Text, 1))
         moddat_g_str_NumDoc = Mid(grd_Listad_Eva.Text, 3)
         
         grd_Listad_Eva.Col = 2
         moddat_g_str_NomCli = grd_Listad_Eva.Text
         
         Call gs_RefrescaGrid(grd_Listad_Eva)
      Case 2
         grd_Listad_Des.Col = 0
         moddat_g_str_NumSol = Mid(grd_Listad_Des.Text, 1, 3) & Mid(grd_Listad_Des.Text, 5, 3) & Mid(grd_Listad_Des.Text, 9, 2) & Mid(grd_Listad_Des.Text, 12, 4)
         
         grd_Listad_Des.Col = 1
         moddat_g_int_TipDoc = CInt(Left(grd_Listad_Des.Text, 1))
         moddat_g_str_NumDoc = Mid(grd_Listad_Des.Text, 3)
         
         grd_Listad_Des.Col = 2
         moddat_g_str_NomCli = grd_Listad_Des.Text
         
         Call gs_RefrescaGrid(grd_Listad_Des)
      
      Case 3
         grd_Listad_Rec.Col = 0
         moddat_g_str_NumSol = Mid(grd_Listad_Rec.Text, 1, 3) & Mid(grd_Listad_Rec.Text, 5, 3) & Mid(grd_Listad_Rec.Text, 9, 2) & Mid(grd_Listad_Rec.Text, 12, 4)
         
         grd_Listad_Rec.Col = 1
         moddat_g_int_TipDoc = CInt(Left(grd_Listad_Rec.Text, 1))
         moddat_g_str_NumDoc = Mid(grd_Listad_Rec.Text, 3)
         
         grd_Listad_Rec.Col = 2
         moddat_g_str_NomCli = grd_Listad_Rec.Text
         
         Call gs_RefrescaGrid(grd_Listad_Rec)
   End Select
   
   frm_SegSol_03.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   Call fs_Carga_SitSol
   
   grd_Listad_Eva.ColWidth(0) = 1525
   grd_Listad_Eva.ColWidth(1) = 1465
   grd_Listad_Eva.ColWidth(2) = 4145
   grd_Listad_Eva.ColWidth(3) = 1385
   grd_Listad_Eva.ColWidth(4) = 4600
   
   grd_Listad_Eva.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_Eva.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_Eva.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad_Eva.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Eva.ColAlignment(4) = flexAlignLeftCenter

   grd_Listad_Rec.ColWidth(0) = 1525
   grd_Listad_Rec.ColWidth(1) = 1465
   grd_Listad_Rec.ColWidth(2) = 4145
   grd_Listad_Rec.ColWidth(3) = 1385
   grd_Listad_Rec.ColWidth(4) = 3155
   grd_Listad_Rec.ColWidth(5) = 1455

   grd_Listad_Rec.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_Rec.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_Rec.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad_Rec.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Rec.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Rec.ColAlignment(5) = flexAlignCenterCenter

   grd_Listad_Des.ColWidth(0) = 1525
   grd_Listad_Des.ColWidth(1) = 1465
   grd_Listad_Des.ColWidth(2) = 4145
   grd_Listad_Des.ColWidth(3) = 1385
   grd_Listad_Des.ColWidth(4) = 3000
   grd_Listad_Des.ColWidth(5) = 1600
   
   grd_Listad_Des.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_Des.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_Des.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad_Des.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Des.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Des.ColAlignment(5) = flexAlignCenterCenter
End Sub

Private Sub fs_Carga_SitSol()
   cmb_Situac.Clear
   
   cmb_Situac.AddItem "SOLICITUDES EN TRAMITE"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 1
   
   cmb_Situac.AddItem "SOLICITUDES DESEMBOLSADAS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 2
   
   cmb_Situac.AddItem "SOLICITUDES RECHAZADAS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 3
   
   cmb_Situac.AddItem "SOLICITUDES ANULADAS"
   cmb_Situac.ItemData(cmb_Situac.NewIndex) = 9
   
   cmb_Situac.ListIndex = -1
End Sub

Private Sub grd_Listad_Des_DblClick()
   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_Eva_DblClick()
   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_Rec_DblClick()
   Call cmd_SegSol_Click
End Sub

Private Sub grd_Listad_Rec_SelChange()
   If grd_Listad_Rec.Rows > 2 Then
      grd_Listad_Rec.RowSel = grd_Listad_Rec.Row
   End If
End Sub

Private Sub grd_Listad_Eva_SelChange()
   If grd_Listad_Eva.Rows > 2 Then
      grd_Listad_Eva.RowSel = grd_Listad_Eva.Row
   End If
End Sub

Private Sub grd_Listad_Des_SelChange()
   If grd_Listad_Des.Rows > 2 Then
      grd_Listad_Des.RowSel = grd_Listad_Des.Row
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
   cmb_Situac.Enabled = p_Activa
   cmb_TipFec.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   If pnl_SolEva.Visible Then
      grd_Listad_Eva.Enabled = Not p_Activa
   End If

   If pnl_SolRec.Visible Then
      grd_Listad_Rec.Enabled = Not p_Activa
   End If
   
   If pnl_SolDes.Visible Then
      grd_Listad_Des.Enabled = Not p_Activa
   End If
   
   cmd_SegSol.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar_SolTra()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad_Eva)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad_Eva.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad_Eva.Rows = grd_Listad_Eva.Rows + 1
      grd_Listad_Eva.Row = grd_Listad_Eva.Rows - 1
      
      'Número de Solicitud
      grd_Listad_Eva.Col = 0
      grd_Listad_Eva.Text = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      
      'ID Cliente
      grd_Listad_Eva.Col = 1
      grd_Listad_Eva.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      'Apellidos y Nombres
      grd_Listad_Eva.Col = 2
      grd_Listad_Eva.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))

      'Fecha de Solicitud
      grd_Listad_Eva.Col = 3
      grd_Listad_Eva.Text = Right(CStr(g_rst_Princi!SOLMAE_FECSOL), 2) & "/" & Mid(CStr(g_rst_Princi!SOLMAE_FECSOL), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLMAE_FECSOL), 4)
      
      'Instancia Actual
      grd_Listad_Eva.Col = 4
      grd_Listad_Eva.Text = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad_Eva.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad_Eva)

   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad_Eva)
End Sub

Private Sub fs_Buscar_SolRec(ByVal p_TipFec As Integer)
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 3 AND "
   
   If p_TipFec = 1 Then
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "SOLMAE_FECREC >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_FECREC <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad_Rec)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad_Rec.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad_Rec.Rows = grd_Listad_Rec.Rows + 1
      grd_Listad_Rec.Row = grd_Listad_Rec.Rows - 1
      
      'Número de Solicitud
      grd_Listad_Rec.Col = 0
      grd_Listad_Rec.Text = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      
      'ID Cliente
      grd_Listad_Rec.Col = 1
      grd_Listad_Rec.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      'Apellidos y Nombres
      grd_Listad_Rec.Col = 2
      grd_Listad_Rec.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      'Fecha de Solicitud
      grd_Listad_Rec.Col = 3
      grd_Listad_Rec.Text = Right(CStr(g_rst_Princi!SOLMAE_FECSOL), 2) & "/" & Mid(CStr(g_rst_Princi!SOLMAE_FECSOL), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLMAE_FECSOL), 4)
      
      'Tipo de Rechazo
      grd_Listad_Rec.Col = 4
      grd_Listad_Rec.Text = moddat_gf_Consulta_ParDes("021", CStr(g_rst_Princi!SOLMAE_TIPREC))
      
      'Fecha de Rechazo
      grd_Listad_Rec.Col = 5
      grd_Listad_Rec.Text = Right(Format(g_rst_Princi!SOLMAE_FECREC, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECREC, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECREC, "00000000"), 4)
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad_Rec.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad_Rec)

   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad_Rec)
End Sub

Private Sub fs_Buscar_SolDes(ByVal p_TipFec As Integer)
   g_str_Parame = "SELECT * FROM CRE_SOLMAE A, CRE_HIPMAE B WHERE "
   g_str_Parame = g_str_Parame & "A.SOLMAE_NUMERO = B.HIPMAE_NUMSOL AND "
   g_str_Parame = g_str_Parame & "A.SOLMAE_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "A.SOLMAE_SITUAC = 2 AND "
   
   If p_TipFec = 1 Then
      g_str_Parame = g_str_Parame & "A.SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "A.SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "B.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "B.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY A.SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad_Des)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad_Des.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad_Des.Rows = grd_Listad_Des.Rows + 1
      grd_Listad_Des.Row = grd_Listad_Des.Rows - 1
      
      'Número de Solicitud
      grd_Listad_Des.Col = 0
      grd_Listad_Des.Text = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      
      'ID Cliente
      grd_Listad_Des.Col = 1
      grd_Listad_Des.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      'Apellidos y Nombres
      grd_Listad_Des.Col = 2
      grd_Listad_Des.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
      
      'Fecha de Solicitud
      grd_Listad_Des.Col = 3
      grd_Listad_Des.Text = Right(CStr(g_rst_Princi!SOLMAE_FECSOL), 2) & "/" & Mid(CStr(g_rst_Princi!SOLMAE_FECSOL), 5, 2) & "/" & Left(CStr(g_rst_Princi!SOLMAE_FECSOL), 4)

      'Número de Operación
      grd_Listad_Des.Col = 4
      grd_Listad_Des.Text = Left(g_rst_Princi!HIPMAE_NUMOPE, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMOPE, 4, 2) & "-" & Right(g_rst_Princi!HIPMAE_NUMOPE, 5)
      
      'Fecha de Desembolso
      grd_Listad_Des.Col = 5
      grd_Listad_Des.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad_Des.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad_Des)

   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad_Des)
End Sub

Private Sub fs_Buscar_SolAnu(ByVal p_TipFec As Integer)
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 9 AND "
   
   If p_TipFec = 1 Then
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_FECSOL <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   Else
      g_str_Parame = g_str_Parame & "SEGFECACT  >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
      g_str_Parame = g_str_Parame & "SEGFECACT <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   Call gs_LimpiaGrid(grd_Listad_Eva)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado Solicitudes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   grd_Listad_Eva.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad_Eva.Rows = grd_Listad_Eva.Rows + 1
      grd_Listad_Eva.Row = grd_Listad_Eva.Rows - 1
      
      'Número de Solicitud
      grd_Listad_Eva.Col = 0
      grd_Listad_Eva.Text = Mid(g_rst_Princi!SOLMAE_NUMERO, 1, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 9, 4)
      
      'ID Cliente
      grd_Listad_Eva.Col = 1
      grd_Listad_Eva.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
      
      'Apellidos y Nombres
      grd_Listad_Eva.Col = 2
      grd_Listad_Eva.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))

      'Fecha de Solicitud
      grd_Listad_Eva.Col = 3
      grd_Listad_Eva.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
      
      'Instancia Actual
      grd_Listad_Eva.Col = 4
      grd_Listad_Eva.Text = moddat_gf_Consulta_ParDes("002", Trim(g_rst_Princi!SOLMAE_CODINS))
      
      g_rst_Princi.MoveNext
   Loop
   grd_Listad_Eva.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad_Eva)

   Screen.MousePointer = 0
   
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad_Eva)
End Sub

