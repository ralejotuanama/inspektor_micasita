VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_Tasaci_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   1380
   ClientTop       =   735
   ClientWidth     =   12855
   Icon            =   "AteCli_frm_016.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10005
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   17648
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
      Begin Threed.SSPanel SSPanel14 
         Height          =   765
         Left            =   30
         TabIndex        =   89
         Top             =   9180
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12030
            Picture         =   "AteCli_frm_016.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11340
            Picture         =   "AteCli_frm_016.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   30
         TabIndex        =   62
         Top             =   3900
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
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_016.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   11310
            Picture         =   "AteCli_frm_016.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_RegInf 
            Height          =   675
            Left            =   1440
            Picture         =   "AteCli_frm_016.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Registrar Evaluación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   750
            Picture         =   "AteCli_frm_016.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir Orden de Tasación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_NueEva 
            Height          =   675
            Left            =   60
            Picture         =   "AteCli_frm_016.frx":15F0
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Generar Orden de Tasación"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   4425
         Left            =   60
         TabIndex        =   52
         Top             =   4710
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   7805
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
         Begin VB.TextBox txt_NomPer 
            Height          =   315
            Left            =   1620
            MaxLength       =   30
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_Observ 
            Height          =   675
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Text            =   "AteCli_frm_016.frx":18FA
            Top             =   3720
            Width           =   11025
         End
         Begin VB.TextBox txt_NumInf 
            Height          =   315
            Left            =   1620
            MaxLength       =   30
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_EmpPer 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   60
            Width           =   3345
         End
         Begin EditLib.fpDateTime ipp_FecEva 
            Height          =   315
            Left            =   1620
            TabIndex        =   14
            Top             =   720
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
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
         Begin EditLib.fpDoubleSingle ipp_ValCom 
            Height          =   315
            Left            =   1620
            TabIndex        =   16
            Top             =   1830
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ValRea 
            Height          =   315
            Left            =   3300
            TabIndex        =   17
            Top             =   1830
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_AreTer 
            Height          =   315
            Left            =   4980
            TabIndex        =   18
            Top             =   1830
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_AreCon 
            Height          =   315
            Left            =   6660
            TabIndex        =   19
            Top             =   1830
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin Threed.SSPanel pnl_FecEmi 
            Height          =   315
            Left            =   8820
            TabIndex        =   57
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "01/10/2004"
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
         Begin EditLib.fpDoubleSingle ipp_VCoEs1 
            Height          =   315
            Left            =   1620
            TabIndex        =   20
            Top             =   2160
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_VReEs1 
            Height          =   315
            Left            =   3300
            TabIndex        =   21
            Top             =   2160
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ATeEs1 
            Height          =   315
            Left            =   4980
            TabIndex        =   22
            Top             =   2160
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ACoEs1 
            Height          =   315
            Left            =   6660
            TabIndex        =   23
            Top             =   2160
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_VCoEs2 
            Height          =   315
            Left            =   1620
            TabIndex        =   24
            Top             =   2490
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_VReEs2 
            Height          =   315
            Left            =   3300
            TabIndex        =   25
            Top             =   2490
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ATeEs2 
            Height          =   315
            Left            =   4980
            TabIndex        =   26
            Top             =   2490
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ACoEs2 
            Height          =   315
            Left            =   6660
            TabIndex        =   27
            Top             =   2490
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_VCoDep 
            Height          =   315
            Left            =   1620
            TabIndex        =   28
            Top             =   2820
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_VReDep 
            Height          =   315
            Left            =   3300
            TabIndex        =   29
            Top             =   2820
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ATeDep 
            Height          =   315
            Left            =   4980
            TabIndex        =   30
            Top             =   2820
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin EditLib.fpDoubleSingle ipp_ACoDep 
            Height          =   315
            Left            =   6660
            TabIndex        =   31
            Top             =   2820
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1620
            TabIndex        =   78
            Top             =   1530
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor Comerc. US$"
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
            Left            =   3300
            TabIndex        =   79
            Top             =   1530
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Valor Fabricac. US$"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4980
            TabIndex        =   80
            Top             =   1530
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Area Terreno m2"
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
            Left            =   6660
            TabIndex        =   81
            Top             =   1530
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Area Constr. m2"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   90
            Left            =   30
            TabIndex        =   82
            Top             =   1410
            Width           =   12645
            _Version        =   65536
            _ExtentX        =   22304
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   90
            Left            =   30
            TabIndex        =   83
            Top             =   3600
            Width           =   12645
            _Version        =   65536
            _ExtentX        =   22304
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnl_TotVCo 
            Height          =   315
            Left            =   1620
            TabIndex        =   85
            Top             =   3240
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotVRe 
            Height          =   315
            Left            =   3300
            TabIndex        =   86
            Top             =   3240
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotACo 
            Height          =   315
            Left            =   6660
            TabIndex        =   87
            Top             =   3240
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotATe 
            Height          =   315
            Left            =   4980
            TabIndex        =   88
            Top             =   3240
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
            Alignment       =   4
         End
         Begin VB.Label Label8 
            Caption         =   "Totales:"
            Height          =   285
            Left            =   60
            TabIndex        =   84
            Top             =   3240
            Width           =   1485
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   1590
            X2              =   8310
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Label Label23 
            Caption         =   "Depósito:"
            Height          =   285
            Left            =   60
            TabIndex        =   77
            Top             =   2820
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Estacionam. 2:"
            Height          =   285
            Left            =   60
            TabIndex        =   76
            Top             =   2490
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Estacionam. 1:"
            Height          =   285
            Left            =   60
            TabIndex        =   75
            Top             =   2160
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "Inmueble:"
            Height          =   285
            Left            =   60
            TabIndex        =   74
            Top             =   1830
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Nombre Perito:"
            Height          =   285
            Left            =   60
            TabIndex        =   61
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_FecEmi 
            Caption         =   "F. Emisión OT:"
            Height          =   315
            Left            =   7470
            TabIndex        =   58
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   3720
            Width           =   1365
         End
         Begin VB.Label Label11 
            Caption         =   "Número Informe:"
            Height          =   285
            Left            =   60
            TabIndex        =   55
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label9 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   60
            TabIndex        =   54
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Empresa Peritaje:"
            Height          =   285
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   36
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
            Height          =   495
            Left            =   600
            TabIndex        =   37
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Tasación del Inmueble"
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
            Left            =   4920
            TabIndex        =   38
            Top             =   120
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
            Picture         =   "AteCli_frm_016.frx":18FE
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2265
         Left            =   30
         TabIndex        =   39
         Top             =   1590
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   3995
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
            TabIndex        =   40
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
            TabIndex        =   41
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
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   50
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   8820
            TabIndex        =   59
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   90
            Left            =   30
            TabIndex        =   63
            Top             =   1410
            Width           =   12705
            _Version        =   65536
            _ExtentX        =   22410
            _ExtentY        =   159
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
            BorderWidth     =   1
            BevelOuter      =   0
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnl_RecDoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   64
            Top             =   1530
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
         Begin Threed.SSPanel pnl_PagGas 
            Height          =   315
            Left            =   1620
            TabIndex        =   66
            Top             =   1860
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
         Begin VB.Label Label27 
            Caption         =   "F. Pago Gastos:"
            Height          =   315
            Left            =   60
            TabIndex        =   67
            Top             =   1860
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "F. Recep. Docum.:"
            Height          =   315
            Left            =   60
            TabIndex        =   65
            Top             =   1530
            Width           =   1455
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   60
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   51
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   46
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   68
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Left            =   10560
            Picture         =   "AteCli_frm_016.frx":21C8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11280
            Picture         =   "AteCli_frm_016.frx":24D2
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_016.frx":27DC
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
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
            TabIndex        =   73
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   72
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   71
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   70
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   69
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Tasaci_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_RecDoc     As String
Dim l_str_PagGas     As String
Dim l_str_IniEva     As String
Dim l_str_Aprueb     As String
Dim l_str_Rechaz     As String
Dim l_int_NumEva     As Integer
Dim l_dbl_ComVta     As Double

Dim l_arr_EmpPer()   As moddat_tpo_Genera

Private Sub cmb_EmpPer_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_EmpPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpPer_Click
   End If
End Sub

Private Sub cmd_Aprueb_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaTas, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaTas, 12, 0, txt_Observ.Text, 0, 0) Then
      Exit Sub
   End If
         
   'Verificar si la Instancia de Seguros ha sido aprobada
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(modatecli_g_con_EvaSeg)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   g_rst_Genera.MoveFirst

   If g_rst_Genera!SEGUIM_SITUAC <> 1 Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      
      Exit Sub
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, modatecli_g_con_AprCre) Then
      Exit Sub
   End If
      
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_AprCre, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, modatecli_g_con_AprCre) Then
      Exit Sub
   End If
   
   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)

   modgen_g_str_Mail_Asunto = "APROBACION DE TASACION DE INMUEBLE Y EVALUACION DE SEGUROS (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
   
   Call cmd_Limpia_Click
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

   'Validación que se encuentre en Instancia de Tasación
   If moddat_g_int_InsAct <> modatecli_g_con_EvaTas Then
      MsgBox "No se encuentra en Instancia de Tasación de Inmueble.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_ActivaItem(False)
   Call fs_Activa(False)

   l_str_RecDoc = ""
   l_str_PagGas = ""
   l_str_IniEva = ""
   l_str_Aprueb = ""
   l_str_Rechaz = ""

   'Obteniendo Información del Seguimiento y Validar
   Call fs_Buscar_SegDet
   
   If Len(Trim(l_str_PagGas)) = 0 Or Len(Trim(l_str_RecDoc)) = 0 Then
      cmd_NueEva.Enabled = False
      cmd_Imprim.Enabled = False
      cmd_RegInf.Enabled = False
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
      
      If Len(Trim(l_str_PagGas)) = 0 Then
         MsgBox "El Cliente no ha pagado los Gastos Administrativos.", vbInformation, modgen_g_str_NomPlt
      End If
      
      If Len(Trim(l_str_RecDoc)) = 0 Then
         MsgBox "No se han recepcionado los Documentos del Cliente.", vbInformation, modgen_g_str_NomPlt
      End If
      
      Call cmd_Limpia_Click
      
      Exit Sub
   End If
   
   If Len(Trim(l_str_Aprueb)) > 0 Then
      MsgBox "El cliente ya ha sido aprobado en esta instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If Len(Trim(l_str_Rechaz)) > 0 Then
      MsgBox "El cliente ya ha sido rechazado en esta instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_Buscar_InfTas
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfTas
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_CodOcu     As Integer
   
   If moddat_g_int_FlgGrb = 1 Then
      If cmb_EmpPer.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_EmpPer)
         Exit Sub
      End If
      
      'Inserta Nuevo Informe
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
   
      Do While moddat_g_int_FlgGOK = False
         Call moddat_gs_FecSis
         
         g_str_Parame = "USP_INSERTA_TRA_EVATAS ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo & "', "
            
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
         g_str_Parame = g_str_Parame & "1)"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If

         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_INSERTA_TRA_EVATAS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   Else
      If Len(Trim(txt_NumInf.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Informe.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumInf)
         Exit Sub
      End If
      
      Call moddat_gs_FecSis
      If CDate(ipp_FecEva.Text) > CDate(moddat_g_str_FecSis) Or CDate(ipp_FecEva.Text) < CDate(pnl_FecEmi.Caption) Then
         MsgBox "La Fecha de Evaluación no puede ser mayor a la de hoy ni menor a la Fecha de Emisión de la OT.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecEva)
         Exit Sub
      End If
   
      If Len(Trim(txt_NomPer.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre del Perito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomPer)
         Exit Sub
      End If
   
      If ipp_ValCom.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ValCom)
         Exit Sub
      End If
      
      If ipp_ValRea.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Fabricación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ValRea)
         Exit Sub
      End If
      
      If ipp_AreTer.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_AreTer)
         Exit Sub
      End If
      
      If ipp_AreCon.Value = 0 Then
         MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_AreCon)
         Exit Sub
      End If
      
      'Registra el Informe de Tasación
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
   
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_MODIFICA_TRA_EVATAS ("
   
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
         g_str_Parame = g_str_Parame & "'" & txt_NumInf.Text & "', "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecEva.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NomPer.Text & "', "
         g_str_Parame = g_str_Parame & CStr(2) & ", "                            'Siempre Dolares
         g_str_Parame = g_str_Parame & CStr(ipp_ValCom.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ValRea.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_AreTer.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_AreCon.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_VCoEs1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_VReEs1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ATeEs1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ACoEs1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_VCoEs2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_VReEs2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ATeEs2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ACoEs2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_VCoDep.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_VReDep.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ATeDep.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ACoDep.Value) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
            
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
         g_str_Parame = g_str_Parame & "1)"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If

         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_MODIFICA_TRA_EVATAS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   End If
   
   'Grabando en Detalle de Seguimiento
   If moddat_g_int_FlgGrb = 1 Then
      r_int_CodOcu = 41
   Else
      r_int_CodOcu = 42
   End If
   
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaTas, r_int_CodOcu, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   MsgBox "Se grabaron los datos correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfTas
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_rst_Direcc  As ADODB.Recordset
   Dim r_str_Direcc  As String
   Dim r_str_UbiGeo  As String
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   Dim r_str_NomVen  As String
   Dim r_str_TelVen  As String
   
   Screen.MousePointer = 11
   
   r_str_Direcc = ""
   r_str_UbiGeo = ""
   r_str_NomVen = ""
   r_str_TelVen = ""
   
   
   'Buscando Información del Inmueble
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLINM_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Direcc, 3) Then
      Exit Sub
   End If
   
   If r_rst_Direcc.BOF And r_rst_Direcc.EOF Then
      r_rst_Direcc.Close
      Set r_rst_Direcc = Nothing
   
      Exit Sub
   End If
   
   r_rst_Direcc.MoveFirst
   
   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Direcc!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(r_rst_Direcc!SOLINM_TIPZON))
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(r_rst_Direcc!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(r_rst_Direcc!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(r_rst_Direcc!SOLINM_UBIGEO))
   
   
   r_str_Direcc = r_str_TipVia & " " & Trim(r_rst_Direcc!SOLINM_NOMVIA) & " " & Trim(r_rst_Direcc!SOLINM_NUMERO)

   If Len(Trim(Trim(r_rst_Direcc!SOLINM_INTDPT))) > 0 Then
      r_str_Direcc = r_str_Direcc & " DPTO/ INT: " & Trim(r_rst_Direcc!SOLINM_INTDPT)
   End If

   If Len(Trim(Trim(r_rst_Direcc!SOLINM_NOMZON))) > 0 Then
      r_str_Direcc = r_str_Direcc & " - " & r_str_TipZon & " " & Trim(r_rst_Direcc!SOLINM_NOMZON)
   End If
      
   r_str_UbiGeo = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   If r_rst_Direcc!SOLINM_TIPPER = 2 Then
      r_str_NomVen = moddat_gf_Consulta_ParDes("203", CStr(r_rst_Direcc!SOLINM_PROTDO)) & "-" & Trim(r_rst_Direcc!SOLINM_PRONDO) & " / " & Trim(r_rst_Direcc!SOLINM_PRORZS)
   Else
      r_str_NomVen = moddat_gf_Consulta_ParDes("203", CStr(r_rst_Direcc!SOLINM_PROTDO)) & "-" & Trim(r_rst_Direcc!SOLINM_PRONDO) & " / " & Trim(r_rst_Direcc!SOLINM_PROAPP) & " " & Trim(r_rst_Direcc!SOLINM_PROAPM) & " " & Trim(r_rst_Direcc!SOLINM_PRONOM)
   End If
   r_str_TelVen = Trim(r_rst_Direcc!SOLINM_PROTL1 & "")
   
   If Len(Trim(r_rst_Direcc!SOLINM_PROTL2 & "")) > 0 Then
      r_str_TelVen = r_str_TelVen & " / " & Trim(r_rst_Direcc!SOLINM_PROTL2 & "")
   End If
   
   r_rst_Direcc.Close
   Set r_rst_Direcc = Nothing

   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   modgen_g_int_NumPag = 1

   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(89) & "Fecha Emisión : " & Format(Date, "dd/mm/yyyy"))
   Call gs_LinImp(Space(89) & "Hora Emisión  :   " & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(40) & "ORDEN DE TRABAJO - TASACION DEL INMUEBLE")
   Call gs_LinImp(Space(40) & "----------------------------------------")
   Call gs_LinImp("")
   

   Call gs_LinImp(Space(5) & "Perito Tasador   : " & cmb_EmpPer.Text)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   Call gs_LinImp(Space(5) & "Número Solicitud : " & pnl_NumSol.Caption)
   Call gs_LinImp(Space(5) & "Producto         : " & moddat_g_str_NomPrd)
   Call gs_LinImp(Space(5) & "Modalidad        : " & moddat_g_str_DesMod)
   Call gs_LinImp(Space(5) & "Cliente          : " & pnl_Client.Caption)
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Valor Venta US$  : " & gf_FormatoNumero(l_dbl_ComVta, 15))
   Call gs_LinImp(Space(5) & "Dirección Inmueb.: " & r_str_Direcc)
   Call gs_LinImp(Space(5) & Space(19) & r_str_UbiGeo)
   Call gs_LinImp(Space(5) & "Propietario      : " & r_str_NomVen)
   Call gs_LinImp(Space(5) & "Teléfono(s)      : " & r_str_TelVen)
   Call gs_LinImp(Space(5) & "Persona Contacto : ")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Sirvase  realizar  el Informe  de  Tasación (Original y Copia) del Inmueble  con datos  adjuntos, para lo cual")
   Call gs_LinImp(Space(5) & "hacemos entrega de los siguientes documentos:")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] JUEGO DE PLANOS DEL INMUELE" & Space(30) & "[  ] COPIA DEL TITULO DE PROPIEDAD")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] MEMORIA DESCRIPTIVA        " & Space(30) & "[  ] CRI O COPIA LITERAL DE FICHA REGISTRAL")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] ESPECIFICACIONES TECNICAS  " & Space(30) & "[  ] CERTIFICADO DE GRAVAMEN (RPU) DEL TERRENO")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] LISTA DE ACABADOS          " & Space(30) & "[  ] PU Y HR DEL TERRENO")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] PRESUPUESTO DE CONSTRUCCION" & Space(30) & "[  ] COPIA DE DECLARATORIA DE FABRICA")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] ESTRUCTURA DE COSTOS       " & Space(30) & "[  ] COPIA DE ESCRITURA DE INDEPENDIZACION")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "[  ] LICENCIA DE CONSTRUCCION   " & Space(30) & "[  ] REGLAMENTO INTERNO")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "OBSERVACIONES:")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "NOMBRE : " & Space(61) & "NOMBRE  : ")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "AREA   : " & Space(61) & "DNI/RUC : ")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-") & Space(50) & String(30, "-"))
   Call gs_LinImp(Space(5) & Space(9) & "P / MICASITA" & Space(9) & Space(50) & Space(6) & "P / PERITO TASADOR")
   
   Call gs_LinImp("")
   
   
   Screen.MousePointer = 0
   
   frm_Imprim_01.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_NueEva_Click()
   moddat_g_int_FlgGrb = 1
   
   Call fs_LimpiaItem
   
   'Activando Botones
   cmd_Grabar.Enabled = True
   cmd_Cancel.Enabled = True
   
   cmd_NueEva.Enabled = False
   cmd_RegInf.Enabled = False
   cmd_Imprim.Enabled = False
   cmd_Aprueb.Enabled = False
   cmd_Rechaz.Enabled = False
   
   'Activando Controles
   cmb_EmpPer.Enabled = True
   txt_NumInf.Enabled = False
   ipp_FecEva.Enabled = False
   txt_NomPer.Enabled = False
   ipp_ValCom.Enabled = False
   ipp_ValRea.Enabled = False
   ipp_AreTer.Enabled = False
   ipp_AreCon.Enabled = False
   txt_Observ.Enabled = False
   
   ipp_VCoEs1.Enabled = False
   ipp_VReEs1.Enabled = False
   ipp_ATeEs1.Enabled = False
   ipp_ACoEs1.Enabled = False

   ipp_VCoEs2.Enabled = False
   ipp_VReEs2.Enabled = False
   ipp_ATeEs2.Enabled = False
   ipp_ACoEs2.Enabled = False
   
   ipp_VCoDep.Enabled = False
   ipp_VReDep.Enabled = False
   ipp_ATeDep.Enabled = False
   ipp_ACoDep.Enabled = False

   Call gs_SetFocus(cmb_EmpPer)
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = modatecli_g_con_EvaTas
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaTas, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaTas, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
   
      r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      r_str_Cadena = r_str_Cadena & Chr(13)
   
   
      modgen_g_str_Mail_Asunto = "RECHAZO DE TASACION DE INMUEBLE  (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = r_str_Cadena
      
      frm_EnvMai_01.Show 1
   
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      
      Call cmd_Limpia_Click
   End If
End Sub

Private Sub cmd_RegInf_Click()
   moddat_g_int_FlgGrb = 2
   
   'Activando Botones
   cmd_Grabar.Enabled = True
   cmd_Cancel.Enabled = True
   
   cmd_NueEva.Enabled = False
   cmd_RegInf.Enabled = False
   cmd_Imprim.Enabled = False
   cmd_Aprueb.Enabled = False
   cmd_Rechaz.Enabled = False
   
   
   cmb_EmpPer.Enabled = False
   
   txt_NumInf.Enabled = True
   ipp_FecEva.Enabled = True
   txt_NomPer.Enabled = True
   ipp_ValCom.Enabled = True
   ipp_ValRea.Enabled = True
   ipp_AreTer.Enabled = True
   ipp_AreCon.Enabled = True
   
   ipp_VCoEs1.Enabled = True
   ipp_VReEs1.Enabled = True
   ipp_ATeEs1.Enabled = True
   ipp_ACoEs1.Enabled = True
   ipp_VCoEs2.Enabled = True
   ipp_VReEs2.Enabled = True
   ipp_ATeEs2.Enabled = True
   ipp_ACoEs2.Enabled = True
   ipp_VCoDep.Enabled = True
   ipp_VReDep.Enabled = True
   ipp_ATeDep.Enabled = True
   ipp_ACoDep.Enabled = True
   
   txt_Observ.Enabled = True
   
   Call gs_SetFocus(txt_NumInf)
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
   
   Call moddat_gs_Carga_LisIte(cmb_EmpPer, l_arr_EmpPer, 1, "507")
End Sub

Private Sub ipp_ACoDep_Change()
   Call ipp_AreCon_Change
End Sub

Private Sub ipp_ACoEs1_Change()
   Call ipp_AreCon_Change
End Sub

Private Sub ipp_ACoEs2_Change()
   Call ipp_AreCon_Change
End Sub

Private Sub ipp_AreCon_Change()
   pnl_TotACo.Caption = Format(CDbl(ipp_AreCon.Text) + CDbl(ipp_ACoEs1.Text) + CDbl(ipp_ACoEs2.Text) + CDbl(ipp_ACoDep.Text), "###,###,##0.00") & " "
End Sub

Private Sub ipp_AreTer_Change()
   pnl_TotATe.Caption = Format(CDbl(ipp_AreTer.Text) + CDbl(ipp_ATeEs1.Text) + CDbl(ipp_ATeEs2.Text) + CDbl(ipp_ATeDep.Text), "###,###,##0.00") & " "
End Sub

Private Sub ipp_ATeDep_Change()
   Call ipp_AreTer_Change
End Sub

Private Sub ipp_ATeEs1_Change()
   Call ipp_AreTer_Change
End Sub

Private Sub ipp_ATeEs2_Change()
   Call ipp_AreTer_Change
End Sub

Private Sub ipp_FecEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomPer)
   End If
End Sub

Private Sub ipp_ValCom_Change()
   pnl_TotVCo.Caption = Format(CDbl(ipp_ValCom.Text) + CDbl(ipp_VCoEs1.Text) + CDbl(ipp_VCoEs2.Text) + CDbl(ipp_VCoDep.Text), "###,###,##0.00") & " "
End Sub

Private Sub ipp_ValCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea)
   End If
End Sub

Private Sub ipp_ValRea_Change()
   pnl_TotVRe.Caption = Format(CDbl(ipp_ValRea.Text) + CDbl(ipp_VReEs1.Text) + CDbl(ipp_VReEs2.Text) + CDbl(ipp_VReDep.Text), "###,###,##0.00") & " "
End Sub

Private Sub ipp_ValRea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreTer)
   End If
End Sub

Private Sub ipp_AreTer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon)
   End If
End Sub

Private Sub ipp_AreCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VCoEs1)
   End If
End Sub

Private Sub ipp_VCoDep_Change()
   Call ipp_ValCom_Change
End Sub

Private Sub ipp_VCoEs1_Change()
   Call ipp_ValCom_Change
End Sub

Private Sub ipp_VCoEs1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VReEs1)
   End If
End Sub

Private Sub ipp_VCoEs2_Change()
   Call ipp_ValCom_Change
End Sub

Private Sub ipp_VReDep_Change()
   Call ipp_ValRea_Change
End Sub

Private Sub ipp_VReEs1_Change()
   Call ipp_ValRea_Change
End Sub

Private Sub ipp_VReEs1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ATeEs1)
   End If
End Sub

Private Sub ipp_ATeEs1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ACoEs1)
   End If
End Sub

Private Sub ipp_ACoEs1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VCoEs2)
   End If
End Sub

Private Sub ipp_VCoEs2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VReEs2)
   End If
End Sub

Private Sub ipp_VReEs2_Change()
   Call ipp_ValRea_Change
End Sub

Private Sub ipp_VReEs2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ATeEs2)
   End If
End Sub

Private Sub ipp_ATeEs2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ACoEs2)
   End If
End Sub

Private Sub ipp_ACoEs2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VCoDep)
   End If
End Sub

Private Sub ipp_VCoDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_VReDep)
   End If
End Sub

Private Sub ipp_VReDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ATeDep)
   End If
End Sub

Private Sub ipp_ATeDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ACoDep)
   End If
End Sub

Private Sub ipp_ACoDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
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
   pnl_IniEva.Caption = ""
   
   pnl_RecDoc.Caption = ""
   pnl_PagGas.Caption = ""
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   cmb_EmpPer.ListIndex = -1
   txt_NumInf.Text = ""
   Call moddat_gs_FecSis
   ipp_FecEva.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   
   txt_NomPer.Text = ""
   ipp_ValCom.Value = 0
   ipp_ValRea.Value = 0
   ipp_AreTer.Value = 0
   ipp_AreCon.Value = 0
   txt_Observ.Text = ""
   
   ipp_VCoEs1.Value = 0
   ipp_VReEs1.Value = 0
   ipp_ATeEs1.Value = 0
   ipp_ACoEs1.Value = 0
   
   ipp_VCoEs2.Value = 0
   ipp_VReEs2.Value = 0
   ipp_ATeEs2.Value = 0
   ipp_ACoEs2.Value = 0
   
   ipp_VCoDep.Value = 0
   ipp_VReDep.Value = 0
   ipp_ATeDep.Value = 0
   ipp_ACoDep.Value = 0
   
   pnl_TotVCo.Caption = "0.00 "
   pnl_TotVRe.Caption = "0.00 "
   pnl_TotATe.Caption = "0.00 "
   pnl_TotACo.Caption = "0.00 "
   
   pnl_FecEmi.Caption = ""
   pnl_FecEmi.Visible = False
   lbl_FecEmi.Visible = False
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_NueEva.Enabled = Not p_Habilita
   cmd_Imprim.Enabled = Not p_Habilita
   cmd_RegInf.Enabled = Not p_Habilita
   cmd_Aprueb.Enabled = Not p_Habilita
   cmd_Rechaz.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   cmb_EmpPer.Enabled = p_Habilita
   txt_NumInf.Enabled = p_Habilita
   ipp_FecEva.Enabled = p_Habilita
   
   txt_NomPer.Enabled = p_Habilita
   ipp_ValCom.Enabled = p_Habilita
   ipp_ValRea.Enabled = p_Habilita
   ipp_AreTer.Enabled = p_Habilita
   ipp_AreCon.Enabled = p_Habilita
   
   ipp_VCoEs1.Enabled = p_Habilita
   ipp_VReEs1.Enabled = p_Habilita
   ipp_ATeEs1.Enabled = p_Habilita
   ipp_ACoEs1.Enabled = p_Habilita
   
   ipp_VCoEs2.Enabled = p_Habilita
   ipp_VReEs2.Enabled = p_Habilita
   ipp_ATeEs2.Enabled = p_Habilita
   ipp_ACoEs2.Enabled = p_Habilita
   
   ipp_VCoDep.Enabled = p_Habilita
   ipp_VReDep.Enabled = p_Habilita
   ipp_ATeDep.Enabled = p_Habilita
   ipp_ACoDep.Enabled = p_Habilita
   
   txt_Observ.Enabled = p_Habilita

   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_NueEva.Enabled = p_Habilita
   cmd_Imprim.Enabled = p_Habilita
   cmd_RegInf.Enabled = p_Habilita
   cmd_Aprueb.Enabled = p_Habilita
   cmd_Rechaz.Enabled = p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
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
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
   
   'Valor Compra-Venta
   l_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_EvaTas) & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_FecOcu = Right(CStr(g_rst_Princi!SEGDET_FECOCU), 2) & "/" & Mid(CStr(g_rst_Princi!SEGDET_FECOCU), 5, 2) & "/" & Left(CStr(g_rst_Princi!SEGDET_FECOCU), 4)
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 11:    l_str_IniEva = r_str_FecOcu
         Case 12:    l_str_Aprueb = r_str_FecOcu
         Case 13:    l_str_Rechaz = r_str_FecOcu
         Case 23:    l_str_RecDoc = r_str_FecOcu
         Case 25:    l_str_PagGas = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If Len(Trim(l_str_IniEva)) > 0 Then
      pnl_IniEva.Caption = l_str_IniEva
   End If
   
   If Len(Trim(l_str_RecDoc)) > 0 Then
      pnl_RecDoc.Caption = l_str_RecDoc
   End If
   
   If Len(Trim(l_str_PagGas)) > 0 Then
      pnl_PagGas.Caption = l_str_PagGas
   End If
End Sub

Private Sub txt_NumInf_GotFocus()
   Call gs_SelecTodo(txt_NumInf)
End Sub

Private Sub txt_NumInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecEva)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;:;")
   End If
End Sub

Private Sub txt_NomPer_GotFocus()
   Call gs_SelecTodo(txt_NomPer)
End Sub

Private Sub txt_NomPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ .")
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Buscar_InfTas()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      cmd_NueEva.Enabled = True
      cmd_Imprim.Enabled = False
      cmd_RegInf.Enabled = False
     
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
     
      cmd_Grabar.Enabled = False
      cmd_Cancel.Enabled = False
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!EVATAS_SITUAC = 1 Then
      cmd_Imprim.Enabled = True
      cmd_RegInf.Enabled = True
      cmd_NueEva.Enabled = False
     
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
   End If
   
   If g_rst_Princi!EVATAS_SITUAC = 2 Then
      cmd_NueEva.Enabled = True
      cmd_Imprim.Enabled = False
      cmd_RegInf.Enabled = True
     
      cmd_Aprueb.Enabled = True
      cmd_Rechaz.Enabled = True
   End If
   
   'Cargar Datos de Evaluación
   cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Princi!EVATAS_CODEMP) - 1
   
   pnl_FecEmi.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEMI))
   pnl_FecEmi.Visible = True
   lbl_FecEmi.Visible = True
   
   txt_NumInf.Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
   ipp_FecEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
   txt_NomPer.Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
   ipp_ValCom.Value = g_rst_Princi!EVATAS_VALCOM
   ipp_ValRea.Value = g_rst_Princi!EVATAS_VALFAB
   ipp_AreTer.Value = g_rst_Princi!EVATAS_ARETER
   ipp_AreCon.Value = g_rst_Princi!EVATAS_ARECON
   ipp_VCoEs1.Value = g_rst_Princi!EVATAS_VCOES1
   ipp_VReEs1.Value = g_rst_Princi!EVATAS_VREES1
   ipp_ATeEs1.Value = g_rst_Princi!EVATAS_ATEES1
   ipp_ACoEs1.Value = g_rst_Princi!EVATAS_ACOES1
   ipp_VCoEs2.Value = g_rst_Princi!EVATAS_VCOES2
   ipp_VReEs2.Value = g_rst_Princi!EVATAS_VREES2
   ipp_ATeEs2.Value = g_rst_Princi!EVATAS_ATEES2
   ipp_ACoEs2.Value = g_rst_Princi!EVATAS_ACOES2
   ipp_VCoDep.Value = g_rst_Princi!EVATAS_VCODEP
   ipp_VReDep.Value = g_rst_Princi!EVATAS_VREDEP
   ipp_ATeDep.Value = g_rst_Princi!EVATAS_ATEDEP
   ipp_ACoDep.Value = g_rst_Princi!EVATAS_ACODEP
   txt_Observ.Text = Trim(g_rst_Princi!EVATAS_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

