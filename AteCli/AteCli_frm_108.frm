VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   1785
   ClientTop       =   2160
   ClientWidth     =   11685
   Icon            =   "AteCli_frm_108.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6465
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   11404
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
         Height          =   4365
         Left            =   30
         TabIndex        =   17
         Top             =   1230
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   7699
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
         Begin VB.TextBox txt_Telef1_2 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   3660
            Width           =   1640
         End
         Begin VB.TextBox txt_Direcc_2 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   2670
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef2_2 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   3660
            Width           =   1640
         End
         Begin VB.TextBox txt_NomArr_2 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   3000
            Width           =   9525
         End
         Begin VB.ComboBox cmb_SegPro 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2340
            Width           =   765
         End
         Begin VB.TextBox txt_Telef1_1 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1530
            Width           =   1640
         End
         Begin VB.TextBox txt_Direcc_1 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   540
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef2_1 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1530
            Width           =   1640
         End
         Begin VB.TextBox txt_NomArr_1 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   870
            Width           =   9525
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2010
            TabIndex        =   0
            Top             =   60
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
         Begin EditLib.fpDateTime ipp_IniAlq_1 
            Height          =   315
            Left            =   2010
            TabIndex        =   3
            Top             =   1200
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   60
            Left            =   60
            TabIndex        =   18
            Top             =   2220
            Width           =   11475
            _Version        =   65536
            _ExtentX        =   20241
            _ExtentY        =   106
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   60
            Left            =   60
            TabIndex        =   30
            Top             =   420
            Width           =   11475
            _Version        =   65536
            _ExtentX        =   20241
            _ExtentY        =   106
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
         Begin EditLib.fpDoubleSingle ipp_AlqMen_1 
            Height          =   315
            Left            =   2010
            TabIndex        =   6
            Top             =   1860
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
         Begin EditLib.fpDateTime ipp_IniAlq_2 
            Height          =   315
            Left            =   2010
            TabIndex        =   10
            Top             =   3330
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
         Begin EditLib.fpDoubleSingle ipp_AlqMen_2 
            Height          =   315
            Left            =   2010
            TabIndex        =   13
            Top             =   3990
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
         Begin VB.Label lbl_General 
            Caption         =   "Alquiler Mensual:"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   37
            Top             =   3990
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Inicio Alquiler:"
            Height          =   315
            Index           =   4
            Left            =   90
            TabIndex        =   36
            Top             =   3330
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Dirección Propiedad:"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   35
            Top             =   2670
            Width           =   1695
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   34
            Top             =   3660
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Arredantario:"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   33
            Top             =   3000
            Width           =   1785
         End
         Begin VB.Label Label11 
            Caption         =   "2da Propiedad:"
            Height          =   285
            Left            =   90
            TabIndex        =   32
            Top             =   2340
            Width           =   1785
         End
         Begin VB.Label lbl_General 
            Caption         =   "Alquiler Mensual:"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   31
            Top             =   1860
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado (S/.):"
            Height          =   285
            Index           =   61
            Left            =   90
            TabIndex        =   23
            Top             =   60
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Inicio Alquiler:"
            Height          =   315
            Index           =   58
            Left            =   90
            TabIndex        =   22
            Top             =   1200
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Dirección Propiedad:"
            Height          =   285
            Index           =   37
            Left            =   90
            TabIndex        =   21
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   20
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Arredantario:"
            Height          =   285
            Index           =   49
            Left            =   90
            TabIndex        =   19
            Top             =   870
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
            TabIndex        =   25
            Top             =   60
            Width           =   10125
            _Version        =   65536
            _ExtentX        =   17859
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes - Actividades Económicas - Rentista"
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
            Picture         =   "AteCli_frm_108.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   26
         Top             =   750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2010
            TabIndex        =   27
            Top             =   60
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07522154 / IKEHARA PUNK MIGUEL ANGEL (1-07521154 / IKEHARA PUNK MIGUEL ANGEL)"
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
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   29
         Top             =   5640
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_108.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10890
            Picture         =   "AteCli_frm_108.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10200
            Picture         =   "AteCli_frm_108.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_SegPro_Click()
   If cmb_SegPro.ListIndex = -1 Then
      txt_Direcc_2.Text = ""
      txt_NomArr_2.Text = ""
      ipp_IniAlq_2.Text = Format(Date, "dd/mm/yyyy")
      txt_Telef1_2.Text = ""
      txt_Telef2_2.Text = ""
      ipp_AlqMen_2.Value = 0
      
      txt_Direcc_2.Enabled = False
      txt_NomArr_2.Enabled = False
      ipp_IniAlq_2.Enabled = False
      txt_Telef1_2.Enabled = False
      txt_Telef2_2.Enabled = False
      ipp_AlqMen_2.Enabled = False
   Else
      If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
         txt_Direcc_2.Enabled = True
         txt_NomArr_2.Enabled = True
         ipp_IniAlq_2.Enabled = True
         txt_Telef1_2.Enabled = True
         txt_Telef2_2.Enabled = True
         ipp_AlqMen_2.Enabled = True
      
         Call gs_SetFocus(txt_Direcc_2)
      Else
         txt_Direcc_2.Text = ""
         txt_NomArr_2.Text = ""
         ipp_IniAlq_2.Text = Format(Date, "dd/mm/yyyy")
         txt_Telef1_2.Text = ""
         txt_Telef2_2.Text = ""
         ipp_AlqMen_2.Value = 0
         
         txt_Direcc_2.Enabled = False
         txt_NomArr_2.Enabled = False
         ipp_IniAlq_2.Enabled = False
         txt_Telef1_2.Enabled = False
         txt_Telef2_2.Enabled = False
         ipp_AlqMen_2.Enabled = False
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_SegPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegPro_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If ipp_IngNet.Value = 0 Then
      MsgBox "El Ingreso Declarado no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If

   If Len(Trim(txt_Direcc_1.Text)) = 0 Then
      MsgBox "Debe ingresar la Dirección de la Propiedad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Direcc_1)
      Exit Sub
   End If

   If Len(Trim(txt_NomArr_1.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArr_1)
      Exit Sub
   End If

   If CDate(ipp_IniAlq_1.Text) > Date Then
      MsgBox "La Fecha de Inicio de Alquiler no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IniAlq_1)
      Exit Sub
   End If

   If Len(Trim(txt_Telef1_1.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Telef1_1)
      Exit Sub
   End If

   If ipp_AlqMen_1.Value = 0 Then
      MsgBox "El Alquiler Mensual no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_AlqMen_1)
      Exit Sub
   End If

   If cmb_SegPro.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente presenta Segunda Propiedad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegPro)
      Exit Sub
   End If
   
   If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
      If Len(Trim(txt_Direcc_2.Text)) = 0 Then
         MsgBox "Debe ingresar la Dirección de la Propiedad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Direcc_2)
         Exit Sub
      End If
   
      If Len(Trim(txt_NomArr_2.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomArr_2)
         Exit Sub
      End If
   
      If CDate(ipp_IniAlq_2.Text) > Date Then
         MsgBox "La Fecha de Inicio de Alquiler no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IniAlq_2)
         Exit Sub
      End If
   
      If Len(Trim(txt_Telef1_2.Text)) = 0 Then
         MsgBox "Debe ingresar el Teléfono del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Telef1_2)
         Exit Sub
      End If
   
      If ipp_AlqMen_2.Value = 0 Then
         MsgBox "El Alquiler Mensual no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_AlqMen_2)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_Inicia_ActEco(moddat_g_int_TipCli, moddat_g_int_OrdAct)
   
   If moddat_g_int_TipCli = 1 Then
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 51
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_IngNet = CDbl(ipp_IngNet.Text)
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Direc1 = txt_Direcc_1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_NomAr1 = txt_NomArr_1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_IniAl1 = ipp_IniAlq_1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele11 = txt_Telef1_1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele21 = txt_Telef2_1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_AlqMe1 = ipp_AlqMen_1.Text
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_SegPro = cmb_SegPro.ItemData(cmb_SegPro.ListIndex)
      
      If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Direc2 = txt_Direcc_2.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_NomAr2 = txt_NomArr_2.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_IniAl2 = ipp_IniAlq_2.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele12 = txt_Telef1_2.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele22 = txt_Telef1_2.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_AlqMe2 = ipp_AlqMen_2.Text
      End If
   Else
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 51
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_IngNet = CDbl(ipp_IngNet.Text)
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Direc1 = txt_Direcc_1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_NomAr1 = txt_NomArr_1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_IniAl1 = ipp_IniAlq_1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele11 = txt_Telef1_1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele21 = txt_Telef2_1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_AlqMe1 = ipp_AlqMen_1.Text
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_SegPro = cmb_SegPro.ItemData(cmb_SegPro.ListIndex)
      
      If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Direc2 = txt_Direcc_2.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_NomAr2 = txt_NomArr_2.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_IniAl2 = ipp_IniAlq_2.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele12 = txt_Telef1_2.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele22 = txt_Telef1_2.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_AlqMe2 = ipp_AlqMen_2.Text
      End If
   End If
   
   moddat_g_int_FlgAct_1 = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_SimCre_Click()
   If moddat_gf_Obtiene_TipCam(1, 2) = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio. No puede ingresar a esta opción.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_SimCre_11.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
      If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 51 Then
         ipp_IngNet.Value = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_IngNet
         
         txt_Direcc_1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Direc1
         txt_NomArr_1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_NomAr1
         ipp_IniAlq_1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_IniAl1
         txt_Telef1_1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele11
         txt_Telef2_1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele21
         ipp_AlqMen_1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_AlqMe1
         
         Call gs_BuscarCombo_Item(cmb_SegPro, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_SegPro)
         
         If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
            txt_Direcc_2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Direc2
            txt_NomArr_2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_NomAr2
            ipp_IniAlq_2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_IniAl2
            txt_Telef1_2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele12
            txt_Telef2_2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_Tele22
            ipp_AlqMen_2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Ren_AlqMe2
            
            txt_Direcc_2.Enabled = True
            txt_NomArr_2.Enabled = True
            ipp_IniAlq_2.Enabled = True
            txt_Telef1_2.Enabled = True
            txt_Telef2_2.Enabled = True
            ipp_AlqMen_2.Enabled = True
         End If
      End If
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & "(" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   
      If moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 51 Then
         ipp_IngNet.Value = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_IngNet
         
         txt_Direcc_1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Direc1
         txt_NomArr_1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_NomAr1
         ipp_IniAlq_1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_IniAl1
         txt_Telef1_1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele11
         txt_Telef2_1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele21
         ipp_AlqMen_1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_AlqMe1
         
         Call gs_BuscarCombo_Item(cmb_SegPro, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_SegPro)
         
         If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
            txt_Direcc_2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Direc2
            txt_NomArr_2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_NomAr2
            ipp_IniAlq_2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_IniAl2
            txt_Telef1_2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele12
            txt_Telef2_2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_Tele22
            ipp_AlqMen_2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Ren_AlqMe2
            
            txt_Direcc_2.Enabled = True
            txt_NomArr_2.Enabled = True
            ipp_IniAlq_2.Enabled = True
            txt_Telef1_2.Enabled = True
            txt_Telef2_2.Enabled = True
            ipp_AlqMen_2.Enabled = True
         End If
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_SegPro, 1, "214")
End Sub

Private Sub fs_Limpia()
   ipp_IngNet.Value = 0
   
   txt_Direcc_1.Text = ""
   txt_NomArr_1.Text = ""
   ipp_IniAlq_1.Text = Format(Date, "dd/mm/yyyy")
   txt_Telef1_1.Text = ""
   txt_Telef2_1.Text = ""
   ipp_AlqMen_1.Value = 0
   
   cmb_SegPro.ListIndex = -1
   
   txt_Direcc_2.Text = ""
   txt_NomArr_2.Text = ""
   ipp_IniAlq_2.Text = Format(Date, "dd/mm/yyyy")
   txt_Telef1_2.Text = ""
   txt_Telef2_2.Text = ""
   ipp_AlqMen_2.Value = 0
   
   txt_Direcc_2.Enabled = False
   txt_NomArr_2.Enabled = False
   ipp_IniAlq_2.Enabled = False
   txt_Telef1_2.Enabled = False
   txt_Telef2_2.Enabled = False
   ipp_AlqMen_2.Enabled = False
End Sub

Private Sub ipp_AlqMen_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_SegPro)
   End If
End Sub

Private Sub ipp_AlqMen_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Direcc_1)
   End If
End Sub

Private Sub ipp_IniAlq_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1_1)
   End If
End Sub

Private Sub ipp_IniAlq_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1_2)
   End If
End Sub

Private Sub txt_Direcc_1_GotFocus()
   Call gs_SelecTodo(txt_Direcc_1)
End Sub

Private Sub txt_Direcc_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomArr_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_NomArr_1_GotFocus()
   Call gs_SelecTodo(txt_NomArr_1)
End Sub

Private Sub txt_NomArr_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IniAlq_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Telef1_1_GotFocus()
   Call gs_SelecTodo(txt_Telef1_1)
End Sub

Private Sub txt_Telef1_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Telef2_1_GotFocus()
   Call gs_SelecTodo(txt_Telef2_1)
End Sub

Private Sub txt_Telef2_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AlqMen_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Direcc_2_GotFocus()
   Call gs_SelecTodo(txt_Direcc_2)
End Sub

Private Sub txt_Direcc_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomArr_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_NomArr_2_GotFocus()
   Call gs_SelecTodo(txt_NomArr_2)
End Sub

Private Sub txt_NomArr_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IniAlq_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Telef1_2_GotFocus()
   Call gs_SelecTodo(txt_Telef1_2)
End Sub

Private Sub txt_Telef1_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Telef2_2_GotFocus()
   Call gs_SelecTodo(txt_Telef2_2)
End Sub

Private Sub txt_Telef2_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AlqMen_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

