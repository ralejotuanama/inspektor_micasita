VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngSol_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   1695
   ClientTop       =   1305
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8355
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   14737
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   85
         Top             =   7560
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_LisOpe 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_002.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   92
            ToolTipText     =   "Lista de Operaciones Crediticias"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_ActEco 
            Height          =   675
            Left            =   6690
            Picture         =   "AteCli_frm_002.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatCyg 
            Height          =   675
            Left            =   7380
            Picture         =   "AteCli_frm_002.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Gastos 
            Height          =   675
            Left            =   8070
            Picture         =   "AteCli_frm_002.frx":0797
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Información Financiera"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Refere 
            Height          =   675
            Left            =   8760
            Picture         =   "AteCli_frm_002.frx":08BA
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Referencias Personales"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_002.frx":0BC4
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Datos del Inmueble"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatSol 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_002.frx":148E
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Datos del Crédito"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_LisRec 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_002.frx":1798
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Lista de Solicitudes Rechazadas"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_002.frx":1AA2
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5895
         Left            =   30
         TabIndex        =   50
         Top             =   1620
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   10398
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
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8190
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   90
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1950
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   750
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodSex 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1080
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   1950
            TabIndex        =   13
            Text            =   "cmb_Paises"
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   14
            Text            =   "cmb_DptNac"
            Top             =   1740
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   1950
            TabIndex        =   15
            Text            =   "cmb_PrvNac"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   16
            Text            =   "cmb_DstNac"
            Top             =   2070
            Width           =   3315
         End
         Begin VB.ComboBox cmb_EstCiv 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2400
            Width           =   3315
         End
         Begin VB.ComboBox cmb_RegCyg 
            Height          =   315
            Left            =   8190
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2400
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2730
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   8190
            TabIndex        =   20
            Text            =   "cmb_Profes"
            Top             =   2730
            Width           =   3315
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   3060
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   3060
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   3870
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1950
            MaxLength       =   120
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   4200
            Width           =   3315
         End
         Begin VB.TextBox txt_Numero 
            Height          =   315
            Left            =   8190
            MaxLength       =   15
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   4200
            Width           =   1640
         End
         Begin VB.TextBox txt_Interi 
            Height          =   315
            Left            =   9870
            MaxLength       =   15
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   4200
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   4530
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   4530
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1950
            TabIndex        =   36
            Text            =   "cmb_DptDir"
            Top             =   4860
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8190
            TabIndex        =   37
            Text            =   "cmb_PrvDir"
            Top             =   4860
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1950
            TabIndex        =   38
            Text            =   "cmb_DstDir"
            Top             =   5190
            Width           =   3315
         End
         Begin VB.CheckBox chk_DirEle 
            Caption         =   "Autoriz. Corresp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9900
            TabIndex        =   23
            Top             =   3090
            Width           =   1485
         End
         Begin VB.TextBox txt_Telefo 
            Height          =   315
            Left            =   1950
            MaxLength       =   8
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   5520
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   5190
            Width           =   3315
         End
         Begin EditLib.fpLongInteger ipp_DepEc1 
            Height          =   315
            Left            =   8190
            TabIndex        =   25
            Top             =   3390
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpDateTime ipp_FecNac 
            Height          =   315
            Left            =   1950
            TabIndex        =   12
            Top             =   1410
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
         Begin EditLib.fpLongInteger ipp_DepEc2 
            Height          =   315
            Left            =   8820
            TabIndex        =   26
            Top             =   3390
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_DepEc3 
            Height          =   315
            Left            =   9480
            TabIndex        =   27
            Top             =   3390
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_DepEc4 
            Height          =   315
            Left            =   10170
            TabIndex        =   28
            Top             =   3390
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_DepEc5 
            Height          =   315
            Left            =   10800
            TabIndex        =   29
            Top             =   3390
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
            Height          =   90
            Left            =   30
            TabIndex        =   51
            Top             =   3750
            Width           =   11505
            _Version        =   65536
            _ExtentX        =   20294
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
         Begin EditLib.fpLongInteger ipp_NumDep 
            Height          =   315
            Left            =   1950
            TabIndex        =   24
            Top             =   3390
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin Threed.SSPanel pnl_EdaCli 
            Height          =   315
            Left            =   3300
            TabIndex        =   88
            Top             =   1410
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin Threed.SSPanel pnl_MesCli 
            Height          =   315
            Left            =   4230
            TabIndex        =   90
            Top             =   1410
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "240 "
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
         Begin VB.Label Label31 
            Caption         =   "Meses"
            Height          =   285
            Left            =   4770
            TabIndex        =   91
            Top             =   1470
            Width           =   555
         End
         Begin VB.Label Label30 
            Caption         =   "Años"
            Height          =   285
            Left            =   3840
            TabIndex        =   89
            Top             =   1470
            Width           =   555
         End
         Begin VB.Label Label29 
            Caption         =   "Apellido Casada:"
            Height          =   285
            Left            =   6210
            TabIndex        =   87
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   86
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   77
            Top             =   90
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   76
            Top             =   750
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "Sexo:"
            Height          =   315
            Left            =   90
            TabIndex        =   75
            Top             =   1080
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   74
            Top             =   1410
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   73
            Top             =   1740
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   72
            Top             =   1740
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   71
            Top             =   2070
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   70
            Top             =   2070
            Width           =   1905
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   315
            Left            =   90
            TabIndex        =   69
            Top             =   2400
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Régimen Conyugal:"
            Height          =   315
            Left            =   6210
            TabIndex        =   68
            Top             =   2400
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   90
            TabIndex        =   67
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión:"
            Height          =   315
            Left            =   6210
            TabIndex        =   66
            Top             =   2730
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   90
            TabIndex        =   65
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6210
            TabIndex        =   64
            Top             =   3060
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "Edades Depend. Econom.:"
            Height          =   285
            Left            =   6210
            TabIndex        =   63
            Top             =   3390
            Width           =   2055
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   62
            Top             =   3870
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   90
            TabIndex        =   61
            Top             =   4200
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6210
            TabIndex        =   60
            Top             =   4200
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   90
            TabIndex        =   59
            Top             =   4530
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6210
            TabIndex        =   58
            Top             =   4530
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   90
            TabIndex        =   57
            Top             =   4860
            Width           =   1905
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6210
            TabIndex        =   56
            Top             =   4860
            Width           =   1905
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   90
            TabIndex        =   55
            Top             =   5190
            Width           =   1905
         End
         Begin VB.Label Label27 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   90
            TabIndex        =   54
            Top             =   5520
            Width           =   1485
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6210
            TabIndex        =   53
            Top             =   5190
            Width           =   1485
         End
         Begin VB.Label Label38 
            Caption         =   "Nro. Depend. Econom.:"
            Height          =   285
            Left            =   90
            TabIndex        =   52
            Top             =   3390
            Width           =   2055
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   825
         Left            =   30
         TabIndex        =   78
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.ComboBox cmb_Modali 
            Height          =   315
            Left            =   5550
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   3735
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   5550
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   3735
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_002.frx":1EE4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   10110
            Picture         =   "AteCli_frm_002.frx":2326
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   9390
            Picture         =   "AteCli_frm_002.frx":2630
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2415
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1950
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   2415
         End
         Begin VB.Label Label40 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   4680
            TabIndex        =   82
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label39 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   4680
            TabIndex        =   81
            Top             =   90
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   80
            Top             =   90
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Docum. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   79
            Top             =   390
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   83
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
            Height          =   495
            Left            =   630
            TabIndex        =   84
            Top             =   60
            Width           =   6465
            _Version        =   65536
            _ExtentX        =   11404
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Ingreso de Solicitud de Crédito"
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
            Picture         =   "AteCli_frm_002.frx":293A
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()   As moddat_tpo_Genera
Dim l_arr_Paises()   As moddat_tpo_Genera
Dim l_arr_Profes()   As moddat_tpo_Genera
Dim l_arr_Parame()   As moddat_tpo_Genera
Dim l_int_FlgCmb     As Integer
Dim l_str_Paises     As String
Dim l_str_Profes     As String
Dim l_str_DptNac     As String
Dim l_str_PrvNac     As String
Dim l_str_DstNac     As String
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String

Private Sub cmb_CodSex_Click()
   Call gs_SetFocus(ipp_FecNac)
End Sub

Private Sub cmb_CodSex_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodSex_Click
   End If
End Sub

Private Sub cmb_DptNac_Change()
   l_str_DptNac = cmb_DptNac.Text
End Sub

Private Sub cmb_DptNac_Click()
   If cmb_DptNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvNac.Clear
         cmb_DstNac.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvNac)
      End If
   End If
End Sub

Private Sub cmb_DptNac_GotFocus()
   Call SendMessage(cmb_DptNac.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_DptNac = cmb_DptNac.Text
End Sub

Private Sub cmb_DptNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptNac, l_str_DptNac)
      l_int_FlgCmb = True
      
      cmb_PrvNac.Clear
      cmb_DstNac.Clear
      If cmb_DptNac.ListIndex > -1 Then
         l_str_DptNac = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvNac)
   End If
End Sub

Private Sub cmb_DptNac_LostFocus()
   Call SendMessage(cmb_DptNac.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_EstCiv_Click()
   cmb_RegCyg.Enabled = False
   Call gs_SetFocus(cmb_NivEst)
   
   If cmb_EstCiv.ListIndex > -1 Then
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         cmb_RegCyg.Enabled = True
         Call gs_SetFocus(cmb_RegCyg)
      Else
         cmb_RegCyg.ListIndex = -1
      End If
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         cmd_DatCyg.Enabled = True
      Else
         Call modatecli_gs_Limpia_DatGen(2)
         ReDim modatecli_g_arr_CygActEco(0)
         
         'Datos Actividades Económicas Cliente Cónyuge
         modatecli_g_str_CodCiu_Cyg = ""
         modatecli_g_str_GirCom_Cyg = ""
         modatecli_g_str_SecEco_Cyg = ""
         modatecli_g_int_TDoEmp_Cyg = 0
         modatecli_g_str_NDoEmp_Cyg = ""
         modatecli_g_int_ActSec_Cyg = 0
   
         atecli_int_CliCyg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Cónyuge)
   
         'Inicializando Arreglos de Solicitudes Rechazadas
         ReDim modatecli_g_arr_CygRec(0)

         'Inicializando Arreglos de Operaciones Vigentes
         ReDim modatecli_g_arr_CygOpe(0)

         'Inicializando Flag de Datos Ingresados
         modatecli_g_int_CygDatGen = 1
   
         'Inicializando Variables de DOI Cónyuge
         moddat_g_int_CygTDo = 0
         moddat_g_str_CygNDo = ""
         
         cmd_DatCyg.Enabled = False
      End If
   End If
End Sub

Private Sub cmb_EstCiv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EstCiv_Click
   End If
End Sub

Private Sub cmb_Modali_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Modali_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Modali_Click
   End If
End Sub

Private Sub cmb_NivEst_Click()
   Call gs_SetFocus(cmb_Profes)
End Sub

Private Sub cmb_NivEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NivEst_Click
   End If
End Sub

Private Sub cmb_Produc_Click()
   If cmb_Produc.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_ParPrd_ComboItem(cmb_Modali, Right(l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 3), "003")
      Screen.MousePointer = 0
   End If
   
   Call gs_SetFocus(cmb_Modali)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_PrvNac_Change()
   l_str_PrvNac = cmb_PrvNac.Text
End Sub

Private Sub cmb_PrvNac_Click()
   If cmb_PrvNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstNac.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"), Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstNac)
      End If
   End If
End Sub

Private Sub cmb_PrvNac_GotFocus()
   Call SendMessage(cmb_PrvNac.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_PrvNac = cmb_PrvNac.Text
End Sub

Private Sub cmb_PrvNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvNac, l_str_PrvNac)
      l_int_FlgCmb = True
      
      cmb_DstNac.Clear
      If cmb_PrvNac.ListIndex > -1 Then
         l_str_DstNac = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"), Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstNac)
   End If
End Sub

Private Sub cmb_DstNac_Change()
   l_str_DstNac = cmb_DstNac.Text
End Sub

Private Sub cmb_DstNac_Click()
   If cmb_DstNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_EstCiv)
      End If
   End If
End Sub

Private Sub cmb_DstNac_GotFocus()
   Call SendMessage(cmb_DstNac.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
   l_str_DstNac = cmb_DstNac.Text
End Sub

Private Sub cmb_DstNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstNac, l_str_DstNac)
      l_int_FlgCmb = True
      
      If cmb_DstNac.ListIndex > -1 Then
         l_str_DstNac = ""
      End If
      
      Call gs_SetFocus(cmb_EstCiv)
   End If
End Sub

Private Sub cmb_Paises_Change()
   l_str_Paises = cmb_Paises.Text
   
   cmb_Paises.SelLength = Len(l_str_Paises)
End Sub

Private Sub cmb_Paises_Click()
   If cmb_Paises.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DptNac.Enabled = True
         cmb_PrvNac.Enabled = True
         cmb_DstNac.Enabled = True
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_EstCiv)
         Else
            Call gs_SetFocus(cmb_DptNac)
         End If
      End If
   Else
      cmb_DptNac.ListIndex = -1
      cmb_PrvNac.Clear
      cmb_DstNac.Clear
      
      cmb_DptNac.Enabled = False
      cmb_PrvNac.Enabled = False
      cmb_DstNac.Enabled = False
   
      Call gs_SetFocus(cmb_EstCiv)
   End If
End Sub

Private Sub cmb_Paises_GotFocus()
   Call SendMessage(cmb_Paises.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
End Sub

Private Sub cmb_Paises_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Paises, l_str_Paises)
      l_int_FlgCmb = True
      
      cmb_DptNac.Enabled = True
      cmb_PrvNac.Enabled = True
      cmb_DstNac.Enabled = True
      
      If cmb_Paises.ListIndex > -1 Then
         l_str_Paises = ""
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_EstCiv)
         Else
            Call gs_SetFocus(cmb_DptNac)
         End If
      Else
         cmb_DptNac.ListIndex = -1
         cmb_PrvNac.Clear
         cmb_DstNac.Clear
      
         cmb_DptNac.Enabled = False
         cmb_PrvNac.Enabled = False
         cmb_DstNac.Enabled = False
   
         Call gs_SetFocus(cmb_EstCiv)
      End If
   End If
End Sub

Private Sub cmb_RegCyg_Click()
   Call gs_SetFocus(cmb_NivEst)
End Sub

Private Sub cmb_RegCyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RegCyg_Click
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

Private Sub cmd_ActEco_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   modatecli_g_int_Tip_ActEco = 1
   frm_IngSol_02.Show 1
End Sub

Private Sub cmd_Buscar_Click()
   Dim r_int_Contad     As Integer
   
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      Case 1
         If Len(Trim(txt_NumDoc.Text)) <> 8 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If

      Case 2
         If Len(Trim(txt_NumDoc.Text)) < 0 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
      
      Case 3
         If Len(Trim(txt_NumDoc.Text)) < 0 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_PltPar
            Call gs_SetFocus(txt_NumDoc)
            Exit Sub
         End If
   End Select

   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_con_PltPar
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   
   If cmb_Modali.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Modalidad.", vbExclamation, modgen_g_con_PltPar
      Call gs_SetFocus(cmb_Modali)
      Exit Sub
   End If
   
   Screen.MousePointer = 11

   moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   moddat_g_str_NumDoc = txt_NumDoc.Text
   
   moddat_g_str_CodPrd = Right(l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, 3)
   moddat_g_str_NomPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Nombre
   
   moddat_g_str_CodMod = Format(cmb_Modali.ItemData(cmb_Modali.ListIndex), "00")
   moddat_g_str_DesMod = cmb_Modali.Text
   
   'Buscando Parámetros por Producto
   Call fs_Buscar_ParPrd
   
   Call fs_Activa(False)
   Call fs_Limpia(2)
   
   'Inicializando y/o Limpiando Arreglos
   Call modatecli_gs_Limpia_DatGen(1)     'Cliente Titular - Datos Generales
   Call modatecli_gs_Limpia_DatGen(2)     'Cónyuge - Datos Generales
   
   Call modatecli_gs_Limpia_Refere(1)     'Referencias - Familiar
   Call modatecli_gs_Limpia_Refere(2)     'Referencias - No Familiar
   Call modatecli_gs_Limpia_DatInm        'Datos del Inmueble
   Call modatecli_gs_Limpia_DatCre        'Datos del Crédito
   
   ReDim modatecli_g_arr_Tit_ActEco(0)    'Cliente Titular - Datos Económicos
   ReDim modatecli_g_arr_Cyg_ActEco(0)    'Cónyuge - Datos Económicos
   
   ReDim modatecli_g_arr_IngresInv(0)     'Ingresos - Inversiones
   ReDim modatecli_g_arr_IngresInm(0)     'Ingresos - Inmuebles
   ReDim modatecli_g_arr_IngresAut(0)     'Ingresos - Autos
   ReDim modatecli_g_arr_IngresEns(0)     'Ingresos - Enseres
   ReDim modatecli_g_arr_GastosTar(0)     'Gastos - Tarjetas
   ReDim modatecli_g_arr_GastosFin(0)     'Gastos - Deudas Financieras
   ReDim modatecli_g_arr_GastosNFi(0)     'Gastos - Deudas No Financieras
   ReDim modatecli_g_arr_GastosGas(0)     'Gastos - Gastos Mensuales
   ReDim modatecli_g_arr_DocCre(0)        'Documentos Recibidos
   
   'Datos Actividades Económicas Cliente Titular
   modatecli_g_str_CodCiu_Tit = ""
   modatecli_g_str_GirCom_Tit = ""
   modatecli_g_str_SecEco_Tit = ""
   modatecli_g_int_TDoEmp_Tit = 0
   modatecli_g_str_NDoEmp_Tit = ""
   modatecli_g_int_ActPri_Tit = 0
   modatecli_g_int_ActSec_Tit = 0
   
   'Datos Actividades Económicas Cliente Cónyuge
   modatecli_g_str_CodCiu_Cyg = ""
   modatecli_g_str_GirCom_Cyg = ""
   modatecli_g_str_SecEco_Cyg = ""
   modatecli_g_int_TDoEmp_Cyg = 0
   modatecli_g_str_NDoEmp_Cyg = ""
   modatecli_g_int_ActSec_Cyg = 0
   
   atecli_int_CliReg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Titular)
   atecli_int_CliCyg = 1              'Flag de Registrado en Base de Datos (1 = No / 2 = Si) (Cónyuge)
   
   'Inicializando Arreglos de Solicitudes Rechazadas
   ReDim modatecli_g_arr_LisRec(0)
   ReDim modatecli_g_arr_CygRec(0)

   ReDim modatecli_g_arr_TitOpe(0)
   ReDim modatecli_g_arr_CygOpe(0)

   'Inicializando Flag de Datos Ingresados
   modatecli_g_int_ActEcoTit = 1
   modatecli_g_int_CygDatGen = 1
   modatecli_g_int_IngresTit = 1
   modatecli_g_int_GastosTit = 1
   
   modatecli_g_int_IngRegInm = 1
   modatecli_g_int_GasRegTar = 1
   modatecli_g_int_GasRegFin = 1
   modatecli_g_int_GasRegGas = 1

   modatecli_g_int_RefereTit = 1
   modatecli_g_int_DatInmTit = 1
   modatecli_g_int_DatCreTit = 1
   
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   


   'VALIDACIONES DE CLIENTE
   'Validar que Cliente no se encuentre en Base Negativa
   If Not atecli_gf_Buscar_BasNeg(moddat_g_int_TipDoc, moddat_g_str_NumDoc) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Titular
   If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Validar que Cliente no tenga una Solicitud de Crédito en Evaluación Como Cónyuge
   If Not atecli_gf_Buscar_SolVig(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   'Buscando Solicitudes Rechazadas
   Call atecli_gs_Buscar_SolRec(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   If UBound(modatecli_g_arr_LisRec) > 0 Then
      cmd_LisRec.Visible = True
   End If
   
   'Buscando Operaciones
   Call atecli_gs_Buscar_CreHip(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   If UBound(modatecli_g_arr_TitOpe) > 0 Then
      If moddat_g_str_CodPrd = "001" Then    'Si Producto es Mivivienda
         MsgBox "El Cliente ya tiene un Crédito Hipotecario registrado.", vbInformation, modgen_g_str_NomPlt
         Call cmd_Limpia_Click
         Exit Sub
      End If
      
      cmd_LisOpe.Visible = True
   End If
   
   'Buscando Información de Cliente Titular
   Call atecli_gs_Buscar_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   
   'Si se encontro Cliente en Base de Datos Asignar Información de Cliente Titular a Controles
   If atecli_int_CliReg = 2 Then
      Call fs_Arreglo_DatCli
   End If
   
   'Si el Titular está registrado como Casado buscar información del Cónyuge
   If moddat_g_int_CygTDo > 0 Then
      Call atecli_gs_Buscar_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2)
   End If
   
   Call gs_SetFocus(txt_ApePat)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DatCyg_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   frm_IngSol_03.Show 1
End Sub

Private Sub cmd_DatInm_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   frm_IngSol_07.Show 1
End Sub

Private Sub cmd_DatSol_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   frm_IngSol_08.Show 1
End Sub

Private Sub cmd_Gastos_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   frm_IngSol_05.Show 1
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_NumSol     As String
   
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   If modatecli_g_int_ActEcoTit = 1 Then
      MsgBox "Debe ingresar la Información Económica del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_ActEco)
      Exit Sub
   End If
   
   If cmd_DatCyg.Enabled Then
      If modatecli_g_int_CygDatGen = 1 Then
         MsgBox "Debe ingresar la Información del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatCyg)
         Exit Sub
      End If
   End If
   
   If modatecli_g_int_GastosTit = 1 Then
      MsgBox "Debe ingresar la Información de Inmuebles, Tarjetas, Deudas y Egresos del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Gastos)
      Exit Sub
   End If
   
   If modatecli_g_int_RefereTit = 1 Then
      MsgBox "Debe ingresar la Información de Referencias Personales del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Refere)
      Exit Sub
   End If
   
   If modatecli_g_int_DatInmTit = 1 Then
      MsgBox "Debe ingresar la Información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatInm)
      Exit Sub
   End If
   
   If modatecli_g_int_DatCreTit = 1 Then
      MsgBox "Debe ingresar la Información del Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_DatSol)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar la Solicitud de Crédito?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabar Datos Generales - Cliente Titular
   'Grabar Datos Generales - Cliente Cónyuge
   'Generar Nro de Solicitud
   'Grabar Ingresos - Inmuebles
   'Grabar Gastos - Tarjetas
   'Grabar Gastos - Deudas
   'Grabar Gastos - Gastos Mensuales
   'Grabar Referencias
   'Grabar Historico de Ejecutivos de Venta
   'Grabar Documentos Recibidos
   'Grabar en Seguimiento
   'Grabar en Detalle de Seguimiento
   'Grabar Datos del Inmueble
   'Grabar Maestro de Solicitudes
   
   
   'Graba Datos Generales de Cliente
   If Not ff_Graba_Cli_DatGen() Then
      Exit Sub
   End If
   
   'Graba Datos Generales de Cónyuge
   If cmd_DatCyg.Enabled Then
      If Not ff_Graba_Cyg_DatGen() Then
         Exit Sub
      End If
   End If


   'Generando Número de Solicitud
   r_str_NumSol = ff_Genera_NumSol()

   'Grabando Información de Ingresos - Inmuebles
   If Not ff_Graba_IngInm(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Tarjetas
   If Not ff_Graba_GasTrj(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Deudas
   If Not ff_Graba_GasDeu(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Gastos Mensuales
   If Not ff_Graba_GasGas(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Referencias
   If Not ff_Graba_Refere(r_str_NumSol) Then
      Exit Sub
   End If

   'Grabando Información de Ejecutivo de Ventas
   If Not ff_Graba_SolEje(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Lista de Documentos Recibidos
   If Not ff_Graba_SolDoc(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando en Seguimiento
   If Not ff_Graba_Seguim(r_str_NumSol) Then
      Exit Sub
   End If
   
   'Grabando Información de Inmueble si tiene identificado el inmueble
   If modatecli_g_arr_DatInm(1).DatInm_InmIde = 1 Then
      If Not ff_Graba_Inmueb(r_str_NumSol) Then
         Exit Sub
      End If
   End If

   'Grabando en Maestro de Solicitudes
   If Not ff_Graba_SolMae(r_str_NumSol) Then
      Exit Sub
   End If

   MsgBox "Ha ingresado correctamente la Solicitud. El Número generado es: " & Left(r_str_NumSol, 3) & "-" & Mid(r_str_NumSol, 4, 3) & "-" & Mid(r_str_NumSol, 7, 2) & "-" & Right(r_str_NumSol, 4), vbInformation, modgen_g_str_NomPlt

   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia(1)
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_TipDoc)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_LisOpe_Click()
   moddat_g_str_NomCli = Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text) & " " & Trim(txt_Nombre.Text)
   frm_LisOpe_01.Show 1
End Sub

Private Sub cmd_LisRec_Click()
   moddat_g_str_NomCli = Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text) & " " & Trim(txt_Nombre.Text)
   
   frm_LisRec_01.Show 1
End Sub

Private Sub cmd_Refere_Click()
   If Not ff_Valida() Then
      Exit Sub
   End If
   
   frm_IngSol_06.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " Ingreso de Solicitud de Crédito"
   
   Call fs_Inicio
   Call fs_Activa(True)
   Call fs_Limpia(1)
      
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
      
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodSex, 1, "207")
   Call moddat_gs_Carga_LisIte_Combo(cmb_EstCiv, 1, "205")
   Call moddat_gs_Carga_LisIte_Combo(cmb_RegCyg, 1, "206")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivEst, 1, "209")
   
   Call moddat_gs_Carga_LisIte(cmb_Paises, l_arr_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_Profes, l_arr_Profes, 1, "501")
      
   Call moddat_gs_Carga_Depart(cmb_DptNac)
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub

Private Sub ipp_DepEc1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc2.Enabled Then
         Call gs_SetFocus(ipp_DepEc2)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc3.Enabled Then
         Call gs_SetFocus(ipp_DepEc3)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc4.Enabled Then
         Call gs_SetFocus(ipp_DepEc4)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_DepEc5.Enabled Then
         Call gs_SetFocus(ipp_DepEc5)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub ipp_DepEc5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   End If
End Sub

Private Sub ipp_FecNac_Change()
   pnl_EdaCli.Caption = Left(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2) & " "
   pnl_MesCli.Caption = Right(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2) & " "
End Sub

Private Sub ipp_FecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Paises)
   End If
End Sub

Private Sub ipp_NumDep_Change()
   If ipp_NumDep.Value = 0 Then
      ipp_DepEc1.Enabled = False
      ipp_DepEc2.Enabled = False
      ipp_DepEc3.Enabled = False
      ipp_DepEc4.Enabled = False
      ipp_DepEc5.Enabled = False
      
      ipp_DepEc1.Value = 0
   Else
      ipp_DepEc1.Enabled = True
      ipp_DepEc2.Enabled = True
      ipp_DepEc3.Enabled = True
      ipp_DepEc4.Enabled = True
      ipp_DepEc5.Enabled = True
      
      If ipp_NumDep.Value < 5 Then
         ipp_DepEc5.Enabled = False
         ipp_DepEc5.Value = 0
      End If
      
      If ipp_NumDep.Value < 4 Then
         ipp_DepEc4.Enabled = False
         ipp_DepEc4.Value = 0
      End If
      
      If ipp_NumDep.Value < 3 Then
         ipp_DepEc3.Enabled = False
         ipp_DepEc3.Value = 0
      End If
      
      If ipp_NumDep.Value < 2 Then
         ipp_DepEc2.Enabled = False
         ipp_DepEc2.Value = 0
      End If
   End If
End Sub

Private Sub ipp_NumDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_NumDep.Value > 0 Then
         Call gs_SetFocus(ipp_DepEc1)
      Else
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeCas)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeCas_GotFocus()
   Call gs_SelecTodo(txt_ApeCas)
End Sub

Private Sub txt_ApeCas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Celula_GotFocus()
   Call gs_SelecTodo(txt_Celula)
End Sub

Private Sub txt_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_DirEle_Change()
   If Len(Trim(txt_DirEle)) > 0 Then
      chk_DirEle.Enabled = True
   Else
      chk_DirEle.Value = 0
      chk_DirEle.Enabled = False
   End If
End Sub

Private Sub txt_NumDoc_LostFocus()
   If cmb_TipDoc.ListIndex > -1 Then
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
   End If
End Sub

Private Sub txt_Telefo_GotFocus()
   Call gs_SelecTodo(txt_Telefo)
End Sub

Private Sub txt_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ActEco)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telefo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodSex)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumDep)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Numero_GotFocus()
   Call gs_SelecTodo(txt_Numero)
End Sub

Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Interi)
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

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Produc)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Limpia(ByVal p_TipLim As Integer)
   Dim r_int_Contad  As Integer
   
   If p_TipLim = 1 Then
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      
      cmb_Produc.ListIndex = -1
      ReDim l_arr_Modali(0)
      cmb_Modali.Clear
   End If
   
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   cmb_CodSex.ListIndex = -1
   
   Call moddat_gs_FecSis
   ipp_FecNac.Text = Format(CDate(moddat_g_str_FecSis) - CDate(18 * 365.25), "dd/mm/yyyy")
   pnl_EdaCli.Caption = "0 "
   pnl_MesCli.Caption = "0 "
   
   cmb_Paises.ListIndex = -1
   cmb_DptNac.ListIndex = -1
   cmb_PrvNac.Clear
   cmb_DstNac.Clear
   cmb_DptNac.Enabled = False
   cmb_PrvNac.Enabled = False
   cmb_DstNac.Enabled = False
   cmb_EstCiv.ListIndex = -1
   cmb_RegCyg.ListIndex = -1
   cmb_RegCyg.Enabled = False
   cmd_DatCyg.Enabled = False
   cmb_NivEst.ListIndex = -1
   cmb_Profes.ListIndex = -1
   txt_DirEle.Text = ""
   txt_Celula.Text = ""
   
   ipp_NumDep.Value = 0
   ipp_DepEc1.Value = 0
   ipp_DepEc2.Value = 0
   ipp_DepEc3.Value = 0
   ipp_DepEc4.Value = 0
   ipp_DepEc5.Value = 0
   
   ipp_DepEc1.Enabled = False
   ipp_DepEc2.Enabled = False
   ipp_DepEc3.Enabled = False
   ipp_DepEc4.Enabled = False
   ipp_DepEc5.Enabled = False
   
   chk_DirEle.Value = 0
   chk_DirEle.Enabled = False
   
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_Numero.Text = ""
   txt_Interi.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Telefo.Text = ""
   
   cmd_LisRec.Visible = False
   cmd_LisOpe.Visible = False
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

Private Sub cmb_Profes_Change()
   l_str_Profes = cmb_Profes.Text
End Sub

Private Sub cmb_Profes_Click()
   If cmb_Profes.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Celula)
      End If
   End If
End Sub

Private Sub cmb_Profes_GotFocus()
   l_int_FlgCmb = True
   l_str_Profes = cmb_Profes.Text
End Sub

Private Sub cmb_Profes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Profes, l_str_Profes)
      l_int_FlgCmb = True
      
      If cmb_Profes.ListIndex > -1 Then
         l_str_Profes = ""
      End If
      
      Call gs_SetFocus(txt_Celula)
   End If
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_TipDoc.Enabled = p_Activa
   txt_NumDoc.Enabled = p_Activa
   cmb_Produc.Enabled = p_Activa
   cmb_Modali.Enabled = p_Activa
   
   txt_ApePat.Enabled = Not p_Activa
   txt_ApeMat.Enabled = Not p_Activa
   txt_Nombre.Enabled = Not p_Activa
   cmb_CodSex.Enabled = Not p_Activa
   ipp_FecNac.Enabled = Not p_Activa
   cmb_Paises.Enabled = Not p_Activa
   cmb_DptNac.Enabled = Not p_Activa
   cmb_PrvNac.Enabled = Not p_Activa
   cmb_DstNac.Enabled = Not p_Activa
   cmb_EstCiv.Enabled = Not p_Activa
   cmb_RegCyg.Enabled = Not p_Activa
   cmb_NivEst.Enabled = Not p_Activa
   cmb_Profes.Enabled = Not p_Activa
   txt_Celula.Enabled = Not p_Activa
   txt_DirEle.Enabled = Not p_Activa
   chk_DirEle.Enabled = Not p_Activa
   
   ipp_NumDep.Enabled = Not p_Activa
   ipp_DepEc1.Enabled = Not p_Activa
   ipp_DepEc2.Enabled = Not p_Activa
   ipp_DepEc3.Enabled = Not p_Activa
   ipp_DepEc4.Enabled = Not p_Activa
   ipp_DepEc5.Enabled = Not p_Activa
   
   cmb_TipVia.Enabled = Not p_Activa
   txt_NomVia.Enabled = Not p_Activa
   txt_Numero.Enabled = Not p_Activa
   txt_Interi.Enabled = Not p_Activa
   cmb_TipZon.Enabled = Not p_Activa
   txt_NomZon.Enabled = Not p_Activa
   cmb_DptDir.Enabled = Not p_Activa
   cmb_PrvDir.Enabled = Not p_Activa
   cmb_DstDir.Enabled = Not p_Activa
   txt_Telefo.Enabled = Not p_Activa
   txt_Refere.Enabled = Not p_Activa
   
   cmd_ActEco.Enabled = Not p_Activa
   cmd_DatCyg.Enabled = Not p_Activa
   cmd_Gastos.Enabled = Not p_Activa
   cmd_Refere.Enabled = Not p_Activa
   cmd_DatInm.Enabled = Not p_Activa
   cmd_DatSol.Enabled = Not p_Activa
   cmd_Grabar.Enabled = Not p_Activa
End Sub

Private Function ff_Valida() As Integer
   ff_Valida = False
   
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Function
   End If
   
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Function
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Function
   End If
   
   If cmb_CodSex.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sexo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodSex)
      Exit Function
   Else
      modatecli_g_arr_DatGen(1).DatGen_CodSex = cmb_CodSex.ItemData(cmb_CodSex.ListIndex)
   End If
   
   If Not IsDate(ipp_FecNac.Text) Then
      MsgBox "La Fecha de Nacimiento no es válida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   If CDate(ipp_FecNac.Text) > Date Then
      MsgBox "Debe ingresar una Fecha de Nacimiento valida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If

   Call moddat_gs_FecSis
   
   If Not (CInt(pnl_EdaCli.Caption) >= modatecli_g_int_Par_EdaMin) Then
      MsgBox "El cliente no cumple con los requisitos de Edad requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   If Not ((CInt(pnl_EdaCli.Caption) < modatecli_g_int_Par_EdaMax) Or (CInt(pnl_EdaCli.Caption) = modatecli_g_int_Par_EdaMax And CInt(pnl_MesCli.Caption) = 0)) Then
      MsgBox "El cliente no cumple con los requisitos de Edad requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Function
   End If
   
   If cmb_Paises.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Paises)
      Exit Function
   End If
   
   If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
      If cmb_DptNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptNac)
         Exit Function
      End If
      
      If cmb_PrvNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvNac)
         Exit Function
      End If
      
      If cmb_DstNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstNac)
         Exit Function
      End If
   End If
   
   If cmb_EstCiv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Estado Civil.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EstCiv)
      Exit Function
   End If
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      If cmb_RegCyg.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Régimen Conyugal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_RegCyg)
         Exit Function
      End If
   End If
   
   If cmb_NivEst.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Estudio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NivEst)
      Exit Function
   End If
   
   If cmb_Profes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Profesión u Oficio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Profes)
      Exit Function
   End If
   
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVia)
      Exit Function
   End If
   
   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomVia)
      Exit Function
   End If
   
   If Len(Trim(txt_Numero.Text)) = 0 Then
      MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Numero)
      Exit Function
   End If
   
   If cmb_TipZon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipZon)
      Exit Function
   End If
   
   If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
      If Len(Trim(txt_NomZon.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomZon)
         Exit Function
      End If
   End If
   
   If cmb_DptDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DptDir)
      Exit Function
   End If
   
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvDir)
      Exit Function
   End If
   
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DstDir)
      Exit Function
   End If
   
   moddat_g_str_NomCli = Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text) & " " & Trim(txt_Nombre.Text)
   
   moddat_g_int_EdaAno = CInt(pnl_EdaCli.Caption)
   moddat_g_int_EdaMes = CInt(pnl_MesCli.Caption)
   
   ff_Valida = True
End Function

Private Sub fs_Buscar_ParPrd()
   Dim r_int_NumUIT     As Integer
   Dim r_dbl_ValUIT     As Double
      
   modatecli_g_dbl_Par_ValViv = 0
   modatecli_g_dbl_Par_PorPre = 0
   modatecli_g_dbl_Par_PorApo = 0
   modatecli_g_int_Par_PlaMin = 0
   modatecli_g_int_Par_PlaMax = 0
   modatecli_g_dbl_Par_PreMin = 0
   modatecli_g_dbl_Par_PreMax = 0
   modatecli_g_int_Par_GraMin = 0
   modatecli_g_int_Par_GraMax = 0
   modatecli_g_int_Par_EdaMin = 0
   modatecli_g_int_Par_EdaMax = 0
   modatecli_g_int_Par_EdaTot = 0

   
   'PARAMETROS DE PRODUCTO
   'Plazo de Créditos en Meses
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "101") Then
      modatecli_g_int_Par_PlaMin = modatecli_g_arr_ParPrd_01(1).Genera_ValMin
      modatecli_g_int_Par_PlaMax = modatecli_g_arr_ParPrd_01(1).Genera_ValMax
   End If
   
   'Período de Gracia
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "801", Format(CInt(moddat_g_str_CodMod), "0") & "01") Then
      modatecli_g_int_Par_GraMin = modatecli_g_arr_ParPrd_01(1).Genera_ValMin
      modatecli_g_int_Par_GraMax = modatecli_g_arr_ParPrd_01(1).Genera_ValMax
   End If
   
   'Edad Cliente
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "321") Then
      modatecli_g_int_Par_EdaMin = modatecli_g_arr_ParPrd_01(1).Genera_ValMin
      modatecli_g_int_Par_EdaMax = modatecli_g_arr_ParPrd_01(1).Genera_ValMax
   End If
   
   'Edad Máxima Cliente
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "322") Then
      modatecli_g_int_Par_EdaTot = modatecli_g_arr_ParPrd_01(1).Genera_Cantid
   End If
   
   'Porcentaje de Aporte Mínimo
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "211") Then
      modatecli_g_dbl_Par_PorApo = modatecli_g_arr_ParPrd_01(1).Genera_Cantid
   End If
   
   'Porcentaje de Porc. Préstamo
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "212") Then
      modatecli_g_dbl_Par_PorPre = modatecli_g_arr_ParPrd_01(1).Genera_Cantid
   End If
   
   'Valor Vivienda (En Soles)
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "502") Then
      r_int_NumUIT = modatecli_g_arr_ParPrd_01(1).Genera_Cantid
      r_dbl_ValUIT = moddat_gf_Consulta_ParVal("001", "002")
      
      modatecli_g_dbl_Par_ValViv = r_int_NumUIT * r_dbl_ValUIT
   End If
   
   'Monto Préstamo (En Dolares)
   If moddat_gf_Consulta_ParPrd(modatecli_g_arr_ParPrd_01(), moddat_g_str_CodPrd, "001", "213") Then
      modatecli_g_dbl_Par_PreMin = modatecli_g_arr_ParPrd_01(1).Genera_ValMin
      modatecli_g_dbl_Par_PreMax = modatecli_g_arr_ParPrd_01(1).Genera_ValMax
   End If
End Sub

Private Function ff_Graba_Cli_DatGen() As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_CntErr     As Integer
   Dim r_int_FlgGOK     As Integer
   
   ff_Graba_Cli_DatGen = False
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CLI_DATGEN ("
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApeCas.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex)) & ", "
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         g_str_Parame = g_str_Parame & CStr(cmb_RegCyg.ItemData(cmb_RegCyg.ListIndex)) & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(cmb_NivEst.ItemData(cmb_NivEst.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_CodSex.ItemData(cmb_CodSex.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(ipp_FecNac.Text, "yyyymmdd") & ", "
      
      If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
         g_str_Parame = g_str_Parame & "'" & Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00") & Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00") & Format(cmb_DstNac.ItemData(cmb_DstNac.ListIndex), "00") & "', "
      Else
         g_str_Parame = g_str_Parame & "'000000', "
      End If
      g_str_Parame = g_str_Parame & "'" & l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo & "', "
      
      g_str_Parame = g_str_Parame & CStr(ipp_NumDep.Text) & ", "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_DepEc1.Text, "000") & Format(ipp_DepEc2.Text, "000") & Format(ipp_DepEc3.Text, "000") & Format(ipp_DepEc4.Text, "000") & Format(ipp_DepEc5.Text, "000") & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Numero.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Interi.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Celula.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telefo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
      
      If chk_DirEle.Value = 1 Then
         g_str_Parame = g_str_Parame & "1, "
      Else
         g_str_Parame = g_str_Parame & "2, "
      End If
      
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_ActPri_Tit) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_CodCiu_Tit & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_GirCom_Tit & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_SecEco_Tit & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_TDoEmp_Tit) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_NDoEmp_Tit & "', "
      
      g_str_Parame = g_str_Parame & "'" & l_arr_Profes(cmb_Profes.ListIndex + 1).Genera_Codigo & "', "
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_NumDoc & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      'Magnitud de Empresa
      If modatecli_g_int_ActPri_Tit = 31 Then
         g_str_Parame = g_str_Parame & "'5', "
      Else
         g_str_Parame = g_str_Parame & "'0', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & CStr(atecli_int_CliReg) & ")"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CLI_DATGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   
   'Grabando Actividades Económicas
   moddat_g_int_CntErr = 0

   'Eliminando Anteriores Actividades Económicas
   g_str_Parame = "USP_BORRAR_CLI_ACTECO (" & CStr(moddat_g_int_TipDoc) & ", '" & moddat_g_str_NumDoc & "', 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Function
   End If
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_Tit_ActEco)
      moddat_g_int_FlgGOK = False
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_INSERTA_CLI_ACTECO ("
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_OrdAct) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodAct) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomCom & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodCiu) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Sucurs & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_GiroCd & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_GiroNm & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipVia) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomVia & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Numero & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Interi & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipZon) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomZon & "', "
         g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_DptDir, "00") & Format(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_PrvDir, "00") & Format(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_DstDir, "00") & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Refere & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Telef1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Telef2 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NumFax & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TeleRH & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_AnexRH & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_IngNet) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FreHab) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_CargoC & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_CargoN & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_NomAre & "', "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FecIng)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FecIng), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_NumAnx & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_TelDir & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_Celula & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_DirEle & "', "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FecCes)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FecCes), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_IngNet) & ", "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FecIni)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FecIni), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_ConLoc) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_TDoEmp) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_NDoEmp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_Tl1Emp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_Tl2Emp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_CargoC & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_CargoN & "', "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FecIng)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FecIng), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_IngNet) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_VtaMen) & ", "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_FecIni)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_FecIni), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_RegTri) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_PorPar) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_TipLoc) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_AlqMen) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_NomArr & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_Tl1Arr & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_Tl2Arr & "', "
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_IngNet) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_PorAcc) & ", "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_FecAnt)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_FecAnt), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Direc1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_NomAr1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele11 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele21 & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_AlqMe1) & ", "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq1)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq1), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Direc2 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_NomAr2 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele12 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele22 & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_AlqMe2) & ", "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq2)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq2), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Direc3 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_NomAr3 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele13 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_Tele23 & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_AlqMe3) & ", "
         
         If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq3)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_FIAlq3), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ren_IngNet) & ", "
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         
         g_str_Parame = g_str_Parame & "1) "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         
         'Creando Archivo de Empresas
         If modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FlgEmp = "NR" Or modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FlgEmp = "NR" Or modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_FlgEmp = "NR" Or modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_FlgEmp = "NR" Then
            r_int_CntErr = 0
            r_int_FlgGOK = False
            
            Do While r_int_FlgGOK = False
               g_str_Parame = "USP_INSERTA_EMP_DATGEN ("
               
               If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Dep_FlgEmp)) > 0 Or Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Com_FlgEmp)) > 0 Or Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Acc_FlgEmp)) > 0 Then
                  g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipDoc) & ", "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NumDoc & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_RazSoc & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomCom & "', "
                  g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_GiroCd & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_GiroNm & "', "
                  
                  If Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Sucurs)) = 0 Then
                     g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipVia) & ", "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomVia & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Numero & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Interi & "', "
                     g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TipZon) & ", "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NomZon & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Refere & "', "
                     g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_DptDir, "00") & Format(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_PrvDir, "00") & Format(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_DstDir, "00") & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Telef1 & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Telef2 & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_NumFax & "', "
                  Else
                     g_str_Parame = g_str_Parame & "0, "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "0, "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'000000', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                  End If
                  
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_TeleRH & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_AnexRH & "', "
               ElseIf Len(Trim(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_FlgEmp)) > 0 Then
                  g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_TDoEmp) & ", "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_NDoEmp & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_RazSoc & "', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'000000', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'000000', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_Tl1Emp & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Tit_ActEco(r_int_Contad).ActEco_Ind_Tl2Emp & "', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
               End If
               
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "9, "
               g_str_Parame = g_str_Parame & "1, "
            
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
               g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
               
               g_str_Parame = g_str_Parame & "1) "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  r_int_CntErr = r_int_CntErr + 1
               Else
                  r_int_FlgGOK = True
               End If
            
               If r_int_CntErr = 6 Then
                  If MsgBox("No se pudo completar el procedimiento USP_INSERTA_EMP_DATGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                     Exit Function
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               End If
            Loop
         End If
         
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_CLI_DATGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Function
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      
      Loop
   Next r_int_Contad
   
   ff_Graba_Cli_DatGen = True
End Function

Private Function ff_Graba_Cyg_DatGen()
   Dim r_int_Contad     As Integer
   Dim r_int_CntErr     As Integer
   Dim r_int_FlgGOK     As Integer
   
   ff_Graba_Cyg_DatGen = False
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CLI_DATGEN ("
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_NumDoc & "', "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_ApePat & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_ApeMat & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_ApeCas & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_Nombre & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex)) & ", "
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         g_str_Parame = g_str_Parame & CStr(cmb_RegCyg.ItemData(cmb_RegCyg.ListIndex)) & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_NivEst) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_CodSex) & ", "
      g_str_Parame = g_str_Parame & Format(modatecli_g_arr_DatGen(2).DatGen_FecNac, "yyyymmdd") & ", "
      
      If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
         g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_DatGen(2).DatGen_DptNac, "00") & Format(modatecli_g_arr_DatGen(2).DatGen_PrvNac, "00") & Format(modatecli_g_arr_DatGen(2).DatGen_DstNac, "00") & "', "
      Else
         g_str_Parame = g_str_Parame & "'000000', "
      End If
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_Paises & "', "
      
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'000000000000000', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Numero.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Interi.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_Celula & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telefo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_DirEle & "', "
      
      If modatecli_g_arr_DatGen(2).DatGen_Autori = 1 Then
         g_str_Parame = g_str_Parame & "1, "
      Else
         g_str_Parame = g_str_Parame & "2, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_ActEco) & ", "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_ActPri_Cyg) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_CodCiu_Cyg & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_GirCom_Cyg & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_SecEco_Cyg & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_TDoEmp_Cyg) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_str_NDoEmp_Cyg & "', "
      
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_Profes & "', "
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      'Magnitud de Empresa
      If modatecli_g_int_ActPri_Cyg = 31 Then
         g_str_Parame = g_str_Parame & "'5', "
      Else
         g_str_Parame = g_str_Parame & "'0', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(atecli_int_CliCyg) & ")"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CLI_DATGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   
   'Grabando Actividades Económicas
   moddat_g_int_CntErr = 0

   'Eliminando Anteriores Actividades Económicas
   g_str_Parame = "USP_BORRAR_CLI_ACTECO (" & CStr(modatecli_g_arr_DatGen(2).DatGen_TipDoc) & ", '" & modatecli_g_arr_DatGen(2).DatGen_NumDoc & "', 2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Function
   End If
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_Cyg_ActEco)
      moddat_g_int_FlgGOK = False
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_INSERTA_CLI_ACTECO ("
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_NumDoc & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_OrdAct) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodAct) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomCom & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodCiu) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Sucurs & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_GiroCd & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_GiroNm & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipVia) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomVia & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Numero & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Interi & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipZon) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomZon & "', "
         g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_DptDir, "00") & Format(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_PrvDir, "00") & Format(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_DstDir, "00") & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Refere & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Telef1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Telef2 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NumFax & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TeleRH & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_AnexRH & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_IngNet) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FreHab) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_CargoC & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_CargoN & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_NomAre & "', "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FecIng)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FecIng), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_NumAnx & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_TelDir & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_Celula & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_DirEle & "', "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FecCes)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FecCes), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_IngNet) & ", "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FecIni)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FecIni), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_ConLoc) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_TDoEmp) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_NDoEmp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_Tl1Emp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_Tl2Emp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_CargoC & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_CargoN & "', "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FecIng)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FecIng), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_IngNet) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_VtaMen) & ", "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_FecIni)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_FecIni), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_RegTri) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_PorPar) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_TipLoc) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_AlqMen) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_NomArr & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_Tl1Arr & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_Tl2Arr & "', "
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_IngNet) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_PorAcc) & ", "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_FecAnt)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_FecAnt), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Direc1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_NomAr1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele11 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele21 & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_AlqMe1) & ", "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq1)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq1), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Direc2 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_NomAr2 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele12 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele22 & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_AlqMe2) & ", "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq2)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq2), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Direc3 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_NomAr3 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele13 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_Tele23 & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_AlqMe3) & ", "
         
         If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq3)) > 0 Then
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_FIAlq3), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ren_IngNet) & ", "
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         
         g_str_Parame = g_str_Parame & "1) "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         
         'Creando Archivo de Empresas
         If modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FlgEmp = "NR" Or modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FlgEmp = "NR" Or modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_FlgEmp = "NR" Or modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_FlgEmp = "NR" Then
            r_int_CntErr = 0
            r_int_FlgGOK = False
            
            Do While r_int_FlgGOK = False
               g_str_Parame = "USP_INSERTA_EMP_DATGEN ("
               
               If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Dep_FlgEmp)) > 0 Or Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Com_FlgEmp)) > 0 Or Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Acc_FlgEmp)) > 0 Then
                  g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipDoc) & ", "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NumDoc & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_RazSoc & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomCom & "', "
                  g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_GiroCd & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_GiroNm & "', "
                  
                  If Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Sucurs)) = 0 Then
                     g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipVia) & ", "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomVia & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Numero & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Interi & "', "
                     g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TipZon) & ", "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NomZon & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Refere & "', "
                     g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_DptDir, "00") & Format(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_PrvDir, "00") & Format(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_DstDir, "00") & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Telef1 & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Telef2 & "', "
                     g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_NumFax & "', "
                  Else
                     g_str_Parame = g_str_Parame & "0, "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "0, "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'000000', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                     g_str_Parame = g_str_Parame & "'', "
                  End If
                  
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_TeleRH & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_AnexRH & "', "
               ElseIf Len(Trim(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_FlgEmp)) > 0 Then
                  g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_TDoEmp) & ", "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_NDoEmp & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_RazSoc & "', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'000000', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'000000', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_Tl1Emp & "', "
                  g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Cyg_ActEco(r_int_Contad).ActEco_Ind_Tl2Emp & "', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
                  g_str_Parame = g_str_Parame & "'', "
               End If
               
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "9, "
               g_str_Parame = g_str_Parame & "1, "
               
            
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
               g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
               
               g_str_Parame = g_str_Parame & "1) "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  r_int_CntErr = r_int_CntErr + 1
               Else
                  r_int_FlgGOK = True
               End If
            
               If r_int_CntErr = 6 Then
                  If MsgBox("No se pudo completar el procedimiento USP_INSERTA_EMP_DATGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                     Exit Function
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               End If
            Loop
         End If
         
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_CLI_DATGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Function
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      
      Loop
   Next r_int_Contad
   
   ff_Graba_Cyg_DatGen = True
End Function

Private Function ff_Genera_NumSol() As String
   Dim r_lng_NumSol     As Long
   Dim r_str_NumSol     As String
   
   ff_Genera_NumSol = ""
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   g_str_Parame = "SELECT * FROM CRE_FOLIOS WHERE "
   g_str_Parame = g_str_Parame & "FOLIOS_TIPFOL = 1 AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_CODSUC = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "FOLIOS_PERANO = " & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumSol = 1
   Else
      r_lng_NumSol = g_rst_Genera!FOLIOS_NUMERO + 1
   End If

   r_str_NumSol = moddat_g_str_CodPrd & modgen_g_str_CodSuc & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & Format(r_lng_NumSol, "0000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_CRE_FOLIOS ("
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumSol) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      
      g_str_Parame = g_str_Parame & "1, "
      
      If r_lng_NumSol = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_FOLIOS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Genera_NumSol = r_str_NumSol
End Function

Private Function ff_Graba_IngInm(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_IngInm = False

   If modatecli_g_int_IngRegInm = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_IngresInm)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLINB ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipInm) & ", "
            
            g_str_Parame = g_str_Parame & Format(CDate(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_FecAdq), "yyyymmdd") & ", "
            
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_ImpVal) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_IngresInm(r_int_Contad).IngInm_TipMon) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_IngresInm(r_int_Contad).IngInm_Direcc & "', "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLINB. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_IngInm = True
End Function

Private Function ff_Graba_GasGas(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasGas = False

   If modatecli_g_int_GasRegGas = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosGas)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLEYM ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_TipGas) & ", "
            g_str_Parame = g_str_Parame & "1, "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosGas(r_int_Contad).GasGas_ImpVal) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLEYM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasGas = True
End Function

Private Function ff_Graba_GasDeu(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasDeu = False

   If modatecli_g_int_GasRegFin = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosFin)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLDEU ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosFin(r_int_Contad).GasFin_InsFin & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosFin(r_int_Contad).GasFin_NumOpe & "', "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_TipMon) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MonOto) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_SalPag) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_CuoMen) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosFin(r_int_Contad).GasFin_MesPag) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLDEU. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasDeu = True
End Function

Private Function ff_Graba_GasTrj(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_GasTrj = False

   If modatecli_g_int_GasRegTar = 1 Then
      For r_int_Contad = 1 To UBound(modatecli_g_arr_GastosTar)
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_INSERTA_CRE_SOLTRJ ("
         
            g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
            g_str_Parame = g_str_Parame & CStr(r_int_Contad) & ", "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_InsFin & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_NumTar & "', "
            g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipTar & "', "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_TipMon) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_LinCre) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_SalPag) & ", "
            g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_GastosTar(r_int_Contad).GasTar_MonMin) & ", "
         
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
            g_str_Parame = g_str_Parame & "1)"
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLTRJ. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Function
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      Next r_int_Contad
   End If
   
   ff_Graba_GasTrj = True
End Function

Private Function ff_Graba_SolDoc(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   Call moddat_gs_FecSis
   
   ff_Graba_SolDoc = False
   
   For r_int_Contad = 1 To UBound(modatecli_g_arr_DocCre)
      If Not moddat_gf_Inserta_SolDoc(p_NumSol, modatecli_g_arr_DocCre(r_int_Contad).DocCre_TipDoc, moddat_g_str_CodPrd, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodAct, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodGrp, modatecli_g_arr_DocCre(r_int_Contad).DocCre_CodIte, Format(CDate(moddat_g_str_FecSis), "yyyymmdd")) Then
         Exit Function
      End If
   Next r_int_Contad
   
   ff_Graba_SolDoc = True
End Function

Private Function ff_Graba_Refere(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_Refere = False
   
   'Grabando Referencia Familiar
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "1, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "1, "                                                                   'Número de Referencia (Familiar)
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(1).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(1).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando Referencia No Familiar
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLREF ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "2, "                                                                   'Tipo de Referencia (Familiar)
      g_str_Parame = g_str_Parame & "1, "                                                                   'Número de Referencia (Familiar)
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_Refere(2).Refere_TipPar) & ", "                    'Tipo de Parentesco
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_ApePat & "', "                   'Apellido Paterno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_ApeMat & "', "                   'Apellido Materno
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Nombre & "', "                   'Nombres
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Telefo & "', "                   'Teléfono
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_Refere(2).Refere_Celula & "', "                   'Celular
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLREF. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_Refere = True
End Function

Private Function ff_Graba_Inmueb(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_Inmueb = False
   

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLINM ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & "1, "                                                          'Número de Registro
      g_str_Parame = g_str_Parame & "1, "                                                          'Situación (Activo)
         
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "                               'Tipo DOI Cliente Titular
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "                              'Numero DOI Cliente Titular
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipInm) & ", "           'Tipo de Inmueble
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_UsoInm) & ", "           'Uso de Inmueble
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipVia) & ", "           'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomVia & "', "          'Nombre de Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Numero & "', "          'Número en Via
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Interi & "', "          'Interior / Dpto
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipZon) & ", "           'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomZon & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Refere & "', "          'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_DatInm(1).DatInm_DptDir, "00") & Format(modatecli_g_arr_DatInm(1).DatInm_PrvDir, "00") & Format(modatecli_g_arr_DatInm(1).DatInm_DstDir, "00") & "', "
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_InmPry) & ", "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_CodPry & "', "
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_NomPry & "', "
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_MCSPry) & ", "
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_TipPro) & ", "
      
      'Si es Persona Natural
      If modatecli_g_arr_DatInm(1).DatInm_TipPro = 1 Then
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Nat_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_ApePat & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_ApeMat & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_Nombre & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Nat_CodSex) & ", "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Nat_EstCiv) & ", "
         
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_Telef1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_Telef2 & "', "
         
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Nat_CygTDo) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_CygNDo & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_CygApp & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_CygApm & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_CygNom & "', "
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Nat_CygSex) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_CygTl1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Nat_CygTl2 & "', "
      Else
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Jur_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_RazSoc & "', "
         
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Jur_TipVia) & ", "           'Tipo de Via
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_NomVia & "', "          'Nombre de Via
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_Numero & "', "          'Número en Via
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_Interi & "', "          'Interior / Dpto
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_Jur_TipZon) & ", "           'Tipo de Zona
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_NomZon & "', "          'Nombre de Zona
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_Refere & "', "          'Referencia
         g_str_Parame = g_str_Parame & "'" & Format(modatecli_g_arr_DatInm(1).DatInm_Jur_DptDir, "00") & Format(modatecli_g_arr_DatInm(1).DatInm_Jur_PrvDir, "00") & Format(modatecli_g_arr_DatInm(1).DatInm_Jur_DstDir, "00") & "', "
         
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_Telef1 & "', "
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_Jur_Telef2 & "', "
         
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                               'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLINM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_Inmueb = True
End Function

Private Function ff_Graba_Seguim(ByVal p_NumSol As String) As Integer
   ff_Graba_Seguim = False

   If Not moddat_gf_Inserta_Seguim(p_NumSol, modatecli_g_con_IngSol) Then
      Exit Function
   End If
   
   If Not moddat_gf_Inserta_SegDet(p_NumSol, modatecli_g_con_IngSol, 11, 0, "", 0, 0) Then
      Exit Function
   End If

   ff_Graba_Seguim = True
End Function

Private Function ff_Graba_SolEje(ByVal p_NumSol As String) As Integer
   Dim r_int_Contad     As Integer

   ff_Graba_SolEje = False
   
   Call moddat_gs_FecSis
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLEJE ("
   
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                                  'Número de Solicitud
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_EjeVta & "', "                   'Ejecutivo de Ventas
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "                                         'Fecha de Asignación
      g_str_Parame = g_str_Parame & "0, "                                                                   'Fecha de Baja
      g_str_Parame = g_str_Parame & "1, "                                                                   'Situación
      g_str_Parame = g_str_Parame & "'', "                                                                  'Observaciones
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                                       'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRE_SOLEJE. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_SolEje = True
End Function

Private Function ff_Graba_SolMae(ByVal p_NumSol As String) As Integer
   Dim r_dbl_TasInt     As Double
   Dim r_dbl_TasCof     As Double
   Dim r_dbl_ComCof     As Double
   
   ff_Graba_SolMae = False
   
   r_dbl_TasInt = 0
   r_dbl_TasCof = 0
   r_dbl_ComCof = 0
      
   'Tasa de Interes de Producto
   If moddat_gf_Consulta_ParPrd(l_arr_Parame, moddat_g_str_CodPrd, "101", Format(modatecli_g_arr_DatCre(1).DatCre_TipMon, "000")) Then
      r_dbl_TasInt = l_arr_Parame(1).Genera_Cantid
   End If
   
   'Tasa de Interes y Comision COFIDE (Si Producto es MiVivienda)
   If moddat_g_str_CodPrd = "001" Then
      If moddat_gf_Consulta_ParPrd(l_arr_Parame, moddat_g_str_CodPrd, "702", "10" & Format(modatecli_g_arr_DatCre(1).DatCre_TipMon, "0")) Then
         r_dbl_TasCof = l_arr_Parame(1).Genera_Cantid
      End If
      
      If moddat_gf_Consulta_ParPrd(l_arr_Parame, moddat_g_str_CodPrd, "702", "20" & Format(modatecli_g_arr_DatCre(1).DatCre_TipMon, "0")) Then
         r_dbl_ComCof = l_arr_Parame(1).Genera_Cantid
      End If
   End If
   
   Call moddat_gs_FecSis
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_CRE_SOLMAE ("
      
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                   'Numero de Solicitud
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "                        'Código Producto
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodMod & "', "                        'Código Modalidad
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "    'Fecha Solicitud
         
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "                            'Tipo DOI Cliente Titular
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "                           'Numero DOI Cliente Titular
      
      'Si Cliente es Casado o Conviviente
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatGen(2).DatGen_TipDoc) & ", "     'Tipo DOI Cliente Cónyuge
         g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatGen(2).DatGen_NumDoc & "', "    'Número DOI Cliente Cónyuge
      Else
         g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI Cliente Cónyuge
         g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI Cliente Cónyuge
      End If
      
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipMon) & ", "        'Tipo de Moneda
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_EjeVta & "', "       'Ejecutivo de Ventas
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_EmpSegDes & "', "                        'Empresa de Seguro de Desgravamen
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_TipSeg) & ", "        'Tipo de Seguro
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_EmpSegViv & "', "                        'Empresa de Seguro de Vivienda
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_CuoAno) & ", "        'Número de Cuotas x Año
      g_str_Parame = g_str_Parame & CStr((modatecli_g_arr_DatCre(1).DatCre_PlaAno) * 12 + modatecli_g_arr_DatCre(1).DatCre_PlaMes) & ", "    'Plazo en Meses
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_PerGra) & ", "        'Período de Gracia
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_CuoMen) & ", "        'Cuota Mensual
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_DiaVct) & ", "        'Dia de Vencimiento
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Dol_MonSol) & ", "    'Monto Solicitado US$
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Sol_MonSol) & ", "    'Monto Solicitado Soles
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Pre_MonSol) & ", "    'Monto Solicitado M.Prestamo
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Dol_ComVta) & ", "    'Compra Venta US$
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Dol_ApoPro) & ", "    'Aporte Propio US$
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Dol_TipCam) & ", "    'Tipo Cambio US$
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatCre(1).DatCre_Pre_TipCam) & ", "    'Tipo Cambio M.Prestamo
      g_str_Parame = g_str_Parame & CStr(r_dbl_TasInt) & ", "                                   'Tasa de Interes
      g_str_Parame = g_str_Parame & CStr(r_dbl_TasCof) & ", "                                   'Tasa de Interes COFIDE
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComCof) & ", "                                   'Tasa de Comisión COFIDE
      g_str_Parame = g_str_Parame & "1, "                                                       'Situación
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_IngRegInm) & ", "                      'Flag de Registro de Ingresos Inmuebles
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegTar) & ", "                      'Flag de Registro de Gastos Tarjetas
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegFin) & ", "                      'Flag de Registro de Gastos Deudas
      g_str_Parame = g_str_Parame & CStr(modatecli_g_int_GasRegGas) & ", "                      'Flag de Registro de Gastos Mensuales
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_InmIde) & ", "        'Flag de Inmueble Identificado
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_ZonPo1 & "', "       'Zona Posible Ubicación 01
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_ZonPo2 & "', "       'Zona Posible Ubicación 02
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatInm(1).DatInm_ZonPo3 & "', "       'Zona Posible Ubicación 03
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_NumDor) & ", "        'Número de Dormitorios
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_NumBan) & ", "        'Número de Baños
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_NumEst) & ", "        'Número de Estacionamientos
      g_str_Parame = g_str_Parame & CStr(modatecli_g_arr_DatInm(1).DatInm_AreCon) & ", "        'Area Construida
      g_str_Parame = g_str_Parame & "'" & modatecli_g_arr_DatCre(1).DatCre_Observ & "', "       'Observaciones
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRESOLMAE. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Graba_SolMae = True
End Function

Private Sub fs_Arreglo_DatCli()
   txt_ApePat.Text = modatecli_g_arr_DatGen(1).DatGen_ApePat
   txt_ApeMat.Text = modatecli_g_arr_DatGen(1).DatGen_ApeMat
   txt_Nombre.Text = modatecli_g_arr_DatGen(1).DatGen_Nombre
   
   Call gs_BuscarCombo_Item(cmb_CodSex, modatecli_g_arr_DatGen(1).DatGen_CodSex)
   ipp_FecNac.Text = modatecli_g_arr_DatGen(1).DatGen_FecNac
   
   cmb_Paises.ListIndex = gf_Busca_Arregl(l_arr_Paises, modatecli_g_arr_DatGen(1).DatGen_Paises) - 1
   
   If modatecli_g_arr_DatGen(1).DatGen_Paises = "004028" Then
      Call gs_BuscarCombo_Item(cmb_DptNac, modatecli_g_arr_DatGen(1).DatGen_DptNac)
   
      Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(modatecli_g_arr_DatGen(1).DatGen_DptNac, "00"))
      Call gs_BuscarCombo_Item(cmb_PrvNac, modatecli_g_arr_DatGen(1).DatGen_PrvNac)

      Call moddat_gs_Carga_Distri(cmb_DstNac, Format(modatecli_g_arr_DatGen(1).DatGen_DptNac, "00"), Format(modatecli_g_arr_DatGen(1).DatGen_PrvNac, "00"))
      Call gs_BuscarCombo_Item(cmb_DstNac, modatecli_g_arr_DatGen(1).DatGen_DstNac)
   End If
   
   Call gs_BuscarCombo_Item(cmb_EstCiv, modatecli_g_arr_DatGen(1).DatGen_EstCiv)
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      Call gs_BuscarCombo_Item(cmb_RegCyg, modatecli_g_arr_DatGen(1).DatGen_RegCyg)
   End If

   Call gs_BuscarCombo_Item(cmb_NivEst, modatecli_g_arr_DatGen(1).DatGen_NivEst)
   
   cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, modatecli_g_arr_DatGen(1).DatGen_Profes) - 1
   
   txt_Celula.Text = modatecli_g_arr_DatGen(1).DatGen_Celula
   txt_DirEle.Text = modatecli_g_arr_DatGen(1).DatGen_DirEle
      
   If modatecli_g_arr_DatGen(1).DatGen_Autori = 1 Then
      chk_DirEle.Value = 1
   End If
      
   ipp_NumDep.Value = modatecli_g_arr_DatGen(1).DatGen_DepEco
      
   If ipp_NumDep.Value > 0 Then
      ipp_DepEc1.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 1, 3))
      ipp_DepEc2.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 4, 3))
      ipp_DepEc3.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 7, 3))
      ipp_DepEc4.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 10, 3))
      ipp_DepEc5.Value = CInt(Mid(modatecli_g_arr_DatGen(1).DatGen_Edades, 13, 3))
   End If
   
   Call gs_BuscarCombo_Item(cmb_TipVia, modatecli_g_arr_DatGen(1).DatGen_TipVia)
   txt_NomVia.Text = modatecli_g_arr_DatGen(1).DatGen_NomVia
   txt_Numero.Text = modatecli_g_arr_DatGen(1).DatGen_Numero
   txt_Interi.Text = modatecli_g_arr_DatGen(1).DatGen_IntDpt
   
   Call gs_BuscarCombo_Item(cmb_TipZon, modatecli_g_arr_DatGen(1).DatGen_TipZon)
   txt_NomZon.Text = modatecli_g_arr_DatGen(1).DatGen_NomZon
   
   Call gs_BuscarCombo_Item(cmb_DptDir, modatecli_g_arr_DatGen(1).DatGen_DptDir)

   Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(modatecli_g_arr_DatGen(1).DatGen_DptDir, "00"))
   Call gs_BuscarCombo_Item(cmb_PrvDir, modatecli_g_arr_DatGen(1).DatGen_PrvDir)

   Call moddat_gs_Carga_Distri(cmb_DstDir, Format(modatecli_g_arr_DatGen(1).DatGen_DptDir, "00"), Format(modatecli_g_arr_DatGen(1).DatGen_PrvDir, "00"))
   Call gs_BuscarCombo_Item(cmb_DstDir, modatecli_g_arr_DatGen(1).DatGen_DstDir)
   
   txt_Refere.Text = modatecli_g_arr_DatGen(1).DatGen_Refere
   txt_Telefo.Text = modatecli_g_arr_DatGen(1).DatGen_Telefo
      
   'Obteniendo DNI del Cónyuge o Conviviente
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 1 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      moddat_g_int_CygTDo = modatecli_g_arr_DatGen(1).DatGen_CygTDo
      moddat_g_str_CygNDo = modatecli_g_arr_DatGen(1).DatGen_CygNDo
   End If
End Sub

