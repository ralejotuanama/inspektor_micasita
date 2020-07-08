VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngSol_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   9585
   ClientLeft      =   1830
   ClientTop       =   825
   ClientWidth     =   11580
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9585
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Width           =   11565
      _Version        =   65536
      _ExtentX        =   20399
      _ExtentY        =   16907
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
         TabIndex        =   120
         Top             =   8760
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   10020
            Picture         =   "AteCli_frm_007.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10740
            Picture         =   "AteCli_frm_007.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   1755
         Left            =   30
         TabIndex        =   119
         Top             =   6960
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
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
         Begin VB.ComboBox cmb_NomZo3 
            Height          =   315
            Left            =   9120
            TabIndex        =   59
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   2325
         End
         Begin VB.ComboBox cmb_DstZo3 
            Height          =   315
            Left            =   6750
            TabIndex        =   58
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   2355
         End
         Begin VB.ComboBox cmb_PrvZo3 
            Height          =   315
            Left            =   4380
            TabIndex        =   57
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   2355
         End
         Begin VB.ComboBox cmb_DptZo3 
            Height          =   315
            Left            =   2010
            TabIndex        =   56
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   2355
         End
         Begin VB.ComboBox cmb_NomZo2 
            Height          =   315
            Left            =   9120
            TabIndex        =   55
            Text            =   "cmb_DptDir"
            Top             =   390
            Width           =   2325
         End
         Begin VB.ComboBox cmb_DstZo2 
            Height          =   315
            Left            =   6750
            TabIndex        =   54
            Text            =   "cmb_DptDir"
            Top             =   390
            Width           =   2355
         End
         Begin VB.ComboBox cmb_PrvZo2 
            Height          =   315
            Left            =   4380
            TabIndex        =   53
            Text            =   "cmb_DptDir"
            Top             =   390
            Width           =   2355
         End
         Begin VB.ComboBox cmb_DptZo2 
            Height          =   315
            Left            =   2010
            TabIndex        =   52
            Text            =   "cmb_DptDir"
            Top             =   390
            Width           =   2355
         End
         Begin VB.ComboBox cmb_NomZo1 
            Height          =   315
            Left            =   9120
            TabIndex        =   51
            Text            =   "cmb_DptDir"
            Top             =   60
            Width           =   2325
         End
         Begin VB.ComboBox cmb_DstZo1 
            Height          =   315
            Left            =   6750
            TabIndex        =   50
            Text            =   "cmb_DptDir"
            Top             =   60
            Width           =   2355
         End
         Begin VB.ComboBox cmb_PrvZo1 
            Height          =   315
            Left            =   4380
            TabIndex        =   49
            Text            =   "cmb_DptDir"
            Top             =   60
            Width           =   2355
         End
         Begin VB.ComboBox cmb_DptZo1 
            Height          =   315
            Left            =   2010
            TabIndex        =   48
            Text            =   "cmb_DptDir"
            Top             =   60
            Width           =   2355
         End
         Begin EditLib.fpLongInteger ipp_NumDor 
            Height          =   315
            Left            =   2010
            TabIndex        =   61
            Top             =   1380
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
         Begin EditLib.fpLongInteger ipp_NumBan 
            Height          =   315
            Left            =   2700
            TabIndex        =   62
            Top             =   1380
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
         Begin EditLib.fpDoubleSingle ipp_AreCon 
            Height          =   315
            Left            =   2010
            TabIndex        =   60
            Top             =   1050
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
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
         Begin EditLib.fpLongInteger ipp_NumEst 
            Height          =   315
            Left            =   3390
            TabIndex        =   63
            Top             =   1380
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
         Begin VB.Label Label29 
            Caption         =   "Dpt / Prv / Dst / Zona 3:"
            Height          =   315
            Left            =   60
            TabIndex        =   125
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label18 
            Caption         =   "Dpt / Prv / Dst / Zona 2:"
            Height          =   315
            Left            =   60
            TabIndex        =   124
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label32 
            Caption         =   "Dpt / Prv / Dst / Zona 1:"
            Height          =   315
            Left            =   60
            TabIndex        =   123
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label33 
            Caption         =   "Area Construc. (m2):"
            Height          =   285
            Left            =   60
            TabIndex        =   122
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label31 
            Caption         =   "Nro. Dorm. / Baños / Est.:"
            Height          =   285
            Left            =   60
            TabIndex        =   121
            Top             =   1380
            Width           =   1935
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   465
         Left            =   30
         TabIndex        =   115
         Top             =   660
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   820
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
         Begin VB.ComboBox cmb_InmIde 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1640
         End
         Begin VB.Label Label17 
            Caption         =   "Inmueble Identificado:"
            Height          =   315
            Left            =   90
            TabIndex        =   116
            Top             =   60
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   585
         Left            =   30
         TabIndex        =   67
         Top             =   30
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   1032
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
            Height          =   405
            Left            =   4770
            TabIndex        =   109
            Top             =   120
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   495
            Left            =   600
            TabIndex        =   110
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Datos del Inmueble"
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
            Picture         =   "AteCli_frm_007.frx":074C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   5745
         Left            =   30
         TabIndex        =   68
         Top             =   1170
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   10134
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
         Begin VB.ComboBox cmb_UsoInm 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipInm 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_InmPry 
            Height          =   315
            Left            =   7980
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   7980
            MaxLength       =   120
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_Numero 
            Height          =   315
            Left            =   2070
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.TextBox txt_Interi 
            Height          =   315
            Left            =   3720
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   7980
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   2070
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   7980
            TabIndex        =   9
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   2070
            TabIndex        =   10
            Text            =   "cmb_PrvDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   7980
            TabIndex        =   11
            Text            =   "cmb_DstDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Proyec 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_Proyec 
            Height          =   315
            Left            =   7980
            MaxLength       =   250
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2880
            Width           =   3315
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   90
            Left            =   30
            TabIndex        =   80
            Top             =   2730
            Width           =   11415
            _Version        =   65536
            _ExtentX        =   20135
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
         Begin Threed.SSPanel pnl_Pro_PerNat 
            Height          =   2565
            Left            =   60
            TabIndex        =   95
            Top             =   3150
            Width           =   11325
            _Version        =   65536
            _ExtentX        =   19976
            _ExtentY        =   4524
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
            BevelOuter      =   0
            Begin VB.TextBox txt_Nat_CygTl2 
               Height          =   315
               Left            =   9540
               MaxLength       =   12
               TabIndex        =   32
               Text            =   "Text1"
               Top             =   2190
               Width           =   1635
            End
            Begin VB.TextBox txt_Nat_CygTl1 
               Height          =   315
               Left            =   7890
               MaxLength       =   12
               TabIndex        =   31
               Text            =   "Text1"
               Top             =   2190
               Width           =   1635
            End
            Begin VB.TextBox txt_Nat_Telef2 
               Height          =   315
               Left            =   9540
               MaxLength       =   12
               TabIndex        =   25
               Text            =   "Text1"
               Top             =   1050
               Width           =   1635
            End
            Begin VB.TextBox txt_Nat_Telef1 
               Height          =   315
               Left            =   7890
               MaxLength       =   12
               TabIndex        =   24
               Text            =   "Text1"
               Top             =   1050
               Width           =   1635
            End
            Begin VB.TextBox txt_Nat_CygNDo 
               Height          =   315
               Left            =   7890
               MaxLength       =   12
               TabIndex        =   27
               Text            =   "Text1"
               Top             =   1530
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Nat_CygTdo 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   1530
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_CygNom 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   30
               Text            =   "Text1"
               Top             =   2190
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_CygApm 
               Height          =   315
               Left            =   7890
               MaxLength       =   30
               TabIndex        =   29
               Text            =   "Text1"
               Top             =   1860
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_CygApp 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   28
               Text            =   "Text1"
               Top             =   1860
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Nat_EstCiv 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   1050
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_ApePat 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   19
               Text            =   "Text1"
               Top             =   390
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_ApeMat 
               Height          =   315
               Left            =   7890
               MaxLength       =   30
               TabIndex        =   20
               Text            =   "Text1"
               Top             =   390
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_Nombre 
               Height          =   315
               Left            =   2010
               MaxLength       =   30
               TabIndex        =   21
               Text            =   "Text1"
               Top             =   720
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Nat_CodSex 
               Height          =   315
               Left            =   7890
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   720
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Nat_TipDoc 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   60
               Width           =   3315
            End
            Begin VB.TextBox txt_Nat_NumDoc 
               Height          =   315
               Left            =   7890
               MaxLength       =   12
               TabIndex        =   18
               Text            =   "Text1"
               Top             =   60
               Width           =   3315
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   90
               Left            =   0
               TabIndex        =   108
               Top             =   1410
               Width           =   11415
               _Version        =   65536
               _ExtentX        =   20135
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
            Begin VB.Label Label8 
               Caption         =   "Teléfonos:"
               Height          =   285
               Left            =   6000
               TabIndex        =   112
               Top             =   2190
               Width           =   1485
            End
            Begin VB.Label Label1 
               Caption         =   "Teléfonos:"
               Height          =   285
               Left            =   6000
               TabIndex        =   111
               Top             =   1050
               Width           =   1485
            End
            Begin VB.Label Label14 
               Caption         =   "Nro. Docum. Identidad:"
               Height          =   285
               Left            =   6000
               TabIndex        =   107
               Top             =   1530
               Width           =   1695
            End
            Begin VB.Label Label9 
               Caption         =   "Tipo Docum. Identidad:"
               Height          =   315
               Left            =   30
               TabIndex        =   106
               Top             =   1530
               Width           =   1905
            End
            Begin VB.Label Label7 
               Caption         =   "Nombres:"
               Height          =   285
               Left            =   60
               TabIndex        =   105
               Top             =   2190
               Width           =   1485
            End
            Begin VB.Label Label6 
               Caption         =   "Apellido Materno:"
               Height          =   285
               Left            =   6000
               TabIndex        =   104
               Top             =   1860
               Width           =   1485
            End
            Begin VB.Label Label5 
               Caption         =   "Apellido Paterno:"
               Height          =   285
               Left            =   60
               TabIndex        =   103
               Top             =   1860
               Width           =   1485
            End
            Begin VB.Label Label4 
               Caption         =   "Estado Civil:"
               Height          =   315
               Left            =   60
               TabIndex        =   102
               Top             =   1050
               Width           =   1905
            End
            Begin VB.Label Label66 
               Caption         =   "Apellido Paterno:"
               Height          =   285
               Left            =   60
               TabIndex        =   101
               Top             =   390
               Width           =   1485
            End
            Begin VB.Label Label65 
               Caption         =   "Apellido Materno:"
               Height          =   285
               Left            =   6000
               TabIndex        =   100
               Top             =   390
               Width           =   1485
            End
            Begin VB.Label Label64 
               Caption         =   "Nombres:"
               Height          =   285
               Left            =   60
               TabIndex        =   99
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label Label63 
               Caption         =   "Sexo:"
               Height          =   315
               Left            =   6000
               TabIndex        =   98
               Top             =   720
               Width           =   1545
            End
            Begin VB.Label Label55 
               Caption         =   "Tipo Docum. Identidad:"
               Height          =   315
               Left            =   30
               TabIndex        =   97
               Top             =   60
               Width           =   1905
            End
            Begin VB.Label Label54 
               Caption         =   "Nro. Docum. Identidad:"
               Height          =   285
               Left            =   6000
               TabIndex        =   96
               Top             =   60
               Width           =   1695
            End
         End
         Begin Threed.SSPanel pnl_Pro_PerJur 
            Height          =   2565
            Left            =   60
            TabIndex        =   69
            Top             =   3150
            Width           =   11325
            _Version        =   65536
            _ExtentX        =   19976
            _ExtentY        =   4524
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
            BevelOuter      =   0
            Begin VB.TextBox txt_Jur_Telef1 
               Height          =   315
               Left            =   7890
               MaxLength       =   12
               TabIndex        =   46
               Text            =   "Text1"
               Top             =   2040
               Width           =   1640
            End
            Begin VB.TextBox txt_Jur_Telef2 
               Height          =   315
               Left            =   9540
               MaxLength       =   12
               TabIndex        =   47
               Text            =   "Text1"
               Top             =   2040
               Width           =   1640
            End
            Begin VB.TextBox txt_Jur_NumDoc 
               Height          =   315
               Left            =   7890
               MaxLength       =   12
               TabIndex        =   34
               Text            =   "50"
               Top             =   60
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Jur_TipDoc 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   60
               Width           =   3315
            End
            Begin VB.TextBox txt_Jur_RazSoc 
               Height          =   315
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   35
               Text            =   "Text1"
               Top             =   390
               Width           =   9195
            End
            Begin VB.TextBox txt_Jur_Refere 
               Height          =   315
               Left            =   2010
               MaxLength       =   250
               TabIndex        =   45
               Text            =   "Text1"
               Top             =   2040
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Jur_DstDir 
               Height          =   315
               Left            =   7890
               TabIndex        =   44
               Text            =   "cmb_DstDir"
               Top             =   1710
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Jur_PrvDir 
               Height          =   315
               Left            =   2010
               TabIndex        =   43
               Text            =   "cmb_PrvDir"
               Top             =   1710
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Jur_DptDir 
               Height          =   315
               Left            =   7890
               TabIndex        =   42
               Text            =   "cmb_DptDir"
               Top             =   1380
               Width           =   3315
            End
            Begin VB.TextBox txt_Jur_NomZon 
               Height          =   315
               Left            =   2010
               MaxLength       =   120
               TabIndex        =   41
               Text            =   "Text1"
               Top             =   1380
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Jur_TipZon 
               Height          =   315
               Left            =   7890
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   1050
               Width           =   3315
            End
            Begin VB.TextBox txt_Jur_Interi 
               Height          =   315
               Left            =   3660
               MaxLength       =   15
               TabIndex        =   39
               Text            =   "Text1"
               Top             =   1050
               Width           =   1640
            End
            Begin VB.TextBox txt_Jur_Numero 
               Height          =   315
               Left            =   2010
               MaxLength       =   15
               TabIndex        =   38
               Text            =   "Text1"
               Top             =   1050
               Width           =   1640
            End
            Begin VB.TextBox txt_Jur_NomVia 
               Height          =   315
               Left            =   7890
               MaxLength       =   120
               TabIndex        =   37
               Text            =   "Text1"
               Top             =   720
               Width           =   3315
            End
            Begin VB.ComboBox cmb_Jur_TipVia 
               Height          =   315
               Left            =   2010
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   720
               Width           =   3315
            End
            Begin VB.Label Label11 
               Caption         =   "Teléfonos:"
               Height          =   285
               Left            =   6030
               TabIndex        =   113
               Top             =   2040
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Nro. Docum. Identidad:"
               Height          =   285
               Left            =   6000
               TabIndex        =   94
               Top             =   60
               Width           =   1695
            End
            Begin VB.Label Label2 
               Caption         =   "Tipo Docum. Identidad:"
               Height          =   315
               Left            =   30
               TabIndex        =   93
               Top             =   60
               Width           =   1905
            End
            Begin VB.Label Label43 
               Caption         =   "Razón Social:"
               Height          =   285
               Left            =   60
               TabIndex        =   79
               Top             =   390
               Width           =   1485
            End
            Begin VB.Label Label45 
               Caption         =   "Referencia:"
               Height          =   285
               Left            =   60
               TabIndex        =   78
               Top             =   2040
               Width           =   1485
            End
            Begin VB.Label Label46 
               Caption         =   "Distrito:"
               Height          =   315
               Left            =   6030
               TabIndex        =   77
               Top             =   1710
               Width           =   1905
            End
            Begin VB.Label Label47 
               Caption         =   "Provincia:"
               Height          =   315
               Left            =   60
               TabIndex        =   76
               Top             =   1710
               Width           =   1905
            End
            Begin VB.Label Label48 
               Caption         =   "Departamento:"
               Height          =   315
               Left            =   6030
               TabIndex        =   75
               Top             =   1380
               Width           =   1905
            End
            Begin VB.Label Label49 
               Caption         =   "Nombre Zona:"
               Height          =   285
               Left            =   60
               TabIndex        =   74
               Top             =   1380
               Width           =   1485
            End
            Begin VB.Label Label50 
               Caption         =   "Tipo de Zona:"
               Height          =   315
               Left            =   6030
               TabIndex        =   73
               Top             =   1050
               Width           =   1905
            End
            Begin VB.Label Label51 
               Caption         =   "Nro - Int/Dpto/Mza/Lote:"
               Height          =   285
               Left            =   60
               TabIndex        =   72
               Top             =   1050
               Width           =   2055
            End
            Begin VB.Label Label52 
               Caption         =   "Nombre Vía:"
               Height          =   285
               Left            =   6030
               TabIndex        =   71
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label Label53 
               Caption         =   "Tipo de Vía:"
               Height          =   315
               Left            =   60
               TabIndex        =   70
               Top             =   720
               Width           =   1905
            End
         End
         Begin VB.Label Label12 
            Caption         =   "Uso de Inmueble:"
            Height          =   315
            Left            =   90
            TabIndex        =   118
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Inmueble:"
            Height          =   315
            Left            =   90
            TabIndex        =   117
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "Inmueble en Proyecto:"
            Height          =   315
            Left            =   6060
            TabIndex        =   114
            Top             =   2040
            Width           =   1665
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   90
            TabIndex        =   92
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   6060
            TabIndex        =   91
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   90
            TabIndex        =   90
            Top             =   1050
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   6060
            TabIndex        =   89
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   90
            TabIndex        =   88
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   6060
            TabIndex        =   87
            Top             =   1380
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   90
            TabIndex        =   86
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   6060
            TabIndex        =   85
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   90
            TabIndex        =   84
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label Label13 
            Caption         =   "Proyecto Hipotecario:"
            Height          =   315
            Left            =   90
            TabIndex        =   83
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Nombre Proyecto:"
            Height          =   285
            Left            =   6060
            TabIndex        =   82
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label Label27 
            Caption         =   "Tipo Propietario:"
            Height          =   315
            Left            =   90
            TabIndex        =   81
            Top             =   2880
            Width           =   1905
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_FlgCmb        As Integer
Dim l_str_DptDir        As String
Dim l_str_PrvDir        As String
Dim l_str_DstDir        As String
Dim l_str_Jur_DptDir    As String
Dim l_str_Jur_PrvDir    As String
Dim l_str_Jur_DstDir    As String
Dim l_str_DptZo1        As String
Dim l_str_PrvZo1        As String
Dim l_str_DstZo1        As String
Dim l_str_NomZo1        As String
Dim l_str_DptZo2        As String
Dim l_str_PrvZo2        As String
Dim l_str_DstZo2        As String
Dim l_str_NomZo2        As String
Dim l_str_DptZo3        As String
Dim l_str_PrvZo3        As String
Dim l_str_DstZo3        As String
Dim l_str_NomZo3        As String
Dim l_arr_Proyec()      As moddat_tpo_Genera

Private Sub cmb_InmIde_Click()
   If cmb_InmIde.ListIndex > -1 Then
      If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
         cmb_DptZo1.ListIndex = -1
         cmb_PrvZo1.Clear
         cmb_DstZo1.Clear
         cmb_NomZo1.Clear
         
         cmb_DptZo2.ListIndex = -1
         cmb_PrvZo2.Clear
         cmb_DstZo2.Clear
         cmb_NomZo2.Clear
         
         cmb_DptZo3.ListIndex = -1
         cmb_PrvZo3.Clear
         cmb_DstZo3.Clear
         cmb_NomZo3.Clear
         
         ipp_NumDor.Value = 0
         ipp_NumBan.Value = 0
         ipp_NumEst.Value = 0
         ipp_AreCon.Value = 0
      
         cmb_DptZo1.Enabled = False
         cmb_PrvZo1.Enabled = False
         cmb_DstZo1.Enabled = False
         cmb_NomZo1.Enabled = False
         
         cmb_DptZo2.Enabled = False
         cmb_PrvZo2.Enabled = False
         cmb_DstZo2.Enabled = False
         cmb_NomZo2.Enabled = False
         
         cmb_DptZo3.Enabled = False
         cmb_PrvZo3.Enabled = False
         cmb_DstZo3.Enabled = False
         cmb_NomZo3.Enabled = False
         
         ipp_NumDor.Enabled = False
         ipp_NumBan.Enabled = False
         ipp_AreCon.Enabled = False
         ipp_NumEst.Enabled = False
      
         cmb_TipInm.Enabled = True
         cmb_UsoInm.Enabled = True
         cmb_TipVia.Enabled = True
         txt_NomVia.Enabled = True
         txt_Numero.Enabled = True
         txt_Interi.Enabled = True
         cmb_TipZon.Enabled = True
         txt_NomZon.Enabled = True
         cmb_DptDir.Enabled = True
         cmb_PrvDir.Enabled = True
         cmb_DstDir.Enabled = True
         txt_Refere.Enabled = True
         cmb_InmPry.Enabled = True
         cmb_Proyec.Enabled = True
         txt_Proyec.Enabled = True
         cmb_Proyec.Enabled = False
         txt_Proyec.Enabled = False
         cmb_TipPro.Enabled = True
         
         Call fs_Limpia_Nat
         Call fs_Limpia_Jur
         
         pnl_Pro_PerNat.Visible = True
         pnl_Pro_PerNat.Enabled = False
         
         pnl_Pro_PerJur.Visible = False
         pnl_Pro_PerJur.Enabled = False
         
         Call gs_SetFocus(cmb_TipInm)
      Else
         cmb_TipInm.ListIndex = -1
         cmb_UsoInm.ListIndex = -1
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
         cmb_InmPry.ListIndex = -1
         cmb_Proyec.ListIndex = -1
         txt_Proyec.Text = ""
         cmb_Proyec.Enabled = False
         txt_Proyec.Enabled = False
         cmb_TipPro.ListIndex = -1
         
         Call fs_Limpia_Nat
         Call fs_Limpia_Jur
         
         pnl_Pro_PerNat.Visible = True
         pnl_Pro_PerNat.Enabled = False
         
         pnl_Pro_PerJur.Visible = False
         pnl_Pro_PerJur.Enabled = False
         
         
         cmb_TipInm.Enabled = False
         cmb_UsoInm.Enabled = False
         cmb_TipVia.Enabled = False
         txt_NomVia.Enabled = False
         txt_Numero.Enabled = False
         txt_Interi.Enabled = False
         cmb_TipZon.Enabled = False
         txt_NomZon.Enabled = False
         cmb_DptDir.Enabled = False
         cmb_PrvDir.Enabled = False
         cmb_DstDir.Enabled = False
         txt_Refere.Enabled = False
         cmb_InmPry.Enabled = False
         cmb_Proyec.Enabled = False
         txt_Proyec.Enabled = False
         cmb_Proyec.Enabled = False
         txt_Proyec.Enabled = False
         cmb_TipPro.Enabled = False
         
         'Para Inmueble No Identificado
         cmb_DptZo1.Enabled = True
         cmb_PrvZo1.Enabled = True
         cmb_DstZo1.Enabled = True
         cmb_NomZo1.Enabled = True
         
         cmb_DptZo2.Enabled = True
         cmb_PrvZo2.Enabled = True
         cmb_DstZo2.Enabled = True
         cmb_NomZo2.Enabled = True
         
         cmb_DptZo3.Enabled = True
         cmb_PrvZo3.Enabled = True
         cmb_DstZo3.Enabled = True
         cmb_NomZo3.Enabled = True
         
         ipp_NumDor.Enabled = True
         ipp_NumBan.Enabled = True
         ipp_AreCon.Enabled = True
         ipp_NumEst.Enabled = True
         
         Call gs_SetFocus(cmb_DptZo1)
      End If
   Else
      'Inmueble No Identificado
      cmb_DptZo1.ListIndex = -1
      cmb_PrvZo1.Clear
      cmb_DstZo1.Clear
      cmb_NomZo1.Clear
      
      cmb_DptZo2.ListIndex = -1
      cmb_PrvZo2.Clear
      cmb_DstZo2.Clear
      cmb_NomZo2.Clear
      
      cmb_DptZo3.ListIndex = -1
      cmb_PrvZo3.Clear
      cmb_DstZo3.Clear
      cmb_NomZo3.Clear
      
      ipp_NumDor.Value = 0
      ipp_NumBan.Value = 0
      ipp_NumEst.Value = 0
      ipp_AreCon.Value = 0
   
      cmb_DptZo1.Enabled = False
      cmb_PrvZo1.Enabled = False
      cmb_DstZo1.Enabled = False
      cmb_NomZo1.Enabled = False
      
      cmb_DptZo2.Enabled = False
      cmb_PrvZo2.Enabled = False
      cmb_DstZo2.Enabled = False
      cmb_NomZo2.Enabled = False
      
      cmb_DptZo3.Enabled = False
      cmb_PrvZo3.Enabled = False
      cmb_DstZo3.Enabled = False
      cmb_NomZo3.Enabled = False
      
      ipp_NumDor.Enabled = False
      ipp_NumBan.Enabled = False
      ipp_AreCon.Enabled = False
      ipp_NumEst.Enabled = False

      'Inmueble Identificado
      cmb_TipInm.ListIndex = -1
      cmb_UsoInm.ListIndex = -1
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
      cmb_InmPry.ListIndex = -1
      cmb_Proyec.ListIndex = -1
      txt_Proyec.Text = ""
      cmb_Proyec.Enabled = False
      txt_Proyec.Enabled = False
      cmb_TipPro.ListIndex = -1
      
      Call fs_Limpia_Nat
      Call fs_Limpia_Jur
      
      pnl_Pro_PerNat.Visible = True
      pnl_Pro_PerNat.Enabled = False
      
      pnl_Pro_PerJur.Visible = False
      pnl_Pro_PerJur.Enabled = False
      
      
      cmb_TipInm.Enabled = False
      cmb_UsoInm.Enabled = False
      cmb_TipVia.Enabled = False
      txt_NomVia.Enabled = False
      txt_Numero.Enabled = False
      txt_Interi.Enabled = False
      cmb_TipZon.Enabled = False
      txt_NomZon.Enabled = False
      cmb_DptDir.Enabled = False
      cmb_PrvDir.Enabled = False
      cmb_DstDir.Enabled = False
      txt_Refere.Enabled = False
      cmb_InmPry.Enabled = False
      cmb_Proyec.Enabled = False
      txt_Proyec.Enabled = False
      cmb_Proyec.Enabled = False
      txt_Proyec.Enabled = False
      cmb_TipPro.Enabled = False
   End If
End Sub

Private Sub cmb_InmPry_Click()
   Call gs_SetFocus(cmb_TipPro)
   
   If cmb_InmPry.ListIndex > -1 Then
      Select Case cmb_InmPry.ItemData(cmb_InmPry.ListIndex)
         Case 1
            cmb_Proyec.Enabled = True
            
            Call gs_SetFocus(cmb_Proyec)
         Case 2
            cmb_Proyec.Enabled = False
            txt_Proyec.Enabled = False
            
            cmb_Proyec.ListIndex = -1
            txt_Proyec.Text = ""
      End Select
   End If
End Sub

Private Sub cmb_InmPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_InmPry_Click
   End If
End Sub

Private Sub cmb_Jur_TipDoc_Click()
   Call gs_SetFocus(txt_Jur_NumDoc)
End Sub

Private Sub cmb_Jur_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Jur_TipDoc_Click
   End If
End Sub

Private Sub cmb_Nat_CodSex_Click()
   Call gs_SetFocus(cmb_Nat_EstCiv)
End Sub

Private Sub cmb_Nat_CygTdo_Click()
   If cmb_Nat_CygTdo.ListIndex > -1 Then
      Select Case cmb_Nat_CygTdo.ItemData(cmb_Nat_CygTdo.ListIndex)
         Case 1:  txt_Nat_CygNDo.MaxLength = 8
         Case 2:  txt_Nat_CygNDo.MaxLength = 12
         Case 3:  txt_Nat_CygNDo.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_Nat_CygNDo)
End Sub

Private Sub cmb_Nat_CygTdo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Nat_CygTdo_Click
   End If
End Sub

Private Sub cmb_Nat_EstCiv_Click()
   Call gs_SetFocus(txt_Nat_Telef1)
   
   If cmb_Nat_EstCiv.ListIndex > -1 Then
      If cmb_Nat_EstCiv.ItemData(cmb_Nat_EstCiv.ListIndex) = 2 Then
         cmb_Nat_CygTdo.Enabled = True
         txt_Nat_CygNDo.Enabled = True
         txt_Nat_CygApp.Enabled = True
         txt_Nat_CygApm.Enabled = True
         txt_Nat_CygNom.Enabled = True
      Else
         cmb_Nat_CygTdo.Enabled = False
         txt_Nat_CygNDo.Enabled = False
         txt_Nat_CygApp.Enabled = False
         txt_Nat_CygApm.Enabled = False
         txt_Nat_CygNom.Enabled = False
      
         cmb_Nat_CygTdo.ListIndex = -1
         txt_Nat_CygNDo.Text = ""
         txt_Nat_CygApp.Text = ""
         txt_Nat_CygApm.Text = ""
         txt_Nat_CygNom.Text = ""
      End If
   End If
End Sub

Private Sub cmb_Nat_TipDoc_Click()
   If cmb_Nat_TipDoc.ListIndex > -1 Then
      Select Case cmb_Nat_TipDoc.ItemData(cmb_Nat_TipDoc.ListIndex)
         Case 1:  txt_Nat_NumDoc.MaxLength = 8
         Case 2:  txt_Nat_NumDoc.MaxLength = 12
         Case 3:  txt_Nat_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_Nat_NumDoc)
End Sub

Private Sub cmb_Nat_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Nat_TipDoc_Click
   End If
End Sub

Private Sub cmb_Proyec_Click()
   If cmb_Proyec.ListIndex > -1 Then
      If l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo = "000000" Then
         txt_Proyec.Enabled = True
         
         Call gs_SetFocus(txt_Proyec)
      Else
         txt_Proyec.Text = ""
         txt_Proyec.Enabled = False
         
         Call gs_SetFocus(cmb_TipPro)
      End If
   Else
      txt_Proyec.Enabled = False
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub cmb_Proyec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Proyec_Click
   End If
End Sub

Private Sub cmb_TipInm_Click()
   Call gs_SetFocus(cmb_UsoInm)
End Sub

Private Sub cmb_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipInm_Click
   End If
End Sub

Private Sub cmb_UsoInm_Click()
   Call gs_SetFocus(cmb_TipVia)
End Sub

Private Sub cmb_UsoInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_UsoInm_Click
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipVia_Click
   End If
End Sub

Private Sub cmd_Acepta_Click()
   If cmb_InmIde.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Cliente tiene identificado el Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_InmIde)
      Exit Sub
   End If

   If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
      If cmb_TipInm.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipInm)
         Exit Sub
      End If
   
      If cmb_UsoInm.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Uso de Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_UsoInm)
         Exit Sub
      End If
      
      If cmb_TipVia.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NomVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de la Via.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_Numero.Text)) = 0 Then
         MsgBox "Debe ingresar el Número en la Via.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Numero)
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
   
      If cmb_InmPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar si el Inmueble se encuentra en un Prooyecto Inmobiliario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_InmPry)
         Exit Sub
      End If
   
      'Si el Inmueble está en un Proyecto Inmobiliario
      If cmb_InmPry.ItemData(cmb_InmPry.ListIndex) = 1 Then
         If cmb_Proyec.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Proyecto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Proyec)
            Exit Sub
         End If
         
         'Selecciono Proyecto No Registrado
         If l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo = "000000" Then
            If Len(Trim(txt_Proyec.Text)) = 0 Then
               MsgBox "Debe ingresar el Nombre del Proyecto.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_Proyec)
               Exit Sub
            End If
         End If
      End If
      
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Propietario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
      
      If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 1 Then
         'Persona Natural
         
         If cmb_Nat_TipDoc.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Nat_TipDoc)
            Exit Sub
         End If
      
         If Len(Trim(txt_Nat_NumDoc)) = 0 Then
            MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Nat_NumDoc)
            Exit Sub
         End If
         
         If Len(Trim(txt_Nat_ApePat)) = 0 Then
            MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Nat_ApePat)
            Exit Sub
         End If
      
         If Len(Trim(txt_Nat_ApeMat)) = 0 Then
            MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Nat_ApeMat)
            Exit Sub
         End If
      
         If Len(Trim(txt_Nat_Nombre)) = 0 Then
            MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Nat_Nombre)
            Exit Sub
         End If
         
         If cmb_Nat_CodSex.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Sexo.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Nat_CodSex)
            Exit Sub
         End If
      
         If cmb_Nat_EstCiv.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Estado Civil.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Nat_EstCiv)
            Exit Sub
         End If
         
         If cmb_Nat_EstCiv.ItemData(cmb_Nat_EstCiv.ListIndex) = 2 Then
            If cmb_Nat_CygTdo.ListIndex = -1 Then
               MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_Nat_CygTdo)
               Exit Sub
            End If
         
            If Len(Trim(txt_Nat_CygNDo)) = 0 Then
               MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_Nat_CygNDo)
               Exit Sub
            End If
            
            If Len(Trim(txt_Nat_CygApp)) = 0 Then
               MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_Nat_CygApp)
               Exit Sub
            End If
         
            If Len(Trim(txt_Nat_CygApm)) = 0 Then
               MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_Nat_CygApm)
               Exit Sub
            End If
         
            If Len(Trim(txt_Nat_CygNom)) = 0 Then
               MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_Nat_CygNom)
               Exit Sub
            End If
         End If
      Else
         'Persona Juridica
         
         If cmb_Jur_TipDoc.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Jur_TipDoc)
            Exit Sub
         End If
      
         If Len(Trim(txt_Jur_NumDoc)) = 0 Then
            MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Jur_NumDoc)
            Exit Sub
         End If
         
         If Not gf_Valida_RUC(Mid(txt_Jur_NumDoc.Text, 1, Len(txt_Jur_NumDoc.Text) - 1), Right(txt_Jur_NumDoc.Text, 1)) Then
            MsgBox "Ingrese correctamente el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Jur_NumDoc)
            Exit Sub
         End If
         
         If Len(Trim(txt_Jur_RazSoc)) = 0 Then
            MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Jur_RazSoc)
            Exit Sub
         End If
         
         If cmb_Jur_TipVia.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Jur_TipVia)
            Exit Sub
         End If
         
         If Len(Trim(txt_Jur_NomVia.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de la Via.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Jur_NomVia)
            Exit Sub
         End If
         
         If Len(Trim(txt_Jur_Numero.Text)) = 0 Then
            MsgBox "Debe ingresar el Número en la Via.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_Jur_Numero)
            Exit Sub
         End If
         
         If cmb_Jur_TipZon.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Jur_TipZon)
            Exit Sub
         End If
      
         If cmb_Jur_TipZon.ItemData(cmb_Jur_TipZon.ListIndex) <> 12 Then
            If Len(Trim(txt_Jur_NomZon.Text)) = 0 Then
               MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(txt_Jur_NomZon)
               Exit Sub
            End If
         End If
      
         If cmb_Jur_DptDir.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Departamento.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Jur_DptDir)
            Exit Sub
         End If
      
         If cmb_Jur_PrvDir.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Provincia.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Jur_PrvDir)
            Exit Sub
         End If
      
         If cmb_Jur_DstDir.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Distrito.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Jur_DstDir)
            Exit Sub
         End If
      End If
   Else
      If cmb_DptZo1.ListIndex = -1 Then
         MsgBox "Debe seleccionar al menos el Departamento de una posible Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptZo1)
         Exit Sub
      End If
      
      If cmb_PrvZo1.ListIndex = -1 Then
         MsgBox "Debe seleccionar al menos la Provincia de una posible Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvZo1)
         Exit Sub
      End If
      
      If cmb_DstZo1.ListIndex = -1 Then
         MsgBox "Debe seleccionar al menos el Distrito de una posible Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstZo1)
         Exit Sub
      End If
      
      If cmb_NomZo1.ListIndex = -1 Then
         MsgBox "Debe seleccionar al menos una posible Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_NomZo1)
         Exit Sub
      End If
      
      If cmb_DptZo2.ListIndex > -1 Then
         If cmb_PrvZo2.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Provincia de la 2da Zona .", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_PrvZo2)
            Exit Sub
         End If
         
         If cmb_DstZo2.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Distrito de la 2da Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_DstZo2)
            Exit Sub
         End If
         
         If cmb_NomZo2.ListIndex = -1 Then
            MsgBox "Debe seleccionar la 2da Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_NomZo2)
            Exit Sub
         End If
      End If
   
      If cmb_DptZo3.ListIndex > -1 Then
         If cmb_PrvZo3.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Provincia de la 3ra Zona .", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_PrvZo3)
            Exit Sub
         End If
         
         If cmb_DstZo3.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Distrito de la 3ra Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_DstZo3)
            Exit Sub
         End If
         
         If cmb_NomZo3.ListIndex = -1 Then
            MsgBox "Debe seleccionar la 3ra Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_NomZo3)
            Exit Sub
         End If
      End If
   End If
   
   'Limpiando Arreglo
   Call modatecli_gs_Limpia_DatInm
   
   'Pasando Información al Arreglo
   modatecli_g_arr_DatInm(1).DatInm_InmIde = cmb_InmIde.ItemData(cmb_InmIde.ListIndex)
   
   If cmb_InmIde.ItemData(cmb_InmIde.ListIndex) = 1 Then
      modatecli_g_arr_DatInm(1).DatInm_TipInm = cmb_TipInm.ItemData(cmb_TipInm.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_UsoInm = cmb_UsoInm.ItemData(cmb_UsoInm.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_NomVia = txt_NomVia.Text
      modatecli_g_arr_DatInm(1).DatInm_Numero = txt_Numero.Text
      modatecli_g_arr_DatInm(1).DatInm_Interi = txt_Interi.Text
      modatecli_g_arr_DatInm(1).DatInm_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_NomZon = txt_NomZon.Text
      modatecli_g_arr_DatInm(1).DatInm_DptDir = cmb_DptDir.ItemData(cmb_DptDir.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_PrvDir = cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_DstDir = cmb_DstDir.ItemData(cmb_DstDir.ListIndex)
      modatecli_g_arr_DatInm(1).DatInm_Refere = txt_Refere.Text
      modatecli_g_arr_DatInm(1).DatInm_InmPry = cmb_InmPry.ItemData(cmb_InmPry.ListIndex)
      
      If cmb_Proyec.ListIndex = -1 Then
         modatecli_g_arr_DatInm(1).DatInm_CodPry = ""
      Else
         modatecli_g_arr_DatInm(1).DatInm_CodPry = l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_Codigo
         modatecli_g_arr_DatInm(1).DatInm_MCSPry = l_arr_Proyec(cmb_Proyec.ListIndex + 1).Genera_TipVal
      End If
      
      modatecli_g_arr_DatInm(1).DatInm_NomPry = txt_Proyec.Text
      modatecli_g_arr_DatInm(1).DatInm_TipPro = cmb_TipPro.ItemData(cmb_TipPro.ListIndex)
      
      If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 1 Then
         'Persona Natural
         modatecli_g_arr_DatInm(1).DatInm_Nat_TipDoc = cmb_Nat_TipDoc.ItemData(cmb_Nat_TipDoc.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Nat_NumDoc = txt_Nat_NumDoc.Text
         modatecli_g_arr_DatInm(1).DatInm_Nat_ApePat = txt_Nat_ApePat.Text
         modatecli_g_arr_DatInm(1).DatInm_Nat_ApeMat = txt_Nat_ApeMat.Text
         modatecli_g_arr_DatInm(1).DatInm_Nat_Nombre = txt_Nat_Nombre.Text
         modatecli_g_arr_DatInm(1).DatInm_Nat_CodSex = cmb_Nat_CodSex.ItemData(cmb_Nat_CodSex.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Nat_EstCiv = cmb_Nat_EstCiv.ItemData(cmb_Nat_EstCiv.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Nat_Telef1 = txt_Nat_Telef1.Text
         modatecli_g_arr_DatInm(1).DatInm_Nat_Telef2 = txt_Nat_Telef2.Text
         
         If cmb_Nat_EstCiv.ItemData(cmb_Nat_EstCiv.ListIndex) = 2 Then
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygTDo = cmb_Nat_CygTdo.ItemData(cmb_Nat_CygTdo.ListIndex)
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygNDo = txt_Nat_CygNDo.Text
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygApp = txt_Nat_CygApp.Text
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygApm = txt_Nat_CygApm.Text
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygNom = txt_Nat_CygNom.Text
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygTl1 = txt_Nat_CygTl1.Text
            modatecli_g_arr_DatInm(1).DatInm_Nat_CygTl2 = txt_Nat_CygTl2.Text
            
            If cmb_Nat_CodSex.ItemData(cmb_Nat_CodSex.ListIndex) = 1 Then
               modatecli_g_arr_DatInm(1).DatInm_Nat_CygSex = 2
            Else
               modatecli_g_arr_DatInm(1).DatInm_Nat_CygSex = 1
            End If
         End If
      Else
         'Persona Juridica
         
         modatecli_g_arr_DatInm(1).DatInm_Jur_TipDoc = cmb_Jur_TipDoc.ItemData(cmb_Jur_TipDoc.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Jur_NumDoc = txt_Jur_NumDoc.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_RazSoc = txt_Jur_RazSoc.Text
   
         modatecli_g_arr_DatInm(1).DatInm_Jur_TipVia = cmb_Jur_TipVia.ItemData(cmb_Jur_TipVia.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Jur_NomVia = txt_Jur_NomVia.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_Numero = txt_Jur_Numero.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_Interi = txt_Jur_Interi.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_TipZon = cmb_Jur_TipZon.ItemData(cmb_Jur_TipZon.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Jur_NomZon = txt_Jur_NomZon.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_DptDir = cmb_Jur_DptDir.ItemData(cmb_Jur_DptDir.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Jur_PrvDir = cmb_Jur_PrvDir.ItemData(cmb_Jur_PrvDir.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Jur_DstDir = cmb_Jur_DstDir.ItemData(cmb_Jur_DstDir.ListIndex)
         modatecli_g_arr_DatInm(1).DatInm_Jur_Refere = txt_Jur_Refere.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_Telef1 = txt_Jur_Telef1.Text
         modatecli_g_arr_DatInm(1).DatInm_Jur_Telef2 = txt_Jur_Telef2.Text
      End If
   Else
      modatecli_g_arr_DatInm(1).DatInm_ZonPo1 = Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00") & Format(cmb_PrvZo1.ItemData(cmb_PrvZo1.ListIndex), "00") & Format(cmb_DstZo1.ItemData(cmb_DstZo1.ListIndex), "00") & Format(cmb_NomZo1.ItemData(cmb_NomZo1.ListIndex), "00")
      
      If cmb_NomZo2.ListIndex > -1 Then
         modatecli_g_arr_DatInm(1).DatInm_ZonPo2 = Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00") & Format(cmb_PrvZo2.ItemData(cmb_PrvZo2.ListIndex), "00") & Format(cmb_DstZo2.ItemData(cmb_DstZo2.ListIndex), "00") & Format(cmb_NomZo2.ItemData(cmb_NomZo2.ListIndex), "00")
      End If
      
      If cmb_NomZo3.ListIndex > -1 Then
         modatecli_g_arr_DatInm(1).DatInm_ZonPo3 = Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00") & Format(cmb_PrvZo3.ItemData(cmb_PrvZo3.ListIndex), "00") & Format(cmb_DstZo3.ItemData(cmb_DstZo3.ListIndex), "00") & Format(cmb_NomZo3.ItemData(cmb_NomZo3.ListIndex), "00")
      End If
      
      modatecli_g_arr_DatInm(1).DatInm_NumDor = ipp_NumDor.Value
      modatecli_g_arr_DatInm(1).DatInm_NumBan = ipp_NumBan.Value
      modatecli_g_arr_DatInm(1).DatInm_NumEst = ipp_NumEst.Value
      modatecli_g_arr_DatInm(1).DatInm_AreCon = ipp_AreCon.Value
   End If
   
   modatecli_g_int_DatInmTit = 2
   
   Unload Me
End Sub

Private Sub ipp_AreCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumDor)
   End If
End Sub

Private Sub ipp_NumBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumEst)
   End If
End Sub

Private Sub ipp_NumDor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumBan)
   End If
End Sub

Private Sub ipp_NumEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   End If
End Sub

Private Sub txt_Jur_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Jur_NumDoc)
End Sub

Private Sub txt_Jur_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Jur_RazSoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Jur_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_Jur_RazSoc)
End Sub

Private Sub txt_Jur_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Jur_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()@#$%&/º*+")
   End If
End Sub

Private Sub txt_Nat_ApePat_GotFocus()
   Call gs_SelecTodo(txt_Nat_ApePat)
End Sub

Private Sub txt_Nat_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ '")
   End If
End Sub

Private Sub txt_Nat_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_Nat_ApeMat)
End Sub

Private Sub txt_Nat_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ '")
   End If
End Sub

Private Sub txt_Nat_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nat_Nombre)
End Sub

Private Sub txt_Nat_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Nat_CodSex)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "-_ '")
   End If
End Sub

Private Sub txt_Nat_CygNDo_GotFocus()
   Call gs_SelecTodo(txt_Nat_CygNDo)
End Sub

Private Sub txt_Nat_CygNDo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_CygApp)
   Else
      If cmb_Nat_CygTdo.ListIndex > -1 Then
         Select Case cmb_Nat_CygTdo.ItemData(cmb_Nat_CygTdo.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_Nat_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_Nat_NumDoc)
End Sub

Private Sub txt_Nat_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_ApePat)
   Else
      If cmb_Nat_TipDoc.ListIndex > -1 Then
         Select Case cmb_Nat_TipDoc.ItemData(cmb_Nat_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Numero_GotFocus()
   Call gs_SelecTodo(txt_Numero)
End Sub

Private Sub txt_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Interi_GotFocus()
   Call gs_SelecTodo(txt_Interi)
End Sub

Private Sub txt_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipZon_Click
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Proyec_GotFocus()
   Call gs_SelecTodo(txt_Proyec)
End Sub

Private Sub txt_Proyec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipPro)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,()(/#@$%&/+*º")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_InmPry)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub cmb_TipPro_Click()
   If cmb_TipPro.ListIndex > -1 Then
      If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 1 Then
         Call fs_Limpia_Nat
         
         pnl_Pro_PerNat.Visible = True
         pnl_Pro_PerNat.Enabled = True
         pnl_Pro_PerJur.Visible = False
         
         Call gs_SetFocus(cmb_Nat_TipDoc)
      Else
         Call fs_Limpia_Jur
         
         pnl_Pro_PerNat.Visible = False
         pnl_Pro_PerJur.Visible = True
         pnl_Pro_PerJur.Enabled = True
         
         Call gs_SetFocus(cmb_Jur_TipDoc)
      End If
   End If
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("Al salir de esta manera perderá la información ingresada. ¿Está seguro de salir de la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_CodDpt     As Integer
   Dim r_int_CodPrv     As Integer
   Dim r_int_CodDst     As Integer
   Dim r_int_CodZon     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " Ingreso de Solicitud de Crédito"
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   Call cmb_InmIde_Click
   
   'Cargando Información del Arreglo
   If modatecli_g_int_DatInmTit = 2 Then
      Call gs_BuscarCombo_Item(cmb_InmIde, modatecli_g_arr_DatInm(1).DatInm_InmIde)
   
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 1 Then
         Call gs_BuscarCombo_Item(cmb_TipInm, modatecli_g_arr_DatInm(1).DatInm_TipInm)
         Call gs_BuscarCombo_Item(cmb_UsoInm, modatecli_g_arr_DatInm(1).DatInm_UsoInm)
         Call gs_BuscarCombo_Item(cmb_TipVia, modatecli_g_arr_DatInm(1).DatInm_TipVia)
         Call gs_BuscarCombo_Item(cmb_TipZon, modatecli_g_arr_DatInm(1).DatInm_TipZon)
         Call gs_BuscarCombo_Item(cmb_DptDir, modatecli_g_arr_DatInm(1).DatInm_DptDir)
         Call gs_BuscarCombo_Item(cmb_PrvDir, modatecli_g_arr_DatInm(1).DatInm_PrvDir)
         Call gs_BuscarCombo_Item(cmb_DstDir, modatecli_g_arr_DatInm(1).DatInm_DstDir)
         Call gs_BuscarCombo_Item(cmb_TipPro, modatecli_g_arr_DatInm(1).DatInm_TipPro)
         
         txt_NomVia.Text = modatecli_g_arr_DatInm(1).DatInm_NomVia
         txt_NomZon.Text = modatecli_g_arr_DatInm(1).DatInm_NomZon
         txt_Numero.Text = modatecli_g_arr_DatInm(1).DatInm_Numero
         txt_Interi.Text = modatecli_g_arr_DatInm(1).DatInm_Interi
         txt_Refere.Text = modatecli_g_arr_DatInm(1).DatInm_Refere
         
         'Pendiente lo de Proyectos (Código)
         Call gs_BuscarCombo_Item(cmb_InmPry, modatecli_g_arr_DatInm(1).DatInm_InmPry)
         
         If modatecli_g_arr_DatInm(1).DatInm_InmPry = 1 Then
            cmb_Proyec.ListIndex = gf_Busca_Arregl(l_arr_Proyec, modatecli_g_arr_DatInm(1).DatInm_CodPry) - 1
            txt_Proyec.Text = modatecli_g_arr_DatInm(1).DatInm_NomPry
         End If
         
         If modatecli_g_arr_DatInm(1).DatInm_TipPro = 1 Then
            'Persona Natural
               
            Call gs_BuscarCombo_Item(cmb_Nat_TipDoc, modatecli_g_arr_DatInm(1).DatInm_Nat_TipDoc)
            Call gs_BuscarCombo_Item(cmb_Nat_CodSex, modatecli_g_arr_DatInm(1).DatInm_Nat_CodSex)
            Call gs_BuscarCombo_Item(cmb_Nat_EstCiv, modatecli_g_arr_DatInm(1).DatInm_Nat_EstCiv)
            
            txt_Nat_NumDoc.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_NumDoc
            txt_Nat_ApePat.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_ApePat
            txt_Nat_ApeMat.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_ApeMat
            txt_Nat_Nombre.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_Nombre
            txt_Nat_Telef1.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_Telef1
            txt_Nat_Telef2.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_Telef2
            
            If modatecli_g_arr_DatInm(1).DatInm_Nat_EstCiv = 2 Then
               Call gs_BuscarCombo_Item(cmb_Nat_CygTdo, modatecli_g_arr_DatInm(1).DatInm_Nat_CygTDo)
               
               txt_Nat_CygNDo.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_CygNDo
               txt_Nat_CygApp.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_CygApp
               txt_Nat_CygApm.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_CygApm
               txt_Nat_CygNom.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_CygNom
               
               txt_Nat_CygTl1.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_CygTl1
               txt_Nat_CygTl2.Text = modatecli_g_arr_DatInm(1).DatInm_Nat_CygTl2
            End If
         Else
            'Persona Juridica
            Call gs_BuscarCombo_Item(cmb_Jur_TipDoc, modatecli_g_arr_DatInm(1).DatInm_Jur_TipDoc)
            
            txt_Jur_NumDoc.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_NumDoc
            txt_Jur_RazSoc.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_RazSoc
            
            Call gs_BuscarCombo_Item(cmb_Jur_TipVia, modatecli_g_arr_DatInm(1).DatInm_Jur_TipVia)
            Call gs_BuscarCombo_Item(cmb_Jur_TipZon, modatecli_g_arr_DatInm(1).DatInm_Jur_TipZon)
            Call gs_BuscarCombo_Item(cmb_Jur_DptDir, modatecli_g_arr_DatInm(1).DatInm_Jur_DptDir)
            Call gs_BuscarCombo_Item(cmb_Jur_PrvDir, modatecli_g_arr_DatInm(1).DatInm_Jur_PrvDir)
            Call gs_BuscarCombo_Item(cmb_Jur_DstDir, modatecli_g_arr_DatInm(1).DatInm_Jur_DstDir)
            
            txt_Jur_NomVia.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_NomVia
            txt_Jur_Numero.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_Numero
            txt_Jur_Interi.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_Interi
            txt_Jur_Refere.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_Refere
            txt_Jur_Telef1.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_Telef1
            txt_Jur_Telef2.Text = modatecli_g_arr_DatInm(1).DatInm_Jur_Telef2
         End If
      Else
         r_int_CodDpt = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo1, 1, 2))
         r_int_CodPrv = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo1, 3, 2))
         r_int_CodDst = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo1, 5, 2))
         r_int_CodZon = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo1, 7, 2))
         
         Call gs_BuscarCombo_Item(cmb_DptZo1, r_int_CodDpt)

         Call moddat_gs_Carga_Provin(cmb_PrvZo1, Format(r_int_CodDpt, "00"))
         Call gs_BuscarCombo_Item(cmb_PrvZo1, r_int_CodPrv)

         Call moddat_gs_Carga_Distri(cmb_DstZo1, Format(r_int_CodDpt, "00"), Format(r_int_CodPrv, "00"))
         Call gs_BuscarCombo_Item(cmb_DstZo1, r_int_CodDst)
         
         Call moddat_gs_Carga_DstZon(cmb_NomZo1, Format(r_int_CodDpt, "00"), Format(r_int_CodPrv, "00"), Format(r_int_CodDst, "00"))
         Call gs_BuscarCombo_Item(cmb_NomZo1, r_int_CodZon)
         
         If modatecli_g_arr_DatInm(1).DatInm_ZonPo2 <> "00000000" Then
            r_int_CodDpt = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo2, 1, 2))
            r_int_CodPrv = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo2, 3, 2))
            r_int_CodDst = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo2, 5, 2))
            r_int_CodZon = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo2, 7, 2))
            
            Call gs_BuscarCombo_Item(cmb_DptZo2, r_int_CodDpt)
   
            Call moddat_gs_Carga_Provin(cmb_PrvZo2, Format(r_int_CodDpt, "00"))
            Call gs_BuscarCombo_Item(cmb_PrvZo2, r_int_CodPrv)
   
            Call moddat_gs_Carga_Distri(cmb_DstZo2, Format(r_int_CodDpt, "00"), Format(r_int_CodPrv, "00"))
            Call gs_BuscarCombo_Item(cmb_DstZo2, r_int_CodDst)
            
            Call moddat_gs_Carga_DstZon(cmb_NomZo2, Format(r_int_CodDpt, "00"), Format(r_int_CodPrv, "00"), Format(r_int_CodDst, "00"))
            Call gs_BuscarCombo_Item(cmb_NomZo2, r_int_CodZon)
         End If
      
         If modatecli_g_arr_DatInm(1).DatInm_ZonPo3 <> "00000000" Then
            r_int_CodDpt = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo3, 1, 2))
            r_int_CodPrv = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo3, 3, 2))
            r_int_CodDst = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo3, 5, 2))
            r_int_CodZon = CInt(Mid(modatecli_g_arr_DatInm(1).DatInm_ZonPo3, 7, 2))
            
            Call gs_BuscarCombo_Item(cmb_DptZo3, r_int_CodDpt)
   
            Call moddat_gs_Carga_Provin(cmb_PrvZo3, Format(r_int_CodDpt, "00"))
            Call gs_BuscarCombo_Item(cmb_PrvZo3, r_int_CodPrv)
   
            Call moddat_gs_Carga_Distri(cmb_DstZo3, Format(r_int_CodDpt, "00"), Format(r_int_CodPrv, "00"))
            Call gs_BuscarCombo_Item(cmb_DstZo3, r_int_CodDst)
            
            Call moddat_gs_Carga_DstZon(cmb_NomZo3, Format(r_int_CodDpt, "00"), Format(r_int_CodPrv, "00"), Format(r_int_CodDst, "00"))
            Call gs_BuscarCombo_Item(cmb_NomZo3, r_int_CodZon)
         End If
         
         ipp_NumDor.Value = modatecli_g_arr_DatInm(1).DatInm_NumDor
         ipp_NumBan.Value = modatecli_g_arr_DatInm(1).DatInm_NumBan
         ipp_NumEst.Value = modatecli_g_arr_DatInm(1).DatInm_NumEst
         ipp_AreCon.Value = modatecli_g_arr_DatInm(1).DatInm_AreCon
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "217")
   Call moddat_gs_Carga_LisIte_Combo(cmb_UsoInm, 1, "218")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPro, 1, "221")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Jur_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Jur_TipZon, 1, "202")
   
   Call moddat_gs_Carga_TipDocIde(cmb_Nat_TipDoc, 1)
   Call moddat_gs_Carga_TipDocIde(cmb_Nat_CygTdo, 1)
   Call moddat_gs_Carga_TipDocIde(cmb_Jur_TipDoc, 2)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Nat_EstCiv, 1, "205")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Nat_CodSex, 1, "207")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_InmPry, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_InmIde, 1, "214")
      
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_Depart(cmb_Jur_DptDir)
   
   Call moddat_gs_Carga_Depart(cmb_DptZo1)
   Call moddat_gs_Carga_Depart(cmb_DptZo2)
   Call moddat_gs_Carga_Depart(cmb_DptZo3)
   
   Call moddat_gs_Carga_Proyec(cmb_Proyec, l_arr_Proyec)
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

Private Sub cmb_Jur_DptDir_Change()
   l_str_Jur_DptDir = cmb_Jur_DptDir.Text
End Sub

Private Sub cmb_Jur_DptDir_Click()
   If cmb_Jur_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Jur_PrvDir.Clear
         cmb_Jur_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Jur_PrvDir, Format(cmb_Jur_DptDir.ItemData(cmb_Jur_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Jur_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_Jur_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Jur_DptDir = cmb_Jur_DptDir.Text
End Sub

Private Sub cmb_Jur_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Jur_DptDir, l_str_Jur_DptDir)
      l_int_FlgCmb = True
      
      cmb_Jur_PrvDir.Clear
      cmb_Jur_DstDir.Clear
      If cmb_Jur_DptDir.ListIndex > -1 Then
         l_str_Jur_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_Jur_PrvDir, Format(cmb_Jur_DptDir.ItemData(cmb_Jur_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Jur_PrvDir)
   End If
End Sub

Private Sub cmb_Jur_PrvDir_Change()
   l_str_Jur_PrvDir = cmb_Jur_PrvDir.Text
End Sub

Private Sub cmb_Jur_PrvDir_Click()
   If cmb_Jur_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_Jur_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Jur_DstDir, Format(cmb_Jur_DptDir.ItemData(cmb_Jur_DptDir.ListIndex), "00"), Format(cmb_Jur_PrvDir.ItemData(cmb_Jur_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_Jur_DstDir)
      End If
   End If
End Sub

Private Sub cmb_Jur_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Jur_PrvDir = cmb_Jur_PrvDir.Text
End Sub

Private Sub cmb_Jur_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Jur_PrvDir, l_str_Jur_PrvDir)
      l_int_FlgCmb = True
      
      cmb_Jur_DstDir.Clear
      If cmb_Jur_PrvDir.ListIndex > -1 Then
         l_str_Jur_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_Jur_DstDir, Format(cmb_Jur_DptDir.ItemData(cmb_Jur_DptDir.ListIndex), "00"), Format(cmb_Jur_PrvDir.ItemData(cmb_Jur_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_Jur_DstDir)
   End If
End Sub

Private Sub cmb_Jur_DstDir_Change()
   l_str_Jur_DstDir = cmb_Jur_DstDir.Text
End Sub

Private Sub cmb_Jur_DstDir_Click()
   If cmb_Jur_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Jur_Refere)
      End If
   End If
End Sub

Private Sub cmb_Jur_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_Jur_DstDir = cmb_Jur_DstDir.Text
End Sub

Private Sub cmb_Jur_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Jur_DstDir, l_str_Jur_DstDir)
      l_int_FlgCmb = True
      
      If cmb_Jur_DstDir.ListIndex > -1 Then
         l_str_Jur_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Jur_Refere)
   End If
End Sub

Private Sub fs_Limpia()
   cmb_InmIde.ListIndex = -1
   
   cmb_DptZo1.ListIndex = -1
   cmb_PrvZo1.Clear
   cmb_DstZo1.Clear
   cmb_NomZo1.Clear
   
   cmb_DptZo2.ListIndex = -1
   cmb_PrvZo2.Clear
   cmb_DstZo2.Clear
   cmb_NomZo2.Clear
   
   cmb_DptZo3.ListIndex = -1
   cmb_PrvZo3.Clear
   cmb_DstZo3.Clear
   cmb_NomZo3.Clear
   
   ipp_NumDor.Value = 0
   ipp_NumBan.Value = 0
   ipp_NumEst.Value = 0
   ipp_AreCon.Value = 0
   
   cmb_TipInm.ListIndex = -1
   cmb_UsoInm.ListIndex = -1
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
   cmb_InmPry.ListIndex = -1
   cmb_Proyec.ListIndex = -1
   txt_Proyec.Text = ""
   cmb_Proyec.Enabled = False
   txt_Proyec.Enabled = False
   cmb_TipPro.ListIndex = -1
   
   Call fs_Limpia_Nat
   Call fs_Limpia_Jur
   
   pnl_Pro_PerNat.Visible = True
   pnl_Pro_PerNat.Enabled = False
   
   pnl_Pro_PerJur.Visible = False
   pnl_Pro_PerJur.Enabled = False
End Sub

Private Sub fs_Limpia_Nat()
   cmb_Nat_TipDoc.ListIndex = -1
   txt_Nat_NumDoc.Text = ""
   txt_Nat_ApePat.Text = ""
   txt_Nat_ApeMat.Text = ""
   txt_Nat_Nombre.Text = ""
   txt_Nat_Telef1.Text = ""
   txt_Nat_Telef2.Text = ""
   cmb_Nat_CodSex.ListIndex = -1
   cmb_Nat_EstCiv.ListIndex = -1
   
   cmb_Nat_CygTdo.ListIndex = -1
   txt_Nat_CygNDo.Text = ""
   txt_Nat_CygApp.Text = ""
   txt_Nat_CygApm.Text = ""
   txt_Nat_CygNom.Text = ""
   txt_Nat_CygTl1.Text = ""
   txt_Nat_CygTl2.Text = ""
   
   cmb_Nat_CygTdo.Enabled = False
   txt_Nat_CygNDo.Enabled = False
   txt_Nat_CygApp.Enabled = False
   txt_Nat_CygApm.Enabled = False
   txt_Nat_CygNom.Enabled = False
End Sub

Private Sub fs_Limpia_Jur()
   cmb_Jur_TipDoc.ListIndex = -1
   txt_Jur_NumDoc.Text = ""
   txt_Jur_RazSoc.Text = ""
   cmb_Jur_TipVia.ListIndex = -1
   txt_Jur_NomVia.Text = ""
   txt_Jur_Numero.Text = ""
   txt_Jur_Interi.Text = ""
   cmb_Jur_TipZon.ListIndex = -1
   txt_Jur_NomZon.Text = ""
   cmb_Jur_DptDir.ListIndex = -1
   cmb_Jur_PrvDir.Clear
   cmb_Jur_DstDir.Clear
   txt_Jur_Refere.Text = ""
   txt_Jur_Telef1.Text = ""
   txt_Jur_Telef2.Text = ""
End Sub

Private Sub cmb_Jur_TipVia_Click()
   Call gs_SetFocus(txt_Jur_NomVia)
End Sub

Private Sub cmb_Jur_TipVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Jur_TipVia_Click
   End If
End Sub

Private Sub txt_Jur_NomVia_GotFocus()
   Call gs_SelecTodo(txt_Jur_NomVia)
End Sub

Private Sub txt_Jur_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Jur_Numero)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Jur_Numero_GotFocus()
   Call gs_SelecTodo(txt_Jur_Numero)
End Sub

Private Sub txt_Jur_Numero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Jur_Interi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Jur_Interi_GotFocus()
   Call gs_SelecTodo(txt_Jur_Interi)
End Sub

Private Sub txt_Jur_Interi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Jur_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub cmb_Jur_TipZon_Click()
   Call gs_SetFocus(txt_Jur_NomZon)
End Sub

Private Sub cmb_Jur_TipZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Jur_TipZon_Click
   End If
End Sub

Private Sub txt_Jur_NomZon_GotFocus()
   Call gs_SelecTodo(txt_Jur_NomZon)
End Sub

Private Sub txt_Jur_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Jur_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Jur_Refere_GotFocus()
   Call gs_SelecTodo(txt_Jur_Refere)
End Sub

Private Sub txt_Jur_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Jur_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()º#@$%=?¿+")
   End If
End Sub

Private Sub txt_Jur_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Jur_Telef1)
End Sub

Private Sub txt_Jur_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Jur_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Jur_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Jur_Telef2)
End Sub

Private Sub txt_Jur_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nat_CygApp_GotFocus()
   Call gs_SelecTodo(txt_Nat_CygApp)
End Sub

Private Sub txt_Nat_CygApp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_CygApm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ '")
   End If
End Sub

Private Sub txt_Nat_CygApm_GotFocus()
   Call gs_SelecTodo(txt_Nat_CygApm)
End Sub

Private Sub txt_Nat_CygApm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_CygNom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ '")
   End If
End Sub

Private Sub txt_Nat_CygNom_GotFocus()
   Call gs_SelecTodo(txt_Nat_CygNom)
End Sub

Private Sub txt_Nat_CygNom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_CygTl1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ '")
   End If
End Sub

Private Sub txt_Nat_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Nat_Telef1)
End Sub

Private Sub txt_Nat_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nat_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Nat_Telef2)
End Sub

Private Sub txt_Nat_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Nat_CygTdo.Enabled Then
         Call gs_SetFocus(cmb_Nat_CygTdo)
      Else
         Call gs_SetFocus(cmd_Acepta)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nat_CygTl1_GotFocus()
   Call gs_SelecTodo(txt_Nat_CygTl1)
End Sub

Private Sub txt_Nat_CygTl1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nat_CygTl2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nat_CygTl2_GotFocus()
   Call gs_SelecTodo(txt_Nat_CygTl2)
End Sub

Private Sub txt_Nat_CygTl2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Acepta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_DptZo1_Change()
   l_str_DptZo1 = cmb_DptZo1.Text
End Sub

Private Sub cmb_DptZo1_Click()
   If cmb_DptZo1.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvZo1.Clear
         cmb_DstZo1.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvZo1, Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvZo1)
      End If
   End If
End Sub

Private Sub cmb_DptZo1_GotFocus()
   l_int_FlgCmb = True
   l_str_DptZo1 = cmb_DptZo1.Text
End Sub

Private Sub cmb_DptZo1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptZo1, l_str_DptZo1)
      l_int_FlgCmb = True
      
      cmb_PrvZo1.Clear
      cmb_DstZo1.Clear
      
      If cmb_DptZo1.ListIndex > -1 Then
         l_str_DptZo1 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvZo1, Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvZo1)
   End If
End Sub

Private Sub cmb_PrvZo1_Change()
   l_str_PrvZo1 = cmb_PrvZo1.Text
End Sub

Private Sub cmb_PrvZo1_Click()
   If cmb_PrvZo1.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstZo1.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstZo1, Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00"), Format(cmb_PrvZo1.ItemData(cmb_PrvZo1.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstZo1)
      End If
   End If
End Sub

Private Sub cmb_PrvZo1_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvZo1 = cmb_PrvZo1.Text
End Sub

Private Sub cmb_PrvZo1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvZo1, l_str_PrvZo1)
      l_int_FlgCmb = True
      
      cmb_DstZo1.Clear
      If cmb_PrvZo1.ListIndex > -1 Then
         l_str_DstZo1 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstZo1, Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00"), Format(cmb_PrvZo1.ItemData(cmb_PrvZo1.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstZo1)
   End If
End Sub

Private Sub cmb_DstZo1_Change()
   l_str_DstZo1 = cmb_DstZo1.Text
End Sub

Private Sub cmb_DstZo1_Click()
   If cmb_DstZo1.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_NomZo1.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_NomZo1, Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00"), Format(cmb_PrvZo1.ItemData(cmb_PrvZo1.ListIndex), "00"), Format(cmb_DstZo1.ItemData(cmb_DstZo1.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_NomZo1)
      End If
   End If
End Sub

Private Sub cmb_DstZo1_GotFocus()
   l_int_FlgCmb = True
   l_str_DstZo1 = cmb_DstZo1.Text
End Sub

Private Sub cmb_DstZo1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstZo1, l_str_DstZo1)
      l_int_FlgCmb = True
      
      cmb_NomZo1.Clear
      If cmb_DstZo1.ListIndex > -1 Then
         l_str_NomZo1 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_NomZo1, Format(cmb_DptZo1.ItemData(cmb_DptZo1.ListIndex), "00"), Format(cmb_PrvZo1.ItemData(cmb_PrvZo1.ListIndex), "00"), Format(cmb_DstZo1.ItemData(cmb_DstZo1.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_NomZo1)
   End If
End Sub

Private Sub cmb_NomZo1_Change()
   l_str_NomZo1 = cmb_NomZo1.Text
End Sub

Private Sub cmb_NomZo1_Click()
   If cmb_NomZo1.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_DptZo2)
      End If
   End If
End Sub

Private Sub cmb_NomZo1_GotFocus()
   l_int_FlgCmb = True
   l_str_NomZo1 = cmb_NomZo1.Text
End Sub

Private Sub cmb_NomZo1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_NomZo1, l_str_NomZo1)
      l_int_FlgCmb = True
      
      If cmb_NomZo1.ListIndex > -1 Then
         l_str_NomZo1 = ""
      End If
      
      Call gs_SetFocus(cmb_DptZo2)
   End If
End Sub

Private Sub cmb_DptZo2_Change()
   l_str_DptZo2 = cmb_DptZo2.Text
End Sub

Private Sub cmb_DptZo2_Click()
   If cmb_DptZo2.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvZo2.Clear
         cmb_DstZo2.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvZo2, Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvZo2)
      End If
   End If
End Sub

Private Sub cmb_DptZo2_GotFocus()
   l_int_FlgCmb = True
   l_str_DptZo2 = cmb_DptZo2.Text
End Sub

Private Sub cmb_DptZo2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptZo2, l_str_DptZo2)
      l_int_FlgCmb = True
      
      cmb_PrvZo2.Clear
      cmb_DstZo2.Clear
      
      If cmb_DptZo2.ListIndex > -1 Then
         l_str_DptZo2 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvZo2, Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvZo2)
   End If
End Sub

Private Sub cmb_PrvZo2_Change()
   l_str_PrvZo2 = cmb_PrvZo2.Text
End Sub

Private Sub cmb_PrvZo2_Click()
   If cmb_PrvZo2.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstZo2.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstZo2, Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00"), Format(cmb_PrvZo2.ItemData(cmb_PrvZo2.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstZo2)
      End If
   End If
End Sub

Private Sub cmb_PrvZo2_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvZo2 = cmb_PrvZo2.Text
End Sub

Private Sub cmb_PrvZo2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvZo2, l_str_PrvZo2)
      l_int_FlgCmb = True
      
      cmb_DstZo2.Clear
      If cmb_PrvZo2.ListIndex > -1 Then
         l_str_DstZo2 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstZo2, Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00"), Format(cmb_PrvZo2.ItemData(cmb_PrvZo2.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstZo2)
   End If
End Sub

Private Sub cmb_DstZo2_Change()
   l_str_DstZo2 = cmb_DstZo2.Text
End Sub

Private Sub cmb_DstZo2_Click()
   If cmb_DstZo2.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_NomZo2.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_NomZo2, Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00"), Format(cmb_PrvZo2.ItemData(cmb_PrvZo2.ListIndex), "00"), Format(cmb_DstZo2.ItemData(cmb_DstZo2.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_NomZo2)
      End If
   End If
End Sub

Private Sub cmb_DstZo2_GotFocus()
   l_int_FlgCmb = True
   l_str_DstZo2 = cmb_DstZo2.Text
End Sub

Private Sub cmb_DstZo2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstZo2, l_str_DstZo2)
      l_int_FlgCmb = True
      
      cmb_NomZo2.Clear
      If cmb_DstZo2.ListIndex > -1 Then
         l_str_NomZo2 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_NomZo2, Format(cmb_DptZo2.ItemData(cmb_DptZo2.ListIndex), "00"), Format(cmb_PrvZo2.ItemData(cmb_PrvZo2.ListIndex), "00"), Format(cmb_DstZo2.ItemData(cmb_DstZo2.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_NomZo2)
   End If
End Sub

Private Sub cmb_NomZo2_Change()
   l_str_NomZo2 = cmb_NomZo2.Text
End Sub

Private Sub cmb_NomZo2_Click()
   If cmb_NomZo2.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_DptZo3)
      End If
   End If
End Sub

Private Sub cmb_NomZo2_GotFocus()
   l_int_FlgCmb = True
   l_str_NomZo2 = cmb_NomZo2.Text
End Sub

Private Sub cmb_NomZo2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_NomZo2, l_str_NomZo2)
      l_int_FlgCmb = True
      
      If cmb_NomZo2.ListIndex > -1 Then
         l_str_NomZo2 = ""
      End If
      
      Call gs_SetFocus(cmb_DptZo3)
   End If
End Sub

Private Sub cmb_DptZo3_Change()
   l_str_DptZo3 = cmb_DptZo3.Text
End Sub

Private Sub cmb_DptZo3_Click()
   If cmb_DptZo3.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvZo3.Clear
         cmb_DstZo3.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvZo3, Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvZo3)
      End If
   End If
End Sub

Private Sub cmb_DptZo3_GotFocus()
   l_int_FlgCmb = True
   l_str_DptZo3 = cmb_DptZo3.Text
End Sub

Private Sub cmb_DptZo3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptZo3, l_str_DptZo3)
      l_int_FlgCmb = True
      
      cmb_PrvZo3.Clear
      cmb_DstZo3.Clear
      
      If cmb_DptZo3.ListIndex > -1 Then
         l_str_DptZo3 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvZo3, Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvZo3)
   End If
End Sub

Private Sub cmb_PrvZo3_Change()
   l_str_PrvZo3 = cmb_PrvZo3.Text
End Sub

Private Sub cmb_PrvZo3_Click()
   If cmb_PrvZo3.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstZo3.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstZo3, Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00"), Format(cmb_PrvZo3.ItemData(cmb_PrvZo3.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstZo3)
      End If
   End If
End Sub

Private Sub cmb_PrvZo3_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvZo3 = cmb_PrvZo3.Text
End Sub

Private Sub cmb_PrvZo3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvZo3, l_str_PrvZo3)
      l_int_FlgCmb = True
      
      cmb_DstZo3.Clear
      If cmb_PrvZo3.ListIndex > -1 Then
         l_str_DstZo3 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstZo3, Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00"), Format(cmb_PrvZo3.ItemData(cmb_PrvZo3.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstZo3)
   End If
End Sub

Private Sub cmb_DstZo3_Change()
   l_str_DstZo3 = cmb_DstZo3.Text
End Sub

Private Sub cmb_DstZo3_Click()
   If cmb_DstZo3.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_NomZo3.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_NomZo3, Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00"), Format(cmb_PrvZo3.ItemData(cmb_PrvZo3.ListIndex), "00"), Format(cmb_DstZo3.ItemData(cmb_DstZo3.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_NomZo3)
      End If
   End If
End Sub

Private Sub cmb_DstZo3_GotFocus()
   l_int_FlgCmb = True
   l_str_DstZo3 = cmb_DstZo3.Text
End Sub

Private Sub cmb_DstZo3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstZo3, l_str_DstZo3)
      l_int_FlgCmb = True
      
      cmb_NomZo3.Clear
      If cmb_DstZo3.ListIndex > -1 Then
         l_str_NomZo3 = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_DstZon(cmb_NomZo3, Format(cmb_DptZo3.ItemData(cmb_DptZo3.ListIndex), "00"), Format(cmb_PrvZo3.ItemData(cmb_PrvZo3.ListIndex), "00"), Format(cmb_DstZo3.ItemData(cmb_DstZo3.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_NomZo3)
   End If
End Sub

Private Sub cmb_NomZo3_Change()
   l_str_NomZo3 = cmb_NomZo3.Text
End Sub

Private Sub cmb_NomZo3_Click()
   If cmb_NomZo3.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(ipp_AreCon)
      End If
   End If
End Sub

Private Sub cmb_NomZo3_GotFocus()
   l_int_FlgCmb = True
   l_str_NomZo3 = cmb_NomZo3.Text
End Sub

Private Sub cmb_NomZo3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_NomZo3, l_str_NomZo3)
      l_int_FlgCmb = True
      
      If cmb_NomZo3.ListIndex > -1 Then
         l_str_NomZo3 = ""
      End If
      
      Call gs_SetFocus(ipp_AreCon)
   End If
End Sub



