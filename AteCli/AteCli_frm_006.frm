VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_IngSol_08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   9375
   ClientLeft      =   1500
   ClientTop       =   720
   ClientWidth     =   11535
   Icon            =   "AteCli_frm_006.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11595
      _Version        =   65536
      _ExtentX        =   20452
      _ExtentY        =   16536
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
         Height          =   3135
         Left            =   30
         TabIndex        =   54
         Top             =   5370
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   5530
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2775
            Left            =   60
            TabIndex        =   13
            Top             =   330
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   9780
            TabIndex        =   55
            Top             =   60
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Recepcionado"
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
            Left            =   90
            TabIndex        =   56
            Top             =   60
            Width           =   9705
            _Version        =   65536
            _ExtentX        =   17119
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
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
         Height          =   765
         Left            =   30
         TabIndex        =   53
         Top             =   8550
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10770
            Picture         =   "AteCli_frm_006.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   675
            Left            =   10080
            Picture         =   "AteCli_frm_006.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   585
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   11505
         _Version        =   65536
         _ExtentX        =   20294
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
            TabIndex        =   46
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
            TabIndex        =   47
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Datos del Crédito"
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
            Picture         =   "AteCli_frm_006.frx":0758
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4665
         Left            =   30
         TabIndex        =   18
         Top             =   660
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   8229
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
         Begin VB.ComboBox cmb_CuoAno 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1710
            Width           =   735
         End
         Begin VB.TextBox txt_Observ 
            Height          =   885
            Left            =   2070
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Text            =   "AteCli_frm_006.frx":0A62
            Top             =   3720
            Width           =   9345
         End
         Begin Threed.SSPanel pnl_ComVta_NSoles 
            Height          =   315
            Left            =   5820
            TabIndex        =   32
            Top             =   390
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_SegDes 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2400
            Width           =   9345
         End
         Begin VB.ComboBox cmb_DiaVct 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   3060
            Width           =   1635
         End
         Begin VB.ComboBox cmb_EjeVta 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   3390
            Width           =   9345
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta_Dolare 
            Height          =   315
            Left            =   2070
            TabIndex        =   1
            Top             =   390
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_ApoPro_Dolare 
            Height          =   315
            Left            =   2070
            TabIndex        =   2
            Top             =   720
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
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
            MinValue        =   "0"
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   2070
            TabIndex        =   4
            Top             =   1380
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
            MaxValue        =   "70"
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
         Begin EditLib.fpLongInteger ipp_PlaMes 
            Height          =   315
            Left            =   3330
            TabIndex        =   5
            Top             =   1380
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
            MaxValue        =   "12"
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
         Begin EditLib.fpLongInteger ipp_PerGra 
            Height          =   315
            Left            =   2070
            TabIndex        =   7
            Top             =   2070
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
            MaxValue        =   "70"
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
         Begin EditLib.fpDoubleSingle ipp_CuoMen 
            Height          =   315
            Left            =   2070
            TabIndex        =   9
            Top             =   2730
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
            MinValue        =   "0"
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
         Begin EditLib.fpDoubleSingle ipp_MonSol_Dolare 
            Height          =   315
            Left            =   2070
            TabIndex        =   3
            Top             =   1050
            Width           =   1995
            _Version        =   196608
            _ExtentX        =   3519
            _ExtentY        =   556
            Enabled         =   0   'False
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
            MinValue        =   "0"
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
         Begin Threed.SSPanel pnl_ComVta_MonPre 
            Height          =   315
            Left            =   9330
            TabIndex        =   34
            Top             =   390
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_MonSol_NSoles 
            Height          =   315
            Left            =   5820
            TabIndex        =   36
            Top             =   1050
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_MonSol_MonPre 
            Height          =   315
            Left            =   9330
            TabIndex        =   38
            Top             =   1050
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ApoPro_NSoles 
            Height          =   315
            Left            =   5820
            TabIndex        =   40
            Top             =   720
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_ApoPro_MonPre 
            Height          =   315
            Left            =   9330
            TabIndex        =   42
            Top             =   720
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
         Begin Threed.SSPanel pnl_PlaTot 
            Height          =   315
            Left            =   5820
            TabIndex        =   49
            Top             =   1380
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
         Begin VB.Label Label21 
            Caption         =   "Cuotas Extraordinarias:"
            Height          =   315
            Left            =   90
            TabIndex        =   52
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Meses"
            Height          =   285
            Left            =   6360
            TabIndex        =   51
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label9 
            Caption         =   "Total:"
            Height          =   315
            Left            =   5040
            TabIndex        =   50
            Top             =   1380
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   3720
            Width           =   1605
         End
         Begin VB.Label Label19 
            Caption         =   "* Día que el cliente desearía pagar"
            Height          =   285
            Left            =   3780
            TabIndex        =   45
            Top             =   3060
            Width           =   2955
         End
         Begin VB.Label Label18 
            Caption         =   "* Importe que el cliente desearía pagar"
            Height          =   285
            Left            =   3780
            TabIndex        =   44
            Top             =   2760
            Width           =   2955
         End
         Begin VB.Label Label17 
            Caption         =   "En MP:"
            Height          =   315
            Left            =   8490
            TabIndex        =   43
            Top             =   750
            Width           =   675
         End
         Begin VB.Label Label16 
            Caption         =   "En S/::"
            Height          =   315
            Left            =   5040
            TabIndex        =   41
            Top             =   750
            Width           =   645
         End
         Begin VB.Label Label13 
            Caption         =   "En MP:"
            Height          =   315
            Left            =   8490
            TabIndex        =   39
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label12 
            Caption         =   "En S/.:"
            Height          =   315
            Left            =   5040
            TabIndex        =   37
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label Label11 
            Caption         =   "En MP:"
            Height          =   315
            Left            =   8490
            TabIndex        =   35
            Top             =   420
            Width           =   765
         End
         Begin VB.Label Label10 
            Caption         =   "En S/.:"
            Height          =   315
            Left            =   5040
            TabIndex        =   33
            Top             =   420
            Width           =   645
         End
         Begin VB.Label Label15 
            Caption         =   "Monto Solicitado US$:"
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Moneda del Crédito:"
            Height          =   315
            Left            =   90
            TabIndex        =   30
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Compra-Venta US$:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "Aporte Propio US$:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label31 
            Caption         =   "Meses"
            Height          =   285
            Left            =   4140
            TabIndex        =   27
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label30 
            Caption         =   "Años"
            Height          =   285
            Left            =   2850
            TabIndex        =   26
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo:"
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   1380
            Width           =   1665
         End
         Begin VB.Label Label3 
            Caption         =   "Seguro de Préstamo:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   2400
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "Período de Gracia:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   2070
            Width           =   1665
         End
         Begin VB.Label Label5 
            Caption         =   "Meses"
            Height          =   285
            Left            =   2850
            TabIndex        =   22
            Top             =   2130
            Width           =   555
         End
         Begin VB.Label Label6 
            Caption         =   "Días de Vencimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   3060
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Cuota Mensual US$:"
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   2730
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Ejec. de Ventas:"
            Height          =   315
            Left            =   90
            TabIndex        =   19
            Top             =   3390
            Width           =   1905
         End
      End
   End
End
Attribute VB_Name = "frm_IngSol_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_Dol_TipCam    As Double
Dim l_dbl_Pre_TipCam    As Double
Dim l_arr_EjeVta()      As moddat_tpo_Genera

Private Sub cmb_CuoAno_Click()
   Call gs_SetFocus(ipp_PerGra)
End Sub

Private Sub cmb_CuoAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CuoAno_Click
   End If
End Sub

Private Sub cmb_DiaVct_Click()
   Call gs_SetFocus(cmb_EjeVta)
End Sub

Private Sub cmb_DiaVct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DiaVct_Click
   End If
End Sub

Private Sub cmb_EjeVta_Click()
   Call gs_SetFocus(txt_Observ)
End Sub

Private Sub cmb_EjeVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EjeVta_Click
   End If
End Sub

Private Sub cmb_SegDes_Click()
   Call gs_SetFocus(ipp_CuoMen)
End Sub

Private Sub cmb_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegDes_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_ComVta_Dolare)
   
   If cmb_TipMon.ListIndex > -1 Then
      l_dbl_Pre_TipCam = moddat_gf_Obtiene_TipCam(1, cmb_TipMon.ItemData(cmb_TipMon.ListIndex))
      
      If l_dbl_Pre_TipCam = 0 Then
         MsgBox "No se encontró el Tipo de Cambio para esta Moneda.", vbExclamation, modgen_g_str_NomPlt
      End If
      
      Call fs_Calcula(1)
      Call fs_Calcula(2)
      Call fs_Calcula(3)
   End If
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Acepta_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_FlgDoc     As Integer
   Dim r_int_PlaAno     As Integer
   Dim r_int_PlaMes     As Integer
   Dim r_str_Selecc     As String
   
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If l_dbl_Dol_TipCam = 0 Then
      MsgBox "Solicite el Ingreso de Tipo de Cambio al Sistema.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If l_dbl_Pre_TipCam = 0 Then
      MsgBox "Solicite el Ingreso de Tipo de Cambio al Sistema.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If CDbl(ipp_ComVta_Dolare.Text) = 0 Then
      MsgBox "Debe ingresar el Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta_Dolare)
      Exit Sub
   End If

   If CDbl(ipp_ApoPro_Dolare.Text) = 0 Then
      MsgBox "Debe ingresar el Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro_Dolare)
      Exit Sub
   End If

   If CDbl(ipp_ComVta_Dolare.Text) <= CDbl(ipp_ApoPro_Dolare.Text) Then
      MsgBox "El Valor Compra Venta no puede ser menor al Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta_Dolare)
      Exit Sub
   End If

   If CDbl(ipp_MonSol_Dolare.Text) <= 0 Then
      MsgBox "Debe ingresar el Monto Solicitado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MonSol_Dolare)
      Exit Sub
   End If
   
   If (CDbl(ipp_ApoPro_Dolare.Text) / CDbl(ipp_ComVta_Dolare.Text) * 100) < modatecli_g_dbl_Par_PorApo Then
      MsgBox "El Aporte Propio no cumple el mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro_Dolare)
      Exit Sub
   End If
   
   If CDbl(pnl_ComVta_NSoles.Caption) > modatecli_g_dbl_Par_ValViv Then
      MsgBox "El Valor de Compra-Venta no se ajusta a los Parámetros requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta_Dolare)
      Exit Sub
   End If
   
   If Not (CDbl(ipp_MonSol_Dolare.Text) >= modatecli_g_dbl_Par_PreMin And CDbl(ipp_MonSol_Dolare.Text) <= modatecli_g_dbl_Par_PreMax) Then
      MsgBox "El Monto Solicitado no se ajusta a los Parámetros requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta_Dolare)
      Exit Sub
   End If
   
   If CInt(ipp_PlaAno.Value) = 0 Then
      MsgBox "Debe ingresar el Plazo en Años.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If Not (CInt(pnl_PlaTot.Caption) >= modatecli_g_int_Par_PlaMin And CInt(pnl_PlaTot.Caption) <= modatecli_g_int_Par_PlaMax) Then
      MsgBox "El Plazo del Préstamo no se ajusta a los Parámetros requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   r_int_PlaAno = Int(CInt(pnl_PlaTot.Caption) / 12) + moddat_g_int_EdaAno
   r_int_PlaMes = CInt(pnl_PlaTot.Caption) - (r_int_PlaAno * 12) + moddat_g_int_EdaMes
   
   If Not ((r_int_PlaAno < modatecli_g_int_Par_EdaTot) Or (r_int_PlaAno = modatecli_g_int_Par_EdaTot And r_int_PlaMes = 0)) Then
      MsgBox "El Plazo del Préstamo no se ajusta al límite de Edad permitido.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If cmb_CuoAno.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Indicador de Cuotas Extraordinarias.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CuoAno)
      Exit Sub
   End If
   
   If Not (CInt(ipp_PerGra.Text) >= modatecli_g_int_Par_GraMin And CInt(ipp_PerGra.Text) <= modatecli_g_int_Par_GraMax) Then
      MsgBox "El Período de Gracia no cumple los Parámetros requeridos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If
   
   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If

   If ipp_CuoMen.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Cuota.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_CuoMen)
      Exit Sub
   End If

   If cmb_DiaVct.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Día de Vencimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaVct)
      Exit Sub
   End If

   If cmb_EjeVta.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Ejecutivo de Ventas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EjeVta)
      Exit Sub
   End If
   
   'Validando Documentos a Recibir
   r_int_FlgDoc = 1
   
   r_str_Selecc = ""
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      r_str_Selecc = Trim(grd_Listad.Text)
      
      grd_Listad.Col = 6
      
      'Si es de Obligatoria Selección
      If Len(Trim(r_str_Selecc)) = 0 And CInt(grd_Listad.Text) = 1 Then
         r_int_FlgDoc = 2
         Exit For
      End If
   Next r_int_Contad
   Call gs_UbiIniGrid(grd_Listad)
   
   If r_int_FlgDoc = 2 Then
      MsgBox "Debe seleccionar los Documentos Crediticios que han sido recibidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   modatecli_g_arr_DatCre(1).DatCre_TipMon = cmb_TipMon.ItemData(cmb_TipMon.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_Dol_ComVta = CDbl(ipp_ComVta_Dolare.Text)
   modatecli_g_arr_DatCre(1).DatCre_Dol_ApoPro = CDbl(ipp_ApoPro_Dolare.Text)
   modatecli_g_arr_DatCre(1).DatCre_Dol_MonSol = CDbl(ipp_MonSol_Dolare.Text)
   modatecli_g_arr_DatCre(1).DatCre_Sol_MonSol = CDbl(pnl_MonSol_NSoles.Caption)
   modatecli_g_arr_DatCre(1).DatCre_Pre_MonSol = CDbl(pnl_MonSol_MonPre.Caption)
   modatecli_g_arr_DatCre(1).DatCre_Dol_TipCam = l_dbl_Dol_TipCam
   modatecli_g_arr_DatCre(1).DatCre_Pre_TipCam = l_dbl_Pre_TipCam
   modatecli_g_arr_DatCre(1).DatCre_PlaAno = ipp_PlaAno.Value
   modatecli_g_arr_DatCre(1).DatCre_PlaMes = ipp_PlaMes.Value
   modatecli_g_arr_DatCre(1).DatCre_PerGra = ipp_PerGra.Value
   modatecli_g_arr_DatCre(1).DatCre_ESgDes = moddat_g_str_EmpSegDes
   modatecli_g_arr_DatCre(1).DatCre_ESgViv = moddat_g_str_EmpSegViv
   modatecli_g_arr_DatCre(1).DatCre_TipSeg = cmb_SegDes.ItemData(cmb_SegDes.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_CuoMen = ipp_CuoMen.Value
   modatecli_g_arr_DatCre(1).DatCre_CuoAno = cmb_CuoAno.ItemData(cmb_CuoAno.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_DiaVct = CInt(cmb_DiaVct.Text)
   modatecli_g_arr_DatCre(1).DatCre_EjeVta = l_arr_EjeVta(cmb_EjeVta.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_Observ = txt_Observ.Text

   ReDim modatecli_g_arr_DocCre(0)
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      If grd_Listad.Text = "X" Then
         ReDim Preserve modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre) + 1)
         
         grd_Listad.Col = 2
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_TipDoc = CInt(grd_Listad.Text)
      
         grd_Listad.Col = 3
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_CodGrp = grd_Listad.Text
      
         grd_Listad.Col = 4
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_CodAct = CInt(grd_Listad.Text)
      
         grd_Listad.Col = 5
         modatecli_g_arr_DocCre(UBound(modatecli_g_arr_DocCre)).DocCre_CodIte = grd_Listad.Text
      End If
   Next r_int_Contad
   
   modatecli_g_int_DatCreTit = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   If MsgBox("Al salir de esta manera perderá la información ingresada. ¿Está seguro de salir de la ventana?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   Dim r_int_ConAux     As Integer
   Dim r_int_TipDoc     As Integer
   Dim r_str_CodGrp     As String
   Dim r_int_CodAct     As Integer
   Dim r_str_CodIte     As String

   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt & " Ingreso de Solicitud de Crédito"
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   If modatecli_g_int_DatCreTit = 2 Then
      Call gs_BuscarCombo_Item(cmb_TipMon, modatecli_g_arr_DatCre(1).DatCre_TipMon)
      
      ipp_ComVta_Dolare.Value = modatecli_g_arr_DatCre(1).DatCre_Dol_ComVta
      ipp_MonSol_Dolare.Value = modatecli_g_arr_DatCre(1).DatCre_Dol_MonSol
      ipp_ApoPro_Dolare.Value = modatecli_g_arr_DatCre(1).DatCre_Dol_ApoPro
      ipp_PlaAno.Value = modatecli_g_arr_DatCre(1).DatCre_PlaAno
      ipp_PlaMes.Value = modatecli_g_arr_DatCre(1).DatCre_PlaMes
      ipp_PerGra.Value = modatecli_g_arr_DatCre(1).DatCre_PerGra
      
      Call gs_BuscarCombo_Item(cmb_CuoAno, modatecli_g_arr_DatCre(1).DatCre_CuoAno)
      Call gs_BuscarCombo_Item(cmb_SegDes, modatecli_g_arr_DatCre(1).DatCre_TipSeg)
      
      ipp_CuoMen.Value = modatecli_g_arr_DatCre(1).DatCre_CuoMen
      Call gs_BuscarCombo_Text(cmb_DiaVct, Format(modatecli_g_arr_DatCre(1).DatCre_DiaVct, "00"), 2)
      
      cmb_EjeVta.ListIndex = gf_Busca_Arregl(l_arr_EjeVta, modatecli_g_arr_DatCre(1).DatCre_EjeVta) - 1
      
      txt_Observ.Text = modatecli_g_arr_DatCre(1).DatCre_Observ
      
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         grd_Listad.Row = r_int_Contad
         
         grd_Listad.Col = 2:     r_int_TipDoc = CInt(grd_Listad.Text)
         grd_Listad.Col = 3:     r_str_CodGrp = grd_Listad.Text
         grd_Listad.Col = 4:     r_int_CodAct = CInt(grd_Listad.Text)
         grd_Listad.Col = 5:     r_str_CodIte = grd_Listad.Text
         
         For r_int_ConAux = 0 To UBound(modatecli_g_arr_DocCre)
            If modatecli_g_arr_DocCre(r_int_ConAux).DocCre_TipDoc = r_int_TipDoc And modatecli_g_arr_DocCre(r_int_ConAux).DocCre_CodGrp = r_str_CodGrp _
               And modatecli_g_arr_DocCre(r_int_ConAux).DocCre_CodAct = r_int_CodAct And modatecli_g_arr_DocCre(r_int_ConAux).DocCre_CodIte = r_str_CodIte Then
               grd_Listad.Col = 1
               grd_Listad.Text = "X"
               
               Exit For
            End If
         Next r_int_ConAux
      Next r_int_Contad
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_ParPrd_ComboItem(cmb_TipMon, moddat_g_str_CodPrd, "002")
   
   moddat_g_str_EmpSegDes = moddat_gf_Consulta_EmpSeg(1)
   moddat_g_str_EmpSegViv = moddat_gf_Consulta_EmpSeg(2)
      
   Call moddat_gs_Carga_TipSeg(cmb_SegDes, moddat_g_str_EmpSegDes)
   Call moddat_gs_Carga_LisIte_Combo(cmb_DiaVct, 1, "222")
   Call moddat_gs_Carga_LisIte_Combo(cmb_CuoAno, 1, "223")
   
   Call moodat_gs_Carga_EjeVta(cmb_EjeVta, l_arr_EjeVta)
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 9690
   grd_Listad.ColWidth(1) = 1290
   grd_Listad.ColWidth(2) = 0
   grd_Listad.ColWidth(3) = 0
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   
   Call fs_Carga_Documen
   
   l_dbl_Dol_TipCam = 0
   l_dbl_Pre_TipCam = 0
   
   l_dbl_Dol_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
   
   If l_dbl_Dol_TipCam = 0 Then
      MsgBox "No se encuentra registrado el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub fs_Limpia()
   cmb_TipMon.ListIndex = -1
   
   ipp_ComVta_Dolare.Value = 0
   pnl_ComVta_NSoles.Caption = "0.00 "
   pnl_ComVta_MonPre.Caption = "0.00 "
   
   ipp_MonSol_Dolare.Value = 0
   pnl_MonSol_NSoles.Caption = "0.00 "
   pnl_MonSol_MonPre.Caption = "0.00 "
   
   ipp_PlaAno.Value = 0
   ipp_PlaMes.Value = 0
   ipp_PerGra.Value = 0
   
   cmb_CuoAno.ListIndex = -1
   cmb_SegDes.ListIndex = -1
   ipp_CuoMen.Value = 0
   cmb_DiaVct.ListIndex = -1
   cmb_EjeVta.ListIndex = -1
   txt_Observ.Text = ""
   pnl_PlaTot.Caption = "0 "
End Sub

Private Sub fs_Carga_Documen()
   Call gs_LimpiaGrid(grd_Listad)
   
   'Documentos Crediticios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '301' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     MsgBox "No se han encontrado Lista de Documentos Crediticios.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
      grd_Listad.Col = 1:     grd_Listad.Text = ""
      grd_Listad.Col = 2:     grd_Listad.Text = "1"
      grd_Listad.Col = 3:     grd_Listad.Text = "301"
      grd_Listad.Col = 4:     grd_Listad.Text = "0"
      grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARPRD_CODITE
      grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARPRD_DESCRI, 1)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Documentos por Actividad Económica Titular - Actividad Principal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(modatecli_g_int_ActPri_Tit) & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '301' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
      grd_Listad.Col = 1:     grd_Listad.Text = ""
      grd_Listad.Col = 2:     grd_Listad.Text = "2"
      grd_Listad.Col = 3:     grd_Listad.Text = "301"
      grd_Listad.Col = 4:     grd_Listad.Text = modatecli_g_int_ActPri_Tit
      grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
      grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARACT_DESCRI, 1)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   
   'Documentos por Actividad Económica Titular - Actividad Secundaria
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(modatecli_g_int_ActSec_Tit) & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '301' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
      grd_Listad.Col = 1:     grd_Listad.Text = ""
      grd_Listad.Col = 2:     grd_Listad.Text = "2"
      grd_Listad.Col = 3:     grd_Listad.Text = "301"
      grd_Listad.Col = 4:     grd_Listad.Text = modatecli_g_int_ActSec_Tit
      grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
      grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARACT_DESCRI, 1)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   
   'Documentos por Actividad Económica Cónyuge - Actividad Principal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(modatecli_g_int_ActPri_Cyg) & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '302' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
      grd_Listad.Col = 1:     grd_Listad.Text = ""
      grd_Listad.Col = 2:     grd_Listad.Text = "2"
      grd_Listad.Col = 3:     grd_Listad.Text = "302"
      grd_Listad.Col = 4:     grd_Listad.Text = modatecli_g_int_ActPri_Cyg
      grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
      grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARACT_DESCRI, 1)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   
   'Documentos por Actividad Económica Cónyuge - Actividad Secundaria
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(modatecli_g_int_ActSec_Cyg) & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '302' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
      grd_Listad.Col = 1:     grd_Listad.Text = ""
      grd_Listad.Col = 2:     grd_Listad.Text = "2"
      grd_Listad.Col = 3:     grd_Listad.Text = "302"
      grd_Listad.Col = 4:     grd_Listad.Text = modatecli_g_int_ActSec_Cyg
      grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
      grd_Listad.Col = 6:     grd_Listad.Text = Left(g_rst_Genera!PARACT_DESCRI, 1)
      
      g_rst_Genera.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 1
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)

   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 32 Then
      Call grd_Listad_DblClick
   End If
End Sub

Private Sub ipp_ApoPro_Dolare_Change()
   If cmb_TipMon.ListIndex > -1 Then
      l_dbl_Pre_TipCam = moddat_gf_Obtiene_TipCam(1, cmb_TipMon.ItemData(cmb_TipMon.ListIndex))
      Call fs_Calcula(3)
   End If

   If ipp_ComVta_Dolare.Value > 0 Then
      ipp_MonSol_Dolare.Value = ipp_ComVta_Dolare.Value - ipp_ApoPro_Dolare.Value
   Else
      ipp_MonSol_Dolare.Value = 0
   End If
End Sub

Private Sub ipp_ApoPro_Dolare_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_ComVta_Dolare_Change()
   If cmb_TipMon.ListIndex > -1 Then
      l_dbl_Pre_TipCam = moddat_gf_Obtiene_TipCam(1, cmb_TipMon.ItemData(cmb_TipMon.ListIndex))
      Call fs_Calcula(1)
   End If
   
   If ipp_ComVta_Dolare.Value > 0 Then
      ipp_MonSol_Dolare.Value = ipp_ComVta_Dolare.Value - ipp_ApoPro_Dolare.Value
   Else
      ipp_MonSol_Dolare.Value = 0
   End If
End Sub

Private Sub ipp_ComVta_Dolare_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ApoPro_Dolare)
   End If
End Sub

Private Sub ipp_CuoMen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DiaVct)
   End If
End Sub

Private Sub ipp_MonSol_Dolare_Change()
   If cmb_TipMon.ListIndex > -1 Then
      l_dbl_Pre_TipCam = moddat_gf_Obtiene_TipCam(1, cmb_TipMon.ItemData(cmb_TipMon.ListIndex))
      Call fs_Calcula(2)
   End If
End Sub

Private Sub ipp_PlaAno_Change()
   pnl_PlaTot.Caption = Format((ipp_PlaAno.Value * 12) + ipp_PlaMes.Value, "##0") & " "
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaMes)
   End If
End Sub

Private Sub ipp_PlaMes_Change()
   Call ipp_PlaAno_Change
End Sub

Private Sub ipp_PlaMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CuoAno)
   End If
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_SegDes)
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(grd_Listad)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Calcula(ByVal p_TipCal As Integer)
   Dim r_dbl_ValSol     As Double
   Dim r_dbl_ValPre     As Double
   
   Select Case p_TipCal
      Case 1
         r_dbl_ValSol = ipp_ComVta_Dolare.Value * l_dbl_Dol_TipCam
         If l_dbl_Pre_TipCam > 0 Then
            r_dbl_ValPre = ipp_ComVta_Dolare.Value * l_dbl_Dol_TipCam / l_dbl_Pre_TipCam
         Else
            r_dbl_ValPre = 0
         End If
         
         pnl_ComVta_NSoles.Caption = Format(r_dbl_ValSol, "###,###,##0.00") & " "
         pnl_ComVta_MonPre.Caption = Format(r_dbl_ValPre, "###,###,##0.00") & " "
         
      Case 2
         r_dbl_ValSol = ipp_MonSol_Dolare.Value * l_dbl_Dol_TipCam
         
         If l_dbl_Pre_TipCam > 0 Then
            r_dbl_ValPre = ipp_MonSol_Dolare.Value * l_dbl_Dol_TipCam / l_dbl_Pre_TipCam
         Else
            r_dbl_ValPre = 0
         End If
         
         pnl_MonSol_NSoles.Caption = Format(r_dbl_ValSol, "###,###,##0.00") & " "
         pnl_MonSol_MonPre.Caption = Format(r_dbl_ValPre, "###,###,##0.00") & " "
   
      Case 3
         r_dbl_ValSol = ipp_ApoPro_Dolare.Value * l_dbl_Dol_TipCam
         
         If l_dbl_Pre_TipCam > 0 Then
            r_dbl_ValPre = ipp_ApoPro_Dolare.Value * l_dbl_Dol_TipCam / l_dbl_Pre_TipCam
         Else
            r_dbl_ValPre = 0
         End If
         
         pnl_ApoPro_NSoles.Caption = Format(r_dbl_ValSol, "###,###,##0.00") & " "
         pnl_ApoPro_MonPre.Caption = Format(r_dbl_ValPre, "###,###,##0.00") & " "
   End Select
End Sub
