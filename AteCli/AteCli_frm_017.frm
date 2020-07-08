VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_Seguro_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   855
   ClientTop       =   1185
   ClientWidth     =   12825
   Icon            =   "AteCli_frm_017.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7995
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   14102
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
         Height          =   765
         Left            =   30
         TabIndex        =   63
         Top             =   7170
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11340
            Picture         =   "AteCli_frm_017.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12030
            Picture         =   "AteCli_frm_017.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1425
         Left            =   30
         TabIndex        =   31
         Top             =   1590
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            TabIndex        =   32
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
            Left            =   8940
            TabIndex        =   33
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
            TabIndex        =   34
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
            TabIndex        =   35
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
            TabIndex        =   36
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
            Left            =   8940
            TabIndex        =   37
            Top             =   720
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
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
            Left            =   8940
            TabIndex        =   38
            Top             =   60
            Width           =   3735
            _Version        =   65536
            _ExtentX        =   6588
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "DOLARES AMERICANOS"
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   44
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Left            =   60
            TabIndex        =   43
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "Modalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   40
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label24 
            Caption         =   "Moneda Prést.:"
            Height          =   315
            Left            =   7470
            TabIndex        =   39
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   3255
         Left            =   30
         TabIndex        =   22
         Top             =   3870
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   5741
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
         Begin VB.TextBox txt_InfViv 
            Height          =   315
            Left            =   1620
            MaxLength       =   30
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   1530
            Width           =   2895
         End
         Begin VB.ComboBox cmb_AplViv 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1860
            Width           =   2925
         End
         Begin VB.ComboBox cmb_AplPre 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   720
            Width           =   2925
         End
         Begin VB.TextBox txt_InfPre 
            Height          =   315
            Left            =   1620
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   390
            Width           =   2895
         End
         Begin VB.TextBox txt_Observ 
            Height          =   885
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Text            =   "AteCli_frm_017.frx":0758
            Top             =   2340
            Width           =   11025
         End
         Begin EditLib.fpDateTime ipp_FecPre 
            Height          =   315
            Left            =   8940
            TabIndex        =   11
            Top             =   390
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
         Begin EditLib.fpDoubleSingle ipp_FoiPre 
            Height          =   315
            Left            =   8940
            TabIndex        =   13
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
            Text            =   "0.000000000"
            DecimalPlaces   =   9
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
         Begin Threed.SSPanel pnl_SegPre 
            Height          =   315
            Left            =   1620
            TabIndex        =   54
            Top             =   60
            Width           =   11025
            _Version        =   65536
            _ExtentX        =   19447
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   90
            Left            =   30
            TabIndex        =   55
            Top             =   1080
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
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
         Begin EditLib.fpDateTime ipp_FecViv 
            Height          =   315
            Left            =   8940
            TabIndex        =   15
            Top             =   1530
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
         Begin EditLib.fpDoubleSingle ipp_FoiViv 
            Height          =   315
            Left            =   8940
            TabIndex        =   17
            Top             =   1860
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
            Text            =   "0.000000000"
            DecimalPlaces   =   9
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
         Begin Threed.SSPanel pnl_SegViv 
            Height          =   315
            Left            =   1620
            TabIndex        =   56
            Top             =   1200
            Width           =   11025
            _Version        =   65536
            _ExtentX        =   19447
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   90
            Left            =   30
            TabIndex        =   57
            Top             =   2220
            Width           =   12675
            _Version        =   65536
            _ExtentX        =   22357
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
         Begin VB.Label Label21 
            Caption         =   "Número Informe:"
            Height          =   315
            Left            =   60
            TabIndex        =   62
            Top             =   1530
            Width           =   1395
         End
         Begin VB.Label Label15 
            Caption         =   "F. Informe:"
            Height          =   315
            Left            =   7470
            TabIndex        =   61
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Factor/Importe:"
            Height          =   315
            Left            =   7470
            TabIndex        =   60
            Top             =   1860
            Width           =   1545
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo de Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   59
            Top             =   1860
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Seguro Vivienda:"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label Label12 
            Caption         =   "Seguro Préstamo:"
            Height          =   315
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo de Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Factor/Importe:"
            Height          =   315
            Left            =   7470
            TabIndex        =   26
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label Label9 
            Caption         =   "F. Informe:"
            Height          =   315
            Left            =   7470
            TabIndex        =   25
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Número Informe:"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label16 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   2340
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   27
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
            Left            =   630
            TabIndex        =   28
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Evaluación de Seguros"
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
            TabIndex        =   29
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
            Picture         =   "AteCli_frm_017.frx":075C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   30
         TabIndex        =   30
         Top             =   3060
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
         Begin VB.CommandButton cmd_RegInf 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_017.frx":0A66
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Registrar Evaluación"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   11280
            Picture         =   "AteCli_frm_017.frx":0D70
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_017.frx":107A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   46
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
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
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_017.frx":14BC
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11280
            Picture         =   "AteCli_frm_017.frx":18FE
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10560
            Picture         =   "AteCli_frm_017.frx":1C08
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
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
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   51
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   50
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   49
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   48
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   47
            Top             =   1740
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_Seguro_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_IniEva     As String
Dim l_str_Aprueb     As String
Dim l_str_Rechaz     As String
Dim l_int_NumEva     As Integer

Private Sub cmb_AplPre_Click()
   Call gs_SetFocus(ipp_FoiPre)
End Sub

Private Sub cmb_AplPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_AplPre_Click
   End If
End Sub

Private Sub cmb_AplViv_Click()
   Call gs_SetFocus(ipp_FoiViv)
End Sub

Private Sub cmb_AplViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_AplViv_Click
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
   
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaSeg, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaSeg, 12, 0, txt_Observ.Text, 0, 0) Then
      Exit Sub
   End If
   
   'Verificar si la Instancia de Tasacion ha sido aprobada
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(modatecli_g_con_EvaTas)

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
      MsgBox "No se encuentra en Instancia de Evaluación de Seguros.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_ActivaItem(False)
   Call fs_Activa(False)

   l_str_IniEva = ""

   'Obteniendo Información del Seguimiento
   Call fs_Buscar_SegDet
   
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
   
   Call fs_Buscar_InfSeg
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfSeg
End Sub

Private Sub cmd_Grabar_Click()
   Call moddat_gs_FecSis
   
   'Seguro de Préstamo
   If Len(Trim(txt_InfPre.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Informe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_InfPre)
      Exit Sub
   End If
      
   If CDate(ipp_FecPre.Text) > CDate(moddat_g_str_FecSis) Or CDate(ipp_FecPre.Text) < CDate(pnl_IniEva.Caption) Then
      MsgBox "La Fecha de Informe no puede ser mayor a la de hoy ni menor a la Fecha de Inicio de Evaluación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPre)
      Exit Sub
   End If
      
   If cmb_AplPre.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Aplicación (Factor o Importe Mensual).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AplPre)
      Exit Sub
   End If
      
   If ipp_FoiPre.Value = 0 Then
      MsgBox "Debe ingresar el Factor/Importe del Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FoiPre)
      Exit Sub
   End If
      
   'Seguro de Vivienda
   If Len(Trim(txt_InfViv.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Informe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_InfViv)
      Exit Sub
   End If
      
   If CDate(ipp_FecViv.Text) > CDate(moddat_g_str_FecSis) Or CDate(ipp_FecViv.Text) < CDate(pnl_IniEva.Caption) Then
      MsgBox "La Fecha de Informe no puede ser mayor a la de hoy ni menor a la Fecha de Inicio de Evaluación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecViv)
      Exit Sub
   End If
      
   If cmb_AplViv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Aplicación (Factor o Importe Mensual).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AplViv)
      Exit Sub
   End If
      
   If ipp_FoiViv.Value = 0 Then
      MsgBox "Debe ingresar el Factor/Importe del Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FoiViv)
      Exit Sub
   End If
   
   'Registra el Informe de Tasación
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_EVASEG ("

      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & txt_InfPre.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPre.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_AplPre.ItemData(cmb_AplPre.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_FoiPre.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_InfViv.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecViv.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_AplViv.ItemData(cmb_AplViv.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_FoiViv.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_EVASEG. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Grabando en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaSeg, 42, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   MsgBox "Se grabaron los datos correctamente.", vbInformation, modgen_g_str_NomPlt
      
   Call fs_ActivaItem(False)
   Call fs_Buscar_InfSeg
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = modatecli_g_con_EvaSeg
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_EvaSeg, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_EvaSeg, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
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
   
      modgen_g_str_Mail_Asunto = "RECHAZO DE EVALUACION DE SEGUROS (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = r_str_Cadena
      
      frm_EnvMai_01.Show 1
      
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      
      Call cmd_Limpia_Click
   End If
End Sub

Private Sub cmd_RegInf_Click()
   'Activando Botones
   cmd_Grabar.Enabled = True
   cmd_Cancel.Enabled = True
   
   cmd_RegInf.Enabled = False
   cmd_Aprueb.Enabled = False
   cmd_Rechaz.Enabled = False
   
   txt_InfPre.Enabled = True
   ipp_FecPre.Enabled = True
   cmb_AplPre.Enabled = True
   ipp_FoiPre.Enabled = True
   
   txt_InfViv.Enabled = True
   ipp_FecViv.Enabled = True
   cmb_AplViv.Enabled = True
   ipp_FoiViv.Enabled = True
   
   txt_Observ.Enabled = True
   
   Call gs_SetFocus(txt_InfPre)
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
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_AplPre, 1, "227")
   Call moddat_gs_Carga_LisIte_Combo(cmb_AplViv, 1, "227")
End Sub

Private Sub ipp_FecPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_AplPre)
   End If
End Sub

Private Sub ipp_FecViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_AplViv)
   End If
End Sub

Private Sub ipp_FoiPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_InfViv)
   End If
End Sub

Private Sub ipp_FoiViv_KeyPress(KeyAscii As Integer)
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
   
   pnl_SegPre.Caption = ""
   pnl_SegViv.Caption = ""
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   Call moddat_gs_FecSis
   
   txt_InfPre.Text = ""
   ipp_FecPre.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   cmb_AplPre.ListIndex = -1
   ipp_FoiPre.Value = 0
   
   txt_InfViv.Text = ""
   ipp_FecViv.Text = Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy")
   cmb_AplViv.ListIndex = -1
   ipp_FoiViv.Value = 0
   
   txt_Observ.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   cmd_RegInf.Enabled = Not p_Habilita
   cmd_Aprueb.Enabled = Not p_Habilita
   cmd_Rechaz.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   txt_InfPre.Enabled = p_Habilita
   ipp_FecPre.Enabled = p_Habilita
   cmb_AplPre.Enabled = p_Habilita
   ipp_FoiPre.Enabled = p_Habilita
   
   txt_InfViv.Enabled = p_Habilita
   ipp_FecViv.Enabled = p_Habilita
   cmb_AplViv.Enabled = p_Habilita
   ipp_FoiViv.Enabled = p_Habilita
   
   txt_Observ.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_RegInf.Enabled = p_Habilita
   cmd_Aprueb.Enabled = p_Habilita
   cmd_Rechaz.Enabled = p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
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
   
   'Información de Seguros
   pnl_SegPre.Caption = Trim(moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES)) & " / " & Trim(moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG))
   pnl_SegViv.Caption = Trim(moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGVIV))
   
   ipp_FoiPre.Value = moddat_gf_Consulta_AplSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG, moddat_g_int_TipMon, g_rst_Princi!SOLMAE_PREMPR, cmb_AplPre)
   ipp_FoiViv.Value = moddat_gf_Consulta_AplSeg(g_rst_Princi!SOLMAE_ESGVIV, 0, moddat_g_int_TipMon, g_rst_Princi!SOLMAE_PREMPR, cmb_AplViv)
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_EvaSeg) & " "
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
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If Len(Trim(l_str_IniEva)) > 0 Then
      pnl_IniEva.Caption = l_str_IniEva
   End If
End Sub

Private Sub txt_InfPre_GotFocus()
   Call gs_SelecTodo(txt_InfPre)
End Sub

Private Sub txt_InfPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecPre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;:;")
   End If
End Sub

Private Sub txt_InfViv_GotFocus()
   Call gs_SelecTodo(txt_InfViv)
End Sub

Private Sub txt_InfViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecViv)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;:;")
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

Private Sub fs_Buscar_InfSeg()
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   cmd_RegInf.Enabled = True
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      cmd_Aprueb.Enabled = False
      cmd_Rechaz.Enabled = False
      
      moddat_g_int_FlgGrb = 1
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   cmd_Aprueb.Enabled = True
   cmd_Rechaz.Enabled = True
      
   'Cargar Datos de Evaluación
   txt_InfPre.Text = Trim(g_rst_Princi!EVASEG_INFPRE & "")
   ipp_FecPre.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAPRE))
   Call gs_BuscarCombo_Item(cmb_AplPre, g_rst_Princi!EVASEG_TAPPRE)
   ipp_FoiPre.Value = g_rst_Princi!EVASEG_TASPRE
   
   txt_InfViv.Text = Trim(g_rst_Princi!EVASEG_INFVIV & "")
   ipp_FecViv.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
   Call gs_BuscarCombo_Item(cmb_AplViv, g_rst_Princi!EVASEG_TAPVIV)
   ipp_FoiViv.Value = g_rst_Princi!EVASEG_TASVIV
   
   txt_Observ.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   moddat_g_int_FlgGrb = 2
End Sub


