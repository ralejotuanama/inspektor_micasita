VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_SimCre_11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   10380
   ClientLeft      =   4575
   ClientTop       =   1455
   ClientWidth     =   12885
   Icon            =   "AteCli_frm_532.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10395
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   12885
      _Version        =   65536
      _ExtentX        =   22728
      _ExtentY        =   18336
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   765
         Left            =   30
         TabIndex        =   34
         Top             =   1380
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   10515
         End
         Begin VB.ComboBox cmb_SubPrd 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   10515
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   36
            Top             =   90
            Width           =   885
         End
         Begin VB.Label Label4 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   35
            Top             =   420
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
            Height          =   615
            Left            =   630
            TabIndex        =   38
            Top             =   30
            Width           =   8505
            _Version        =   65536
            _ExtentX        =   15002
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "Simulación de Créditos Hipotecarios"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10920
            Top             =   60
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_532.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   645
         Left            =   30
         TabIndex        =   39
         Top             =   720
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   1830
            Picture         =   "AteCli_frm_532.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Imprimir Cronograma de Simulación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CalCuo 
            Height          =   585
            Left            =   630
            Picture         =   "AteCli_frm_532.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Calcular Cuota"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1230
            Picture         =   "AteCli_frm_532.frx":0932
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Imprimir Simulación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   30
            Picture         =   "AteCli_frm_532.frx":0D74
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12180
            Picture         =   "AteCli_frm_532.frx":107E
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   4455
         Left            =   30
         TabIndex        =   40
         Top             =   2160
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
         _ExtentY        =   7858
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
         Begin VB.ComboBox cmb_BMSTas 
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2850
            Width           =   915
         End
         Begin VB.ComboBox cmb_TasEsp 
            Height          =   315
            Left            =   11220
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   780
            Width           =   1515
         End
         Begin VB.ComboBox cmb_CuoDbl 
            Height          =   315
            Left            =   7710
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   450
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipSeg 
            Height          =   315
            Left            =   7710
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   780
            Width           =   1995
         End
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   11220
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   450
            Width           =   1515
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta 
            Height          =   315
            Left            =   2310
            TabIndex        =   3
            Top             =   450
            Width           =   1785
            _Version        =   196608
            _ExtentX        =   3149
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
         Begin EditLib.fpDoubleSingle ipp_ApoPro 
            Height          =   315
            Left            =   2310
            TabIndex        =   6
            Top             =   1530
            Width           =   1785
            _Version        =   196608
            _ExtentX        =   3149
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
            Left            =   7710
            TabIndex        =   16
            Top             =   120
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
            MaxValue        =   "0"
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
            Left            =   11220
            TabIndex        =   17
            Top             =   120
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_CuoMen 
            Height          =   315
            Left            =   8130
            TabIndex        =   41
            Top             =   3270
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_IngReq 
            Height          =   315
            Left            =   8130
            TabIndex        =   42
            Top             =   3600
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_IngReq_Sol 
            Height          =   315
            Left            =   11430
            TabIndex        =   43
            Top             =   3600
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_CuoPBP 
            Height          =   315
            Left            =   11430
            TabIndex        =   44
            Top             =   2190
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_MtoPBP 
            Height          =   315
            Left            =   11430
            TabIndex        =   45
            Top             =   2520
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_ValEst_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   46
            Top             =   780
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoPre_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   47
            Top             =   3270
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoPre 
            Height          =   315
            Left            =   2070
            TabIndex        =   12
            Top             =   3270
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_ApoPro_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   48
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipCam 
            Height          =   315
            Left            =   11430
            TabIndex        =   49
            Top             =   1530
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000  "
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
         Begin Threed.SSPanel pnl_TasInt 
            Height          =   315
            Left            =   8130
            TabIndex        =   50
            Top             =   1530
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000  "
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
         Begin Threed.SSPanel pnl_CosEfe 
            Height          =   315
            Left            =   8130
            TabIndex        =   51
            Top             =   2520
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000  "
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
         Begin Threed.SSPanel pnl_SegDes 
            Height          =   315
            Left            =   8130
            TabIndex        =   52
            Top             =   1860
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000  "
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
         Begin Threed.SSPanel pnl_SegInm 
            Height          =   315
            Left            =   8130
            TabIndex        =   53
            Top             =   2190
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.0000  "
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
         Begin Threed.SSPanel pnl_CuoMen_Sol 
            Height          =   315
            Left            =   11430
            TabIndex        =   54
            Top             =   3270
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_CuoSBP 
            Height          =   315
            Left            =   11430
            TabIndex        =   110
            Top             =   1860
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_PorIni 
            Height          =   285
            Left            =   4110
            TabIndex        =   113
            Top             =   1215
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   " 0.00%"
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_FmvBbp 
            Height          =   315
            Left            =   2310
            TabIndex        =   7
            Top             =   1860
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin EditLib.fpDoubleSingle ipp_MtoAFP 
            Height          =   315
            Left            =   2310
            TabIndex        =   9
            Top             =   2520
            Width           =   1785
            _Version        =   196608
            _ExtentX        =   3149
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
         Begin Threed.SSPanel pnl_MtoBMS 
            Height          =   315
            Left            =   3240
            TabIndex        =   11
            Top             =   2850
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_MtoAFP_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   122
            Top             =   2520
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MtoBMS_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   124
            Top             =   2850
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_FmvBbp_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   126
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoIni 
            Height          =   315
            Left            =   2070
            TabIndex        =   5
            Top             =   1200
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_CuoIni_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   129
            Top             =   1200
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_MefPbp 
            Height          =   315
            Left            =   2310
            TabIndex        =   8
            Top             =   2190
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_MefPbp_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   132
            Top             =   2190
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValTot 
            Height          =   315
            Left            =   2070
            TabIndex        =   2
            Top             =   120
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_ValTot_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   136
            Top             =   120
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
            Alignment       =   4
         End
         Begin EditLib.fpDoubleSingle ipp_ValEst 
            Height          =   315
            Left            =   2310
            TabIndex        =   4
            Top             =   780
            Width           =   1785
            _Version        =   196608
            _ExtentX        =   3149
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
         Begin Threed.SSPanel pnl_ComVta_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   142
            Top             =   450
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValGas_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   144
            Top             =   3600
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValGas 
            Height          =   315
            Left            =   2070
            TabIndex        =   14
            Top             =   3600
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_TotPre_Sol 
            Height          =   315
            Left            =   5070
            TabIndex        =   147
            Top             =   3930
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotPre 
            Height          =   315
            Left            =   2070
            TabIndex        =   15
            Top             =   3930
            Width           =   2025
            _Version        =   65536
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSCheck chk_Gastos 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   3630
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Incluye Gastos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label20 
            Caption         =   "Ingreso Requerido:"
            Height          =   315
            Left            =   6420
            TabIndex        =   151
            Top             =   3660
            Width           =   1485
         End
         Begin VB.Label Label46 
            Caption         =   "Total del Prestamo"
            Height          =   285
            Left            =   90
            TabIndex        =   150
            Top             =   3990
            Width           =   1515
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   19
            Left            =   1530
            TabIndex        =   149
            Top             =   3990
            Width           =   465
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   148
            Top             =   3990
            Width           =   435
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   18
            Left            =   1530
            TabIndex        =   146
            Top             =   3660
            Width           =   465
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   145
            Top             =   3660
            Width           =   435
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   143
            Top             =   510
            Width           =   435
         End
         Begin VB.Label Label41 
            Caption         =   "Valor Estacio."
            Height          =   285
            Left            =   480
            TabIndex        =   141
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   17
            Left            =   1770
            TabIndex        =   140
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label40 
            Caption         =   "Valor Total Vivienda"
            Height          =   285
            Left            =   120
            TabIndex        =   139
            Top             =   180
            Width           =   1515
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   138
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   16
            Left            =   1530
            TabIndex        =   137
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label38 
            Caption         =   "Bono (PBP)"
            Height          =   285
            Left            =   480
            TabIndex        =   135
            Top             =   2250
            Width           =   1215
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   15
            Left            =   1770
            TabIndex        =   134
            Top             =   2250
            Width           =   465
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   133
            Top             =   2250
            Width           =   435
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   14
            Left            =   1530
            TabIndex        =   131
            Top             =   1260
            Width           =   465
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   130
            Top             =   1260
            Width           =   435
         End
         Begin VB.Label Label34 
            Caption         =   "Cuota Inicial"
            Height          =   285
            Left            =   90
            TabIndex        =   128
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   127
            Top             =   3330
            Width           =   435
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   125
            Top             =   2880
            Width           =   435
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   123
            Top             =   2580
            Width           =   435
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   13
            Left            =   1530
            TabIndex        =   121
            Top             =   3330
            Width           =   465
         End
         Begin VB.Label Label23 
            Caption         =   "Monto del Prestamo"
            Height          =   285
            Left            =   90
            TabIndex        =   120
            Top             =   3330
            Width           =   1515
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   12
            Left            =   1770
            TabIndex        =   119
            Top             =   2910
            Width           =   465
         End
         Begin VB.Label Label22 
            Caption         =   "BMS (B. Verde)"
            Height          =   285
            Left            =   480
            TabIndex        =   118
            Top             =   2910
            Width           =   1215
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   11
            Left            =   1770
            TabIndex        =   117
            Top             =   2580
            Width           =   465
         End
         Begin VB.Label Label21 
            Caption         =   "AFP (25%)"
            Height          =   285
            Left            =   480
            TabIndex        =   116
            Top             =   2580
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Tasa Especial:"
            Height          =   285
            Left            =   9870
            TabIndex        =   115
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label Label17 
            Caption         =   "Cuotas Dobles:"
            Height          =   285
            Left            =   6420
            TabIndex        =   114
            Top             =   510
            Width           =   1305
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   7
            Left            =   10980
            TabIndex        =   112
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label lbl_CuoSBP 
            Caption         =   "Cuota s / PBP:"
            Height          =   315
            Left            =   9870
            TabIndex        =   111
            Top             =   1920
            Width           =   1155
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Left            =   10980
            TabIndex        =   82
            Top             =   1590
            Width           =   375
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   255
            Left            =   10920
            TabIndex        =   81
            Top             =   3330
            Width           =   465
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   255
            Left            =   10920
            TabIndex        =   80
            Top             =   3660
            Width           =   465
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   5
            Left            =   10980
            TabIndex        =   79
            Top             =   2250
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "Seguro Inmueble:"
            Height          =   315
            Left            =   6420
            TabIndex        =   78
            Top             =   2250
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Seguro Desgravamen:"
            Height          =   315
            Left            =   6420
            TabIndex        =   77
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Costo Efectivo:"
            Height          =   315
            Left            =   6420
            TabIndex        =   76
            Top             =   2580
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "Tasa Interés:"
            Height          =   315
            Left            =   6420
            TabIndex        =   75
            Top             =   1590
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo Cambio:"
            Height          =   315
            Left            =   9870
            TabIndex        =   74
            Top             =   1590
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   73
            Top             =   1920
            Width           =   435
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   2
            Left            =   1770
            TabIndex        =   72
            Top             =   1920
            Width           =   465
         End
         Begin VB.Label Label6 
            Caption         =   "Bono (BBP)"
            Height          =   285
            Left            =   480
            TabIndex        =   71
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   70
            Top             =   1590
            Width           =   435
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   285
            Left            =   4560
            TabIndex        =   69
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   1
            Left            =   1770
            TabIndex        =   68
            Top             =   1590
            Width           =   465
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   0
            Left            =   1770
            TabIndex        =   67
            Top             =   510
            Width           =   465
         End
         Begin VB.Label lbl_CuoPBP 
            Caption         =   "Cuota c / PBP:"
            Height          =   315
            Left            =   9870
            TabIndex        =   66
            Top             =   2250
            Width           =   1155
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   6
            Left            =   10980
            TabIndex        =   65
            Top             =   2580
            Width           =   375
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   4
            Left            =   7590
            TabIndex        =   64
            Top             =   3330
            Width           =   375
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   " US$"
            Height          =   285
            Index           =   3
            Left            =   7590
            TabIndex        =   63
            Top             =   3330
            Width           =   375
         End
         Begin VB.Label Label30 
            Caption         =   "Cuota Mensual:"
            Height          =   315
            Left            =   6420
            TabIndex        =   62
            Top             =   3330
            Width           =   1305
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Seg. Desg.:"
            Height          =   285
            Left            =   6420
            TabIndex        =   61
            Top             =   840
            Width           =   1305
         End
         Begin VB.Label Label18 
            Caption         =   "Día de Pago:"
            Height          =   285
            Left            =   9870
            TabIndex        =   60
            Top             =   510
            Width           =   1305
         End
         Begin VB.Label Label25 
            Caption         =   "P.Gracia (meses):"
            Height          =   255
            Left            =   9870
            TabIndex        =   59
            Top             =   180
            Width           =   1305
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo (años):"
            Height          =   285
            Left            =   6420
            TabIndex        =   58
            Top             =   180
            Width           =   1305
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Inmueble:"
            Height          =   285
            Left            =   480
            TabIndex        =   57
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label Label27 
            Caption         =   "Aporte Propio"
            Height          =   285
            Left            =   480
            TabIndex        =   56
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label lbl_MtoPBP 
            Caption         =   "Monto PBP:"
            Height          =   315
            Left            =   9870
            TabIndex        =   55
            Top             =   2580
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2880
         Left            =   30
         TabIndex        =   83
         Top             =   6630
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
         _ExtentY        =   5080
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
         Begin TabDlg.SSTab tab_Cronog 
            Height          =   2775
            Left            =   60
            TabIndex        =   30
            Top             =   30
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   4895
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Cronograma Tramo No Concesional"
            TabPicture(0)   =   "AteCli_frm_532.frx":14C0
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_Totale(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel63"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel60"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel58"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel43"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel42"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel41"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel38"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel5"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel3"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "pnl_Tot_TotCuo_NCo"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_Tot_OtrCar_NCo"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "grd_Listad_NCo"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_Tot_Intere_NCo"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_Tot_SegPre_NCo"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_Tot_SegViv_NCo"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_Tot_Capita_NCo"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).ControlCount=   17
            TabCaption(1)   =   "Cronograma Tramo Concesional"
            TabPicture(1)   =   "AteCli_frm_532.frx":14DC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_Totale(1)"
            Tab(1).Control(1)=   "SSPanel21"
            Tab(1).Control(2)=   "SSPanel20"
            Tab(1).Control(3)=   "SSPanel19"
            Tab(1).Control(4)=   "SSPanel18"
            Tab(1).Control(5)=   "SSPanel17"
            Tab(1).Control(6)=   "SSPanel16"
            Tab(1).Control(7)=   "pnl_Tot_TotCuo_Con"
            Tab(1).Control(8)=   "grd_Listad_Con"
            Tab(1).Control(9)=   "pnl_Tot_Intere_Con"
            Tab(1).Control(10)=   "pnl_Tot_Capita_Con"
            Tab(1).ControlCount=   11
            Begin Threed.SSPanel pnl_Tot_Capita_NCo 
               Height          =   285
               Left            =   2550
               TabIndex        =   84
               Top             =   2370
               Width           =   1420
               _Version        =   65536
               _ExtentX        =   2505
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tot_SegViv_NCo 
               Height          =   285
               Left            =   6755
               TabIndex        =   85
               Top             =   2370
               Width           =   1400
               _Version        =   65536
               _ExtentX        =   2469
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tot_SegPre_NCo 
               Height          =   285
               Left            =   5370
               TabIndex        =   86
               Top             =   2370
               Width           =   1400
               _Version        =   65536
               _ExtentX        =   2469
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tot_Intere_NCo 
               Height          =   285
               Left            =   3965
               TabIndex        =   87
               Top             =   2370
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2505
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad_NCo 
               Height          =   1605
               Left            =   30
               TabIndex        =   31
               Top             =   690
               Width           =   12615
               _ExtentX        =   22251
               _ExtentY        =   2831
               _Version        =   393216
               Rows            =   10
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Tot_OtrCar_NCo 
               Height          =   285
               Left            =   8055
               TabIndex        =   88
               Top             =   2370
               Width           =   1440
               _Version        =   65536
               _ExtentX        =   2531
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tot_TotCuo_NCo 
               Height          =   285
               Left            =   9450
               TabIndex        =   89
               Top             =   2370
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   285
               Left            =   3930
               TabIndex        =   90
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   60
               TabIndex        =   91
               Top             =   390
               Width           =   1125
               _Version        =   65536
               _ExtentX        =   1984
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   1170
               TabIndex        =   92
               Top             =   390
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   2550
               TabIndex        =   93
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   9450
               TabIndex        =   94
               Top             =   390
               Width           =   1260
               _Version        =   65536
               _ExtentX        =   2222
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   10695
               TabIndex        =   95
               Top             =   390
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   5310
               TabIndex        =   96
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   6690
               TabIndex        =   97
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   8070
               TabIndex        =   98
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
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
            Begin Threed.SSPanel pnl_Tot_Capita_Con 
               Height          =   285
               Left            =   -71490
               TabIndex        =   100
               Top             =   3000
               Width           =   1830
               _Version        =   65536
               _ExtentX        =   3228
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tot_Intere_Con 
               Height          =   285
               Left            =   -69660
               TabIndex        =   101
               Top             =   3000
               Width           =   1830
               _Version        =   65536
               _ExtentX        =   3228
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad_Con 
               Height          =   2025
               Left            =   -74970
               TabIndex        =   32
               Top             =   690
               Width           =   12615
               _ExtentX        =   22251
               _ExtentY        =   3572
               _Version        =   393216
               Rows            =   10
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Tot_TotCuo_Con 
               Height          =   285
               Left            =   -67830
               TabIndex        =   102
               Top             =   3000
               Width           =   1830
               _Version        =   65536
               _ExtentX        =   3228
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
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
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -69870
               TabIndex        =   103
               Top             =   390
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -74940
               TabIndex        =   104
               Top             =   390
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -73770
               TabIndex        =   105
               Top             =   390
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -72210
               TabIndex        =   106
               Top             =   390
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -67530
               TabIndex        =   107
               Top             =   390
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -65190
               TabIndex        =   108
               Top             =   390
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   1
               Left            =   -73410
               TabIndex        =   109
               Top             =   3030
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> "
               Height          =   315
               Index           =   0
               Left            =   240
               TabIndex        =   99
               Top             =   2400
               Width           =   1845
            End
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   825
         Left            =   30
         TabIndex        =   152
         Top             =   9525
         Width           =   12825
         _Version        =   65536
         _ExtentX        =   22622
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
         Begin VB.ComboBox cmb_TipIng 
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   90
            Width           =   4755
         End
         Begin VB.CommandButton cmd_CalEst 
            Caption         =   "&Calcular"
            Height          =   315
            Left            =   10170
            TabIndex        =   28
            Top             =   420
            Width           =   1395
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2220
            TabIndex        =   26
            Top             =   90
            Width           =   1845
            _Version        =   196608
            _ExtentX        =   3254
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
         Begin Threed.SSPanel pnl_MtoMax 
            Height          =   315
            Left            =   2220
            TabIndex        =   153
            Top             =   420
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
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
         Begin Threed.SSPanel pnl_CuoApr 
            Height          =   315
            Left            =   6840
            TabIndex        =   154
            Top             =   420
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Cuota Mensual:"
            Height          =   195
            Left            =   4860
            TabIndex        =   162
            Top             =   450
            Width           =   1110
         End
         Begin VB.Label lbl_General 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso Neto:"
            Height          =   195
            Index           =   61
            Left            =   90
            TabIndex        =   161
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Monto Máximo Prést.:"
            Height          =   195
            Left            =   90
            TabIndex        =   160
            Top             =   450
            Width           =   1530
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Ingreso:"
            Height          =   195
            Left            =   4860
            TabIndex        =   159
            Top             =   120
            Width           =   930
         End
         Begin VB.Label lbl_SimMon 
            AutoSize        =   -1  'True
            Caption         =   " "
            Height          =   195
            Index           =   8
            Left            =   5580
            TabIndex        =   158
            Top             =   405
            Width           =   45
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "US$"
            Height          =   195
            Index           =   9
            Left            =   1830
            TabIndex        =   157
            Top             =   480
            Width           =   315
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "US$"
            Height          =   195
            Index           =   10
            Left            =   6480
            TabIndex        =   156
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "S/."
            Height          =   195
            Left            =   1920
            TabIndex        =   155
            Top             =   150
            Width           =   225
         End
      End
   End
End
Attribute VB_Name = "frm_SimCre_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_SubPrd()      As moddat_tpo_Genera
Dim l_arr_DiaPag()      As moddat_tpo_Genera
Dim l_str_CodPrd        As String
Dim l_str_CodSub        As String
Dim l_int_TipMon        As Integer
Dim l_dbl_TasInt        As Double
Dim l_dbl_TipCam        As Double
Dim l_dbl_TasMVi        As Double
Dim l_dbl_ComCof        As Double
Dim l_dbl_TasCof        As Double
Dim l_dbl_Portes        As Double

Dim l_Arr_TNC_Cli()     As String
Dim l_Arr_TC_Cli()      As String
Dim l_Arr_TNC_Cof()     As String
Dim l_Arr_TC_Cof()      As String
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_dbl_MPSMS         As Double

Private Sub chk_Gastos_Click(Value As Integer)
   If chk_Gastos.Value = True Then
'      ipp_ValEst.Enabled = True
   Else
'      ipp_ValEst.Enabled = False
      pnl_ValGas.Caption = "0.00 "
   End If
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub fs_Calcular_Prestamo()
   'Valor inmueble
   pnl_ValTot.Caption = Format(CDbl(ipp_ComVta.Value) + CDbl(ipp_ValEst.Value), "##,###,##0.00") & " "
   If CDbl(pnl_ValTot.Caption) < 0 Then pnl_ValTot.Caption = "0.00 "
   
   'Inicial
   Call fs_Calcular_MtoSBMS
   If cmb_BMSTas.ListIndex > -1 Then
      If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
         Call fs_Bono_Verde(modatecli_g_dbl_BMSTas * 100)
      Else
         Call fs_Bono_Verde(cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex))
      End If
   End If
   
   pnl_CuoIni.Caption = Format(CDbl(ipp_ApoPro.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption), "##,###,##0.00") & " "
   If CDbl(pnl_CuoIni.Caption) < 0 Then pnl_CuoIni.Caption = "0.00 "
   
   'Prestamo
   pnl_MtoPre.Caption = Format(CDbl(pnl_ValTot.Caption) - CDbl(pnl_CuoIni.Caption), "##,###,##0.00") & " "
   If CDbl(pnl_MtoPre.Caption) < 0 Then pnl_MtoPre.Caption = "0.00 "
   pnl_TotPre.Caption = Format(CDbl(pnl_MtoPre.Caption) + CDbl(pnl_ValGas.Caption), "##,###,##0.00") & " "
   DoEvents
End Sub

Private Sub fs_Calcular_GCierre()
   If chk_Gastos.Value = True Then
      pnl_ValGas.Caption = Format(CDbl(fs_Genera_Gastos_Cierre), "##,###,##0.00") & " "
      If CDbl(pnl_ValGas.Caption) < 0 Then pnl_ValGas.Caption = "0.00 "
      If CDbl(pnl_ValGas.Caption) = 0 And CDbl(ipp_ValEst.Value = 0) Then pnl_ValGas.Caption = "0.00 "
   Else
      pnl_ValGas.Caption = "0.00 "
   End If
   pnl_TotPre.Caption = Format(CDbl(pnl_MtoPre.Caption) + CDbl(pnl_ValGas.Caption), "##,###,##0.00") & " "
End Sub

Private Function fs_Genera_Gastos_Cierre() As Double
Dim r_dbl_GasTas        As Double
Dim r_dbl_GasNot        As Double
Dim r_dbl_BloReg        As Double
Dim r_dbl_GasMin        As Double
Dim r_dbl_GasMin_Inm    As Double
Dim r_dbl_GasMin_Est    As Double
Dim r_dbl_GasMin_Gar    As Double
Dim r_dbl_GasHip        As Double
Dim r_dbl_ValITF        As Double

   fs_Genera_Gastos_Cierre = 0
   r_dbl_GasTas = 0
   r_dbl_GasNot = 0
   r_dbl_GasMin_Est = 0
   r_dbl_GasMin_Inm = 0
   r_dbl_GasMin = 0
   
   'Gastos de Tasación
   r_dbl_GasTas = modgen_g_dbl_GasTas
    
   'Gastos Notariales
   r_dbl_GasNot = modgen_g_dbl_GasNot
   
   'Bloqueo Registral
   If modgen_g_int_TipBie > 0 Then
      If modgen_g_int_TipBie = 2 Then
         If ipp_ValEst.Value = 0 Then
            r_dbl_BloReg = modgen_g_dbl_inm_ficha
         Else
            r_dbl_BloReg = modgen_g_dbl_inm_ficha * 2
         End If
      Else
         r_dbl_BloReg = 0
      End If
   Else
      r_dbl_BloReg = 0
   End If
   
   'Gastos por registro de minuta
   If ipp_ComVta.Value > 0 Then
      r_dbl_GasMin_Inm = (ipp_ComVta.Value * modgen_g_dbl_inm_factor) + modgen_g_dbl_inm_ficha
   End If
   
   If ipp_ValEst.Value > 0 Then
      r_dbl_GasMin_Est = (ipp_ValEst.Value * modgen_g_dbl_est_factor) + modgen_g_dbl_est_ficha
   End If
   
   r_dbl_GasMin = r_dbl_GasMin_Inm + r_dbl_GasMin_Est + 37
   
   'Gastos por registro de Hipoteca
   r_dbl_GasMin_Gar = CDbl(CDbl(CDbl(CDbl(ipp_ComVta.Value) + CDbl(ipp_ValEst.Value)) * modgen_g_dbl_gar_factor) + modgen_g_dbl_gar_ficha) * 1.1
   
   r_dbl_GasHip = CDbl(r_dbl_GasMin_Gar * 1.15) + 37
   
   'Valor ITF
   r_dbl_ValITF = CDbl(pnl_MtoPre.Caption) * (0.005 / 100) * 2
   
   'Totaliza los gastos de cierre
   fs_Genera_Gastos_Cierre = r_dbl_GasNot + r_dbl_BloReg + r_dbl_GasMin + r_dbl_GasHip + r_dbl_ValITF
End Function

Private Sub chk_Gastos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub chk_Gastos_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub cmb_BMSTas_Click()
    If cmb_BMSTas.ListIndex > -1 Then
      Call gs_SetFocus(chk_Gastos)
   End If
End Sub

Private Sub cmb_BMSTas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BMSTas_Click
   End If
End Sub

Private Sub cmb_BMSTas_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub fs_Calcular_MtoSBMS()
   l_dbl_MPSMS = Format(CDbl(pnl_ValTot.Caption) - CDbl(ipp_ApoPro.Text) - (CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption)) - CDbl(ipp_MtoAFP.Text), "###,###,##0.00") & " " 'CDbl(ipp_ComVta.Text)
End Sub

'Private Sub fs_Calcular()
'   Call fs_Calcular_MtoSBMS
'
'   If cmb_BMSTas.ListIndex > -1 Then
'      If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
'         Call fs_Bono_Verde(modatecli_g_dbl_BMSTas * 100)
'      Else
'         Call fs_Bono_Verde(cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex))
'      End If
'   End If
'
'   If CDbl(ipp_ComVta.Value) - (CDbl(ipp_ApoPro.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption)) > 0 Then
'      pnl_MtoPre.Caption = Format(CDbl(ipp_ComVta.Value) - (CDbl(ipp_ApoPro.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption)), "##,###,##0.00") & " "
'   Else
'      pnl_MtoPre.Caption = "0.00 "
'   End If
'
'   pnl_CuoIni.Caption = Format(CDbl(ipp_ApoPro.Value) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption), "##,###,##0.00") & " "
'   If CDbl(pnl_CuoIni.Caption) < 0 Then
'      pnl_CuoIni.Caption = "0.00 "
'   End If
'End Sub

Private Sub cmb_DiaPag_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub cmb_TipSeg_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

'**************************************************************************************************
'********************************************* BOTONES ********************************************
'**************************************************************************************************
Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_Produc)
End Sub

Private Sub cmd_CalCuo_Click()
'variables antiguas
Dim r_arr_ParPrd()      As moddat_tpo_Genera
Dim r_int_TipVal_Des    As Integer
Dim r_int_TipVal_Viv    As Integer
Dim r_dbl_Import_Des    As Double
Dim r_dbl_Import_Viv    As Double
Dim r_dbl_Portes        As Double
Dim r_dbl_CuoRta        As Double
Dim r_dbl_PorCon        As Double
Dim r_dbl_TopCon        As Double
Dim r_dbl_ValMin_ComVta As Double
Dim r_dbl_ValMax_ComVta As Double
Dim r_dbl_PorMin_ApoPro As Double
Dim r_dbl_PorMax_ApoPro As Double
Dim r_dbl_PorMax_MtoPre As Double
Dim r_dbl_ValMin_MtoPre As Double
Dim r_dbl_ValMax_MtoPre As Double
Dim r_dbl_PrcMin        As Double
Dim r_dbl_PrcMax        As Double
Dim r_str_Moneda        As String
   
'variables nueva para la generacion del cronograma
Dim obj_Cronog          As Object
Dim int_Produc          As Integer
Dim int_CuoDbl          As Integer
Dim dbl_ValInm          As Double
Dim dbl_CuoIni          As Double
Dim dbl_MtoCon          As Double
Dim int_PlaPre          As Integer
Dim dbl_TasInt          As Double
Dim dbl_TasCof          As Double
Dim dbl_ComCof          As Double
Dim dat_FecDes          As Date
Dim int_DiaVct          As Integer
Dim int_PerGra          As Integer
Dim str_PriVct          As String
Dim dbl_Portes          As Double
Dim dbl_SegViv          As Double
Dim int_TipSDe          As Integer
Dim dbl_SegDes          As Double
Dim dbl_CuoMen          As Double
Dim dbl_CuoPbp          As Double
Dim dbl_IngReq          As Double
Dim str_CodCiu          As String

   'Valida ingreso de informacion
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   If cmb_SubPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sub-Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SubPrd)
      Exit Sub
   End If
   If CDbl(ipp_ComVta.Text) <= 0# Then
      MsgBox "Debe ingresar el Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   If CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Value) <= 0# Then
      MsgBox "Debe ingresar el Aporte Propio y/o el monto de AFP.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   If CDbl(ipp_ApoPro.Text) >= CDbl(ipp_ComVta.Text) Then
      MsgBox "La Cuota Inicial no puede ser mayor o igual al Valor del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   If CDbl(ipp_PlaAno.Text) <= 0# Then
      MsgBox "Debe ingresar el Plazo del prestamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If CDbl(ipp_PerGra.Text) < 0# Then
      MsgBox "Debe ingresar el Periodo de Gracia correctamente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If
   If cmb_CuoDbl.ListIndex = -1 Then
      MsgBox "Debe seleccionar si desea Cuotas Dobles.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CuoDbl)
      Exit Sub
   End If
   If cmb_TipSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipSeg)
      Exit Sub
   End If
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Día de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
      Exit Sub
   End If
   
   'Valida parametros minimo y maximo de plazo y gracia
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo está fuera del rango permitido (Entre " & CStr(ipp_PlaAno.MinValue) & " y " & CStr(ipp_PlaAno.MaxValue) & " años).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   If Not (CInt(ipp_PerGra.Text) >= ipp_PerGra.MinValue And CInt(ipp_PerGra.Text) <= ipp_PerGra.MaxValue) Then
      MsgBox "El Período de Gracia está fuera del rango permitido (Entre " & CStr(ipp_PerGra.MinValue) & " y " & CStr(ipp_PerGra.MaxValue) & " meses).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If
   
   'muestra datos
   pnl_MtoPre.Caption = Format(CDbl(pnl_ValTot.Caption) - (CDbl(ipp_ApoPro.Text) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption)), "###,##0.00") & " "
   'pnl_MtoPre.Caption = Format(CDbl(ipp_ComVta.Text) + CDbl(Me.ipp_ValEst.Value) - (CDbl(ipp_ApoPro.Text) + CDbl(pnl_FmvBbp.Caption) + CDbl(pnl_MefPbp.Caption) + CDbl(ipp_MtoAFP.Text) + CDbl(pnl_MtoBMS.Caption)), "###,##0.00") & " "

   'Recalcula datos del prestamo
'   Call fs_Calcular_Prestamo
'   Call fs_Calcular_GCierre
   
   If l_str_CodPrd = "023" Then
      pnl_PorIni.Caption = "(" & Format((CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption)) * 100, "##0.00") & "%) "  'CDbl(ipp_ComVta.Text)
   Else
      pnl_PorIni.Caption = "(" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Value)) / (CDbl(pnl_ValTot.Caption)) * 100, "##0.00") & "%) "  'CDbl(ipp_ComVta.Text) + CDbl(Me.ipp_ValEst.Value)
   End If
   r_str_Moneda = moddat_gf_Consulta_ParDes("229", CStr(l_int_TipMon))
   
   If l_int_TipMon = 2 Then
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(pnl_MtoPre.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_FmvBbp_Sol.Caption = Format(CDbl(pnl_FmvBbp.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_MefPbp_Sol.Caption = Format(CDbl(pnl_MefPbp.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_MtoAFP_Sol.Caption = Format(CDbl(ipp_MtoAFP.Text) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_MtoBMS_Sol.Caption = Format(CDbl(pnl_MtoBMS.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text) * l_dbl_TipCam, "###,##0.00") & " "
      'pnl_CuoIni_Sol.Caption = Format(CDbl(pnl_CuoIni.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_ValTot_Sol.Caption = Format(CDbl(pnl_ValTot.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_ValEst_Sol.Caption = Format(CDbl(ipp_ValEst.Text) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_ValGas_Sol.Caption = Format(CDbl(pnl_ValGas.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      pnl_TotPre_Sol.Caption = Format(CDbl(pnl_TotPre.Caption) * l_dbl_TipCam, "###,##0.00") & " "
   Else
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text), "###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(pnl_MtoPre.Caption), "###,##0.00") & " "
      pnl_FmvBbp_Sol.Caption = Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & " "
      pnl_MefPbp_Sol.Caption = Format(CDbl(pnl_MefPbp.Caption), "###,##0.00") & " "
      pnl_MtoAFP_Sol.Caption = Format(CDbl(ipp_MtoAFP.Text), "###,##0.00") & " "
      pnl_MtoBMS_Sol.Caption = Format(CDbl(pnl_MtoBMS.Caption), "###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text), "###,##0.00") & " "
      'pnl_CuoIni_Sol.Caption = Format(CDbl(pnl_CuoIni.Caption), "###,##0.00") & " "
      pnl_ValTot_Sol.Caption = Format(CDbl(pnl_ValTot.Caption), "###,##0.00") & " "
      pnl_ValEst_Sol.Caption = Format(CDbl(ipp_ValEst.Text), "###,##0.00") & " "
      pnl_ValGas_Sol.Caption = Format(CDbl(pnl_ValGas.Caption), "###,##0.00") & " "
      pnl_TotPre_Sol.Caption = Format(CDbl(pnl_TotPre.Caption), "###,##0.00") & " "
   End If
   pnl_CuoIni_Sol.Caption = Format(CDbl(pnl_ComVta_Sol.Caption) + CDbl(ipp_ValEst.Text) - CDbl(pnl_MtoPre_Sol.Caption), "###,##0.00") & " "
   
   'Obteniendo Parámetros para Validación de Montos
   r_dbl_ValMin_ComVta = 0
   r_dbl_ValMax_ComVta = 0
   r_dbl_PorMin_ApoPro = 0
   r_dbl_PorMax_ApoPro = 0
   r_dbl_PorMax_MtoPre = 0
   r_dbl_ValMin_MtoPre = 0
   r_dbl_ValMax_MtoPre = 0
   r_dbl_PrcMin = 0
   r_dbl_PrcMax = 0
   
   '******************************************
   'Obtiene Valor Máximo y Minimo del Inmueble
   Select Case l_str_CodPrd > 0
      'En Monto
      Case InStr(moddat_g_str_AgrTMIC, l_str_CodPrd)
         If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "021") Then
            r_dbl_ValMax_ComVta = r_arr_ParPrd(1).Genera_Cantid
         End If
      
      'En UIT
      Case InStr(moddat_g_str_AgrCME, l_str_CodPrd)
         If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = r_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = r_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      'En UIT (Mínimo y Máximo)
      Case InStr(moddat_g_str_AgrTFMV, l_str_CodPrd)
         If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = r_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = r_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMin_ComVta = Format(r_dbl_ValMin_ComVta, "########0.00")
            r_dbl_ValMax_ComVta = Format(r_dbl_ValMax_ComVta, "########0.00")
         End If
   End Select
   
   '**************************************
   'Obtiene % Mínimo de Aporte Propio
   If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "022") Then
      r_dbl_PorMin_ApoPro = r_arr_ParPrd(1).Genera_Cantid
   End If
   
   '*************************************
   'Obtiene % Máximo de Monto de Préstamo
   If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "023") Then
      r_dbl_PorMax_MtoPre = r_arr_ParPrd(1).Genera_Cantid
   End If
   
   '********************************
   'Obtiene Monto Máximo de Préstamo
   If InStr(moddat_g_str_Agr1MIC, l_str_CodPrd) > 0 Then
      'En Montos
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "024") Then
         r_dbl_ValMax_MtoPre = r_arr_ParPrd(1).Genera_Cantid
      End If
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "026") Then
         r_dbl_ValMin_MtoPre = r_arr_ParPrd(1).Genera_Cantid
      End If
   
   ElseIf InStr(moddat_g_str_AgrCME, l_str_CodPrd) > 0 Then
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "023") Then
         r_dbl_ValMax_MtoPre = r_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
   
   ElseIf InStr(moddat_g_str_AgrMIHG, l_str_CodPrd) > 0 Then
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = r_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = r_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
      End If
   
   ElseIf InStr(moddat_g_str_Agr1FMV, l_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, l_str_CodPrd) > 0 Then
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = r_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = r_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
         r_dbl_ValMin_MtoPre = Format(r_dbl_ValMin_MtoPre, "#######0.00")
         r_dbl_ValMax_MtoPre = Format(r_dbl_ValMax_MtoPre, "#######0.00")
      End If
   End If
   
   '*******************************
   'Validando Valor de Compra Venta
   If InStr(moddat_g_str_AgrTFMV, l_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, l_str_CodPrd) > 0 Then
      If CDbl(pnl_ValTot.Caption) < CDbl(r_dbl_ValMin_ComVta) Then
         MsgBox "El Valor de Compra-Venta no cubre el mínimo requerido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMin_ComVta, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      If CDbl(pnl_ValTot_Sol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMax_ComVta, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      
   ElseIf InStr(moddat_g_str_Agr1MIC, l_str_CodPrd) > 0 Then
      If CDbl(pnl_ValTot.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMax_ComVta, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      
   End If
   
   '**************************
   'Validando la Cuota Inicial
   If l_str_CodPrd = "023" Then
      If CDbl(Format((CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption)) * 100, "##0.000000")) < r_dbl_PorMin_ApoPro Then
         MsgBox "El Porcentaje de Cuota Inicial (" & Format((CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption)) * 100, "###0.000000") & "%) no cubre el mínimo permitido para el Producto (" & Format(r_dbl_PorMin_ApoPro, "##0.000000") & "%).", vbExclamation, modgen_g_str_NomPlt   'CDbl(ipp_ComVta.Text)
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   Else
      If CDbl(Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Text)) / (CDbl(pnl_ValTot.Caption)) * 100, "##0.000000")) < r_dbl_PorMin_ApoPro Then   'CDbl(ipp_ComVta.Text) + CDbl(ipp_ValEst.Text)
         MsgBox "El Porcentaje de Cuota Inicial (" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100, "###0.000000") & "%) no cubre el mínimo permitido para el Producto (" & Format(r_dbl_PorMin_ApoPro, "##0.000000") & "%).", vbExclamation, modgen_g_str_NomPlt  'CDbl(ipp_ComVta.Text)
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   '**************************
   'Validando financiamiento permitido
   If CDbl(Format(CDbl(pnl_MtoPre.Caption) / (CDbl(pnl_ValTot.Caption)) * 100, "##0.000000")) > r_dbl_PorMax_MtoPre Then   'CDbl(ipp_ComVta.Text) + CDbl(ipp_ValEst.Text)
      MsgBox "El Porcentaje de Financiamiento (" & Format(CDbl(pnl_MtoPre.Caption) / CDbl(pnl_ValTot.Caption) * 100, "##0.000000") & "%) excede el máximo permitido para el Producto (" & Format(r_dbl_PorMax_MtoPre, "##0.000000") & "%).", vbExclamation, modgen_g_str_NomPlt   'CDbl(ipp_ComVta.Text)
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If InStr(moddat_g_str_AgrTFMV, l_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, l_str_CodPrd) > 0 Then '"003" "004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      'Para obtener % Maximo de Aporte Propio
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "027") Then
         r_dbl_PorMax_ApoPro = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      If l_str_CodPrd = "023" Then
         If CDbl(Format((CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption)) * 100, "##0.000000")) > r_dbl_PorMax_ApoPro Then   'CDbl(ipp_ComVta.Text)
            MsgBox "El Porcentaje de Cuota Inicial (" & Format((CDbl(pnl_CuoIni.Caption) / CDbl(pnl_ValTot.Caption)) * 100, "###0.000000") & "%) sobrepasa el máximo permitido para el Producto (" & Format(r_dbl_PorMax_ApoPro, "##0.000000") & "%).", vbExclamation, modgen_g_str_NomPlt    ' CDbl(ipp_ComVta.Text)
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
         If CDbl(Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100, "##0.000000")) > r_dbl_PorMax_ApoPro Then  'CDbl(ipp_ComVta.Text)
            MsgBox "El Porcentaje de Cuota Inicial (" & Format((CDbl(ipp_ApoPro.Text) + CDbl(ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption) * 100, "###0.000000") & "%) sobrepasa el máximo permitido para el Producto (" & Format(r_dbl_PorMax_ApoPro, "##0.000000") & "%).", vbExclamation, modgen_g_str_NomPlt  'CDbl(ipp_ComVta.Text)
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   '***************************
   'Validando Monto de Préstamo
   If InStr(moddat_g_str_Agr1MIC, l_str_CodPrd) > 0 Then      '"002" "011"
      If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMin_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      If CDbl(pnl_MtoPre.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMax_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   
   ElseIf InStr(moddat_g_str_AgrCME, l_str_CodPrd) > 0 Then   '"003"
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMax_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   
   ElseIf InStr(moddat_g_str_AgrMIHG, l_str_CodPrd) > 0 Then  '"004"
      If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMin_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMax_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   
   ElseIf InStr(moddat_g_str_Agr1FMV, l_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, l_str_CodPrd) > 0 Then '"007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMin_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto (" & r_str_Moneda & " " & Format(r_dbl_ValMax_MtoPre, "###,##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   cmd_CalCuo.Enabled = False
   Screen.MousePointer = 11
   
'   str_CodCiu = "9999"
'   If cmb_TasEsp.ListIndex > -1 Then
'      If CInt(cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex)) = 2 Then
'         str_CodCiu = "7522"
'      End If
'   End If
   
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000003", Format(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex), "000"), l_int_TipMon, CDbl(pnl_MtoPre.Caption), r_int_TipVal_Des, r_dbl_Import_Des, cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex))
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000003", 0, l_int_TipMon, CDbl(pnl_ValTot.Caption), r_int_TipVal_Viv, r_dbl_Import_Viv, cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex))
   
   r_dbl_Portes = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   l_dbl_Portes = r_dbl_Portes
   
   pnl_SegDes.Caption = Format(r_dbl_Import_Des, "###,##0.000000") & " "
   pnl_SegInm.Caption = Format(r_dbl_Import_Viv, "###,##0.000000") & " "
  
   'Relación Cuota / Renta
   r_dbl_CuoRta = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), l_str_CodPrd, l_str_CodSub, "001", "013") Then
      r_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If

   l_dbl_TasMVi = 0
   l_dbl_ComCof = 0
   l_dbl_TasCof = 0
   
   If InStr(moddat_g_str_AgrCME, l_str_CodPrd) > 0 Then
      l_dbl_TasMVi = moddat_gf_ComMVi(l_str_CodPrd, 3, l_int_TipMon, ipp_PlaAno.Value)
      l_dbl_ComCof = moddat_gf_ComMVi(l_str_CodPrd, 4, l_int_TipMon, ipp_PlaAno.Value)
      l_dbl_TasCof = moddat_gf_ComMVi(l_str_CodPrd, 5, l_int_TipMon, ipp_PlaAno.Value)
   ElseIf InStr(moddat_g_str_AgrTFMV, l_str_CodPrd) > 0 Then
      l_dbl_ComCof = moddat_gf_ComMVi(l_str_CodPrd, 4, l_int_TipMon, ipp_PlaAno.Value)
      l_dbl_TasCof = moddat_gf_ComMVi(l_str_CodPrd, 5, l_int_TipMon, ipp_PlaAno.Value)
   End If
   
   '********************************************************************************************************
   '*************************** GENERACION DE CRONOGRAMAS SEGUN TIPO DE PRODUCTO ***************************
   '********************************************************************************************************
   Select Case l_str_CodPrd > 0
      'COMENTADO PORQUE EL PRODUCTO HA CADUCADO
      'Case InStr(moddat_g_str_AgrCME, l_str_CodPrd)
      '   'Para obtener porcentaje de TC
      '   r_dbl_PorCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "011") Then
      '      r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   'Para obtener tope de TC
      '   r_dbl_TopCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
      '      r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   'NUEVA rutina de generacion de cronogramas
      '   int_Produc = 1
      '   int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
      '   dbl_ValInm = CDbl(pnl_ValTot.Caption) + CDbl(pnl_ValGas.Caption)  'CDbl(ipp_ComVta.Text)
      '   dbl_CuoIni = CDbl(pnl_CuoIni.Caption)
      '   dbl_MtoCon = CDbl(pnl_MtoPre.Caption) * (r_dbl_PorCon / 100)
      '   If dbl_MtoCon > r_dbl_TopCon Then dbl_MtoCon = r_dbl_TopCon
      '   int_PlaPre = CInt(ipp_PlaAno.Text) * 12
      '   dbl_TasInt = CDbl(pnl_TasInt.Caption)
      '   dbl_TasCof = l_dbl_TasCof
      '   dbl_ComCof = l_dbl_ComCof
      '   dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
      '   int_DiaVct = CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)
      '   int_PerGra = CInt(ipp_PerGra.Text)
      '   str_PriVct = ""
      '   dbl_Portes = CDbl(r_dbl_Portes)
      '   dbl_SegViv = CDbl(pnl_SegInm.Caption)
      '   int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
      '   dbl_SegDes = CDbl(pnl_SegDes.Caption)
      '
      '   'Calculando cronogramas
      '   Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
      '   Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
      '
      '   'Mostrando Cronograma 1
      '   Call fs_Muestra_Cronograma1
      '
      '   'Mostrando Cronograma 2
      '   Call fs_Muestra_Cronograma2
      '
      '   dbl_CuoMen = 0
      '   dbl_CuoPbp = 0
      '   dbl_IngReq = 0
      '   Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, l_str_CodPrd, l_str_CodSub)
      '   pnl_CuoMen.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
      '   pnl_IngReq.Caption = Format(dbl_IngReq, "###,###,##0.00") & " "
      '   pnl_CuoSBP.Caption = Format(dbl_CuoPbp, "###,###,##0.00") & " "
      '   pnl_CuoPBP.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
      '   pnl_MtoPBP.Caption = Format(dbl_MtoCon, "###,###,##0.00") & " "
      '
      '   If l_int_TipMon <> 1 Then
      '      pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      '      pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption) * l_dbl_TipCam, "###,##0.00") & " "
      '   Else
      '      pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption), "###,##0.00") & " "
      '      pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption), "###,##0.00") & " "
      '   End If
      '
      '   'Calculando Costo Efectivo
      '   pnl_CosEfe.Caption = Format(gf_Calculo_CostoEfectivo(l_Arr_TNC_Cli(), l_dbl_TasInt, CDbl(pnl_TotPre.Caption)), "###,##0.00") & " "
      
      Case InStr(moddat_g_str_Agr1MIC, l_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = 0
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = CDbl(pnl_TasInt.Caption)
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(pnl_SegInm.Caption)
         int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
         dbl_SegDes = CDbl(pnl_SegDes.Caption)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, l_str_CodPrd, l_str_CodSub)
         pnl_CuoMen.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         pnl_IngReq.Caption = Format(dbl_IngReq, "###,###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption) * l_dbl_TipCam, "###,##0.00") & " "
            pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption) * l_dbl_TipCam, "###,##0.00") & " "
         Else
            pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption), "###,##0.00") & " "
            pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption), "###,##0.00") & " "
         End If
         
         pnl_CosEfe.Caption = Format(gf_Calculo_CostoEfectivo(l_Arr_TNC_Cli(), l_dbl_TasInt, CDbl(pnl_TotPre.Caption)), "###,##0.00") & " "
         
         
      Case InStr(moddat_g_str_AgrMIHG, l_str_CodPrd) Or InStr(moddat_g_str_Agr2MIC, l_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, l_str_CodPrd)   '"004", "006", "007", "009", "010", "013", "014", "015", "016", "017", "018"
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If CDbl(pnl_ComVta_Sol.Caption) > (50 * moddat_gf_Consulta_ParVal("001", "002")) Then
            r_dbl_TopCon = 5000
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = r_dbl_TopCon
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = CDbl(pnl_TasInt.Caption)
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(pnl_SegInm.Caption)
         int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
         dbl_SegDes = CDbl(pnl_SegDes.Caption)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         'Mostrando Cronograma 2
         Call fs_Muestra_Cronograma2
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, l_str_CodPrd, l_str_CodSub)
         pnl_CuoMen.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         pnl_IngReq.Caption = Format(dbl_IngReq, "###,###,##0.00") & " "
         pnl_CuoSBP.Caption = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         pnl_CuoPBP.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         pnl_MtoPBP.Caption = Format(dbl_MtoCon, "###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption) * l_dbl_TipCam, "###,##0.00") & " "
            pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption) * l_dbl_TipCam, "###,##0.00") & " "
         Else
            pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption), "###,##0.00") & " "
            pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption), "###,##0.00") & " "
         End If
         
         'Calculando Costo Efectivo
         pnl_CosEfe.Caption = Format(gf_Calculo_CostoEfectivo(l_Arr_TNC_Cli(), l_dbl_TasInt, CDbl(pnl_TotPre.Caption)), "###,##0.00") & " "
      
      
      Case InStr(moddat_g_str_Agr1FMV, l_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = CInt(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex))
         dbl_ValInm = CDbl(pnl_ValTot.Caption)
         dbl_CuoIni = CDbl(pnl_CuoIni.Caption) - CDbl(pnl_ValGas.Caption)
         dbl_MtoCon = 0
         int_PlaPre = CInt(ipp_PlaAno.Text) * 12
         dbl_TasInt = CDbl(pnl_TasInt.Caption)
         dbl_TasCof = l_dbl_TasCof
         dbl_ComCof = l_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)
         int_PerGra = CInt(ipp_PerGra.Text)
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = CDbl(pnl_SegInm.Caption)
         int_TipSDe = CInt(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex)) - 10
         dbl_SegDes = CDbl(pnl_SegDes.Caption)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, 0, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         'Mostrando Cronograma 1
         Call fs_Muestra_Cronograma1
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, l_str_CodPrd, l_str_CodSub)
         pnl_CuoMen.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         pnl_IngReq.Caption = Format(dbl_IngReq, "###,###,##0.00") & " "
         pnl_CuoSBP.Caption = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         pnl_CuoPBP.Caption = Format(dbl_CuoMen, "###,###,##0.00") & " "
         pnl_MtoPBP.Caption = Format(dbl_MtoCon, "###,##0.00") & " "
         
         If l_int_TipMon <> 1 Then
            pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption) * l_dbl_TipCam, "###,##0.00") & " "
            pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption) * l_dbl_TipCam, "###,##0.00") & " "
         Else
            pnl_CuoMen_Sol.Caption = Format(CDbl(pnl_CuoMen.Caption), "###,##0.00") & " "
            pnl_IngReq_Sol.Caption = Format(CDbl(pnl_IngReq.Caption), "###,##0.00") & " "
         End If
         
         'Calculando Costo Efectivo
         pnl_CosEfe.Caption = Format(gf_Calculo_CostoEfectivo(l_Arr_TNC_Cli(), l_dbl_TasInt, CDbl(pnl_TotPre.Caption)), "###,##0.00") & " "
   End Select
   
   Screen.MousePointer = 0
   cmd_CalCuo.Enabled = True
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_FecImp  As String
   Dim r_str_HorImp  As String
   
   If grd_Listad_NCo.Rows = 0 Then
      MsgBox "Debe realizar algún cálculo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir la Hoja de Simulación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   r_str_FecImp = Format(date, "yyyymmdd")
   r_str_HorImp = Format(Time, "hhmmss")
   
   'Generamos la cadena con los campos para compararlo en la BD si es que ya existe
   g_str_Parame = "SELECT * FROM RPT_SIMCRE WHERE "
   g_str_Parame = g_str_Parame & "SIMCRE_FECCRE = " & r_str_FecImp & " AND "
   g_str_Parame = g_str_Parame & "SIMCRE_HORCRE = " & r_str_HorImp & " AND "
   g_str_Parame = g_str_Parame & "SIMCRE_TERCRE = '" & modgen_g_str_NombPC & "' "
      
   'Condicion si No se ejecuta la sentencia SQL
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   'Condicion si No se encuentra al comienzo o al final del archivo y lo evalua
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      'Cerramos la conexion a la BD
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
         
      'Si ya se encuentra en la BD se procede a eliminar
      g_str_Parame = "USP_RPT_SIMCRE_BORRAR (" & "'" & r_str_FecImp & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_HorImp & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
   End If
      
   'Cerramos la conexion a la BD
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
                        
   'Se llama al procedure y se ejecuta el ingreso de la data en la base de datos
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_RPT_SIMCRE ("
      g_str_Parame = g_str_Parame & r_str_FecImp & ", "
      g_str_Parame = g_str_Parame & r_str_HorImp & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon) & ", "
      g_str_Parame = g_str_Parame & CDbl(pnl_ValTot.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_ComVta.Value) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_ValEst.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotPre.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ValGas.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPBP.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value) & ", "
      g_str_Parame = g_str_Parame & CStr((ipp_PlaAno.Value * 12)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PerGra.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo & ", "
      g_str_Parame = g_str_Parame & "'" & cmb_TipSeg.Text & "', "
      g_str_Parame = g_str_Parame & CStr(pnl_TasInt.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_SegDes.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_SegInm.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Portes) & ", "
      
      If InStr(moddat_g_str_Agr1MIC, l_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, l_str_CodPrd) > 0 Then  '"002" "011" "019" "021" "022" "023"
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMen.Caption)) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoSBP.Caption)) & ", "
      End If
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoPBP.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngReq.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngReq_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_TipCam.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_CosEfe.Caption) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
                                    
      'Se genera el mensaje de error por la concurrencia que exista
      If moddat_g_int_CntErr = 5 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "RPT_SIMCRE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "CRE_SUBPRD"
   crp_Imprim.DataFiles(3) = ""
   
   'Se selecciona la formula con el tipo de producto
   crp_Imprim.SelectionFormula = "{RPT_SIMCRE.SIMCRE_FECCRE} = " & r_str_FecImp & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SIMCRE.SIMCRE_HORCRE} = " & r_str_HorImp & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SIMCRE.SIMCRE_TERCRE} = '" & modgen_g_str_NombPC & "'"
   
   'Se pregunta para saber que codigo mostrará la data en su respectivo reporte
   If l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "002" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "001" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "011" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "019" Then
      crp_Imprim.Formulas(0) = "F_ApoPro = '" & Format(ipp_ApoPro.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(1) = "F_MtoBBP = '" & Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & "'"
      crp_Imprim.Formulas(2) = "F_MtoAFP = '" & Format(ipp_MtoAFP.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(3) = "F_MtoBMS = '" & Format(pnl_MtoBMS.Caption, "###,##0.00") & "'"
      If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
         If cmb_BMSTas.ListIndex > -1 Then
            If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & modatecli_g_dbl_BMSTas * 100 & "%)" & "'"
            Else
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex) & "%)" & "'"
            End If
         End If
      End If
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_13.RPT"
   ElseIf l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "003" Then
      crp_Imprim.Formulas(0) = "F_ApoPro = '" & Format(ipp_ApoPro.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(1) = "F_MtoBBP = '" & Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & "'"
      crp_Imprim.Formulas(2) = "F_MtoAFP = '" & Format(ipp_MtoAFP.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(3) = "F_MtoBMS = '" & Format(pnl_MtoBMS.Caption, "###,##0.00") & "'"
      If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
         If cmb_BMSTas.ListIndex > -1 Then
            If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & modatecli_g_dbl_BMSTas * 100 & "%)" & "'"
            Else
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex) & "%)" & "'"
            End If
         End If
      End If
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_14.RPT"
   ElseIf l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "004" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "006" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "007" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "009" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "010" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "012" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "013" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "014" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "015" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "016" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "017" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "018" Then
      crp_Imprim.Formulas(0) = "F_ApoPro = '" & Format(ipp_ApoPro.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(1) = "F_MtoBBP = '" & Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & "'"
      crp_Imprim.Formulas(2) = "F_MtoAFP = '" & Format(ipp_MtoAFP.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(3) = "F_MtoBMS = '" & Format(pnl_MtoBMS.Caption, "###,##0.00") & "'"
      
      If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Or InStr(moddat_g_str_Agr2MIC, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
         If cmb_BMSTas.ListIndex > -1 Then
            If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & modatecli_g_dbl_BMSTas * 100 & "%)" & "'"
            Else
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex) & "%)" & "'"
            End If
         End If
      End If
      crp_Imprim.Formulas(5) = "F_SimMon = 'S/' "
      crp_Imprim.Formulas(6) = "F_SalFin = 0 "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_15.RPT"
   ElseIf l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "021" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "022" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "023" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "024" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "025" Then
      crp_Imprim.Formulas(0) = "Formula20 = '" & Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & "'"
      crp_Imprim.Formulas(1) = "F_CuoIni = '" & Format(CDbl(pnl_CuoIni.Caption), "###,##0.00") & "'"
      
      If l_str_CodPrd = "023" Then
         crp_Imprim.Formulas(2) = "F_PorIni = '" & Format(100 * (CDbl(pnl_CuoIni.Caption)) / CDbl(pnl_ValTot.Caption), "###,##0.00") & "'"
      Else
         crp_Imprim.Formulas(2) = "F_PorIni = '" & Format(100 * (CDbl(ipp_ApoPro.Text) + CDbl(Me.ipp_MtoAFP.Text)) / CDbl(pnl_ValTot.Caption), "###,##0.00") & "'"  'ipp_ComVta.Text
      End If
      
      crp_Imprim.Formulas(3) = "F_ApoPro = '" & Format(ipp_ApoPro.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(4) = "F_MtoAFP = '" & Format(ipp_MtoAFP.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(5) = "F_MtoBMS = '" & Format(pnl_MtoBMS.Caption, "###,##0.00") & "'"
         
      If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
         If cmb_BMSTas.ListIndex > -1 Then
            If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
               crp_Imprim.Formulas(6) = "F_BMSTas = '" & "(" & modatecli_g_dbl_BMSTas * 100 & "%)" & "'"
            Else
               crp_Imprim.Formulas(6) = "F_BMSTas = '" & "(" & cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex) & "%)" & "'"
            End If
         End If
      End If
      crp_Imprim.Formulas(7) = "F_MtoPBP = '" & Format(pnl_MefPbp.Caption, "###,##0.00") & "'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_21.RPT"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_ImpCro_Click()
Dim r_str_FecImp  As String
Dim r_str_HorImp  As String
Dim r_int_Contad  As Integer
Dim r_int_NumCuo  As Integer
Dim r_str_FecVct  As String
Dim r_dbl_Capita  As Double
Dim r_dbl_Intere  As Double
Dim r_dbl_SegDes  As Double
Dim r_dbl_SegInm  As Double
Dim r_dbl_Portes  As Double
Dim r_dbl_TotCuo  As Double
Dim r_dbl_SalCap  As Double

   If grd_Listad_NCo.Rows = 0 Then
      MsgBox "Debe realizar algún cálculo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir la Hoja de Simulación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   r_str_FecImp = Format(date, "yyyymmdd")
   r_str_HorImp = Format(Time, "hhmmss")
   
   'Generamos la cadena con los campos para compararlo en la BD si es que ya existe
   g_str_Parame = "SELECT * FROM RPT_SIMCRE WHERE "
   g_str_Parame = g_str_Parame & "SIMCRE_FECCRE = " & r_str_FecImp & " AND "
   g_str_Parame = g_str_Parame & "SIMCRE_HORCRE = " & r_str_HorImp & " AND "
   g_str_Parame = g_str_Parame & "SIMCRE_TERCRE = '" & modgen_g_str_NombPC & "' "
   
   'Condicion si No se ejecuta la sentencia SQL
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   'Condicion si No se encuentra al comienzo o al final del archivo y lo evalua
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      'Cerramos la conexion a la BD
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
         
      'Si ya se encuentra en la BD se procede a eliminar
      g_str_Parame = "USP_RPT_SIMCRE_BORRAR (" & "'" & r_str_FecImp & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_HorImp & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
   End If
      
   'Cerramos la conexion a la BD
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
                        
   'Se llama al procedure y se ejecuta el ingreso de la data en la base de datos
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_RPT_SIMCRE ("
      g_str_Parame = g_str_Parame & r_str_FecImp & ", "
      g_str_Parame = g_str_Parame & r_str_HorImp & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon) & ", "
      g_str_Parame = g_str_Parame & CDbl(pnl_ValTot.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_ComVta.Value) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_ValEst.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TotPre.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ValGas.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPre.Caption)) & ","
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_MtoPBP.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaAno.Value) & ", "
      g_str_Parame = g_str_Parame & CStr((ipp_PlaAno.Value * 12)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PerGra.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_CuoDbl.ItemData(cmb_CuoDbl.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo & ", "
      g_str_Parame = g_str_Parame & "'" & cmb_TipSeg.Text & "', "
      g_str_Parame = g_str_Parame & CStr(pnl_TasInt.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_SegDes.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_SegInm.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_Portes) & ", "
      
      If InStr(moddat_g_str_Agr1MIC, l_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, l_str_CodPrd) > 0 Then '"002" "011" "019" "020" "021" "022" "023"
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoMen.Caption)) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoSBP.Caption)) & ", "
      End If
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CuoPBP.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngReq.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_IngReq_Sol.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_TipCam.Caption) & ", "
      g_str_Parame = g_str_Parame & CStr(pnl_CosEfe.Caption) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
                                    
      'Se genera el mensaje de error por la concurrencia que exista
      If moddat_g_int_CntErr = 5 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Cronograma No Concesional
   If tab_Cronog.Tab = 0 Then
      grd_Listad_NCo.Redraw = False
      For r_int_Contad = 0 To grd_Listad_NCo.Rows - 1
         grd_Listad_NCo.Row = r_int_Contad
         
         grd_Listad_NCo.Col = 0
         r_int_NumCuo = CInt(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 1
         r_str_FecVct = grd_Listad_NCo.Text
      
         grd_Listad_NCo.Col = 2
         r_dbl_Capita = CDbl(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 3
         r_dbl_Intere = CDbl(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 4
         r_dbl_SegDes = CDbl(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 5
         r_dbl_SegInm = CDbl(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 6
         r_dbl_Portes = CDbl(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 7
         r_dbl_TotCuo = CDbl(grd_Listad_NCo.Text)
      
         grd_Listad_NCo.Col = 8
         r_dbl_SalCap = CDbl(grd_Listad_NCo.Text)
         
         'Si ya se encuentra en la BD se procede a eliminar
         g_str_Parame = "INSERT INTO RPT_SIMCUO ("
         g_str_Parame = g_str_Parame & "SIMCUO_FECCRE, "
         g_str_Parame = g_str_Parame & "SIMCUO_HORCRE, "
         g_str_Parame = g_str_Parame & "SIMCUO_TERCRE, "
         g_str_Parame = g_str_Parame & "SIMCUO_NUMCUO, "
         g_str_Parame = g_str_Parame & "SIMCUO_FECVCT, "
         g_str_Parame = g_str_Parame & "SIMCUO_CAPITA, "
         g_str_Parame = g_str_Parame & "SIMCUO_INTERE, "
         g_str_Parame = g_str_Parame & "SIMCUO_SEGDES, "
         g_str_Parame = g_str_Parame & "SIMCUO_SEGINM, "
         g_str_Parame = g_str_Parame & "SIMCUO_PORTES, "
         g_str_Parame = g_str_Parame & "SIMCUO_TOTCUO, "
         g_str_Parame = g_str_Parame & "SIMCUO_SALCAP) "
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & r_str_FecImp & ", "
         g_str_Parame = g_str_Parame & r_str_HorImp & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_NumCuo) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_FecVct & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Capita) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Intere) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_SegDes) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_SegInm) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Portes) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_TotCuo) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_SalCap) & ") "
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
      Next r_int_Contad
      
      grd_Listad_NCo.Redraw = True
      Call gs_UbiIniGrid(grd_Listad_NCo)
   Else
   
      grd_Listad_Con.Redraw = True
      Call gs_UbiIniGrid(grd_Listad_Con)
      grd_Listad_Con.Redraw = False
   
      For r_int_Contad = 0 To grd_Listad_Con.Rows - 1
         grd_Listad_Con.Row = r_int_Contad
         
         grd_Listad_Con.Col = 0
         r_int_NumCuo = CInt(grd_Listad_Con.Text)
      
         grd_Listad_Con.Col = 1
         r_str_FecVct = grd_Listad_Con.Text
      
         grd_Listad_Con.Col = 2
         r_dbl_Capita = CDbl(grd_Listad_Con.Text)
      
         grd_Listad_Con.Col = 3
         r_dbl_Intere = CDbl(grd_Listad_Con.Text)
      
         grd_Listad_Con.Col = 4
         r_dbl_TotCuo = CDbl(grd_Listad_Con.Text)
      
         grd_Listad_Con.Col = 5
         r_dbl_SalCap = CDbl(grd_Listad_Con.Text)
         
         'Si ya se encuentra en la BD se procede a eliminar
         g_str_Parame = "INSERT INTO RPT_SIMCUO ("
         g_str_Parame = g_str_Parame & "SIMCUO_FECCRE, "
         g_str_Parame = g_str_Parame & "SIMCUO_HORCRE, "
         g_str_Parame = g_str_Parame & "SIMCUO_TERCRE, "
         g_str_Parame = g_str_Parame & "SIMCUO_NUMCUO, "
         g_str_Parame = g_str_Parame & "SIMCUO_FECVCT, "
         g_str_Parame = g_str_Parame & "SIMCUO_CAPITA, "
         g_str_Parame = g_str_Parame & "SIMCUO_INTERE, "
         g_str_Parame = g_str_Parame & "SIMCUO_SEGDES, "
         g_str_Parame = g_str_Parame & "SIMCUO_SEGINM, "
         g_str_Parame = g_str_Parame & "SIMCUO_PORTES, "
         g_str_Parame = g_str_Parame & "SIMCUO_TOTCUO, "
         g_str_Parame = g_str_Parame & "SIMCUO_SALCAP) "
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & r_str_FecImp & ", "
         g_str_Parame = g_str_Parame & r_str_HorImp & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_NumCuo) & ", "
         g_str_Parame = g_str_Parame & "'" & r_str_FecVct & "', "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Capita) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Intere) & ", "
         g_str_Parame = g_str_Parame & CStr(0) & ", "
         g_str_Parame = g_str_Parame & CStr(0) & ", "
         g_str_Parame = g_str_Parame & CStr(0) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_TotCuo) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_SalCap) & ") "
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
      Next r_int_Contad
      
      grd_Listad_Con.Redraw = True
      Call gs_UbiIniGrid(grd_Listad_Con)
   End If
   
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "RPT_SIMCRE"
   crp_Imprim.DataFiles(1) = "CRE_PRODUC"
   crp_Imprim.DataFiles(2) = "RPT_SIMCUO"
   crp_Imprim.DataFiles(3) = "CRE_SUBPRD"
   
   'Se selecciona la formula con el tipo de producto
   crp_Imprim.SelectionFormula = "{RPT_SIMCRE.SIMCRE_FECCRE} = " & r_str_FecImp & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SIMCRE.SIMCRE_HORCRE} = " & r_str_HorImp & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_SIMCRE.SIMCRE_TERCRE} = '" & modgen_g_str_NombPC & "'"
   
   'Se pregunta para saber que codigo mostrará la data en su respectivo reporte
   If tab_Cronog.Tab = 0 Then
      crp_Imprim.Formulas(0) = "F_ApoPro = '" & Format(ipp_ApoPro.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(1) = "F_MtoBBP = '" & Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & "'"
      crp_Imprim.Formulas(2) = "F_MtoAFP = '" & Format(ipp_MtoAFP.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(3) = "F_MtoBMS = '" & Format(pnl_MtoBMS.Caption, "###,##0.00") & "'"
      
      If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
         If cmb_BMSTas.ListIndex > -1 Then
            If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & modatecli_g_dbl_BMSTas * 100 & "%)" & "'"
            Else
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex) & "%)" & "'"
            End If
         End If
      End If
      crp_Imprim.Formulas(5) = "F_MtoPBP = '" & Format(CDbl(pnl_MefPbp.Caption), "###,##0.00") & "'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_16.RPT"
   Else
      crp_Imprim.Formulas(0) = "F_ApoPro = '" & Format(ipp_ApoPro.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(1) = "F_MtoBBP = '" & Format(CDbl(pnl_FmvBbp.Caption), "###,##0.00") & "'"
      crp_Imprim.Formulas(2) = "F_MtoAFP = '" & Format(ipp_MtoAFP.Text, "###,##0.00") & "'"
      crp_Imprim.Formulas(3) = "F_MtoBMS = '" & Format(pnl_MtoBMS.Caption, "###,##0.00") & "'"
      If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
         If cmb_BMSTas.ListIndex > -1 Then
            If CDbl(l_dbl_MPSMS) <= modatecli_g_dbl_MtoFin And cmb_BMSTas.ListIndex <> 0 Then
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & modatecli_g_dbl_BMSTas * 100 & "%)" & "'"
            Else
               crp_Imprim.Formulas(4) = "F_BMSTas = '" & "(" & cmb_BMSTas.ItemData(cmb_BMSTas.ListIndex) & "%)" & "'"
            End If
         End If
      End If
      crp_Imprim.Formulas(5) = "F_MtoPBP = '" & Format(CDbl(pnl_MefPbp.Caption), "###,##0.00") & "'"
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "COM_SIMCRE_17.RPT"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_CalEst_Click()
   cmd_CalEst.Enabled = False

   Screen.MousePointer = 11
   Call fs_Calcul_MtoMax
   Screen.MousePointer = 0
   
   cmd_CalEst.Enabled = True
End Sub

'**************************************************************************************************
'******************************************* FORMULARIO *******************************************
'**************************************************************************************************
Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Produc)
   Screen.MousePointer = 0
End Sub

'**************************************************************************************************
'***************************************** PROCEDIMIENTOS *****************************************
'**************************************************************************************************
Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc_Comerc(cmb_Produc, l_arr_Produc, 4)
   Call moddat_gs_Carga_TipSeg(cmb_TipSeg, "000001")
   
   'Cargando Tipo de Ingreso
   cmb_TipIng.Clear
   cmb_TipIng.AddItem "INDIVIDUAL"
   cmb_TipIng.ItemData(cmb_TipIng.NewIndex) = 1
   cmb_TipIng.AddItem "CONYUGAL"
   cmb_TipIng.ItemData(cmb_TipIng.NewIndex) = 2
   cmb_TipIng.ListIndex = -1
   
   cmb_BMSTas.Clear
   cmb_BMSTas.AddItem "S/BMS"
   cmb_BMSTas.ItemData(cmb_BMSTas.NewIndex) = 0
   cmb_BMSTas.AddItem "G1-3%"
   cmb_BMSTas.ItemData(cmb_BMSTas.NewIndex) = 3
   cmb_BMSTas.AddItem "G2-4%"
   cmb_BMSTas.ItemData(cmb_BMSTas.NewIndex) = 4
   cmb_BMSTas.ListIndex = -1
   
   'Cargando Tasa Especial
   Call moddat_gs_Carga_LisIte_Combo(cmb_TasEsp, 1, "522")
   
   'Cuotas Dobles
   Call moddat_gs_Carga_LisIte_Combo(cmb_CuoDbl, 1, "277")

   grd_Listad_NCo.ColWidth(0) = 1110
   grd_Listad_NCo.ColWidth(1) = 1380
   grd_Listad_NCo.ColWidth(2) = 1395
   grd_Listad_NCo.ColWidth(3) = 1390
   grd_Listad_NCo.ColWidth(4) = 1385
   grd_Listad_NCo.ColWidth(5) = 1380
   grd_Listad_NCo.ColWidth(6) = 1370
   grd_Listad_NCo.ColWidth(7) = 1260
   grd_Listad_NCo.ColWidth(8) = 1410
   grd_Listad_NCo.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_NCo.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_NCo.ColAlignment(2) = flexAlignRightCenter
   grd_Listad_NCo.ColAlignment(3) = flexAlignRightCenter
   grd_Listad_NCo.ColAlignment(4) = flexAlignRightCenter
   grd_Listad_NCo.ColAlignment(5) = flexAlignRightCenter
   grd_Listad_NCo.ColAlignment(6) = flexAlignRightCenter
   grd_Listad_NCo.ColAlignment(7) = flexAlignRightCenter
   grd_Listad_NCo.ColAlignment(8) = flexAlignRightCenter
   
   grd_Listad_Con.ColWidth(0) = 1620
   grd_Listad_Con.ColWidth(1) = 1830
   grd_Listad_Con.ColWidth(2) = 1830
   grd_Listad_Con.ColWidth(3) = 1830
   grd_Listad_Con.ColWidth(4) = 1830
   grd_Listad_Con.ColWidth(5) = 1830
   grd_Listad_Con.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_Con.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_Con.ColAlignment(2) = flexAlignRightCenter
   grd_Listad_Con.ColAlignment(3) = flexAlignRightCenter
   grd_Listad_Con.ColAlignment(4) = flexAlignRightCenter
   grd_Listad_Con.ColAlignment(5) = flexAlignRightCenter
   
   'Bono - Producto FMV MAS BBP y BBP COMPLEMENTO INICIAL
   pnl_FmvBbp.Caption = "0.00 "
   pnl_MefPbp.Caption = "0.00 "
'   If InStr(moddat_g_str_Agr1FMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 And l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo <> "019" Then
'      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo, l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo, "051", "012") Then
'         pnl_FmvBbp.Caption = Format(l_arr_ParPrd(1).Genera_Cantid, "###,###,##0.00") & " "
'      End If
'   End If
End Sub

Private Sub fs_Limpia()
Dim r_int_Contad     As Integer
   
   cmb_Produc.ListIndex = -1
   cmb_SubPrd.Clear
   cmb_BMSTas.ListIndex = -1
   chk_Gastos.Value = 0
   
   pnl_ValTot.Caption = "0.00 "
   ipp_ComVta.Value = 0
   ipp_ValEst.Value = 0
   pnl_CuoIni.Caption = "0.00 "
   ipp_ApoPro.Value = 0
   pnl_FmvBbp.Caption = "0.00 "
   pnl_MefPbp.Caption = "0.00 "
   ipp_MtoAFP.Value = 0
   pnl_MtoBMS.Caption = "0.00 "
   pnl_MtoPre.Caption = "0.00 "
   pnl_ValGas.Caption = "0.00 "
   pnl_TotPre.Caption = "0.00 "
   
   pnl_ValTot_Sol.Caption = "0.00 "
   pnl_ComVta_Sol.Caption = "0.00 "
   pnl_ValEst_Sol.Caption = "0.00 "
   pnl_CuoIni_Sol.Caption = "0.00 "
   pnl_ApoPro_Sol.Caption = "0.00 "
   pnl_FmvBbp_Sol.Caption = "0.00 "
   pnl_MefPbp_Sol.Caption = "0.00 "
   pnl_MtoAFP_Sol.Caption = "0.00 "
   pnl_MtoBMS_Sol.Caption = "0.00 "
   pnl_MtoPre_Sol.Caption = "0.00 "
   pnl_ValGas_Sol.Caption = "0.00 "
   pnl_TotPre_Sol.Caption = "0.00 "
   
   ipp_PlaAno.Value = 0
   ipp_PerGra.Value = 0
   cmb_CuoDbl.ListIndex = -1
   cmb_TipSeg.ListIndex = -1
   cmb_TasEsp.ListIndex = -1
   cmb_DiaPag.Clear
   
   pnl_CuoMen.Caption = "0.00 "
   pnl_IngReq.Caption = "0.00 "
   pnl_CuoMen_Sol.Caption = "0.00 "
   pnl_IngReq_Sol.Caption = "0.00 "
   pnl_TipCam.Caption = "0.0000 "
   pnl_TasInt.Caption = "0.00 "
   pnl_SegDes.Caption = "0.000000 "
   pnl_SegInm.Caption = "0.000000 "
   pnl_CosEfe.Caption = "0.00 "
   pnl_CuoSBP.Caption = "0.00 "
   pnl_MtoPBP.Caption = "0.00 "
   pnl_CuoPBP.Caption = "0.00 "
   
   For r_int_Contad = 0 To lbl_SimMon.Count - 1
      lbl_SimMon(r_int_Contad).Caption = ""
   Next r_int_Contad
   
   lbl_Totale(0).Caption = "Totales ==> "
   lbl_Totale(1).Caption = "Totales ==> "
   tab_Cronog.TabCaption(0) = "Cronograma de Pagos"
   tab_Cronog.TabVisible(1) = False
   
   'Limpiando Grid TNC
   Call gs_LimpiaGrid(grd_Listad_NCo)
   pnl_Tot_Capita_NCo.Caption = "0.00 "
   pnl_Tot_Intere_NCo.Caption = "0.00 "
   pnl_Tot_SegPre_NCo.Caption = "0.00 "
   pnl_Tot_SegViv_NCo.Caption = "0.00 "
   pnl_Tot_OtrCar_NCo.Caption = "0.00 "
   pnl_Tot_TotCuo_NCo.Caption = "0.00 "

   'Limpiando Grid TC
   Call gs_LimpiaGrid(grd_Listad_Con)
   pnl_Tot_Capita_Con.Caption = "0.00 "
   pnl_Tot_Intere_Con.Caption = "0.00 "
   pnl_Tot_TotCuo_Con.Caption = "0.00 "
   
   ipp_IngNet.Value = 0
   cmb_TipIng.ListIndex = -1
   pnl_MtoMax.Caption = "0.00 "
   pnl_CuoApr.Caption = "0.00 "
End Sub

Private Sub fs_Bono_Verde(ByVal p_ValAfe As Double)
   p_ValAfe = p_ValAfe / 100
   
   If InStr(moddat_g_str_AgrTFMV, l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo) > 0 Then
      If cmb_BMSTas.ListIndex > -1 Then
         pnl_MtoBMS.Caption = Format((CDbl(l_dbl_MPSMS) * p_ValAfe) / (1 + p_ValAfe), "###,###,##0.00") & " "
         If pnl_MtoBMS.Caption < 0 Then pnl_MtoBMS.Caption = "0.00 "
      End If
      pnl_MtoPre.Caption = Format(CDbl(l_dbl_MPSMS) - CDbl(pnl_MtoBMS.Caption), "###,###,##0.00") & " "
      If pnl_MtoPre.Caption < 0 Then pnl_MtoPre.Caption = "0.00 "
   Else
      pnl_MtoBMS.Caption = "0.00 "
   End If
End Sub

Private Sub fs_Limpia_Calcul()
   pnl_CosEfe.Caption = "0.00 "
   pnl_SegDes.Caption = "0.00 "
   pnl_SegInm.Caption = "0.00 "
   pnl_CuoMen.Caption = "0.00 "
   pnl_IngReq.Caption = "0.00 "
   pnl_CuoMen_Sol.Caption = "0.00 "
   pnl_IngReq_Sol.Caption = "0.00 "
   pnl_CuoSBP.Caption = "0.00 "
   pnl_CuoPBP.Caption = "0.00 "
   pnl_MtoPBP.Caption = "0.00 "

   'Limpiando Grid TNC
   Call gs_LimpiaGrid(grd_Listad_NCo)
   pnl_Tot_Capita_NCo.Caption = "0.00 "
   pnl_Tot_Intere_NCo.Caption = "0.00 "
   pnl_Tot_SegPre_NCo.Caption = "0.00 "
   pnl_Tot_SegViv_NCo.Caption = "0.00 "
   pnl_Tot_OtrCar_NCo.Caption = "0.00 "
   pnl_Tot_TotCuo_NCo.Caption = "0.00 "

   'Limpiando Grid TC
   Call gs_LimpiaGrid(grd_Listad_Con)
   pnl_Tot_Capita_Con.Caption = "0.00 "
   pnl_Tot_Intere_Con.Caption = "0.00 "
   pnl_Tot_TotCuo_Con.Caption = "0.00 "
End Sub

Private Function fs_Obtiene_MontoBBP(ByVal p_CodPrd As Integer, ByVal p_CodSub As Integer, ByVal p_CodIte As String) As Double
Dim r_str_Parame     As String
Dim r_rst_MtoBBP     As ADODB.Recordset

   fs_Obtiene_MontoBBP = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT * FROM CRE_PARPRD "
   r_str_Parame = r_str_Parame & " WHERE PARPRD_CODPRD = '" & Format(p_CodPrd, "000") & "' "
   r_str_Parame = r_str_Parame & "   AND PARPRD_CODSUB = '" & Format(p_CodSub, "000") & "' "
   r_str_Parame = r_str_Parame & "   AND PARPRD_CODGRP = '051' "
   r_str_Parame = r_str_Parame & "   AND PARPRD_CODITE = '" & p_CodIte & "'" '012'

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_MtoBBP, 3) Then
       Exit Function
   End If
   
   If r_rst_MtoBBP.BOF And r_rst_MtoBBP.EOF Then
     r_rst_MtoBBP.Close
     Set r_rst_MtoBBP = Nothing
     Exit Function
   End If
   
   r_rst_MtoBBP.MoveFirst
   fs_Obtiene_MontoBBP = r_rst_MtoBBP!PARPRD_CANTID
   
   r_rst_MtoBBP.Close
   Set r_rst_MtoBBP = Nothing
End Function

Private Sub fs_Muestra_Cronograma1()
Dim r_dbl_Cuo_Capita    As Double
Dim r_dbl_Cuo_Intere    As Double
Dim r_dbl_Cuo_SegPre    As Double
Dim r_dbl_Cuo_SegViv    As Double
Dim r_dbl_Cuo_Portes    As Double
Dim r_dbl_Cuo_TotCuo    As Double
Dim r_dbl_Tot_Capita    As Double
Dim r_dbl_Tot_Intere    As Double
Dim r_dbl_Tot_SegPre    As Double
Dim r_dbl_Tot_SegViv    As Double
Dim r_dbl_Tot_Portes    As Double
Dim r_dbl_Tot_TotCuo    As Double
Dim r_int_Contad        As Integer

   grd_Listad_NCo.Redraw = False
   Call gs_LimpiaGrid(grd_Listad_NCo)
   r_dbl_Tot_Capita = 0
   r_dbl_Tot_Intere = 0
   r_dbl_Tot_SegPre = 0
   r_dbl_Tot_SegViv = 0
   r_dbl_Tot_Portes = 0
   r_dbl_Tot_TotCuo = 0
   
   For r_int_Contad = 1 To UBound(l_Arr_TNC_Cli)
      grd_Listad_NCo.Rows = grd_Listad_NCo.Rows + 1
      grd_Listad_NCo.Row = grd_Listad_NCo.Rows - 1
      
      r_dbl_Cuo_Capita = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 4), "###,##0.00"))
      r_dbl_Cuo_Intere = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 5), "###,##0.00"))
      r_dbl_Cuo_SegPre = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 6), "###,##0.00"))
      r_dbl_Cuo_SegViv = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 7), "###,##0.00"))
      r_dbl_Cuo_Portes = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 8), "###,##0.00"))
      r_dbl_Cuo_TotCuo = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 9), "###,##0.00"))
      r_dbl_Tot_Capita = r_dbl_Tot_Capita + r_dbl_Cuo_Capita
      r_dbl_Tot_Intere = r_dbl_Tot_Intere + r_dbl_Cuo_Intere
      r_dbl_Tot_SegPre = r_dbl_Tot_SegPre + r_dbl_Cuo_SegPre
      r_dbl_Tot_SegViv = r_dbl_Tot_SegViv + r_dbl_Cuo_SegViv
      r_dbl_Tot_Portes = r_dbl_Tot_Portes + r_dbl_Cuo_Portes
      r_dbl_Tot_TotCuo = r_dbl_Tot_TotCuo + r_dbl_Cuo_TotCuo
      
      grd_Listad_NCo.Col = 0
      grd_Listad_NCo.Text = Format(r_int_Contad, "000")
      
      grd_Listad_NCo.Col = 1
      grd_Listad_NCo.Text = Format(l_Arr_TNC_Cli(r_int_Contad, 2), "dd/mm/yyyy")
      
      grd_Listad_NCo.Col = 2
      grd_Listad_NCo.Text = Format(r_dbl_Cuo_Capita, "###,##0.00")
      
      grd_Listad_NCo.Col = 3
      grd_Listad_NCo.Text = Format(r_dbl_Cuo_Intere, "###,##0.00")
      
      grd_Listad_NCo.Col = 4
      grd_Listad_NCo.Text = Format(r_dbl_Cuo_SegPre, "###,##0.00")
      
      grd_Listad_NCo.Col = 5
      grd_Listad_NCo.Text = Format(r_dbl_Cuo_SegViv, "###,##0.00")
      
      grd_Listad_NCo.Col = 6
      grd_Listad_NCo.Text = Format(r_dbl_Cuo_Portes, "###,##0.00")
      
      grd_Listad_NCo.Col = 7
      grd_Listad_NCo.Text = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
      
      grd_Listad_NCo.Col = 8
      grd_Listad_NCo.Text = Format(l_Arr_TNC_Cli(r_int_Contad, 10), "###,##0.00")
   Next r_int_Contad
   
   grd_Listad_NCo.Redraw = True
   Call gs_UbiIniGrid(grd_Listad_NCo)
   pnl_Tot_Capita_NCo.Caption = Format(r_dbl_Tot_Capita, "###,##0.00") & " "
   pnl_Tot_Intere_NCo.Caption = Format(r_dbl_Tot_Intere, "###,##0.00") & " "
   pnl_Tot_SegPre_NCo.Caption = Format(r_dbl_Tot_SegPre, "###,##0.00") & " "
   pnl_Tot_SegViv_NCo.Caption = Format(r_dbl_Tot_SegViv, "###,##0.00") & " "
   pnl_Tot_OtrCar_NCo.Caption = Format(r_dbl_Tot_Portes, "###,##0.00") & " "
   pnl_Tot_TotCuo_NCo.Caption = Format(r_dbl_Tot_TotCuo, "###,##0.00") & " "
End Sub

Private Sub fs_Muestra_Cronograma2()
Dim r_dbl_Cuo_Capita    As Double
Dim r_dbl_Cuo_Intere    As Double
Dim r_dbl_Cuo_TotCuo    As Double
Dim r_dbl_Tot_Capita    As Double
Dim r_dbl_Tot_Intere    As Double
Dim r_dbl_Tot_TotCuo    As Double
Dim r_int_Contad        As Integer
   
   grd_Listad_Con.Redraw = False
   Call gs_LimpiaGrid(grd_Listad_Con)
   r_dbl_Tot_Capita = 0
   r_dbl_Tot_Intere = 0
   r_dbl_Tot_TotCuo = 0
   
   If UBound(l_Arr_TC_Cli) > 0 Then
      For r_int_Contad = 1 To UBound(l_Arr_TC_Cli)
         grd_Listad_Con.Rows = grd_Listad_Con.Rows + 1
         grd_Listad_Con.Row = grd_Listad_Con.Rows - 1
         
         r_dbl_Cuo_Capita = CDbl(Format(l_Arr_TC_Cli(r_int_Contad, 4), "###,##0.00"))
         r_dbl_Cuo_Intere = CDbl(Format(l_Arr_TC_Cli(r_int_Contad, 5), "###,##0.00"))
         r_dbl_Cuo_TotCuo = CDbl(Format(l_Arr_TC_Cli(r_int_Contad, 7), "###,##0.00"))
         r_dbl_Tot_Capita = r_dbl_Tot_Capita + r_dbl_Cuo_Capita
         r_dbl_Tot_Intere = r_dbl_Tot_Intere + r_dbl_Cuo_Intere
         r_dbl_Tot_TotCuo = r_dbl_Tot_TotCuo + r_dbl_Cuo_TotCuo
         
         grd_Listad_Con.Col = 0
         grd_Listad_Con.Text = Format(r_int_Contad, "000")
         
         grd_Listad_Con.Col = 1
         grd_Listad_Con.Text = Format(l_Arr_TC_Cli(r_int_Contad, 2), "dd/mm/yyyy")
         
         grd_Listad_Con.Col = 2
         grd_Listad_Con.Text = Format(r_dbl_Cuo_Capita, "###,##0.00")
         
         grd_Listad_Con.Col = 3
         grd_Listad_Con.Text = Format(r_dbl_Cuo_Intere, "###,##0.00")
         
         grd_Listad_Con.Col = 4
         grd_Listad_Con.Text = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
         
         grd_Listad_Con.Col = 5
         grd_Listad_Con.Text = Format(l_Arr_TC_Cli(r_int_Contad, 8), "###,##0.00")
      Next r_int_Contad
      
      grd_Listad_Con.Redraw = True
      Call gs_UbiIniGrid(grd_Listad_Con)
      pnl_Tot_Capita_Con.Caption = Format(r_dbl_Tot_Capita, "###,##0.00") & " "
      pnl_Tot_Intere_Con.Caption = Format(r_dbl_Tot_Intere, "###,##0.00") & " "
      pnl_Tot_TotCuo_Con.Caption = Format(r_dbl_Tot_TotCuo, "###,##0.00") & " "
   End If
End Sub

Private Sub fs_Calcul_MtoMax()
Dim r_int_TipVal_Viv As Integer
Dim r_dbl_Import_Viv As Double
Dim r_int_TipVal_Des As Integer
Dim r_dbl_Import_Des As Double
Dim r_dbl_SegViv     As Double
Dim r_dbl_Portes     As Double
Dim r_dbl_CuoRta     As Double
Dim r_dbl_CuoMen     As Double
Dim r_dbl_PlzMax     As Double
Dim r_dbl_PorCon     As Double
Dim r_dbl_TopCon     As Double
Dim r_dbl_CuoFin     As Double
Dim r_int_TipSeg     As Integer
Dim r_dbl_CuoSol     As Double
Dim r_dbl_CuoApr     As Double
Dim r_dbl_ValInm     As Double
Dim r_arr_ParPrd()   As moddat_tpo_Genera
   
   'inicializa
   pnl_CuoApr.Caption = "0.00 "
   pnl_MtoMax.Caption = "0.00 "
   
   'validaciones
   If cmb_Produc.ListIndex = -1 Then
      Exit Sub
   End If
   If cmb_SubPrd.ListIndex = -1 Then
      Exit Sub
   End If
   If cmb_TipIng.ListIndex = -1 Then
      Exit Sub
   End If
   If cmb_TipIng.ItemData(cmb_TipIng.ListIndex) = 1 Then
      r_int_TipSeg = 11
   Else
      r_int_TipSeg = 12
   End If
   If ipp_IngNet.Value = 0 Then
      Exit Sub
   End If
   If ipp_IngNet.Value < 100 Then
      Exit Sub
   End If
   If cmb_TipIng.ListIndex = -1 Then
      Exit Sub
   End If
   
   'Obtiene seguros
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000003", r_int_TipSeg, l_int_TipMon, 1, r_int_TipVal_Des, r_dbl_Import_Des, "9999")
   Call moddat_gf_Consulta_ValSeg(l_str_CodPrd, l_str_CodSub, "000003", 0, l_int_TipMon, 1, r_int_TipVal_Viv, r_dbl_Import_Viv, "9999")
   
   'Obtiene portes
   r_dbl_Portes = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Obtiene Plazo Maximo del Producto
   r_dbl_PlzMax = 0
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub) Then
      r_dbl_PlzMax = moddat_g_arr_Genera(1).Genera_PlzMax
   End If
   
   'Obtiene Relación Cuota / Renta
   r_dbl_CuoRta = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), l_str_CodPrd, l_str_CodSub, "001", "013") Then
      r_dbl_CuoRta = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   '******************************************************************************
   'FECHA: 29/08/2012
   'RAFAEL DURAND BANDA
   '******************************************************************************
   'SE COLOCA TEMPORALMENTE PARA CUMPLIR CON LAS POLITICAS.
   'LUEGO SE MODIFICARA PARA QUE LA CUOTA RENTA SE OBTENGA DE PARAMETROS
   '******************************************************************************
   If InStr(moddat_g_str_AgrTMIC, l_str_CodPrd) > 0 Then 'l_str_CodPrd = "002" Or l_str_CodPrd = "006" Or l_str_CodPrd = "011" Or l_str_CodPrd = "012" Then
      If CDbl(ipp_IngNet.Text) >= 1000 And CDbl(ipp_IngNet.Text) < 2000 Then
         r_dbl_CuoRta = 30
      End If
      If CDbl(ipp_IngNet.Text) >= 2000 And CDbl(ipp_IngNet.Text) < 4000 Then
         r_dbl_CuoRta = 35
      End If
      If CDbl(ipp_IngNet.Text) >= 4000 Then
         r_dbl_CuoRta = 40
      End If
   End If
   '******************************************************************************
   '******************************************************************************
   '******************************************************************************
   
   r_dbl_CuoSol = r_dbl_CuoRta / 100 * ipp_IngNet.Value
   r_dbl_CuoApr = 0
   
   If l_int_TipMon = 1 Then
      r_dbl_CuoApr = r_dbl_CuoRta / 100 * ipp_IngNet.Value
   Else
      r_dbl_CuoApr = (r_dbl_CuoRta / 100 * ipp_IngNet.Value) / CDbl(pnl_TipCam.Caption)
   End If
   
   'Obtene Valor Máximo del Inmueble
   r_dbl_ValInm = 0
   Select Case l_str_CodPrd > 0
      'En Monto
      Case InStr(moddat_g_str_AgrTMIC, l_str_CodPrd)   '"002", "006", "011"
         If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "001", "021") Then
            r_dbl_ValInm = r_arr_ParPrd(1).Genera_Cantid
         End If
      
      'En UIT
      Case InStr(moddat_g_str_AgrCME, l_str_CodPrd)    '"003"
         If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "022") Then
            r_dbl_ValInm = r_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      'En UIT (Mínimo y Máximo)
      Case InStr(moddat_g_str_AgrTFMV, l_str_CodPrd)   '"004", "007", "009", "010", "012", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023"
         If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, l_str_CodPrd, l_str_CodSub, "051", "022") Then
            r_dbl_ValInm = r_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
   End Select
   
   If r_int_TipVal_Viv = 1 Then
      r_dbl_SegViv = CDbl(Format(r_dbl_Import_Viv / 100 * r_dbl_ValInm, "###0.00"))
   Else
      r_dbl_SegViv = r_dbl_Import_Viv
   End If
   
   Select Case l_str_CodPrd > 0
      Case InStr(moddat_g_str_Agr1MIC, l_str_CodPrd) Or InStr(moddat_g_str_Agr1FMV, l_str_CodPrd)   '"002", "011", "019", "021", "022", "023"
         r_dbl_CuoMen = r_dbl_CuoApr - r_dbl_SegViv - r_dbl_Portes
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_miCasita(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, r_dbl_ValInm, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, r_dbl_CuoApr, l_dbl_TasInt, r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   
      Case InStr(moddat_g_str_AgrCME, l_str_CodPrd)    '"003"
         r_dbl_PorCon = 0
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, 0, l_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         r_dbl_CuoMen = CDbl(Format((r_dbl_CuoApr / ((100 - r_dbl_PorCon) / 100)) - r_dbl_SegViv - r_dbl_Portes, "####0.00"))
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_CME(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, r_dbl_ValInm, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, r_dbl_CuoApr, l_dbl_TasInt, r_dbl_PorCon, r_dbl_TopCon, r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   
      Case InStr(moddat_g_str_AgrTFMV, l_str_CodPrd)   '"004", "007", "009", "010", "012", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023"
         r_dbl_TopCon = 0
         r_dbl_CuoMen = CDbl(Format(r_dbl_CuoApr - r_dbl_SegViv - r_dbl_Portes, "####0.00"))
         pnl_MtoMax.Caption = Format(modcal_gf_Calcul_MtoMax_MiHogar(r_dbl_CuoMen, l_dbl_TasInt + r_dbl_Import_Des, Format(date, "dd/mm/yyyy"), r_dbl_PlzMax * 12, r_dbl_ValInm, r_dbl_Import_Des, r_dbl_Import_Viv, r_dbl_Portes, r_dbl_CuoApr, l_dbl_TasInt, r_dbl_TopCon, r_dbl_CuoFin), "###,##0.00") & " "
         pnl_CuoApr.Caption = Format(r_dbl_CuoFin, "###,##0.00") & " "
   End Select
End Sub

'**************************************************************************************************
'******************************************* CONTROLES ********************************************
'**************************************************************************************************
Private Sub cmb_Produc_Click()
Dim r_int_Contad     As Integer
   
   cmb_SubPrd.Clear
   pnl_FmvBbp.Caption = "0.00 "
   pnl_MefPbp.Caption = "0.00 "
   pnl_FmvBbp_Sol.Caption = "0.00 "
   pnl_MefPbp_Sol.Caption = "0.00 "
   pnl_ComVta_Sol.Caption = "0.00 "
   pnl_MtoPre_Sol.Caption = "0.00 "
   pnl_ApoPro_Sol.Caption = "0.00 "
   pnl_TipCam.Caption = "0.0000 "
   pnl_TasInt.Caption = "0.0000 "
   pnl_SegDes.Caption = "0.000000 "
   pnl_SegInm.Caption = "0.000000 "
   pnl_CosEfe.Caption = "0.00 "
   pnl_CuoSBP.Caption = "0.00 "
   pnl_MtoPBP.Caption = "0.00 "
   pnl_CuoPBP.Caption = "0.00 "
   
   pnl_CuoIni.Caption = "0.00 "
   pnl_CuoIni_Sol.Caption = "0.00 "
   pnl_MtoBMS.Caption = "0.00 "
   pnl_MtoBMS_Sol.Caption = "0.00 "
   pnl_MtoAFP_Sol.Caption = "0.00 "
   
   cmb_CuoDbl.ListIndex = -1
   cmb_TipSeg.ListIndex = -1
   cmb_TasEsp.ListIndex = -1
   cmb_DiaPag.Clear
   pnl_CuoMen.Caption = "0.00 "
   pnl_CuoMen_Sol.Caption = "0.00 "
   pnl_IngReq.Caption = "0.00 "
   pnl_IngReq_Sol.Caption = "0.00 "
   
   'Limpiando Grid TNC
   Call gs_LimpiaGrid(grd_Listad_NCo)
   pnl_Tot_Capita_NCo.Caption = "0.00 "
   pnl_Tot_Intere_NCo.Caption = "0.00 "
   pnl_Tot_SegPre_NCo.Caption = "0.00 "
   pnl_Tot_SegViv_NCo.Caption = "0.00 "
   pnl_Tot_OtrCar_NCo.Caption = "0.00 "
   pnl_Tot_TotCuo_NCo.Caption = "0.00 "

   'Limpiando Grid TC
   Call gs_LimpiaGrid(grd_Listad_Con)
   pnl_Tot_Capita_Con.Caption = "0.00 "
   pnl_Tot_Intere_Con.Caption = "0.00 "
   pnl_Tot_TotCuo_Con.Caption = "0.00 "
   
   For r_int_Contad = 0 To lbl_SimMon.Count - 1
      lbl_SimMon(r_int_Contad).Caption = ""
   Next r_int_Contad
   
   lbl_Totale(0).Caption = "Totales ==> "
   lbl_Totale(1).Caption = "Totales ==> "
   l_str_CodPrd = ""
   
   'Simulación por Ingreso
   ipp_IngNet.Value = 0
   cmb_TipIng.ListIndex = -1
   pnl_MtoMax.Caption = "0.00 "
   pnl_CuoApr.Caption = "0.00 "
   
   If cmb_Produc.ListIndex > -1 Then
      Screen.MousePointer = 11
      l_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
      Call moddat_gs_Carga_SubPrd(cmb_SubPrd, l_arr_SubPrd, l_str_CodPrd)
      
      If l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "004" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "003" Or _
         l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "007" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "009" Or _
         l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "010" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "013" Or _
         l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "014" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "015" Or _
         l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "016" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "017" Or _
         l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "018" Or l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "006" Then
         
         tab_Cronog.TabCaption(0) = "Cronograma Tramo No Concesional"
         tab_Cronog.TabCaption(1) = "Cronograma Tramo Concesional"
         tab_Cronog.TabVisible(1) = True
         
         lbl_CuoSBP.Visible = True
         lbl_MtoPBP.Visible = True
         lbl_CuoPBP.Visible = True
         pnl_MtoPBP.Visible = True
         pnl_CuoPBP.Visible = True
         pnl_CuoSBP.Visible = True
         lbl_SimMon(5).Visible = True
         lbl_SimMon(6).Visible = True
         lbl_SimMon(7).Visible = True
      Else
         tab_Cronog.TabCaption(0) = "Cronograma de Pagos"
         tab_Cronog.TabVisible(1) = False
         
         lbl_MtoPBP.Visible = False
         lbl_CuoSBP.Visible = False
         lbl_CuoPBP.Visible = False
         pnl_MtoPBP.Visible = False
         pnl_CuoPBP.Visible = False
         pnl_CuoSBP.Visible = False
         lbl_SimMon(5).Visible = False
         lbl_SimMon(6).Visible = False
         lbl_SimMon(7).Visible = False
         
         If l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo = "021" Then
            pnl_FmvBbp.Caption = "12,500.00 "
         End If
      End If
      
      Call gs_SetFocus(cmb_SubPrd)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_SubPrd_Click()
Dim r_str_SimMon     As String
Dim r_int_Contad     As Integer

   pnl_TipCam.Caption = "0.0000 "
   pnl_TasInt.Caption = "0.00 "
   pnl_SegDes.Caption = "0.000000 "
   pnl_SegInm.Caption = "0.000000 "
   pnl_CosEfe.Caption = "0.00 "
   pnl_CuoSBP.Caption = "0.00 "
   pnl_MtoPBP.Caption = "0.00 "
   pnl_CuoPBP.Caption = "0.00 "
   pnl_ComVta_Sol.Caption = "0.00 "
   pnl_MtoPre_Sol.Caption = "0.00 "
   pnl_ApoPro_Sol.Caption = "0.00 "
   
   pnl_CuoIni.Caption = "0.00 "
   pnl_CuoIni_Sol.Caption = "0.00 "
   pnl_MtoBMS.Caption = "0.00 "
   pnl_MtoBMS_Sol.Caption = "0.00 "
   pnl_MtoAFP_Sol.Caption = "0.00 "
   
   cmb_TipSeg.ListIndex = -1
   cmb_DiaPag.Clear
   pnl_CuoMen.Caption = "0.00 "
   pnl_IngReq.Caption = "0.00 "
   pnl_CuoMen_Sol.Caption = "0.00 "
   pnl_IngReq_Sol.Caption = "0.00 "
   
   'Limpiando Grid TNC
   Call gs_LimpiaGrid(grd_Listad_NCo)
   pnl_Tot_Capita_NCo.Caption = "0.00 "
   pnl_Tot_Intere_NCo.Caption = "0.00 "
   pnl_Tot_SegPre_NCo.Caption = "0.00 "
   pnl_Tot_SegViv_NCo.Caption = "0.00 "
   pnl_Tot_OtrCar_NCo.Caption = "0.00 "
   pnl_Tot_TotCuo_NCo.Caption = "0.00 "

   'Limpiando Grid TC
   Call gs_LimpiaGrid(grd_Listad_Con)
   pnl_Tot_Capita_Con.Caption = "0.00 "
   pnl_Tot_Intere_Con.Caption = "0.00 "
   pnl_Tot_TotCuo_Con.Caption = "0.00 "
   
   For r_int_Contad = 0 To lbl_SimMon.Count - 1
      lbl_SimMon(r_int_Contad).Caption = ""
   Next r_int_Contad
   
   lbl_Totale(0).Caption = "Totales ==> "
   lbl_Totale(1).Caption = "Totales ==> "
   
   'Simulación por Ingreso
   ipp_IngNet.Value = 0
   cmb_TipIng.ListIndex = -1
   pnl_MtoMax.Caption = "0.00 "
   pnl_CuoApr.Caption = "0.00 "
   
   l_str_CodSub = ""
   l_int_TipMon = 0
   l_dbl_TipCam = 0
   l_dbl_TasInt = 0
   
   If cmb_SubPrd.ListIndex > -1 Then
      l_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
      l_int_TipMon = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_TipMon
      r_str_SimMon = moddat_gf_Consulta_ParDes("229", CStr(l_int_TipMon))
   
      For r_int_Contad = 0 To lbl_SimMon.Count - 1
         lbl_SimMon(r_int_Contad).Caption = r_str_SimMon
      Next r_int_Contad
      
      lbl_Totale(0).Caption = "Totales ==> " & r_str_SimMon & " "
      lbl_Totale(1).Caption = "Totales ==> " & r_str_SimMon & " "
   
      'Tipo de Cambio
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, l_int_TipMon)
      pnl_TipCam.Caption = Format(l_dbl_TipCam, "###,##0.0000") & " "
      
      'Plazo del Crédito
      ipp_PlaAno.MinValue = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_PlzMin
      ipp_PlaAno.MaxValue = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_PlzMax
      
      'Periodo de Gracia
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "008", "002") Then
         ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
         ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      End If
      
      'Día de Pago
      Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), l_str_CodPrd, l_str_CodSub, "009")
      
      If InStr(moddat_g_str_Agr1FMV, l_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, l_str_CodPrd) > 0 Then   '"007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
         If Day(date) < 29 Then
            cmb_DiaPag.ListIndex = gf_Busca_Arregl(l_arr_DiaPag, Format(Day(date), "000")) - 1
         Else
            cmb_DiaPag.ListIndex = gf_Busca_Arregl(l_arr_DiaPag, "001") - 1
         End If
      End If
      
      'Tasa de Interes
      l_dbl_TasInt = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, l_str_CodPrd, l_str_CodSub, "002", "101") Then
         l_dbl_TasInt = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      pnl_TasInt.Caption = Format(l_dbl_TasInt, "##0.00") & " "
      
      'Obtiene Bono para todos los productos FMV
      If InStr(moddat_g_str_Agr2FMV, l_str_CodPrd) = 0 And InStr(moddat_g_str_Agr2MIC, l_str_CodPrd) = 0 Then
         pnl_FmvBbp.Caption = Format(fs_Obtiene_MontoBBP(l_str_CodPrd, l_str_CodSub, "012"), "###,##0.00") & " "
      End If
      'Obtiene Bono MEF
      If l_str_CodPrd = "023" Then
         pnl_MefPbp.Caption = Format(fs_Obtiene_MontoBBP(l_str_CodPrd, l_str_CodSub, "013"), "###,##0.00") & " "
      End If
      
      Call gs_SetFocus(ipp_ComVta)
   End If
End Sub

Private Sub cmb_SubPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SubPrd_Click
   End If
End Sub

Private Sub ipp_ComVta_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEst)
   End If
End Sub

Private Sub ipp_ComVta_LostFocus()
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_ApoPro_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_ApoPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoAFP)
   End If
End Sub

Private Sub ipp_ApoPro_LostFocus()
   Call fs_Calcular_Prestamo
End Sub

Private Sub ipp_MtoAFP_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub ipp_PerGra_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub ipp_PlaAno_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub ipp_PerGra_Change()
   Call fs_Limpia_Calcul
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CuoDbl)
   End If
End Sub

Private Sub cmb_CuoDbl_Click()
   If cmb_CuoDbl.ListIndex > -1 Then
      Call gs_SetFocus(cmb_DiaPag)
   End If
End Sub

Private Sub cmb_CuoDbl_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CuoDbl_Click
   End If
End Sub

Private Sub cmb_TipSeg_Click()
   If cmb_TipSeg.ListIndex > -1 Then
      Call fs_Limpia_Calcul
      Call Calculo_TasEsp
      Call gs_SetFocus(cmb_TasEsp)
   End If
End Sub

Private Sub cmb_TipSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipSeg_Click
   End If
End Sub

Private Sub cmb_TasEsp_Click()
   If cmb_TasEsp.ListIndex > -1 Then
      Call Calculo_TasEsp
      Call gs_SetFocus(cmd_CalCuo) 'cmb_DiaPag
   End If
End Sub

Private Sub cmb_TasEsp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TasEsp_Click
   End If
End Sub

Public Sub Calculo_TasEsp()
Dim p_TipSeg   As Integer
Dim p_TipMon   As Integer
Dim p_MtoPre   As Double
Dim p_Import   As Double
Dim p_TipVal   As Integer
      
   p_MtoPre = CDbl(pnl_MtoPre.Caption)
   
   If cmb_TipSeg.ListIndex = -1 Then
      Exit Sub
   End If
   If l_str_CodPrd = "" Then
      Exit Sub
   End If
   If p_MtoPre = 0 Then
      Exit Sub
   End If
   
   p_TipSeg = Format(cmb_TipSeg.ItemData(cmb_TipSeg.ListIndex), "000")
   p_Import = 0
   
   If p_TipSeg > 0 Then    'Seguro de Préstamo
      g_str_Parame = "SELECT * FROM MNT_SEGPRE WHERE "
      g_str_Parame = g_str_Parame & "SEGPRE_CODPRD = '" & l_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODSUB = '" & l_str_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODIGO = '000003' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPSEG = " & p_TipSeg & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPMON = " & l_int_TipMon & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMAX >= " & CStr(p_MtoPre) & " "
   Else                    'Seguro de Vivienda
      g_str_Parame = "SELECT * FROM MNT_SEGVIV WHERE "
      g_str_Parame = g_str_Parame & "SEGVIV_CODPRD = '" & l_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODSUB = '" & l_str_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODIGO = '000003' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_TIPMON = " & CStr(l_int_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMAX >= " & CStr(p_MtoPre) & " "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   If p_TipSeg > 0 Then
      p_TipVal = g_rst_Genera!SEGPRE_VTATIP
      p_Import = g_rst_Genera!SEGPRE_VTAFOI
   Else
      p_TipVal = g_rst_Genera!SEGVIV_VTATIP
      p_Import = g_rst_Genera!SEGVIV_VTAFOI
   End If
   
   If cmb_TasEsp.ListIndex > -1 Then
      If cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex) = 1 Then
         p_Import = p_Import + (p_Import * 0)
      ElseIf cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex) = 2 Then
         p_Import = p_Import + (p_Import * (50 / 100))
      ElseIf cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex) = 3 Then
         p_Import = p_Import + (p_Import * (100 / 100))
      ElseIf cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex) = 4 Then
         p_Import = p_Import + (p_Import * (150 / 100))
      ElseIf cmb_TasEsp.ItemData(cmb_TasEsp.ListIndex) = 5 Then
         p_Import = p_Import + (p_Import * (200 / 100))
      End If
   End If
   
   pnl_SegDes.Caption = Format(p_Import, "###,##0.000000") & " "
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub cmb_DiaPag_Click()
   If cmb_DiaPag.ListIndex > -1 Then
      Call fs_Limpia_Calcul
      Call gs_SetFocus(cmb_TipSeg)
   End If
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DiaPag_Click
   End If
End Sub

Private Sub grd_Listad_Con_SelChange()
   If grd_Listad_Con.Rows > 2 Then
      grd_Listad_Con.RowSel = grd_Listad_Con.Row
   End If
End Sub

Private Sub grd_Listad_NCo_SelChange()
   If grd_Listad_NCo.Rows > 2 Then
      grd_Listad_NCo.RowSel = grd_Listad_NCo.Row
   End If
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipIng)
   End If
End Sub

Private Sub cmb_TipIng_Click()
   If cmb_TipIng.ListIndex > -1 Then
      Call gs_SetFocus(cmd_CalEst)
   End If
End Sub

Private Sub cmb_TipIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmb_TipIng_Click
   End If
End Sub

Private Sub ipp_MtoAFP_Change()
    Call fs_Limpia_Calcul
End Sub

Private Sub ipp_MtoAFP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BMSTas)
   End If
End Sub

Private Sub ipp_PlaAno_LostFocus()
    Call fs_Calcular_Prestamo
    Call fs_Calcular_GCierre
End Sub

Private Sub ipp_ValEst_Change()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

Private Sub ipp_ValEst_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ApoPro)
   End If
End Sub

Private Sub ipp_ValEst_LostFocus()
   Call fs_Calcular_Prestamo
   Call fs_Calcular_GCierre
End Sub

