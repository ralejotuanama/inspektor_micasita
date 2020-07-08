VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_SolCre_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10425
   ClientLeft      =   4215
   ClientTop       =   1965
   ClientWidth     =   11640
   Icon            =   "AteCli_frm_114.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10425
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   18389
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   1395
         Left            =   30
         TabIndex        =   54
         Top             =   4770
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2461
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
         Begin VB.ComboBox cmb_MonAho 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_InsFin 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   60
            Width           =   3315
         End
         Begin EditLib.fpDoubleSingle ipp_MtoAho 
            Height          =   315
            Left            =   2040
            TabIndex        =   59
            Top             =   720
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
         Begin EditLib.fpLongInteger ipp_MesAho 
            Height          =   315
            Left            =   2040
            TabIndex        =   61
            Top             =   1050
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
         Begin VB.Label Label22 
            Caption         =   "Meses Ahorrados:"
            Height          =   285
            Left            =   60
            TabIndex        =   62
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label Label18 
            Caption         =   "Monto Mínimo de Ahorro:"
            Height          =   285
            Left            =   60
            TabIndex        =   60
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   390
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Institución Financiera:"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   60
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1935
         Left            =   30
         TabIndex        =   50
         Top             =   7620
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3413
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
            Height          =   1545
            Left            =   30
            TabIndex        =   51
            Top             =   360
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   2725
            _Version        =   393216
            Rows            =   12
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   9750
            TabIndex        =   52
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   60
            TabIndex        =   53
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   1365
         Left            =   30
         TabIndex        =   43
         Top             =   6210
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2408
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
         Begin VB.TextBox txt_Observ 
            Height          =   585
            Left            =   2040
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Text            =   "AteCli_frm_114.frx":000C
            Top             =   60
            Width           =   9345
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   660
            Width           =   9345
         End
         Begin VB.ComboBox cmb_EjeSeg 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   990
            Width           =   9345
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label8 
            Caption         =   "Consejero Hipotecario:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   660
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Ejecutivo de Seguimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   47
            Top             =   990
            Width           =   1905
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   12
         Top             =   9600
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Calcul 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_114.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Calculadora"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_114.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_114.frx":0624
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Aceptar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "AteCli_frm_114.frx":092E
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3225
         Left            =   30
         TabIndex        =   13
         Top             =   1500
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5689
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
         Begin VB.ComboBox cmb_TipEva 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9375
         End
         Begin VB.ComboBox cmb_EmpSeg 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2190
            Width           =   9345
         End
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2850
            Width           =   1635
         End
         Begin VB.ComboBox cmb_SegDes 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2520
            Width           =   9345
         End
         Begin EditLib.fpDoubleSingle ipp_ComVta 
            Height          =   315
            Left            =   2070
            TabIndex        =   1
            Top             =   390
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
            Left            =   2070
            TabIndex        =   2
            Top             =   720
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
         Begin EditLib.fpDoubleSingle ipp_MtoPre 
            Height          =   315
            Left            =   2070
            TabIndex        =   3
            Top             =   1050
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
         Begin EditLib.fpLongInteger ipp_PlaAno 
            Height          =   315
            Left            =   2040
            TabIndex        =   4
            Top             =   1530
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
            MaxValue        =   "30"
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
            Left            =   2040
            TabIndex        =   5
            Top             =   1860
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
            MaxValue        =   "6"
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
         Begin Threed.SSPanel pnl_ComVta_Sol 
            Height          =   315
            Left            =   5640
            TabIndex        =   28
            Top             =   390
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_MtoPre_Sol 
            Height          =   315
            Left            =   5640
            TabIndex        =   29
            Top             =   1050
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_ApoPro_Sol 
            Height          =   315
            Left            =   5640
            TabIndex        =   30
            Top             =   720
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_ComVta_Dol 
            Height          =   315
            Left            =   9120
            TabIndex        =   34
            Top             =   390
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_MtoPre_Dol 
            Height          =   315
            Left            =   9120
            TabIndex        =   35
            Top             =   1050
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel pnl_ApoPro_Dol 
            Height          =   315
            Left            =   9120
            TabIndex        =   36
            Top             =   720
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   90
            Left            =   30
            TabIndex        =   40
            Top             =   1410
            Width           =   11460
            _Version        =   65536
            _ExtentX        =   20205
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
         Begin VB.Label Label9 
            Caption         =   "Tipo de Evaluación:"
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Compañía de Seguros:"
            Height          =   315
            Left            =   60
            TabIndex        =   41
            Top             =   2190
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Valores en US$:"
            Height          =   315
            Left            =   7830
            TabIndex        =   39
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label12 
            Caption         =   "Valores en US$:"
            Height          =   315
            Left            =   7830
            TabIndex        =   38
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "Valores en US$:"
            Height          =   315
            Left            =   7830
            TabIndex        =   37
            Top             =   750
            Width           =   1185
         End
         Begin VB.Label Label11 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4470
            TabIndex        =   33
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label13 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4470
            TabIndex        =   32
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label17 
            Caption         =   "Valores en S/.:"
            Height          =   315
            Left            =   4470
            TabIndex        =   31
            Top             =   750
            Width           =   1185
         End
         Begin VB.Label Label6 
            Caption         =   "Día de Pago:"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   2850
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   2520
            Width           =   1905
         End
         Begin VB.Label Label4 
            Caption         =   "Período de Gracia:"
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   1860
            Width           =   1665
         End
         Begin VB.Label Label29 
            Caption         =   "Plazo:"
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   1530
            Width           =   1665
         End
         Begin VB.Label Label2 
            Caption         =   "Aporte Propio:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label35 
            Caption         =   "Valor Compra-Venta:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "Monto Solicitado:"
            Height          =   285
            Left            =   90
            TabIndex        =   21
            Top             =   1050
            Width           =   1815
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1085
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   495
            Left            =   660
            TabIndex        =   15
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Solicitud de Crédito - Datos del Crédito"
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
            Picture         =   "AteCli_frm_114.frx":0D70
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   16
         Top             =   690
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   2040
            TabIndex        =   17
            Top             =   60
            Width           =   9405
            _Version        =   65536
            _ExtentX        =   16589
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   390
            Width           =   9405
            _Version        =   65536
            _ExtentX        =   16589
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   19
            Top             =   390
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_SolCre_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_EmpSeg()      As moddat_tpo_Genera
Dim l_arr_Modali()      As moddat_tpo_Genera
Dim l_arr_CuoExt()      As moddat_tpo_Genera
Dim l_arr_DiaPag()      As moddat_tpo_Genera
Dim l_arr_ConHip()      As moddat_tpo_Genera
Dim l_arr_EjeSeg()      As moddat_tpo_Genera
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_arr_InsFin()      As moddat_tpo_Genera
Dim l_arr_TipEva()      As moddat_tpo_Genera
Dim l_dbl_TipCam        As Double
Dim l_int_GraMax        As Integer

Private Sub cmb_ConHip_Click()
   Call gs_SetFocus(cmb_EjeSeg)
End Sub

Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConHip_Click
   End If
End Sub

Private Sub cmb_CuoExt_Click()
   Call gs_SetFocus(cmb_EmpSeg)
End Sub

Private Sub cmb_CuoExt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CuoExt_Click
   End If
End Sub

Private Sub cmb_DiaPag_Click()
   If cmb_DiaPag.ListIndex > -1 Then
      If cmb_TipEva.ListIndex > -1 Then
         If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
            Call gs_SetFocus(cmb_InsFin)
         Else
            Call gs_SetFocus(txt_Observ)
         End If
      End If
   End If
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DiaPag_Click
   End If
End Sub

Private Sub cmb_EjeSeg_Click()
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmb_EjeSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EjeSeg_Click
   End If
End Sub

Private Sub cmb_EmpSeg_Click()
   If cmb_EmpSeg.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
      
      Call gs_SetFocus(cmb_SegDes)
   Else
      cmb_SegDes.Clear
   End If
End Sub

Private Sub cmb_EmpSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpSeg_Click
   End If
End Sub


Private Sub cmb_InsFin_Click()
   Call gs_SetFocus(cmb_MonAho)
End Sub

Private Sub cmb_InsFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_InsFin_Click
   End If
End Sub

Private Sub cmb_MonAho_Click()
   Call gs_SetFocus(ipp_MtoAho)
End Sub

Private Sub cmb_MonAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonAho_Click
   End If
End Sub

Private Sub cmb_SegDes_Click()
   Call gs_SetFocus(cmb_DiaPag)
End Sub

Private Sub cmb_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegDes_Click
   End If
End Sub

Private Sub cmb_TipMon_Change()
   Call fs_Calcul
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_ComVta)
   Call fs_Calcul
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmb_TipEva_Click()
   Call gs_SetFocus(ipp_ComVta)
   
   If cmb_TipEva.ListIndex > -1 Then
      If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
         cmb_InsFin.Enabled = True
         cmb_MonAho.Enabled = True
         ipp_MtoAho.Enabled = True
         ipp_MesAho.Enabled = True
      Else
         cmb_InsFin.Enabled = False
         cmb_MonAho.Enabled = False
         ipp_MtoAho.Enabled = False
         ipp_MesAho.Enabled = False
         
         cmb_InsFin.ListIndex = -1
         cmb_MonAho.ListIndex = -1
         ipp_MtoAho.Value = 0
         ipp_MesAho.Value = 0
      End If
   End If
End Sub

Private Sub cmb_TipEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipEva_Click
   End If
End Sub

Private Sub cmd_Calcul_Click()
   Dim r_lng_NumPid    As Long
   
   r_lng_NumPid = Shell("c:\winnt\system32\calc.exe", vbNormalFocus)
   
   If r_lng_NumPid = 0 Then
      MsgBox "Error Iniciando la Aplicación", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad           As Integer
   Dim r_int_FlgDoc           As Integer
   Dim r_dbl_ValMin_ComVta    As Double
   Dim r_dbl_ValMax_ComVta    As Double
   Dim r_dbl_PorMin_ApoPro    As Double
   Dim r_dbl_PorMax_MtoPre    As Double
   Dim r_dbl_ValMin_MtoPre    As Double
   Dim r_dbl_ValMax_MtoPre    As Double
   Dim r_int_EdaMax           As Integer
   Dim r_int_EdaAct           As Integer
   Dim r_dbl_Aho_ApoMin       As Double
   Dim r_dbl_Aho_ApoTp1       As Double
   Dim r_dbl_Aho_ApoTp2       As Double
   Dim r_dbl_Aho_ApoRgI       As Double
   Dim r_dbl_Aho_ApoRgF       As Double
   Dim r_dbl_ApoMin           As Double
   Dim r_dbl_CuoAho           As Double
   Dim r_dbl_Aho_CuoMin       As Double
   Dim r_dbl_Ini_ApoMin       As Double
   Dim r_dbl_Ini_PlaMin       As Double
   Dim r_dbl_Ini_PlaMax       As Double
   Dim r_dbl_PrcMin           As Double
   Dim r_dbl_PrcMax           As Double
   
   If cmb_TipEva.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEva)
      Exit Sub
   End If
   
   'If moddat_g_str_CodPrd = "001" Then
   '   If cmb_TipEva.ItemData(cmb_TipEva.ListIndex) = 4 Then
   '      MsgBox "El Producto no acepta este Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(cmb_TipEva)
   '      Exit Sub
   '   End If
   'ElseIf moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "002" Then
   '   If cmb_TipEva.ItemData(cmb_TipEva.ListIndex) = 3 Or cmb_TipEva.ItemData(cmb_TipEva.ListIndex) = 4 Then
   '      MsgBox "El Producto no acepta este Tipo de Evaluación.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(cmb_TipEva)
   '      Exit Sub
   '   End If
   'End If
   
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 3 Or CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 4 Then
      If modatecli_g_arr_DatInm(1).DatInm_InmIde = 2 Then
         MsgBox "Este Tipo de Evaluación exige que el Inmueble sea identificado. Registre la información del Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipEva)
         Exit Sub
      End If
   End If
   
   'Buscar Parámetros en Productos
   r_dbl_ValMin_ComVta = 0
   r_dbl_ValMax_ComVta = 0
   r_dbl_PorMin_ApoPro = 0
   r_dbl_PorMax_MtoPre = 0
   r_dbl_ValMin_MtoPre = 0
   r_dbl_ValMax_MtoPre = 0
   r_dbl_PrcMin = 0
   r_dbl_PrcMax = 0

   'Edad Máxima del Cliente
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "012") Then
      r_int_EdaMax = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener Valor Máximo del Inmueble
   Select Case moddat_g_str_CodPrd
      Case "001"  'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case "002"  'En Dólares
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "021") Then
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_Cantid
         End If
      
      Case "003"  'En UIT
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
      
      Case "004"  'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
   
      Case "007"  'En UIT (Mínimo y Máximo)
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "022") Then
            r_dbl_ValMin_ComVta = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002")
            r_dbl_ValMax_ComVta = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002")
         End If
   End Select
   
   'Para obtener % Mínimo de Aporte Propio
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "022") Then
      r_dbl_PorMin_ApoPro = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener % Máximo de Monto de Préstamo
   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "023") Then
      r_dbl_PorMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
   End If
   
   'Para obtener Monto Máximo de Préstamo
   If moddat_g_str_CodPrd = "002" Then
      'En Dólares
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "024") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid
      End If
   ElseIf moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
   ElseIf moddat_g_str_CodPrd = "004" Then
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
      End If
   ElseIf moddat_g_str_CodPrd = "007" Then
      'Porcentaje para Valor Minimo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "024") Then
         r_dbl_PrcMin = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'Porcentaje para Valor Máximo
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "025") Then
         r_dbl_PrcMax = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      'En UIT
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "023") Then
         r_dbl_ValMin_MtoPre = l_arr_ParPrd(1).Genera_ValMin * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMin / 100
         r_dbl_ValMax_MtoPre = l_arr_ParPrd(1).Genera_ValMax * moddat_gf_Consulta_ParVal("001", "002") * r_dbl_PrcMax / 100
      End If
   End If
   
   If CDbl(ipp_ComVta.Text) = 0 Then
      MsgBox "Debe ingresar el Valor de Compra-Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ComVta)
      Exit Sub
   End If
   
   'Validando Valor de Compra Venta
   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Then
      If CDbl(pnl_ComVta_Sol.Caption) < r_dbl_ValMin_ComVta Then
         MsgBox "El Valor de Compra-Venta no cubre el mínimo requerido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
      
      If CDbl(pnl_ComVta_Sol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   ElseIf moddat_g_str_CodPrd = "002" Then
      If CDbl(pnl_ComVta_Dol.Caption) > r_dbl_ValMax_ComVta Then
         MsgBox "El Valor de Compra-Venta excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ComVta)
         Exit Sub
      End If
   End If

   If CDbl(ipp_ApoPro.Text) = 0 Then
      MsgBox "Debe ingresar el Aporte Propio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If CDbl(ipp_ApoPro.Text) > CDbl(ipp_ComVta.Text) Then
      MsgBox "El Aporte Propio no puede ser mayor al Valor de Compra Venta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   If CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100 < r_dbl_PorMin_ApoPro Then
      MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If

   If CDbl(ipp_MtoPre.Text) / CDbl(ipp_ComVta.Text) * 100 > r_dbl_PorMax_MtoPre Then
      MsgBox "El Aporte Propio no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ApoPro)
      Exit Sub
   End If
   
   'Validando Monto de Préstamo
   If moddat_g_str_CodPrd = "002" Then
      If CDbl(pnl_MtoPre_Dol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   ElseIf moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   ElseIf moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Then
      If CDbl(pnl_MtoPre_Sol.Caption) < r_dbl_ValMin_MtoPre Then
         MsgBox "El Monto del Préstamo no cubre el mínimo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
      
      If CDbl(pnl_MtoPre_Sol.Caption) > r_dbl_ValMax_MtoPre Then
         MsgBox "El Monto del Préstamo excede el máximo permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If

   If CDbl(ipp_PlaAno.Text) = 0 Then
      MsgBox "Debe ingresar el Plazo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If
   
   If Not (CInt(ipp_PlaAno.Text) >= ipp_PlaAno.MinValue And CInt(ipp_PlaAno.Text) <= ipp_PlaAno.MaxValue) Then
      MsgBox "El Plazo indicado no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Tit), Date), 2))
   
   If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
      MsgBox "La Edad del Cliente más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PlaAno)
      Exit Sub
   End If

   If Not (CInt(ipp_PerGra.Text) >= ipp_PerGra.MinValue And CInt(ipp_PerGra.Text) <= ipp_PerGra.MaxValue) Then
      MsgBox "El Período de Gracia no se ajusta a los Parámetros permitidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerGra)
      Exit Sub
   End If

   If cmb_EmpSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Seguros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpSeg)
      Exit Sub
   End If

   If cmb_SegDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegDes)
      Exit Sub
   End If
   
   If moddat_g_int_EstCiv <> 2 And moddat_g_int_EstCiv <> 5 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
         MsgBox "El Cliente no requiere tomar Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   'Si cliente complementa Renta
   If moddat_g_int_ComRta = 1 Then
      If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) <> 12 Then
         MsgBox "El Cliente presenta Complemento de Renta debe seleccionar el Seguro de Desgravamen Mancomunado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SegDes)
         Exit Sub
      End If
   End If
   
   If cmb_SegDes.ItemData(cmb_SegDes.ListIndex) = 12 Then
      r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(moddat_g_str_FecNac_Cyg), Date), 2))
      
      If r_int_EdaAct + CInt(ipp_PlaAno.Text) > r_int_EdaMax Then
         MsgBox "La Edad del Cónyuge más el Plazo del Préstamo excede el parámetro permitido. El Plazo máximo podría ser de " & CStr(r_int_EdaMax - r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaAno)
         Exit Sub
      End If
   End If
   
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Día de Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
      Exit Sub
   End If
   
   'Evaluación Normal
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 1 Then
      'Validando que Clientes de Provincias cumplan con Aporte Inicial mínimo
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "025") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial mínimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   'Ahorro Programado
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
      'Clientes de Provincia no tienen acceso a este Tipo de Evaluación
      'If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
      '   MsgBox "Este tipo de Evaluación sólo está permitida para clientes que residen en Lima Metropolitana o Callao.", vbExclamation, modgen_g_str_NomPlt
      '   Call gs_SetFocus(cmb_InsFin)
      '   Exit Sub
      'End If
   
      If cmb_InsFin.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Institución Financiera donde tiene sus ahorros.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_InsFin)
         Exit Sub
      End If
      
      If cmb_MonAho.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Moneda de su ahorro.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_MonAho)
         Exit Sub
      End If
   
      If ipp_MtoAho.Value = 0 Then
         MsgBox "Debe ingresar el Monto Mínimo Mensual de su Ahorro.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MtoAho)
         Exit Sub
      End If
   
      If ipp_MesAho.Value = 0 Then
         MsgBox "Debe ingresar los Meses Ahorrados.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MesAho)
         Exit Sub
      End If
      
      r_dbl_CuoAho = 0
      If cmb_MonAho.ItemData(cmb_MonAho.ListIndex) = 1 Then
         r_dbl_CuoAho = CDbl(ipp_MtoAho.Text) / l_dbl_TipCam
      ElseIf cmb_MonAho.ItemData(cmb_MonAho.ListIndex) = 2 Then
         r_dbl_CuoAho = CDbl(ipp_MtoAho.Text)
      End If
      r_dbl_CuoAho = CDbl(Format(r_dbl_CuoAho, "###,##0.00"))
      
      r_dbl_Aho_CuoMin = 0
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "001") Then
         r_dbl_Aho_CuoMin = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      If r_dbl_CuoAho < r_dbl_Aho_CuoMin Then
         MsgBox "El Importe de la Cuota Mensual Ahorrada no cumple con el mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MtoAho)
         Exit Sub
      End If
      
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      r_dbl_Aho_ApoTp1 = 0
      
      If moddat_g_str_CodPrd = "002" Then
         If r_dbl_ApoMin < 20 Then
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
         
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
            r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
         End If
      
         If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
            MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MesAho)
            Exit Sub
         End If
      ElseIf moddat_g_str_CodPrd = "003" Then
         If r_dbl_ApoMin >= 20 And r_dbl_ApoMin < 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "013") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 30 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "014") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (20%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      ElseIf moddat_g_str_CodPrd = "004" Then
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (10%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      ElseIf moddat_g_str_CodPrd = "007" Then
         If r_dbl_ApoMin >= 10 And r_dbl_ApoMin < 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "011") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         ElseIf r_dbl_ApoMin >= 20 Then
            If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "052", "012") Then
               r_dbl_Aho_ApoTp1 = l_arr_ParPrd(1).Genera_Cantid
            End If
            
            If CInt(ipp_MesAho.Text) < r_dbl_Aho_ApoTp1 Then
               MsgBox "El Cliente no cumple con el Tiempo Mínimo de Ahorro requerido. (" & CStr(r_dbl_Aho_ApoTp1) & " meses).", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_MesAho)
               Exit Sub
            End If
         
         Else
            MsgBox "El Aporte Propio no cubre el mínimo permitido para el Tipo de Evaluación. (10%).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
   End If
   
   '30%-35%
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 3 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      If modatecli_g_arr_DatInm(1).DatInm_PryMCs = 1 Then
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "001") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      Else
         r_dbl_Ini_ApoMin = 0
         If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "002") Then
            r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
         End If
         
         If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
            MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ApoPro)
            Exit Sub
         End If
      End If
      
      'If moddat_g_str_UbiGeo <> "1501" And moddat_g_str_UbiGeo <> "0701" Then
      '   r_dbl_Ini_ApoMin = 0
      '   If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "053", "004") Then
      '      r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
      '   End If
      '
      '   If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
      '      MsgBox "Cliente de Provincias. El Aporte Inicial es menor al Aporte Inicial mínimo requerido. (" & CStr(r_dbl_Ini_ApoMin) & "%).", vbExclamation, modgen_g_str_NomPlt
      '      Call gs_SetFocus(ipp_ApoPro)
      '      Exit Sub
      '   End If
      'End If
   End If
   
   '50% Inicial Sin Evaluación
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 4 Then
      r_dbl_ApoMin = CDbl(ipp_ApoPro.Text) / CDbl(ipp_ComVta.Text) * 100
      
      r_dbl_Ini_ApoMin = 0
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "054", "001") Then
         r_dbl_Ini_ApoMin = l_arr_ParPrd(1).Genera_Cantid
      End If
      
      If r_dbl_ApoMin < r_dbl_Ini_ApoMin Then
         MsgBox "El Aporte Inicial es menor al Aporte Inicial mínimo requerido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ApoPro)
         Exit Sub
      End If
   End If
   
   If cmb_ConHip.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConHip)
      Exit Sub
   End If
   
   If cmb_EjeSeg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Ejecutivo de Seguimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EjeSeg)
      Exit Sub
   End If
   
   'Validando Documentos a Recibir
   grd_Listad.Redraw = False
   
   r_int_FlgDoc = 1
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 1
      If Trim(grd_Listad.Text) = "X" Then
         r_int_FlgDoc = 2
         Exit For
      End If
   Next r_int_Contad
   
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
   
   If r_int_FlgDoc = 1 Then
      MsgBox "Debe seleccionar los Documentos Crediticios que han sido recibidos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call modatecli_gs_Limpia_DatCre
   
   modatecli_g_arr_DatCre(1).DatCre_TipEva = CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo)
   modatecli_g_arr_DatCre(1).DatCre_TipMon = moddat_g_int_TipMon
   modatecli_g_arr_DatCre(1).DatCre_ComVta = CDbl(ipp_ComVta.Text)
   modatecli_g_arr_DatCre(1).DatCre_ApoPro = CDbl(ipp_ApoPro.Text)
   modatecli_g_arr_DatCre(1).DatCre_MtoPre = CDbl(ipp_MtoPre.Text)
   modatecli_g_arr_DatCre(1).DatCre_TipCam = l_dbl_TipCam
   modatecli_g_arr_DatCre(1).DatCre_ComVta_Sol = CDbl(pnl_ComVta_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_ApoPro_Sol = CDbl(pnl_ApoPro_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoPre_Sol = CDbl(pnl_MtoPre_Sol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_ComVta_Dol = CDbl(pnl_ComVta_Dol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_ApoPro_Dol = CDbl(pnl_ApoPro_Dol.Caption)
   modatecli_g_arr_DatCre(1).DatCre_MtoPre_Dol = CDbl(pnl_MtoPre_Dol.Caption)
   
   modatecli_g_arr_DatCre(1).DatCre_PlaAno = ipp_PlaAno.Value
   modatecli_g_arr_DatCre(1).DatCre_PerGra = ipp_PerGra.Value
   modatecli_g_arr_DatCre(1).DatCre_CuoExt = 2
   
   modatecli_g_arr_DatCre(1).DatCre_ESgDes = l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_TipSeg = cmb_SegDes.ItemData(cmb_SegDes.ListIndex)
   modatecli_g_arr_DatCre(1).DatCre_ESgViv = l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo

   modatecli_g_arr_DatCre(1).DatCre_DiaPag = CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo)
   
   If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
      modatecli_g_arr_DatCre(1).DatCre_InsFin = l_arr_InsFin(cmb_InsFin.ListIndex + 1).Genera_Codigo
      modatecli_g_arr_DatCre(1).DatCre_MonAho = cmb_MonAho.ItemData(cmb_MonAho.ListIndex)
      modatecli_g_arr_DatCre(1).DatCre_MtoAho = CDbl(ipp_MtoAho.Text)
      modatecli_g_arr_DatCre(1).DatCre_MesAho = CDbl(ipp_MesAho.Text)
   Else
      modatecli_g_arr_DatCre(1).DatCre_InsFin = ""
      modatecli_g_arr_DatCre(1).DatCre_MonAho = 0
      modatecli_g_arr_DatCre(1).DatCre_MtoAho = 0
      modatecli_g_arr_DatCre(1).DatCre_MesAho = 0
   End If
   
   modatecli_g_arr_DatCre(1).DatCre_ConHip = l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_EjeSeg = l_arr_EjeSeg(cmb_EjeSeg.ListIndex + 1).Genera_Codigo
   modatecli_g_arr_DatCre(1).DatCre_Observ = txt_Observ.Text

   'Cargando Documentos
   ReDim modatecli_g_arr_DocCre(0)
   
   grd_Listad.Redraw = False
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
   
   grd_Listad.Redraw = True
   
   modatecli_g_int_DatCreTit = 2
   
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
   Dim r_int_Contad     As Integer
   Dim r_int_ConAux     As Integer
   Dim r_int_TipDoc     As Integer
   Dim r_str_CodGrp     As String
   Dim r_int_CodAct     As Integer
   Dim r_str_CodIte     As String
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Produc.Caption = moddat_gf_Consulta_Produc(moddat_g_str_CodPrd)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call fs_Carga_Docume
   
   l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, 2)
   
   ipp_ComVta.Enabled = True
   ipp_ApoPro.Enabled = True
   'ipp_MtoPre.Enabled = True
   ipp_PlaAno.Enabled = True
   ipp_PerGra.Enabled = True
      
   cmb_EmpSeg.Enabled = True
   cmb_SegDes.Enabled = True
   cmb_DiaPag.Enabled = True
   
   If modatecli_g_int_DatCreTit = 2 Then
      'Call gs_BuscarCombo_Item(cmb_TipEva, modatecli_g_arr_DatCre(1).DatCre_TipEva)
      
      cmb_TipEva.ListIndex = gf_Busca_Arregl(l_arr_TipEva, Format(modatecli_g_arr_DatCre(1).DatCre_TipEva, "000")) - 1
      
      ipp_ComVta.Value = modatecli_g_arr_DatCre(1).DatCre_ComVta
      ipp_ApoPro.Value = modatecli_g_arr_DatCre(1).DatCre_ApoPro
      ipp_MtoPre.Value = modatecli_g_arr_DatCre(1).DatCre_MtoPre
      
      Call fs_Calcul
   
      ipp_PlaAno.Value = modatecli_g_arr_DatCre(1).DatCre_PlaAno
      ipp_PerGra.Value = modatecli_g_arr_DatCre(1).DatCre_PerGra
      
      cmb_EmpSeg.ListIndex = gf_Busca_Arregl(l_arr_EmpSeg, modatecli_g_arr_DatCre(1).DatCre_ESgDes) - 1
      Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo)
      Call gs_BuscarCombo_Item(cmb_SegDes, modatecli_g_arr_DatCre(1).DatCre_TipSeg)
      
      cmb_DiaPag.ListIndex = gf_Busca_Arregl(l_arr_DiaPag, Format(modatecli_g_arr_DatCre(1).DatCre_DiaPag, "000")) - 1
      
      If CInt(l_arr_TipEva(cmb_TipEva.ListIndex + 1).Genera_Codigo) = 2 Then
         cmb_InsFin.ListIndex = gf_Busca_Arregl(l_arr_InsFin, modatecli_g_arr_DatCre(1).DatCre_InsFin) - 1
         Call gs_BuscarCombo_Item(cmb_MonAho, modatecli_g_arr_DatCre(1).DatCre_MonAho)
         ipp_MtoAho.Text = modatecli_g_arr_DatCre(1).DatCre_MtoAho
         ipp_MesAho.Text = modatecli_g_arr_DatCre(1).DatCre_MesAho
         
         cmb_InsFin.Enabled = True
         cmb_MonAho.Enabled = True
         ipp_MtoAho.Enabled = True
         ipp_MesAho.Enabled = True
      End If
      
      cmb_ConHip.ListIndex = gf_Busca_Arregl(l_arr_ConHip, modatecli_g_arr_DatCre(1).DatCre_ConHip) - 1
      cmb_EjeSeg.ListIndex = gf_Busca_Arregl(l_arr_EjeSeg, modatecli_g_arr_DatCre(1).DatCre_EjeSeg) - 1
   
      txt_Observ.Text = modatecli_g_arr_DatCre(1).DatCre_Observ
   
      grd_Listad.Redraw = False
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
   
      grd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Call moddat_gs_Carga_LisIte_Combo(cmb_TipEva, 1, "038")
   
   Call moddat_gs_Carga_ParSubPrd(cmb_TipEva, l_arr_TipEva(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "014")

   'Plazo de Crédito
   If moddat_gf_Consulta_SubPrd_Arregl(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub) Then
      ipp_PlaAno.MinValue = moddat_g_arr_Genera(1).Genera_PlzMin
      ipp_PlaAno.MaxValue = moddat_g_arr_Genera(1).Genera_PlzMax
   End If
   
   'Periodo de Gracia
   l_int_GraMax = 0
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "008", "002") Then
      ipp_PerGra.MinValue = moddat_g_arr_Genera(1).Genera_ValMin
      ipp_PerGra.MaxValue = moddat_g_arr_Genera(1).Genera_ValMax
      
      l_int_GraMax = moddat_g_arr_Genera(1).Genera_ValMax
   End If
   
   
   Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
      
   Call moddat_gs_Carga_EmpSeg(cmb_EmpSeg, l_arr_EmpSeg)
   
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call moddat_gs_Carga_EjecMC(cmb_EjeSeg, l_arr_EjeSeg, 131)

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

   'Ahorro Programado
   Call moddat_gs_Carga_LisIte(cmb_InsFin, l_arr_InsFin, 1, "505")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonAho, 1, "204")
   
End Sub

Private Sub fs_Limpia()
   ipp_ComVta.Value = 0
   ipp_ApoPro.Value = 0
   ipp_MtoPre.Value = 0
   
   pnl_ComVta_Sol.Caption = "0.00 "
   pnl_ApoPro_Sol.Caption = "0.00 "
   pnl_MtoPre_Sol.Caption = "0.00 "
   
   pnl_ComVta_Dol.Caption = "0.00 "
   pnl_ApoPro_Dol.Caption = "0.00 "
   pnl_MtoPre_Dol.Caption = "0.00 "
   
   ipp_PlaAno.Value = ipp_PlaAno.MinValue
   ipp_PerGra.Value = 0
   cmb_EmpSeg.ListIndex = -1
   cmb_SegDes.Clear
   cmb_DiaPag.ListIndex = -1
   
   ipp_ComVta.Enabled = False
   ipp_ApoPro.Enabled = False
   'ipp_MtoPre.Enabled = False
   ipp_PlaAno.Enabled = False
   ipp_PerGra.Enabled = False
   cmb_EmpSeg.Enabled = False
   cmb_SegDes.Enabled = False
   cmb_DiaPag.Enabled = False
   
   txt_Observ.Text = ""
   
   cmb_ConHip.ListIndex = -1
   cmb_EjeSeg.ListIndex = -1
   
   Call gs_LimpiaGrid(grd_Listad)
   
   cmb_InsFin.ListIndex = -1
   cmb_MonAho.ListIndex = -1
   ipp_MtoAho.Value = 0
   ipp_MesAho.Value = 0

   cmb_InsFin.Enabled = False
   cmb_MonAho.Enabled = False
   ipp_MtoAho.Enabled = False
   ipp_MesAho.Enabled = False
End Sub

Private Sub ipp_ApoPro_Change()
   ipp_MtoPre.Value = CDbl(ipp_ComVta.Text) - CDbl(ipp_ApoPro.Text)
   
   Call fs_Calcul
End Sub

Private Sub ipp_ApoPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_ComVta_Change()
   ipp_MtoPre.Value = CDbl(ipp_ComVta.Text) - CDbl(ipp_ApoPro.Text)
   
   Call fs_Calcul
End Sub

Private Sub ipp_ComVta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ApoPro)
   End If
End Sub

Private Sub ipp_MesAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub ipp_MtoAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MesAho)
   End If
End Sub

Private Sub ipp_MtoPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PlaAno)
   End If
End Sub

Private Sub ipp_PerGra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_EmpSeg)
   End If
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerGra)
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ConHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
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

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Carga_Docume()
   Dim r_int_ActPri_Cli    As Integer
   Dim r_int_ActSec_Cli    As Integer
   Dim r_int_ActPri_Cyg    As Integer
   Dim r_int_ActSec_Cyg    As Integer
   
   '0 - Descripción
   '1 - Selección
   '2 - Tipo de Origen de Documento
   '3 - Código de Grupo
   '4 - Código de Actividad Económica
   '5 - Código de Item
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'Documentos Crediticios
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '011' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARPRD_DESCRI)
         grd_Listad.Col = 1:     grd_Listad.Text = ""
         grd_Listad.Col = 2:     grd_Listad.Text = "1"
         grd_Listad.Col = 3:     grd_Listad.Text = "011"
         grd_Listad.Col = 4:     grd_Listad.Text = "0"
         grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARPRD_CODITE
         
         g_rst_Genera.MoveNext
      Loop
      
      grd_Listad.Redraw = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   r_int_ActPri_Cli = moddat_gf_Consulta_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   r_int_ActSec_Cli = moddat_gf_Consulta_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2)
   
   r_int_ActPri_Cyg = moddat_gf_Consulta_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   r_int_ActSec_Cyg = moddat_gf_Consulta_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
   
   
   'Documentos por Actividad Económica Titular - Actividad Principal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(r_int_ActPri_Cli) & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '002' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
         grd_Listad.Col = 1:     grd_Listad.Text = ""
         grd_Listad.Col = 2:     grd_Listad.Text = "2"
         grd_Listad.Col = 3:     grd_Listad.Text = "002"
         grd_Listad.Col = 4:     grd_Listad.Text = r_int_ActPri_Cli
         grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
         
         g_rst_Genera.MoveNext
      Loop
      
      grd_Listad.Redraw = True
   End If
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   
   'Documentos por Actividad Económica Titular - Actividad Secundaria
   If r_int_ActPri_Cli <> r_int_ActSec_Cli And r_int_ActSec_Cli > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
      g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(r_int_ActSec_Cli) & " AND "
      g_str_Parame = g_str_Parame & "PARACT_CODGRP = '002' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
      g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         grd_Listad.Redraw = False
         
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
            grd_Listad.Col = 1:     grd_Listad.Text = ""
            grd_Listad.Col = 2:     grd_Listad.Text = "2"
            grd_Listad.Col = 3:     grd_Listad.Text = "002"
            grd_Listad.Col = 4:     grd_Listad.Text = r_int_ActSec_Cli
            grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
            
            g_rst_Genera.MoveNext
         Loop
         
         grd_Listad.Redraw = True
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   'Documentos por Actividad Económica Cónyuge - Actividad Principal
   If r_int_ActPri_Cyg > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
      g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(r_int_ActPri_Cyg) & " AND    "
      g_str_Parame = g_str_Parame & "PARACT_CODGRP = '003' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
      g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         grd_Listad.Redraw = False
         
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
            grd_Listad.Col = 1:     grd_Listad.Text = ""
            grd_Listad.Col = 2:     grd_Listad.Text = "2"
            grd_Listad.Col = 3:     grd_Listad.Text = "003"
            grd_Listad.Col = 4:     grd_Listad.Text = r_int_ActPri_Cyg
            grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
            
            g_rst_Genera.MoveNext
         Loop
         
         grd_Listad.Redraw = True
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   'Documentos por Actividad Económica Cónyuge - Actividad Secundaria
   If r_int_ActPri_Cyg <> r_int_ActSec_Cyg And r_int_ActPri_Cyg > 0 And r_int_ActSec_Cyg > 0 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
      g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & moddat_g_str_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODACT = " & CStr(r_int_ActSec_Cyg) & " AND    "
      g_str_Parame = g_str_Parame & "PARACT_CODGRP = '003' AND "
      g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
      g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         grd_Listad.Redraw = False
         
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            grd_Listad.Rows = grd_Listad.Rows + 1
            
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = Trim(g_rst_Genera!PARACT_DESCRI)
            grd_Listad.Col = 1:     grd_Listad.Text = ""
            grd_Listad.Col = 2:     grd_Listad.Text = "2"
            grd_Listad.Col = 3:     grd_Listad.Text = "003"
            grd_Listad.Col = 4:     grd_Listad.Text = r_int_ActSec_Cyg
            grd_Listad.Col = 5:     grd_Listad.Text = g_rst_Genera!PARACT_CODITE
            
            g_rst_Genera.MoveNext
         Loop
         
         grd_Listad.Redraw = True
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Calcul()
   If moddat_g_int_TipMon = 1 Then
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text), "###,###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text), "###,###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(ipp_MtoPre.Text), "###,###,##0.00") & " "
      
      pnl_ComVta_Dol.Caption = Format(CDbl(ipp_ComVta.Text) / l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ApoPro_Dol.Caption = Format(CDbl(ipp_ApoPro.Text) / l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoPre_Dol.Caption = Format(CDbl(ipp_MtoPre.Text) / l_dbl_TipCam, "###,###,##0.00") & " "
   Else
      pnl_ComVta_Sol.Caption = Format(CDbl(ipp_ComVta.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_ApoPro_Sol.Caption = Format(CDbl(ipp_ApoPro.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      pnl_MtoPre_Sol.Caption = Format(CDbl(ipp_MtoPre.Text) * l_dbl_TipCam, "###,###,##0.00") & " "
      
      pnl_ComVta_Dol.Caption = Format(CDbl(ipp_ComVta.Text), "###,###,##0.00") & " "
      pnl_ApoPro_Dol.Caption = Format(CDbl(ipp_ApoPro.Text), "###,###,##0.00") & " "
      pnl_MtoPre_Dol.Caption = Format(CDbl(ipp_MtoPre.Text), "###,###,##0.00") & " "
   End If
End Sub
