VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pla_Aho_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   2010
   ClientTop       =   3165
   ClientWidth     =   10815
   Icon            =   "AteCli_frm_555.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel4 
      Height          =   9030
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   10845
      _Version        =   65536
      _ExtentX        =   19129
      _ExtentY        =   15928
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   33
         Top             =   60
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   690
            TabIndex        =   34
            Top             =   225
            Width           =   6570
            _Version        =   65536
            _ExtentX        =   11589
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Mantenimiento de Planes de Ahorro"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_555.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   35
         Top             =   780
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10140
            Picture         =   "AteCli_frm_555.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   45
            Picture         =   "AteCli_frm_555.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   45
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2010
         Left            =   30
         TabIndex        =   36
         Top             =   1500
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
         _ExtentY        =   3545
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
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   2925
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   1830
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1230
            Width           =   1455
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   7170
            MaxLength       =   8
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   150
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef3 
            Height          =   315
            Left            =   7170
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1230
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3315
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1230
            Width           =   1455
         End
         Begin VB.TextBox txt_Email 
            Height          =   315
            Left            =   1830
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1575
            Width           =   8640
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1830
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   885
            Width           =   2925
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1830
            MaxLength       =   30
            TabIndex        =   2
            Top             =   555
            Width           =   2925
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   7170
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   555
            Width           =   3300
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo de Documento:"
            Height          =   315
            Left            =   90
            TabIndex        =   59
            Top             =   180
            Width           =   1710
         End
         Begin VB.Label lbl_General 
            Caption         =   "Email:"
            Height          =   285
            Index           =   39
            Left            =   105
            TabIndex        =   43
            Top             =   1635
            Width           =   1710
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   5265
            TabIndex        =   42
            Top             =   210
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfonos Fijo/Movil:"
            Height          =   285
            Index           =   46
            Left            =   105
            TabIndex        =   41
            Top             =   1290
            Width           =   1710
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono Trabajo:"
            Height          =   285
            Index           =   55
            Left            =   5265
            TabIndex        =   40
            Top             =   1290
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   105
            TabIndex        =   39
            Top             =   945
            Width           =   1710
         End
         Begin VB.Label Label1 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   105
            TabIndex        =   38
            Top             =   615
            Width           =   1710
         End
         Begin VB.Label Label6 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   5265
            TabIndex        =   37
            Top             =   615
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1935
         Left            =   30
         TabIndex        =   44
         Top             =   3555
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
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
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   150
            Width           =   2925
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1830
            MaxLength       =   120
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   480
            Width           =   2925
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   7170
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   480
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   8820
            MaxLength       =   15
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   480
            Width           =   1665
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   810
            Width           =   2925
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   7170
            MaxLength       =   120
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   810
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1830
            TabIndex        =   15
            Text            =   "cmb_DptDir"
            Top             =   1140
            Width           =   2925
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   7170
            TabIndex        =   16
            Top             =   1140
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1830
            TabIndex        =   17
            Top             =   1470
            Width           =   2925
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   7170
            MaxLength       =   250
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1470
            Width           =   3315
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   105
            TabIndex        =   53
            Top             =   150
            Width           =   1710
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   105
            TabIndex        =   52
            Top             =   480
            Width           =   1710
         End
         Begin VB.Label Label21 
            Caption         =   "Nro/Mza/Lote - Int/Dpto:"
            Height          =   285
            Left            =   5265
            TabIndex        =   51
            Top             =   480
            Width           =   1800
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   105
            TabIndex        =   50
            Top             =   810
            Width           =   1710
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   5265
            TabIndex        =   49
            Top             =   810
            Width           =   1800
         End
         Begin VB.Label Label3 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   105
            TabIndex        =   48
            Top             =   1140
            Width           =   1710
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   5265
            TabIndex        =   47
            Top             =   1140
            Width           =   1800
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   105
            TabIndex        =   46
            Top             =   1470
            Width           =   1710
         End
         Begin VB.Label Label4 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   5265
            TabIndex        =   45
            Top             =   1470
            Width           =   1800
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3390
         Left            =   30
         TabIndex        =   54
         Top             =   5535
         Width           =   10755
         _Version        =   65536
         _ExtentX        =   18971
         _ExtentY        =   5980
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
         Begin VB.ComboBox cmb_SubPrd 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   480
            Width           =   8655
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   150
            Width           =   8655
         End
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2970
            Width           =   3225
         End
         Begin VB.ComboBox cmb_EjeSeg 
            Height          =   315
            Left            =   7170
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2970
            Width           =   3315
         End
         Begin VB.ComboBox cmb_InsFin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7170
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   900
            Visible         =   0   'False
            Width           =   3315
         End
         Begin VB.ComboBox cmb_MonAho 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1230
            Width           =   2925
         End
         Begin VB.ComboBox cmb_DiaPag 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   900
            Width           =   2925
         End
         Begin EditLib.fpDoubleSingle ipp_MtoAho 
            Height          =   315
            Left            =   1830
            TabIndex        =   25
            Top             =   1950
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
            Left            =   1830
            TabIndex        =   26
            Top             =   2280
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            MaxValue        =   "18"
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
         Begin EditLib.fpDoubleSingle ipp_TotAho 
            Height          =   315
            Left            =   1830
            TabIndex        =   27
            Top             =   2610
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ControlType     =   1
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
         Begin EditLib.fpDateTime ipp_IniVct 
            Height          =   315
            Left            =   1830
            TabIndex        =   24
            Top             =   1605
            Width           =   1650
            _Version        =   196608
            _ExtentX        =   2910
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
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Primer Vcto.:"
            Height          =   270
            Index           =   58
            Left            =   105
            TabIndex        =   66
            Top             =   1620
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Ejecutivo de Seguimiento:"
            Height          =   315
            Left            =   5265
            TabIndex        =   65
            Top             =   3030
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Institución Financiera:"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5265
            TabIndex        =   64
            Top             =   930
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label16 
            Caption         =   "Sub-Producto:"
            Height          =   315
            Left            =   105
            TabIndex        =   63
            Top             =   510
            Width           =   1710
         End
         Begin VB.Label Label15 
            Caption         =   "Consejero Hipotecario:"
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   3030
            Width           =   1710
         End
         Begin VB.Label Label13 
            Caption         =   "Total a Ahorrar:"
            Height          =   285
            Left            =   105
            TabIndex        =   61
            Top             =   2685
            Width           =   1710
         End
         Begin VB.Label Label12 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   105
            TabIndex        =   60
            Top             =   210
            Width           =   1710
         End
         Begin VB.Label Label8 
            Caption         =   "Moneda de Ahorro:"
            Height          =   315
            Left            =   105
            TabIndex        =   58
            Top             =   1275
            Width           =   1710
         End
         Begin VB.Label Label18 
            Caption         =   "Monto de Ahorro:"
            Height          =   285
            Left            =   105
            TabIndex        =   57
            Top             =   1980
            Width           =   1710
         End
         Begin VB.Label Label9 
            Caption         =   "Meses Ahorrados:"
            Height          =   285
            Left            =   105
            TabIndex        =   56
            Top             =   2340
            Width           =   1710
         End
         Begin VB.Label Label10 
            Caption         =   "Día de Pago:"
            Height          =   315
            Left            =   105
            TabIndex        =   55
            Top             =   930
            Width           =   1710
         End
      End
   End
End
Attribute VB_Name = "frm_Pla_Aho_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()   As moddat_tpo_Genera
Dim l_arr_SubPrd()   As moddat_tpo_Genera
Dim l_arr_ConHip()   As moddat_tpo_Genera
Dim l_arr_EjeSeg()   As moddat_tpo_Genera
Dim l_arr_InsFin()   As moddat_tpo_Genera
Dim l_arr_DiaPag()   As moddat_tpo_Genera
Dim l_int_FlgCmb     As Integer
Dim l_str_PrvDir     As String
Dim l_str_DptDir     As String
Dim r_str_FecRec     As String
Dim l_str_DstDir     As String
Dim r_str_Numero     As String

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "236")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_Produc_Comerc(cmb_Produc, l_arr_Produc, 4)
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonAho, 1, "204")
   Call moddat_gs_Carga_LisIte(cmb_InsFin, l_arr_InsFin, 1, "505")
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
   Call moddat_gs_Carga_EjecMC(cmb_EjeSeg, l_arr_EjeSeg, 131)
   txt_NumDoc.MaxLength = 8
End Sub

Private Sub fs_bloquea(ByVal est As Boolean)
    cmb_Produc.Enabled = est
    cmb_SubPrd.Enabled = est
    cmb_DiaPag.Enabled = est
    cmb_InsFin.Enabled = est
    cmb_MonAho.Enabled = est
    ipp_MtoAho.Enabled = est
    ipp_MesAho.Enabled = est
    ipp_TotAho.Enabled = est
    ipp_IniVct.Enabled = est
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_Telef3.Text = ""
   txt_Email.Text = ""
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.ListIndex = -1
   cmb_DstDir.ListIndex = -1
   txt_Refere.Text = ""
   cmb_Produc.ListIndex = -1
   cmb_SubPrd.ListIndex = -1
   cmb_DiaPag.ListIndex = -1
   cmb_InsFin.ListIndex = -1
   cmb_MonAho.ListIndex = -1
   ipp_IniVct.Text = Format(date, "DD/MM/YYYY")
   ipp_MtoAho.Value = 0
   ipp_MesAho.Value = 0
   ipp_TotAho.Value = 0
   cmb_ConHip.ListIndex = -1
   cmb_EjeSeg.ListIndex = -1
End Sub

Private Sub fs_Buscar()
   'Si selecciono adicionar
   If modmip_g_int_FlgGrb_1 = 1 Then
       cmb_TipDoc.Enabled = True
       txt_NumDoc.Enabled = True
       Exit Sub
   End If
   
   'Busca datos de la operacion
   g_str_Parame = "SELECT * "
   g_str_Parame = g_str_Parame & " FROM CRE_AHOMAE "
   g_str_Parame = g_str_Parame & "WHERE AHOMAE_NUMERO = '" & Trim(moddat_g_str_NumOpe) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontro datos de la operación seleccionada.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If Not g_rst_Princi.EOF Then
      'Si es modificacion valida que la operacion este vigente
      If modmip_g_int_FlgGrb_1 = 2 Then
         If CInt(g_rst_Princi!AHOMAE_SITUAC) <> 2 Then
            MsgBox "Operación seleccionada no se encuentra en estado VIGENTE.", vbExclamation, modgen_g_str_NomPlt
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
            Unload Me
         End If
      End If
      
      'Muestra datos
      r_str_Numero = Trim(g_rst_Princi!AHOMAE_NUMERO & "")
      ipp_IniVct.Text = gf_FormatoFecha(CStr(Trim(g_rst_Princi!AHOMAE_PRIVCT)))
      r_str_FecRec = Trim(g_rst_Princi!AHOMAE_PRIVCT)
      cmb_Produc.ListIndex = gf_Busca_Arregl(l_arr_Produc, Trim(g_rst_Princi!AHOMAE_CODPRD & "")) - 1
      
      cmb_SubPrd.Clear
      moddat_g_str_CodPrd = ""
      moddat_g_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
      Call moddat_gs_Carga_SubPrd(cmb_SubPrd, l_arr_SubPrd, moddat_g_str_CodPrd)
                              
      cmb_SubPrd.ListIndex = gf_Busca_Arregl(l_arr_SubPrd, Trim(g_rst_Princi!AHOMAE_SUBPRD & "")) - 1
      moddat_g_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
      Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
      cmb_DiaPag.ListIndex = gf_Busca_Arregl(l_arr_DiaPag, Format(g_rst_Princi!AHOMAE_DIAPAG & "", "000")) - 1
      cmb_InsFin.ListIndex = gf_Busca_Arregl(l_arr_InsFin, Format(g_rst_Princi!AHOMAE_CODBAN & "", "000000")) - 1
      ipp_IniVct.Text = gf_FormatoFecha(CStr(Trim(g_rst_Princi!AHOMAE_PRIVCT)))
      
      Call gs_BuscarCombo_Item(cmb_MonAho, CInt(g_rst_Princi!AHOMAE_MONAHO))
      ipp_MtoAho.Value = Trim(g_rst_Princi!AHOMAE_MTOAHO & "")
      ipp_MesAho.Value = Trim(g_rst_Princi!AHOMAE_NUMMES & "")
      ipp_TotAho.Value = Trim(g_rst_Princi!AHOMAE_TOTAHO & "")
      
      cmb_ConHip.ListIndex = gf_Busca_Arregl(l_arr_ConHip, Trim(g_rst_Princi!AHOMAE_CONHIP & "")) - 1
      cmb_EjeSeg.ListIndex = gf_Busca_Arregl(l_arr_EjeSeg, Trim(g_rst_Princi!AHOMAE_EJESEG & "")) - 1
      
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      Call fs_bloquea(False)
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Busca datos del cliente
   g_str_Parame = "SELECT * "
   g_str_Parame = g_str_Parame & " FROM CRE_AHOCLI  "
   g_str_Parame = g_str_Parame & "WHERE AHOCLI_TIPCLI = 1  "
   g_str_Parame = g_str_Parame & "  AND AHOCLI_TIPDOC = " & CStr(moddat_g_str_TipDoc) & " "
   g_str_Parame = g_str_Parame & "  AND AHOCLI_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se ha encontrado datos del cliente para la operación seleccionada.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If Not g_rst_Princi.EOF Then
      Call gs_BuscarCombo_Item(cmb_TipDoc, CInt(Left(g_rst_Princi!AHOCLI_TIPDOC, 2)))
      txt_NumDoc.Text = Trim(g_rst_Princi!AHOCLI_NUMDOC & "")
      txt_ApePat.Text = Trim(g_rst_Princi!AHOCLI_APEPAT & "")
      txt_ApeMat.Text = Trim(g_rst_Princi!AHOCLI_APEMAT & "")
      txt_Nombre.Text = Trim(g_rst_Princi!AHOCLI_NOMBRE & "")
      txt_Telef1.Text = Trim(g_rst_Princi!AHOCLI_TELFIJ & "")
      txt_Telef2.Text = Trim(g_rst_Princi!AHOCLI_TELTRA & "")
      txt_Telef3.Text = Trim(g_rst_Princi!AHOCLI_TELCEL & "")
      txt_Email.Text = Trim(g_rst_Princi!AHOCLI_DIRELE & "")
      
      Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!AHOCLI_TIPVIA)
      txt_NomVia.Text = Trim(g_rst_Princi!AHOCLI_NOMVIA & "")
      txt_NumVia.Text = Trim(g_rst_Princi!AHOCLI_NUMVIA & "")
      txt_IntDpt.Text = Trim(g_rst_Princi!AHOCLI_INTDPT & "")
      Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!AHOCLI_TIPZON)
      txt_NomZon.Text = Trim(g_rst_Princi!AHOCLI_NOMZON & "")

      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!AHOCLI_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!AHOCLI_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!AHOCLI_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!AHOCLI_UBIGEO, 2), Mid(g_rst_Princi!AHOCLI_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!AHOCLI_UBIGEO, 2)))
      txt_Refere.Text = Trim(g_rst_Princi!AHOCLI_REFERE & "")
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Si es consulta se bloquea campos
   If modmip_g_int_FlgGrb_1 = 3 Then
      Call fs_bloquea(False)
      cmb_ConHip.Enabled = False
      cmb_EjeSeg.Enabled = False
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      txt_ApePat.Enabled = False
      txt_ApeMat.Enabled = False
      txt_Nombre.Enabled = False
      txt_Telef1.Enabled = False
      txt_Telef2.Enabled = False
      txt_Telef3.Enabled = False
      txt_Email.Enabled = False
      cmb_TipVia.Enabled = False
      txt_NomVia.Enabled = False
      txt_NumVia.Enabled = False
      txt_IntDpt.Enabled = False
      cmb_TipZon.Enabled = False
      txt_NomZon.Enabled = False
      cmb_DptDir.Enabled = False
      cmb_PrvDir.Enabled = False
      cmb_DstDir.Enabled = False
      txt_Refere.Enabled = False
      cmd_Grabar.Visible = False
      cmd_Grabar.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicio
   Call fs_Limpia
   Call fs_bloquea(True)
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   'Validaciones
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Numero de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ListIndex = 3 Then
      MsgBox "La Opción del Numero de Ruc no esta permitida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If cmb_TipDoc.ListIndex = 0 Or cmb_TipDoc.ListIndex = 1 Or cmb_TipDoc.ListIndex = 2 Then
      If Len(Trim(txt_NumDoc.Text)) <> 8 Then
         MsgBox "El Numero de Documento debe tener 8 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Sub
   End If
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   If cmb_TipVia.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipVia)
      Exit Sub
   End If
   If Len(Trim(txt_NomVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomVia)
      Exit Sub
   End If
   If Len(Trim(txt_NumVia.Text)) = 0 Then
      MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumVia)
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
      MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DptDir)
      Exit Sub
   End If
   If cmb_PrvDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvDir)
      Exit Sub
   End If
   If cmb_DstDir.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DstDir)
      Exit Sub
   End If
      
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   If cmb_SubPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar el SubProducto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SubPrd)
      Exit Sub
   End If
   If cmb_DiaPag.ListIndex = -1 Then
      MsgBox "Debe seleccionar el dia de pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DiaPag)
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
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If modmip_g_int_FlgGrb_1 = 1 Then
      r_str_FecRec = Format(ipp_IniVct.Text, "YYYYMMDD")
      r_str_Numero = Year(date) & Trim(txt_NumDoc.Text)
   End If
   
   'Grabando Información del Cliente, el maestro de ahorro y cuotas
   g_str_Parame = "USP_CRE_AHOMAE ("
   g_str_Parame = g_str_Parame & "'" & r_str_FecRec & "', "
   g_str_Parame = g_str_Parame & "'" & r_str_Numero & "', "
   g_str_Parame = g_str_Parame & "1,"
   g_str_Parame = g_str_Parame & "" & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_NumDoc.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApePat & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApeMat & "', "
   g_str_Parame = g_str_Parame & "'" & txt_Nombre & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
   g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Telef1.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Telef2.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Telef3.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Email.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo & "', "  '' CARACTER 3
   g_str_Parame = g_str_Parame & "'" & l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo & "', "  '' CARACTER 3
   g_str_Parame = g_str_Parame & "" & CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo) & ", "  '' CARACTER 3
   g_str_Parame = g_str_Parame & "" & CInt(ipp_MesAho.Text) & ", "
   g_str_Parame = g_str_Parame & "" & CDbl(ipp_MtoAho.Text) & ", "
   g_str_Parame = g_str_Parame & "" & CDbl(ipp_TotAho.Text) & ", "
   g_str_Parame = g_str_Parame & "" & CStr(cmb_MonAho.ItemData(cmb_MonAho.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "0, "  '' CARACTER 6
   g_str_Parame = g_str_Parame & "'0', "
   g_str_Parame = g_str_Parame & "'" & l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & l_arr_EjeSeg(cmb_EjeSeg.ListIndex + 1).Genera_Codigo & "', "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   
   If modmip_g_int_FlgGrb_1 = 1 Then
      g_str_Parame = g_str_Parame & "" & CInt(modmip_g_int_FlgGrb_1) & ") "
   Else
      g_str_Parame = g_str_Parame & "" & CInt(modmip_g_int_FlgAct_1) & ") "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CRE_AHOMAE.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
    
   'Busca si el cliente tiene una operacion vigente
   g_str_Parame = "SELECT AHOMAE_NUMERO "
   g_str_Parame = g_str_Parame & " FROM CRE_AHOMAE  "
   g_str_Parame = g_str_Parame & "WHERE AHOMAE_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " "
   g_str_Parame = g_str_Parame & "  AND AHOMAE_NUMDOC = '" & Trim(txt_NumDoc.Text) & "' "
   g_str_Parame = g_str_Parame & "  AND AHOMAE_SITUAC = '2' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Unload Me
   End If
   
   MsgBox "Se genero la siguiente operación: " & Mid(g_rst_Princi!AHOMAE_NUMERO, 1, 4) & "-" & Mid(g_rst_Princi!AHOMAE_NUMERO, 5, 8) & "-" & Mid(g_rst_Princi!AHOMAE_NUMERO, 13, 3) & "  ", vbExclamation, modgen_g_str_NomPlt
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub fs_BuscaDatCli(ByVal tipdoc As Integer, ByVal numdoc As String)
   'Busca datos del cliente en maestro de clientes ahorro
   g_str_Parame = "SELECT * "
   g_str_Parame = g_str_Parame & " FROM CRE_AHOCLI  "
   g_str_Parame = g_str_Parame & "WHERE AHOCLI_TIPCLI = 1 "
   g_str_Parame = g_str_Parame & "  AND AHOCLI_TIPDOC = " & CStr(tipdoc) & " "
   g_str_Parame = g_str_Parame & "  AND AHOCLI_NUMDOC = '" & Trim(numdoc) & "' "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      txt_ApePat.Text = ""
      txt_ApeMat.Text = ""
      txt_Nombre.Text = ""
      txt_Telef1.Text = ""
      txt_Telef2.Text = ""
      txt_Telef3.Text = ""
      txt_Email.Text = ""
      cmb_TipVia.ListIndex = -1
      txt_NomVia.Text = ""
      txt_NumVia.Text = ""
      txt_IntDpt.Text = ""
      cmb_TipZon.ListIndex = -1
      txt_NomZon.Text = ""
      cmb_DptDir.ListIndex = -1
      cmb_PrvDir.ListIndex = -1
      cmb_DstDir.ListIndex = -1
      txt_Refere.Text = ""
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If Not g_rst_Princi.EOF Then
      txt_ApePat.Text = Trim(g_rst_Princi!AHOCLI_APEPAT & "")
      txt_ApeMat.Text = Trim(g_rst_Princi!AHOCLI_APEMAT & "")
      txt_Nombre.Text = Trim(g_rst_Princi!AHOCLI_NOMBRE & "")
      txt_Telef1.Text = Trim(g_rst_Princi!AHOCLI_TELFIJ & "")
      txt_Telef2.Text = Trim(g_rst_Princi!AHOCLI_TELTRA & "")
      txt_Telef3.Text = Trim(g_rst_Princi!AHOCLI_TELCEL & "")
      txt_Email.Text = Trim(g_rst_Princi!AHOCLI_DIRELE & "")
      
      Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!AHOCLI_TIPVIA)
      txt_NomVia.Text = Trim(g_rst_Princi!AHOCLI_NOMVIA & "")
      txt_NumVia.Text = Trim(g_rst_Princi!AHOCLI_NUMVIA & "")
      txt_IntDpt.Text = Trim(g_rst_Princi!AHOCLI_INTDPT & "")
   
      Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!AHOCLI_TIPZON)
      txt_NomZon.Text = Trim(g_rst_Princi!AHOCLI_NOMZON & "")

      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!AHOCLI_UBIGEO, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!AHOCLI_UBIGEO, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!AHOCLI_UBIGEO, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!AHOCLI_UBIGEO, 2), Mid(g_rst_Princi!AHOCLI_UBIGEO, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!AHOCLI_UBIGEO, 2)))
      txt_Refere.Text = Trim(g_rst_Princi!AHOCLI_REFERE & "")
      Call gs_SetFocus(cmb_Produc)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Busca si el cliente tiene una operacion vigente
   g_str_Parame = "SELECT AHOMAE_NUMERO "
   g_str_Parame = g_str_Parame & " FROM CRE_AHOMAE  "
   g_str_Parame = g_str_Parame & "WHERE AHOMAE_TIPDOC = " & CStr(tipdoc) & " "
   g_str_Parame = g_str_Parame & "  AND AHOMAE_NUMDOC = '" & Trim(numdoc) & "' "
   g_str_Parame = g_str_Parame & "  AND AHOMAE_SITUAC = '2' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      MsgBox "El cliente tiene una operación registrada: " & Mid(g_rst_Princi!AHOMAE_NUMERO, 1, 4) & "-" & Mid(g_rst_Princi!AHOMAE_NUMERO, 5, 8) & "-" & Mid(g_rst_Princi!AHOMAE_NUMERO, 13, 3), vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Unload Me
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 8
         Case 3:  txt_NumDoc.MaxLength = 8
      End Select
   End If
   If modmip_g_int_FlgGrb_1 = 1 Then
       Call gs_SetFocus(txt_NumDoc)
   Else
       Call gs_SetFocus(txt_ApePat)
   End If
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
    Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApePat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_LostFocus()
    If cmb_TipDoc.ListIndex <> -1 Then
        Call fs_BuscaDatCli(CInt(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)), txt_NumDoc.Text)
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
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Telef1_GotFocus()
    Call gs_SelecTodo(txt_Telef1)
End Sub

Private Sub txt_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef2_GotFocus()
    Call gs_SelecTodo(txt_Telef2)
End Sub

Private Sub txt_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef3)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef3_GotFocus()
    Call gs_SelecTodo(txt_Telef3)
End Sub

Private Sub txt_Telef3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Email)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Email_GotFocus()
    Call gs_SelecTodo(txt_Email)
End Sub

Private Sub txt_Email_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. @,;:()/º")
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Click
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
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

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Produc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub cmb_Produc_Click()
   cmb_SubPrd.Clear
   moddat_g_str_CodPrd = ""
   
   If cmb_Produc.ListIndex > -1 Then
      Screen.MousePointer = 11
      moddat_g_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
      Call moddat_gs_Carga_SubPrd(cmb_SubPrd, l_arr_SubPrd, moddat_g_str_CodPrd)
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
   If cmb_SubPrd.ListIndex > -1 Then
      moddat_g_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
      Call moddat_gs_Carga_ParSubPrd(cmb_DiaPag, l_arr_DiaPag(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "009")
      Call gs_SetFocus(cmb_DiaPag)
   End If
End Sub

Private Sub cmb_SubPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SubPrd_Click
   End If
End Sub

Private Sub cmb_DiaPag_Click()
Dim fecvcto As String
    If cmb_DiaPag.ListIndex > -1 Then
       Call gs_SetFocus(cmb_InsFin)
       fecvcto = Format(CInt(l_arr_DiaPag(cmb_DiaPag.ListIndex + 1).Genera_Codigo), "00") & "/" & Format(IIf(Month(date) = 12, 1, Month(date) + 1), "00") & "/" & Format(IIf(Month(date) = 12, Year(date) + 1, Year(date)), "0000")
       ipp_IniVct.Text = Format(fecvcto, "DD/MM/YYYY")
       Call gs_SetFocus(cmb_MonAho)
   End If
End Sub

Private Sub cmb_DiaPag_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      Call gs_SetFocus(cmb_MonAho)
   End If
End Sub

Private Sub cmb_MonAho_Click()
    Call gs_SetFocus(ipp_IniVct)
End Sub

Private Sub cmb_MonAho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmb_MonAho_Click
    End If
End Sub

Private Sub ipp_IniVct_GotFocus()
    Call gs_SelecTodo(ipp_IniVct)
End Sub

Private Sub ipp_IniVct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoAho)
   End If
End Sub

Private Sub ipp_MtoAho_GotFocus()
    Call gs_SelecTodo(ipp_MtoAho)
End Sub

Private Sub ipp_MtoAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MesAho)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "- '")
   End If
End Sub

Private Sub ipp_MesAho_GotFocus()
    Call gs_SelecTodo(ipp_MesAho)
End Sub

Private Sub ipp_MesAho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_TotAho)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "- '")
   End If
End Sub

Private Sub ipp_MesAho_LostFocus()
    ipp_TotAho.Text = CDbl(ipp_MtoAho.Text) * CInt(ipp_MesAho.Text)
End Sub

Private Sub ipp_PlaAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DiaPag)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "- '")
   End If
End Sub

Private Sub ipp_MtoAho_LostFocus()
    ipp_TotAho.Text = CDbl(ipp_MtoAho.Text) * CInt(ipp_MesAho.Text)
End Sub

Private Sub ipp_TotAho_GotFocus()
    Call gs_SelecTodo(ipp_TotAho)
End Sub

Private Sub ipp_TotAho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ConHip)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "- '")
   End If
End Sub

Private Sub cmb_ConHip_Click()
    Call gs_SetFocus(cmb_EjeSeg)
End Sub

Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call cmb_ConHip_Click
   End If
End Sub

Private Sub cmb_EjeSeg_Click()
    Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_EjeSeg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmb_EjeSeg_Click
    End If
End Sub

Private Sub cmb_InsFin_Click()
    Call gs_SetFocus(cmb_MonAho)
End Sub

Private Sub cmb_InsFin_KeyPress(KeyAscii As Integer)
    Call cmb_InsFin_Click
End Sub
