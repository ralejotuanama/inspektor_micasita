VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   3000
   ClientTop       =   765
   ClientWidth     =   11655
   Icon            =   "AteCli_frm_102.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9885
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   17436
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
         Height          =   435
         Left            =   30
         TabIndex        =   78
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1950
            TabIndex        =   79
            Top             =   60
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1 - 07522154"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Caption         =   "Docum. de Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   80
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   46
         Top             =   30
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
            TabIndex        =   47
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
            Picture         =   "AteCli_frm_102.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7815
         Left            =   30
         TabIndex        =   48
         Top             =   1230
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   13785
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
         Begin VB.ComboBox cmb_DocAlt 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1065
         End
         Begin VB.ComboBox cmb_TDoAlt 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NDoAlt 
            Height          =   315
            Left            =   8160
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_ClasMC 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   7440
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ClaSbs 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   7110
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CarDom 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   6630
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8160
            MaxLength       =   250
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   5970
            Width           =   3315
         End
         Begin VB.TextBox txt_Telefo 
            Height          =   315
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   6300
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
            Left            =   9870
            TabIndex        =   19
            Top             =   3840
            Width           =   1485
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1920
            TabIndex        =   34
            Text            =   "cmb_DstDir"
            Top             =   5970
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8160
            TabIndex        =   33
            Text            =   "cmb_PrvDir"
            Top             =   5640
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1920
            TabIndex        =   32
            Text            =   "cmb_DptDir"
            Top             =   5640
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8160
            MaxLength       =   120
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   5310
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   5310
            Width           =   3315
         End
         Begin VB.TextBox txt_Interi 
            Height          =   315
            Left            =   9840
            MaxLength       =   15
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   4980
            Width           =   1640
         End
         Begin VB.TextBox txt_Numero 
            Height          =   315
            Left            =   8160
            MaxLength       =   15
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   4980
            Width           =   1640
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1920
            MaxLength       =   120
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   4980
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   4650
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8160
            MaxLength       =   120
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   3840
            Width           =   1665
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   1920
            MaxLength       =   9
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   3840
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   8160
            TabIndex        =   16
            Text            =   "cmb_Profes"
            Top             =   3510
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   3510
            Width           =   3315
         End
         Begin VB.ComboBox cmb_RegCyg 
            Height          =   315
            Left            =   8160
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3180
            Width           =   3315
         End
         Begin VB.ComboBox cmb_EstCiv 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   3180
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8160
            TabIndex        =   12
            Text            =   "cmb_DstNac"
            Top             =   2850
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   1920
            TabIndex        =   11
            Text            =   "cmb_PrvNac"
            Top             =   2850
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8160
            TabIndex        =   10
            Text            =   "cmb_DptNac"
            Top             =   2520
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Text            =   "cmb_Paises"
            Top             =   2520
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodSex 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1860
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1530
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   870
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1200
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1200
            Width           =   3315
         End
         Begin EditLib.fpLongInteger ipp_DepEc1 
            Height          =   315
            Left            =   8160
            TabIndex        =   21
            Top             =   4170
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
            Left            =   1920
            TabIndex        =   8
            Top             =   2190
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
            Left            =   8790
            TabIndex        =   22
            Top             =   4170
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
            Left            =   9450
            TabIndex        =   23
            Top             =   4170
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
            Left            =   10110
            TabIndex        =   24
            Top             =   4170
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
            Left            =   10740
            TabIndex        =   25
            Top             =   4170
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
            Left            =   0
            TabIndex        =   49
            Top             =   4530
            Width           =   11505
            _Version        =   65536
            _ExtentX        =   20294
            _ExtentY        =   159
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Left            =   1920
            TabIndex        =   20
            Top             =   4170
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   90
            Left            =   0
            TabIndex        =   81
            Top             =   6990
            Width           =   11505
            _Version        =   65536
            _ExtentX        =   20294
            _ExtentY        =   159
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin EditLib.fpLongInteger ipp_AnoDom 
            Height          =   315
            Left            =   8160
            TabIndex        =   38
            Top             =   6630
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   90
            Left            =   0
            TabIndex        =   89
            Top             =   750
            Width           =   11505
            _Version        =   65536
            _ExtentX        =   20294
            _ExtentY        =   159
            _StockProps     =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label35 
            Caption         =   "Personal FF.AA / FF.PP:"
            Height          =   315
            Left            =   60
            TabIndex        =   90
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   88
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label33 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   6180
            TabIndex        =   87
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label31 
            Caption         =   "Clasificación miCasita:"
            Height          =   315
            Left            =   60
            TabIndex        =   85
            Top             =   7440
            Width           =   1785
         End
         Begin VB.Label Label30 
            Caption         =   "Clasificación SBS:"
            Height          =   315
            Left            =   60
            TabIndex        =   84
            Top             =   7110
            Width           =   1545
         End
         Begin VB.Label Label32 
            Caption         =   "Nro. Años en Domicilio:"
            Height          =   285
            Left            =   6180
            TabIndex        =   83
            Top             =   6630
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Caract. Domicilio:"
            Height          =   315
            Left            =   60
            TabIndex        =   82
            Top             =   6630
            Width           =   1905
         End
         Begin VB.Label Label38 
            Caption         =   "Nro. Depend. Econom.:"
            Height          =   285
            Left            =   60
            TabIndex        =   77
            Top             =   4170
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6180
            TabIndex        =   76
            Top             =   5970
            Width           =   1485
         End
         Begin VB.Label Label27 
            Caption         =   "Teléfono:"
            Height          =   285
            Left            =   60
            TabIndex        =   75
            Top             =   6300
            Width           =   1485
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   60
            TabIndex        =   74
            Top             =   5970
            Width           =   1905
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6180
            TabIndex        =   73
            Top             =   5640
            Width           =   1905
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   60
            TabIndex        =   72
            Top             =   5640
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6180
            TabIndex        =   71
            Top             =   5310
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   60
            TabIndex        =   70
            Top             =   5310
            Width           =   1905
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6180
            TabIndex        =   69
            Top             =   4980
            Width           =   2055
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   60
            TabIndex        =   68
            Top             =   4980
            Width           =   1485
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   60
            TabIndex        =   67
            Top             =   4650
            Width           =   1905
         End
         Begin VB.Label Label18 
            Caption         =   "Edades Depend. Econom.:"
            Height          =   285
            Left            =   6180
            TabIndex        =   66
            Top             =   4170
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6180
            TabIndex        =   65
            Top             =   3840
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   60
            TabIndex        =   64
            Top             =   3840
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión o Actividad:"
            Height          =   315
            Left            =   6180
            TabIndex        =   63
            Top             =   3510
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   60
            TabIndex        =   62
            Top             =   3510
            Width           =   1905
         End
         Begin VB.Label Label13 
            Caption         =   "Régimen Conyugal:"
            Height          =   315
            Left            =   6180
            TabIndex        =   61
            Top             =   3180
            Width           =   1905
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   315
            Left            =   60
            TabIndex        =   60
            Top             =   3180
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6180
            TabIndex        =   59
            Top             =   2850
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   58
            Top             =   2850
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6180
            TabIndex        =   57
            Top             =   2520
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   2520
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   55
            Top             =   2190
            Width           =   1905
         End
         Begin VB.Label Label6 
            Caption         =   "Sexo:"
            Height          =   315
            Left            =   60
            TabIndex        =   54
            Top             =   1860
            Width           =   1905
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   60
            TabIndex        =   53
            Top             =   1530
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   60
            TabIndex        =   52
            Top             =   870
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   51
            Top             =   1200
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Apellido Casada:"
            Height          =   285
            Left            =   6180
            TabIndex        =   50
            Top             =   1200
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   735
         Left            =   30
         TabIndex        =   86
         Top             =   9090
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
         Begin VB.CommandButton cmd_SimCre 
            Height          =   675
            Left            =   30
            Picture         =   "AteCli_frm_102.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10860
            Picture         =   "AteCli_frm_102.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_102.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_DatCyg 
            Height          =   675
            Left            =   1410
            Picture         =   "AteCli_frm_102.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_ActEco 
            Height          =   675
            Left            =   720
            Picture         =   "AteCli_frm_102.frx":101B
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Paises()   As moddat_tpo_Genera
Dim l_arr_Profes()   As moddat_tpo_Genera
Dim l_str_Paises     As String
Dim l_str_Profes     As String
Dim l_str_DptNac     As String
Dim l_str_PrvNac     As String
Dim l_str_DstNac     As String
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_int_FlgCmb     As Integer

Private Sub cmb_CarDom_Click()
   Call gs_SetFocus(ipp_AnoDom)
End Sub

Private Sub cmb_CarDom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CarDom_Click
   End If
End Sub

Private Sub cmb_ClaSbs_Click()
   Call gs_SetFocus(cmb_ClasMC)
End Sub

Private Sub cmb_ClaSbs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClaSbs_Click
   End If
End Sub

Private Sub cmb_ClasMC_Click()
   Call gs_SetFocus(cmd_ActEco)
End Sub

Private Sub cmb_ClasMC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClasMC_Click
   End If
End Sub

Private Sub cmb_CodSex_Click()
   Call gs_SetFocus(ipp_FecNac)
End Sub

Private Sub cmb_CodSex_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodSex_Click
   End If
End Sub

Private Sub cmb_DocAlt_Click()
   If cmb_DocAlt.ListIndex = -1 Then
      cmb_TDoAlt.ListIndex = -1
      txt_NDoAlt.Text = ""
      
      cmb_TDoAlt.Enabled = False
      txt_NDoAlt.Enabled = False
   Else
      If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
         cmb_TDoAlt.Enabled = True
         txt_NDoAlt.Enabled = True
         
         Call gs_SetFocus(cmb_TDoAlt)
      Else
         cmb_TDoAlt.ListIndex = -1
         txt_NDoAlt.Text = ""
         
         cmb_TDoAlt.Enabled = False
         txt_NDoAlt.Enabled = False
      
         Call gs_SetFocus(txt_ApePat)
      End If
   End If
End Sub

Private Sub cmb_DocAlt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DocAlt_Click
   End If
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
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
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

Private Sub cmb_DptDir_LostFocus()
   Call SendMessage(cmb_DptDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_DstNac_LostFocus()
   Call SendMessage(cmb_DstNac.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_Paises_LostFocus()
   Call SendMessage(cmb_Paises.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
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

Private Sub cmb_PrvDir_LostFocus()
   Call SendMessage(cmb_PrvDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 1, 0&)

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

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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
   If cmb_EstCiv.ListIndex > -1 Then
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         cmb_RegCyg.Enabled = True
         cmd_DatCyg.Enabled = True
         
         Call gs_SetFocus(cmb_RegCyg)
      ElseIf cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         cmd_DatCyg.Enabled = True
         
         Call gs_SetFocus(cmb_NivEst)
      Else
         'Blanquear información registrada en memoria del Cónyuge
         moddat_g_int_FlgCyg = 1
   
         moddat_g_int_CygTDo = 0
         moddat_g_str_CygNDo = ""
         moddat_g_str_CygNom = ""
         
         Call moddat_gs_Inicia_DatCyg
         
         Call moddat_gs_Inicia_ActEco(2, 1)
         Call moddat_gs_Inicia_ActEco(2, 2)
      
         cmb_RegCyg.ListIndex = -1
         cmb_RegCyg.Enabled = False
         
         Call gs_SetFocus(cmb_NivEst)
      End If
   Else
      cmb_RegCyg.ListIndex = -1
      cmb_RegCyg.Enabled = False
   End If
End Sub

Private Sub cmb_EstCiv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EstCiv_Click
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
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./<>*+#,()" + Chr(34))
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

Private Sub cmb_Profes_LostFocus()
   Call SendMessage(cmb_Profes.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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

Private Sub cmb_PrvNac_LostFocus()
   Call SendMessage(cmb_PrvNac.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_RegCyg_Click()
   Call gs_SetFocus(cmb_NivEst)
End Sub

Private Sub cmb_RegCyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RegCyg_Click
   End If
End Sub

Private Sub cmb_TDoAlt_Click()
   If cmb_TDoAlt.ListIndex > -1 Then
      Select Case cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)
         Case 1:  txt_NDoAlt.MaxLength = 8
         Case Else:  txt_NDoAlt.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NDoAlt)
End Sub

Private Sub cmb_TDoAlt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TDoAlt_Click
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
   If Not ff_Valida_DatCli() Then
      Exit Sub
   End If
   
   moddat_g_str_NomCli = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre
   moddat_g_int_TipCli = 1
   
   frm_MntCli_03.Show 1
End Sub

Private Sub cmd_DatCyg_Click()
   If Not ff_Valida_DatCli() Then
      Exit Sub
   End If
   
   moddat_g_str_NomCli = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre
   
   frm_MntCli_09.Show 1
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_FlgAcc  As String
   Dim r_str_RelLab  As String
   Dim r_str_Parame  As String
   Dim r_int_Multip  As Integer
   Dim r_dbl_IngMin  As Double
   Dim r_dbl_IngDec  As Double

   If Not ff_Valida_DatCli() Then
      Exit Sub
   End If
   
   If moddat_g_arr_ActEco_Tit(1).ActEco_TipAct = 0 Then
      MsgBox "Debe ingresar la Actividad Económica del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_ActEco)
      Exit Sub
   End If

   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
      If moddat_g_arr_DatCyg(1).DatCli_TipDoc = 0 Then
         MsgBox "Debe ingresar la Información del Cónyuge o Conviviente.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmd_DatCyg)
         Exit Sub
      End If
      
      If moddat_g_arr_DatCyg(1).DatCli_ActEco = 1 Then
         If moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct = 0 Then
            MsgBox "Debe ingresar la Actividad Económica del Cónyuge o Conviviente.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmd_DatCyg)
            Exit Sub
         End If
      End If
   End If

   moddat_g_dbl_IngDec = 0
   If Len(Trim(moddat_g_str_CodPrd)) > 0 Then
      'Para Obtener Ingreso Total
      r_dbl_IngDec = 0
      
      'Actividad Principal del Cliente
      r_int_Multip = 0
      
      Select Case moddat_g_arr_ActEco_Tit(1).ActEco_Dep_FreHab
         Case 1:  r_int_Multip = 1
         Case 2:  r_int_Multip = 2
         Case 3:  r_int_Multip = 4
      End Select
   
      r_dbl_IngDec = r_dbl_IngDec + (moddat_g_arr_ActEco_Tit(1).ActEco_Dep_IngNet * r_int_Multip) + moddat_g_arr_ActEco_Tit(1).ActEco_Ind_IngNet + moddat_g_arr_ActEco_Tit(1).ActEco_Com_IngNet + moddat_g_arr_ActEco_Tit(1).ActEco_Acc_IngNet + moddat_g_arr_ActEco_Tit(1).ActEco_Ren_IngNet
   
      'Actividad Secundaria del Cliente
      r_int_Multip = 0
      
      Select Case moddat_g_arr_ActEco_Tit(2).ActEco_Dep_FreHab
         Case 1:  r_int_Multip = 1
         Case 2:  r_int_Multip = 2
         Case 3:  r_int_Multip = 4
      End Select
   
      r_dbl_IngDec = r_dbl_IngDec + (moddat_g_arr_ActEco_Tit(2).ActEco_Dep_IngNet * r_int_Multip) + moddat_g_arr_ActEco_Tit(2).ActEco_Ind_IngNet + moddat_g_arr_ActEco_Tit(2).ActEco_Com_IngNet + moddat_g_arr_ActEco_Tit(2).ActEco_Acc_IngNet + moddat_g_arr_ActEco_Tit(2).ActEco_Ren_IngNet
   
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         'Actividad Principal del Cónyuge
         r_int_Multip = 0
         
         Select Case moddat_g_arr_ActEco_Cyg(1).ActEco_Dep_FreHab
            Case 1:  r_int_Multip = 1
            Case 2:  r_int_Multip = 2
            Case 3:  r_int_Multip = 4
         End Select
         
         r_dbl_IngDec = r_dbl_IngDec + (moddat_g_arr_ActEco_Cyg(1).ActEco_Dep_IngNet * r_int_Multip) + moddat_g_arr_ActEco_Cyg(1).ActEco_Ind_IngNet + moddat_g_arr_ActEco_Cyg(1).ActEco_Com_IngNet + moddat_g_arr_ActEco_Cyg(1).ActEco_Acc_IngNet + moddat_g_arr_ActEco_Cyg(1).ActEco_Ren_IngNet
      
         'Actividad Secundaria del Cliente
         r_int_Multip = 0
         
         Select Case moddat_g_arr_ActEco_Cyg(2).ActEco_Dep_FreHab
            Case 1:  r_int_Multip = 1
            Case 2:  r_int_Multip = 2
            Case 3:  r_int_Multip = 4
         End Select
      
         r_dbl_IngDec = r_dbl_IngDec + (moddat_g_arr_ActEco_Cyg(2).ActEco_Dep_IngNet * r_int_Multip) + moddat_g_arr_ActEco_Cyg(2).ActEco_Ind_IngNet + moddat_g_arr_ActEco_Cyg(2).ActEco_Com_IngNet + moddat_g_arr_ActEco_Cyg(2).ActEco_Acc_IngNet + moddat_g_arr_ActEco_Cyg(2).ActEco_Ren_IngNet
      End If
   
      moddat_g_dbl_IngDec = r_dbl_IngDec
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información del Cliente
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CLI_DATGEN ("
      
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex)) & ", "
      
      If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & CStr(cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NDoAlt.Text & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
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
      g_str_Parame = g_str_Parame & "'" & l_arr_Profes(cmb_Profes.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_CodSex.ItemData(cmb_CodSex.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecNac.Text), "yyyymmdd") & ", "
      
      If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
         g_str_Parame = g_str_Parame & "'" & Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00") & Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00") & Format(cmb_DstNac.ItemData(cmb_DstNac.ListIndex), "00") & "', "
      Else
         g_str_Parame = g_str_Parame & "'000000', "
      End If
      g_str_Parame = g_str_Parame & "'" & l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo & "', "
      
      g_str_Parame = g_str_Parame & CStr(ipp_NumDep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DepEc1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DepEc2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DepEc3.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DepEc4.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_DepEc5.Value) & ", "
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
      ElseIf chk_DirEle.Value = 0 Then
         g_str_Parame = g_str_Parame & "2, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(cmb_CarDom.ItemData(cmb_CarDom.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AnoDom.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & Left(cmb_ClaSbs.Text, 1) & "', "
      g_str_Parame = g_str_Parame & "'" & Left(cmb_ClasMC.Text, 1) & "', "
      g_str_Parame = g_str_Parame & "'1', "
      
      'Buscar en BD de Vinculados
      r_str_FlgAcc = "0"
      r_str_RelLab = "0"
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
      r_str_Parame = r_str_Parame & "PERVIN_TDOTIT = " & CStr(moddat_g_int_TipDoc) & " AND "
      r_str_Parame = r_str_Parame & "PERVIN_NDOTIT = '" & moddat_g_str_NumDoc & "' AND "
      r_str_Parame = r_str_Parame & "PERVIN_TDOVIN = 0 "

      If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_str_FlgAcc = Trim(g_rst_Genera!PERVIN_FLGACC)
         r_str_RelLab = Trim(g_rst_Genera!PERVIN_RELLAB)
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      If Len(Trim(r_str_FlgAcc)) = 0 And Len(Trim(r_str_RelLab)) = 0 Then
         r_str_Parame = ""
         r_str_Parame = r_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
         r_str_Parame = r_str_Parame & "PERVIN_TDOVIN = " & CStr(moddat_g_int_TipDoc) & " AND "
         r_str_Parame = r_str_Parame & "PERVIN_NDOVIN = '" & moddat_g_str_NumDoc & "' "
   
         If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
      
            r_str_FlgAcc = Trim(g_rst_Genera!PERVIN_FLGACC)
            r_str_RelLab = Trim(g_rst_Genera!PERVIN_RELLAB)
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      End If
      
      g_str_Parame = g_str_Parame & "'" & r_str_FlgAcc & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_RelLab & "', "
      
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_TipAct) & ", "
      
      Select Case moddat_g_arr_ActEco_Tit(1).ActEco_TipAct
         Case 11
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Dep_CodCiu) & ", "
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Dep_TipDoc) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(1).ActEco_Dep_NumDoc & "', "
            
         Case 21
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Ind_CodCiu) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            
         Case 31
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Com_CodCiu) & ", "
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Com_TipDoc) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(1).ActEco_Com_NumDoc & "', "
            
         Case 41
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Acc_CodCiu) & ", "
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Acc_TipDoc) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(1).ActEco_Acc_NumDoc & "', "
            
         Case 51
            g_str_Parame = g_str_Parame & "9999, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
      
         Case 61
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Otr_CodCiu) & ", "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
      End Select
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   'Grabando Actividad Económica
   g_str_Parame = "USP_CLI_ACTECO_BORRAR ("
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   Call fs_Grabar_ActEco_Tit(1)
   
   If moddat_g_arr_ActEco_Tit(2).ActEco_TipAct > 0 Then
      Call fs_Grabar_ActEco_Tit(2)
   End If
   
   'Si cliente presenta Cónyuge
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
      moddat_g_int_CygTDo = moddat_g_arr_DatCyg(1).DatCli_TipDoc
      moddat_g_str_CygNDo = moddat_g_arr_DatCyg(1).DatCli_NumDoc
   
      'Verificando si Cónyuge ya existe en Base Datos
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         moddat_g_int_FlgGrb_1 = 1
      Else
         moddat_g_int_FlgGrb_1 = 2
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Grabando Información del Cónyuge
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_CLI_DATGEN ("
         
         g_str_Parame = g_str_Parame & "2, "
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_DocAlt) & ", "
         
         If moddat_g_arr_DatCyg(1).DatCli_DocAlt = 1 Then
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_TDoAlt) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_NDoAlt & "', "
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_ApePat & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_ApeMat & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_ApeCas & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_Nombre & "', "
         
         g_str_Parame = g_str_Parame & CStr(cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex)) & ", "
         
         If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
            g_str_Parame = g_str_Parame & CStr(cmb_RegCyg.ItemData(cmb_RegCyg.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "0, "
         End If
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_NivEst) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_Profes & "', "
         
         If cmb_CodSex.ItemData(cmb_CodSex.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & "2, "
         Else
            g_str_Parame = g_str_Parame & "1, "
         End If
         
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_DatCyg(1).DatCli_FecNac), "yyyymmdd") & ", "
         
         If CInt(moddat_g_arr_DatCyg(1).DatCli_Paises) = 4028 Then
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_UbiGeo & "', "
         Else
            g_str_Parame = g_str_Parame & "'000000', "
         End If
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_Paises & "', "
         
         g_str_Parame = g_str_Parame & CStr(ipp_NumDep.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_DepEc1.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_DepEc2.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_DepEc3.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_DepEc4.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_DepEc5.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Numero.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Interi.Text & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_Celula & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Telefo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_DirEle & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_ChkEle) & ", "
         
         g_str_Parame = g_str_Parame & CStr(cmb_CarDom.ItemData(cmb_CarDom.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_AnoDom.Value) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_ClaSbs & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_ClasMC & "', "
         g_str_Parame = g_str_Parame & "'1', "
         
         'Buscar en BD de Vinculados
         r_str_FlgAcc = ""
         r_str_RelLab = ""
         
         r_str_Parame = ""
         r_str_Parame = r_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
         r_str_Parame = r_str_Parame & "PERVIN_TDOTIT = " & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & " AND "
         r_str_Parame = r_str_Parame & "PERVIN_NDOTIT = '" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "' AND "
         r_str_Parame = r_str_Parame & "PERVIN_TDOVIN = 0 "
   
         If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
      
            r_str_FlgAcc = Trim(g_rst_Genera!PERVIN_FLGACC)
            r_str_RelLab = Trim(g_rst_Genera!PERVIN_RELLAB)
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         If Len(Trim(r_str_FlgAcc)) = 0 And Len(Trim(r_str_RelLab)) = 0 Then
            r_str_Parame = ""
            r_str_Parame = r_str_Parame & "SELECT * FROM CRE_PERVIN WHERE "
            r_str_Parame = r_str_Parame & "PERVIN_TDOVIN = " & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & " AND "
            r_str_Parame = r_str_Parame & "PERVIN_NDOVIN = '" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "' "
      
            If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
               Exit Sub
            End If
         
            If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
               g_rst_Genera.MoveFirst
         
               r_str_FlgAcc = Trim(g_rst_Genera!PERVIN_FLGACC)
               r_str_RelLab = Trim(g_rst_Genera!PERVIN_RELLAB)
            End If
            
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
         End If
         
         g_str_Parame = g_str_Parame & "'" & r_str_FlgAcc & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_RelLab & "', "
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_ActEco) & ", "
         
         If moddat_g_arr_DatCyg(1).DatCli_ActEco = 1 Then
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct) & ", "
         
            Select Case moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct
               Case 11
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Dep_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Dep_TipDoc) & ", "
                  g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(1).ActEco_Dep_NumDoc & "', "
                  
               Case 21
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Ind_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
                  
               Case 31
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Com_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Com_TipDoc) & ", "
                  g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(1).ActEco_Com_NumDoc & "', "
                  
               Case 41
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Acc_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(1).ActEco_Acc_TipDoc) & ", "
                  g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(1).ActEco_Acc_NumDoc & "', "
                  
               Case 51
                  g_str_Parame = g_str_Parame & "9999, "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
            
               Case 61
                  g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(1).ActEco_Otr_CodCiu) & ", "
                  g_str_Parame = g_str_Parame & "0, "
                  g_str_Parame = g_str_Parame & "'', "
            End Select
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
            g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb_1) & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
      
      'Grabando Actividad Económica
      g_str_Parame = "USP_CLI_ACTECO_BORRAR ("
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
          Exit Sub
      End If
      
      If moddat_g_arr_DatCyg(1).DatCli_ActEco = 1 Then
         Call fs_Grabar_ActEco_Cyg(1)
      
         If moddat_g_arr_ActEco_Cyg(2).ActEco_TipAct > 0 Then
            Call fs_Grabar_ActEco_Cyg(2)
         End If
      End If
   Else
      moddat_g_int_CygTDo = 0
      moddat_g_str_CygNDo = ""
   End If
   
   moddat_g_int_FlgAct = 2
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
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_DocIde.Caption = moddat_gf_Consulta_ParDes("230", moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc
   
   Call fs_Inicio
   Call fs_Limpia
   
   'Inicializando Arreglos Actividades Económicas del Titular
   Call moddat_gs_Inicia_ActEco(1, 1)
   Call moddat_gs_Inicia_ActEco(1, 2)

   'Inicializando Arreglo de Datos del Cónyuge
   Call moddat_gs_Inicia_DatCyg
   moddat_g_int_FlgCyg = 1
   
   'Inicializando Arreglos Actividades Económicas del Cónyuge
   Call moddat_gs_Inicia_ActEco(2, 1)
   Call moddat_gs_Inicia_ActEco(2, 2)
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         Call gs_BuscarCombo_Item(cmb_DocAlt, g_rst_Princi!DatGen_FLGDOA)
         
         If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
            Call gs_BuscarCombo_Item(cmb_TDoAlt, g_rst_Princi!DatGen_TIPDOA)
            txt_NDoAlt.Text = Trim(g_rst_Princi!DatGen_NUMDOA)
            
            cmb_TDoAlt.Enabled = True
            txt_NDoAlt.Enabled = True
         End If
         
         txt_ApePat.Text = Trim(g_rst_Princi!DatGen_ApePat & "")
         txt_ApeMat.Text = Trim(g_rst_Princi!DatGen_ApeMat & "")
         txt_ApeCas.Text = Trim(g_rst_Princi!DatGen_ApeCas & "")
         txt_Nombre.Text = Trim(g_rst_Princi!DatGen_Nombre & "")
         
         Call gs_BuscarCombo_Item(cmb_CodSex, g_rst_Princi!DatGen_CodSex)
         ipp_FecNac.Text = Right(CStr(g_rst_Princi!DATGEN_NACFEC), 2) & "/" & Mid(CStr(g_rst_Princi!DATGEN_NACFEC), 5, 2) & "/" & Left(CStr(g_rst_Princi!DATGEN_NACFEC), 4)
         
         cmb_Paises.ListIndex = gf_Busca_Arregl(l_arr_Paises, g_rst_Princi!DATGEN_NACPAI) - 1
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
            Call gs_BuscarCombo_Item(cmb_DptNac, CInt(Left(g_rst_Princi!DATGEN_NACLUG, 2)))
            Call moddat_gs_Carga_Provin(cmb_PrvNac, Left(g_rst_Princi!DATGEN_NACLUG, 2))
            Call gs_BuscarCombo_Item(cmb_PrvNac, CInt(Mid(g_rst_Princi!DATGEN_NACLUG, 3, 2)))
            Call moddat_gs_Carga_Distri(cmb_DstNac, Left(g_rst_Princi!DATGEN_NACLUG, 2), Mid(g_rst_Princi!DATGEN_NACLUG, 3, 2))
            Call gs_BuscarCombo_Item(cmb_DstNac, CInt(Right(g_rst_Princi!DATGEN_NACLUG, 2)))
            
            cmb_DptNac.Enabled = True
            cmb_PrvNac.Enabled = True
            cmb_DstNac.Enabled = True
         End If
         
         Call gs_BuscarCombo_Item(cmb_EstCiv, g_rst_Princi!DATGEN_ESTCIV)
         
         If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
            Call gs_BuscarCombo_Item(cmb_RegCyg, g_rst_Princi!DatGen_RegCyg)
            
            cmb_RegCyg.Enabled = True
         End If
         
         Call gs_BuscarCombo_Item(cmb_NivEst, g_rst_Princi!DatGen_NivEst)
         cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, g_rst_Princi!DatGen_Profes) - 1
         
         txt_Celula.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
         txt_DirEle.Text = Trim(g_rst_Princi!DatGen_DirEle & "")
         
         If g_rst_Princi!DATGEN_AUTENV = 1 Then
            chk_DirEle.Value = 1
            chk_DirEle.Enabled = True
         End If
         
         ipp_NumDep.Value = g_rst_Princi!DatGen_DepEco
         
         ipp_DepEc1.Value = g_rst_Princi!DatGen_EDAD01
         ipp_DepEc2.Value = g_rst_Princi!DatGen_EDAD02
         ipp_DepEc3.Value = g_rst_Princi!DatGen_EDAD03
         ipp_DepEc4.Value = g_rst_Princi!DatGen_EDAD04
         ipp_DepEc5.Value = g_rst_Princi!DatGen_EDAD05
         
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!DatGen_TipVia)
         txt_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_Numero.Text = Trim(g_rst_Princi!DatGen_Numero & "")
         txt_Interi.Text = Trim(g_rst_Princi!DatGen_IntDpt & "")
         
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!DatGen_TipZon)
         txt_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
      
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!DatGen_Ubigeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!DatGen_Ubigeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!DatGen_Ubigeo, 2), Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!DatGen_Ubigeo, 2)))
         
         txt_Refere.Text = Trim(g_rst_Princi!DatGen_Refere & "")
         txt_Telefo.Text = Trim(g_rst_Princi!DatGen_Telefo & "")
         
         Call gs_BuscarCombo_Item(cmb_CarDom, g_rst_Princi!DATGEN_CARDOM)
         ipp_AnoDom.Value = g_rst_Princi!DatGen_ANODOM
         
         Call gs_BuscarCombo_Text(cmb_ClaSbs, g_rst_Princi!DATGEN_CLASBS, 1)
         Call gs_BuscarCombo_Text(cmb_ClasMC, g_rst_Princi!DATGEN_CLASMC, 1)
         
         moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
         moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Cargar Actividades Económicas del Titular
      Call fs_Cargar_ActEco_Tit(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
      Call fs_Cargar_ActEco_Tit(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2)
      
      'Buscar Información del Cónyuge
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
      
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            moddat_g_int_FlgGrb_1 = 2
            moddat_g_int_FlgCyg = 2
            
            moddat_g_arr_DatCyg(1).DatCli_TipDoc = moddat_g_int_CygTDo
            moddat_g_arr_DatCyg(1).DatCli_NumDoc = moddat_g_str_CygNDo
            
            moddat_g_arr_DatCyg(1).DatCli_DocAlt = g_rst_Princi!DatGen_FLGDOA
         
            If moddat_g_arr_DatCyg(1).DatCli_DocAlt = 1 Then
               moddat_g_arr_DatCyg(1).DatCli_TDoAlt = g_rst_Princi!DatGen_TIPDOA
               moddat_g_arr_DatCyg(1).DatCli_NDoAlt = Trim(g_rst_Princi!DatGen_NUMDOA & "")
            End If
         
            moddat_g_arr_DatCyg(1).DatCli_ApePat = Trim(g_rst_Princi!DatGen_ApePat & "")
            moddat_g_arr_DatCyg(1).DatCli_ApeMat = Trim(g_rst_Princi!DatGen_ApeMat & "")
            moddat_g_arr_DatCyg(1).DatCli_ApeCas = Trim(g_rst_Princi!DatGen_ApeCas & "")
            moddat_g_arr_DatCyg(1).DatCli_Nombre = Trim(g_rst_Princi!DatGen_Nombre & "")
            
            moddat_g_str_CygNom = Trim(g_rst_Princi!DatGen_ApePat & "") & " " & Trim(g_rst_Princi!DatGen_ApeMat & "") & " " & Trim(g_rst_Princi!DatGen_Nombre & "")
         
            moddat_g_arr_DatCyg(1).DatCli_FecNac = Right(CStr(g_rst_Princi!DATGEN_NACFEC), 2) & "/" & Mid(CStr(g_rst_Princi!DATGEN_NACFEC), 5, 2) & "/" & Left(CStr(g_rst_Princi!DATGEN_NACFEC), 4)
            moddat_g_arr_DatCyg(1).DatCli_Paises = g_rst_Princi!DATGEN_NACPAI
            
            moddat_g_arr_DatCyg(1).DatCli_UbiGeo = g_rst_Princi!DATGEN_NACLUG
            moddat_g_arr_DatCyg(1).DatCli_NivEst = g_rst_Princi!DatGen_NivEst
            moddat_g_arr_DatCyg(1).DatCli_Profes = Trim(g_rst_Princi!DatGen_Profes & "")
            
            moddat_g_arr_DatCyg(1).DatCli_Celula = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
            moddat_g_arr_DatCyg(1).DatCli_DirEle = Trim(g_rst_Princi!DatGen_DirEle & "")
            moddat_g_arr_DatCyg(1).DatCli_ChkEle = g_rst_Princi!DATGEN_AUTENV
            moddat_g_arr_DatCyg(1).DatCli_ClaSbs = Trim(g_rst_Princi!DATGEN_CLASBS & "")
            moddat_g_arr_DatCyg(1).DatCli_ClasMC = Trim(g_rst_Princi!DATGEN_CLASMC & "")
            moddat_g_arr_DatCyg(1).DatCli_ActEco = g_rst_Princi!DATGEN_ACTECO
         End If
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      
         'Cargar Actividades Económicas del Cónyuge
         Call fs_Cargar_ActEco_Cyg(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
         Call fs_Cargar_ActEco_Cyg(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
      End If
   End If
   
   Call gs_CentraForm(Me)
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_DocAlt, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoAlt, 1, "231")
   
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
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_CarDom, 1, "233")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_ClaSbs, 1, "058")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ClasMC, 1, "058")
End Sub
   
Private Sub fs_Limpia()
   cmb_DocAlt.ListIndex = -1
   cmb_TDoAlt.ListIndex = -1
   txt_NDoAlt.Text = ""
   
   cmb_TDoAlt.Enabled = False
   txt_NDoAlt.Enabled = False
   
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   cmb_CodSex.ListIndex = -1
   ipp_FecNac.Text = Format(Date, "dd/mm/yyyy")
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
   
   cmb_CarDom.ListIndex = -1
   ipp_AnoDom.Value = 0
   
   cmb_ClaSbs.ListIndex = -1
   cmb_ClasMC.ListIndex = -1

   'Limpiando Arreglos de Actividad Economica del Titular
   Call moddat_gs_Inicia_ActEco(1, 1)
   Call moddat_gs_Inicia_ActEco(1, 2)

   'Limpiando Arreglos de Actividad Economica del Cónyuge
   Call moddat_gs_Inicia_ActEco(2, 1)
   Call moddat_gs_Inicia_ActEco(2, 2)
End Sub

Private Sub ipp_AnoDom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ClaSbs)
   End If
End Sub

Private Sub ipp_DepEc1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc2)
   End If
End Sub

Private Sub ipp_DepEc2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc3)
   End If
End Sub

Private Sub ipp_DepEc3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc4)
   End If
End Sub

Private Sub ipp_DepEc4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc5)
   End If
End Sub

Private Sub ipp_DepEc5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   End If
End Sub

Private Sub ipp_FecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Paises)
   End If
End Sub

Private Sub ipp_NumDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc1)
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

Private Sub txt_NDoAlt_GotFocus()
   Call gs_SelecTodo(txt_NDoAlt)
End Sub

Private Sub txt_NDoAlt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApePat)
   Else
      If cmb_TDoAlt.ListIndex > -1 Then
         Select Case cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)
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

Private Sub txt_Telefo_GotFocus()
   Call gs_SelecTodo(txt_Telefo)
End Sub

Private Sub txt_Telefo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CarDom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Function ff_Valida_DatCli() As Integer
   Dim r_int_EdaMin     As Integer
   Dim r_int_EdaMax     As Integer
   Dim r_int_EdaAct     As Integer
   
   ff_Valida_DatCli = False
   
   If cmb_DocAlt.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente es miembro de las FF.AA o FF.PP.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DocAlt)
      Exit Function
   End If
   
   If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
      If cmb_TDoAlt.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoAlt)
         Exit Function
      End If
      
      If Len(Trim(txt_NDoAlt.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NDoAlt)
         Exit Function
      End If
   End If
   
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
   End If
   
   'Si es Masculino
   If cmb_CodSex.ItemData(cmb_CodSex.ListIndex) = 1 Then
      If Len(Trim(txt_ApeCas.Text)) > 0 Then
         MsgBox "El cliente no puede presentar Apellido de Casada.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_ApeCas)
         Exit Function
      End If
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
   
   'Rango de Edades del Cliente
   If Len(Trim(moddat_g_str_CodPrd)) > 0 Then
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "011") Then
         r_int_EdaMin = moddat_g_arr_Genera(1).Genera_ValMin
         r_int_EdaMax = moddat_g_arr_Genera(1).Genera_ValMax
      End If
      
      r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(ipp_FecNac.Text), Date), 2))
      
      If Not (r_int_EdaAct >= r_int_EdaMin And r_int_EdaAct <= r_int_EdaMax) Then
         MsgBox "El Cliente no cumple con los requisitos de Edad requeridos. Tiene " & CStr(r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecNac)
         Exit Function
      End If
   End If
   
   If cmb_Paises.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Paises)
      Exit Function
   End If
   
   If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
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
   
   If cmb_CarDom.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Característica del Domicilio actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CarDom)
      Exit Function
   End If
   
   If cmb_ClaSbs.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación de la SBS.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClaSbs)
      Exit Function
   End If
   
   If cmb_ClasMC.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clasificación en miCasita.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClasMC)
      Exit Function
   End If
   
   moddat_g_str_NomCli = Trim(txt_ApePat.Text) & " " & Trim(txt_ApeMat.Text) & " " & Trim(txt_Nombre.Text)
   
   ff_Valida_DatCli = True
End Function

Private Sub fs_Grabar_ActEco_Tit(ByVal p_Indice As Integer)
   'Grabando Información de Actividad Económica
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CLI_ACTECO ("
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_OrdAct) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_TipAct) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipOfi) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_SitTra) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FreHab) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomAre & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumAnx & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TelDir & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Celula & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_DirEle & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IngNet) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCar & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_GirCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_VtaMen) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCar & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RegTri) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_PorPar) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_AlqMen) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomArr & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TelArr & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_PorPar) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr1 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele11 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele21 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe1) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_SegPro) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr2 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele12 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele22 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe2) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Activi & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Observ & "', "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   'Verificando si tiene empresas por Crear en Maestro de Empresas (Dependiente)
   If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp = 9 And moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc > 0 Then
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_EMP_DATGEN ("
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom & "', "
         g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu & ", "
         g_str_Parame = g_str_Parame & "'', "
         
         If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipOfi = 1 Then
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt & "', "
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax & "', "
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "000000, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "6, "
         g_str_Parame = g_str_Parame & "'0', "
         g_str_Parame = g_str_Parame & "'', "
         
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
            If MsgBox("No se pudo completar la grabación de los datos (EMP_DATGEN - Titular - <Dependiente>). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   End If
   
   'Verificando si tiene empresas por Crear en Maestro de Empresas (Dependiente - Trabajo Anterior)
   If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant)) > 0 Then
      If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant = 9 Then
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            Screen.MousePointer = 11
            
            g_str_Parame = "USP_EMP_DATGEN ("
            
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant & "', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "000000, "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "6, "
            g_str_Parame = g_str_Parame & "'0', "
            g_str_Parame = g_str_Parame & "'', "
            
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
               If MsgBox("No se pudo completar la grabación de los datos (EMP_DATGEN - Titular - <Dependiente - Trabajo Anterior>). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
            
            Screen.MousePointer = 0
         Loop
      End If
   End If
   
   'Verificando si tiene empresas por Crear en Maestro de Empresas (Accionista)
   If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FlgEmp = 9 And moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc > 0 Then
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_EMP_DATGEN ("
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom & "', "
         g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu & ", "
         g_str_Parame = g_str_Parame & "'', "
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "6, "
         g_str_Parame = g_str_Parame & "'0', "
         g_str_Parame = g_str_Parame & "'', "
         
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
            If MsgBox("No se pudo completar la grabación de los datos (EMP_DATGEN - Titular - <Accionista>). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   End If
End Sub

Private Sub fs_Grabar_ActEco_Cyg(ByVal p_Indice As Integer)
   'Grabando Información de Actividad Económica
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CLI_ACTECO ("
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_DatCyg(1).DatCli_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_DatCyg(1).DatCli_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_OrdAct) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_TipAct) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipOfi) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_SitTra) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FreHab) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomAre & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumAnx & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TelDir & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Celula & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_DirEle & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IngNet) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCar & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_GirCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_VtaMen) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCar & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RegTri) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_PorPar) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_AlqMen) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomArr & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TelArr & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_PorPar) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr1 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele11 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele21 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe1) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_SegPro) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr2 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele12 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele22 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe2) & ", "
      
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Activi & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Observ & "', "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop

   'Verificando si tiene empresas por Crear en Maestro de Empresas (Dependiente)
   If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp = 9 And moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc > 0 Then
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_EMP_DATGEN ("
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom & "', "
         g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu & ", "
         g_str_Parame = g_str_Parame & "'', "
         
         If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipOfi = 1 Then
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt & "', "
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax & "', "
         Else
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "000000, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "6, "
         g_str_Parame = g_str_Parame & "'0', "
         g_str_Parame = g_str_Parame & "'', "
         
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
            If MsgBox("No se pudo completar la grabación de los datos (EMP_DATGEN - Titular - <Dependiente>). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   End If
   
   'Verificando si tiene empresas por Crear en Maestro de Empresas (Dependiente - Trabajo Anterior)
   If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant)) > 0 Then
      If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant = 9 Then
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            Screen.MousePointer = 11
            
            g_str_Parame = "USP_EMP_DATGEN ("
            
            g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant) & ", "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant & "', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "000000, "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant & "', "
            g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant & "', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "6, "
            g_str_Parame = g_str_Parame & "'0', "
            g_str_Parame = g_str_Parame & "'', "
            
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
               If MsgBox("No se pudo completar la grabación de los datos (EMP_DATGEN - Titular - <Dependiente - Trabajo Anterior>). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
            
            Screen.MousePointer = 0
         Loop
      End If
   End If
   
   'Verificando si tiene empresas por Crear en Maestro de Empresas (Accionista)
   If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FlgEmp = 9 And moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc > 0 Then
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         
         g_str_Parame = "USP_EMP_DATGEN ("
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom & "', "
         g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu & ", "
         g_str_Parame = g_str_Parame & "'', "
         
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "6, "
         g_str_Parame = g_str_Parame & "'0', "
         g_str_Parame = g_str_Parame & "'', "
         
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
            If MsgBox("No se pudo completar la grabación de los datos (EMP_DATGEN - Titular - <Accionista>). ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
   End If
End Sub

Private Sub fs_Cargar_ActEco_Tit(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_Indice) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_OrdAct = p_Indice
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_TipAct = g_rst_Princi!ActEco_CodAct
   
      'Dependiente
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_SitTra = g_rst_Princi!ActEco_Dep_SitTra
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc = g_rst_Princi!ActEco_Dep_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipOfi = g_rst_Princi!ActEco_Dep_TipOfi
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu = g_rst_Princi!ActEco_Dep_CodCiu
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      Else
         g_rst_Genera.MoveFirst
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu = g_rst_Genera!DATGEN_CODCIU
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Genera!DATGEN_TELERH & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      
         If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = 1 Then
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = g_rst_Genera!DatGen_TipVia
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = g_rst_Genera!DatGen_TipZon
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         Else
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IngNet = g_rst_Princi!ActEco_Dep_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FreHab = g_rst_Princi!ActEco_Dep_FreHab
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCar = Trim(g_rst_Princi!ActEco_Dep_CodCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCar = Trim(g_rst_Princi!ActEco_Dep_NomCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomAre = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumAnx = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TelDir = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Celula = Trim(g_rst_Princi!ActEco_Dep_Celula & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_DirEle = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt = g_rst_Princi!ActEco_Dep_TraAnt
      
      If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt = 1 Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant = g_rst_Princi!ActEco_Dep_TipDoc_Ant
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant = g_rst_Princi!ActEco_Dep_NumDoc_Ant
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant = "9"
         
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")
         Else
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
      End If
      
      'Independiente
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc = g_rst_Princi!ActEco_Ind_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc = Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipVia = g_rst_Princi!ActEco_Ind_TipVia
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomVia = Trim(g_rst_Princi!ActEco_Ind_NomVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumVia = Trim(g_rst_Princi!ActEco_Ind_NumVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IntDpt = Trim(g_rst_Princi!ActEco_Ind_IntDpt & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipZon = g_rst_Princi!ActEco_Ind_TipZon
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomZon = Trim(g_rst_Princi!ActEco_Ind_NomZon & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_UbiGeo = Trim(g_rst_Princi!ActEco_Ind_UbiGeo & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Refere = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1 = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2 = Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumFax = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCiu = g_rst_Princi!ActEco_Ind_CodCiu
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IngNet = g_rst_Princi!ActEco_Ind_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc = g_rst_Princi!ActEco_Ind_ConLoc
      
      If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc = 1 Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp = g_rst_Princi!ActEco_Ind_TipDoc_Emp
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp = g_rst_Princi!ActEco_Ind_NumDoc_Emp
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FlgEmp = "9"
         
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")
         Else
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCar = Trim(g_rst_Princi!ActEco_Ind_CodCar & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCar = Trim(g_rst_Princi!ActEco_Ind_NomCar & "")
      End If
         
      'Comerciante
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipDoc = g_rst_Princi!ActEco_Com_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumDoc = Trim(g_rst_Princi!ActEco_Com_NumDoc & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RazSoc = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCom = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipVia = g_rst_Princi!ActEco_Com_TipVia
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomVia = Trim(g_rst_Princi!ActEco_Com_NomVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumVia = Trim(g_rst_Princi!ActEco_Com_NumVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IntDpt = Trim(g_rst_Princi!ActEco_Com_IntDpt & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipZon = g_rst_Princi!ActEco_Com_TipZon
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomZon = Trim(g_rst_Princi!ActEco_Com_NomZon & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_UbiGeo = Trim(g_rst_Princi!ActEco_Com_UbiGeo & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Refere = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef1 = Trim(g_rst_Princi!ActEco_Com_Telef1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef2 = Trim(g_rst_Princi!ActEco_Com_Telef2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumFax = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCiu = g_rst_Princi!ActEco_Com_CodCiu
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_GirCom = Trim(g_rst_Princi!ActEco_Com_GirCom & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IngNet = g_rst_Princi!ActEco_Com_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_VtaMen = g_rst_Princi!ActEco_Com_VtaMen
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCar = Trim(g_rst_Princi!ActEco_Com_CodCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCar = Trim(g_rst_Princi!ActEco_Com_NomCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RegTri = g_rst_Princi!ActEco_Com_RegTri
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_PorPar = g_rst_Princi!ActEco_Com_PorPar
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipLoc = g_rst_Princi!ActEco_Com_TipLoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_AlqMen = g_rst_Princi!ActEco_Com_AlqMen
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomArr = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TelArr = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
      
      'Accionista
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc = g_rst_Princi!ActEco_Acc_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc = Trim(g_rst_Princi!ActEco_Com_NumDoc & "")
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia = g_rst_Princi!ActEco_Acc_TipVia
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Princi!ActEco_Acc_NomVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Princi!ActEco_Acc_NumVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Princi!ActEco_Acc_IntDpt & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon = g_rst_Princi!ActEco_Acc_TipZon
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Princi!ActEco_Acc_NomZon & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Princi!ActEco_Acc_UbiGeo & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Princi!ActEco_Acc_Refere & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu = g_rst_Princi!ActEco_Acc_CodCiu
      Else
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia = g_rst_Genera!DatGen_TipVia
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon = g_rst_Genera!DatGen_TipZon
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu = g_rst_Genera!DATGEN_CODCIU
      End If
         
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IngNet = g_rst_Princi!ActEco_Acc_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_PorPar = g_rst_Princi!ActEco_Acc_PorPar
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
      
      'Rentista
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IngNet = g_rst_Princi!ActEco_Ren_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc1 = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr1 = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele11 = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele21 = Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe1 = g_rst_Princi!ActEco_Ren_AlqMe1
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_SegPro = g_rst_Princi!ActEco_Ren_SegPro
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc2 = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr2 = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele12 = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele22 = Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe2 = g_rst_Princi!ActEco_Ren_AlqMe2
   
      'Otros
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_IngNet = g_rst_Princi!ActEco_Otr_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Activi = Trim(g_rst_Princi!ActEco_Otr_Activi & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_CodCiu = g_rst_Princi!ActEco_Otr_CodCiu
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Observ = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

Private Sub fs_Cargar_ActEco_Cyg(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_Indice) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_OrdAct = p_Indice
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_TipAct = g_rst_Princi!ActEco_CodAct
   
      'Dependiente
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_SitTra = g_rst_Princi!ActEco_Dep_SitTra
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc = g_rst_Princi!ActEco_Dep_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipOfi = g_rst_Princi!ActEco_Dep_TipOfi
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu = g_rst_Princi!ActEco_Dep_CodCiu
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      Else
         g_rst_Genera.MoveFirst
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu = g_rst_Genera!DATGEN_CODCIU
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Genera!DATGEN_TELERH & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      
         If moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = 1 Then
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = g_rst_Genera!DatGen_TipVia
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = g_rst_Genera!DatGen_TipZon
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         Else
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IngNet = g_rst_Princi!ActEco_Dep_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FreHab = g_rst_Princi!ActEco_Dep_FreHab
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCar = Trim(g_rst_Princi!ActEco_Dep_CodCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCar = Trim(g_rst_Princi!ActEco_Dep_NomCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomAre = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumAnx = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TelDir = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Celula = Trim(g_rst_Princi!ActEco_Dep_Celula & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_DirEle = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt = g_rst_Princi!ActEco_Dep_TraAnt
      
      If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt = 1 Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant = g_rst_Princi!ActEco_Dep_TipDoc_Ant
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant = g_rst_Princi!ActEco_Dep_NumDoc_Ant
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant = "9"
         
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")
         Else
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
      End If
      
      'Independiente
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc = g_rst_Princi!ActEco_Ind_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc = Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipVia = g_rst_Princi!ActEco_Ind_TipVia
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomVia = Trim(g_rst_Princi!ActEco_Ind_NomVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumVia = Trim(g_rst_Princi!ActEco_Ind_NumVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IntDpt = Trim(g_rst_Princi!ActEco_Ind_IntDpt & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipZon = g_rst_Princi!ActEco_Ind_TipZon
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomZon = Trim(g_rst_Princi!ActEco_Ind_NomZon & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_UbiGeo = Trim(g_rst_Princi!ActEco_Ind_UbiGeo & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Refere = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1 = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2 = Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumFax = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCiu = g_rst_Princi!ActEco_Ind_CodCiu
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IngNet = g_rst_Princi!ActEco_Ind_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc = g_rst_Princi!ActEco_Ind_ConLoc
      
      If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc = 1 Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp = g_rst_Princi!ActEco_Ind_TipDoc_Emp
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp = g_rst_Princi!ActEco_Ind_NumDoc_Emp
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FlgEmp = "9"
         
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")
         Else
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCar = Trim(g_rst_Princi!ActEco_Ind_CodCar & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCar = Trim(g_rst_Princi!ActEco_Ind_NomCar & "")
      End If
         
      'Comerciante
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipDoc = g_rst_Princi!ActEco_Com_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumDoc = Trim(g_rst_Princi!ActEco_Com_NumDoc & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RazSoc = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCom = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipVia = g_rst_Princi!ActEco_Com_TipVia
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomVia = Trim(g_rst_Princi!ActEco_Com_NomVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumVia = Trim(g_rst_Princi!ActEco_Com_NumVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IntDpt = Trim(g_rst_Princi!ActEco_Com_IntDpt & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipZon = g_rst_Princi!ActEco_Com_TipZon
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomZon = Trim(g_rst_Princi!ActEco_Com_NomZon & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_UbiGeo = Trim(g_rst_Princi!ActEco_Com_UbiGeo & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Refere = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef1 = Trim(g_rst_Princi!ActEco_Com_Telef1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef2 = Trim(g_rst_Princi!ActEco_Com_Telef2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumFax = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCiu = g_rst_Princi!ActEco_Com_CodCiu
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_GirCom = Trim(g_rst_Princi!ActEco_Com_GirCom & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IngNet = g_rst_Princi!ActEco_Com_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_VtaMen = g_rst_Princi!ActEco_Com_VtaMen
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCar = Trim(g_rst_Princi!ActEco_Com_CodCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCar = Trim(g_rst_Princi!ActEco_Com_NomCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RegTri = g_rst_Princi!ActEco_Com_RegTri
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_PorPar = g_rst_Princi!ActEco_Com_PorPar
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipLoc = g_rst_Princi!ActEco_Com_TipLoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_AlqMen = g_rst_Princi!ActEco_Com_AlqMen
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomArr = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TelArr = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
      
      'Accionista
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc = g_rst_Princi!ActEco_Acc_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia = g_rst_Princi!ActEco_Acc_TipVia
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Princi!ActEco_Acc_NomVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Princi!ActEco_Acc_NumVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Princi!ActEco_Acc_IntDpt & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon = g_rst_Princi!ActEco_Acc_TipZon
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Princi!ActEco_Acc_NomZon & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Princi!ActEco_Acc_UbiGeo & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Princi!ActEco_Acc_Refere & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu = g_rst_Princi!ActEco_Acc_CodCiu
      Else
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia = g_rst_Genera!DatGen_TipVia
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon = g_rst_Genera!DatGen_TipZon
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu = g_rst_Genera!DATGEN_CODCIU
      End If
         
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IngNet = g_rst_Princi!ActEco_Acc_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_PorPar = g_rst_Princi!ActEco_Acc_PorPar
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
      
      'Rentista
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IngNet = g_rst_Princi!ActEco_Ren_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc1 = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr1 = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele11 = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele21 = Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe1 = g_rst_Princi!ActEco_Ren_AlqMe1
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_SegPro = g_rst_Princi!ActEco_Ren_SegPro
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc2 = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr2 = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele12 = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele22 = Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe2 = g_rst_Princi!ActEco_Ren_AlqMe2
      
      'Otros
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_IngNet = g_rst_Princi!ActEco_Otr_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Activi = Trim(g_rst_Princi!ActEco_Otr_Activi & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_CodCiu = g_rst_Princi!ActEco_Otr_CodCiu
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Observ = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

