VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   10410
   ClientLeft      =   2340
   ClientTop       =   585
   ClientWidth     =   11685
   Icon            =   "AteCli_frm_104.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10395
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   8295
         Left            =   30
         TabIndex        =   51
         Top             =   1230
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   14631
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
         Begin VB.ComboBox cmb_SitTra 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TraAnt 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   6270
            Width           =   945
         End
         Begin VB.TextBox txt_Telef2_Ant 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   7590
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef1_Ant 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   7590
            Width           =   1640
         End
         Begin VB.TextBox txt_NomCom_Ant 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   7260
            Width           =   9525
         End
         Begin VB.TextBox txt_RazSoc_Ant 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   37
            Text            =   "Text1"
            Top             =   6930
            Width           =   9525
         End
         Begin VB.ComboBox cmb_TipDoc_Ant 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   6600
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc_Ant 
            Height          =   315
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   6600
            Width           =   2355
         End
         Begin VB.CommandButton cmd_BusEmp_Ant 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   36
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   6600
            Width           =   435
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   60
            Left            =   60
            TabIndex        =   82
            Top             =   6150
            Width           =   11475
            _Version        =   65536
            _ExtentX        =   20241
            _ExtentY        =   106
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
            BevelOuter      =   1
         End
         Begin VB.ComboBox cmb_TipOfi 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   2010
            TabIndex        =   6
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   9525
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1050
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   3690
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   8220
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8220
            MaxLength       =   15
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   2370
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9870
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2370
            Width           =   1665
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8220
            MaxLength       =   120
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2700
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8220
            TabIndex        =   17
            Text            =   "cmb_PrvDir"
            Top             =   3030
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   3360
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   3690
            Width           =   1640
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   9525
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2355
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2700
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   16
            Text            =   "cmb_DptDir"
            Top             =   3030
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   18
            Text            =   "cmb_DstDir"
            Top             =   3360
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   3690
            Width           =   1640
         End
         Begin VB.TextBox txt_AnexRH 
            Height          =   315
            Left            =   3660
            MaxLength       =   5
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1710
            Width           =   1640
         End
         Begin VB.TextBox txt_TeleRH 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1710
            Width           =   1640
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   5460
            Width           =   1640
         End
         Begin VB.TextBox txt_NomAre 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   5130
            Width           =   3315
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   5790
            Width           =   3315
         End
         Begin VB.TextBox txt_NumAnx 
            Height          =   315
            Left            =   8220
            MaxLength       =   5
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   5130
            Width           =   1640
         End
         Begin VB.TextBox txt_TelDir 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   5460
            Width           =   1640
         End
         Begin VB.TextBox txt_NomCar 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   4800
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NomCar 
            Height          =   315
            Left            =   2010
            TabIndex        =   26
            Text            =   "cmb_Dep_NomCar"
            Top             =   4800
            Width           =   3315
         End
         Begin VB.ComboBox cmb_FreHab 
            Height          =   315
            Left            =   8220
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   4140
            Width           =   3315
         End
         Begin VB.CommandButton cmd_BusEmp_Act 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   3
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   390
            Width           =   435
         End
         Begin Threed.SSPanel pnl_FlgEmp_Act 
            Height          =   315
            Left            =   11100
            TabIndex        =   52
            Top             =   390
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NR"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2010
            TabIndex        =   23
            Top             =   4140
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
         Begin EditLib.fpDateTime ipp_FecIng 
            Height          =   315
            Left            =   2010
            TabIndex        =   25
            Top             =   4470
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
         Begin EditLib.fpDateTime ipp_FecCes_Ant 
            Height          =   315
            Left            =   8220
            TabIndex        =   42
            Top             =   7950
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
            TabIndex        =   83
            Top             =   4050
            Width           =   11475
            _Version        =   65536
            _ExtentX        =   20241
            _ExtentY        =   106
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
            BevelOuter      =   1
         End
         Begin Threed.SSPanel pnl_FlgEmp_Ant 
            Height          =   315
            Left            =   11100
            TabIndex        =   84
            Top             =   6600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NR"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
         End
         Begin EditLib.fpDateTime ipp_FecIng_Ant 
            Height          =   315
            Left            =   2010
            TabIndex        =   41
            Top             =   7920
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
         Begin VB.Label lbl_General 
            Caption         =   "Situación Trabajador:"
            Height          =   285
            Index           =   8
            Left            =   90
            TabIndex        =   93
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label lbl_General 
            Caption         =   "Trabajo Anterior:"
            Height          =   285
            Index           =   7
            Left            =   90
            TabIndex        =   92
            Top             =   6270
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   6
            Left            =   90
            TabIndex        =   90
            Top             =   7920
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   89
            Top             =   7590
            Width           =   1815
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   88
            Top             =   7260
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   87
            Top             =   6600
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   2
            Left            =   6210
            TabIndex        =   86
            Top             =   6600
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   85
            Top             =   6930
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Celular de la Empresa:"
            Height          =   285
            Index           =   0
            Left            =   6210
            TabIndex        =   81
            Top             =   5460
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   49
            Left            =   90
            TabIndex        =   80
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   41
            Left            =   6210
            TabIndex        =   79
            Top             =   2040
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Oficina:"
            Height          =   285
            Index           =   40
            Left            =   90
            TabIndex        =   78
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6210
            TabIndex        =   77
            Top             =   2370
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6210
            TabIndex        =   76
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6210
            TabIndex        =   75
            Top             =   3030
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6210
            TabIndex        =   74
            Top             =   3360
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6210
            TabIndex        =   73
            Top             =   3690
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   42
            Left            =   90
            TabIndex        =   72
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   90
            TabIndex        =   71
            Top             =   2700
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   90
            TabIndex        =   70
            Top             =   3030
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   90
            TabIndex        =   69
            Top             =   3360
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   68
            Top             =   3690
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   36
            Left            =   90
            TabIndex        =   67
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   6210
            TabIndex        =   66
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   37
            Left            =   90
            TabIndex        =   65
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   90
            TabIndex        =   64
            Top             =   1380
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono/Anexo RR.HH:"
            Height          =   285
            Index           =   47
            Left            =   90
            TabIndex        =   63
            Top             =   1710
            Width           =   1815
         End
         Begin VB.Label lbl_General 
            Caption         =   "Area:"
            Height          =   285
            Index           =   63
            Left            =   90
            TabIndex        =   62
            Top             =   5130
            Width           =   1605
         End
         Begin VB.Label lbl_General 
            Caption         =   "E-mail:"
            Height          =   285
            Index           =   60
            Left            =   90
            TabIndex        =   61
            Top             =   5790
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo (Especificar):"
            Height          =   285
            Index           =   57
            Left            =   6210
            TabIndex        =   60
            Top             =   4800
            Width           =   1665
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Cese:"
            Height          =   315
            Index           =   66
            Left            =   6210
            TabIndex        =   59
            Top             =   7950
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "Anexo:"
            Height          =   285
            Index           =   64
            Left            =   6210
            TabIndex        =   58
            Top             =   5130
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   58
            Left            =   90
            TabIndex        =   57
            Top             =   4470
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   62
            Left            =   90
            TabIndex        =   56
            Top             =   4800
            Width           =   975
         End
         Begin VB.Label lbl_General 
            Caption         =   "Frecuencia Haberes:"
            Height          =   315
            Index           =   56
            Left            =   6210
            TabIndex        =   55
            Top             =   4140
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado (S/.):"
            Height          =   285
            Index           =   61
            Left            =   90
            TabIndex        =   54
            Top             =   4140
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono Directo:"
            Height          =   285
            Index           =   59
            Left            =   90
            TabIndex        =   53
            Top             =   5460
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   46
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
            TabIndex        =   47
            Top             =   60
            Width           =   10125
            _Version        =   65536
            _ExtentX        =   17859
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes - Actividades Económicas - Dependiente o Pensionista"
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
            Picture         =   "AteCli_frm_104.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   48
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
            TabIndex        =   49
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
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   91
         Top             =   9570
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
            Picture         =   "AteCli_frm_104.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Simulación de Créditos Hipotecarios"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10200
            Picture         =   "AteCli_frm_104.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10890
            Picture         =   "AteCli_frm_104.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_NomCar()   As moddat_tpo_Genera
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_str_CodCiu     As String
Dim l_str_NomCar     As String
Dim l_int_FlgCmb     As Integer

Private Sub cmb_CodCiu_Change()
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_Click()
   If cmb_CodCiu.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_TeleRH)
      End If
   End If
End Sub

Private Sub cmb_CodCiu_GotFocus()
   l_int_FlgCmb = True
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + modgen_g_con_NUMERO + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CodCiu, l_str_CodCiu)
      l_int_FlgCmb = True
      
      If cmb_CodCiu.ListIndex > -1 Then
         l_str_CodCiu = ""
      End If
      
      Call gs_SetFocus(txt_TeleRH)
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

Private Sub cmb_DstDir_LostFocus()
   Call SendMessage(cmb_DstDir.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_FreHab_Click()
   Call gs_SetFocus(ipp_FecIng)
End Sub

Private Sub cmb_FreHab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIng)
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

Private Sub cmb_SitTra_Click()
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub cmb_SitTra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SitTra_Click
   End If
End Sub

Private Sub cmb_TipDoc_Ant_Click()
   If cmb_TipDoc_Ant.ListIndex > -1 Then
      Select Case cmb_TipDoc_Ant.ItemData(cmb_TipDoc_Ant.ListIndex)
         Case 1:     txt_NumDoc_Ant.MaxLength = 8
         Case 7:     txt_NumDoc_Ant.MaxLength = 11
         Case Else:  txt_NumDoc_Ant.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc_Ant)
End Sub

Private Sub cmb_TipDoc_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Ant_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumDoc.MaxLength = 8
         Case 7:     txt_NumDoc.MaxLength = 11
         Case Else:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmb_TipOfi_Click()
   If cmb_TipOfi.ListIndex > -1 Then
      'Si es Oficina Principal
      If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 1 Then
         If pnl_FlgEmp_Act.Tag = "9" Or Len(Trim(pnl_FlgEmp_Act.Tag)) = 0 Then
            If cmb_TipVia.Enabled Then
               Call gs_SetFocus(cmb_TipVia)
            Else
               Call gs_SetFocus(ipp_IngNet)
            End If
         Else
            'Buscar Dirección en Empresas
            g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & txt_NumDoc.Text & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
               Exit Sub
            End If
            
            g_rst_Princi.MoveFirst
   
            Call gs_BuscarCombo_Item(cmb_TipVia, CInt(Trim(g_rst_Princi!DatGen_TipVia)))
            txt_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
            txt_NumVia.Text = Trim(g_rst_Princi!DatGen_numVia & "")
            txt_IntDpt.Text = Trim(g_rst_Princi!DatGen_IntDpt & "")
            Call gs_BuscarCombo_Item(cmb_TipZon, CInt(Trim(g_rst_Princi!DatGen_TipZon)))
            txt_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
            txt_Refere.Text = Trim(g_rst_Princi!DatGen_Refere & "")
               
            If g_rst_Princi!DatGen_Ubigeo <> "000000" Then
               Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 1, 2)))
               Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
               Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 3, 2)))
               Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
               Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Mid(Trim(g_rst_Princi!DatGen_Ubigeo), 5, 2)))
            Else
               cmb_DptDir.ListIndex = -1
               cmb_PrvDir.Clear
               cmb_DstDir.Clear
            End If
               
            txt_Telef1.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
            txt_Telef2.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
            txt_NumFax.Text = Trim(g_rst_Princi!DatGen_NUMFAX & "")
         
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
         
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
            txt_Telef1.Enabled = False
            txt_Telef2.Enabled = False
            txt_NumFax.Enabled = False
         End If
      Else
         cmb_TipVia.Enabled = True
         txt_NomVia.Enabled = True
         txt_NumVia.Enabled = True
         txt_IntDpt.Enabled = True
         cmb_TipZon.Enabled = True
         txt_NomZon.Enabled = True
         cmb_DptDir.Enabled = True
         cmb_PrvDir.Enabled = True
         cmb_DstDir.Enabled = True
         txt_Refere.Enabled = True
         txt_Telef1.Enabled = True
         txt_Telef2.Enabled = True
         txt_NumFax.Enabled = True
      
         Call gs_SetFocus(cmb_TipVia)
      End If
   End If
End Sub

Private Sub cmb_TipOfi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipOfi_Click
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

Private Sub cmb_TraAnt_Click()
   If cmb_TraAnt.ListIndex > -1 Then
      If cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex) = 1 Then
         cmb_TipDoc_Ant.Enabled = True
         txt_NumDoc_Ant.Enabled = True
         cmd_BusEmp_Ant.Enabled = True
         'txt_RazSoc_Ant.Enabled = True
         'txt_NomCom_Ant.Enabled = True
         'txt_Telef1_Ant.Enabled = True
         'txt_Telef2_Ant.Enabled = True
         ipp_FecIng_Ant.Enabled = True
         ipp_FecCes_Ant.Enabled = True
         
         Call gs_SetFocus(cmb_TipDoc_Ant)
      Else
         cmb_TipDoc_Ant.ListIndex = -1
         txt_NumDoc_Ant.Text = ""
         pnl_FlgEmp_Ant.Caption = ""
         
         txt_RazSoc_Ant.Text = ""
         txt_NomCom_Ant.Text = ""
         txt_Telef1_Ant.Text = ""
         txt_Telef2_Ant.Text = ""
         ipp_FecIng_Ant.Text = Format(Date, "dd/mm/yyyy")
         ipp_FecCes_Ant.Text = Format(Date, "dd/mm/yyyy")
         
         cmb_TipDoc_Ant.Enabled = False
         txt_NumDoc_Ant.Enabled = False
         cmd_BusEmp_Ant.Enabled = False
         txt_RazSoc_Ant.Enabled = False
         txt_NomCom_Ant.Enabled = False
         txt_Telef1_Ant.Enabled = False
         txt_Telef2_Ant.Enabled = False
         ipp_FecIng_Ant.Enabled = False
         ipp_FecCes_Ant.Enabled = False
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   Else
      cmb_TipDoc_Ant.ListIndex = -1
      txt_NumDoc_Ant.Text = ""
      pnl_FlgEmp_Ant.Caption = ""
      
      txt_RazSoc_Ant.Text = ""
      txt_NomCom_Ant.Text = ""
      txt_Telef1_Ant.Text = ""
      txt_Telef2_Ant.Text = ""
      ipp_FecIng_Ant.Text = Format(Date, "dd/mm/yyyy")
      ipp_FecCes_Ant.Text = Format(Date, "dd/mm/yyyy")
      
      cmb_TipDoc_Ant.Enabled = False
      txt_NumDoc_Ant.Enabled = False
      cmd_BusEmp_Ant.Enabled = False
      txt_RazSoc_Ant.Enabled = False
      txt_NomCom_Ant.Enabled = False
      txt_Telef1_Ant.Enabled = False
      txt_Telef2_Ant.Enabled = False
      ipp_FecIng_Ant.Enabled = False
      ipp_FecCes_Ant.Enabled = False
   End If
End Sub

Private Sub cmb_TraAnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TraAnt_Click
   End If
End Sub

Private Sub cmd_BusEmp_Act_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc.Text)) <> 11 Then
         MsgBox "El Número de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   End If
   
   'Buscando Empresa
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & txt_NumDoc.Text & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      pnl_FlgEmp_Act.Caption = moddat_gf_Consulta_ParDes("019", "9")
      pnl_FlgEmp_Act.Tag = "9"
      
      txt_RazSoc.Enabled = True
      txt_NomCom.Enabled = True
      cmb_CodCiu.Enabled = True
      txt_TeleRH.Enabled = True
      txt_AnexRH.Enabled = True
      cmb_TipOfi.Enabled = True
      cmb_TipVia.Enabled = True
      txt_NomVia.Enabled = True
      txt_NumVia.Enabled = True
      txt_IntDpt.Enabled = True
      cmb_TipZon.Enabled = True
      txt_NomZon.Enabled = True
      cmb_DptDir.Enabled = True
      cmb_PrvDir.Enabled = True
      cmb_DstDir.Enabled = True
      txt_Refere.Enabled = True
      txt_Telef1.Enabled = True
      txt_Telef2.Enabled = True
      txt_NumFax.Enabled = True
      
      txt_RazSoc.Text = ""
      txt_NomCom.Text = ""
      cmb_CodCiu.ListIndex = -1
      txt_TeleRH.Text = ""
      txt_AnexRH.Text = ""
      cmb_TipVia.ListIndex = -1
      txt_NomVia.Text = ""
      txt_NumVia.Text = ""
      txt_IntDpt.Text = ""
      cmb_TipZon.ListIndex = -1
      txt_NomZon.Text = ""
      cmb_DptDir.ListIndex = -1
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      txt_Refere.Text = ""
      txt_Telef1.Text = ""
      txt_Telef2.Text = ""
      txt_NumFax.Text = ""
      
      Call gs_SetFocus(txt_RazSoc)
   Else
      g_rst_Princi.MoveFirst
   
      pnl_FlgEmp_Act.Caption = moddat_gf_Consulta_ParDes("019", g_rst_Princi!DATGEN_CLASIF)
      pnl_FlgEmp_Act.Tag = CStr(g_rst_Princi!DATGEN_CLASIF)
      
      txt_RazSoc.Text = Trim(g_rst_Princi!DATGEN_RAZSOC)
      txt_NomCom.Text = Trim(g_rst_Princi!DATGEN_NOMCOM)
      
      Call gs_BuscarCombo_Item(cmb_CodCiu, g_rst_Princi!DATGEN_CODCIU)
      txt_TeleRH.Text = Trim(g_rst_Princi!DATGEN_TELERH & "")
      txt_AnexRH.Text = Trim(g_rst_Princi!DATGEN_ANEXRH & "")
      
      cmb_TipOfi.Enabled = True
      
      If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) And moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumDoc = txt_NumDoc.Text And moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = 2 Then
         'Call gs_BuscarCombo_Item(cmb_TipVia, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipVia)
         'txt_NomVia.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomVia
         'txt_NumVia.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumVia
         'txt_IntDpt.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_IntDpt
         'Call gs_BuscarCombo_Item(cmb_TipZon, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipZon)
         'txt_NomZon.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomZon
         'Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2)))
         'Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2))
         'Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 3, 2)))
         'Call moddat_gs_Carga_Distri(cmb_DstDir, Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2), Mid(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 3, 2))
         'Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2)))
         'txt_Refere.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Refere
         'txt_Telef1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef1
         'txt_Telef2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef2
         'txt_NumFax.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumFax
         
         'cmb_TipVia.Enabled = True
         'txt_NomVia.Enabled = True
         'txt_NumVia.Enabled = True
         'txt_IntDpt.Enabled = True
         'cmb_TipZon.Enabled = True
         'txt_NomZon.Enabled = True
         'cmb_DptDir.Enabled = True
         'cmb_PrvDir.Enabled = True
         'cmb_DstDir.Enabled = True
         'txt_Refere.Enabled = True
         'txt_Telef1.Enabled = True
         'txt_Telef2.Enabled = True
         'txt_NumFax.Enabled = True
         
         Call gs_BuscarCombo_Item(cmb_TipOfi, 2)
      Else
         Call gs_BuscarCombo_Item(cmb_TipOfi, 1)
      End If
      
         
      txt_RazSoc.Enabled = False
      txt_NomCom.Enabled = False
      cmb_CodCiu.Enabled = False
      txt_TeleRH.Enabled = False
      txt_AnexRH.Enabled = False
      
      Call gs_SetFocus(cmb_TipOfi)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BusEmp_Ant_Click()
   If cmb_TipDoc_Ant.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc_Ant)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc_Ant.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc_Ant)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc_Ant.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc_Ant.Text)) <> 11 Then
         MsgBox "El Número de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Ant)
         Exit Sub
      End If
   End If
   
   'Buscando Empresa
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(cmb_TipDoc_Ant.ItemData(cmb_TipDoc_Ant.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & txt_NumDoc_Ant.Text & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      pnl_FlgEmp_Ant.Caption = moddat_gf_Consulta_ParDes("019", "9")
      pnl_FlgEmp_Ant.Tag = "9"
      
      txt_RazSoc_Ant.Enabled = True
      txt_NomCom_Ant.Enabled = True
      txt_Telef1_Ant.Enabled = True
      txt_Telef2_Ant.Enabled = True
      
      txt_RazSoc_Ant.Text = ""
      txt_NomCom_Ant.Text = ""
      txt_Telef1_Ant.Text = ""
      txt_Telef2_Ant.Text = ""
      
      Call gs_SetFocus(txt_RazSoc_Ant)
   Else
      g_rst_Princi.MoveFirst
   
      pnl_FlgEmp_Ant.Caption = moddat_gf_Consulta_ParDes("019", g_rst_Princi!DATGEN_CLASIF)
      pnl_FlgEmp_Ant.Tag = CStr(g_rst_Princi!DATGEN_CLASIF)
      
      txt_RazSoc_Ant.Text = Trim(g_rst_Princi!DATGEN_RAZSOC)
      txt_NomCom_Ant.Text = Trim(g_rst_Princi!DATGEN_NOMCOM)
      
      txt_Telef1_Ant.Text = Trim(g_rst_Princi!DATGEN_TELEF1 & "")
      txt_Telef2_Ant.Text = Trim(g_rst_Princi!DATGEN_TELEF2 & "")
         
      txt_RazSoc_Ant.Enabled = False
      txt_NomCom_Ant.Enabled = False
      txt_Telef1_Ant.Enabled = False
      txt_Telef2_Ant.Enabled = False
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_SitTra.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación del Cliente como Trabajador.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitTra)
      Exit Sub
   End If
   
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc.Text)) <> 11 Then
         MsgBox "El Número de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   End If

   If Len(Trim(txt_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RazSoc)
      Exit Sub
   End If

   If Len(Trim(txt_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomCom)
      Exit Sub
   End If

   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar el CIIU.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If

   If cmb_TipOfi.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Oficina.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOfi)
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

   If Len(Trim(txt_Telef1.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Telef1)
      Exit Sub
   End If


   If ipp_IngNet.Value = 0 Then
      MsgBox "El Ingreso Declarado no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If
   
   If CDate(ipp_FecIng.Text) > Date Then
      MsgBox "La Fecha de Ingreso no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIng)
      Exit Sub
   End If
   
   If cmb_FreHab.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Frecuencia de Haberes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FreHab)
      Exit Sub
   End If
   
   If cmb_NomCar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Cargo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NomCar)
      Exit Sub
   End If
   
   If l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
      If Len(Trim(txt_NomCar.Text)) = 0 Then
         MsgBox "Debe ingresar el Cargo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomCar)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_NomAre.Text)) = 0 Then
      MsgBox "Debe ingresar el Area.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomAre)
      Exit Sub
   End If
   
   If cmb_TraAnt.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el cliente brinda información del trabajo anterior.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TraAnt)
      Exit Sub
   End If
   
   If cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex) = 1 Then
      If cmb_TipDoc_Ant.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc_Ant)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc_Ant.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc_Ant)
         Exit Sub
      End If
      
      If cmb_TipDoc_Ant.ItemData(cmb_TipDoc_Ant.ListIndex) = 7 Then
         If Len(Trim(txt_NumDoc_Ant.Text)) <> 11 Then
            MsgBox "El Número de Documento ingresado no es correcto.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumDoc_Ant)
            Exit Sub
         End If
      End If
   
      If Len(Trim(txt_RazSoc_Ant.Text)) = 0 Then
         MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_RazSoc_Ant)
         Exit Sub
      End If
   
      If Len(Trim(txt_NomCom_Ant.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomCom_Ant)
         Exit Sub
      End If
   
      If Len(Trim(txt_Telef1_Ant.Text)) = 0 Then
         MsgBox "Debe ingresar el Teléfono.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Telef1_Ant)
         Exit Sub
      End If
   
      If CDate(ipp_FecIng_Ant.Text) > CDate(ipp_FecIng.Text) Then
         MsgBox "La Fecha de Ingreso del trabajo anterio no puede ser mayor a la Fecha de Ingreso del trabajo actual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIng_Ant)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_Inicia_ActEco(moddat_g_int_TipCli, moddat_g_int_OrdAct)
   
   If moddat_g_int_TipCli = 1 Then
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 11
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_SitTra = cmb_SitTra.ItemData(cmb_SitTra.ListIndex)
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumDoc = txt_NumDoc.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp = pnl_FlgEmp_Act.Tag
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_RazSoc = txt_RazSoc.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomCom = txt_NomCom.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TeleRH = txt_TeleRH.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_AnexRH = txt_AnexRH.Text
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomVia = txt_NomVia.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumVia = txt_NumVia.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_IntDpt = txt_IntDpt.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomZon = txt_NomZon.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo = Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00")
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Refere = txt_Refere.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef1 = txt_Telef1.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef2 = txt_Telef2.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumFax = txt_NumFax.Text
      
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_IngNet = CDbl(ipp_IngNet.Text)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FreHab = cmb_FreHab.ItemData(cmb_FreHab.ListIndex)
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FecIng = ipp_FecIng.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_CodCar = l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomCar = txt_NomCar.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomAre = txt_NomAre.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumAnx = txt_NumAnx.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TelDir = txt_TelDir.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Celula = txt_Celula.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_DirEle = txt_DirEle.Text
      moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TraAnt = cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex)
      
      If cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex) = 1 Then
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipDoc_Ant = cmb_TipDoc_Ant.ItemData(cmb_TipDoc_Ant.ListIndex)
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumDoc_Ant = txt_NumDoc_Ant.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp_Ant = pnl_FlgEmp_Ant.Tag
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_RazSoc_Ant = txt_RazSoc_Ant.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomCom_Ant = txt_NomCom_Ant.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef1_Ant = txt_Telef1_Ant.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef2_Ant = txt_Telef2_Ant.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FecIng_Ant = ipp_FecIng_Ant.Text
         moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FecCes_Ant = ipp_FecCes_Ant.Text
      End If
   Else
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_OrdAct = moddat_g_int_OrdAct
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 11
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_SitTra = cmb_SitTra.ItemData(cmb_SitTra.ListIndex)
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumDoc = txt_NumDoc.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp = pnl_FlgEmp_Act.Tag
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_RazSoc = txt_RazSoc.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomCom = txt_NomCom.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_CodCiu = cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TeleRH = txt_TeleRH.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_AnexRH = txt_AnexRH.Text
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipVia = cmb_TipVia.ItemData(cmb_TipVia.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomVia = txt_NomVia.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumVia = txt_NumVia.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_IntDpt = txt_IntDpt.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipZon = cmb_TipZon.ItemData(cmb_TipZon.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomZon = txt_NomZon.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo = Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00")
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Refere = txt_Refere.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef1 = txt_Telef1.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef2 = txt_Telef2.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumFax = txt_NumFax.Text
      
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_IngNet = CDbl(ipp_IngNet.Text)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FreHab = cmb_FreHab.ItemData(cmb_FreHab.ListIndex)
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FecIng = ipp_FecIng.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_CodCar = l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomCar = txt_NomCar.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomAre = txt_NomAre.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumAnx = txt_NumAnx.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TelDir = txt_TelDir.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Celula = txt_Celula.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_DirEle = txt_DirEle.Text
      moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TraAnt = cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex)
      
      If cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex) = 1 Then
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipDoc_Ant = cmb_TipDoc_Ant.ItemData(cmb_TipDoc_Ant.ListIndex)
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumDoc_Ant = txt_NumDoc_Ant.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp_Ant = pnl_FlgEmp_Ant.Tag
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_RazSoc_Ant = txt_RazSoc_Ant.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomCom_Ant = txt_NomCom_Ant.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef1_Ant = txt_Telef1_Ant.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef2_Ant = txt_Telef2_Ant.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FecIng_Ant = ipp_FecIng_Ant.Text
         moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FecCes_Ant = ipp_FecCes_Ant.Text
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
      
      If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_TipAct = 11 Then
         Call gs_BuscarCombo_Item(cmb_SitTra, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_SitTra)
         
         Call gs_BuscarCombo_Item(cmb_TipDoc, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipDoc)
         txt_NumDoc.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumDoc
         
         pnl_FlgEmp_Act.Tag = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp
         pnl_FlgEmp_Act.Caption = moddat_gf_Consulta_ParDes("019", pnl_FlgEmp_Act.Tag)
         
         txt_RazSoc.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_RazSoc
         txt_NomCom.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomCom
         
         Call gs_BuscarCombo_Item(cmb_CodCiu, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_CodCiu)
         txt_TeleRH.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TeleRH
         txt_AnexRH.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_AnexRH
         
         Call gs_BuscarCombo_Item(cmb_TipOfi, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipOfi)
         
         Call gs_BuscarCombo_Item(cmb_TipVia, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipVia)
         txt_NomVia.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomVia
         txt_NumVia.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumVia
         txt_IntDpt.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_IntDpt
         Call gs_BuscarCombo_Item(cmb_TipZon, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipZon)
         txt_NomZon.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomZon
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2), Mid(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2)))
         txt_Refere.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Refere
         txt_Telef1.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef1
         txt_Telef2.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef2
         txt_NumFax.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumFax
         
         cmb_TipOfi.Enabled = True
         If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp = 9 Then
            cmb_TipVia.Enabled = True
            txt_NomVia.Enabled = True
            txt_NumVia.Enabled = True
            txt_IntDpt.Enabled = True
            cmb_TipZon.Enabled = True
            txt_NomZon.Enabled = True
            cmb_DptDir.Enabled = True
            cmb_PrvDir.Enabled = True
            cmb_DstDir.Enabled = True
            txt_Refere.Enabled = True
            txt_Telef1.Enabled = True
            txt_Telef2.Enabled = True
            txt_NumFax.Enabled = True
         Else
            If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 2 Then
               cmb_TipVia.Enabled = True
               txt_NomVia.Enabled = True
               txt_NumVia.Enabled = True
               txt_IntDpt.Enabled = True
               cmb_TipZon.Enabled = True
               txt_NomZon.Enabled = True
               cmb_DptDir.Enabled = True
               cmb_PrvDir.Enabled = True
               cmb_DstDir.Enabled = True
               txt_Refere.Enabled = True
               txt_Telef1.Enabled = True
               txt_Telef2.Enabled = True
               txt_NumFax.Enabled = True
            Else
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
               txt_Telef1.Enabled = False
               txt_Telef2.Enabled = False
               txt_NumFax.Enabled = False
            End If
         End If
         
         ipp_IngNet.Value = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_IngNet
         
         Call gs_BuscarCombo_Item(cmb_FreHab, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FreHab)
         ipp_FecIng.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FecIng
         
         cmb_NomCar.ListIndex = gf_Busca_Arregl(l_arr_NomCar, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_CodCar) - 1
         txt_NomCar.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomCar
         txt_NomAre.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomAre
         txt_NumAnx.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumAnx
         txt_TelDir.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TelDir
         txt_Celula.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Celula
         txt_DirEle.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_DirEle
         
         Call gs_BuscarCombo_Item(cmb_TraAnt, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TraAnt)
         
         If cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex) = 1 Then
            Call gs_BuscarCombo_Item(cmb_TipDoc_Ant, moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipDoc_Ant)
            txt_NumDoc_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NumDoc_Ant
            
            pnl_FlgEmp_Ant.Tag = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp_Ant
            pnl_FlgEmp_Ant.Caption = moddat_gf_Consulta_ParDes("019", pnl_FlgEmp_Ant.Tag)
            
            txt_RazSoc_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_RazSoc_Ant
            txt_NomCom_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_NomCom_Ant
            
            txt_Telef1_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef1_Ant
            txt_Telef2_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_Telef2_Ant
         
            ipp_FecIng_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FecIng_Ant
            ipp_FecCes_Ant.Text = moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_FecCes_Ant
         End If
      End If
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & "(" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   
      If moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_TipAct = 11 Then
         Call gs_BuscarCombo_Item(cmb_SitTra, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_SitTra)
         
         Call gs_BuscarCombo_Item(cmb_TipDoc, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipDoc)
         txt_NumDoc.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumDoc
         
         pnl_FlgEmp_Act.Tag = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp
         pnl_FlgEmp_Act.Caption = moddat_gf_Consulta_ParDes("019", pnl_FlgEmp_Act.Tag)
         
         txt_RazSoc.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_RazSoc
         txt_NomCom.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomCom
         
         Call gs_BuscarCombo_Item(cmb_CodCiu, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_CodCiu)
         txt_TeleRH.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TeleRH
         txt_AnexRH.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_AnexRH
         
         Call gs_BuscarCombo_Item(cmb_TipOfi, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipOfi)
         
         Call gs_BuscarCombo_Item(cmb_TipVia, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipVia)
         txt_NomVia.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomVia
         txt_NumVia.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumVia
         txt_IntDpt.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_IntDpt
         Call gs_BuscarCombo_Item(cmb_TipZon, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipZon)
         txt_NomZon.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomZon
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2), Mid(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_UbiGeo, 2)))
         txt_Refere.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Refere
         txt_Telef1.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef1
         txt_Telef2.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef2
         txt_NumFax.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumFax
         
         cmb_TipOfi.Enabled = True
         If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 2 Then
            cmb_TipVia.Enabled = True
            txt_NomVia.Enabled = True
            txt_NumVia.Enabled = True
            txt_IntDpt.Enabled = True
            cmb_TipZon.Enabled = True
            txt_NomZon.Enabled = True
            cmb_DptDir.Enabled = True
            cmb_PrvDir.Enabled = True
            cmb_DstDir.Enabled = True
            txt_Refere.Enabled = True
            txt_Telef1.Enabled = True
            txt_Telef2.Enabled = True
            txt_NumFax.Enabled = True
         Else
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
            txt_Telef1.Enabled = False
            txt_Telef2.Enabled = False
            txt_NumFax.Enabled = False
         End If
         
         ipp_IngNet.Value = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_IngNet
         
         Call gs_BuscarCombo_Item(cmb_FreHab, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FreHab)
         ipp_FecIng.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FecIng
         
         cmb_NomCar.ListIndex = gf_Busca_Arregl(l_arr_NomCar, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_CodCar) - 1
         txt_NomCar.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomCar
         txt_NomAre.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomAre
         txt_NumAnx.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumAnx
         txt_TelDir.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TelDir
         txt_Celula.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Celula
         txt_DirEle.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_DirEle
         
         Call gs_BuscarCombo_Item(cmb_TraAnt, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TraAnt)
         
         If cmb_TraAnt.ItemData(cmb_TraAnt.ListIndex) = 1 Then
            Call gs_BuscarCombo_Item(cmb_TipDoc_Ant, moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipDoc_Ant)
            txt_NumDoc_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NumDoc_Ant
            
            pnl_FlgEmp_Ant.Tag = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FlgEmp_Ant
            pnl_FlgEmp_Ant.Caption = moddat_gf_Consulta_ParDes("019", pnl_FlgEmp_Ant.Tag)
            
            txt_RazSoc_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_RazSoc_Ant
            txt_NomCom_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_NomCom_Ant
            
            txt_Telef1_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef1_Ant
            txt_Telef2_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_Telef2_Ant
         
            ipp_FecIng_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FecIng_Ant
            ipp_FecCes_Ant.Text = moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_FecCes_Ant
         End If
      End If
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "232")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipOfi, 1, "234")
   Call moddat_gs_Carga_LisIte_Combo(cmb_SitTra, 1, "235")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FreHab, 1, "210")
   Call moddat_gs_Carga_LisIte(cmb_NomCar, l_arr_NomCar, 1, "503")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TraAnt, 1, "214")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Ant, 1, "232")
End Sub

Private Sub fs_Limpia()
   cmb_SitTra.ListIndex = -1
   
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   pnl_FlgEmp_Act.Caption = ""
   pnl_FlgEmp_Act.Tag = ""
   
   txt_RazSoc.Text = ""
   txt_NomCom.Text = ""
   cmb_CodCiu.ListIndex = -1
   txt_TeleRH.Text = ""
   txt_AnexRH.Text = ""
   cmb_TipOfi.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_NumFax.Text = ""
   
   txt_RazSoc.Enabled = False
   txt_NomCom.Enabled = False
   cmb_CodCiu.Enabled = False
   txt_TeleRH.Enabled = False
   txt_AnexRH.Enabled = False
   cmb_TipOfi.Enabled = False
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
   txt_Telef1.Enabled = False
   txt_Telef2.Enabled = False
   txt_NumFax.Enabled = False
   
   ipp_IngNet.Value = 0
   cmb_FreHab.ListIndex = -1
   ipp_FecIng.Text = Format(Date, "dd/mm/yyyy")
   cmb_NomCar.ListIndex = -1
   txt_NomCar.Text = ""
   txt_NomCar.Enabled = False
   txt_NomAre.Text = ""
   txt_NumAnx.Text = ""
   txt_TelDir.Text = ""
   txt_Celula.Text = ""
   txt_DirEle.Text = ""
   
   cmb_TraAnt.ListIndex = -1
   
   cmb_TipDoc_Ant.ListIndex = -1
   txt_NumDoc_Ant.Text = ""
   pnl_FlgEmp_Ant.Caption = ""
   
   txt_RazSoc_Ant.Text = ""
   txt_NomCom_Ant.Text = ""
   txt_Telef1_Ant.Text = ""
   txt_Telef2_Ant.Text = ""
   ipp_FecIng_Ant.Text = Format(Date, "dd/mm/yyyy")
   ipp_FecCes_Ant.Text = Format(Date, "dd/mm/yyyy")
   
   cmb_TipDoc_Ant.Enabled = False
   txt_NumDoc_Ant.Enabled = False
   cmd_BusEmp_Ant.Enabled = False
   txt_RazSoc_Ant.Enabled = False
   txt_NomCom_Ant.Enabled = False
   txt_Telef1_Ant.Enabled = False
   txt_Telef2_Ant.Enabled = False
   ipp_FecIng_Ant.Enabled = False
   ipp_FecCes_Ant.Enabled = False
End Sub

Private Sub ipp_FecCes_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_FecIng_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecCes_Ant)
   End If
End Sub

Private Sub ipp_FecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NomCar)
   End If
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_FreHab)
   End If
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

Private Sub txt_NumDoc_Ant_Change()
   pnl_FlgEmp_Ant.Caption = ""
   pnl_FlgEmp_Ant.Tag = ""
   
   txt_RazSoc_Ant.Text = ""
   txt_NomCom_Ant.Text = ""
   txt_Telef1_Ant.Text = ""
   txt_Telef2_Ant.Text = ""
   
   txt_RazSoc_Ant.Enabled = False
   txt_NomCom_Ant.Enabled = False
   txt_Telef1_Ant.Enabled = False
   txt_Telef2_Ant.Enabled = False
End Sub

Private Sub txt_NumDoc_Ant_GotFocus()
   Call gs_SelecTodo(txt_NumDoc_Ant)
End Sub

Private Sub txt_NumDoc_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(pnl_FlgEmp_Ant.Caption)) > 0 Then
         Call gs_SetFocus(ipp_FecIng_Ant)
      Else
         Call gs_SetFocus(cmd_BusEmp_Ant)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_Change()
   pnl_FlgEmp_Act.Caption = ""
   pnl_FlgEmp_Act.Tag = ""
   
   txt_RazSoc.Enabled = False
   txt_NomCom.Enabled = False
   cmb_CodCiu.Enabled = False
   txt_TeleRH.Enabled = False
   txt_AnexRH.Enabled = False
   cmb_TipOfi.Enabled = False
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
   txt_Telef1.Enabled = False
   txt_Telef2.Enabled = False
   txt_NumFax.Enabled = False
   
   txt_RazSoc.Text = ""
   txt_NomCom.Text = ""
   cmb_CodCiu.ListIndex = -1
   txt_TeleRH.Text = ""
   txt_AnexRH.Text = ""
   cmb_TipOfi.ListIndex = -1
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_NumFax.Text = ""
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

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(pnl_FlgEmp_Act.Caption)) > 0 Then
         Call gs_SetFocus(cmb_TipOfi)
      Else
         Call gs_SetFocus(cmd_BusEmp_Act)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()#%&")
   End If
End Sub

Private Sub txt_NomCom_GotFocus()
   Call gs_SelecTodo(txt_NomCom)
End Sub

Private Sub txt_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodCiu)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
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
      Call gs_SetFocus(txt_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumFax_GotFocus()
   Call gs_SelecTodo(txt_NumFax)
End Sub

Private Sub txt_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IngNet)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TeleRH_GotFocus()
   Call gs_SelecTodo(txt_TeleRH)
End Sub

Private Sub txt_TeleRH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_AnexRH)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_AnexRH_GotFocus()
   Call gs_SelecTodo(txt_AnexRH)
End Sub

Private Sub txt_AnexRH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipOfi)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub cmb_NomCar_Change()
   l_str_NomCar = cmb_NomCar.Text
End Sub

Private Sub cmb_NomCar_Click()
   txt_NomCar.Enabled = False
   txt_NomCar.Text = ""
   
   If cmb_NomCar.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
            txt_NomCar.Enabled = True
            Call gs_SetFocus(txt_NomCar)
         Else
            Call gs_SetFocus(txt_NomAre)
         End If
      End If
   End If
End Sub

Private Sub cmb_NomCar_GotFocus()
   l_int_FlgCmb = True
   l_str_NomCar = cmb_NomCar.Text
End Sub

Private Sub cmb_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ ./*+#,()<>" + Chr(34))
   Else
      txt_NomCar.Enabled = False
      txt_NomCar.Text = ""
      
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_NomCar, l_str_NomCar)
      l_int_FlgCmb = True
      
      If cmb_NomCar.ListIndex > -1 Then
         l_str_NomCar = ""
      End If
      
      If l_arr_NomCar(cmb_NomCar.ListIndex + 1).Genera_Codigo = "999999" Then
         txt_NomCar.Enabled = True
         Call gs_SetFocus(txt_NomCar)
      Else
         Call gs_SetFocus(txt_NomAre)
      End If
   End If
End Sub

Private Sub txt_NomCar_GotFocus()
   Call gs_SelecTodo(txt_NomCar)
End Sub

Private Sub txt_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomAre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_NomAre_GotFocus()
   Call gs_SelecTodo(txt_NomAre)
End Sub

Private Sub txt_NomAre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAnx)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_NumAnx_GotFocus()
   Call gs_SelecTodo(txt_NumAnx)
End Sub

Private Sub txt_NumAnx_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TelDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TelDir_GotFocus()
   Call gs_SelecTodo(txt_TelDir)
End Sub

Private Sub txt_TelDir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Celula)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
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

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TraAnt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_RazSoc_Ant_GotFocus()
   Call gs_SelecTodo(txt_RazSoc_Ant)
End Sub

Private Sub txt_RazSoc_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCom_Ant)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_NomCom_Ant_GotFocus()
   Call gs_SelecTodo(txt_NomCom_Ant)
End Sub

Private Sub txt_NomCom_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1_Ant)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,:()'#%&")
   End If
End Sub

Private Sub txt_Telef1_Ant_GotFocus()
   Call gs_SelecTodo(txt_Telef1_Ant)
End Sub

Private Sub txt_Telef1_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2_Ant)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef2_Ant_GotFocus()
   Call gs_SelecTodo(txt_Telef2_Ant)
End Sub

Private Sub txt_Telef2_Ant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIng_Ant)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub



