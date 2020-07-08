VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12120
   ClientLeft      =   2475
   ClientTop       =   375
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   12120
   ScaleWidth      =   11685
   Begin Threed.SSPanel SSPanel1 
      Height          =   12105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   21352
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
         Height          =   2775
         Left            =   30
         TabIndex        =   68
         Top             =   8460
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   4895
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
         Begin VB.ComboBox cmb_TraAnt 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   60
            Width           =   945
         End
         Begin VB.CommandButton Command3 
            Height          =   675
            Left            =   10860
            Picture         =   "AteCli_frm_136.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   510
            Width           =   675
         End
         Begin VB.CommandButton Command2 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_136.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Limpiar Datos"
            Top             =   510
            Width           =   675
         End
         Begin VB.CommandButton Command1 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_136.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Buscar Datos"
            Top             =   510
            Width           =   675
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   70
            Text            =   "Text1"
            Top             =   840
            Width           =   2775
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   510
            Width           =   2775
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   1005
            Left            =   30
            TabIndex        =   74
            Top             =   1230
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   1773
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   60
            Left            =   30
            TabIndex        =   79
            Top             =   420
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   60
            Left            =   30
            TabIndex        =   80
            Top             =   2280
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
         Begin EditLib.fpDateTime ipp_FecCes_Ant 
            Height          =   315
            Left            =   8220
            TabIndex        =   81
            Top             =   2370
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
         Begin EditLib.fpDateTime ipp_FecIng_Ant 
            Height          =   315
            Left            =   2010
            TabIndex        =   82
            Top             =   2370
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
            Caption         =   "Fecha de Cese:"
            Height          =   315
            Index           =   66
            Left            =   6210
            TabIndex        =   84
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   6
            Left            =   90
            TabIndex        =   83
            Top             =   2370
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Trabajo Anterior:"
            Height          =   285
            Index           =   7
            Left            =   90
            TabIndex        =   78
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   76
            Top             =   540
            Width           =   1785
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   90
            TabIndex        =   75
            Top             =   870
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   2295
         Left            =   30
         TabIndex        =   59
         Top             =   1230
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   4048
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
         Begin VB.ComboBox cmb_TipDoc_Pro 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   30
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc_Pro 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Buscar_Pro 
            Height          =   675
            Left            =   9450
            Picture         =   "AteCli_frm_136.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia_Pro 
            Height          =   675
            Left            =   10140
            Picture         =   "AteCli_frm_136.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Editar_Pro 
            Height          =   675
            Left            =   10860
            Picture         =   "AteCli_frm_136.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   30
            Width           =   675
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Pro 
            Height          =   1515
            Left            =   30
            TabIndex        =   65
            Top             =   750
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   2672
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   90
            TabIndex        =   67
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   66
            Top             =   60
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   4845
         Left            =   30
         TabIndex        =   1
         Top             =   3570
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   8546
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
         Begin VB.ComboBox cmb_FreHab 
            Height          =   315
            Left            =   8220
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2820
            Width           =   3315
         End
         Begin VB.ComboBox cmb_NomCar 
            Height          =   315
            Left            =   2010
            TabIndex        =   23
            Text            =   "cmb_Dep_NomCar"
            Top             =   3480
            Width           =   3315
         End
         Begin VB.TextBox txt_NomCar 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   3480
            Width           =   3315
         End
         Begin VB.TextBox txt_TelDir 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   4140
            Width           =   1640
         End
         Begin VB.TextBox txt_NumAnx 
            Height          =   315
            Left            =   8220
            MaxLength       =   5
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   3810
            Width           =   1640
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   4470
            Width           =   3315
         End
         Begin VB.TextBox txt_NomAre 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   3810
            Width           =   3315
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   4140
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   2370
            Width           =   1640
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   15
            Text            =   "cmb_DstDir"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   14
            Text            =   "cmb_DptDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   2370
            Width           =   1640
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8220
            TabIndex        =   9
            Text            =   "cmb_PrvDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8220
            MaxLength       =   120
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9870
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1050
            Width           =   1665
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8220
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   2370
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipOfi 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_SitTra 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   3315
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   2010
            TabIndex        =   25
            Top             =   2820
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
            TabIndex        =   26
            Top             =   3150
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
            TabIndex        =   27
            Top             =   2730
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
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono Directo:"
            Height          =   285
            Index           =   59
            Left            =   90
            TabIndex        =   50
            Top             =   4140
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado (S/.):"
            Height          =   285
            Index           =   61
            Left            =   90
            TabIndex        =   49
            Top             =   2820
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Frecuencia Haberes:"
            Height          =   315
            Index           =   56
            Left            =   6210
            TabIndex        =   48
            Top             =   2820
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   62
            Left            =   90
            TabIndex        =   47
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   58
            Left            =   90
            TabIndex        =   46
            Top             =   3150
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Anexo:"
            Height          =   285
            Index           =   64
            Left            =   6210
            TabIndex        =   45
            Top             =   3810
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo (Especificar):"
            Height          =   285
            Index           =   57
            Left            =   6210
            TabIndex        =   44
            Top             =   3480
            Width           =   1665
         End
         Begin VB.Label lbl_General 
            Caption         =   "E-mail:"
            Height          =   285
            Index           =   60
            Left            =   90
            TabIndex        =   43
            Top             =   4470
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Area:"
            Height          =   285
            Index           =   63
            Left            =   90
            TabIndex        =   42
            Top             =   3810
            Width           =   1605
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   41
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   90
            TabIndex        =   40
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   90
            TabIndex        =   39
            Top             =   1710
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   90
            TabIndex        =   38
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   42
            Left            =   90
            TabIndex        =   37
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6210
            TabIndex        =   36
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6210
            TabIndex        =   35
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6210
            TabIndex        =   34
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6210
            TabIndex        =   33
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6210
            TabIndex        =   32
            Top             =   1050
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Oficina:"
            Height          =   285
            Index           =   40
            Left            =   90
            TabIndex        =   31
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   41
            Left            =   90
            TabIndex        =   30
            Top             =   720
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Celular de la Empresa:"
            Height          =   285
            Index           =   0
            Left            =   6210
            TabIndex        =   29
            Top             =   4140
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Situación Trabajador:"
            Height          =   285
            Index           =   8
            Left            =   90
            TabIndex        =   28
            Top             =   60
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   51
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
            TabIndex        =   52
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
            Picture         =   "AteCli_frm_136.frx":123C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   53
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
            TabIndex        =   54
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
            TabIndex        =   55
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   56
         Top             =   11280
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10890
            Picture         =   "AteCli_frm_136.frx":1546
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10200
            Picture         =   "AteCli_frm_136.frx":1988
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

